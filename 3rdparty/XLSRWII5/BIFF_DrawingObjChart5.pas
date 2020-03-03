unit BIFF_DrawingObjChart5;

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

uses Classes, SysUtils, Contnrs,
     BIFF_RecsII5, BIFF_Utils5, BIFF_DrawingObjAnchor5, BIFF_Escher5,
     BIFF_RecordStorageChart5, BIFF_FormulaHandler5, BIFF_EscherTypes5,
     BIFF_Stream5, BIFF_DecodeFormula5,
     Xc12Utils5, Xc12DataStylesheet5,
     XLSUtils5;

//* Line style for lines on charts charts.
type TChartLineStyle = (clsSolid,      //* Solid line
                        clsDash,       //* Dashed line
                        clsDot,        //* Dotted line
                        clsDashDot,    //* Dash dot line
                        clsDashDotDot, //* Dash dot dot line
                        clsNone,       //* No line
                        clsDarkGray,   //* Dark gray line
                        clsMediumGray, //* Medium gray line
                        clsLightGray   //* Light gray line
                        );
{$ifdef ver130}
type TChartLineWeight = (clwNarrow,clwMedium,clwWide,clwHairline);
{$else}
//* Line weight for lines on charts.
type TChartLineWeight = (clwNarrow,          //* Narrow line
                         clwMedium,          //* Medium line
                         clwWide,            //* Wide line
                         clwHairline = $FFFF //* Hair line
                         );
{$endif}
//* Line markers for values on charts.
type TChartMarkerStyle = (cmsNone,              //* No marker
                          cmsSquare,            //* Square marker
                          cmsDiamond,           //* Diamond marker
                          cmsTriangle,          //* Triangle marker
                          cmsX,                 //* 'X' marker
                          cmsStar,              //* Star marker
                          cmsDowJones,          //* Stock chart marker
                          cmsStandardDeviation, //* Standard deviation marker
                          cmsCircle,            //* Circle marker
                          cmsPlus               //* Plus marker
                          );
//* What kind of data label that is displayed.
type TChartDataLabel = (cdlValue,        //* The value
                        cdlValuePercent, //* The value as percent
                        cdlSmoothedLine, //* A smooted line
                        cdlCategory,     //* The category
                        cdlBubbleSize    //* Bubble sixe on bubble charts.
                        );
     //* Set of data label options.
     TChartDataLabels = set of TChartDataLabel;
{$ifdef ver130}
type TChartLabelType = (cltNA1,cltTitle,cltYAxis,cltXAxis,cltDataPoint,cltNA2,cltNA3,cltZAxis);
{$else}
//* What kind of chart label that is displayed.
type TChartLabelType = (cltTitle = 1,    //* Chart title
                        cltYAxis = 2,    //* Y axis value
                        cltXAxis = 3,    //* X axis value
                        cltDataPoint = 4,//* Data point value
                        cltZAxis = 7     //* Z axis value
                        );
{$endif}
//* Options for the chart.
type TChartSheetOption = (csoOnlyVisibleCells, //* Include only visible cells
                          csoDoNotSize         //* Do not resize the chart
                          );
     //* Set of chart options.
     TChartSheetOptions = set of TChartSheetOption;
//* How empty cells are handled.
type TChartEmptyCells = (cecNotPlotted,  //* Do not plot them
                         cecZero,        //* Treat them as zero values
                         cecInterpolated //* Interpolate a value from the surrounding values
                         );
//* Type of tick mark.
type TTickType = (ttInvisible, //* No tick mark is shown.
                  ttInside,    //* Inside of axis line
                  ttOutside,   //* Outside of axis line
                  ttCrossAxis  //* Cross axis line
                  );
//* Tick label position relative to axis line.
type TTickLablePos = (tlpInvisible, //* No tick lable is shown.
                      tlpLowEnd,    //* Low end of plot area.
                      tlpHighEnd,   //* High end of plot area.
                      tlpNextToAxis //* Next to axis.
                      );
//* Defines the point on the category or serie axis where the value axis crosses
type TCatSerAxisScalingOption = (csasoCrossesMidCategory, //* Axis crosses midcategory.
                                 csasoValueCrossesRight,  //* Value axis crosses at the far right category (in a line, bar, column, scatter, or area chart; 2-D charts only).
                                 csasoReverseCategories   //* Display categories in reverse order.
                                 );
     //* Set of axis options.
type TCatSerAxisScalingOptions = set of TCatSerAxisScalingOption;
//* Scaling options for chart axis.
type TValueAxisScalingOption = (vasoAutoMin,              //* Automatic minimum selected.
                                vasoAutoMax,              //* Automatic maximum selected.
                                vasoAutoMajor,            //* Automatic major unit selected.
                                vasoAutoMinor,            //* Automatic minor unit selected.
                                vasoAutoCross,            //* Automatic category crossing point selected.
                                vasoLogScale,             //* Logarithmic scale.
                                vasoReverse,              //* Values in reverse order.
                                vasoCategoryAxisCrossing  //* Category axis to cross at maximum value.
                                );
     //* Set of axis scaling options.
type TValueAxisScalingOptions = set of TValueAxisScalingOption;
//* Options for 3D charts.
type TChart3DOption = (c3doPerspectiveTrans, //* Use perspective transform.
                       c3doClusterdCols,     //* 3-D columns are clustered or stacked.
                       c3doAutoScaling,      //* Use auto scaling.
                       c3do2dWalls           //* Use 2D walls and gridlines
                       );
     //* Set 3D chart options.
     TChart3DOptions = set of TChart3DOption;
//* Defines the location of the legend on the display and its overall size. The displayed legend contains all series on the chart.
type TChartLegendType = (cltBottom,   //* Legend at bottom.
                         cltCorner,   //* Legend at the corner.
                         cltTop,      //* Legend at top.
                         cltRight,    //* Legend at right.
                         cltLeft,     //* Legend at left.
                         cltNotDocked //* Legend is not docked.
                         );
//* Legend spacing. not used for Excel.
type TChartLegendSpacing = (clsClose,clsMedium,clsOpen);
//* What kind of object that can be used for bar chart series.
type TSerieItemGeomtery = (sigDefault,       //* No object is used.
                           sigCylinder,      //* Cylinder.
                           sigPyramid,       //* Pyramid.
                           sigChoppedPyramid,//* Chopped pyramid.
                           sigCone,          //* Cone.
                           sigChoppedCone,   //* Chopped cone.
                           sigLast           //* This value is not used.
                           );
//* Gradient fill style.
type TGradientFillStyle = (gfsHorizontal,   //* Horizontal fill.
                           gfsVertical,     //* Vertical fill.
                           gfsDiagonalUp,   //* Diagonal up fill.
                           gfsDiagonalDown, //* Diagonal down fill.
                           gfsFromCorner,   //* From corner fill.
                           gfsFromCenter    //* From center fill.
                           );

//* ~exclude
type TBmpFileHeader = packed record
     ID: array[0..2] of char;
     Size: longword;
     Unknown1: array[0..2] of byte;
     Unknown2: word;
     Unknown3: word;
     end;

//* What kind of chart that is displayed.
type TXLSChartType = (xctNone,   //* No chart.
                      xctArea,   //* Area chart.
                      xctBubble, //* Bubble chart.
                      xctColumn, //* Column chart.
                      xctLine,   //* Line chart.
                      xctPie,    //* Pie chart.
                      xctRadar,  //* Radar chart.
                      xctScatter,//* Scatter chart.
                      xctSurface //* Surface chart.
                      );

//* ~exclude
type TRecordObject = class(TObject)
private
     FRecord: TChartRecord;

     function GetValid: boolean;
protected
     procedure Check; overload;
     procedure Check(Index: integer); overload;
     procedure Assign(Rec: TChartRecord); virtual;
public
     property Valid: boolean read GetValid;
     property Rec: TChartRecord read FRecord;
     end;


{
type TFontBasis = class(TObject)
private
}

type
//* Object that stores settings for how texts are shown in the chart.
    TChartText = class(TRecordObject)
private
     FFont: TXc12Font;

     function GetColor: TXc12IndexColor;
     procedure SetColor(const Value: TXc12IndexColor);
protected
     procedure Assign(Rec: TChartRecord); override;
public
     //* The text color. Please note that the Color property of the ~[link Font] is
     //* not used for this.
     property Color: TXc12IndexColor read GetColor write SetColor;
     //* The font used for the text.
     property Font: TXc12Font read FFont;
     end;

type
//* Options for lines in the chart.
    TLineFormat = class(TRecordObject)
private
     function  GetAutomatic: boolean;
     function  GetLineColor: TXc12IndexColor;
     function  GetLineStyle: TChartLineStyle;
     function  GetLineWeight: TChartLineWeight;
     procedure SetAutomatic(const Value: boolean);
     procedure SetLineColor(const Value: TXc12IndexColor);
     procedure SetLineStyle(const Value: TChartLineStyle);
     procedure SetLineWeight(const Value: TChartLineWeight);
protected
     procedure SetDefault;
public
     //* Color of the line.
     property LineColor: TXc12IndexColor read GetLineColor write SetLineColor;
     //* Style of the line.
     property LineStyle: TChartLineStyle read GetLineStyle write SetLineStyle;
     //* Width of of the line.
     property LineWeight: TChartLineWeight read GetLineWeight write SetLineWeight;
     //* True if the line format is automatic (default values are used).
     property Automatic: boolean read GetAutomatic write SetAutomatic;
     end;

type
//* Options for areas that are filled with colors and
//* optionally a pattern. The border line is given by ~[link TLineFormat].
    TAreaFormat = class(TRecordObject)
private
     function  GetBackgroundColor: TXc12IndexColor;
     function  GetForegroundColor: TXc12IndexColor;
     function  GetPattern: TXc12FillPattern;
     function  GetSwapColorWhenNeg: boolean;
     procedure SetBackgroundColor(const Value: TXc12IndexColor);
     procedure SetForegroundColor(const Value: TXc12IndexColor);
     procedure SetPattern(const Value: TXc12FillPattern);
     procedure SetSwapColorWhenNeg(const Value: boolean);
     function  GetAutomatic: boolean;
     procedure SetAutomatic(const Value: boolean);
protected
     procedure SetDefault;
public
     //* Fill pattern style.
     property Pattern: TXc12FillPattern read GetPattern write SetPattern;
     //* Foreground color for the fill pattern. If no fill pattern is used, this
     //* is the color of the area.
     property ForegroundColor: TXc12IndexColor read GetForegroundColor write SetForegroundColor;
     //* Background color for the fill pattern. Only used with fill patterns,
     //* or if SwapColorWhenNeg is True.
     property BackgroundColor: TXc12IndexColor read GetBackgroundColor write SetBackgroundColor;
     //* Set to true if foreground and background colors shall be swaped if
     //* the value that the area correspnds to is negativ.
     property SwapColorWhenNeg: boolean read GetSwapColorWhenNeg write SetSwapColorWhenNeg;
     //* True if the area format is automatic (default values are used).
     property Automatic: boolean read GetAutomatic write SetAutomatic;
     end;

type
//* Options for line markers on line and scatter charts.
    TMarkerFormat = class(TRecordObject)
private
     function  GetAutomatic: boolean;
     function  GetBackgroundColor: TXc12IndexColor;
     function  GetForegroundColor: TXc12IndexColor;
     function  GetMarkerStyle: TChartMarkerStyle;
     function  GetSize: integer;
     procedure SetAutomatic(const Value: boolean);
     procedure SetBackgroundColor(const Value: TXc12IndexColor);
     procedure SetForegroundColor(const Value: TXc12IndexColor);
     procedure SetMarkerStyle(const Value: TChartMarkerStyle);
     procedure SetSize(const Value: integer);
public
     //* Style of the marker, see ~[link TChartMarkerStyle].
     property MarkerStyle: TChartMarkerStyle read GetMarkerStyle write SetMarkerStyle;
     //* Foreground color of the marker.
     property ForegroundColor: TXc12IndexColor read GetForegroundColor write SetForegroundColor;
     //* Background color of the marker.
     property BackgroundColor: TXc12IndexColor read GetBackgroundColor write SetBackgroundColor;
     //* True if the colors automatic.
     property Automatic: boolean read GetAutomatic write SetAutomatic;
     //* Size of the markers in units of 1/20th of a point.
     property Size: integer read GetSize write SetSize;
     end;


type
//* A TGelFrame contains optional settings for filling effects and pictures on areas in the
//* chart, such as plot area legend, etc. The TGelFrame is not used to
//* define simple colors of an area. See also ~[link TPaintFrame].
    TGelFrame = class(TRecordObject)
private
     FOPT: TOPT;

     procedure InsertPICRec;
     procedure RemovePICRec;
     function  GetPictureType: TMSOBlipType;
protected
     procedure Assign(Rec: TChartRecord); override;
     procedure SetDefault;
     procedure UpdateRecord(Sender: TObject);
     // Loads a picture from a TStream. ~[link PictType] is the type of picture.
     procedure LoadPictFromStream(Stream: TStream; PictType: TMSOBlipType);
     // Loads a picture from a file.
     procedure LoadPictFromFile(Filename: AxUCString);
public
     //* ~exclude
     constructor Create;
     //* ~exclude
     destructor Destroy; override;
     //* Saves the picture to a TStream. This assumes that the filling effect
     //* is a picture.
     //* ~param Stream The stream to write the picture to.
     procedure SavePictToStream(Stream: TStream);
     //* Saves the picture to a file. This assumes that the filling effect
     //* is a picture. If AutoExt is True, the extension of the filename is
     //* then added automatically, depending of the type of picture
     //* (bmp, jpg, etc). The Filename shall then only contain the name and
     //* path, without a dot.
     //* ~param Filename The name of the disk file to write the picture to.
     //* ~param AutoExt Set this to True if the correct extension of the picture type shall be added automatically.
     procedure SavePictToFile(Filename: AxUCString; AutoExt: boolean);
     //* Loads a picture as fill effect from Stream.
     //* ~param Stream The stream to read the picture from.
     //* ~param PictType The type of the picture. As the file is in a stream (and not a disk file), there is
     //* no file extension from where the picture type can be detected.
     procedure FillEffectPicture(Stream: TStream; PictType: TMSOBlipType); overload;
     //* Loads a picture in a file as fill effect. Acceppted file formats are:
     //* BMP, JPG, PNG, WMF, EMF, PICT.
     //* ~param Filename The name of the disk file where the picture is.
     procedure FillEffectPicture(Filename: AxUCString); overload;
     //* Loads a picture as fill effect from Stream. ~[link PictType] is the type of picture.
     //* ~param Stream The stream to read the picture from.
     //* ~param PictType The type of the picture. As the file is in a stream (and not a disk file), there is
     //* no file extension from where the picture type can be detected.
     procedure FillEffectTexture(Stream: TStream; PictType: TMSOBlipType); overload;
     //* Loads a texture as fill effect from Stream. ~[link PictType] is the type of picture.
     //* A texture is the same as a picture, but textures are tiled, whithout
     //* changing the scale, while pictures are scaled to fit the area.
     //* ~param Filename The name of the disk file where the picture is.
     procedure FillEffectTexture(Filename: AxUCString); overload;
     //* Sets the fill effect to a gradient from Color1 and Color2.
     //* GradientFillStyle is the style, see ~[link TGradientFillStyle].
     //* ~param Color1 The first color.
     //* ~param Color2 The second color.
     //* ~param GradientStyle The gradient fill style.
     procedure FillEffectGradient(Color1,Color2: TXc12IndexColor; GradientStyle: TGradientFillStyle);

     //* The type of picture, if a picture is used. Read only.
     property PictureType: TMSOBlipType read GetPictureType;
     end;

type
//* Paint settings for areas.
    TPaintFrame = class(TRecordObject)
private
     FLineFormat: TLineFormat;
     FAreaFormat: TAreaFormat;
     FGelFrame: TGelFrame;

     function  GetHasFillEffects: boolean;
     procedure SetHasFillEffects(const Value: boolean);
protected
     procedure Assign(Rec: TChartRecord); override;
     procedure SetDefault(Parent: TChartRecord);
public
     //* ~exclude
     constructor Create;
     //* ~exclude
     destructor Destroy; override;

     //* Settings for the border line.
     property LineFormat: TLineFormat read FLineFormat write FLineFormat;
     //* Settings for the area.
     property AreaFormat: TAreaFormat read FAreaFormat write FAreaFormat;
     //* Settings for optional fill effects. If there are no fill effects,
     //* ~[link HasFillEffects] must be set to True.
     property FillEffects: TGelFrame read FGelFrame write FGelFrame;
     //* True if the frame has fill effects. In order to add fill effects,
     //* HasFillEffects must be set to True. If set to False, any fill effects
     //* will be removed.
     property HasFillEffects: boolean read GetHasFillEffects write SetHasFillEffects;
     end;

//* ~exclude
type TTextFrame = class(TRecordObject)
private
     FLineFormat: TLineFormat;
     FAreaFormat: TAreaFormat;

     function  GetAutoPosition: boolean;
     function  GetShadow: boolean;
     procedure SetAutoPosition(const Value: boolean);
     procedure SetShadow(const Value: boolean);
     function  GetAutoSize: boolean;
     procedure SetAutoSize(const Value: boolean);
protected
     procedure Assign(Rec: TChartRecord); override;
public
     constructor Create;
     destructor Destroy; override;

     property LineFormat: TLineFormat read FLineFormat write FLineFormat;
     property AreaFormat: TAreaFormat read FAreaFormat write FAreaFormat;
     property Shadow: boolean read GetShadow write SetShadow;
     property AutoSize: boolean read GetAutoSize write SetAutoSize;
     property AutoPosition: boolean read GetAutoPosition write SetAutoPosition;
     end;

type
//* Optional text lables displayed in the chart.
    TChartLabel = class(TCollectionItem)
private
     FRecord: TChartRecord;
     FRecPOS: TChartRecord;
     FRecAI: TChartRecord;
     FRecOBJECTLINK: TChartRecord;
     FRecSERIESTEXT: TChartRecord;

     function  GetDataPointIndex: integer;
     function  GetLabelType: TChartLabelType;
     function  GetSerieIndex: integer;
     function  GetText: AxUCString;
     function  GetXOffset: integer;
     function  GetYOffset: integer;
     procedure SetDataPointIndex(const Value: integer);
     procedure SetLabelType(const Value: TChartLabelType);
     procedure SetSerieIndex(const Value: integer);
     procedure SetText(const Value: AxUCString);
     procedure SetXOffset(const Value: integer);
     procedure SetYOffset(const Value: integer);
     function  GetRotation: integer;
     procedure SetRotation(const Value: integer);
protected
     procedure AssignRecords(Rec: TChartRecord);
     procedure SetDefault;
public
     //* ~exclude
     constructor Create(Collection: TCollection); override;
     //* ~exclude
     destructor Destroy; override;

     //* What kind of data the the label belongs to. See ~[link TChartLabelType].
     property LabelType: TChartLabelType read GetLabelType;
     //* X offset for the text.
     property XOffset: integer read GetXOffset write SetXOffset;
     //* Y offset for the text.
     property YOffset: integer read GetYOffset write SetYOffset;
     //* Index for the serie that the label belongs to. Only used when linking
     //* to a serie value.
     property SerieIndex: integer read GetSerieIndex write SetSerieIndex;
     //* Which data point in the serie the lable belongs. Only used when
     //* linking to a serie value.
     property DataPointIndex: integer read GetDataPointIndex write SetDataPointIndex;
     //* The text for the lable.
     property Text: AxUCString read GetText write SetText;
     //* Text rotation.~[br]
     //* 0 = Horizontal.~[br]
     //* 90 = Up 90 degrees.~[br]
     //* 180 = Down 180 degrees.~[br]
     //* -90 = Down (inverted from 180).
     property Rotation: integer read GetRotation write SetRotation;
     end;

type
//* List of TChartLable.
    TChartLabels = class(TCollection)
private
     FOwner: TPersistent;

     function  GetItems(Index: integer): TChartLabel;
protected
     function  GetOwner: TPersistent; override;
public
     //* ~exclude
     constructor Create(AOwner: TPersistent);
     //* Add a new TChartLable.
     //* ~param LabelType What kind of lable that shall be added.
     //* ~result The new TChartLabel object.
     function  Add(LabelType: TChartLabelType): TChartLabel;
     //* ~exclude
     function  AddFromFile: TChartLabel;
     //* Delete TChartLable at Index in the list.
     //* ~param Index Index of the TChartLable to be deleted.
     procedure Delete(Index: integer);
     //* The TChartLable items in the list.
     property  Items[Index: integer]: TChartLabel read GetItems; default;
     end;

//* A TDataPoint is a serie value on the chart.
type TDataPoint = class(TObject)
private
     FRecord: TChartRecord;

     FLineFormat: TLineFormat;
     FAreaFormat: TAreaFormat;
     FMarkerFormat: TMarkerFormat;

     FRecPIEFORMAT: TChartRecord;
     FRecATTACHEDLABEL: TChartRecord;

     function  GetPointIndex: integer;
     function  GetDataLablel: TChartDataLabels;
     function  GetPieSliceDist: integer;
     procedure SetPointIndex(const Value: integer);
     procedure SetDataLablel(const Value: TChartDataLabels);
     procedure SetPieSliceDist(const Value: integer);
protected
     procedure AssignRecords(Rec: TChartRecord);
public
     //* ~exclude
     constructor Create;
     //* ~exclude
     destructor Destroy; override;

     //* Index of the serie value that the TDataPoint represents.
     property PointIndex: integer read GetPointIndex write SetPointIndex;
     //* Distance of pie slice from center of pie, in percents. Only used for pie charts.
     property PieSliceDist: integer read GetPieSliceDist write SetPieSliceDist;
     //* Data label options.
     property DataLablel: TChartDataLabels read GetDataLablel write SetDataLablel;
     //* Line format.
     property LineFormat: TLineFormat read FLineFormat write FLineFormat;
     //* Area fill format.
     property AreaFormat: TAreaFormat read FAreaFormat write FAreaFormat;
     //* Marker format.
     property MarkerFormat: TMarkerFormat read FMarkerFormat write FMarkerFormat;
     end;

//* A list of TDataPoint
type TDataPoints = class(TObjectList)
private
     FOwner: TObject;

     function  GetItems(Index: integer): TDataPoint;
public
     //* ~exclude
     constructor Create(AOwner: TObject);
     //* Add a new TDataPoint.
     //* ~result The new TDataPoint object.
     function  Add: TDataPoint;
     //* ~exclude
     function  AddFromFile: TDataPoint;
     //* Deletes a TDataPoint from the list.
     //* ~param Index Index of the TDataPoint to be deleted.
     procedure Delete(Index: integer);
     //* The TDataPoint's in the list.
     property Items[Index: integer]: TDataPoint read GetItems; default;
     end;

type
//* TChartSerie is a serie of values in the chart, that referrers to an area
//* of cells in the workbook. There has to be at least one serie in a chart.
    TXLSChartSerie = class(TCollectionItem)
private
     FRecord: TChartRecord;
     FFormulaHandler: TFormulaHandler;
     FiRecAI_Values: integer;
     FiRecAI_Values2: integer;
     FiRecAI_Category: integer;
     FiRecAI_SerieName: integer;
     PRecGeomtery: TChartRecord;
     FDefDatapoint: TDataPoint;
     FDataPoints: TDataPoints;

     function  GetAIFormula(AI: PCRecAI): AxUCString;
     procedure SetAIFormula(AI: TChartRecord; Value: AxUCString);
     function  GetCategory: AxUCString;
     function  GetValues: AxUCString;
     procedure SetCategory(const Value: AxUCString);
     procedure SetValues(const Value: AxUCString);
     function  GetValues2: AxUCString;
     procedure SetValues2(const Value: AxUCString);
     function  GetSerieName: AxUCString;
     procedure SetSerieName(const Value: AxUCString);
     function  GetGeomtery: TSerieItemGeomtery;
     procedure SetGeomtery(const Value: TSerieItemGeomtery);
     function  GetHasDefDataPoint: boolean;
     procedure SetHasDefDataPoint(const Value: boolean);
protected
     procedure AssignRecords(Rec: TChartRecord);
     procedure CalcValueCount;
public
     //* ~exclude
     constructor Create(Collection: TCollection); override;
     //* ~exclude
     destructor Destroy; override;
     //* ~exclude
     function GetValuesFmla: TDynByteArray;
published
     //* Source cells for the serie. This shall be a vector, i.e the area
     //* given shall be a one dimensional list of rows or columns. The worksheet
     //* name must be included, an the cell references shall be absolute,
     //* otherwise they may change if the chart is moved.
     //* Example, source cells in one column: Sheet1!$D$7:$D$18
     //* Example, source cells in one row: Sheet3!$B$7:$H$7
     property Values: AxUCString read GetValues write SetValues;
     //* Second set of source values. Values2 are used in bubble charts, to
     //* set the size of the bubbles. See also ~[link Values].
     property Values2: AxUCString read GetValues2 write SetValues2;
     //* Source values for category axis. As there only can be one category
     //* source, it shall be assigned to the first serie used. As the first
     //* serie is added automatically with the chart, there is nothing
     //* to choose of. But if there is more than one serie, the catoegory
     //* shall be assigned to the first serie ADDED, that is, the serire
     //* with index 1 (Series[1]).
     property Category: AxUCString read GetCategory write SetCategory;
     //* The name of the serie.
     //* If the serie name not is a reference to a cell, the name mus be within double quotes, as in: "My serie".
     property SerieName: AxUCString read GetSerieName write SetSerieName;
     //* Optional TDataPoint's that can format how each value in the serie is displayed.
     property DataPoints: TDataPoints read FDataPoints write FDataPoints;
     //* Default settings for all data points (values). See also ~[link HasDefDataPoint].
     property DefDataPoint: TDataPoint read FDefDatapoint write FDefDatapoint;
     //* Returns True if there is a default setting for the data points (values).
     //* Set this property to True in order to change the default settings.
     //* See also ~[link DefDataPoint].
     property HasDefDataPoint: boolean read GetHasDefDataPoint write SetHasDefDataPoint;
     //* The values of the series is reprecented by a geometric figure.
     property Geomtery: TSerieItemGeomtery read GetGeomtery write SetGeomtery;
     end;

//* List of TXLSChartSerie
type TXLSChartSeries = class(TCollection)
private
     function GetChartSerie(Index: integer): TXLSChartSerie;
protected
     FOwner: TPersistent;
     FFormulaHandler: TFormulaHandler;

     function  GetOwner: TPersistent; override;
public
     //* ~exclude
     constructor Create(AOwner: TPersistent; FormulaHandler: TFormulaHandler);
     //* Add a new TXLSChartSerie.
     //* ~result The new TXLSChartSerie object.
     function  Add: TXLSChartSerie;
     //* ~exclude
     function  AddFromFile: TXLSChartSerie;
     //* Delete the TXLSChartSerie at Index.
     //* ~param Index Index of the TXLSChartSerie to be deleted.
     procedure Delete(Index: integer);
     //* The TXLSChartSerie's in the list.
     property Items[Index: integer]: TXLSChartSerie read GetChartSerie; default;
     end;

//* Format for tick marks on the chart.
type TTickFormat = class(TRecordObject)
private
     function GetLabelColor: TXc12IndexColor;
     function GetLabelPosition: TTickLablePos;
     function GetMajorTick: TTickType;
     function GetMinorTick: TTickType;
     function GetTextRotation: integer;
     procedure SetLabelColor(const Value: TXc12IndexColor);
     procedure SetLabelPosition(const Value: TTickLablePos);
     procedure SetMajorTick(const Value: TTickType);
     procedure SetMinorTick(const Value: TTickType);
     procedure SetTextRotation(const Value: integer);
public
     //* Tick style for major ticks. See also ~[link TCatSerAxisScaling.TickMarksFrequency]
     property MajorTick: TTickType read GetMajorTick write SetMajorTick;
     //* Tick style for minor ticks. See also ~[link TCatSerAxisScaling.TickMarksFrequency]
     property MinorTick: TTickType read GetMinorTick write SetMinorTick;
     //* Position for tick labels.
     property LabelPosition: TTickLablePos read GetLabelPosition write SetLabelPosition;
     //* Color for tick labels.
     property LabelColor: TXc12IndexColor read GetLabelColor write SetLabelColor;
     //* Rotation of tick labels.
     property TextRotation: integer read GetTextRotation write SetTextRotation;
     end;

//* Scaling settings for a category or series axis
type TCatSerAxisScaling = class(TRecordObject)
private
     function  GetCrossingPoint: integer;
     function  GetLabelsFrequency: integer;
     function  GetOptions: TCatSerAxisScalingOptions;
     function  GetTickMarksFrequency: integer;
     procedure SetCrossingPoint(const Value: integer);
     procedure SetLabelsFrequency(const Value: integer);
     procedure SetOptions(const Value: TCatSerAxisScalingOptions);
     procedure SetTickMarksFrequency(const Value: integer);
public
     //* The CrossingPoint defines the point on the category axis where the value axis crosses.
     //* A value of 1 indicates that the value axis crosses to the left, or in the center,
     //* of the first category (depending on if the csasoCrossesMidCategory in Options is set);
     //* a value of 2 indicates that the value axis crosses to the left or center of the
     //* second category, and so on.
     property CrossingPoint: integer read GetCrossingPoint write SetCrossingPoint;
     //* The LabelsFrequency property defines how often labels appear along the category or series axis.
     //* A value of 1 indicates that a category label will appear with each category,
     //* a value of 2 eans a label appears every other category, and so on.
     property LabelsFrequency: integer read GetLabelsFrequency write SetLabelsFrequency;
     //* The TickMarksFrequency property defines how often tick marks appear along the category or series axis.
     //* A value of 1 indicates that a tick mark will appear between each category or series;
     //* a value of 2 means a label appears between every other category or series, etc.
     property TickMarksFrequency: integer read GetTickMarksFrequency write SetTickMarksFrequency;
     // Options for TCatSerAxisScaling.
     property Options: TCatSerAxisScalingOptions read GetOptions write SetOptions;
     end;

//* Settings for the values axis.
type TValueAxisScaling = class(TRecordObject)
private
     procedure CheckFlag(Flag: word; Value: double);

     function  GetCategoryAxisCrossing: double;
     function  GetMajorInc: double;
     function  GetMaxValue: double;
     function  GetMinorInc: double;
     function  GetMinValue: double;
     function  GetOptions: TValueAxisScalingOptions;
     procedure SetCategoryAxisCrossing(const Value: double);
     procedure SetMajorInc(const Value: double);
     procedure SetMaxValue(const Value: double);
     procedure SetMinorInc(const Value: double);
     procedure SetMinValue(const Value: double);
     procedure SetOptions(const Value: TValueAxisScalingOptions);
public
     //* MinValue defines the minimum numeric value that appears along the value axis.
     //* This value is all zeros if Auto Minimum is selected on the Scale tab of the Format Axis dialog box.
     property MinValue: double read GetMinValue write SetMinValue;
     //* MaxValue defines the maximum value displayed along the value axis
     //* and is all zeros if Auto Maximum is selected.
     property MaxValue: double read GetMaxValue write SetMaxValue;
     //* MajorInc defines the increment (unit) of the major value divisions
     //* (gridlines) along the value axis. MajorInc is all zeros if Auto Major
     //* Unit is selected on the Scale tab of the Format Axis dialog box.
     property MajorInc: double read GetMajorInc write SetMajorInc;
     //* MinorInc defines the minor value divisions (gridlines) along
     //* the value axis and is all zeros if Auto Minor Unit is selected.
     property MinorInc: double read GetMinorInc write SetMinorInc;
     //* CategoryAxisCrossing defines the value along the value axis at which
     //* the category axis crosses. This field is all zeros if Auto Category
     //* Axis Crosses At is selected.
     property CategoryAxisCrossing: double read GetCategoryAxisCrossing write SetCategoryAxisCrossing;
     //* Options for TValueAxisScaling.
     property Options: TValueAxisScalingOptions read GetOptions write SetOptions;
     end;

//* Base class for chart axis.
type TChartAxis = class(TRecordObject)
private
     FTick: TTickFormat;
protected
     procedure Assign(Rec: TChartRecord); override;
public
     constructor Create;
     destructor Destroy; override;
     //* Tick marks for the axis.
     property Tick: TTickFormat read FTick;
     end;

//* Category axis settings.
type TCategoryAxis = class(TChartAxis)
private
     FCatSerAxisScaling: TCatSerAxisScaling;
protected
     procedure Assign(Rec: TChartRecord); override;
public
     //* ~exclude
     constructor Create;
     //* ~exclude
     destructor Destroy; override;

     //* Scaling settings of the axis.
     property CatSerAxisScaling: TCatSerAxisScaling read FCatSerAxisScaling write FCatSerAxisScaling;
     end;

//* Serie axis settings.
type TSerieAxis = class(TChartAxis)
private
     FCatSerAxisScaling: TCatSerAxisScaling;
protected
     procedure Assign(Rec: TChartRecord); override;
public
     //* ~exclude
     constructor Create;
     //* ~exclude
     destructor Destroy; override;

     //* Scaling settings of the axis.
     property CatSerAxisScaling: TCatSerAxisScaling read FCatSerAxisScaling write FCatSerAxisScaling;
     end;

//* Value axis settings.
type TValueAxis = class(TChartAxis)
private
     FValueAxisScaling: TValueAxisScaling;
protected
     procedure Assign(Rec: TChartRecord); override;
public
     //* ~exclude
     constructor Create;
     //* ~exclude
     destructor Destroy; override;

     //* Scaling settings of the axis.
     property ValueAxisScaling: TValueAxisScaling read FValueAxisScaling write FValueAxisScaling;
     end;

//* Settings for 3D charts.
type TChart3D = class(TRecordObject)
private
     function  GetDistance: integer;
     function  GetElevation: integer;
     function  GetOptions: TChart3DOptions;
     function  GetRelDepth: integer;
     function  GetRelHeight: integer;
     function  GetRotation: integer;
     function  GetSeriesSpace: integer;
     procedure SetDistance(const Value: integer);
     procedure SetElevation(const Value: integer);
     procedure SetOptions(const Value: TChart3DOptions);
     procedure SetRelDepth(const Value: integer);
     procedure SetRelHeight(const Value: integer);
     procedure SetRotation(const Value: integer);
     procedure SetSeriesSpace(const Value: integer);
public
     //* Rotation angle (0 to 360 degrees).
     property Rotation: integer read GetRotation write SetRotation;
     //* Elevation angle (–90 to +90 degrees).
     property Elevation: integer read GetElevation write SetElevation;
     //* Distance from eye to chart (0 to 100).
     property Distance: integer read GetDistance write SetDistance;
     //* Height of plot volume relative to width and depth.
     property RelHeight: integer read GetRelHeight write SetRelHeight;
     //* Depth of points relative to width.
     property RelDepth: integer read GetRelDepth write SetRelDepth;
     //* Space between series:
     property SeriesSpace: integer read GetSeriesSpace write SetSeriesSpace;
     //* Options for 3D charts.
     property Options: TChart3DOptions read GetOptions write SetOptions;
     end;

//* Settings for the chart's legend.
type TChartLegend = class(TRecordObject)
private
     FFrame: TPaintFrame;

     function  GetLegendSpacing: TChartLegendSpacing;
     function  GetLegendType: TChartLegendType;
     function  GetVertical: boolean;
     procedure SetLegendSpacing(const Value: TChartLegendSpacing);
     procedure SetLegendType(const Value: TChartLegendType);
     procedure SetVertical(const Value: boolean);
     function  GetHasFrame: boolean;
     procedure SetHasFrame(const Value: boolean);
protected
     procedure Assign(Rec: TChartRecord); override;
public
     //* ~exclude
     constructor Create;
     //* ~exclude
     destructor Destroy; override;

     //* The type of the legend.
     property LegendType: TChartLegendType read GetLegendType write SetLegendType;
     //* Not used.
     property LegendSpacing: TChartLegendSpacing read GetLegendSpacing write SetLegendSpacing;
     //* True if the legend is vertical.
     property Vertical: boolean read GetVertical write SetVertical;
     //* Settings for the legends frame area. Changing this will override the default settings.
     //* See also ~[link HasFrame]
     property Frame: TPaintFrame read FFrame write FFrame;
     //* Set HasFrame to True in order to change the default settings for the legends frame area.
     //* See also ~[link Frame]
     property HasFrame: boolean read GetHasFrame write SetHasFrame;
     end;

//* Base class for chart styles.
//* See also ~[link TChartPlotArea.ChartStyle]
type TChartStyle = class(TRecordObject)
private
     FChart3D: TChart3D;
     FLegend: TChartLegend;
     FDefDatapoint: TDataPoint;

     function  GetVaryColors: boolean;
     procedure SetVaryColors(const Value: boolean);
     function  GetHasDefDataPoint: boolean;
     procedure SetHasDefDataPoint(const Value: boolean);
protected
     procedure Assign(Rec: TChartRecord); override;
     function ChartType: TXLSChartType; virtual;
public
     //* ~exclude
     constructor Create;
     //* ~exclude
     destructor Destroy; override;

     //* ~exclude
     property Chart3D: TChart3D read FChart3D;
     //* Legend for the chart.
     property Legend: TChartLegend read FLegend write FLegend;
     //* Vary color for each data point.
     property VaryColors: boolean read GetVaryColors write SetVaryColors;
     //* Default settings for the data points (values). See also ~[link HasDefDataPoint]
     property DefDatapoint: TDataPoint read FDefDatapoint;
     //* Set HasDefDataPoint to True in order to change the values of ~[link DefDataPoint]
     property HasDefDataPoint: boolean read GetHasDefDataPoint write SetHasDefDataPoint;
     end;

//* Settings for area style charts.
//* See also ~[link TChartPlotArea.ChartStyle]
type TChartStyleArea = class(TChartStyle)
protected
     procedure Assign(Rec: TChartRecord); override;
     function ChartType: TXLSChartType; override;
public
     constructor Create;
     destructor Destroy; override;
     end;

//* Settings for bar and column style charts.
//* See also ~[link TChartPlotArea.ChartStyle]
type TChartStyleBarColumn = class(TChartStyle)
private
     function  GetIsBar: boolean;
     procedure SetIsBar(const Value: boolean);
     function  GetStacked: boolean;
     procedure SetStacked(const Value: boolean);
protected
     procedure Assign(Rec: TChartRecord); override;
     function  ChartType: TXLSChartType; override;
public
     //* ~exclude
     constructor Create;
     //* ~exclude
     destructor Destroy; override;

     //* IsBar is True when the chart is bar chart, when False, the chart is a column chart.
     property IsBar: boolean read GetIsBar write SetIsBar;
     //* WHen True, stack the displayed values
     property Stacked: boolean read GetStacked write SetStacked;
     end;

//* Settings for surface style charts.
//* See also ~[link TChartPlotArea.ChartStyle]
type TChartStyleSurface = class(TChartStyle)
protected
     procedure Assign(Rec: TChartRecord); override;
public
     //* ~exclude
     constructor Create;
     //* ~exclude
     destructor Destroy; override;
     end;

//* Settings for bubble and scatter style charts.
//* See also ~[link TChartPlotArea.ChartStyle]
type TChartStyleScatter = class(TChartStyle)
private
     function  GetBubbleIsArea: boolean;
     function  GetBubblePercent: integer;
     function  GetIsBubble: boolean;
     function  GetNegativeBubble: boolean;
     function  GetShadow: boolean;
     procedure SetBubbleIsArea(const Value: boolean);
     procedure SetBubblePercent(const Value: integer);
     procedure SetIsBubble(const Value: boolean);
     procedure SetNegativeBubble(const Value: boolean);
     procedure SetShadow(const Value: boolean);
protected
     procedure Assign(Rec: TChartRecord); override;
public
     constructor Create;
     destructor Destroy; override;

     //* Percent of largest bubble compared to chart in general
     property BubblePercent: integer read GetBubblePercent write SetBubblePercent;
     //* When True, the size of the bubbles is the bubbles area, when False, the size is the bubbles width.
     property BubbleIsArea: boolean read GetBubbleIsArea write SetBubbleIsArea;
     //* When True, it's a bubble chart, otherwise it's a scatter chart.
     property IsBubble: boolean read GetIsBubble write SetIsBubble;
     //* Show negative bubbles.
     property NegativeBubble: boolean read GetNegativeBubble write SetNegativeBubble;
     //* Bubble series has a shadow
     property Shadow: boolean read GetShadow write SetShadow;
     end;

//* Settings for pie style charts.
//* See also ~[link TChartPlotArea.ChartStyle]
type TChartStylePie = class(TChartStyle)
private
     function  GetAngle: integer;
     function  GetDonutHole: integer;
     function  GetLeaderLines: boolean;
     function  GetShadow: boolean;
     procedure SetAngle(const Value: integer);
     procedure SetDonutHole(const Value: integer);
     procedure SetLeaderLines(const Value: boolean);
     procedure SetShadow(const Value: boolean);
protected
     procedure Assign(Rec: TChartRecord); override;
public
     //* ~exclude
     constructor Create;
     //* ~exclude
     destructor Destroy; override;

     //* Angle of the first pie slice expressed in degrees.
     property Angle: integer read GetAngle write SetAngle;
     //* 0 = true pie chart. Non-zero = size of center hole in a donut chart (as a percentage).
     property DonutHole: integer read GetDonutHole write SetDonutHole;
     //* When True, his pie has a shadow.
     property Shadow: boolean read GetShadow write SetShadow;
     //* When True, show leader lines to data labels.
     property LeaderLines: boolean read GetLeaderLines write SetLeaderLines;
     end;

// Bar of Pie/Pie of Pie
//* Settings for Bar of Pie/Pie of Pie style charts.
//* See also ~[link TChartPlotArea.ChartStyle]
type TChartStyleBopPop = class(TChartStyle)
protected
     procedure Assign(Rec: TChartRecord); override;
public
     constructor Create;
     destructor Destroy; override;
     end;

//* Settings for line style charts.
//* See also ~[link TChartPlotArea.ChartStyle]
type TChartStyleLine = class(TChartStyle)
private
     function  GetCategoryAsPercent: boolean;
     function  GetShadow: boolean;
     function  GetStacked: boolean;
     procedure SetCategoryAsPercent(const Value: boolean);
     procedure SetShadow(const Value: boolean);
     procedure SetStacked(const Value: boolean);
protected
     procedure Assign(Rec: TChartRecord); override;
public
     //* ~exclude
     constructor Create;
     //* ~exclude
     destructor Destroy; override;
     //* When True, stack the displayed values.
     property Stacked: boolean read GetStacked write SetStacked;
     //* When True, each category is broken down as a percentage
     property CategoryAsPercent: boolean read GetCategoryAsPercent write SetCategoryAsPercent;
     //* When True, the line has a shadow.
     property Shadow: boolean read GetShadow write SetShadow;
     end;

//* Settings for radar style charts.
//* See also ~[link TChartPlotArea.ChartStyle]
type TChartStyleRadar = class(TChartStyle)
private
     function  GetAxisLabels: boolean;
     function  GetShadow: boolean;
     procedure SetAxisLabels(const Value: boolean);
     procedure SetShadow(const Value: boolean);
protected
     procedure Assign(Rec: TChartRecord); override;
public
     //* ~exclude
     constructor Create;
     //* ~exclude
     destructor Destroy; override;

     //* When True, chart contains radar axis labels.
     property AxisLabels: boolean read GetAxisLabels write SetAxisLabels;
     //* When True, the radar serie has a shadow.
     property Shadow: boolean read GetShadow write SetShadow;
     end;

//* Settings for radar area style charts.
//* See also ~[link TChartPlotArea.ChartStyle]
type TChartStyleRadarArea = class(TChartStyle)
private
     function  GetAxisLabels: boolean;
     function  GetShadow: boolean;
     procedure SetAxisLabels(const Value: boolean);
     procedure SetShadow(const Value: boolean);
protected
     procedure Assign(Rec: TChartRecord); override;
public
     //* ~exclude
     constructor Create;
     //* ~exclude
     destructor Destroy; override;

     //* When True, chart contains radar axis labels.
     property AxisLabels: boolean read GetAxisLabels write SetAxisLabels;
     //* When True, the radar serie has a shadow.
     property Shadow: boolean read GetShadow write SetShadow;
     end;

//* Settings for the area where the chart's series is displayed.
type TChartPlotArea = class(TRecordObject)
private
     FCategoryAxis: TCategoryAxis;
     FSerieAxis: TSerieAxis;
     FValueAxis: TValueAxis;
     FChartStyle: TChartStyle;
     FFrame: TPaintFrame;

     function  GetChartType: TXLSChartType;
     procedure SetChartType(const Value: TXLSChartType);
     function  GetIs3D: boolean;
     procedure SetIs3D(const Value: boolean);
     function  GetHasLegend: boolean;
     procedure SetHasLegend(const Value: boolean);
protected
     procedure Assign(Rec: TChartRecord); override;
public
     //* ~exclude
     constructor Create;
     //* ~exclude
     destructor Destroy; override;

     //* The type of the chart that is displayed.
     property ChartType: TXLSChartType read GetChartType write SetChartType;
     //* The style of the chart type that is created when setting the ~[link ChartType] property.
     //* In order to change the style, ChartStyle must be typecasted to the correct descendant
     //* of TChartStyle. ~[br]
     //* Use these typecastings for the values of ChartType: ~[br]
     //* xctArea = ~[link TChartStyleArea](ChartStyle) ~[br]
     //* xctBubble = ~[link TChartStyleScatter](ChartStyle) ~[br]
     //* xctColumn = ~[link TChartStyleBarColumn](ChartStyle) ~[br]
     //* xctLine = ~[link TChartStyleLine](ChartStyle) ~[br]
     //* xctPie = ~[link TChartStylePie](ChartStyle) ~[br]
     //* xctRadar = ~[link TChartStyleRadar](ChartStyle) ~[br]
     //* xctScatter = ~[link TChartStyleScatter](ChartStyle) ~[br]
     //* xctSurface = ~[link TChartStyleSurface](ChartStyle) ~[br]
     property ChartStyle: TChartStyle read FChartStyle write FChartStyle;
     //* Settings for the chart frame.
     property Frame: TPaintFrame read FFrame write FFrame;
     //* When True, the chart is displayed as 3D. This may not be valid for all chart types.
     property Is3D: boolean read GetIs3D write SetIs3D;
     //* When True, a legend is displayed.
     property HasLegend: boolean read GetHasLegend write SetHasLegend;
     //* Settings for category axis.
     property CategoryAxis: TCategoryAxis read FCategoryAxis write FCategoryAxis;
     //* Settings for serie axis.
     property SerieAxis: TSerieAxis read FSerieAxis write FSerieAxis;
     //* Settings for value axis.
     property ValueAxis: TValueAxis read FValueAxis write FValueAxis;
     end;

//* The chart object. TDrwChart reprecents the whole chart as it is in Excel.
//* ~[br]
//* ~[br]
//* ~[b Example:]
//* ~[br]
//* ~[sample
//*  var
//*    Serie: TXLSChartSerie;
//*  // Add some values.
//*  XLS.Sheet[0~[].FillRandom('A1:C10',1000);
//*  // Add a chart.
//*  with XLS.Sheets[0~[].AddChart do begin
//*    // Position the chart.
//*    Col1 := 5;
//*    Row1 := 4;
//*    Col2 := 11;
//*    Row2 := 19;
//*
//*    // By default, there is one serie in the chart.
//*    Serie := Series[0~[];
//*    // Set the source area to A1:A10
//*    Serie.Values := 'Sheet1!$A$1:$A$10';
//*
//*    // Add another serie.
//*    Serie := Series.Add;
//*    // Set the source area to B1:B10
//*    Serie.Values := 'Sheet1!$B$1:$B$10';
//*    // Set the color of this serie.
//*    Serie.HasDefDataPoint := True;
//*    Serie.DefDataPoint.AreaFormat.ForegroundColor := xcYellow;
//*
//*    // Set the chart type.
//*    PlotArea.ChartType := xctColumn;
//*
//*    // Show a legend.
//*    PlotArea.HasLegend := True;
//*    // Not 3D.
//*    PlotArea.Is3D := False;
//*
//*    // Background color of the chart's area.
//*    PlotArea.Frame.AreaFormat.ForegroundColor := xc27;
//*    // The ChartStyle property must be typecasted to set properties that
//*    // is specific to the created chart type.
//*    TChartStyleBarColumn(PlotArea.ChartStyle).IsBar := True;
//*  end;
//*  // Add a title to the chart.
//*  with XLS.Sheets[0~[].Charts[0~[].Labels.Add(cltTitle) do begin
//*    Text := 'My sweet chart';
//*  end;
//* ]
type TDrwChart = class(TDrwAnchor)
private
     FRecord: TChartRecord;
     FRecCHART: TChartRecord;
     FRecSHTPROPS: TChartRecord;
     FDefaultTextA: TChartText;
     FDefaultTextB: TChartText;
     FLabels: TChartLabels;
     FSeries: TXLSChartSeries;
     FPlotArea: TChartPlotArea;
     FFrame: TPaintFrame;

     function  GetEmptyCells: TChartEmptyCells;
     function  GetSheetOptions: TChartSheetOptions;
     procedure SetEmptyCells(const Value: TChartEmptyCells);
     procedure SetSheetOptions(const Value: TChartSheetOptions);
     function  GetPlotVisibleCellsOnly: boolean;
     procedure SetPlotVisibleCellsOnly(const Value: boolean);
protected
     FFormulaHandler: TFormulaHandler;

     procedure AssignRecords(Records: TChartRecord);
public
     //* ~exclude
     constructor Create(Collection: TCollection); override;
     //* ~exclude
     destructor Destroy; override;
     //* Set the worksheet to SheetIndex. This will change all source data to this sheet.
     procedure SourceSheet(SheetIndex: integer);
     //* How to handle empty cells in the chart source data.
     property EmptyCells: TChartEmptyCells read GetEmptyCells write SetEmptyCells;
     //* When True only visible cells are plotted.
     property PlotVisibleCellsOnly: boolean read GetPlotVisibleCellsOnly write SetPlotVisibleCellsOnly;
published
     //* Font settings for first default font.
     property DefaultTextA: TChartText read FDefaultTextA write FDefaultTextA;
     //* Font settings for second default font.
     property DefaultTextB: TChartText read FDefaultTextB write FDefaultTextB;
     //* The series of the chart.
     property Series: TXLSChartSeries read FSeries write FSeries;
     //* Settings for the charts plot area.
     property PlotArea: TChartPlotArea read FPlotArea write FPlotArea;
     //* Optional labels on the chart.
     property Labels: TChartLabels read FLabels write FLabels;
     //* Options for the chart.
     property SheetOptions: TChartSheetOptions read GetSheetOptions write SetSheetOptions;
     //* ~exclude
     property Rec: TChartRecord read FRecord;
     end;

//* List of TDrwChart objects.
type TDrwCharts = class(TCollection)
private
     function  GetDrwChart(Index: integer): TDrwChart;
protected
     FOwner: TPersistent;
     FDrawing: TEscherDrawing;
     FFileAdd: boolean;
     FFormulaHandler: TFormulaHandler;
     FFonts: TXc12Fonts;

     function  GetOwner: TPersistent; override;
public
     //* ~exclude
     constructor Create(AOwner: TPersistent; Drawing: TEscherDrawing; FormulaHandler: TFormulaHandler; Fonts: TXc12Fonts);
     //* Add a new TDrwChart
     //* ~param SheetIndex The index of the worksheet where the chart shall be.
     //* ~result The new TDrwChart object.
     function  Add(SheetIndex: integer): TDrwChart;
     //* Set the worksheet to Index for all charts. This will change all source data to this sheet.
     procedure SourceSheet(Index: integer);
     //* ~exclude
     function  AddClean(SheetIndex: integer): TDrwChart;
     //* ~exclude
     procedure AddFromFile(Shape: TShapeClientAnchor);
     //* ~exclude
     procedure Clear;

     property Items[Index: integer]: TDrwChart read GetDrwChart; default;
     end;

//* ~exclude
type TSheetChart = class(TCollectionItem)
private
     FSheetIndex: integer;
     FName: AxUCString;
     FRecord: TChartRecord;
     FDrawing: TEscherDrawing;
published
     constructor Create(Collection: TCollection); override;
     destructor Destroy; override;
     procedure AddDrawing(Drawing: TEscherDrawing);
     property SheetIndex: integer read FSheetIndex write FSheetIndex;
     property Name: AxUCString read FName write FName;
     property Drawing: TEscherDrawing read FDrawing;
     end;

//* ~exclude
type TSheetCharts = class(TCollection)
private
     FOwner: TPersistent;
     FDrawingGroup: TEscherGroup;
     FCurrIndex: integer;

     function GetSheetChart(Index: integer): TSheetChart;
protected
     function  GetOwner: TPersistent; override;
     function  Add: TSheetChart;
public
     constructor Create(AOwner: TPersistent; DrawingGroup: TEscherGroup);

     function  Find(const AName: AxUCString): integer;

     procedure LoadFromStream(Stream: TXLSStream; Name: AxUCString; PBuf: PByteArray; Fonts: TXc12Fonts; SheetIndex: integer);
     procedure SaveToStream(Index: integer; Stream: TXLSStream);

     property Items[Index: integer]: TSheetChart read GetSheetChart; default;
     property DrawingGroup: TEscherGroup read FDrawingGroup;
     property CurrIndex: integer read FCurrIndex;
     end;

implementation

procedure CheckRec(Rec: TChartRecord);
begin
  if Rec = Nil then
    raise XLSRWException.Create('Property has no data');
end;

{ TDrwCharts }

function TDrwCharts.Add(SheetIndex: integer): TDrwChart;
begin
  Result := TDrwChart(inherited Add);
  FFormulaHandler.ExternalNames.AddSelf(SheetIndex,1);
  Result.FRecord := TShapeOutsideMsoChart(Result.FShape.ExShape).Records;
  Result.FFormulaHandler := FFormulaHandler;
  Result.FRecord.ReadDefaultRecords(drdAll);
  Result.AssignRecords(Result.FRecord);
end;

function TDrwCharts.AddClean(SheetIndex: integer): TDrwChart;
begin
  Result := TDrwChart(inherited Add);
  FFormulaHandler.ExternalNames.AddSelf(SheetIndex,1);
  Result.FRecord := TShapeOutsideMsoChart(Result.FShape.ExShape).Records;
  Result.FFormulaHandler := FFormulaHandler;
end;

procedure TDrwCharts.AddFromFile(Shape: TShapeClientAnchor);
var
  Chart: TDrwChart;
begin
  FFileAdd := True;
  try
    Chart := TDrwChart(inherited Add);
    Chart.FRecord := TShapeOutsideMsoChart(Shape.ExShape).Records;
    Chart.FFormulaHandler := FFormulaHandler;
    Chart.FShape := Shape;
    Chart.AssignRecords(Chart.FRecord);
  finally
    FFileAdd := False;
  end;
end;

procedure TDrwCharts.Clear;
var
  i: integer;
begin
  for i := 0 to Count - 1 do
    FDrawing.DeleteShape(Items[i].FShape.SpId);
  inherited Clear;
end;

constructor TDrwCharts.Create(AOwner: TPersistent; Drawing: TEscherDrawing; FormulaHandler: TFormulaHandler; Fonts: TXc12Fonts);
begin
  inherited Create(TDrwChart);
  FOwner := AOwner;
  FDrawing := Drawing;
  FFormulaHandler := FormulaHandler;
  FFonts := Fonts;
end;

function TDrwCharts.GetDrwChart(Index: integer): TDrwChart;
begin
  Result := TDrwChart(inherited Items[Index]);
end;

function TDrwCharts.GetOwner: TPersistent;
begin
  Result := FOwner;
end;

procedure TDrwCharts.SourceSheet(Index: integer);
var
  i: integer;
begin
  for i := 0 to Count - 1 do
    Items[i].SourceSheet(Index);
end;

{ TDrwChart }

procedure TDrwChart.AssignRecords(Records: TChartRecord);
var
  i: integer;
begin
  i := 0;
  if (Records.Count > 0) and (Records[i].RecId = BIFFRECID_BOF8) then begin
    while i < Records.Count do begin
      if Records[i].RecId = CHARTRECID_CHART then begin
        FRecCHART := Records[i];
        Records := Records[i];
        Break;
      end;
      Inc(i);
    end;
  end;
  if Records = Nil then
    raise XLSRWException.Create('Can not find CHART record.');
  i := 0;
  while i < Records.Count do begin
    case Records[i].RecId of
      CHARTRECID_DEFAULTTEXT: begin
        Inc(i);
        if (i >= Records.Count) or (Records[i].RecId <> CHARTRECID_TEXT) then
          raise XLSRWException.Create('Excpected Chart record missing: TEXT');
        case PCRecDEFAULTTEXT(Records[i - 1].Data).Id of
          2: FDefaultTextA.Assign(Records[i]);
          3: FDefaultTextB.Assign(Records[i]);
        end;
      end;
      CHARTRECID_FRAME:
        FFrame.Assign(Records[i]);
      CHARTRECID_SERIES:
        FSeries.AddFromFile.AssignRecords(Records[i]);
      CHARTRECID_TEXT:
        FLabels.AddFromFile.AssignRecords(Records[i]);
      CHARTRECID_SHTPROPS:
        FRecSHTPROPS := Records[i];
      CHARTRECID_AXISPARENT:
        FPlotArea.Assign(Records[i]);
      CHARTRECID_FONTX: begin
        Inc(i);
      end;
    end;
    Inc(i);
  end;
end;

constructor TDrwChart.Create(Collection: TCollection);
begin
  FDefaultTextA := TChartText.Create;
  FDefaultTextB := TChartText.Create;
  FLabels := TChartLabels.Create(Self);
  FSeries := TXLSChartSeries.Create(Self,TDrwCharts(Collection).FFormulaHandler);
  FPlotArea := TChartPlotArea.Create;
  FFrame := TPaintFrame.Create;

  if not TDrwCharts(Collection).FFileAdd then
    FShape := TDrwCharts(Collection).FDrawing.AddChart;
  inherited Create(Collection);
  FName := 'Chart ' + IntToStr(ID);
end;

destructor TDrwChart.Destroy;
begin
  FLabels.Free;
  FSeries.Free;
  FDefaultTextA.Free;
  FDefaultTextB.Free;
  FPlotArea.Free;
  FFrame.Free;
//  TDrwCharts(Collection).FDrawing.DeleteShape(FShape.SpId);
  inherited;
end;

function TDrwChart.GetEmptyCells: TChartEmptyCells;
begin
  CheckRec(FRecSHTPROPS);
  Result := TChartEmptyCells(PCRecSHTPROPS(FRecSHTPROPS.Data).BlanksAs);
end;

function TDrwChart.GetPlotVisibleCellsOnly: boolean;
begin
  CheckRec(FRecSHTPROPS);
  Result := (PCRecSHTPROPS(FRecSHTPROPS.Data).Flags and $0002) = $0002;
end;

function TDrwChart.GetSheetOptions: TChartSheetOptions;
begin
  CheckRec(FRecSHTPROPS);
  Result := [];
  if (PCRecSHTPROPS(FRecSHTPROPS.Data).Flags and $0002) = $0002 then
    Result := Result + [csoOnlyVisibleCells];
  if (PCRecSHTPROPS(FRecSHTPROPS.Data).Flags and $0004) = $0004 then
    Result := Result + [csoDoNotSize];
end;

procedure TDrwChart.SetEmptyCells(const Value: TChartEmptyCells);
begin
  CheckRec(FRecSHTPROPS);
  PCRecSHTPROPS(FRecSHTPROPS.Data).BlanksAs := Integer(Value);
end;

procedure TDrwChart.SetPlotVisibleCellsOnly(const Value: boolean);
begin
  CheckRec(FRecSHTPROPS);
  if Value then
    PCRecSHTPROPS(FRecSHTPROPS.Data).Flags := PCRecSHTPROPS(FRecSHTPROPS.Data).Flags or $0002
  else
    PCRecSHTPROPS(FRecSHTPROPS.Data).Flags := PCRecSHTPROPS(FRecSHTPROPS.Data).Flags and not $0002;
end;

procedure TDrwChart.SetSheetOptions(const Value: TChartSheetOptions);
begin
  CheckRec(FRecSHTPROPS);
  PCRecSHTPROPS(FRecSHTPROPS.Data).Flags := PCRecSHTPROPS(FRecSHTPROPS.Data).Flags and not ($0002 + $0004);
  if csoOnlyVisibleCells in Value then
    PCRecSHTPROPS(FRecSHTPROPS.Data).Flags := PCRecSHTPROPS(FRecSHTPROPS.Data).Flags or $0002;
  if csoDoNotSize in Value then
    PCRecSHTPROPS(FRecSHTPROPS.Data).Flags := PCRecSHTPROPS(FRecSHTPROPS.Data).Flags or $0004;
end;

procedure TDrwChart.SourceSheet(SheetIndex: integer);
var
  ExtIndex: integer;

procedure DoSourceSheet(Records: TChartRecord; Index: integer);
var
  i: integer;
begin
  for i := 0 to Records.Count - 1 do begin
    if Records[i].Count > 0 then
      DoSourceSheet(Records[i],Index);
    if Records[i].RecId = CHARTRECID_AI then
      AdjustSheet(@PCRecAI(Records[i].Data).Formula,PCRecAI(Records[i].Data).FormulaSize,ExtIndex)
    else if Records[i].RecId = CHARTRECID_FONTX then
      TXc12Font(Records[i].Data).Use;
  end;
end;

begin
  ExtIndex := FFormulaHandler.ExternalNames.AddSelf(SheetIndex,1);
  DoSourceSheet(FRecord,Index);
end;

{ TChartText }

procedure TChartText.Assign(Rec: TChartRecord);
var
  i: integer;
begin
  inherited;
  FFont := Nil;
  for i := 0 to Rec.Count - 1 do begin
    if Rec[i].RecId = CHARTRECID_FONTX then begin
      FFont := TXc12Font(Rec[i].Data);
    end;
  end;
end;

function TChartText.GetColor: TXc12IndexColor;
begin
  Check;
  if PCRecTEXT(FRecord.Data).ColorIndex > Integer(High(TXc12IndexColor)) then
    Result := xcAutomatic
  else
    Result := TXc12IndexColor(PCRecTEXT(FRecord.Data).ColorIndex);
end;

procedure TChartText.SetColor(const Value: TXc12IndexColor);
begin
  Check;
  PCRecTEXT(FRecord.Data).ColorIndex := Word(Value);
  PCRecTEXT(FRecord.Data).Options1 := PCRecTEXT(FRecord.Data).Options1 and (not $0001);
  PCRecTEXT(FRecord.Data).Options1 := PCRecTEXT(FRecord.Data).Options1 and (not $0020);
end;

{ TChartSeries }

function TXLSChartSeries.Add: TXLSChartSerie;
const
  D: array[0..11] of byte = ($01,$00,$01,$00,$0A,$00,$0A,$00,$01,$00,$00,$00);
var
  i,Index: integer;
  Rec,Par: TChartRecord;
begin
  Index := -1;
  Par := Items[Count - 1].FRecord.Parent;
  for i := 0 to Par.Count - 1 do begin
    if Par[i].RecId = CHARTRECID_SERIES then
      Index := i;
  end;
  if Index < 0 then
    raise XLSRWException.Create('Can not find insertion point for serie');
  Rec := Par.InsertRecord(Index + 1,CHARTRECID_SERIES,SizeOf(TCRecSERIES));
  Move(D,Rec.Data^,SizeOf(TCRecSERIES));
  Rec.ReadDefaultRecords(drdSerie);

  Result := TXLSChartSerie.Create(Self);
  Result.FFormulaHandler := FFormulaHandler;
  Result.AssignRecords(Rec);

  Result.CalcValueCount;
end;

function TXLSChartSeries.AddFromFile: TXLSChartSerie;
begin
  Result := TXLSChartSerie.Create(Self);
  Result.FFormulaHandler := FFormulaHandler;
end;

constructor TXLSChartSeries.Create(AOwner: TPersistent; FormulaHandler: TFormulaHandler);
begin
  inherited Create(TXLSChartSerie);
  FOwner := AOwner;
  FFormulaHandler := FormulaHandler;
end;

procedure TXLSChartSeries.Delete(Index: integer);
var
  i: integer;
begin
  i := Items[Index].FRecord.Parent.IndexOf(Items[Index].FRecord);
  if i >= 0 then
    Items[Index].FRecord.Parent.Delete(i);
//  Items[Index].FRecord.Free;
  inherited Delete(Index);
end;

function TXLSChartSeries.GetChartSerie(Index: integer): TXLSChartSerie;
begin
  Result := TXLSChartSerie(inherited Items[Index]);
end;

function TXLSChartSeries.GetOwner: TPersistent;
begin
  Result := FOwner;
end;

{ TChartSerie }

procedure TXLSChartSerie.AssignRecords(Rec: TChartRecord);
var
  i: integer;
begin
  FRecord := Rec;
  FiRecAI_Values := -1;
  FiRecAI_Values2 := -1;
  FiRecAI_Category := -1;
  FiRecAI_SerieName := -1;
  PRecGeomtery := Nil;
  i := 0;
  while i < FRecord.Count do begin
    case FRecord[i].RecId of
      CHARTRECID_AI: begin
        case PCRecAI(FRecord[i].Data).LinkType of
          0: FiRecAI_SerieName := i;
          1: FiRecAI_Values := i;
          2: FiRecAI_Category := i;
          3: FiRecAI_Values2 := i;
        end;
      end;
      CHARTRECID_DATAFORMAT: begin
        if PCRecDATAFORMAT(FRecord[i].Data).PointNumber = $FFFF then begin
          FDefDatapoint.AssignRecords(FRecord[i]);
          if (FRecord[i].Count > 0) and (FRecord[i][0].RecId = CHARTRECID_GEOMETRY) then
            PRecGeomtery := FRecord[i][0];
        end
        else
          FDatapoints.AddFromFile.AssignRecords(FRecord[i]);
      end;
    end;
    Inc(i);
  end;
end;

constructor TXLSChartSerie.Create(Collection: TCollection);
begin
  inherited Create(Collection);
  FDataPoints := TDataPoints.Create(Self);
  FDefDatapoint := TDataPoint.Create;
end;

destructor TXLSChartSerie.Destroy;
begin
  FDataPoints.Free;
  FDefDatapoint.Free;
  inherited;
end;

procedure TXLSChartSerie.CalcValueCount;
var
  S: AxUCString;
  C1,R1,C2,R2: integer;
begin
  S := GetAIFormula(PCRecAI(FRecord[FiRecAI_Values].Data));
  if AreaStrToColRow(S,C1,R1,C2,R2) then
    PCRecSERIES(FRecord.Data).ValuesCount := (R2 - R1 + 1) * (C2 - C1 + 1);
end;

function TXLSChartSerie.GetAIFormula(AI: PCRecAI): AxUCString;
begin
  if AI.FormulaSize > 0 then
    Result := FFormulaHandler.DecodeFormula(@AI.Formula,AI.FormulaSize)
  else
    Result := '';
end;

function TXLSChartSerie.GetCategory: AxUCString;
begin
  if FiRecAI_Category < 0 then
    raise XLSRWException.Create('Property has no data');
  Result := GetAIFormula(PCRecAI(FRecord[FiRecAI_Category].Data));
end;
{
function TChartSerie.GetCategoriesCount: integer;
begin
  Result := PCRecSERIES(FRecord.Data).CategoriesCount;
end;
}
function TXLSChartSerie.GetGeomtery: TSerieItemGeomtery;
var
  i: integer;
begin
  Result := sigDefault;
  if PRecGeomtery = Nil then
    Exit;
  for i := 0 to Ord(sigLast) - 1 do begin
    if PCRecGEOMETRY(PRecGeomtery.Data).ItemType = Word(TSerieItemGeomtery(i)) then begin
      Result := TSerieItemGeomtery(i);
      Exit;
    end;
  end;
end;

function TXLSChartSerie.GetHasDefDataPoint: boolean;
begin
  Result := FDefDataPoint.FRecord.Count >= 3;
end;

function TXLSChartSerie.GetSerieName: AxUCString;
var
  i: integer;
begin
  if FiRecAI_SerieName < 0 then
    raise XLSRWException.Create('Property has no data');
  i := FiRecAI_SerieName;
  if PCRecAI(FRecord[i].Data).ReferenceType = 1 then begin
    Inc(i);
    if (i < FRecord.Count) and (FRecord[i].RecId = CHARTRECID_SERIESTEXT) then begin
      Result := ByteStrToWideString(@PCRecSERIESTEXT(FRecord[i].Data).Text,PCRecSERIESTEXT(FRecord[i].Data).Length);
      if Result <> '' then
        Result := '"' + Result + '"';
    end
    else
      Result := '';
  end
  else
    Result := GetAIFormula(PCRecAI(FRecord[i].Data));
end;

function TXLSChartSerie.GetValues: AxUCString;
begin
  if FiRecAI_Values < 0 then
    raise XLSRWException.Create('Property has no data');
  Result := GetAIFormula(PCRecAI(FRecord[FiRecAI_Values].Data));
end;

function TXLSChartSerie.GetValues2: AxUCString;
begin
  if FiRecAI_Values2 < 0 then
    raise XLSRWException.Create('Property has no data');
  Result := GetAIFormula(PCRecAI(FRecord[FiRecAI_Values2].Data));
end;

function TXLSChartSerie.GetValuesFmla: TDynByteArray;
var
  AI: PCRecAI;
begin
  if FiRecAI_Values < 0 then begin
    SetLength(Result,0);
    Exit;
  end;
  AI := PCRecAI(FRecord[FiRecAI_Values].Data);
  if AI.FormulaSize > 0 then begin
    SetLength(Result,AI.FormulaSize);
    Move(AI.Formula,Result[0],AI.FormulaSize);
  end
  else
    SetLength(Result,0)
end;

procedure TXLSChartSerie.SetAIFormula(AI: TChartRecord; Value: AxUCString);
var
  Sz,Delta: integer;
  P: PByteArray;
begin
  GetMem(P,MAXRECSZ_97);
  try
    Sz := FFormulaHandler.EncodeFormula(Value,0,P,MAXRECSZ_97);
    if Sz > 0 then begin
      // Ensure correct ref class
      if P[0] = $5B then
        P[0] := $3B;
    end;
    Delta := Sz - (AI.Length - SizeOf(TCRecAI));
    AI.Resize(Delta);
    Move(P^,PCRecAI(AI.Data).Formula,Sz);
    PCRecAI(AI.Data).FormulaSize := Sz;
    if Sz = 0 then
      PCRecAI(AI.Data).ReferenceType := 1
    else
      PCRecAI(AI.Data).ReferenceType := 2;
  finally
    FreeMem(P);
  end;
end;

procedure TXLSChartSerie.SetCategory(const Value: AxUCString);
begin
  if FiRecAI_Category < 0 then
    raise XLSRWException.Create('Property has no data');
  SetAIFormula(FRecord[FiRecAI_Category],Value);
end;

procedure TXLSChartSerie.SetGeomtery(const Value: TSerieItemGeomtery);
begin
  if PRecGeomtery = Nil then
    Exit;
  case Value of
    sigDefault        : PCRecGEOMETRY(PRecGeomtery.Data).ItemType := $0000;
    sigCylinder       : PCRecGEOMETRY(PRecGeomtery.Data).ItemType := $0001;
    sigPyramid        : PCRecGEOMETRY(PRecGeomtery.Data).ItemType := $0100;
    sigChoppedPyramid : PCRecGEOMETRY(PRecGeomtery.Data).ItemType := $0200;
    sigCone           : PCRecGEOMETRY(PRecGeomtery.Data).ItemType := $0101;
    sigChoppedCone    : PCRecGEOMETRY(PRecGeomtery.Data).ItemType := $0201;
  end;
end;

procedure TXLSChartSerie.SetHasDefDataPoint(const Value: boolean);
begin
  if Value = GetHasDefDataPoint then
    Exit;
  FDefDatapoint.FRecord.ReadDefaultRecords(drdDataformat);
  FDefDatapoint.AssignRecords(FDefDatapoint.FRecord);
end;

procedure TXLSChartSerie.SetSerieName(const Value: AxUCString);
var
  i: integer;
  S: AxUCString;
  L,Delta: integer;
  Rec: TChartRecord;
begin
  if FiRecAI_SerieName < 0 then
    raise XLSRWException.Create('Property has no data');
  if (Value <> '') and (Value[1] = '"') then begin
    S := Copy(Value,2,MAXINT);
    if Copy(S,Length(S),1) = '"' then
      S := Copy(S,1,Length(S) - 1);
    i := FiRecAI_SerieName + 1;
    if i < FRecord.Count then begin
      if FRecord[i].RecId <> CHARTRECID_SERIESTEXT then begin
        Rec := FRecord.InsertRecord(i,CHARTRECID_SERIESTEXT,SizeOf(TCRecSERIESTEXT));
        PCRecSERIESTEXT(Rec.Data).TextId := 0;
        if FiRecAI_Values > i then
          Inc(FiRecAI_Values);
        if FiRecAI_Values2 > i then
          Inc(FiRecAI_Values2);
        if FiRecAI_Category > i then
          Inc(FiRecAI_Category);
      end
      else
        Rec := FRecord[i];
      L := Length(S) * 2 + 1;
      Delta := L - (Rec.Length - SizeOf(TCRecSERIESTEXT));
      Rec.Resize(Delta);
      WideStringToByteStr(S,@PCRecSERIESTEXT(Rec.Data).Text);
      PCRecSERIESTEXT(Rec.Data).Length := Length(S);
      PCRecAI(FRecord[FiRecAI_SerieName].Data).ReferenceType := 1;
    end;
  end
  else
    SetAIFormula(FRecord[FiRecAI_SerieName],Value);
end;

procedure TXLSChartSerie.SetValues(const Value: AxUCString);
begin
  if FiRecAI_Values < 0 then
    raise XLSRWException.Create('Property has no data');
  SetAIFormula(FRecord[FiRecAI_Values],Value);
  CalcValueCount;
end;

procedure TXLSChartSerie.SetValues2(const Value: AxUCString);
begin
  if FiRecAI_Values2 < 0 then
    raise XLSRWException.Create('Property has no data');
  SetAIFormula(FRecord[FiRecAI_Values2],Value);
end;

{ TDataPoint }

procedure TDataPoint.AssignRecords(Rec: TChartRecord);
var
  i: integer;
begin
  if Rec.Count = 0 then
    raise XLSRWException.Create('DATAFORMAT has no child records');
  FRecord := Rec;
  for i := 0 to FRecord.Count - 1 do begin
    case FRecord[i].RecId of
      CHARTRECID_LINEFORMAT:     FLineFormat.Assign(FRecord[i]);
      CHARTRECID_AREAFORMAT:     FAreaFormat.Assign(FRecord[i]);
      CHARTRECID_PIEFORMAT:      FRecPIEFORMAT := FRecord[i];
      CHARTRECID_MARKERFORMAT:   FMarkerFormat.Assign(FRecord[i]);
      CHARTRECID_ATTACHEDLABEL:  FRecATTACHEDLABEL := FRecord[i];
    end;
  end;
end;

constructor TDataPoint.Create;
begin
  FLineFormat := TLineFormat.Create;
  FAreaFormat := TAreaFormat.Create;
  FMarkerFormat := TMarkerFormat.Create;
end;

destructor TDataPoint.Destroy;
begin
  FLineFormat.Free;
  FAreaFormat.Free;
  FMarkerFormat.Free;
  inherited;
end;

function TDataPoint.GetDataLablel: TChartDataLabels;
begin
  Result := TChartDataLabels(Byte(PCRecATTACHEDLABEL(FRecATTACHEDLABEL.Data).Flags));
end;

function TDataPoint.GetPieSliceDist: integer;
begin
  if FRecPIEFORMAT = Nil then
    raise XLSRWException.Create('Property has no data.');
  Result := PCRecPIEFORMAT(FRecPIEFORMAT.Data).Percent;
end;

function TDataPoint.GetPointIndex: integer;
begin
  Result := PCRecDATAFORMAT(FRecord.Data).PointNumber;
end;

procedure TDataPoint.SetDataLablel(const Value: TChartDataLabels);
begin
  PCRecATTACHEDLABEL(FRecATTACHEDLABEL.Data).Flags := Byte(Value);
end;

procedure TDataPoint.SetPieSliceDist(const Value: integer);
begin
  if FRecPIEFORMAT = Nil then
    raise XLSRWException.Create('Property has no data.');
  if (Value < 0) or (Value > 100) then
    raise XLSRWException.Create('Value out of range');
  PCRecPIEFORMAT(FRecPIEFORMAT.Data).Percent := Value;
end;

procedure TDataPoint.SetPointIndex(const Value: integer);
begin
  if Value < 0 then
    raise XLSRWException.Create('Point index out of range.');
  PCRecDATAFORMAT(FRecord.Data).PointNumber := Value;
end;

{ TDataPoints }

function TDataPoints.Add: TDataPoint;
var
  i: integer;
begin
  Result := TDataPoint.Create;
  inherited Add(Result);
  i := TXLSChartSerie(FOwner).FRecord.FindRecord(CHARTRECID_DATAFORMAT);
  if i < 0 then
    raise XLSRWException.Create('Can not find insertion point.');
  Result.FRecord := TXLSChartSerie(FOwner).FRecord.InsertRecord(i + 1,CHARTRECID_DATAFORMAT,SizeOf(TCRecDATAFORMAT));
  PCRecDATAFORMAT(Result.FRecord.Data).PointNumber := 0;
  PCRecDATAFORMAT(Result.FRecord.Data).SeriesIndex := TXLSChartSerie(FOwner).Index;
  PCRecDATAFORMAT(Result.FRecord.Data).SeriesNumber := TXLSChartSerie(FOwner).Index;
  PCRecDATAFORMAT(Result.FRecord.Data).Flags := 0;
  Result.FRecord.ReadDefaultRecords(drdDataformat);
  Result.AssignRecords(Result.FRecord);
end;

function TDataPoints.AddFromFile: TDataPoint;
begin
  Result := TDataPoint.Create;
  inherited Add(Result);
end;

constructor TDataPoints.Create(AOwner: TObject);
begin
  inherited Create;
  FOwner := AOwner;
end;

procedure TDataPoints.Delete(Index: integer);
var
  i: integer;
  Rec: TChartRecord;
begin
  Rec := TXLSChartSerie(FOwner).FRecord;
  for i := 0 to Rec.Count - 1 do begin
    if Rec[i] = Items[Index].FRecord then begin
      Rec.Delete(i);
      Break;
    end;
  end;
  inherited Delete(Index);
end;

function TDataPoints.GetItems(Index: integer): TDataPoint;
begin
  Result := TDataPoint(inherited Items[Index]);
end;

{ TLineFormat }

function TLineFormat.GetAutomatic: boolean;
begin
  Check;
  Result := (PCRecLINEFORMAT(FRecord.Data).Format and $0001) = $0001;
end;

function TLineFormat.GetLineColor: TXc12IndexColor;
begin
  Check;
  Result := IntToXColor(PCRecLINEFORMAT(FRecord.Data).ColorIndex);
end;

function TLineFormat.GetLineStyle: TChartLineStyle;
begin
  Check;
  Result := TChartLineStyle(PCRecLINEFORMAT(FRecord.Data).Pattern);
end;

function TLineFormat.GetLineWeight: TChartLineWeight;
begin
  Check;
  Result := TChartLineWeight(PCRecLINEFORMAT(FRecord.Data).Weight);
end;

procedure TLineFormat.SetAutomatic(const Value: boolean);
begin
  Check;
  if Value then
    PCRecLINEFORMAT(FRecord.Data).Format := PCRecLINEFORMAT(FRecord.Data).Format or $0001
  else
    PCRecLINEFORMAT(FRecord.Data).Format := PCRecLINEFORMAT(FRecord.Data).Format and not $0001;
end;

procedure TLineFormat.SetDefault;
begin
  PCRecLINEFORMAT(FRecord.Data).ColorRGB := $00000000;
  PCRecLINEFORMAT(FRecord.Data).Pattern := $0000;
  PCRecLINEFORMAT(FRecord.Data).Weight := $FFFF;
  PCRecLINEFORMAT(FRecord.Data).Format := $0009;
  PCRecLINEFORMAT(FRecord.Data).ColorIndex := $004D;
end;

procedure TLineFormat.SetLineColor(const Value: TXc12IndexColor);
begin
  Check;
  PCRecLINEFORMAT(FRecord.Data).ColorIndex := Word(Value);
  PCRecLINEFORMAT(FRecord.Data).ColorRGB := XColorToRGB(Value);
end;

procedure TLineFormat.SetLineStyle(const Value: TChartLineStyle);
begin
  Check;
  PCRecLINEFORMAT(FRecord.Data).Pattern := Word(Value);
end;

procedure TLineFormat.SetLineWeight(const Value: TChartLineWeight);
begin
  Check;
  PCRecLINEFORMAT(FRecord.Data).Weight := Word(Value);
end;

{ TAreaFormat }

function TAreaFormat.GetAutomatic: boolean;
begin
  Check;
  Result := (PCRecAREAFORMAT(FRecord.Data).Format and $0001) = $0001;
end;

function TAreaFormat.GetBackgroundColor: TXc12IndexColor;
begin
  Check;
  Result := IntToXColor(PCRecAREAFORMAT(FRecord.Data).ColorBgIndex);
end;

function TAreaFormat.GetForegroundColor: TXc12IndexColor;
begin
  Check;
  Result := IntToXColor(PCRecAREAFORMAT(FRecord.Data).ColorFgIndex);
end;

function TAreaFormat.GetPattern: TXc12FillPattern;
begin
  Check;
  Result := TXc12FillPattern(PCRecAREAFORMAT(FRecord.Data).Pattern);
end;

function TAreaFormat.GetSwapColorWhenNeg: boolean;
begin
  Check;
  Result := (PCRecAREAFORMAT(FRecord.Data).Format and $0002) = $0002;
end;

procedure TAreaFormat.SetAutomatic(const Value: boolean);
begin
  Check;
  if Value then
    PCRecAREAFORMAT(FRecord.Data).Format := PCRecAREAFORMAT(FRecord.Data).Format or $0001
  else
    PCRecAREAFORMAT(FRecord.Data).Format := PCRecAREAFORMAT(FRecord.Data).Format and not $0001;
end;

procedure TAreaFormat.SetBackgroundColor(const Value: TXc12IndexColor);
begin
  Check;
  PCRecAREAFORMAT(FRecord.Data).ColorBgIndex := Word(Value);
  PCRecAREAFORMAT(FRecord.Data).ColorBgRGB := XColorToRGB(Value);
end;

procedure TAreaFormat.SetDefault;
begin
  PCRecAREAFORMAT(FRecord.Data).ColorFgRGB := $00FFFFFF;
  PCRecAREAFORMAT(FRecord.Data).ColorBgRGB := $00000000;
  PCRecAREAFORMAT(FRecord.Data).Pattern := $0001;
  PCRecAREAFORMAT(FRecord.Data).Format := $0001;
  PCRecAREAFORMAT(FRecord.Data).ColorFgIndex := $004E;
  PCRecAREAFORMAT(FRecord.Data).ColorBgIndex := $004E;
end;

procedure TAreaFormat.SetForegroundColor(const Value: TXc12IndexColor);
begin
  Check;
  PCRecAREAFORMAT(FRecord.Data).ColorFgIndex := Word(Value);
  PCRecAREAFORMAT(FRecord.Data).ColorFgRGB := XColorToRGB(Value);
end;

procedure TAreaFormat.SetPattern(const Value: TXc12FillPattern);
begin
  Check;
  PCRecAREAFORMAT(FRecord.Data).Pattern := Word(Value);
end;

procedure TAreaFormat.SetSwapColorWhenNeg(const Value: boolean);
begin
  Check;
  if Value then
    PCRecAREAFORMAT(FRecord.Data).Format := PCRecAREAFORMAT(FRecord.Data).Format or $0002
  else
    PCRecAREAFORMAT(FRecord.Data).Format := PCRecAREAFORMAT(FRecord.Data).Format and not $0002;
end;

{ TRecordObject }

procedure TRecordObject.Assign(Rec: TChartRecord);
begin
  FRecord := Rec;
end;

procedure TRecordObject.Check;
begin
  if FRecord = Nil then
    raise XLSRWException.Create('Property has no data.');
end;

procedure TRecordObject.Check(Index: integer);
begin
  Check;
  if Index >= FRecord.Count then
    raise XLSRWException.Create('Property has no data.');
end;

function TRecordObject.GetValid: boolean;
begin
  Result := FRecord <> Nil;
end;

{ TMarkerFormat }

function TMarkerFormat.GetAutomatic: boolean;
begin
  Check;
  Result := (PCRecMARKERFORMAT(FRecord.Data).Format and $0001) = $0001;
end;

function TMarkerFormat.GetBackgroundColor: TXc12IndexColor;
begin
  Check;
  Result := IntToXColor(PCRecMARKERFORMAT(FRecord.Data).ColorBgIndex);
end;

function TMarkerFormat.GetForegroundColor: TXc12IndexColor;
begin
  Check;
  Result := IntToXColor(PCRecMARKERFORMAT(FRecord.Data).ColorFgIndex);
end;

function TMarkerFormat.GetMarkerStyle: TChartMarkerStyle;
begin
  Check;
  Result := TChartMarkerStyle(PCRecMARKERFORMAT(FRecord.Data).TypeOfMarker);
end;

function TMarkerFormat.GetSize: integer;
begin
  Check;
  Result := PCRecMARKERFORMAT(FRecord.Data).Size;
end;

procedure TMarkerFormat.SetAutomatic(const Value: boolean);
begin
  Check;
  if Value then
    PCRecMARKERFORMAT(FRecord.Data).Format := PCRecMARKERFORMAT(FRecord.Data).Format or $0001
  else
    PCRecMARKERFORMAT(FRecord.Data).Format := PCRecMARKERFORMAT(FRecord.Data).Format and not $0001;
end;

procedure TMarkerFormat.SetBackgroundColor(const Value: TXc12IndexColor);
begin
  Check;
  PCRecMARKERFORMAT(FRecord.Data).ColorBgIndex := Word(Value);
  PCRecMARKERFORMAT(FRecord.Data).ColorBgRGB := XColorToRGB(Value);
end;

procedure TMarkerFormat.SetForegroundColor(const Value: TXc12IndexColor);
begin
  Check;
  PCRecMARKERFORMAT(FRecord.Data).ColorFgIndex := Word(Value);
  PCRecMARKERFORMAT(FRecord.Data).ColorFgRGB := XColorToRGB(Value);
end;

procedure TMarkerFormat.SetMarkerStyle(const Value: TChartMarkerStyle);
begin
  Check;
  PCRecMARKERFORMAT(FRecord.Data).TypeOfMarker := Word(Value);
  PCRecMARKERFORMAT(FRecord.Data).Format := 0
{
  if Value = cmsNone then
    PCRecMARKERFORMAT(FRecord.Data).Format := 0
  else
    PCRecMARKERFORMAT(FRecord.Data).Format := 1;
}
end;

procedure TMarkerFormat.SetSize(const Value: integer);
begin
  Check;
  PCRecMARKERFORMAT(FRecord.Data).Size := Value;
end;

{ TTextFrame }

procedure TTextFrame.Assign(Rec: TChartRecord);
var
  i: integer;
begin
  inherited Assign(Rec);
  for i := 0 to FRecord.Count - 1 do begin
    case FRecord[i].RecId of
      CHARTRECID_LINEFORMAT: FLineFormat.Assign(FRecord[i]);
      CHARTRECID_AREAFORMAT: FAreaFormat.Assign(FRecord[i]);
    end;
  end;
end;

constructor TTextFrame.Create;
begin
  FLineFormat := TLineFormat.Create;
  FAreaFormat := TAreaFormat.Create;
end;

destructor TTextFrame.Destroy;
begin
  FLineFormat.Free;
  FAreaFormat.Free;
  inherited;
end;

function TTextFrame.GetAutoPosition: boolean;
begin
  Result := (PCRecFRAME(FRecord.Data).Flags and $0001) = $0001;
end;

function TTextFrame.GetAutoSize: boolean;
begin
  Result := (PCRecFRAME(FRecord.Data).Flags and $0002) = $0002;
end;

function TTextFrame.GetShadow: boolean;
begin
  Result := PCRecFRAME(FRecord.Data).FrameType = 4;
end;

procedure TTextFrame.SetAutoPosition(const Value: boolean);
begin
  if Value then
    PCRecFRAME(FRecord.Data).Flags := PCRecFRAME(FRecord.Data).Flags or $0001
  else
    PCRecFRAME(FRecord.Data).Flags := PCRecFRAME(FRecord.Data).Flags and not $0001
end;

procedure TTextFrame.SetAutoSize(const Value: boolean);
begin
  if Value then
    PCRecFRAME(FRecord.Data).Flags := PCRecFRAME(FRecord.Data).Flags or $0002
  else
    PCRecFRAME(FRecord.Data).Flags := PCRecFRAME(FRecord.Data).Flags and not $0002
end;

procedure TTextFrame.SetShadow(const Value: boolean);
begin
  if Value then
    PCRecFRAME(FRecord.Data).FrameType := 4
  else
    PCRecFRAME(FRecord.Data).FrameType := 0;
end;

{ TChartLabel }

procedure TChartLabel.AssignRecords(Rec: TChartRecord);
var
  i: integer;
begin
  FRecord := Rec;
  for i := 0 to FRecord.Count - 1 do begin
    case FRecord[i].RecId of
      CHARTRECID_POS:        FRecPOS := FRecord[i];
      CHARTRECID_AI:         FRecAI := FRecord[i];
      CHARTRECID_SERIESTEXT: FRecSERIESTEXT := FRecord[i];
      CHARTRECID_OBJECTLINK: FRecOBJECTLINK := FRecord[i];
    end;
  end;
end;

constructor TChartLabel.Create(Collection: TCollection);
begin
  inherited;

end;

destructor TChartLabel.Destroy;
begin

  inherited;
end;

function TChartLabel.GetDataPointIndex: integer;
begin
  Result := PCRecOBJECTLINK(FRecOBJECTLINK.Data).IndexData
end;

function TChartLabel.GetLabelType: TChartLabelType;
begin
  Result := TChartLabelType(PCRecOBJECTLINK(FRecOBJECTLINK.Data).LinkedTo);
end;

function TChartLabel.GetRotation: integer;
begin
  Result := PCRecTEXT(FRecord.Data).Rotation;
end;

function TChartLabel.GetSerieIndex: integer;
begin
  Result := PCRecOBJECTLINK(FRecOBJECTLINK.Data).IndexSeries;
end;

function TChartLabel.GetText: AxUCString;
begin
  if FRecSERIESTEXT = Nil then
    Result := ''
  else
    Result := ByteStrToWideString(@PCRecSERIESTEXT(FRecSERIESTEXT.Data).Text,PCRecSERIESTEXT(FRecSERIESTEXT.Data).Length);
end;

function TChartLabel.GetXOffset: integer;
begin
  Result := PCRecPOS(FRecPOS.Data).X1;
end;

function TChartLabel.GetYOffset: integer;
begin
  Result := PCRecPOS(FRecPOS.Data).Y1;
end;

procedure TChartLabel.SetDataPointIndex(const Value: integer);
begin
  PCRecOBJECTLINK(FRecOBJECTLINK.Data).IndexData := Value;
end;

procedure TChartLabel.SetDefault;
begin
  FillChar(FRecord.Data^,SizeOf(TCRecTEXT),#0);
  PCRecTEXT(FRecord.Data).HorizAlignment := $02;
  PCRecTEXT(FRecord.Data).VertAlignment := $02;
  PCRecTEXT(FRecord.Data).BkgMode := $0001;
  PCRecTEXT(FRecord.Data).Options1 := $0081;
  PCRecTEXT(FRecord.Data).Options2 := $3D20;
  PCRecTEXT(FRecord.Data).ColorIndex := 77;

  FRecPOS := FRecord.InsertRecord(-1,CHARTRECID_POS,SizeOf(TCRecPOS));
  FillChar(FRecPOS.Data^,SizeOf(TCRecPOS),#0);
  PCRecPOS(FRecPOS.Data).TopLt := 2;
  PCRecPOS(FRecPOS.Data).TopRt := 2;

  FRecAI := FRecord.InsertRecord(-1,CHARTRECID_AI,SizeOf(TCRecAI));
  FillChar(FRecAI.Data^,SizeOf(TCRecAI),#0);
  PCRecAI(FRecAI.Data).ReferenceType := $01;

  FRecOBJECTLINK := FRecord.InsertRecord(-1,CHARTRECID_OBJECTLINK,SizeOf(TCRecOBJECTLINK));
  FillChar(FRecOBJECTLINK.Data^,SizeOf(TCRecOBJECTLINK),#0);
  PCRecOBJECTLINK(FRecOBJECTLINK.Data).LinkedTo := $0001;
end;

procedure TChartLabel.SetLabelType(const Value: TChartLabelType);
begin
  PCRecOBJECTLINK(FRecOBJECTLINK.Data).LinkedTo := Word(Value);
end;

procedure TChartLabel.SetRotation(const Value: integer);
begin
  if (Value < -360) or (Value > 360) then
    raise XLSRWException.Create('Invalid rotation.');
  PCRecTEXT(FRecord.Data).Rotation := Value;
end;

procedure TChartLabel.SetSerieIndex(const Value: integer);
begin
  PCRecOBJECTLINK(FRecOBJECTLINK.Data).IndexSeries := Word(Value);
end;

procedure TChartLabel.SetText(const Value: AxUCString);
var
  i,L,Delta: integer;
begin
  if (Value <> '') and (FRecSERIESTEXT = Nil) then begin
    i := FRecord.FindRecord(CHARTRECID_AI);
    if i < 0 then
      raise XLSRWException.Create('Can not find insertion point.');
    FRecSERIESTEXT := FRecord.InsertRecord(i + 1,CHARTRECID_SERIESTEXT,SizeOf(TCRecSERIESTEXT));
    PCRecSERIESTEXT(FRecSERIESTEXT.Data).Length := 0;
  end
  else if (Value = '') and (FRecSERIESTEXT <> Nil) then begin
    i := FRecord.FindRecord(CHARTRECID_SERIESTEXT);
    if i >= 0 then
      FRecord.Delete(i);
    FRecSERIESTEXT := Nil;
    PCRecTEXT(FRecord.Data).Options1 := PCRecTEXT(FRecord).Options1 or $0010;
  end;
  if Value <> '' then begin
    PCRecTEXT(FRecord.Data).Options1 := PCRecTEXT(FRecord.Data).Options1 and not $0010;
    L := Length(Value) * 2 + 1;
    Delta := L - (FRecSERIESTEXT.Length - SizeOf(TCRecSERIESTEXT));
    FRecSERIESTEXT.Resize(Delta);
    WideStringToByteStr(Value,@PCRecSERIESTEXT(FRecSERIESTEXT.Data).Text);
    PCRecSERIESTEXT(FRecSERIESTEXT.Data).Length := Length(Value);
    PCRecAI(FRecAI.Data).ReferenceType := 1;
  end;
end;

procedure TChartLabel.SetXOffset(const Value: integer);
begin
  PCRecPOS(FRecPOS.Data).X1 := Longword(Value);
end;

procedure TChartLabel.SetYOffset(const Value: integer);
begin
  PCRecPOS(FRecPOS.Data).Y1 := Longword(Value);
end;

{ TChartLabels }

function TChartLabels.Add(LabelType: TChartLabelType): TChartLabel;
var
  i: integer;
begin
  Result := TChartLabel(inherited Add);

  if LabelType in [cltYAxis,cltXAxis,cltZAxis] then begin
    i := TDrwChart(FOwner).FRecCHART.FindRecord(CHARTRECID_AXISPARENT);
    if i >= 0 then begin
      Result.FRecord := TChartRecord.Create(TDrwChart(FOwner).FRecCHART[i],CHARTRECID_TEXT,SizeOf(TCRecTEXT));
      TDrwChart(FOwner).FRecCHART[i].Add(Result.FRecord);
    end;
  end
  else
    Result.FRecord := TDrwChart(FOwner).FRecCHART.InsertRecord(-1,CHARTRECID_TEXT,SizeOf(TCRecTEXT));
  Result.SetDefault;
  Result.SetLabelType(LabelType);
end;

function TChartLabels.AddFromFile: TChartLabel;
begin
  Result := TChartLabel(inherited Add);
end;

constructor TChartLabels.Create(AOwner: TPersistent);
begin
  inherited Create(TChartLabel);
  FOwner := AOwner;
end;

procedure TChartLabels.Delete(Index: integer);
var
  i: integer;
  Rec: TChartRecord;
begin
  Rec := TDrwChart(FOwner).FRecCHART;
  for i := 0 to Rec.Count - 1 do begin
    if Rec[i] = Items[Index].FRecord then begin
      Rec.Delete(i);
      Break;
    end;
  end;
  inherited Delete(Index);
end;

function TChartLabels.GetItems(Index: integer): TChartLabel;
begin
  Result := TChartLabel(inherited Items[Index]);
end;

function TChartLabels.GetOwner: TPersistent;
begin
  Result := FOwner;
end;

{ TTickFormat }

function TTickFormat.GetLabelColor: TXc12IndexColor;
begin
  Check;
  if PCRecTICK(FRecord.Data).ColorIndex > Integer(High(TXc12IndexColor)) then
    Result := xcAutomatic
  else
    Result := TXc12IndexColor(PCRecTICK(FRecord.Data).ColorIndex);
end;

function TTickFormat.GetLabelPosition: TTickLablePos;
begin
  Check;
  Result := TTickLablePos(PCRecTICK(FRecord.Data).LabelPos);
end;

function TTickFormat.GetMajorTick: TTickType;
begin
  Check;
  Result := TTickType(PCRecTICK(FRecord.Data).MajorMark);
end;

function TTickFormat.GetMinorTick: TTickType;
begin
  Check;
  Result := TTickType(PCRecTICK(FRecord.Data).MinorMark);
end;

function TTickFormat.GetTextRotation: integer;
begin
  Check;
  Result := PCRecTICK(FRecord.Data).Rotation;
end;

procedure TTickFormat.SetLabelColor(const Value: TXc12IndexColor);
begin
  Check;
  PCRecTICK(FRecord.Data).ColorIndex := Word(Value);
  PCRecTICK(FRecord.Data).ColorRGB := XColorToRGB(Value);
  if Value <> xcAutomatic then
    PCRecTICK(FRecord.Data).Flags := PCRecTICK(FRecord.Data).Flags and not $0001
  else
    PCRecTICK(FRecord.Data).Flags := PCRecTICK(FRecord.Data).Flags or $0001;
end;

procedure TTickFormat.SetLabelPosition(const Value: TTickLablePos);
begin
  Check;
  PCRecTICK(FRecord.Data).LabelPos := Word(Value);
end;

procedure TTickFormat.SetMajorTick(const Value: TTickType);
begin
  Check;
  PCRecTICK(FRecord.Data).MajorMark := Word(Value);
end;

procedure TTickFormat.SetMinorTick(const Value: TTickType);
begin
  Check;
  PCRecTICK(FRecord.Data).MinorMark := Word(Value);
end;

procedure TTickFormat.SetTextRotation(const Value: integer);
begin
  Check;
  PCRecTICK(FRecord.Data).Rotation := Value;
  if Value <> 0 then
    PCRecTICK(FRecord.Data).Flags := PCRecTICK(FRecord.Data).Flags and not $0020
  else
    PCRecTICK(FRecord.Data).Flags := PCRecTICK(FRecord.Data).Flags or $0020;
end;

{ TChartPlotArea }

procedure TChartPlotArea.Assign(Rec: TChartRecord);
var
  i: integer;
begin
  inherited Assign(Rec);
  for i := 0 to FRecord.Count - 1do begin
    case FRecord[i].RecId of
      CHARTRECID_AXIS: begin
        case PCRecAXIS(FRecord[i].Data).AxisType of
          0: FCategoryAxis.Assign(FRecord[i]);
          1: FValueAxis.Assign(FRecord[i]);
          2: FSerieAxis.Assign(FRecord[i]);
          else
            raise XLSRWException.CreateFmt('Unknown axis type %d',[PCRecAXIS(FRecord[i].Data).AxisType]);
        end;
      end;
      CHARTRECID_CHARTFORMAT: begin
        if FRecord[i].Count <= 0 then
          raise XLSRWException.Create('Chart style record is missing');
        case FRecord[i][0].RecId of
          CHARTRECID_AREA:
            FChartStyle := TChartStyleArea.Create;
          CHARTRECID_BAR:
            FChartStyle := TChartStyleBarColumn.Create;
          CHARTRECID_BOPPOP:
            FChartStyle := TChartStyleBopPop.Create;
          CHARTRECID_LINE:
            FChartStyle := TChartStyleLine.Create;
          CHARTRECID_PIE:
            FChartStyle := TChartStylePie.Create;
          CHARTRECID_RADAR:
            FChartStyle := TChartStyleRadar.Create;
          CHARTRECID_RADARAREA:
            FChartStyle := TChartStyleRadarArea.Create;
          CHARTRECID_SCATTER:
            FChartStyle := TChartStyleScatter.Create;
          CHARTRECID_SURFACE:
            FChartStyle := TChartStyleSurface.Create;
          else
            raise XLSRWException.CreateFmt('Unknown chart style %.4X',[FRecord[i][0].RecId]);
        end;
        FChartStyle.Assign(FRecord[i]);
      end;
      CHARTRECID_FRAME:
        FFrame.Assign(FRecord[i]);
    end;
  end;
end;

constructor TChartPlotArea.Create;
begin
  FCategoryAxis := TCategoryAxis.Create;
  FSerieAxis := TSerieAxis.Create;
  FValueAxis := TValueAxis.Create;
  FFrame := TPaintFrame.Create;
end;

destructor TChartPlotArea.Destroy;
begin
  FCategoryAxis.Free;
  FSerieAxis.Free;
  FValueAxis.Free;
  FFrame.Free;
  if FChartStyle <> Nil then
    FChartStyle.Free;
  inherited;
end;

function TChartPlotArea.GetChartType: TXLSChartType;
begin
  Result := FChartStyle.ChartType;
end;

function TChartPlotArea.GetHasLegend: boolean;
begin
  Result := FChartStyle.FLegend.Valid;
end;

function TChartPlotArea.GetIs3D: boolean;
begin
  Result := FChartStyle.FChart3D.Valid;
end;

procedure TChartPlotArea.SetChartType(const Value: TXLSChartType);
var
  Rec: TChartRecord;
begin
  if Value = xctNone then
    raise XLSRWException.Create('Invalid chart type');
  if Value = GetChartType then
    Exit;
  Rec := FChartStyle.FRecord;
  Rec.Clear;

  FChartStyle.FRecord := Nil;
  FChartStyle.Free;

  case Value of
    xctArea: begin
      FChartStyle := TChartStyleArea.Create;
      FChartStyle.FRecord := Rec;
      FChartStyle.FRecord.ReadDefaultRecords(drdStyleArea);
    end;
    xctBubble: begin
      FChartStyle := TChartStyleScatter.Create;
      FChartStyle.FRecord := Rec;
      FChartStyle.FRecord.ReadDefaultRecords(drdStyleScatter);
    end;
    xctColumn: begin
      FChartStyle := TChartStyleBarColumn.Create;
      FChartStyle.FRecord := Rec;
      FChartStyle.FRecord.ReadDefaultRecords(drdStyleBarColumn);
      TChartStyleBarColumn(FChartStyle).IsBar := False;
    end;
    xctLine: begin
      FChartStyle := TChartStyleLine.Create;
      FChartStyle.FRecord := Rec;
      FChartStyle.FRecord.ReadDefaultRecords(drdStyleLine);
    end;
    xctPie: begin
      FChartStyle := TChartStylePie.Create;
      FChartStyle.FRecord := Rec;
      FChartStyle.FRecord.ReadDefaultRecords(drdStylePie);
    end;
    xctRadar: begin
      FChartStyle := TChartStyleRadar.Create;
      FChartStyle.FRecord := Rec;
      FChartStyle.FRecord.ReadDefaultRecords(drdStyleRadar);
    end;
    xctScatter: begin
      FChartStyle := TChartStyleScatter.Create;
      FChartStyle.FRecord := Rec;
      FChartStyle.FRecord.ReadDefaultRecords(drdStyleScatter);
    end;
    xctSurface: begin
      FChartStyle := TChartStyleSurface.Create;
      FChartStyle.FRecord := Rec;
      FChartStyle.FRecord.ReadDefaultRecords(drdStyleSurface);
    end;
  end;
  FChartStyle.Assign(FChartStyle.FRecord);
end;

procedure TChartPlotArea.SetHasLegend(const Value: boolean);
const
  DefData: array[0..19] of byte = ($3C,$0E,$00,$00,$CE,$06,$00,$00,$50,$01,$00,$00,$5E,$01,$00,$00,$03,$01,$1F,$00);
begin
  if Value = HasLegend then
    Exit;
  if Value then begin
    FChartStyle.FLegend.FRecord := FChartStyle.FRecord.InsertRecord(-1,CHARTRECID_LEGEND,SizeOf(TCRecLEGEND_));
    Move(DefData,FChartStyle.FLegend.FRecord.Data^,SizeOf(TCRecLEGEND_));
    FChartStyle.FLegend.FRecord.ReadDefaultRecords(drdLegend);
  end
  else begin
    FChartStyle.FRecord.RemoveRecord(CHARTRECID_LEGEND);
    FChartStyle.FLegend.FRecord := Nil;
  end;
end;

procedure TChartPlotArea.SetIs3D(const Value: boolean);
const
  DefData: array[0..13] of byte = ($14,$00,$0F,$00,$1E,$00,$64,$00,$64,$00,$96,$00,$16,$00);
var
  Rec: TChartRecord;
begin
  if Value = GetIs3D then
    Exit;
  if Value then begin
    if FChartStyle.FRecord.Count < 2 then
      raise XLSRWException.Create('Invalid record');
    if FChartStyle.FRecord[1].RecId = CHARTRECID_CHARTFORMATLINK then
      Rec := FChartStyle.FRecord.InsertRecord(2,CHARTRECID_3D,SizeOf(TCRec3D))
    else
      Rec := FChartStyle.FRecord.InsertRecord(1,CHARTRECID_3D,SizeOf(TCRec3D));
    Move(DefData,Rec.Data^,SizeOf(TCRec3D));
  end
  else begin
    FChartStyle.FRecord.RemoveRecord(CHARTRECID_3D);
    FChartStyle.FChart3D.FRecord.Free;
    FChartStyle.FChart3D.FRecord := Nil;
  end;
end;

{ TChartAxis }

procedure TChartAxis.Assign(Rec: TChartRecord);
var
  i: integer;
begin
  inherited Assign(Rec);
  for i := 0 to FRecord.Count - 1 do begin
    case FRecord[i].RecId of
      CHARTRECID_TICK: FTick.Assign(FRecord[i]);
    end;
  end;
end;

constructor TChartAxis.Create;
begin
  FTick := TTickFormat.Create;
end;

destructor TChartAxis.Destroy;
begin
  FTick.Free;
  inherited;
end;

{ TAxisScaling }

function TCatSerAxisScaling.GetCrossingPoint: integer;
begin
  Check;
  Result := PCRecCATSERRANGE(FRecord.Data).CrossingPoint;
end;

function TCatSerAxisScaling.GetLabelsFrequency: integer;
begin
  Check;
  Result := PCRecCATSERRANGE(FRecord.Data).FreqLabels;
end;

function TCatSerAxisScaling.GetOptions: TCatSerAxisScalingOptions;
begin
  Check;
  Result := [];
  if (PCRecCATSERRANGE(FRecord.Data).Flags and $0001) = $0001 then
    Result := Result + [csasoCrossesMidCategory];
  if (PCRecCATSERRANGE(FRecord.Data).Flags and $0002) = $0002 then
    Result := Result + [csasoValueCrossesRight];
  if (PCRecCATSERRANGE(FRecord.Data).Flags and $0004) = $0004 then
    Result := Result + [csasoReverseCategories];
end;

function TCatSerAxisScaling.GetTickMarksFrequency: integer;
begin
  Check;
  Result := PCRecCATSERRANGE(FRecord.Data).FreqMarks;
end;

procedure TCatSerAxisScaling.SetCrossingPoint(const Value: integer);
begin
  Check;
  if (Value < 1) then
    raise XLSRWException.Create('Value out of range');
  PCRecCATSERRANGE(FRecord.Data).CrossingPoint := Word(Value);
end;

procedure TCatSerAxisScaling.SetLabelsFrequency(const Value: integer);
begin
  Check;
  if (Value < 1) then
    raise XLSRWException.Create('Value out of range');
  PCRecCATSERRANGE(FRecord.Data).FreqLabels := Word(Value);
end;

procedure TCatSerAxisScaling.SetOptions(const Value: TCatSerAxisScalingOptions);
begin
  Check;
  PCRecCATSERRANGE(FRecord.Data).Flags := PCRecCATSERRANGE(FRecord.Data).Flags and not ($0001 + $0002 + $0004);
  if csasoCrossesMidCategory in Value then
    PCRecCATSERRANGE(FRecord.Data).Flags := PCRecCATSERRANGE(FRecord.Data).Flags or $0001;
  if csasoValueCrossesRight in Value then
    PCRecCATSERRANGE(FRecord.Data).Flags := PCRecCATSERRANGE(FRecord.Data).Flags or $0002;
  if csasoReverseCategories in Value then
    PCRecCATSERRANGE(FRecord.Data).Flags := PCRecCATSERRANGE(FRecord.Data).Flags or $0004;
end;

procedure TCatSerAxisScaling.SetTickMarksFrequency(const Value: integer);
begin
  Check;
  if (Value < 1) then
    raise XLSRWException.Create('Value out of range');
  PCRecCATSERRANGE(FRecord.Data).FreqMarks := Word(Value);
end;

{ TChartXAxis }

procedure TCategoryAxis.Assign(Rec: TChartRecord);
var
  i: integer;
begin
  inherited Assign(Rec);
  for i := 0 to FRecord.Count - 1 do begin
    case FRecord[i].RecId of
      CHARTRECID_CATSERRANGE:
        FCatSerAxisScaling.Assign(FRecord[i]);
    end;
  end;
end;

constructor TCategoryAxis.Create;
begin
  inherited Create;
  FCatSerAxisScaling := TCatSerAxisScaling.Create;
end;

destructor TCategoryAxis.Destroy;
begin
  FCatSerAxisScaling.Free;
  inherited;
end;

{ TChartYAxis }

procedure TSerieAxis.Assign(Rec: TChartRecord);
var
  i: integer;
begin
  inherited Assign(Rec);
  for i := 0 to FRecord.Count - 1 do begin
    case FRecord[i].RecId of
      CHARTRECID_CATSERRANGE:
        FCatSerAxisScaling.Assign(FRecord[i]);
    end;
  end;
end;

constructor TSerieAxis.Create;
begin
  inherited Create;
  FCatSerAxisScaling := TCatSerAxisScaling.Create;
end;

destructor TSerieAxis.Destroy;
begin
  FCatSerAxisScaling.Free;
  inherited;
end;

{ TChartZAxis }

procedure TValueAxis.Assign(Rec: TChartRecord);
var
  i: integer;
begin
  inherited Assign(Rec);
  for i := 0 to FRecord.Count - 1 do begin
    case FRecord[i].RecId of
      CHARTRECID_VALUERANGE:
        FValueAxisScaling.Assign(FRecord[i]);
    end;
  end;
end;

constructor TValueAxis.Create;
begin
  inherited Create;
  FValueAxisScaling := TValueAxisScaling.Create;
end;

destructor TValueAxis.Destroy;
begin
  FValueAxisScaling.Free;
  inherited;
end;

{ TValueAxisScaling }

procedure TValueAxisScaling.CheckFlag(Flag: word; Value: double);
begin
  if Value = 0 then
    PCRecVALUERANGE(FRecord.Data).Flags := PCRecVALUERANGE(FRecord.Data).Flags or Flag
  else
    PCRecVALUERANGE(FRecord.Data).Flags := PCRecVALUERANGE(FRecord.Data).Flags and not Flag;
end;

function TValueAxisScaling.GetCategoryAxisCrossing: double;
begin
  Check;
  Result := PCRecVALUERANGE(FRecord.Data).AxisCrosses;
end;

function TValueAxisScaling.GetMajorInc: double;
begin
  Check;
  Result := PCRecVALUERANGE(FRecord.Data).MajorInc
end;

function TValueAxisScaling.GetMaxValue: double;
begin
  Check;
  Result := PCRecVALUERANGE(FRecord.Data).MaxValue
end;

function TValueAxisScaling.GetMinorInc: double;
begin
  Check;
  Result := PCRecVALUERANGE(FRecord.Data).MinorInc
end;

function TValueAxisScaling.GetMinValue: double;
begin
  Check;
  Result := PCRecVALUERANGE(FRecord.Data).MinValue
end;

function TValueAxisScaling.GetOptions: TValueAxisScalingOptions;
begin
  Check;
  Result := [];
  if (PCRecVALUERANGE(FRecord.Data).Flags and $0001) = $0001 then
    Result := Result + [vasoAutoMin];
  if (PCRecVALUERANGE(FRecord.Data).Flags and $0002) = $0002 then
    Result := Result + [vasoAutoMax];
  if (PCRecVALUERANGE(FRecord.Data).Flags and $0004) = $0004 then
    Result := Result + [vasoAutoMajor];
  if (PCRecVALUERANGE(FRecord.Data).Flags and $0008) = $0008 then
    Result := Result + [vasoAutoMinor];
  if (PCRecVALUERANGE(FRecord.Data).Flags and $0010) = $0010 then
    Result := Result + [vasoAutoCross];
  if (PCRecVALUERANGE(FRecord.Data).Flags and $0020) = $0020 then
    Result := Result + [vasoLogScale];
  if (PCRecVALUERANGE(FRecord.Data).Flags and $0040) = $0040 then
    Result := Result + [vasoReverse];
  if (PCRecVALUERANGE(FRecord.Data).Flags and $0080) = $0080 then
    Result := Result + [vasoCategoryAxisCrossing];
end;

procedure TValueAxisScaling.SetCategoryAxisCrossing(const Value: double);
begin
  Check;
  PCRecVALUERANGE(FRecord.Data).AxisCrosses := Value;
  CheckFlag($0010,Value);
end;

procedure TValueAxisScaling.SetMajorInc(const Value: double);
begin
  Check;
  PCRecVALUERANGE(FRecord.Data).MajorInc := Value;
  CheckFlag($0004,Value);
end;

procedure TValueAxisScaling.SetMaxValue(const Value: double);
begin
  Check;
  PCRecVALUERANGE(FRecord.Data).MaxValue := Value;
  CheckFlag($0002,Value);
end;

procedure TValueAxisScaling.SetMinorInc(const Value: double);
begin
  Check;
  PCRecVALUERANGE(FRecord.Data).MinorInc := Value;
  CheckFlag($0008,Value);
end;

procedure TValueAxisScaling.SetMinValue(const Value: double);
begin
  Check;
  PCRecVALUERANGE(FRecord.Data).MinValue := Value;
  CheckFlag($0001,Value);
end;

procedure TValueAxisScaling.SetOptions(const Value: TValueAxisScalingOptions);
begin
  Check;
  PCRecVALUERANGE(FRecord.Data).Flags := PCRecVALUERANGE(FRecord.Data).Flags and not ($0001 + $0002 + $0004 + $0008 + $0010 + $0020 + $0040 + $0080);
  if vasoAutoMin in Value then
    PCRecVALUERANGE(FRecord.Data).Flags := PCRecVALUERANGE(FRecord.Data).Flags or $0001;
  if vasoAutoMax in Value then
    PCRecVALUERANGE(FRecord.Data).Flags := PCRecVALUERANGE(FRecord.Data).Flags or $0002;
  if vasoAutoMajor in Value then
    PCRecVALUERANGE(FRecord.Data).Flags := PCRecVALUERANGE(FRecord.Data).Flags or $0004;
  if vasoAutoMinor in Value then
    PCRecVALUERANGE(FRecord.Data).Flags := PCRecVALUERANGE(FRecord.Data).Flags or $0008;
  if vasoAutoCross in Value then
    PCRecVALUERANGE(FRecord.Data).Flags := PCRecVALUERANGE(FRecord.Data).Flags or $0010;
  if vasoLogScale in Value then
    PCRecVALUERANGE(FRecord.Data).Flags := PCRecVALUERANGE(FRecord.Data).Flags or $0020;
  if vasoReverse in Value then
    PCRecVALUERANGE(FRecord.Data).Flags := PCRecVALUERANGE(FRecord.Data).Flags or $0040;
  if vasoCategoryAxisCrossing in Value then
    PCRecVALUERANGE(FRecord.Data).Flags := PCRecVALUERANGE(FRecord.Data).Flags or $0080;
end;

{ TChartStyle }

procedure TChartStyle.Assign(Rec: TChartRecord);
var
  i: integer;
begin
  inherited Assign(Rec);
  for i := 0 to FRecord.Count - 1 do begin
    case FRecord[i].RecId of
      CHARTRECID_3D:
        FChart3D.Assign(FRecord[i]);
      CHARTRECID_LEGEND:
        FLegend.Assign(FRecord[i]);
      CHARTRECID_DATAFORMAT: begin
        if PCRecDATAFORMAT(FRecord[i].Data).SeriesNumber > $F000 then begin
          FDefDatapoint.AssignRecords(FRecord[i]);
        end;
      end;
    end;
  end;
end;

function TChartStyle.ChartType: TXLSChartType;
begin
  Result := xctNone;
end;

constructor TChartStyle.Create;
begin
  FChart3D := TChart3D.Create;
  FLegend := TChartLegend.Create;
  FDefDatapoint := TDataPoint.Create;
end;

destructor TChartStyle.Destroy;
begin
  FChart3D.Free;
  FLegend.Free;
  FDefDatapoint.Free;
  inherited;
end;

function TChartStyle.GetHasDefDataPoint: boolean;
begin
  Result := FDefDataPoint.FRecord <> Nil;
end;

function TChartStyle.GetVaryColors: boolean;
begin
  Check;
  Result := (PCRecCHARTFORMAT(FRecord.Data).Flags and $0001) = $0001;
end;

procedure TChartStyle.SetHasDefDataPoint(const Value: boolean);
const
  D: array[0..7] of byte = ($00,$00,$00,$00,$FD,$FF,$00,$00);
begin
  if Value = GetHasDefDataPoint then
    Exit;
  FDefDatapoint.FRecord := FRecord.InsertRecord(-1,CHARTRECID_DATAFORMAT,SizeOf(TCRecDATAFORMAT));
  Move(D,FDefDatapoint.FRecord.Data^,SizeOf(TCRecDATAFORMAT));
  FDefDatapoint.FRecord.ReadDefaultRecords(drdDataformat);
  FDefDatapoint.AssignRecords(FDefDatapoint.FRecord);

{
  FDefDatapoint.FRecord.ReadDefaultRecords(drdDataformat);
  FDefDatapoint.AssignRecords(FDefDatapoint.FRecord);
}
end;

procedure TChartStyle.SetVaryColors(const Value: boolean);
begin
  Check;
  if Value then
    PCRecCHARTFORMAT(FRecord.Data).Flags := PCRecCHARTFORMAT(FRecord.Data).Flags or $0001
  else
    PCRecCHARTFORMAT(FRecord.Data).Flags := PCRecCHARTFORMAT(FRecord.Data).Flags and not $0001;
end;

{ TChartStyleBarColumn }

procedure TChartStyleBarColumn.Assign(Rec: TChartRecord);
begin
  inherited Assign(rec);
//  PCRecBAR(FRecord[0].Data).SpaceBars := 75;
//  PCRecBAR(FRecord[0].Data).SpaceCategories := 5;
end;

function TChartStyleBarColumn.ChartType: TXLSChartType;
begin
  Result := xctColumn;
end;

constructor TChartStyleBarColumn.Create;
begin
  inherited Create;
end;

destructor TChartStyleBarColumn.Destroy;
begin

  inherited;
end;

function TChartStyleBarColumn.GetIsBar: boolean;
begin
  Check(0);
  Result := (PCRecBAR(FRecord[0].Data).Flags and $0001) = $0001;
end;

function TChartStyleBarColumn.GetStacked: boolean;
begin
  Check(0);
  Result := (PCRecBAR(FRecord[0].Data).Flags and $0002) = $0002;
end;

procedure TChartStyleBarColumn.SetIsBar(const Value: boolean);
begin
  Check(0);
  PCRecBAR(FRecord[0].Data).Flags := PCRecBAR(FRecord[0].Data).Flags and not $0001;
  if Value then
    PCRecBAR(FRecord[0].Data).Flags := PCRecBAR(FRecord[0].Data).Flags or $0001;
end;

procedure TChartStyleBarColumn.SetStacked(const Value: boolean);
begin
  Check(0);
  PCRecBAR(FRecord[0].Data).Flags := PCRecBAR(FRecord[0].Data).Flags and not $0002;
  if Value then
    PCRecBAR(FRecord[0].Data).Flags := PCRecBAR(FRecord[0].Data).Flags or $0002;
end;

{ TChartStyleSurface }

procedure TChartStyleSurface.Assign(Rec: TChartRecord);
begin
  inherited;

end;

constructor TChartStyleSurface.Create;
begin
  inherited Create;
end;

destructor TChartStyleSurface.Destroy;
begin

  inherited;
end;

{ TChartStyleArea }

procedure TChartStyleArea.Assign(Rec: TChartRecord);
begin
  inherited;

end;

function TChartStyleArea.ChartType: TXLSChartType;
begin
  Result := xctArea;
end;

constructor TChartStyleArea.Create;
begin
  inherited Create;
end;

destructor TChartStyleArea.Destroy;
begin

  inherited;
end;

{ TChartStyleScatter }

procedure TChartStyleScatter.Assign(Rec: TChartRecord);
begin
  inherited;

end;

constructor TChartStyleScatter.Create;
begin
  inherited Create;
end;

destructor TChartStyleScatter.Destroy;
begin

  inherited;
end;

function TChartStyleScatter.GetBubbleIsArea: boolean;
begin
  Check(0);
  Result := PCRecSCATTER(FRecord[0].Data).BubbleSizeIs = $0001;
end;

function TChartStyleScatter.GetBubblePercent: integer;
begin
  Check(0);
  Result := PCRecSCATTER(FRecord[0].Data).BubblePercent;
end;

function TChartStyleScatter.GetIsBubble: boolean;
begin
  Check(0);
  Result := (PCRecSCATTER(FRecord[0].Data).Flags and $0001) = $0001;
end;

function TChartStyleScatter.GetNegativeBubble: boolean;
begin
  Check(0);
  Result := (PCRecSCATTER(FRecord[0].Data).Flags and $0002) = $0002;
end;

function TChartStyleScatter.GetShadow: boolean;
begin
  Check(0);
  Result := (PCRecSCATTER(FRecord[0].Data).Flags and $0004) = $0004;
end;

procedure TChartStyleScatter.SetBubbleIsArea(const Value: boolean);
begin
  Check(0);
  if Value then
    PCRecSCATTER(FRecord[0].Data).BubbleSizeIs := $0001
  else
    PCRecSCATTER(FRecord[0].Data).BubbleSizeIs := $0002;
end;

procedure TChartStyleScatter.SetBubblePercent(const Value: integer);
begin
  Check(0);
  if (Value < 0) or (Value > 100) then
    raise XLSRWException.Create('Value out of range');
  PCRecSCATTER(FRecord[0].Data).BubblePercent := Value;
end;

procedure TChartStyleScatter.SetIsBubble(const Value: boolean);
begin
  Check(0);
  if Value then
    PCRecSCATTER(FRecord[0].Data).Flags := PCRecSCATTER(FRecord[0].Data).Flags or $0001
  else
    PCRecSCATTER(FRecord[0].Data).Flags := PCRecSCATTER(FRecord[0].Data).Flags and not $0001;
end;

procedure TChartStyleScatter.SetNegativeBubble(const Value: boolean);
begin
  Check(0);
  if Value then
    PCRecSCATTER(FRecord[0].Data).Flags := PCRecSCATTER(FRecord[0].Data).Flags or $0002
  else
    PCRecSCATTER(FRecord[0].Data).Flags := PCRecSCATTER(FRecord[0].Data).Flags and not $0002;
end;

procedure TChartStyleScatter.SetShadow(const Value: boolean);
begin
  Check(0);
  if Value then
    PCRecSCATTER(FRecord[0].Data).Flags := PCRecSCATTER(FRecord[0].Data).Flags or $0004
  else
    PCRecSCATTER(FRecord[0].Data).Flags := PCRecSCATTER(FRecord[0].Data).Flags and not $0004;
end;

{ TChartStylePie }

procedure TChartStylePie.Assign(Rec: TChartRecord);
begin
  inherited;

end;

constructor TChartStylePie.Create;
begin
  inherited Create;
end;

destructor TChartStylePie.Destroy;
begin

  inherited;
end;

function TChartStylePie.GetAngle: integer;
begin
  Check(0);
  Result := PCRecPIE(FRecord[0].Data).Angle;
end;

function TChartStylePie.GetDonutHole: integer;
begin
  Check(0);
  Result := PCRecPIE(FRecord[0].Data).Donut;
end;

function TChartStylePie.GetLeaderLines: boolean;
begin
  Check(0);
  Result := (PCRecPIE(FRecord[0].Data).Flags and $0002) = $0002;
end;

function TChartStylePie.GetShadow: boolean;
begin
  Check(0);
  Result := (PCRecPIE(FRecord[0].Data).Flags and $0001) = $0001;
end;

procedure TChartStylePie.SetAngle(const Value: integer);
begin
  Check(0);
  if (Value < 0) or (Value > 360) then
    raise XLSRWException.Create('Value out of range');
  PCRecPIE(FRecord[0].Data).Angle := Value;
end;

procedure TChartStylePie.SetDonutHole(const Value: integer);
begin
  Check;
  if (Value < 0) or (Value > 100) then
    raise XLSRWException.Create('Value out of range');
  PCRecPIE(FRecord[0].Data).Donut := Value;
end;

procedure TChartStylePie.SetLeaderLines(const Value: boolean);
begin
  Check(0);
  if Value then
    PCRecPIE(FRecord[0].Data).Flags := PCRecPIE(FRecord[0].Data).Flags or $0002
  else
    PCRecPIE(FRecord[0].Data).Flags := PCRecPIE(FRecord[0].Data).Flags and not $0002;
end;

procedure TChartStylePie.SetShadow(const Value: boolean);
begin
  Check(0);
  if Value then
    PCRecPIE(FRecord[0].Data).Flags := PCRecPIE(FRecord[0].Data).Flags or $0001
  else
    PCRecPIE(FRecord[0].Data).Flags := PCRecPIE(FRecord[0].Data).Flags and not $0001;
end;

{ TChartStyleLine }

procedure TChartStyleLine.Assign(Rec: TChartRecord);
begin
  inherited;

end;

constructor TChartStyleLine.Create;
begin
  inherited Create;
end;

destructor TChartStyleLine.Destroy;
begin

  inherited;
end;

function TChartStyleLine.GetCategoryAsPercent: boolean;
begin
  Check(0);
  Result := (PCRecLINE(FRecord[0].Data).Flags and $0002) = $0002;
end;

function TChartStyleLine.GetShadow: boolean;
begin
  Check(0);
  Result := (PCRecLINE(FRecord[0].Data).Flags and $0004) = $0004;
end;

function TChartStyleLine.GetStacked: boolean;
begin
  Check(0);
  Result := (PCRecLINE(FRecord[0].Data).Flags and $0001) = $0001;
end;

procedure TChartStyleLine.SetCategoryAsPercent(const Value: boolean);
begin
  Check(0);
  if Value then
    PCRecLINE(FRecord[0].Data).Flags := PCRecLINE(FRecord[0].Data).Flags or $0002
  else
    PCRecLINE(FRecord[0].Data).Flags := PCRecLINE(FRecord[0].Data).Flags and not $0002;
end;

procedure TChartStyleLine.SetShadow(const Value: boolean);
begin
  Check(0);
  if Value then
    PCRecLINE(FRecord[0].Data).Flags := PCRecLINE(FRecord[0].Data).Flags or $0004
  else
    PCRecLINE(FRecord[0].Data).Flags := PCRecLINE(FRecord[0].Data).Flags and not $0004;
end;

procedure TChartStyleLine.SetStacked(const Value: boolean);
begin
  Check(0);
  if Value then
    PCRecLINE(FRecord[0].Data).Flags := PCRecLINE(FRecord[0].Data).Flags or $0001
  else
    PCRecLINE(FRecord[0].Data).Flags := PCRecLINE(FRecord[0].Data).Flags and not $0001;
end;

{ TChartStyleBopPop }

procedure TChartStyleBopPop.Assign(Rec: TChartRecord);
begin
  inherited;

end;

constructor TChartStyleBopPop.Create;
begin
  inherited Create;
end;

destructor TChartStyleBopPop.Destroy;
begin

  inherited;
end;

{ TChartStyleRadar }

procedure TChartStyleRadar.Assign(Rec: TChartRecord);
begin
  inherited;

end;

constructor TChartStyleRadar.Create;
begin
  inherited Create;
end;

destructor TChartStyleRadar.Destroy;
begin

  inherited;
end;

function TChartStyleRadar.GetAxisLabels: boolean;
begin
  Check(0);
  Result := (PCRecRADAR(FRecord[0].Data).Flags and $0001) = $0001;
end;

function TChartStyleRadar.GetShadow: boolean;
begin
  Check(0);
  Result := (PCRecRADAR(FRecord[0].Data).Flags and $0002) = $0002;
end;

procedure TChartStyleRadar.SetAxisLabels(const Value: boolean);
begin
  Check(0);
  if Value then
    PCRecRADAR(FRecord[0].Data).Flags := PCRecRADAR(FRecord[0].Data).Flags or $0001
  else
    PCRecRADAR(FRecord[0].Data).Flags := PCRecRADAR(FRecord[0].Data).Flags and not $0001;
end;

procedure TChartStyleRadar.SetShadow(const Value: boolean);
begin
  Check(0);
  if Value then
    PCRecRADAR(FRecord[0].Data).Flags := PCRecRADAR(FRecord[0].Data).Flags or $0002
  else
    PCRecRADAR(FRecord[0].Data).Flags := PCRecRADAR(FRecord[0].Data).Flags and not $0002;
end;

{ TChartStyleRadarArea }

procedure TChartStyleRadarArea.Assign(Rec: TChartRecord);
begin
  inherited;

end;

constructor TChartStyleRadarArea.Create;
begin
  inherited Create;
end;

destructor TChartStyleRadarArea.Destroy;
begin

  inherited;
end;

function TChartStyleRadarArea.GetAxisLabels: boolean;
begin
  Check(0);
  Result := (PCRecRADARAREA(FRecord[0].Data).Flags and $0001) = $0001;
end;

function TChartStyleRadarArea.GetShadow: boolean;
begin
  Check(0);
  Result := (PCRecRADARAREA(FRecord[0].Data).Flags and $0002) = $0002;
end;

procedure TChartStyleRadarArea.SetAxisLabels(const Value: boolean);
begin
  Check(0);
  if Value then
    PCRecRADARAREA(FRecord[0].Data).Flags := PCRecRADARAREA(FRecord[0].Data).Flags or $0001
  else
    PCRecRADARAREA(FRecord[0].Data).Flags := PCRecRADARAREA(FRecord[0].Data).Flags and not $0001;
end;

procedure TChartStyleRadarArea.SetShadow(const Value: boolean);
begin
  Check(0);
  if Value then
    PCRecRADARAREA(FRecord[0].Data).Flags := PCRecRADARAREA(FRecord[0].Data).Flags or $0002
  else
    PCRecRADARAREA(FRecord[0].Data).Flags := PCRecRADARAREA(FRecord[0].Data).Flags and not $0002;
end;

{ TChart3D }

function TChart3D.GetDistance: integer;
begin
  Result := PCRec3D(FRecord.Data).Distance;
end;

function TChart3D.GetElevation: integer;
begin
  Result := PCRec3D(FRecord.Data).Elevation;
end;

function TChart3D.GetOptions: TChart3DOptions;
begin
  Result := [];
  if (PCRec3D(FRecord.Data).Flags and $0001) = $0001 then
    Result := Result + [c3doPerspectiveTrans];
  if (PCRec3D(FRecord.Data).Flags and $0002) = $0002 then
    Result := Result + [c3doClusterdCols];
  if (PCRec3D(FRecord.Data).Flags and $0004) = $0004 then
    Result := Result + [c3doAutoScaling];
  if (PCRec3D(FRecord.Data).Flags and $0020) = $0020 then
    Result := Result + [c3do2dWalls];
end;

function TChart3D.GetRelDepth: integer;
begin
  Result := PCRec3D(FRecord.Data).Depth;
end;

function TChart3D.GetRelHeight: integer;
begin
  Result := PCRec3D(FRecord.Data).PlotHeight;
end;

function TChart3D.GetRotation: integer;
begin
  Result := PCRec3D(FRecord.Data).Rotation;
end;

function TChart3D.GetSeriesSpace: integer;
begin
  Result := PCRec3D(FRecord.Data).Space;
end;

procedure TChart3D.SetDistance(const Value: integer);
begin
  if (Value < 0) or (Value > 100) then
    raise XLSRWException.Create('Invalid value');
  PCRec3D(FRecord.Data).Distance := Value;
end;

procedure TChart3D.SetElevation(const Value: integer);
begin
  if (Value < -90) or (Value > 90) then
    raise XLSRWException.Create('Invalid value');
  PCRec3D(FRecord.Data).Elevation := Value;
end;

procedure TChart3D.SetOptions(const Value: TChart3DOptions);
begin
  PCRec3D(FRecord.Data).Flags := PCRec3D(FRecord.Data).Flags and not ($0001 + $0002 + $0004 + $0020);
  if c3doPerspectiveTrans in Value then
    PCRec3D(FRecord.Data).Flags := PCRec3D(FRecord.Data).Flags + $0001;
  if c3doClusterdCols in Value then
    PCRec3D(FRecord.Data).Flags := PCRec3D(FRecord.Data).Flags + $0002;
  if c3doAutoScaling in Value then
    PCRec3D(FRecord.Data).Flags := PCRec3D(FRecord.Data).Flags + $0004;
  if c3do2dWalls in Value then
    PCRec3D(FRecord.Data).Flags := PCRec3D(FRecord.Data).Flags + $0020;
end;

procedure TChart3D.SetRelDepth(const Value: integer);
begin
  PCRec3D(FRecord.Data).Depth := Value;
end;

procedure TChart3D.SetRelHeight(const Value: integer);
begin
  PCRec3D(FRecord.Data).PlotHeight := Value;
end;

procedure TChart3D.SetRotation(const Value: integer);
begin
  if (Value < 0) or (Value > 360) then
    raise XLSRWException.Create('Invalid value');
  PCRec3D(FRecord.Data).Rotation := Value;
end;

procedure TChart3D.SetSeriesSpace(const Value: integer);
begin
  PCRec3D(FRecord.Data).Space := Value;
end;

{ TPaintFrame }

procedure TPaintFrame.Assign(Rec: TChartRecord);
var
  i: integer;
begin
  inherited Assign(Rec);
  i := 0;
  while i < FRecord.Count do begin
    case FRecord[i].RecId of
      CHARTRECID_LINEFORMAT:
        FLineFormat.Assign(FRecord[i]);
      CHARTRECID_AREAFORMAT:
        FAreaFormat.Assign(FRecord[i]);
      CHARTRECID_GELFRAME:
        FGelFrame.Assign(FRecord[i]);
    end;
    Inc(i);
  end;
end;

constructor TPaintFrame.Create;
begin
  inherited Create;
  FLineFormat := TLineFormat.Create;
  FAreaFormat := TAreaFormat.Create;
  FGelFrame := TGelFrame.Create;
end;

destructor TPaintFrame.Destroy;
begin
  FLineFormat.Free;
  FAreaFormat.Free;
  FGelFrame.Free;
  inherited;
end;

function TPaintFrame.GetHasFillEffects: boolean;
begin
  Result := FGelFrame.FRecord <> Nil;
end;

procedure TPaintFrame.SetDefault(Parent: TChartRecord);
begin
  FRecord := Parent.InsertRecord(-1,CHARTRECID_FRAME,SizeOf(TCRecFRAME));
  PCRecFRAME(FRecord.Data).FrameType := $0000;
  PCRecFRAME(FRecord.Data).Flags := $0003;
  FLineformat.FRecord := FRecord.InsertRecord(-1,CHARTRECID_LINEFORMAT,SizeOf(TCRecLINEFORMAT));
  FLineformat.SetDefault;
  FAreaformat.FRecord := FRecord.InsertRecord(-1,CHARTRECID_AREAFORMAT,SizeOf(TCRecAREAFORMAT));
  FAreaformat.SetDefault;
end;

procedure TPaintFrame.SetHasFillEffects(const Value: boolean);
begin
  if Value = HasFillEffects then
    Exit;
  if Value then begin
    FGelFrame.FRecord := FRecord.InsertRecord(-1,CHARTRECID_GELFRAME,0,True);
    FGelFrame.SetDefault;
  end
  else begin
    FRecord.RemoveRecord(CHARTRECID_GELFRAME);
    FGelFrame.FRecord.Free;
    FGelFrame.FRecord := Nil;
  end;
end;

{ TChartLegend }

procedure TChartLegend.Assign(Rec: TChartRecord);
var
  i: integer;
begin
  inherited Assign(Rec);
  for i := 0 to FRecord.Count - 1 do begin
    case FRecord[i].RecId of
      CHARTRECID_FRAME:
        FFrame.Assign(FRecord[i]);
    end;
  end;
end;

constructor TChartLegend.Create;
begin
  FFrame := TPaintFrame.Create;
end;

destructor TChartLegend.Destroy;
begin
  FFrame.Free;
  inherited;
end;

function TChartLegend.GetHasFrame: boolean;
begin
  Result := FFrame.FRecord <> Nil;
end;

function TChartLegend.GetLegendSpacing: TChartLegendSpacing;
begin
  Check;
  Result := TChartLegendSpacing(PCRecLEGEND_(FRecord.Data).Spacing);
end;

function TChartLegend.GetLegendType: TChartLegendType;
begin
  Check;
  Result := TChartLegendType(PCRecLEGEND_(FRecord.Data).LegendType);
end;

function TChartLegend.GetVertical: boolean;
begin
  Check;
  Result := (PCRecLEGEND_(FRecord.Data).Flags and $0010) = $0010;
end;

procedure TChartLegend.SetHasFrame(const Value: boolean);
begin
  if Value = GetHasFrame then
    Exit;
  if Value then
    FFrame.SetDefault(FRecord)
  else begin
    FFrame.FRecord.Free;
    FFrame.FRecord := Nil;
  end;
end;

procedure TChartLegend.SetLegendSpacing(const Value: TChartLegendSpacing);
begin
  Check;
  PCRecLEGEND_(FRecord.Data).Spacing := Word(Value);
end;

procedure TChartLegend.SetLegendType(const Value: TChartLegendType);
begin
  Check;
  PCRecLEGEND_(FRecord.Data).LegendType := Word(Value);
end;

procedure TChartLegend.SetVertical(const Value: boolean);
begin
  if Value then
    PCRecLEGEND_(FRecord.Data).Flags := PCRecLEGEND_(FRecord.Data).Flags or $0010
  else
    PCRecLEGEND_(FRecord.Data).Flags := PCRecLEGEND_(FRecord.Data).Flags and not $0010;
end;

{ TGelFrame }

procedure TGelFrame.Assign(Rec: TChartRecord);
var
  P: PByteArray;
  PHeader: PMSOHeader;
  L: integer;
begin
  inherited Assign(Rec);

  if (FRecord is TChartRecordUpdate) then
    TChartRecordUpdate(FRecord).OnUpdate := UpdateRecord;

  FOPT.Clear;
  P := Rec.Data;
  L := Rec.Length;
  if PMSOHeader(P).FBT <> MSO_OPT then
    raise XLSRWException.Create('First MSO record in Gelframe is not OPT');
  while L > 0 do begin
    PHeader := PMSOHeader(P);
    P := PByteArray(NativeInt(P) + SizeOf(TMSOHeader));
    if PHeader.FBT = MSO_OPT then
      FOPT.LoadFromBuffer(PHeader.Length,P,(PHeader.VerInst shr 4) and $0FFF);
    Dec(L,PHeader.Length + SizeOf(TMSOHeader));
    P := PByteArray(NativeInt(P) + NativeInt(PHeader.Length));
  end;

  if (FRecord is TChartRecordUpdate) then
    FRecord.Resize(-FRecord.Length);
end;

constructor TGelFrame.Create;
begin
  FOPT := TOPT.Create;
end;

destructor TGelFrame.Destroy;
begin
  FOPT.Free;
  inherited;
end;

procedure TGelFrame.FillEffectPicture(Stream: TStream; PictType: TMSOBlipType);
begin
  LoadPictFromStream(Stream,PictType);
  FOPT.UpdateValue(msooptFillType,msofillPicture);
  InsertPICRec;
end;

procedure TGelFrame.FillEffectGradient(Color1, Color2: TXc12IndexColor; GradientStyle: TGradientFillStyle);

function To1616(ValInt,ValFrac: integer): longword;
begin
  Result := (ValInt shl 16) + ValFrac;
end;

begin
  RemovePICRec;
  FOPT.UpdateValue(msooptFillColor,$08000000 + Longword(Color1));
  FOPT.UpdateValue(msooptFillBackColor,$08000000 + Longword(Color2));
  case GradientStyle of
    gfsHorizontal: begin
      FOPT.UpdateValue(msooptFillType,msofillShadeScale);
      FOPT.UpdateValue(msooptFillAngle,To1616(0,0));
    end;
    gfsVertical: begin
      FOPT.UpdateValue(msooptFillType,msofillShadeScale);
      FOPT.UpdateValue(msooptFillAngle,To1616(-90,0));
    end;
    gfsDiagonalUp: begin
      FOPT.UpdateValue(msooptFillType,msofillShadeScale);
      FOPT.UpdateValue(msooptFillAngle,To1616(-135,0));
    end;
    gfsDiagonalDown: begin
      FOPT.UpdateValue(msooptFillType,msofillShadeScale);
      FOPT.UpdateValue(msooptFillAngle,To1616(-45,0));
    end;
    gfsFromCorner: begin
      FOPT.UpdateValue(msooptFillType,msofillShadeCenter);
      FOPT.UpdateValue(msooptFillAngle,To1616(-45,0));
    end;
    gfsFromCenter: begin
      FOPT.UpdateValue(msooptFillType,msofillShadeShape);
      FOPT.UpdateValue(msooptFillAngle,To1616(-45,0));
    end;
  end;
end;

procedure TGelFrame.FillEffectPicture(Filename: AxUCString);
begin
  LoadPictFromFile(Filename);
  FOPT.UpdateValue(msooptFillType,msofillPicture);
  InsertPICRec;
end;

procedure TGelFrame.FillEffectTexture(Stream: TStream; PictType: TMSOBlipType);
begin
  LoadPictFromStream(Stream,PictType);
  FOPT.UpdateValue(msooptFillType,msofillTexture);
  RemovePICRec;
end;

procedure TGelFrame.FillEffectTexture(Filename: AxUCString);
begin
  LoadPictFromFile(Filename);
  FOPT.UpdateValue(msooptFillType,msofillTexture);
//  FOPT.UpdateString(msooptFillBlipName,'Blip');
  RemovePICRec;
end;

function TGelFrame.GetPictureType: TMSOBlipType;
var
  i: integer;
begin
  Result := msoblipERROR;
  i := FOPT.Find(msooptFillBlip);
  if (i >= 0) or (FOPT.Value[i] > 0) then begin
    case (PMSOHeader(FOPT.Extra[i]).VerInst shr 4) and $0FFF of
      $000     : Result := msoblipUNKNOWN;
      BLIP_WMF : Result := msoblipWMF;
      BLIP_EMF : Result := msoblipEMF;
      BLIP_PICT: Result := msoblipPICT;
      BLIP_PNG : Result := msoblipPNG;
      BLIP_JPEG: Result := msoblipJPEG;
      BLIP_DIB : Result := msoblipDIB;
      $800     : Result := msoblipUNKNOWN;
      else
        Result := msoblipUNKNOWN;
    end;
  end;
end;

procedure TGelFrame.InsertPICRec;
var
  i: integer;
  Rec: TChartRecord;
begin
  i := FRecord.FindRecord(CHARTRECID_PICF);
  if i >= 0 then
    Rec := FRecord[i]
  else
    Rec := FRecord.InsertRecord(-1,CHARTRECID_PICF,SizeOf(TCRecPICF));
  FillChar(Rec.Data^,SizeOf(TCRecPICF),#0);
  PCRecPICF(Rec.Data).PaintMode := 1;
  PCRecPICF(Rec.Data).Flags := $0E;
end;

procedure TGelFrame.LoadPictFromFile(Filename: AxUCString);
var
  PictType: TMSOBlipType;
  Stream: TFileStream;
begin
  PictType := BlipFromExt(Filename);
  Stream := TFileStream.Create(Filename,fmOpenRead);
  try
    LoadPictFromStream(Stream,PictType);
  finally
    Stream.Free;
  end;
end;

procedure TGelFrame.LoadPictFromStream(Stream: TStream; PictType: TMSOBlipType);
var
  i: integer;
  P: PByteArray;
begin
  i := FOPT.Find(msooptFillBlip);
  if i < 0 then
    raise XLSRWException.Create('There is no picture');
  if PictType in [msoblipERROR,msoblipUNKNOWN] then
    raise XLSRWException.Create('Can not load this picture');
  case PictType of
    msoblipEMF,msoblipWMF,msoblipPICT: begin
      raise XLSRWException.Create('Picture format not implemented yet');
    end;
    msoblipDIB: begin
      FOPT.UpdateExtra(msooptFillBlip,Stream.Size - SizeOf(TBmpFileHeader) + SizeOf(TMSOHeader) + 16 + 1);
      Stream.Seek(SizeOf(TBmpFileHeader),soFromBeginning);
      P := FOPT.Extra[i];
      PMSOHeader(P).VerInst := BLIP_DIB shl 4;
      PMSOHeader(P).FBT := MSO_BLIPSTART + Word(PictType);
      PMSOHeader(P).Length := Stream.Size - SizeOf(TBmpFileHeader) + 16 + 1;
      P := PByteArray(NativeInt(P) + SizeOf(TMSOHeader) + 16 + 1);
      Stream.Read(P^,Stream.Size - SizeOf(TBmpFileHeader));
    end;
    else begin
      FOPT.UpdateExtra(msooptFillBlip,Stream.Size + SizeOf(TMSOHeader) + 16 + 1);
      P := FOPT.Extra[i];
      if PictType = msoblipJPEG then
        PMSOHeader(P).VerInst := BLIP_JPEG shl 4
      else
        PMSOHeader(P).VerInst := BLIP_PNG shl 4;
      PMSOHeader(P).FBT := MSO_BLIPSTART + Word(PictType);
      PMSOHeader(P).Length := Stream.Size + 16 + 1;
      P := PByteArray(NativeInt(P) + SizeOf(TMSOHeader) + 16 + 1);
      Stream.Read(P^,Stream.Size);
    end;
  end;
  FOPT.UpdateString(msooptFillBlipName,'');
  FOPT.UpdateValue(msooptFillBlipFlags,2);
end;

procedure TGelFrame.RemovePICRec;
begin
  FRecord.RemoveRecord(CHARTRECID_PICF);
end;

procedure TGelFrame.SavePictToFile(Filename: AxUCString; AutoExt: boolean);
var
  Stream: TFileStream;
begin
  if AutoExt then begin
    case GetPictureType of
      msoblipWMF  : Filename := Filename + '.wmf';
      msoblipEMF  : Filename := Filename + '.emf';
      msoblipPICT : Filename := Filename + '.pict';
      msoblipJPEG : Filename := Filename + '.jpg';
      msoblipPNG  : Filename := Filename + '.png';
      msoblipDIB  : Filename := Filename + '.bmp';
    end;
  end;
  Stream := TFileStream.Create(Filename,fmCreate);
  try
    SavePictToStream(Stream);
  finally
    Stream.Free;
  end;
end;

procedure TGelFrame.SavePictToStream(Stream: TStream);
var
  i: integer;
  P: PByteArray;
  PictType: TMSOBlipType;
  BmpHdr: TBmpFileHeader;
begin
  i := FOPT.Find(msooptFillBlip);
  if (i < 0) or (FOPT.Value[i] <= 0) then
    raise XLSRWException.Create('There is no picture');
  PictType := GetPictureType;
  if PictType in [msoblipERROR,msoblipUNKNOWN] then
    raise XLSRWException.Create('Can not save this picture');
  P := PByteArray(NativeInt(FOPT.Extra[i]) + SizeOf(TMSOHeader) + 16 + 1);
  if PictType in [msoblipEMF,msoblipWMF,msoblipPICT] then begin
    raise XLSRWException.Create('Picture format not implemented yet');
  end
  else begin
    if PictType = msoblipDIB then begin
      BmpHdr.ID[0] := 'B';
      BmpHdr.ID[1] := 'M';
      BmpHdr.ID[2] := '8';
      // Not sure why 10 shall be added...
      BmpHdr.Size := PMSOHeader(FOPT.Extra[i]).Length - (SizeOf(TMSOHeader) + 16 + 1) + SizeOf(TBmpFileHeader) + 10;
      BmpHdr.Unknown1[0] := $00;
      BmpHdr.Unknown1[1] := $00;
      BmpHdr.Unknown1[2] := $00;
      BmpHdr.Unknown2 := $0036;
      BmpHdr.Unknown3 := $0000;
      Stream.Write(BmpHdr,SizeOf(TBmpFileHeader));
    end;
    // Not sure why 10 shall be added...
    Stream.Write(P^,PMSOHeader(FOPT.Extra[i]).Length - (SizeOf(TMSOHeader) + 16 + 1) + 10);
  end;
end;

procedure TGelFrame.SetDefault;
var
  i: integer;
  Root: TChartRecord;
  Shape: TShapeOutsideMsoChart;
begin
  FOPT.Clear;
  FOPT.AddValue(msooptFillType,$00000007);
  FOPT.AddValue(msooptFillColor,$08000009);
  FOPT.AddValue(msooptFillOpacity,$00010000);
  FOPT.AddValue(msooptFillBackColor,$08000016);
  FOPT.AddValue(msooptFillBackOpacity,$00010000);
  FOPT.AddValue(msooptFillCrMod,$100000F4);
  FOPT.AddExtra(msooptFillBlip,0,True);
  FOPT.AddExtra(msooptFillBlipName,0);
  FOPT.AddValue(msooptFillBlipFlags,$00000000);
  FOPT.AddValue(msooptFillWidth,$00000000);
  FOPT.AddValue(msooptFillHeight,$00000000);
  FOPT.AddValue(msooptFillAngle,$00000000);
  FOPT.AddValue(msooptFillFocus,$00000064);
  FOPT.AddValue(msooptFillToLeft,$00000000);
  FOPT.AddValue(msooptFillToTop,$00000000);
  FOPT.AddValue(msooptFillToRight,$00000000);
  FOPT.AddValue(msooptFillToBottom,$00000000);
  FOPT.AddValue(msooptFillRectLeft,$00000000);
  FOPT.AddValue(msooptFillRectTop,$00000000);
  FOPT.AddValue(msooptFillRectRight,$00000000);
  FOPT.AddValue(msooptFillRectBottom,$00000000);
  FOPT.AddValue(msooptFillDztype,$00000000);
  FOPT.AddValue(msooptFillShadePreset,$00000000);
  FOPT.AddExtra(msooptFillShadeColors,0);
  FOPT.AddValue(msooptFillOriginX,$00000000);
  FOPT.AddValue(msooptFillOriginY,$00000000);
  FOPT.AddValue(msooptFillShapeOriginX,$00000000);
  FOPT.AddValue(msooptFillShapeOriginY,$00000000);
  FOPT.AddValue(msooptFillShadeType,$40000003);
  FOPT.AddValue(msooptFNoFillHitTest,$001F001C);
  FOPT.SetItemCount;
  if (FRecord is TChartRecordUpdate) then
    TChartRecordUpdate(FRecord).OnUpdate := UpdateRecord;

  Root := FRecord.Root;
  if Root = Nil then
    raise XLSRWException.Create('Can not find chart root record');
  Shape := TShapeOutsideMsoChart(Root.Parent);
  if Shape.Drawing = Nil then begin
    Shape.Drawing := TEscherDrawing.Create(Shape.DrawingGroup,Nil,Nil);
    Shape.Drawing.AddEmpty;
    i := Root.FindRecord(BIFFRECID_PROTECT);
    if i < 0 then
      raise XLSRWException.Create('Can not find insertion point for MSODRAWING');
    Root.InsertRecord(i + 1,BIFFRECID_MSODRAWING,0);
  end;
end;

procedure TGelFrame.UpdateRecord(Sender: TObject);
begin
  FRecord.Resize(SizeOf(TMSOHeader) + FOPT.Size);
  FOPT.WriteToBuffer(FRecord.Data);
end;

{ TSheetCharts }

constructor TSheetCharts.Create(AOwner: TPersistent; DrawingGroup: TEscherGroup);
begin
  inherited Create(TSheetChart);
  FOwner := AOwner;
  FDrawingGroup := DrawingGroup;
end;

function TSheetCharts.Find(const AName: AxUCString): integer;
begin
  for Result := 0 to Count - 1 do begin
    if Items[Result].Name = AName then
      Exit;
  end;
  Result := -1;
end;

function TSheetCharts.GetSheetChart(Index: integer): TSheetChart;
begin
  Result := TSheetChart(inherited Items[Index]);
end;

function TSheetCharts.GetOwner: TPersistent;
begin
  Result := FOwner;
end;

procedure TSheetCharts.LoadFromStream(Stream: TXLSStream; Name: AxUCString; PBuf: PByteArray; Fonts: TXc12Fonts; SheetIndex: integer);
var
  H: TBIFFHeader;
  Chart: TSheetChart;
begin
  Chart := Add;
  Chart.Name := Name;
  Chart.FSheetIndex := SheetIndex;
  Chart.FRecord := TChartRecordUpdate.Create(TChartRecord(Self),ID_CHARTRECORDROOT,0);
  H.RecID := BIFFRECID_BOF8;
  H.Length := SizeOf(TRecBOF8);
  Chart.FRecord.Add(TChartRecord.Create(Chart.FRecord,H,PBuf));
  Chart.FRecord.Read(Stream,PBuf,Fonts);
end;

function TSheetCharts.Add: TSheetChart;
begin
  Result := TSheetChart(inherited Add);
end;

procedure TSheetCharts.SaveToStream(Index: integer; Stream: TXLSStream);
begin
  FCurrIndex := Index;
  Items[Index].FRecord.Write(Stream);
end;

{ TSheetChart }

procedure TSheetChart.AddDrawing(Drawing: TEscherDrawing);
begin
  // The TEscherDrawing is now owned by TSheetChart
  FDrawing := Drawing;
end;

constructor TSheetChart.Create(Collection: TCollection);
begin
  inherited Create(Collection);
  FName := 'Chart';
end;

destructor TSheetChart.Destroy;
begin
  if FDrawing <> Nil then begin
    FDrawing.Free;
    FDrawing := Nil;
  end;
  if FRecord <> Nil then begin
    FRecord.Free;
    FRecord := Nil;
  end;
  inherited;
end;

end.

