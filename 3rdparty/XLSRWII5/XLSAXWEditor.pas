unit XLSAXWEditor;

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

uses Classes, SysUtils, Contnrs, IniFiles,
     XLSUtils5;

const AXW_COLOR_AUTOMATIC          = $FF000000;
const AXW_COLOR_BLACK              = $00000000;
const AXW_COLOR_WHITE              = $00FFFFFF;

const AXW_COLOR_PAPER              = AXW_COLOR_WHITE;

const AXW_CHP_DEFAULT_FONT         = 'Arial';

const AXW_CHP_SUPERSCRIPT_SCALE    = 0.65;
const AXW_CHP_SUBSCRIPT_SCALE      = 0.65;

const AXW_HTML_DEF_FONTNAME        = 'Times New Roman';
const AXW_HTML_DEF_FONTSIZE        = 12;

type TST_Border =  (stbNil,stbNone,stbSingle,stbThick,stbDouble,stbDotted,stbDashed,stbDotDash,stbDotDotDash,stbTriple,stbThinThickSmallGap,stbThickThinSmallGap,stbThinThickThinSmallGap,stbThinThickMediumGap,stbThickThinMediumGap,stbThinThickThinMediumGap,stbThinThickLargeGap,stbThickThinLargeGap,stbThinThickThinLargeGap,stbWave,stbDoubleWave,stbDashSmallGap,stbDashDotStroked,stbThreeDEmboss,stbThreeDEngrave,stbOutset,stbInset,stbApples,stbArchedScallops,stbBabyPacifier,stbBabyRattle,stbBalloons3Colors,
stbBalloonsHotAir,stbBasicBlackDashes,stbBasicBlackDots,stbBasicBlackSquares,stbBasicThinLines,stbBasicWhiteDashes,stbBasicWhiteDots,stbBasicWhiteSquares,stbBasicWideInline,stbBasicWideMidline,stbBasicWideOutline,stbBats,stbBirds,stbBirdsFlight,stbCabins,stbCakeSlice,stbCandyCorn,stbCelticKnotwork,stbCertificateBanner,stbChainLink,stbChampagneBottle,stbCheckedBarBlack,stbCheckedBarColor,stbCheckered,stbChristmasTree,stbCirclesLines,stbCirclesRectangles,stbClassicalWave,stbClocks,stbCompass,stbConfetti,
stbConfettiGrays,stbConfettiOutline,stbConfettiStreamers,stbConfettiWhite,stbCornerTriangles,stbCouponCutoutDashes,stbCouponCutoutDots,stbCrazyMaze,stbCreaturesButterfly,stbCreaturesFish,stbCreaturesInsects,stbCreaturesLadyBug,stbCrossStitch,stbCup,stbDecoArch,stbDecoArchColor,stbDecoBlocks,stbDiamondsGray,stbDoubleD,stbDoubleDiamonds,stbEarth1,stbEarth2,stbEclipsingSquares1,stbEclipsingSquares2,stbEggsBlack,stbFans,stbFilm,stbFirecrackers,stbFlowersBlockPrint,stbFlowersDaisies,stbFlowersModern1,stbFlowersModern2,
stbFlowersPansy,stbFlowersRedRose,stbFlowersRoses,stbFlowersTeacup,stbFlowersTiny,stbGems,stbGingerbreadMan,stbGradient,stbHandmade1,stbHandmade2,stbHeartBalloon,stbHeartGray,stbHearts,stbHeebieJeebies,stbHolly,stbHouseFunky,stbHypnotic,stbIceCreamCones,stbLightBulb,stbLightning1,stbLightning2,stbMapPins,stbMapleLeaf,stbMapleMuffins,stbMarquee,stbMarqueeToothed,stbMoons,stbMosaic,stbMusicNotes,stbNorthwest,stbOvals,stbPackages,stbPalmsBlack,stbPalmsColor,stbPaperClips,stbPapyrus,stbPartyFavor,stbPartyGlass,stbPencils,
stbPeople,stbPeopleWaving,stbPeopleHats,stbPoinsettias,stbPostageStamp,stbPumpkin1,stbPushPinNote2,stbPushPinNote1,stbPyramids,stbPyramidsAbove,stbQuadrants,stbRings,stbSafari,stbSawtooth,stbSawtoothGray,stbScaredCat,stbSeattle,stbShadowedSquares,stbSharksTeeth,stbShorebirdTracks,stbSkyrocket,stbSnowflakeFancy,stbSnowflakes,stbSombrero,stbSouthwest,stbStars,stbStarsTop,stbStars3d,stbStarsBlack,stbStarsShadowed,stbSun,stbSwirligig,stbTornPaper,stbTornPaperBlack,stbTrees,stbTriangleParty,stbTriangles,stbTribal1,
stbTribal2,stbTribal3,stbTribal4,stbTribal5,stbTribal6,stbTwistedLines1,stbTwistedLines2,stbVine,stbWaveline,stbWeavingAngles,stbWeavingBraid,stbWeavingRibbon,stbWeavingStrips,stbWhiteFlowers,stbWoodwork,stbXIllusions,stbZanyTriangles,stbZigZag,stbZigZagStitch);

// **********************************************Ä******************************
// ****************************** CHPX properties ******************************
// **********************************************Ä******************************
type TAXWChpId = (
// Boolean
axciBold,axciItalic,axciStrikeTrough,axciHidden,axciShadow,axciOutline,axciEmboss,
axciEngrave,axciSmallCaps,axciAllCaps,axciMisspelled,axciRTL,

// Integer
axciColor,axciFillColor,axciUnderline,axciUnderlineColor,axciSubSuperscript,
axciColorHLinkHoover,axciColorHLinkVisited,

// Float
axciSize,

// Object
axciStyle,

// Owner object
axciBorder,

// String
axciFontName
);

const CHPXBoolProps     = [axciBold,axciItalic,axciStrikeTrough,axciMisspelled,axciHidden,axciShadow,axciOutline,axciEmboss,
                           axciEngrave,axciSmallCaps,axciAllCaps,axciRTL];
const CHPXIntProps      = [axciColor,axciFillColor,axciColorHLinkHoover,axciColorHLinkVisited,axciUnderline,axciUnderlineColor,axciSubSuperscript];
const CHPXFloatProps    = [axciSize];
const CHPXObjProps      = [axciStyle];
const CHPXOwnedObjProps = [axciBorder];
const CHPXStringProps   = [axciFontName];

// *****************************************************************************
// ****************************** PAPX properties ******************************
// *****************************************************************************
type TAXWPapId = (
// Boolean
axpiContextualSpacing,axpiNoWrap,

// Integer
axpiAlignment,axpiColor,axpiOutlineLevel,axpiLineSpacingType,

// Float
axpiIndentLeft,axpiIndentRight,axpiIndentFirstLine,axpiIndentHanging,axpiLineSpacing,
axpiSpaceBefore,axpiSpaceAfter,

// Object
axpiStyle,

// Owned object
axpiBorderLeft,axpiBorderTop,axpiBorderRight,axpiBorderBottom

// String
);

// **********************************************Ä******************************
// ****************************** TAPX properties ******************************
// **********************************************Ä******************************
type TAXWTapId = (
// Integer
axtiAlignment,axtiWidthType,
// Float
axtiIndent,axtiWidth,axtiCellMargLeft,axtiCellMargTop,axtiCellMargRight,axtiCellMargBottom,

// Object
axtiStyle,

// Owned object
axtiBorderLeft,axtiBorderTop,axtiBorderRight,axtiBorderBottom,axtiBorderInsideVert,axtiBorderInsideHoriz
);

type TAXWPictureType = (aptBMP,aptJPG,aptPNG,aptGIF,aptEMF,aptWMF,aptPIC,aptRAW,aptUnknown);
const AXWPictureTypeName: array[TAXWPictureType] of AxUCString = ('bmp','jpg','png','gif','emf','wmf','pic','raw','_unknown_');

type TAXWTabStopAlignment = (atsaClear,atsaLeft,atsaCenter,atsaRight,atsaDecimal,atsaNum);

type TAXWTabStopLeader = (atslNone,atslDot,atslHyphen,atslUnderscore,atslHeavy,atslMiddleDot);

type TAXWIntegerList = class(TList)
private
     function  GetItems(Index: integer): integer;
     procedure SetItems(Index: integer; const Value: integer);
     function  GetAsString: AxUCString;
     procedure SetAsString(Value: AxUCString);
protected
public
     procedure Assign(AList: TAXWIntegerList);

     procedure Add(AValue: integer);

     procedure SetMinSize(ASize, AValue: integer);

     property AsString: AxUCString read GetAsString write SetAsString;
     property Items[Index: integer]: integer read GetItems write SetItems; default;
     end;

type TAXWHyperlink = class(TObject)
protected
     FURL: AxUCString;
public
     property URL: AxUCString read FURL write FURL;
     property Address: AxUCString read FURL write FURL;
     end;

type TAXWHyperlinks = class(TObjectList)
private
     function  GetItems(Index: integer): TAXWHyperlink;
protected
public
     constructor Create;

     function Add(AURL: AxUCString): TAXWHyperlink;

     property Items[Index: integer]: TAXWHyperlink read GetItems; default;
     end;

     TAXWDocPropBorder = class(TObject)
protected
     FWidth: double;
     FColor: longword;
     FStyle: TST_Border;
     FSpace: double;
public
     procedure Clear;

     procedure Assign(ABorder: TAXWDocPropBorder);
     function  Assigned: boolean;

     function  Equal(ABorder: TAXWDocPropBorder): boolean;

     procedure SetVals(AWidth: double; AColor: longword; AStyle: TST_Border; ASpace: double);

     function Margin: double;

     procedure ColorRGB(AR, AG, AB: byte);

     property Width: double read FWidth write FWidth;
     property Color: longword read FColor write FColor;
     property Style: TST_Border read FStyle write FStyle;
     property Space: double read FSpace write FSpace;
     end;

     TAXWDocPropBorders = class(TObject)
protected
     FOwnsBr   : boolean;
     FDestroyBr: boolean;
     FLeft     : TAXWDocPropBorder;
     FTop      : TAXWDocPropBorder;
     FRight    : TAXWDocPropBorder;
     FBottom   : TAXWDocPropBorder;
public
     constructor Create(AOwnsBorders,ADestroyBorders: boolean);
     destructor Destroy; override;

     procedure Clear;

     procedure Assign(ABorders: TAXWDocPropBorders);
     function  Assigned: boolean;
     function  AllAssigned: boolean;

     procedure DeleteNil;

     function  AllEqual: boolean;

     function  AddLeft: TAXWDocPropBorder;
     function  AddTop: TAXWDocPropBorder;
     function  AddRight: TAXWDocPropBorder;
     function  AddBottom: TAXWDocPropBorder;
     procedure AddAll;

     property Left  : TAXWDocPropBorder read FLeft write FLeft;
     property Top   : TAXWDocPropBorder read FTop write FTop;
     property Right : TAXWDocPropBorder read FRight write FRight;
     property Bottom: TAXWDocPropBorder read FBottom write FBottom;
     end;

     TAXWDocPropBordersList = class(TObjectList)
private
     function GetItems(Index: integer): TAXWDocPropBorders;
     function GetLastAdded: TAXWDocPropBorder;
protected
     FLastAdded: TAXWDocPropBorder;
public
     constructor Create;

     procedure Assign(ABorders: TAXWDocPropBordersList);

     function  Add: TAXWDocPropBorders;
     function  AddLeft: TAXWDocPropBorder;
     function  AddTop: TAXWDocPropBorder;
     function  AddRight: TAXWDocPropBorder;
     function  AddBottom: TAXWDocPropBorder;

     function  Last: TAXWDocPropBorders;

     function LastAllAssigned: boolean;

     property LastAdded: TAXWDocPropBorder read GetLastAdded;
     property Items[Index: integer]: TAXWDocPropBorders read GetItems; default;
     end;

type TAXWPapTextAlign = (axptaLeft,axptaCenter,axptaRight,axptaJustify);

type TAXWTableCellAlignVert = (atcavTop,atcavCenter,atcavBottom,atcavBoth);

type TAXWPropType = (aptBoolean,aptInteger,aptFloat,aptString,aptObject,aptOwnedObject);

type TAXWProp = class(TObject)
protected
     FId: integer;

     function  GetAsBoolean: boolean; virtual;
     procedure SetAsBoolean(const Value: boolean); virtual;
     function  GetAsInteger: integer; virtual;
     function  GetAsString: AxUCString; virtual;
     procedure SetAsInteger(const Value: integer); virtual;
     procedure SetAsString(const Value: AxUCString); virtual;
     function  GetAsObject: TObject; virtual;
     procedure SetAsObject(const Value: TObject); virtual;
     function  GetAsOwnedObject: TObject; virtual;
     procedure SetAsOwnedObject(const Value: TObject); virtual;
     function  GetAsFloat: double; virtual;
     procedure SetAsFloat(const Value: double); virtual;
public
     function Type_: TAXWPropType; virtual; abstract;

     property Id: integer read FId;
     property AsBoolean: boolean read GetAsBoolean write SetAsBoolean;
     property AsInteger: integer read GetAsInteger write SetAsInteger;
     property AsFloat: double read GetAsFloat write SetAsFloat;
     property AsString: AxUCString read GetAsString write SetAsString;
     property AsObject: TObject read GetAsObject write SetAsObject;
     property AsOwnedObject: TObject read GetAsOwnedObject write SetAsOwnedObject;
     end;

type TAXWPropBoolean = class(TAXWProp)
protected
     FValue: boolean;

     function  GetAsBoolean: boolean; override;
     procedure SetAsBoolean(const Value: boolean); override;
public
     function Type_: TAXWPropType; override;
     end;

type TAXWPropInteger = class(TAXWProp)
protected
     FValue: integer;

     function  GetAsInteger: integer; override;
     procedure SetAsInteger(const Value: integer); override;
public
     function Type_: TAXWPropType; override;
     end;

type TAXWPropFloat = class(TAXWProp)
protected
     FValue: double;

     function  GetAsFloat: double; override;
     procedure SetAsFloat(const Value: double); override;
public
     function Type_: TAXWPropType; override;
     end;

type TAXWPropString = class(TAXWProp)
protected
     FValue: AxUCString;

     function  GetAsString: AxUCString; override;
     procedure SetAsString(const Value: AxUCString); override;
public
     function Type_: TAXWPropType; override;
     end;

type TAXWPropObject = class(TAXWProp)
protected
     FValue: TObject;

     function  GetAsObject: TObject; override;
     procedure SetAsObject(const Value: TObject); override;
public
     destructor Destroy; override;

     function Type_: TAXWPropType; override;
     end;

type TAXWPropOwnedObject = class(TAXWProp)
protected
     FValue: TObject;

     function  GetAsOwnedObject: TObject; override;
     procedure SetAsOwnedObject(const Value: TObject); override;
public
     destructor Destroy; override;

     function Type_: TAXWPropType; override;
     end;

type TAXWProps = class(TObjectList)
private
     function  GetItems(Index: integer): TAXWProp;
protected
     FMaster: TAXWProps;
public
     constructor Create(AMaster: TAXWProps);

     procedure Assign(AProps: TAXWProps);

     function AddBoolean(AId: integer; AValue: Boolean): TAXWPropBoolean;
     function AddInteger(AId: integer; AValue: integer): TAXWPropInteger;
     function AddFloat(AId: integer; AValue: double): TAXWPropFloat;
     function AddString(AId: integer; AValue: AxUCString): TAXWPropString;
     function AddObject(AId: integer; AValue: TObject): TAXWPropObject;
     function AddOwnedObject(AId: integer; AValue: TObject): TAXWPropOwnedObject;

     function Find(AId: integer): integer;

     function FindBoolean(AId: integer; out AValue: boolean): boolean;
     function FindInteger(AId: integer; out AValue: integer): boolean;
     function FindFloat(AId: integer; out AValue: double): boolean;
     function FindString(AId: integer; out AValue: AxUCString): boolean;
     function FindObject(AId: integer; out AObject: TObject): boolean;

     function FindBooleanDef(AId: integer; ADefault: boolean): boolean;
     function FindIntegerDef(AId: integer; ADefault: integer): integer;
     function FindFloatDef(AId: integer; ADefault: double): double;
     function FindStringDef(AId: integer; ADefault: AxUCString): AxUCString;
     function FindObjectDef(AId: integer; ADefault: TObject): TObject;

     procedure Remove(AId: integer);

     property Items[Index: integer]: TAXWProp read GetItems; default;
     end;

type TAXWCharStyle = class;
     TAXWParaStyle = class;

     TAXWPAP = class(TAXWProps)
private
     function  GetAlignment: TAXWPapTextAlign;
     procedure SetAlignment(const Value: TAXWPapTextAlign);
     function  GetIndentFirstLine: double;
     function  GetIndentHanging: double;
     function  GetIndentLeft: double;
     function  GetIndentRight: double;
     procedure SetIndentFirstLine(const Value: double);
     procedure SetIndentHanging(const Value: double);
     procedure SetIndentLeft(const Value: double);
     procedure SetIndentRight(const Value: double);
     function  GetColor: longword;
     function  GetSpaceAfter: double;
     function  GetSpaceBefore: double;
     procedure SetColor(const Value: longword);
     procedure SetSpaceAfter(const Value: double);
     procedure SetSpaceBefore(const Value: double);
     function  GetLineSpacing: double;
     procedure SetLineSpacing(const Value: double);
     function  GetMaster: TAXWPAP;
     function  GetBorderBottom: TAXWDocPropBorder;
     function  GetBorderLeft: TAXWDocPropBorder;
     function  GetBorderRight: TAXWDocPropBorder;
     function  GetBorderTop: TAXWDocPropBorder;
     function  GetStyle: TAXWCharStyle;
     procedure SetStyle(const Value: TAXWCharStyle);
protected
public
     function  AddBorderLeft: TAXWDocPropBorder;
     function  AddBorderRight: TAXWDocPropBorder;
     function  AddBorderTop: TAXWDocPropBorder;
     function  AddBorderBottom: TAXWDocPropBorder;

     property Master         : TAXWPAP read GetMaster;

     property Style          : TAXWCharStyle read GetStyle write SetStyle;

     property Alignment      : TAXWPapTextAlign read GetAlignment write SetAlignment;

     property IndentLeft     : double read GetIndentLeft write SetIndentLeft;
     property IndentRight    : double read GetIndentRight write SetIndentRight;
     property IndentFirstLine: double read GetIndentFirstLine write SetIndentFirstLine;
     property IndentHanging  : double read GetIndentHanging write SetIndentHanging;

     property Color          : longword read GetColor write SetColor;

     property SpaceBefore    : double read GetSpaceBefore write SetSpaceBefore;
     property SpaceAfter     : double read GetSpaceAfter write SetSpaceAfter;
     property LineSpacing    : double read GetLineSpacing write SetLineSpacing;

     property BorderLeft     : TAXWDocPropBorder read GetBorderLeft;
     property BorderRight    : TAXWDocPropBorder read GetBorderRight;
     property BorderTop      : TAXWDocPropBorder read GetBorderTop;
     property BorderBottom   : TAXWDocPropBorder read GetBorderBottom;
     end;

     TAXWPAPX = TAXWPAP;

     TAXWChpUnderline =
      (axcuNone,axcuSingle,axcuWords,axcuDouble,axcuThick,axcuDotted,axcuDottedHeavy,
       axcuDash,axcuDashedHeavy,axcuDashLong,axcuDashLongHeavy,axcuDotDash,axcuDashDotHeavy,
       axcuDotDotDash,axcuDashDotDotHeavy,axcuWave,axcuWavyHeavy,axcuWavyDouble);

     TAXWChpSubSuperscript = (axcssNone,axcssSubscript,axcssSuperscript);


     TAXWCHP = class(TAXWProps)
private
     function  GetStyle: TAXWCharStyle;
     procedure SetStyle(const Value: TAXWCharStyle);
     function  GetBold: boolean;
     function  GetColor: longword;
     function  GetFontName: AxUCString;
     function  GetItalic: boolean;
     function  GetSize: double;
     function  GetStrikeTrough: boolean;
     function  GetUnderline: TAXWChpUnderline;
     function  GetUnderlineColor: longword;
     procedure SetBold(const Value: boolean);
     procedure SetColor(const Value: longword);
     procedure SetFontName(const Value: AxUCString);
     procedure SetItalic(const Value: boolean);
     procedure SetSize(const Value: double);
     procedure SetStrikeTrough(const Value: boolean);
     procedure SetUnderline(const Value: TAXWChpUnderline);
     procedure SetUnderlineColor(const Value: longword);
     function  GetMaster: TAXWCHP;
     function  GetFillColor: longword;
     function  GetSubSuperscript: TAXWChpSubSuperscript;
     procedure SetFillColor(const Value: longword);
     procedure SetSubSuperscript(const Value: TAXWChpSubSuperscript);
protected
public
     function AddBorder: TAXWDocPropBorder;

     property Master        : TAXWCHP read GetMaster;

     property Style         : TAXWCharStyle read GetStyle write SetStyle;
     property FontName      : AxUCString read GetFontName write SetFontName;
     property Size          : double read GetSize write SetSize;
     property Bold          : boolean read GetBold write SetBold;
     property Italic        : boolean read GetItalic write SetItalic;
     property StrikeTrough  : boolean read GetStrikeTrough write SetStrikeTrough;
     property Color         : longword read GetColor write SetColor;
     property FillColor     : longword read GetFillColor write SetFillColor;
     property Underline     : TAXWChpUnderline read GetUnderline write SetUnderline;
     property UnderlineColor: longword read GetUnderlineColor write SetUnderlineColor;
     property SubSuperscript: TAXWChpSubSuperscript read GetSubSuperscript write SetSubSuperscript;
     end;

     TAXWCHPX = TAXWCHP;

     TAXWTAP = class(TAXWProps)
private
     function  GetCellMargBottom: double;
     function  GetCellMargLeft: double;
     function  GetCellMargRight: double;
     function  GetCellMargTop: double;
     procedure SetCellMargBottom(const Value: double);
     procedure SetCellMargLeft(const Value: double);
     procedure SetCellMargRight(const Value: double);
     procedure SetCellMargTop(const Value: double);
     function  GetBorderBottom: TAXWDocPropBorder;
     function  GetBorderLeft: TAXWDocPropBorder;
     function  GetBorderRight: TAXWDocPropBorder;
     function  GetBorderTop: TAXWDocPropBorder;
     function  GetBorderInsideHoriz: TAXWDocPropBorder;
     function  GetBorderInsideVert: TAXWDocPropBorder;
protected
     function  DoAddBorder(AId: TAXWTapId): TAXWDocPropBorder;
public
     function  AddBorderLeft: TAXWDocPropBorder;
     function  AddBorderRight: TAXWDocPropBorder;
     function  AddBorderTop: TAXWDocPropBorder;
     function  AddBorderBottom: TAXWDocPropBorder;

     procedure SetDefaultBorders;

     property CellMargLeft      : double read GetCellMargLeft write SetCellMargLeft;
     property CellMargTop       : double read GetCellMargTop write SetCellMargTop;
     property CellMargRight     : double read GetCellMargRight write SetCellMargRight;
     property CellMargBottom    : double read GetCellMargBottom write SetCellMargBottom;

     property BorderLeft: TAXWDocPropBorder read GetBorderLeft;
     property BorderRight: TAXWDocPropBorder read GetBorderRight;
     property BorderTop: TAXWDocPropBorder read GetBorderTop;
     property BorderBottom: TAXWDocPropBorder read GetBorderBottom;
     property BorderInsideVert: TAXWDocPropBorder read GetBorderInsideVert;
     property BorderInsideHoriz: TAXWDocPropBorder read GetBorderInsideHoriz;
     end;

     TAXWTAPX = TAXWTAP;

     TAXWSEP = class(TObject)
private
     function  GetLandscape: boolean;
     procedure SetLandscape(const Value: boolean);
protected
     FPageCode  : integer;
     FPageWidth : double;
     FPageHeight: double;

     FMargLeft  : double;
     FMargTop   : double;
     FMargRight : double;
     FMargBottom: double;
     FMargHeader: double;
     FMargFooter: double;
     FMargWeb   : double;
     FGutter    : double;
     FTitlePg   : boolean;
public
     property PageCode    : integer read FPageCode write FPageCode;
     property PageWidth   : double read FPageWidth;
     property PageHeight  : double read FPageHeight;
     property Landscape   : boolean read GetLandscape write SetLandscape;

     property MargLeft    : double read FMargLeft write FMargLeft;
     property MargTop     : double read FMargTop write FMargTop;
     property MargRight   : double read FMargRight write FMargRight;
     property MargBottom  : double read FMargBottom write FMargBottom;
     property MargHeader  : double read FMargHeader write FMargHeader;
     property MargWeb     : double read FMargWeb write FMargWeb;
     property Gutter      : double read FGutter write FGutter;
     property TitlePg     : boolean read FTitlePg write FTitlePg;
     end;

     TAXWRootData = class(TObject)
protected
     FMasterCHP: TAXWCHP;
     FMasterPAP: TAXWPAP;
     FMasterTAP: TAXWTAP;
public
     constructor Create;
     destructor Destroy; override;

     property MasterCHP: TAXWCHP read FMasterCHP;
     property MasterPAP: TAXWPAP read FMasterPAP;
     property MasterTAP: TAXWTAP read FMasterTAP;
     end;

     TAXWStyleType = (astPara,astCharRun,astTable,astNumbering);

     TAXWStyles = class;

     TAXWStyle = class(TObject)
protected
     FStyles  : TAXWStyles;

     FId      : AxUCString;
     FName    : AxUCString;
     FBasedOn : TAXWStyle;
public
     constructor Create(AStyles: TAXWStyles); virtual;

     function  Type_: TAXWStyleType; virtual; abstract;

     property Id      : AxUCString read FId write FId;
     property Name    : AxUCString read FName write FName;
     property BasedOn : TAXWStyle read FBasedOn write FBasedOn;
     end;

     TAXWCharStyle = class(TAXWStyle)
protected
     FCHPX: TAXWCHPX;
public
     constructor Create(AStyles: TAXWStyles); override;
     destructor Destroy; override;

     function  Type_: TAXWStyleType; override;

     property CHPX: TAXWCHPX read FCHPX;
     end;

     TAXWParaStyle = class(TAXWStyle)
protected
     FPAPX: TAXWPAPX;
public
     constructor Create(AStyles: TAXWStyles); override;
     destructor Destroy; override;

     function  Type_: TAXWStyleType; override;

     property PAPX: TAXWPAPX read FPAPX;
     end;

     TAXWStyleClass = class of TAXWStyle;

     TAXWStyles = class(TObject)
private
     function  GetItems(Index: integer): TAXWStyle;
protected
     FRootData : TAXWRootData;
     FItems    : THashedStringList;

     function Add(const AName: AxUCString; AClass: TAXWStyleClass): TAXWStyle;
public
     constructor Create(ARootData : TAXWRootData);
     destructor Destroy; override;

     procedure Clear;

     function  AddChar(const AName: AxUCString): TAXWCharStyle;
     function  AddPara(const AName: AxUCString): TAXWParaStyle;

     procedure GetAllStyles(AList: TStrings);

     function  FindByName(AName: AxUCString; AType: TAXWStyleType): TAXWStyle;

     property Items[Index: integer]: TAXWStyle read GetItems;
     end;

type TAXWFontData = class(TObject)
protected
     FName          : AxUCString;
     FSize          : double;
     FRotation      : integer;
     FBold          : boolean;
     FItalic        : boolean;
     FUnderline     : TAXWChpUnderline;
     FULColor       : longword;
     FULPos         : double;
     FULSize        : double;
     FCharSet       : byte;
     FStrikeOut     : boolean;
     FSubSuperOffs  : double;
     FSuperscript   : boolean;
     FSubSuperscript: TAXWChpSubSuperscript;
     FColor         : longword;
     FFillColor     : longword;
     FBorder        : TAXWDocPropBorder;
public
     constructor Create;
     procedure Clear;

     procedure Assign(AFontData: TAXWFontData);

     property Name          : AxUCString read FName write FName;
     property Size          : double read FSize write FSize;
     property Rotation      : integer read FRotation write FRotation;
     property Bold          : boolean read FBold write FBold;
     property Italic        : boolean read FItalic write FItalic;
     property Underline     : TAXWChpUnderline read FUnderline write FUnderline;
     property ULColor       : longword read FULColor write FULColor;
     property UnderlinePos  : double read FULPos write FULPos;
     property UnderlineSize : double read FULSize write FULSize;
     property StrikeOut     : boolean read FStrikeOut write FStrikeOut;
     property SubSuperscript: TAXWChpSubSuperscript read FSubSuperscript write FSubSuperscript;
     property SubsSuperOffs : double read FSubSuperOffs write FSubSuperOffs;
     property CharSet       : Byte read FCharSet write FCharSet;
     property Color         : longword read FColor write FColor;
     property FillColor     : longword read FFillColor write FFillColor;
     property Border        : TAXWDocPropBorder read FBorder write FBorder;
     end;

type TAXWFontDataDoc = class(TAXWFontData)
protected
public
     procedure Apply(AStyle: TAXWStyle); overload;
     procedure Apply(ACHP: TAXWCHP); overload;
     end;

type TAXWParaData = class(TObject)
protected
     FAlignment      : TAXWPapTextAlign;
     FIndentLeft     : double;
     FIndentRight    : double;
     FIndentFirstLine: double;
     FIndentHanging  : double;
     FLineSpacing    : double;
     FSpaceBefore    : double;
     FSpaceAfter     : double;
     FBorders        : TAXWDocPropBorders;
     FColor          : longword;
public
     constructor Create;
     destructor Destroy; override;

     procedure Reset(APAP: TAXWPAP);

     procedure Apply(AStyle: TAXWStyle); overload;
     procedure Apply(APAPX: TAXWPAPX); overload;

     function  MargLeft: double;
     function  MargRight: double;
     function  MargHoriz: double;

     function  HasBorders: boolean; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}

     property Alignment      : TAXWPapTextAlign read FAlignment write FAlignment;
     property IndentLeft     : double read FIndentLeft write FIndentLeft;
     property IndentRight    : double read FIndentRight write FIndentRight;
     property IndentFirstLine: double read FIndentFirstLine write FIndentFirstLine;
     property IndentHanging  : double read FIndentHanging write FIndentHanging;
     property LineSpacing    : double read FLineSpacing write FLineSpacing;
     property SpaceBefore    : double read FSpaceBefore write FSpaceBefore;
     property SpaceAfter     : double read FSpaceAfter write FSpaceAfter;
     property Borders        : TAXWDocPropBorders read FBorders;
     property Color          : longword read FColor write FColor;
     end;

type TAXWCharRunType = (acrtText,acrtTab,acrtBreak,acrtGraphic,acrtPicture,acrtHyperlink,acrtSimpleField,acrtLinked,acrtFootnote,acrtFootnoteRef);

const  AXWCharRunTypeText = [acrtText,acrtHyperlink,acrtSimpleField];

type  TXWCharRunBreakType = (acrbtLineBreak,acrbtPageBreak,acrbtColumn,acrbtSeparator,acrbtContinuationSeparator);

type TAXWCharRuns = class;

     TAXWCharRun = class(TObject)
protected
     FParent: TAXWCharRuns;
     FText  : AxUCString;
     FCHPX  : TAXWCHPX;
public
     constructor Create(AParent: TAXWCharRuns);
     destructor Destroy; override;

     function  Type_: TAXWCharRunType; virtual;

     procedure AddCHPX;

     property Text   : AxUCString read FText write FText;
     property RawText: AxUCString read FText write FText;
     property CHPX   : TAXWCHPX read FCHPX;
     end;

     TAXWCharRunHyperlink = class(TAXWCharRun)
protected
     FHyperlink: TAXWHyperlink;
public
     function  Type_: TAXWCharRunType; override;

     property Hyperlink: TAXWHyperlink read FHyperlink;
     end;

     TAXWCharRunTab = class(TAXWCharRun)
protected
public
     function  Type_: TAXWCharRunType; override;
     end;

     TAXWCharRunBreak = class(TAXWCharRun)
protected
     FBreakType: TXWCharRunBreakType;
public
     constructor Create(AParent: TAXWCharRuns; ABreakType: TXWCharRunBreakType);

     function  Type_: TAXWCharRunType; override;

     property BreakType: TXWCharRunBreakType read FBreakType;
     end;

     TAXWCharRuns = class(TObjectList)
private
     function  GetItems(Index: integer): TAXWCharRun;
protected
     FRootData: TAXWRootData;
public
     constructor Create(ARootData: TAXWRootData);

     function First: TAXWCharRun;

     function  Add: TAXWCharRun;
     function  AddTab: TAXWCharRunTab;
     function  AddBreak(ABreakType: TXWCharRunBreakType): TAXWCharRunBreak;

     function AddHyperlink(ACHPX: TAXWCHPX; AHyperlink: TAXWHyperlink): TAXWCharRunHyperlink;

     property RootData: TAXWRootData read FRootData;

     property Items[Index: integer]: TAXWCharRun read GetItems; default;
     end;

type TAXWTabStop = class(TObject)
protected
     FPosition : double;
     FAlignment: TAXWTabStopAlignment;
     FLeader   : TAXWTabStopLeader;
public
     constructor Create;

     property Position : double read FPosition write FPosition;
     property Alignment: TAXWTabStopAlignment read FAlignment write FAlignment;
     property Leader   : TAXWTabStopLeader read FLeader write FLeader;
     end;

type TAXWTabStops = class(TObjectList)
private
     function  GetItems(Index: integer): TAXWTabStop;
protected
public
     constructor Create;

     function  Last: TAXWTabStop;

     function  Add(const APosition: double; AAlignment: TAXWTabStopAlignment = atsaLeft): TAXWTabStop; overload;

     property Items[Index: integer]: TAXWTabStop read GetItems; default;
     end;

type TAXWLogParaType = (alptPara,alptTable);

type TAXWTable = class;
     TAXWTableRow = class;
     TAXWTableRows = class;
     TAXWLogParas = class;

     TAXWLogParaItem = class(TObject)
protected
     FParent: TAXWLogParas;
public
     constructor Create(AParent: TAXWLogParas);

     function Type_: TAXWLogParaType; virtual; abstract;

     property Parent: TAXWLogParas read FParent;
     end;

     TAXWNumberingStyleType = (anstNone,anstBullet,anstDecimal,anstLowerRoman,anstUpperRoman,anstLowerLetter,anstUpperLetter);

     TAXWLogPara = class(TAXWLogParaItem)
private
     function  GetPlainText: AxUCString;
protected
     FPAPX         : TAXWPAPX;
     FRuns         : TAXWCharRuns;
     FTabs         : TAXWTabStops;
public
     constructor Create(AParent: TAXWLogParas);
     destructor Destroy; override;

     procedure AddPAPX;

     procedure AddTabStops;

     function Type_: TAXWLogParaType; override;

     procedure SetupParaData(AParaData: TAXWParaData);

     property PAPX         : TAXWPAPX read FPAPX;
     property Runs         : TAXWCharRuns read FRuns;
     property TabStops     : TAXWTabStops read FTabs;
     property PlainText    : AxUCString read GetPlainText;
     end;

     TAXWTableBorderProps = class(TObject)
protected
     FBorders      : TAXWDocPropBorders;
     FBorderInsideH: TAXWDocPropBorder;
     FBorderInsideV: TAXWDocPropBorder;
public
     constructor Create;
     destructor Destroy; override;

     procedure Assign(AProps: TAXWTableBorderProps);

     procedure AddBorders(AWidth: double; AColor: longword; AStyle: TST_Border; ASpace: double);

     function  AddBorderTop: TAXWDocPropBorder; overload;
     function  AddBorderTop(AWidth: double; AColor: longword; AStyle: TST_Border; ASpace: double): TAXWDocPropBorder; overload;
     function  AddBorderLeft: TAXWDocPropBorder; overload;
     function  AddBorderLeft(AWidth: double; AColor: longword; AStyle: TST_Border; ASpace: double): TAXWDocPropBorder; overload;
     function  AddBorderBottom: TAXWDocPropBorder; overload;
     function  AddBorderBottom(AWidth: double; AColor: longword; AStyle: TST_Border; ASpace: double): TAXWDocPropBorder; overload;
     function  AddBorderRight: TAXWDocPropBorder;  overload;
     function  AddBorderRight(AWidth: double; AColor: longword; AStyle: TST_Border; ASpace: double): TAXWDocPropBorder; overload;
     function  AddBorderInsideH: TAXWDocPropBorder;
     function  AddBorderInsideV: TAXWDocPropBorder;

     property Borders          : TAXWDocPropBorders read FBorders;
     property BorderInsideHoriz: TAXWDocPropBorder read FBorderInsideH;
     property BorderInsideVert : TAXWDocPropBorder read FBorderInsideV;
     end;

     TAXWTableCellProps = class(TAXWTableBorderProps)
protected
     FBorders      : TAXWDocPropBorders;
     FBorderTl2br  : TAXWDocPropBorder;
     FBorderTr2bl  : TAXWDocPropBorder;
     FFillColor    : longword;
     FMargLeft     : double;
     FMargTop      : double;
     FMargRight    : double;
     FMargBottom   : double;
public
     constructor Create;
     destructor Destroy; override;

     //! Set the properties from AProps.
     procedure Assign(AProps: TAXWTableCellProps);

     function  AddBorderTl2br: TAXWDocPropBorder;
     function  AddBorderTr2bl: TAXWDocPropBorder;

     //! Returns True if any border is created.
     function  HasBorders: boolean;

     //! Returns True if any marginal is set.
     function  HasMargs: boolean;

     property BorderTl2br      : TAXWDocPropBorder read FBorderTl2br;
     property BorderTr2bl      : TAXWDocPropBorder read FBorderTr2bl;

     property FillColor        : longword read FFillColor write FFillColor;

     property MargLeft         : double read FMargLeft write FMargLeft;
     property MargTop          : double read FMargTop write FMargTop;
     property MargRight        : double read FMargRight write FMargRight;
     property MargBottom       : double read FMargBottom write FMargBottom;
     end;

     TAXWTableCellFlag = (atcfSelected,atcfMerged,atcfMergedRoot,atcfMergedLast);
     TAXWTableCellFlags = set of TAXWTableCellFlag;

     TAXWTableCell = class(TObject)
protected
     FFlags    : TAXWTableCellFlags;

     FRow      : TAXWTableRow;
     FParas    : TAXWLogParas;

     FProps    : TAXWTableCellProps;

     FAlignVert: TAXWTableCellAlignVert;
     FWidth    : double;
public
     constructor Create(ARow: TAXWTableRow);
     destructor Destroy; override;

     function  AddProps: TAXWTableCellProps;

     property Flags    : TAXWTableCellFlags read FFlags write FFlags;
     property Row      : TAXWTableRow read FRow;
     property Paras    : TAXWLogParas read FParas;
     property Width    : double read FWidth write FWidth;
     property Props    : TAXWTableCellProps read FProps;
     property AlignVert: TAXWTableCellAlignVert read FAlignVert write FAlignVert;
     end;

     TAXWTableRow = class(TObjectList)
private
     function  GetCells(Index: integer): TAXWTableCell;
     function  GetPlainText: AxUCString;
protected
     FParent : TAXWTableRows;

     FAHeight: double;
public
     constructor Create(AParent: TAXWTableRows);

     function  Add: TAXWTableCell;

     function  LastPara: TAXWLogPara;

     function  Last: TAXWTableCell;

     function  Table: TAXWTable;

     property Parent   : TAXWTableRows read FParent;
     property AHeight  : double read FAHeight write FAHeight;
     property PlainText: AxUCString read GetPlainText;
     property Cells[Index: integer]: TAXWTableCell read GetCells; default;
     end;

     TAXWTableRows = class(TObjectList)
private
     function  GetItems(Index: integer): TAXWTableRow;
protected
     FParent: TAXWTable;
public
     constructor Create(AParent: TAXWTable);

     function Add: TAXWTableRow;

     function Last: TAXWTableRow;

     property Parent: TAXWTable read FParent;
     property Items[Index: integer]: TAXWTableRow read GetItems; default;
     end;

     TAXWTable = class(TAXWLogParaItem)
private
     function  GetItems(Index: integer): TAXWTableRow;
     function  GetPlainText: AxUCString;
protected
     FRows  : TAXWTableRows;
     FTAPX  : TAXWTAPX;
public
     constructor Create(AParent: TAXWLogParas);
     destructor Destroy; override;

     function Type_: TAXWLogParaType; override;

     function Count: integer;

     function Add: TAXWTableRow;

     function LastPara: TAXWLogPara;

     function LastCell: TAXWTableCell;

     property TAPX     : TAXWTAPX read FTAPX;
     property PlainText: AxUCString read GetPlainText;
     property Items[Index: integer]: TAXWTableRow read GetItems; default;
     end;

     TAXWLogParas = class(TObjectList)
private
     function  GetItems(Index: integer): TAXWLogParaItem;
     function  GetPlainText: AxUCString;
     function  GetPara(Index: integer): TAXWLogPara;
     function  GetTable(Index: integer): TAXWTable;
protected
     FRootData     : TAXWRootData;

     FAutoNumbering: TAXWNumberingStyleType;
public
     constructor Create(ARootData: TAXWRootData);

     function CreatePara: TAXWLogPara;

     function AppendPara: TAXWLogPara;

     function _Add: TAXWLogPara;

     function AddTable: TAXWTable;

     function _Last: TAXWLogPara;

     function  Last: TAXWLogPara;

     property RootData     : TAXWRootData read FRootData;
     property AutoNumbering: TAXWNumberingStyleType read FAutoNumbering write FAutoNumbering;
     property PlainText    : AxUCString read GetPlainText;
     property Items[Index: integer]: TAXWLogParaItem read GetItems;
     property Para[Index: integer]: TAXWLogPara read GetPara; default;
     property Tables[Index: integer]: TAXWTable read GetTable;
     end;

type TAXWLogDocEditor = class(TObject)
private
     function  GetMasterCHP: TAXWCHP;
     function  GetMasterPAP: TAXWPAP;
     function  GetPLainText: AxUCString;
protected
     FRootData  : TAXWRootData;

     FSEP       : TAXWSEP;

     FStyles    : TAXWStyles;

     FHyperlinks: TAXWHyperlinks;

     FParas     : TAXWLogParas;
public
     constructor Create;
     destructor Destroy; override;

     procedure Clear;

     property MasterCHP  : TAXWCHP read GetMasterCHP;
     property MasterPAP  : TAXWPAP read GetMasterPAP;

     property SEP        : TAXWSEP read FSEP;

     property Styles     : TAXWStyles read FStyles;

     property Hyperlinks : TAXWHyperlinks read FHyperlinks;

     property Paras      : TAXWLogParas read FParas;

     property PlainText  : AxUCString read GetPLainText;
     end;


procedure Swap(var A,B: integer); overload;
procedure Swap(var A,B: double); overload;

implementation

procedure Swap(var A,B: integer);  overload;
var
  T: integer;
begin
  T := A;
  A := B;
  B := T;
end;

procedure Swap(var A,B: double); overload;
var
  T: double;
begin
  T := A;
  A := B;
  B := T;
end;


{ TAXWIntegerList }

procedure TAXWIntegerList.Add(AValue: integer);
begin
  inherited Add(Pointer(AValue));
end;

procedure TAXWIntegerList.Assign(AList: TAXWIntegerList);
var
  i: integer;
begin
  Clear;

  for i := 0 to AList.Count - 1 do
    Add(AList[i]);
end;

function TAXWIntegerList.GetAsString: AxUCString;
var
  i: integer;
begin
  Result := '';
  for i := 0 to Count - 1 do
    Result := Result + IntToStr(Items[i]) + ',';
  Result := Copy(Result,1,Length(Result) - 1);
end;

function TAXWIntegerList.GetItems(Index: integer): integer;
begin
  Result := Integer(inherited Items[Index]);
end;

procedure TAXWIntegerList.SetAsString(Value: AxUCString);
var
  S: AxUCString;
begin
  Clear;

  while Value <> '' do begin
    S := SplitAtChar(',',Value);
    // Missing value, use default
    if S = '' then
      Add($70000000)
    else
      Add(StrToIntDef(S,0));
  end;
end;

procedure TAXWIntegerList.SetItems(Index: integer; const Value: integer);
begin
  inherited Items[Index] := Pointer(Value);
end;

procedure TAXWIntegerList.SetMinSize(ASize, AValue: integer);
begin
  while Count < ASize do
    Add(AValue);
end;

{ TAXWLogPara }

procedure TAXWLogPara.AddPAPX;
begin
  if FPAPX = Nil then
    FPAPX := TAXWPAPX.Create(FParent.RootData.MasterPAP);
end;

procedure TAXWLogPara.AddTabStops;
begin
  if FTabs = Nil then
    FTabs := TAXWTabStops.Create;
end;

constructor TAXWLogPara.Create(AParent: TAXWLogParas);
begin
  inherited Create(AParent);

  FRuns := TAXWCharRuns.Create(FParent.RootData);
end;

destructor TAXWLogPara.Destroy;
begin
  FRuns.Free;

  if FTabs <> Nil then
    FTabs.Free;

  if FPAPX <> Nil then
    FPAPX.Free;

  inherited;
end;

function TAXWLogPara.GetPlainText: AxUCString;
var
  i: integer;
begin
  Result := '';

  for i := 0 to FRuns.Count - 1 do
    Result := Result + FRuns[i].Text;
end;

procedure TAXWLogPara.SetupParaData(AParaData: TAXWParaData);
//var
//  TStyle: TAXWTableStyle;
begin
  AParaData.Reset(FParent.FRootData.MasterPAP);

//  if FParas.Cell <> Nil then begin
//    TStyle := FParas.Cell.Table.TAPX.Style;
//    if (TStyle <> Nil) and (TStyle.PAPX <> NiL) then
//      AParaData.Apply(TStyle.PAPX)
//    else begin
//      AParaData.SpaceBefore := 0;
//      AParaData.SpaceAfter := 0;
//    end;
//  end;

  if FPAPX <> Nil then
    AParaData.Apply(FPAPX);
//  if (FNumbering <> Nil) and (FNumbering.PAPX <> Nil) then
//    AParaData.Apply(FNumbering.PAPX);
end;

function TAXWLogPara.Type_: TAXWLogParaType;
begin
  Result := alptPara;
end;

{ TAXWCharRuns }

function TAXWCharRuns.Add: TAXWCharRun;
begin
  Result := TAXWCharRun.Create(Self);

  inherited Add(Result);
end;

function TAXWCharRuns.AddBreak(ABreakType: TXWCharRunBreakType): TAXWCharRunBreak;
begin
  Result := TAXWCharRunBreak.Create(Self,ABreakType);

  inherited Add(Result);
end;

function TAXWCharRuns.AddHyperlink(ACHPX: TAXWCHPX; AHyperlink: TAXWHyperlink): TAXWCharRunHyperlink;
begin
  Result := TAXWCharRunHyperlink.Create(Self);
  Result.FCHPX := ACHPX;
  Result.FHyperlink := AHyperlink;

  inherited Add(Result);
end;

function TAXWCharRuns.AddTab: TAXWCharRunTab;
begin
  Result := TAXWCharRunTab.Create(Self);

  inherited Add(Result);
end;

constructor TAXWCharRuns.Create(ARootData: TAXWRootData);
begin
  inherited Create;

  FRootData := ARootData;
end;

function TAXWCharRuns.First: TAXWCharRun;
begin
  Result := TAXWCharRun(inherited First);
end;

function TAXWCharRuns.GetItems(Index: integer): TAXWCharRun;
begin
  Result := TAXWCharRun(inherited Items[Index]);
end;

{ TAXWProp }

function TAXWProp.GetAsBoolean: boolean;
begin
  raise Exception.Create('Invalid AXW Prop');
end;

function TAXWProp.GetAsFloat: double;
begin
  raise Exception.Create('Invalid AXW Prop');
end;

function TAXWProp.GetAsInteger: integer;
begin
  raise Exception.Create('Invalid AXW Prop');
end;

function TAXWProp.GetAsObject: TObject;
begin
  raise Exception.Create('Invalid AXW Prop');
end;

function TAXWProp.GetAsOwnedObject: TObject;
begin
  raise Exception.Create('Invalid AXW Prop');
end;

function TAXWProp.GetAsString: AxUCString;
begin
  raise Exception.Create('Invalid AXW Prop');
end;

procedure TAXWProp.SetAsBoolean(const Value: boolean);
begin
  raise Exception.Create('Invalid AXW Prop');
end;

procedure TAXWProp.SetAsFloat(const Value: double);
begin
  raise Exception.Create('Invalid AXW Prop');
end;

procedure TAXWProp.SetAsInteger(const Value: integer);
begin
  raise Exception.Create('Invalid AXW Prop');
end;

procedure TAXWProp.SetAsObject(const Value: TObject);
begin
  raise Exception.Create('Invalid AXW Prop');
end;

procedure TAXWProp.SetAsOwnedObject(const Value: TObject);
begin
  raise Exception.Create('Invalid AXW Prop');
end;

procedure TAXWProp.SetAsString(const Value: AxUCString);
begin
  raise Exception.Create('Invalid AXW Prop');
end;

{ TAXWPropInteger }

function TAXWPropInteger.GetAsInteger: integer;
begin
  Result := FValue;
end;

procedure TAXWPropInteger.SetAsInteger(const Value: integer);
begin
  FValue := Value;
end;

function TAXWPropInteger.Type_: TAXWPropType;
begin
  Result := aptInteger;
end;

{ TAXWPropString }

function TAXWPropString.GetAsString: AxUCString;
begin
  Result := FValue;
end;

procedure TAXWPropString.SetAsString(const Value: AxUCString);
begin
  FValue := Value;
end;

function TAXWPropString.Type_: TAXWPropType;
begin
  Result := aptString;
end;

{ TAXWProps }

function TAXWProps.AddBoolean(AId: integer; AValue: Boolean): TAXWPropBoolean;
var
  i: integer;
begin
  i := Find(AId);
  if i >= 0 then
    Result := TAXWPropBoolean(Items[i])
  else begin
    Result := TAXWPropBoolean.Create;
    Result.FId := AId;

    inherited Add(Result);
  end;
  Result.AsBoolean := AValue;
end;

function TAXWProps.AddFloat(AId: integer; AValue: double): TAXWPropFloat;
var
  i: integer;
begin
  i := Find(AId);
  if i >= 0 then
    Result := TAXWPropFloat(Items[i])
  else begin
    Result := TAXWPropFloat.Create;
    Result.FId := AId;

    inherited Add(Result);
  end;
  Result.AsFloat := AValue;
end;

function TAXWProps.AddInteger(AId, AValue: integer): TAXWPropInteger;
var
  i: integer;
begin
  i := Find(AId);
  if i >= 0 then
    Result := TAXWPropInteger(Items[i])
  else begin
    Result := TAXWPropInteger.Create;
    Result.FId := AId;

    inherited Add(Result);
  end;
  Result.AsInteger := AValue;
end;

function TAXWProps.AddObject(AId: integer; AValue: TObject): TAXWPropObject;
var
  i: integer;
begin
  i := Find(AId);
  if i >= 0 then
    Result := TAXWPropObject(Items[i])
  else begin
    Result := TAXWPropObject.Create;
    Result.FId := AId;

    inherited Add(Result);
  end;
  Result.AsObject := AValue;
end;

function TAXWProps.AddOwnedObject(AId: integer; AValue: TObject): TAXWPropOwnedObject;
var
  i: integer;
begin
  i := Find(AId);
  if i >= 0 then
    Result := TAXWPropOwnedObject(Items[i])
  else begin
    Result := TAXWPropOwnedObject.Create;
    Result.FId := AId;
    Result.FValue := AValue;

    inherited Add(Result);
  end;
  Result.AsOwnedObject := AValue;
end;

function TAXWProps.AddString(AId: integer; AValue: AxUCString): TAXWPropString;
var
  i: integer;
begin
  i := Find(AId);
  if i >= 0 then
    Result := TAXWPropString(Items[i])
  else begin
    Result := TAXWPropString.Create;
    Result.FId := AId;

    inherited Add(Result);
  end;
  Result.AsString := AValue;
end;

procedure TAXWProps.Assign(AProps: TAXWProps);
var
  i: integer;
begin
  Clear;

  for i := 0 to AProps.Count - 1 do begin
    case AProps[i].Type_ of
      aptBoolean: AddBoolean(AProps[i].Id,AProps[i].AsBoolean);
      aptInteger: AddInteger(AProps[i].Id,AProps[i].AsInteger);
      aptFloat  : AddFloat(AProps[i].Id,AProps[i].AsFloat);
      aptString : AddString(AProps[i].Id,AProps[i].AsString);
      else        raise Exception.Create('__TODO__');
    end;
  end;
end;

constructor TAXWProps.Create(AMaster: TAXWProps);
begin
  inherited Create;

  FMaster := AMaster;
end;

function TAXWProps.Find(AId: integer): integer;
begin
  for Result := 0 to Count - 1 do begin
    if Items[Result].FId = AId then
      Exit;
  end;

  Result := -1;
end;

function TAXWProps.FindBoolean(AId: integer; out AValue: boolean): boolean;
var
  i: integer;
begin
  i := Find(AId);

  if i >= 0 then begin
    AValue := Items[i].AsBoolean;
    Result := True;
  end
  else
    Result := False;
end;

function TAXWProps.FindBooleanDef(AId: integer; ADefault: boolean): boolean;
begin
  Result := ADefault;

  if not FindBoolean(AId,Result) and (FMaster <> Nil) then
    FMaster.FindBoolean(AId,Result);
end;

function TAXWProps.FindObject(AId: integer; out AObject: TObject): boolean;
var
  i: integer;
begin
  for i := 0 to Count - 1 do begin
    if Items[i].FId = AId then begin
      if Items[i] is TAXWPropOwnedObject then
        AObject := Items[i].AsOwnedObject
      else
        AObject := Items[i].AsObject;
      Result := True;
      Exit;
    end;
  end;
  Result := False;
end;

function TAXWProps.FindObjectDef(AId: integer; ADefault: TObject): TObject;
begin
  Result := ADefault;

  if not FindObject(AId,Result) and (FMaster <> Nil) then
    FMaster.FindObject(AId,Result);
end;

function TAXWProps.FindString(AId: integer; out AValue: AxUCString): boolean;
var
  i: integer;
begin
  i := Find(AId);

  if i >= 0 then begin
    AValue := Items[i].AsString;
    Result := True;
  end
  else
    Result := False;
end;

function TAXWProps.FindStringDef(AId: integer; ADefault: AxUCString): AxUCString;
begin
  Result := ADefault;

  if not FindString(AId,Result) and (FMaster <> Nil) then
    FMaster.FindString(AId,Result);
end;

function TAXWProps.GetItems(Index: integer): TAXWProp;
begin
  Result := TAXWProp(inherited Items[Index]);
end;

procedure TAXWProps.Remove(AId: integer);
var
  i: integer;
begin
  i := Find(AId);

  if i >= 0 then
    Delete(i);
end;

function TAXWProps.FindFloat(AId: integer; out AValue: double): boolean;
var
  i: integer;
begin
  i := Find(AId);

  if i >= 0 then begin
    AValue := Items[i].AsFloat;
    Result := True;
  end
  else
    Result := False;
end;

function TAXWProps.FindFloatDef(AId: integer; ADefault: double): double;
begin
  Result := ADefault;

  if not FindFloat(AId,Result) and (FMaster <> Nil) then
    FMaster.FindFloat(AId,Result);
end;

function TAXWProps.FindInteger(AId: integer; out AValue: integer): boolean;
var
  i: integer;
begin
  i := Find(AId);

  if i >= 0 then begin
    AValue := Items[i].AsInteger;
    Result := True;
  end
  else
    Result := False;
end;

function TAXWProps.FindIntegerDef(AId, ADefault: integer): integer;
begin
  Result := ADefault;

  if not FindInteger(AId,Result) and (FMaster <> Nil) then
    FMaster.FindInteger(AId,Result);
end;

{ TAXWCharRun }

procedure TAXWCharRun.AddCHPX;
begin
  if FCHPX = Nil then
    FCHPX := TAXWCHPX.Create(FParent.RootData.MasterCHP);
end;

constructor TAXWCharRun.Create(AParent: TAXWCharRuns);
begin
  FParent := AParent;
end;

destructor TAXWCharRun.Destroy;
begin
  if FCHPX <> Nil then
    FCHPX.Free;

  inherited;
end;

function TAXWCharRun.Type_: TAXWCharRunType;
begin
  Result := acrtText;
end;

{ TAXWLogDocEditor }

procedure TAXWLogDocEditor.Clear;
begin
  FParas.Clear;
end;

constructor TAXWLogDocEditor.Create;
begin
  FRootData := TAXWRootData.Create;

  FSEP := TAXWSEP.Create;

  FStyles := TAXWStyles.Create(FrootData);

  FHyperlinks := TAXWHyperlinks.Create;

  FParas := TAXWLogParas.Create(FRootData);
end;

destructor TAXWLogDocEditor.Destroy;
begin
  FParas.Free;

  FRootData.Free;

  FSEP.Free;

  FStyles.Free;

  FHyperlinks.Free;

  inherited;
end;

function TAXWLogDocEditor.GetMasterCHP: TAXWCHP;
begin
  Result := FRootData.MasterCHP;
end;

function TAXWLogDocEditor.GetMasterPAP: TAXWPAP;
begin
  Result := FRootData.MasterPAP;
end;

function TAXWLogDocEditor.GetPLainText: AxUCString;
begin
  Result := FParas.PlainText;
end;

{ TAXWLogParas }

function TAXWLogParas.AddTable: TAXWTable;
begin
  Result := TAXWTable.Create(Self);

  inherited Add(Result);
end;

function TAXWLogParas.AppendPara: TAXWLogPara;
begin
  Result := CreatePara;

  inherited Add(Result);
end;

constructor TAXWLogParas.Create(ARootData: TAXWRootData);
begin
  inherited Create;

  FRootData := ARootData;
end;

function TAXWLogParas.CreatePara: TAXWLogPara;
begin
  Result := TAXWLogPara.Create(Self);

//{$ifdef _AXOLOT_DEBUG}
//{$message warn '__TODO__ inherit para'}
//{$endif}
//  raise Exception.Create('__TODO__');
end;

function TAXWLogParas.GetItems(Index: integer): TAXWLogParaItem;
begin
  Result := TAXWLogParaItem(inherited Items[Index]);
end;

function TAXWLogParas.GetPara(Index: integer): TAXWLogPara;
begin
  if Items[Index].Type_ <> alptPara then
    raise XLSRWException.Create('Item is not para');

  Result := TAXWLogPara(Items[Index]);
end;

function TAXWLogParas.GetPlainText: AxUCString;
var
  i: integer;
begin
  Result := '';

  for i := 0 to Count - 1 do begin
    case Items[i].Type_ of
      alptPara : Result := Result + Para[i].PlainText;
      alptTable: Result := Result + Tables[i].PlainText;
    end;

  end;
end;

function TAXWLogParas.GetTable(Index: integer): TAXWTable;
begin
  if Items[Index].Type_ <> alptTable then
    XLSRWException.Create('Item is not table');

  Result := TAXWTable(Items[Index]);
end;

function TAXWLogParas.Last: TAXWLogPara;
begin
  Result := Nil;
  if Count > 0 then begin
    case Items[Count - 1].Type_ of
      alptPara : Result := TAXWLogPara(Items[Count - 1]);
      alptTable: Result := TAXWTable(Items[Count - 1]).LastPara;
    end;
  end;
end;

function TAXWLogParas._Add: TAXWLogPara;
begin
  Result := TAXWLogPara.Create(Self);

  inherited Add(Result);
end;

function TAXWLogParas._Last: TAXWLogPara;
begin
  Result := TAXWLogPara(inherited Last);
end;

{ TAXWTableCell }

function TAXWTableCell.AddProps: TAXWTableCellProps;
begin
  if FProps = Nil then
    FProps := TAXWTableCellProps.Create;

  Result := FProps;
end;

constructor TAXWTableCell.Create(ARow: TAXWTableRow);
begin
  FRow := ARow;

  FParas := TAXWLogParas.Create(FRow.Parent.Parent.Parent.RootData);
end;

destructor TAXWTableCell.Destroy;
begin
  FParas.Free;

  if FProps <> Nil then
    FProps.Free;

  inherited;
end;

{ TAXWTableRow }

function TAXWTableRow.Add: TAXWTableCell;
begin
  Result := TAXWTableCell.Create(Self);

  inherited Add(Result);
end;

constructor TAXWTableRow.Create(AParent: TAXWTableRows);
begin
  inherited Create;

  FParent := AParent;
end;

function TAXWTableRow.GetCells(Index: integer): TAXWTableCell;
begin
  Result := TAXWTableCell(inherited Items[Index]);
end;

function TAXWTableRow.GetPlainText: AxUCString;
var
  i: integer;
begin
  Result := '';

  for i := 0 to Count - 1 do
    Result := Result + Cells[i].Paras.PlainText;
end;

function TAXWTableRow.Last: TAXWTableCell;
begin
  Result := TAXWTableCell(inherited Last);
end;

function TAXWTableRow.LastPara: TAXWLogPara;
begin
  Result := Nil;
  if Count > 0  then
    Result := Cells[Count - 1].Paras.Last;
end;

function TAXWTableRow.Table: TAXWTable;
begin
  Result := FParent.FParent;
end;

{ TAXWTableRows }

function TAXWTableRows.Add: TAXWTableRow;
begin
  Result := TAXWTableRow.Create(Self);

  inherited Add(Result);
end;

constructor TAXWTableRows.Create(AParent: TAXWTable);
begin
  inherited Create;

  FParent := AParent;
end;

function TAXWTableRows.GetItems(Index: integer): TAXWTableRow;
begin
  Result := TAXWTableRow(inherited Items[Index]);
end;

function TAXWTableRows.Last: TAXWTableRow;
begin
  Result := TAXWTableRow(inherited Last);
end;

{ TAXWTable }

function TAXWTable.Add: TAXWTableRow;
begin
  Result := FRows.Add;
end;

function TAXWTable.Count: integer;
begin
  Result := FRows.Count;
end;

constructor TAXWTable.Create(AParent: TAXWLogParas);
begin
  inherited Create(AParent);

  FRows := TAXWTableRows.Create(Self);

  FTAPX := TAXWTAPX.Create(FParent.RootData.MasterTAP);
end;

destructor TAXWTable.Destroy;
begin
  FRows.Free;
  FTAPX.Free;

  inherited;
end;

function TAXWTable.GetItems(Index: integer): TAXWTableRow;
begin
  Result := FRows[Index];
end;

function TAXWTable.GetPlainText: AxUCString;
var
  i: integer;
begin
  Result := '';

  for i := 0 to FRows.Count - 1 do
    Result := Result + FRows[i].PlainText;
end;

function TAXWTable.LastCell: TAXWTableCell;
var
  Row: TAXWTableRow;
begin
  Result := Nil;
  if Count > 0 then begin
    Row := FRows.Last;
    if Row.Count > 0 then
      Result := Row.Last;
  end;
end;

function TAXWTable.LastPara: TAXWLogPara;
begin
  Result := Nil;
  if Count > 0 then
    Result := FRows.Last.LastPara;
end;

function TAXWTable.Type_: TAXWLogParaType;
begin
  Result := alptTable;
end;

{ TAXWTabStops }

function TAXWTabStops.Add(const APosition: double; AAlignment: TAXWTabStopAlignment): TAXWTabStop;
begin
  Result := TAXWTabStop.Create;
  Result.Position := APosition;
  Result.Alignment := AAlignment;

  inherited Add(Result);
end;

constructor TAXWTabStops.Create;
begin
  inherited Create;

end;

function TAXWTabStops.GetItems(Index: integer): TAXWTabStop;
begin
  Result := TAXWTabStop(inherited Items[Index]);
end;

function TAXWTabStops.Last: TAXWTabStop;
begin
  Result := TAXWTabStop(inherited Last);
end;

{ TAXWTabStop }

constructor TAXWTabStop.Create;
begin
  FAlignment := atsaLeft;
  FPosition := 0;
end;

{ TAXWDocPropBorder }

procedure TAXWDocPropBorder.Assign(ABorder: TAXWDocPropBorder);
begin
  FWidth := ABorder.FWidth;
  FColor := ABorder.FColor;
  FStyle := ABorder.FStyle;
  FSpace := ABorder.FSpace;
end;

function TAXWDocPropBorder.Assigned: boolean;
begin
  Result := FStyle > stbNone;
end;

procedure TAXWDocPropBorder.Clear;
begin
  FWidth := 0;
  FColor := AXW_COLOR_AUTOMATIC;
  FStyle := stbNone;
  FSpace := 0;
end;

procedure TAXWDocPropBorder.ColorRGB(AR, AG, AB: byte);
begin
  FColor := (AR shl 16) + (AG shl 8) + AB;
end;

function TAXWDocPropBorder.Equal(ABorder: TAXWDocPropBorder): boolean;
begin
  Result := (ClassType = ABorder.ClassType) and
            (FWidth = ABorder.FWidth) and
            (FColor = ABorder.FColor) and
            (FStyle = ABorder.FStyle) and
            (FSpace = ABorder.FSpace);
end;

function TAXWDocPropBorder.Margin: double;
begin
  Result := FWidth + FSpace;
end;

procedure TAXWDocPropBorder.SetVals(AWidth: double; AColor: longword; AStyle: TST_Border; ASpace: double);
begin
  FWidth := AWidth;
  FColor := AColor;
  FStyle := AStyle;
  FSpace := ASpace;
end;

{ TAXWDocPropBorders }

procedure TAXWDocPropBorders.AddAll;
begin
  AddLeft;
  AddTop;
  AddRight;
  AddBottom;
end;

function TAXWDocPropBorders.AddBottom: TAXWDocPropBorder;
begin
  if FBottom = Nil then
    FBottom := TAXWDocPropBorder.Create;
  Result := FBottom;
end;

function TAXWDocPropBorders.AddLeft: TAXWDocPropBorder;
begin
  if FLeft = Nil then
    FLeft := TAXWDocPropBorder.Create;
  Result := FLeft;
end;

function TAXWDocPropBorders.AddRight: TAXWDocPropBorder;
begin
  if FRight = Nil then
    FRight := TAXWDocPropBorder.Create;
  Result := FRight;
end;

function TAXWDocPropBorders.AddTop: TAXWDocPropBorder;
begin
  if FTop = Nil then
    FTop := TAXWDocPropBorder.Create;
  Result := FTop;
end;

function TAXWDocPropBorders.AllAssigned: boolean;
begin
  Result := (FLeft <> Nil) and (FTop <> Nil) and (FRight <> Nil) and (FBottom <> Nil);
end;

function TAXWDocPropBorders.AllEqual: boolean;
begin
  Result := (FLeft <> Nil) and (FTop <> Nil) and (FRight <> Nil) and (FBottom <> Nil);
  if Result then begin
    Result := ((FLeft.Width = FTop.Width) and (FLeft.Width = FRight.Width) and (FLeft.Width = FBottom.Width)) and
              ((FLeft.Style = FTop.Style) and (FLeft.Style = FRight.Style) and (FLeft.Style = FBottom.Style)) and
              ((FLeft.Color = FTop.Color) and (FLeft.Color = FRight.Color) and (FLeft.Color = FBottom.Color)) and
              ((FLeft.Space = FTop.Space) and (FLeft.Space = FRight.Space) and (FLeft.Space = FBottom.Space));
  end;
end;

procedure TAXWDocPropBorders.Assign(ABorders: TAXWDocPropBorders);
begin
  Clear;

  if ABorders.Left <> Nil then begin
    AddLeft;
    FLeft.Assign(ABorders.FLeft);
  end;
  if ABorders.Top <> Nil then begin
    AddTop;
    FTop.Assign(ABorders.FTop);
  end;
  if ABorders.Right <> Nil then begin
    AddRight;
    FRight.Assign(ABorders.FRight);
  end;
  if ABorders.Bottom <> Nil then begin
    AddBottom;
    FBottom.Assign(ABorders.FBottom);
  end;
end;

function TAXWDocPropBorders.Assigned: boolean;
begin
  Result := (FLeft <> Nil) or (FTop <> Nil) or (FRight <> Nil) or (FBottom <> Nil);
end;

procedure TAXWDocPropBorders.Clear;
begin
  if FLeft   <> Nil then FLeft.Free;
  if FTop    <> Nil then FTop.Free;
  if FRight  <> Nil then FRight.Free;
  if FBottom <> Nil then FBottom.Free;

  FLeft     := Nil;
  FTop      := Nil;
  FRight    := Nil;
  FBottom   := Nil;
end;

constructor TAXWDocPropBorders.Create(AOwnsBorders, ADestroyBorders: boolean);
begin
  FOwnsBr := AOwnsBorders;
  FDestroyBr := ADestroyBorders;
  if FOwnsBr then begin
    FLeft   := TAXWDocPropBorder.Create;
    FTop    := TAXWDocPropBorder.Create;
    FRight  := TAXWDocPropBorder.Create;
    FBottom := TAXWDocPropBorder.Create;
  end;
end;

procedure TAXWDocPropBorders.DeleteNil;
begin
  if (FLeft <> Nil) and (FLeft.Style in [stbNil]) then begin
    FLeft.Free;
    FLeft := Nil;
  end;
  if (FTop <> Nil) and (FTop.Style in [stbNil]) then begin
    FTop.Free;
    FTop := Nil;
  end;
  if (FRight <> Nil) and (FRight.Style in [stbNil]) then begin
    FRight.Free;
    FRight := Nil;
  end;
  if (FBottom <> Nil) and (FBottom.Style in [stbNil]) then begin
    FBottom.Free;
    FBottom := Nil;
  end;
end;

destructor TAXWDocPropBorders.Destroy;
begin
  if FOwnsBr then begin
    FLeft.Free;
    FTop.Free;
    FRight.Free;
    FBottom.Free;
  end
  else if FDestroyBr then
    Clear;

  inherited;
end;

{ TAXWDocPropBordersList }

function TAXWDocPropBordersList.Add: TAXWDocPropBorders;
begin
  Result := TAXWDocPropBorders.Create(False,True);

  inherited Add(Result);
end;

function TAXWDocPropBordersList.AddBottom: TAXWDocPropBorder;
begin
  if LastAllAssigned then
    Add;
  FLastAdded := Last.AddBottom;
  Result := FLastAdded;
end;

function TAXWDocPropBordersList.AddLeft: TAXWDocPropBorder;
begin
  if LastAllAssigned then
    Add;
  FLastAdded := Last.AddLeft;
  Result := FLastAdded;
end;

function TAXWDocPropBordersList.AddRight: TAXWDocPropBorder;
begin
  if LastAllAssigned then
    Add;
  FLastAdded := Last.AddRight;
  Result := FLastAdded;
end;

function TAXWDocPropBordersList.AddTop: TAXWDocPropBorder;
begin
  if LastAllAssigned then
    Add;
  FLastAdded := Last.AddTop;
  Result := FLastAdded;
end;

procedure TAXWDocPropBordersList.Assign(ABorders: TAXWDocPropBordersList);
var
  i: integer;
begin
  Clear;

  for i := 0 to ABorders.Count - 1 do
    Add.Assign(ABorders[i]);
end;

constructor TAXWDocPropBordersList.Create;
begin
  inherited Create;
end;

function TAXWDocPropBordersList.GetItems(Index: integer): TAXWDocPropBorders;
begin
  Result := TAXWDocPropBorders(inherited Items[Index]);
end;

function TAXWDocPropBordersList.GetLastAdded: TAXWDocPropBorder;
begin
  if FLastAdded = Nil then begin
    Add;
    AddTop;
  end;
  Result := FLastAdded;
end;

function TAXWDocPropBordersList.Last: TAXWDocPropBorders;
begin
  Result := TAXWDocPropBorders(inherited Last);
end;

function TAXWDocPropBordersList.LastAllAssigned: boolean;
begin
  Result := Count = 0;
  if not Result then
    Result := Last.AllAssigned;
end;

{ TAXWTAP }

function TAXWTAP.AddBorderBottom: TAXWDocPropBorder;
begin
  Result := DoAddBorder(axtiBorderBottom);
end;

function TAXWTAP.AddBorderLeft: TAXWDocPropBorder;
begin
  Result := DoAddBorder(axtiBorderLeft);
end;

function TAXWTAP.AddBorderRight: TAXWDocPropBorder;
begin
  Result := DoAddBorder(axtiBorderRight);
end;

function TAXWTAP.AddBorderTop: TAXWDocPropBorder;
begin
  Result := DoAddBorder(axtiBorderTop);
end;

function TAXWTAP.DoAddBorder(AId: TAXWTapId): TAXWDocPropBorder;
begin
  if not FindObject(Integer(AId),TObject(Result)) then begin
    Result := TAXWDocPropBorder.Create;
    AddOwnedObject(Integer(AId),Result);
  end;
end;

function TAXWTAP.GetBorderBottom: TAXWDocPropBorder;
begin
  Result := TAXWDocPropBorder(FindObjectDef(Integer(axtiBorderBottom),Nil));
end;

function TAXWTAP.GetBorderInsideHoriz: TAXWDocPropBorder;
begin
  Result := TAXWDocPropBorder(FindObjectDef(Integer(axtiBorderInsideHoriz),Nil));
end;

function TAXWTAP.GetBorderInsideVert: TAXWDocPropBorder;
begin
  Result := TAXWDocPropBorder(FindObjectDef(Integer(axtiBorderInsideVert),Nil));
end;

function TAXWTAP.GetBorderLeft: TAXWDocPropBorder;
begin
  Result := TAXWDocPropBorder(FindObjectDef(Integer(axtiBorderLeft),Nil));
end;

function TAXWTAP.GetBorderRight: TAXWDocPropBorder;
begin
  Result := TAXWDocPropBorder(FindObjectDef(Integer(axtiBorderRight),Nil));
end;

function TAXWTAP.GetBorderTop: TAXWDocPropBorder;
begin
  Result := TAXWDocPropBorder(FindObjectDef(Integer(axtiBorderTop),Nil));
end;

function TAXWTAP.GetCellMargBottom: double;
begin
  Result := FindFloatDef(Integer(axtiCellMargBottom),0.0);
end;

function TAXWTAP.GetCellMargLeft: double;
begin
  Result := FindFloatDef(Integer(axtiCellMargLeft),5.6);
end;

function TAXWTAP.GetCellMargRight: double;
begin
  Result := FindFloatDef(Integer(axtiCellMargRight),5.6);
end;

function TAXWTAP.GetCellMargTop: double;
begin
  Result := FindFloatDef(Integer(axtiCellMargTop),0.0);
end;

procedure TAXWTAP.SetCellMargBottom(const Value: double);
begin
  AddFloat(Integer(axtiCellMargBottom),Value);
end;

procedure TAXWTAP.SetCellMargLeft(const Value: double);
begin
  AddFloat(Integer(axtiCellMargLeft),Value);
end;

procedure TAXWTAP.SetCellMargRight(const Value: double);
begin
  AddFloat(Integer(axtiCellMargRight),Value);
end;

procedure TAXWTAP.SetCellMargTop(const Value: double);
begin
  AddFloat(Integer(axtiCellMargTop),Value);
end;

procedure TAXWTAP.SetDefaultBorders;
begin
  AddBorderLeft.SetVals(0.5,AXW_COLOR_BLACK,stbSingle,0);
  AddBorderTop.SetVals(0.5,AXW_COLOR_BLACK,stbSingle,0);
  AddBorderRight.SetVals(0.5,AXW_COLOR_BLACK,stbSingle,0);
  AddBorderBottom.SetVals(0.5,AXW_COLOR_BLACK,stbSingle,0);
end;

{ TAXWPropOwnedObject }

destructor TAXWPropOwnedObject.Destroy;
begin
  FValue.Free;

  inherited;
end;

function TAXWPropOwnedObject.GetAsOwnedObject: TObject;
begin
  Result := FValue;
end;

procedure TAXWPropOwnedObject.SetAsOwnedObject(const Value: TObject);
begin
  FValue := Value;
end;

function TAXWPropOwnedObject.Type_: TAXWPropType;
begin
  Result := aptOwnedObject;
end;

{ TAXWPropObject }

destructor TAXWPropObject.Destroy;
begin

  inherited;
end;

function TAXWPropObject.GetAsObject: TObject;
begin
  Result := FValue;
end;

procedure TAXWPropObject.SetAsObject(const Value: TObject);
begin
  FValue := Value;
end;

function TAXWPropObject.Type_: TAXWPropType;
begin
  Result := aptObject;
end;

{ TAXWCHP }

function TAXWCHP.AddBorder: TAXWDocPropBorder;
begin
  if not FindObject(Integer(axciBorder),TObject(Result)) then begin
    Result := TAXWDocPropBorder.Create;
    Result.Style := stbSingle;
    AddOwnedObject(Integer(axciBorder),Result);
  end;
end;

function TAXWCHP.GetBold: boolean;
begin
  Result := FindBooleanDef(Integer(axciBold),False);
end;

function TAXWCHP.GetColor: longword;
begin
  Result := FindIntegerDef(Integer(axciColor),AXW_COLOR_BLACK);
end;

function TAXWCHP.GetFillColor: longword;
begin
  Result := FindIntegerDef(Integer(axciFillColor),AXW_COLOR_WHITE);
end;

function TAXWCHP.GetFontName: AxUCString;
begin
  Result := FindStringDef(Integer(axciFontName),'Arial');
end;

function TAXWCHP.GetItalic: boolean;
begin
  Result := FindBooleanDef(Integer(axciItalic),False);
end;

function TAXWCHP.GetMaster: TAXWCHP;
begin
  Result := TAXWCHP(FMaster);
end;

function TAXWCHP.GetSize: double;
begin
  Result := FindFloatDef(Integer(axciSize),10);
end;

function TAXWCHP.GetStrikeTrough: boolean;
begin
  Result := FindBooleanDef(Integer(axciStrikeTrough),False);
end;

function TAXWCHP.GetStyle: TAXWCharStyle;
begin
  if not FindObject(Integer(axciStyle),TObject(Result)) then
    Result := Nil;
end;

function TAXWCHP.GetSubSuperscript: TAXWChpSubSuperscript;
begin
  Result := TAXWChpSubSuperscript(FindIntegerDef(Integer(axciFillColor),Integer(axcssNone)));
end;

function TAXWCHP.GetUnderline: TAXWChpUnderline;
begin
  Result := TAXWChpUnderline(FindIntegerDef(Integer(axciUnderline),Integer(axcuNone)));
end;

function TAXWCHP.GetUnderlineColor: longword;
begin
  Result := FindIntegerDef(Integer(axciUnderlineColor),AXW_COLOR_BLACK);
end;

procedure TAXWCHP.SetBold(const Value: boolean);
begin
  AddBoolean(Integer(axciBold),Value);
end;

procedure TAXWCHP.SetColor(const Value: longword);
begin
  AddInteger(Integer(axciColor),Value);
end;

procedure TAXWCHP.SetFillColor(const Value: longword);
begin
  AddInteger(Integer(axciFillColor),Value);
end;

procedure TAXWCHP.SetFontName(const Value: AxUCString);
begin
  AddString(Integer(axciFontName),Value);
end;

procedure TAXWCHP.SetItalic(const Value: boolean);
begin
  AddBoolean(Integer(axciItalic),Value);
end;

procedure TAXWCHP.SetSize(const Value: double);
begin
  AddFloat(Integer(axciSize),Value);
end;

procedure TAXWCHP.SetStrikeTrough(const Value: boolean);
begin
  AddBoolean(Integer(axciStrikeTrough),Value);
end;

procedure TAXWCHP.SetStyle(const Value: TAXWCharStyle);
begin
  if Value <> Nil then
    AddObject(Integer(axciStyle),Value)
  else
    Remove(Integer(axciStyle));
end;

procedure TAXWCHP.SetSubSuperscript(const Value: TAXWChpSubSuperscript);
begin
  AddInteger(Integer(axciSubSuperscript),Integer(Value));
end;

procedure TAXWCHP.SetUnderline(const Value: TAXWChpUnderline);
begin
  AddInteger(Integer(axciUnderline),Integer(Value));
end;

procedure TAXWCHP.SetUnderlineColor(const Value: longword);
begin
  AddInteger(Integer(axciUnderlineColor),Integer(Value));
end;

{ TAXWCharStyle }

constructor TAXWCharStyle.Create(AStyles: TAXWStyles);
begin
  inherited Create(AStyles);

  FCHPX := TAXWCHPX.Create(FStyles.FRootData.MasterCHP);
end;

destructor TAXWCharStyle.Destroy;
begin
  FCHPX.Free;

  inherited;
end;

function TAXWCharStyle.Type_: TAXWStyleType;
begin
  Result := astCharRun;
end;

{ TAXWSEP }

function TAXWSEP.GetLandscape: boolean;
begin
  Result := FPageWidth > FPageHeight;
end;

procedure TAXWSEP.SetLandscape(const Value: boolean);
begin
  if Value then begin
    if FPageWidth < FPageHeight then
      Swap(FPageWidth,FPageHeight);
  end
  else begin
    if FPageWidth > FPageHeight then
      Swap(FPageWidth,FPageHeight);
  end;
end;

{ TAXWCharRunBreak }

constructor TAXWCharRunBreak.Create(AParent: TAXWCharRuns; ABreakType: TXWCharRunBreakType);
begin
  inherited Create(AParent);

  FBreakType := ABreakType;
end;

function TAXWCharRunBreak.Type_: TAXWCharRunType;
begin
  Result := acrtBreak;
end;

{ TAXWCharRunHyperlink }

function TAXWCharRunHyperlink.Type_: TAXWCharRunType;
begin
  Result := acrtHyperlink;
end;

{ TAXWCharRunTab }

function TAXWCharRunTab.Type_: TAXWCharRunType;
begin
  Result := acrtTab;
end;

{ TAXWPropFloat }

function TAXWPropFloat.GetAsFloat: double;
begin
  Result := FValue;
end;

procedure TAXWPropFloat.SetAsFloat(const Value: double);
begin
  FValue := Value;
end;

function TAXWPropFloat.Type_: TAXWPropType;
begin
  Result := aptFloat;
end;

{ TAXWPAP }

function TAXWPAP.AddBorderBottom: TAXWDocPropBorder;
begin
  if not FindObject(Integer(axpiBorderBottom),TObject(Result)) then begin
    Result := TAXWDocPropBorder.Create;
    Result.Style := stbSingle;
    AddOwnedObject(Integer(axpiBorderBottom),Result);
  end;
end;

function TAXWPAP.AddBorderLeft: TAXWDocPropBorder;
begin
  if not FindObject(Integer(axpiBorderLeft),TObject(Result)) then begin
    Result := TAXWDocPropBorder.Create;
    Result.Style := stbSingle;
    AddOwnedObject(Integer(axpiBorderLeft),Result);
  end;
end;

function TAXWPAP.AddBorderRight: TAXWDocPropBorder;
begin
  if not FindObject(Integer(axpiBorderRight),TObject(Result)) then begin
    Result := TAXWDocPropBorder.Create;
    Result.Style := stbSingle;
    AddOwnedObject(Integer(axpiBorderRight),Result);
  end;
end;

function TAXWPAP.AddBorderTop: TAXWDocPropBorder;
begin
  if not FindObject(Integer(axpiBorderTop),TObject(Result)) then begin
    Result := TAXWDocPropBorder.Create;
    Result.Style := stbSingle;
    AddOwnedObject(Integer(axpiBorderTop),Result);
  end;
end;

function TAXWPAP.GetAlignment: TAXWPapTextAlign;
begin
  Result := TAXWPapTextAlign(FindIntegerDef(Integer(axpiAlignment),Integer(axptaLeft)));
end;

function TAXWPAP.GetBorderBottom: TAXWDocPropBorder;
begin
  Result := TAXWDocPropBorder(FindObjectDef(Integer(axpiBorderBottom),Nil));
end;

function TAXWPAP.GetBorderLeft: TAXWDocPropBorder;
begin
  Result := TAXWDocPropBorder(FindObjectDef(Integer(axpiBorderLeft),Nil));
end;

function TAXWPAP.GetBorderRight: TAXWDocPropBorder;
begin
  Result := TAXWDocPropBorder(FindObjectDef(Integer(axpiBorderRight),Nil));
end;

function TAXWPAP.GetBorderTop: TAXWDocPropBorder;
begin
  Result := TAXWDocPropBorder(FindObjectDef(Integer(axpiBorderTop),Nil));
end;

function TAXWPAP.GetColor: longword;
begin
  Result := FindIntegerDef(Integer(axpiColor),AXW_COLOR_WHITE);
end;

function TAXWPAP.GetIndentFirstLine: double;
begin
  Result := FindFloatDef(Integer(axpiIndentFirstLine),0.0);
end;

function TAXWPAP.GetIndentHanging: double;
begin
  Result := FindFloatDef(Integer(axpiIndentHanging),0.0);
end;

function TAXWPAP.GetIndentLeft: double;
begin
  Result := FindFloatDef(Integer(axpiIndentLeft),0.0);
end;

function TAXWPAP.GetIndentRight: double;
begin
  Result := FindFloatDef(Integer(axpiIndentRight),0.0);
end;

function TAXWPAP.GetLineSpacing: double;
begin
  Result := FindFloatDef(Integer(axpiLineSpacing),1.0);
end;

function TAXWPAP.GetMaster: TAXWPAP;
begin
  Result := TAXWPAP(FMaster);
end;

function TAXWPAP.GetSpaceAfter: double;
begin
  Result := FindFloatDef(Integer(axpiSpaceAfter),0.0);
end;

function TAXWPAP.GetSpaceBefore: double;
begin
  Result := FindFloatDef(Integer(axpiSpaceBefore),0.0);
end;

function TAXWPAP.GetStyle: TAXWCharStyle;
begin
  if not FindObject(Integer(axpiStyle),TObject(Result)) then
    Result := Nil;
end;

procedure TAXWPAP.SetAlignment(const Value: TAXWPapTextAlign);
begin
  AddInteger(Integer(axpiAlignment),Integer(Value));
end;

procedure TAXWPAP.SetColor(const Value: longword);
begin
  AddInteger(Integer(axpiColor),Value);
end;

procedure TAXWPAP.SetIndentFirstLine(const Value: double);
begin
  AddFloat(Integer(axpiIndentFirstLine),Value);
end;

procedure TAXWPAP.SetIndentHanging(const Value: double);
begin
  AddFloat(Integer(axpiIndentHanging),Value);
end;

procedure TAXWPAP.SetIndentLeft(const Value: double);
begin
  AddFloat(Integer(axpiIndentLeft),Value);
end;

procedure TAXWPAP.SetIndentRight(const Value: double);
begin
  AddFloat(Integer(axpiIndentRight),Value);
end;

procedure TAXWPAP.SetLineSpacing(const Value: double);
begin
  AddFloat(Integer(axpiLineSpacing),Value);
end;

procedure TAXWPAP.SetSpaceAfter(const Value: double);
begin
  AddFloat(Integer(axpiSpaceAfter),Value);
end;

procedure TAXWPAP.SetSpaceBefore(const Value: double);
begin
  AddFloat(Integer(axpiSpaceBefore),Value);
end;

procedure TAXWPAP.SetStyle(const Value: TAXWCharStyle);
begin
  if Value <> Nil then
    AddObject(Integer(axpiStyle),Value)
  else
    Remove(Integer(axpiStyle));
end;

{ TAXWRootData }

constructor TAXWRootData.Create;
begin
  FMasterCHP := TAXWCHP.Create(Nil);
  FMasterPAP := TAXWPAP.Create(Nil);
  FMasterTAP := TAXWTAP.Create(Nil);
end;

destructor TAXWRootData.Destroy;
begin
  FMasterCHP.Free;
  FMasterPAP.Free;
  FMasterTAP.Free;

  inherited;
end;

{ TAXWLogParaItem }

constructor TAXWLogParaItem.Create(AParent: TAXWLogParas);
begin
  FParent := AParent;
end;

{ TAXWPropBoolean }

function TAXWPropBoolean.GetAsBoolean: boolean;
begin
  Result := FValue;
end;

procedure TAXWPropBoolean.SetAsBoolean(const Value: boolean);
begin
  FValue := Value;
end;

function TAXWPropBoolean.Type_: TAXWPropType;
begin
  Result := aptBoolean;
end;

{ TAXWParaStyle }

constructor TAXWParaStyle.Create(AStyles: TAXWStyles);
begin
  inherited Create(AStyles);

  FPAPX := TAXWPAPX.Create(FStyles.FRootData.MasterPAP);
end;

destructor TAXWParaStyle.Destroy;
begin
  FPAPX.Free;

  inherited;
end;

function TAXWParaStyle.Type_: TAXWStyleType;
begin
  Result := astPara;
end;

{ TAXWStyles }

function TAXWStyles.Add(const AName: AxUCString; AClass: TAXWStyleClass): TAXWStyle;
var
  i: integer;
  S: AxUCString;
begin
  S := AName;

  i := 1;
  while FItems.IndexOf(S) >= 0 do begin
    S := S + IntToStr(i);
    if i > 1000 then
      raise XLSRWException.Create('Can not create style id');
  end;

  Result := AClass.Create(Self);
  Result.FId := S;
  Result.FName := AName;

  FItems.AddObject(S,Result);
end;

function TAXWStyles.AddChar(const AName: AxUCString): TAXWCharStyle;
begin
  Result := TAXWCharStyle(Add(AName,TAXWCharStyle));
end;

function TAXWStyles.AddPara(const AName: AxUCString): TAXWParaStyle;
begin
  Result := TAXWParaStyle(Add(AName,TAXWParaStyle));
end;

procedure TAXWStyles.Clear;
begin
  FItems.Clear;
end;

constructor TAXWStyles.Create(ARootData : TAXWRootData);
begin
  FRootData := ARootData;

  FItems := THashedStringList.Create;
end;

destructor TAXWStyles.Destroy;
var
  i: integer;
begin
  for i := 0 to FItems.Count - 1 do
    FItems.Objects[i].Free;

  FItems.Free;

  inherited;
end;

function TAXWStyles.FindByName(AName: AxUCString; AType: TAXWStyleType): TAXWStyle;
var
  i: integer;
begin
  AName := Lowercase(AName);

  for i := 0 to FItems.Count - 1 do begin
    if (Items[i].Type_ = AType) and (Lowercase(Items[i].Name) = AName) then begin
      Result := Items[i];
      Exit;
    end;
  end;
  Result := Nil;
end;

procedure TAXWStyles.GetAllStyles(AList: TStrings);
begin
  AList.Assign(FItems);
end;

function TAXWStyles.GetItems(Index: integer): TAXWStyle;
begin
  Result := TAXWStyle(FItems.Objects[Index]);
end;

{ TAXWStyle }

constructor TAXWStyle.Create(AStyles: TAXWStyles);
begin
  FStyles := AStyles;
end;

{ TAXWHyperlinks }

function TAXWHyperlinks.Add(AURL: AxUCString): TAXWHyperlink;
begin
  Result := TAXWHyperlink.Create;
  Result.URL := AURL;

  inherited Add(Result);
end;

constructor TAXWHyperlinks.Create;
begin
  inherited Create;
end;

function TAXWHyperlinks.GetItems(Index: integer): TAXWHyperlink;
begin
  Result := TAXWHyperlink(inherited Items[Index]);
end;

{ TAXWTableBorderProps }

function TAXWTableBorderProps.AddBorderBottom: TAXWDocPropBorder;
begin
  if FBorders.Bottom = Nil then
    FBorders.Bottom := TAXWDocPropBorder.Create;
  Result := FBorders.Bottom;
end;

function TAXWTableBorderProps.AddBorderBottom(AWidth: double; AColor: longword; AStyle: TST_Border; ASpace: double): TAXWDocPropBorder;
begin
  Result := AddBorderBottom;
  Result.SetVals(AWidth,AColor,AStyle,ASpace);
end;

function TAXWTableBorderProps.AddBorderInsideH: TAXWDocPropBorder;
begin
  if FBorderInsideH = Nil then
    FBorderInsideH  := TAXWDocPropBorder.Create;
  Result := FBorderInsideH;
end;

function TAXWTableBorderProps.AddBorderInsideV: TAXWDocPropBorder;
begin
  if FBorderInsideV = Nil then
    FBorderInsideV  := TAXWDocPropBorder.Create;
  Result := FBorderInsideV;
end;

function TAXWTableBorderProps.AddBorderLeft: TAXWDocPropBorder;
begin
  if FBorders.Left = Nil then
    FBorders.Left  := TAXWDocPropBorder.Create;
  Result := FBorders.Left;
end;

function TAXWTableBorderProps.AddBorderLeft(AWidth: double; AColor: longword; AStyle: TST_Border; ASpace: double): TAXWDocPropBorder;
begin
  Result := AddBorderLeft;
  Result.SetVals(AWidth,AColor,AStyle,ASpace);
end;

function TAXWTableBorderProps.AddBorderRight: TAXWDocPropBorder;
begin
  if FBorders.Right = Nil then
    FBorders.Right  := TAXWDocPropBorder.Create;
  Result := FBorders.Right;
end;

function TAXWTableBorderProps.AddBorderRight(AWidth: double; AColor: longword; AStyle: TST_Border; ASpace: double): TAXWDocPropBorder;
begin
  Result := AddBorderRight;
  Result.SetVals(AWidth,AColor,AStyle,ASpace);
end;

procedure TAXWTableBorderProps.AddBorders(AWidth: double; AColor: longword; AStyle: TST_Border; ASpace: double);
begin
  AddBorderLeft(AWidth,AColor,AStyle,ASpace);
  AddBorderTop(AWidth,AColor,AStyle,ASpace);
  AddBorderRight(AWidth,AColor,AStyle,ASpace);
  AddBorderBottom(AWidth,AColor,AStyle,ASpace);
end;

function TAXWTableBorderProps.AddBorderTop(AWidth: double; AColor: longword; AStyle: TST_Border; ASpace: double): TAXWDocPropBorder;
begin
  Result := AddBorderTop;
  Result.SetVals(AWidth,AColor,AStyle,ASpace);
end;

function TAXWTableBorderProps.AddBorderTop: TAXWDocPropBorder;
begin
  if FBorders.Top = Nil then
    FBorders.Top  := TAXWDocPropBorder.Create;
  Result := FBorders.Top;
end;

procedure TAXWTableBorderProps.Assign(AProps: TAXWTableBorderProps);
begin
  if AProps.Borders.Top <> Nil then begin
    AddBorderTop;
    FBorders.Top.Assign(AProps.FBorders.Top);
  end;
  if AProps.Borders.Left <> Nil then begin
    AddBorderLeft;
    FBorders.Left.Assign(AProps.FBorders.Left);
  end;
  if AProps.Borders.Bottom <> Nil then begin
    AddBorderBottom;
    FBorders.Bottom.Assign(AProps.FBorders.Bottom);
  end;
  if AProps.Borders.Right <> Nil then begin
    AddBorderRight;
    FBorders.Right.Assign(AProps.FBorders.Right);
  end;
  if AProps.FBorderInsideH <> Nil then begin
    AddBorderInsideH;
    FBorderInsideH.Assign(AProps.FBorderInsideH);
  end;
  if AProps.FBorderInsideV <> Nil then begin
    AddBorderInsideV;
    FBorderInsideV.Assign(AProps.FBorderInsideV);
  end;
end;

constructor TAXWTableBorderProps.Create;
begin
  FBorders := TAXWDocPropBorders.Create(False,False);
end;

destructor TAXWTableBorderProps.Destroy;
begin
  if FBorders.Top <> Nil then FBorders.Top.Free;
  if FBorders.Left <> Nil then FBorders.Left.Free;
  if FBorders.Bottom <> Nil then FBorders.Bottom.Free;
  if FBorders.Right <> Nil then FBorders.Right.Free;
  if FBorderInsideH <> Nil then FBorderInsideH.Free;
  if FBorderInsideV <> Nil then FBorderInsideV.Free;

  FBorders.Free;

  inherited;
end;

{ TAXWTableCellProps }

function TAXWTableCellProps.AddBorderTl2br: TAXWDocPropBorder;
begin
  if FBorderTl2br = Nil then
    FBorderTl2br  := TAXWDocPropBorder.Create;
  Result := FBorderTl2br;
end;

function TAXWTableCellProps.AddBorderTr2bl: TAXWDocPropBorder;
begin
  if FBorderTr2bl = Nil then
    FBorderTr2bl  := TAXWDocPropBorder.Create;
  Result := FBorderTr2bl;
end;

procedure TAXWTableCellProps.Assign(AProps: TAXWTableCellProps);
begin
  inherited Assign(AProps);

  if AProps.FBorderTl2br <> Nil then begin
    AddBorderTl2br;
    FBorderTl2br.Assign(AProps.FBorderTl2br);
  end;
  if AProps.FBorderTr2bl <> Nil then begin
    AddBorderTr2bl;
    FBorderTr2bl.Assign(AProps.FBorderTr2bl);
  end;
  FFillColor := AProps.FFillColor;

  FMargLeft := AProps.FMargLeft;
  FMargTop := AProps.FMargTop;
  FMargRight := AProps.FMargRight;
  FMargBottom := AProps.FMargBottom;
end;

constructor TAXWTableCellProps.Create;
begin
  inherited Create;

  FFillColor := AXW_COLOR_AUTOMATIC;
end;

destructor TAXWTableCellProps.Destroy;
begin
  if FBorderTl2br <> Nil then FBorderTl2br.Free;
  if FBorderTr2bl <> Nil then FBorderTr2bl.Free;

  inherited;
end;

function TAXWTableCellProps.HasBorders: boolean;
begin
  Result := (FBorders.Left <> Nil) or (FBorders.Top <> Nil) or (FBorders.Right <> Nil) or (FBorders.Bottom <> Nil);
end;

function TAXWTableCellProps.HasMargs: boolean;
begin
  Result := (FMargLeft <> 0) or (FMargTop <> 0) or (FMargRight <> 0) or (FMargBottom <> 0);
end;

{ TAXWFontDataDoc }

procedure TAXWFontDataDoc.Apply(AStyle: TAXWStyle);
begin

end;

procedure TAXWFontDataDoc.Apply(ACHP: TAXWCHP);
var
  i    : integer;
  Style: TAXWStyle;
begin
  Style := TAXWStyle(ACHP.Style);
  if Style <> Nil then
    Apply(Style);

  for i := 0 to ACHP.Count - 1 do begin
    case TAXWChpId(ACHP[i].Id) of
      axciBold          : FBold := ACHP[i].AsBoolean = True;
      axciItalic        : FItalic := ACHP[i].AsBoolean = True;
      axciSize          : FSize := ACHP[i].AsFloat;
      axciColor         : FColor := ACHP[i].AsInteger;
      axciFillColor     : FFillColor := ACHP[i].AsInteger;
      axciUnderline     : FUnderline := TAXWChpUnderline(ACHP[i].AsInteger);
      axciUnderlineColor: FULColor := ACHP[i].AsInteger;
      axciFontName      : FName := ACHP.FontName;
      axciSubSuperscript: FSubSuperscript := TAXWChpSubSuperscript(ACHP[i].AsInteger);
      axciStrikeTrough  : FStrikeOut := ACHP[i].AsBoolean = True;
//      axciBorder        : FBorder := ACHP.Border;
    end;
  end;
  case FSubSuperscript of
    axcssSubscript  : begin
      FSubSuperOffs := FSize * (1 - AXW_CHP_SUBSCRIPT_SCALE);
      FSize := FSize * AXW_CHP_SUBSCRIPT_SCALE;
    end;
    axcssSuperscript: begin
      FSubSuperOffs := FSize * -(1 - AXW_CHP_SUPERSCRIPT_SCALE);
      FSize := FSize * AXW_CHP_SUPERSCRIPT_SCALE;
    end;
  end;
end;

{ TAXWFontData }

procedure TAXWFontData.Assign(AFontData: TAXWFontData);
begin

end;

procedure TAXWFontData.Clear;
begin
  FName           := AXW_HTML_DEF_FONTNAME;
  FSize           := AXW_HTML_DEF_FONTSIZE;
  FRotation       := 0;
  FBold           := False;
  FItalic         := False;
  FUnderline      := axcuNone;
  FULColor        := AXW_COLOR_BLACK;
  FULPos          := 0;
  FULSize         := 0;
  FStrikeOut      := False;
  FCharSet        := 0;
  FColor          := AXW_COLOR_BLACK;
  FSubSuperscript := axcssNone;
  FSubSuperOffs   := 0;
  FBorder         := Nil;
end;

constructor TAXWFontData.Create;
begin
  Clear;
end;

{ TAXWParaData }

procedure TAXWParaData.Apply(AStyle: TAXWStyle);
begin
  if AStyle.BasedOn <> Nil then
    Apply(AStyle.BasedOn);

  if AStyle is TAXWParaStyle then
    Apply(TAXWParaStyle(AStyle).PAPX);
end;

procedure TAXWParaData.Apply(APAPX: TAXWPAPX);
var
  i: integer;
  Style: TAXWStyle;
begin
  Style := TAXWStyle(APAPX.Style);
  if Style <> Nil then
    Apply(Style);

  for i := 0 to APAPX.Count - 1 do begin
    case TAXWPapId(APAPX[i].Id) of
      axpiAlignment      : FAlignment       := APAPX.Alignment;
      axpiLineSpacing    : FLineSpacing     := APAPX.LineSpacing;
      axpiSpaceBefore    : FSpaceBefore     := APAPX.SpaceBefore;
      axpiSpaceAfter     : FSpaceAfter      := APAPX.SpaceAfter;
      axpiIndentLeft     : FIndentLeft      := APAPX.IndentLeft;
      axpiIndentRight    : FIndentRight     := APAPX.IndentRight;
      axpiIndentFirstLine: FIndentFirstLine := APAPX.IndentFirstLine;
      axpiIndentHanging  : FIndentHanging   := APAPX.IndentHanging;
      axpiBorderLeft     : FBorders.Left    := TAXWDocPropBorder(APAPX[i].AsObject);
      axpiBorderTop      : FBorders.Top     := TAXWDocPropBorder(APAPX[i].AsObject);
      axpiBorderRight    : FBorders.Right   := TAXWDocPropBorder(APAPX[i].AsObject);
      axpiBorderBottom   : FBorders.Bottom  := TAXWDocPropBorder(APAPX[i].AsObject);
      axpiColor          : FColor           := APAPX.Color;
    end;
  end;
  if FColor = AXW_COLOR_AUTOMATIC then
    FColor := AXW_COLOR_PAPER;
end;

constructor TAXWParaData.Create;
begin
  FBorders := TAXWDocPropBorders.Create(False,False);
end;

destructor TAXWParaData.Destroy;
begin
  FBorders.Free;

  inherited;
end;

function TAXWParaData.HasBorders: boolean;
begin
  Result := False;
end;

function TAXWParaData.MargHoriz: double;
begin
  Result := 0;
end;

function TAXWParaData.MargLeft: double;
begin
  Result := 0;
end;

function TAXWParaData.MargRight: double;
begin
  Result := 0;
end;

procedure TAXWParaData.Reset(APAP: TAXWPAP);
begin

end;

end.
