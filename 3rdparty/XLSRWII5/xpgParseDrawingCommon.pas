unit xpgParseDrawingCommon;

// Copyright (c) 2010,2012 Axolot Data
// Web : http://www.axolot.com/xpg
// Mail: xpg@axolot.com
//
// X   X  PPP    GGG
//  X X   P  P  G
//   X    PPP   G  GG
//  X X   P     G   G
// X   X  P      GGG
//
// File generated with Axolot XPG, Xml Parser Generator.
// Version 0.00.90.
// File created on 2012-10-30 16:01:18

{$I AxCompilers.inc}

{$MINENUMSIZE 2}
{$BOOLEVAL OFF}
{$RANGECHECKS OFF}

interface

uses Classes, SysUtils, Contnrs, IniFiles, Math,
     xpgPUtils, xpgPLists, xpgPXMLUtils, xpgPXML,
     Xc12Utils5, Xc12Graphics,
     XLSUtils5, XLSReadWriteOPC5;

const XPG_UNKNOWN_ENUM = $0000FFFF;

type TST_SchemeColorVal =  (stscvBg1,stscvTx1,stscvBg2,stscvTx2,stscvAccent1,stscvAccent2,stscvAccent3,stscvAccent4,stscvAccent5,stscvAccent6,stscvHlink,stscvFolHlink,stscvPhClr,stscvDk1,stscvLt1,stscvDk2,stscvLt2);
const StrTST_SchemeColorVal: array[0..16] of AxUCString = ('bg1','tx1','bg2','tx2','accent1','accent2','accent3','accent4','accent5','accent6','hlink','folHlink','phClr','dk1','lt1','dk2','lt2');
type TST_BlackWhiteMode =  (stbwmClr,stbwmAuto,stbwmGray,stbwmLtGray,stbwmInvGray,stbwmGrayWhite,stbwmBlackGray,stbwmBlackWhite,stbwmBlack,stbwmWhite,stbwmHidden);
const StrTST_BlackWhiteMode: array[0..10] of AxUCString = ('clr','auto','gray','ltGray','invGray','grayWhite','blackGray','blackWhite','black','white','hidden');
type TST_PresetColorVal =  (stpcvAliceBlue,stpcvAntiqueWhite,stpcvAqua,stpcvAquamarine,stpcvAzure,stpcvBeige,stpcvBisque,stpcvBlack,stpcvBlanchedAlmond,stpcvBlue,stpcvBlueViolet,stpcvBrown,stpcvBurlyWood,stpcvCadetBlue,stpcvChartreuse,stpcvChocolate,stpcvCoral,stpcvCornflowerBlue,stpcvCornsilk,stpcvCrimson,stpcvCyan,stpcvDkBlue,stpcvDkCyan,stpcvDkGoldenrod,stpcvDkGray,stpcvDkGreen,stpcvDkKhaki,stpcvDkMagenta,stpcvDkOliveGreen,stpcvDkOrange,stpcvDkOrchid,stpcvDkRed,stpcvDkSalmon,stpcvDkSeaGreen,stpcvDkSlateBlue,stpcvDkSlateGray,stpcvDkTurquoise,stpcvDkViolet,stpcvDeepPink,
stpcvDeepSkyBlue,stpcvDimGray,stpcvDodgerBlue,stpcvFirebrick,stpcvFloralWhite,stpcvForestGreen,stpcvFuchsia,stpcvGainsboro,stpcvGhostWhite,stpcvGold,stpcvGoldenrod,stpcvGray,stpcvGreen,stpcvGreenYellow,stpcvHoneydew,stpcvHotPink,stpcvIndianRed,stpcvIndigo,stpcvIvory,stpcvKhaki,stpcvLavender,stpcvLavenderBlush,stpcvLawnGreen,stpcvLemonChiffon,stpcvLtBlue,stpcvLtCoral,stpcvLtCyan,stpcvLtGoldenrodYellow,stpcvLtGray,stpcvLtGreen,stpcvLtPink,stpcvLtSalmon,stpcvLtSeaGreen,stpcvLtSkyBlue,stpcvLtSlateGray,stpcvLtSteelBlue,stpcvLtYellow,stpcvLime,stpcvLimeGreen,stpcvLinen,
stpcvMagenta,stpcvMaroon,stpcvMedAquamarine,stpcvMedBlue,stpcvMedOrchid,stpcvMedPurple,stpcvMedSeaGreen,stpcvMedSlateBlue,stpcvMedSpringGreen,stpcvMedTurquoise,stpcvMedVioletRed,stpcvMidnightBlue,stpcvMintCream,stpcvMistyRose,stpcvMoccasin,stpcvNavajoWhite,stpcvNavy,stpcvOldLace,stpcvOlive,stpcvOliveDrab,stpcvOrange,stpcvOrangeRed,stpcvOrchid,stpcvPaleGoldenrod,stpcvPaleGreen,stpcvPaleTurquoise,stpcvPaleVioletRed,stpcvPapayaWhip,stpcvPeachPuff,stpcvPeru,stpcvPink,stpcvPlum,stpcvPowderBlue,stpcvPurple,stpcvRed,stpcvRosyBrown,stpcvRoyalBlue,stpcvSaddleBrown,stpcvSalmon,stpcvSandyBrown,
stpcvSeaGreen,stpcvSeaShell,stpcvSienna,stpcvSilver,stpcvSkyBlue,stpcvSlateBlue,stpcvSlateGray,stpcvSnow,stpcvSpringGreen,stpcvSteelBlue,stpcvTan,stpcvTeal,stpcvThistle,stpcvTomato,stpcvTurquoise,stpcvViolet,stpcvWheat,stpcvWhite,stpcvWhiteSmoke,stpcvYellow,stpcvYellowGreen);
const StrTST_PresetColorVal: array[0..139] of AxUCString = ('aliceBlue','antiqueWhite','aqua','aquamarine','azure','beige','bisque','black','blanchedAlmond','blue','blueViolet','brown','burlyWood','cadetBlue','chartreuse','chocolate','coral','cornflowerBlue','cornsilk','crimson','cyan','dkBlue','dkCyan','dkGoldenrod','dkGray','dkGreen','dkKhaki','dkMagenta','dkOliveGreen','dkOrange','dkOrchid','dkRed','dkSalmon','dkSeaGreen','dkSlateBlue','dkSlateGray','dkTurquoise','dkViolet','deepPink','deepSkyBlue','dimGray','dodgerBlue','firebrick','floralWhite','forestGreen','fuchsia',
'gainsboro','ghostWhite','gold','goldenrod','gray','green','greenYellow','honeydew','hotPink','indianRed','indigo','ivory','khaki','lavender','lavenderBlush','lawnGreen','lemonChiffon','ltBlue','ltCoral','ltCyan','ltGoldenrodYellow','ltGray','ltGreen','ltPink','ltSalmon','ltSeaGreen','ltSkyBlue','ltSlateGray','ltSteelBlue','ltYellow','lime','limeGreen','linen','magenta','maroon','medAquamarine','medBlue','medOrchid','medPurple','medSeaGreen','medSlateBlue','medSpringGreen','medTurquoise','medVioletRed','midnightBlue',
'mintCream','mistyRose','moccasin','navajoWhite','navy','oldLace','olive','oliveDrab','orange','orangeRed','orchid','paleGoldenrod','paleGreen','paleTurquoise','paleVioletRed','papayaWhip','peachPuff','peru','pink','plum','powderBlue','purple','red','rosyBrown','royalBlue','saddleBrown','salmon','sandyBrown','seaGreen','seaShell','sienna','silver','skyBlue','slateBlue','slateGray','snow','springGreen','steelBlue','tan','teal','thistle','tomato','turquoise','violet','wheat','white','whiteSmoke','yellow','yellowGreen');
type TST_SystemColorVal =  (stscvScrollBar,stscvBackground,stscvActiveCaption,stscvInactiveCaption,stscvMenu,stscvWindow,stscvWindowFrame,stscvMenuText,stscvWindowText,stscvCaptionText,stscvActiveBorder,stscvInactiveBorder,stscvAppWorkspace,stscvHighlight,stscvHighlightText,stscvBtnFace,stscvBtnShadow,stscvGrayText,stscvBtnText,stscvInactiveCaptionText,stscvBtnHighlight,stscv3dDkShadow,stscv3dLight,stscvInfoText,stscvInfoBk,stscvHotLight,stscvGradientActiveCaption,stscvGradientInactiveCaption,stscvMenuHighlight,stscvMenuBar);
const StrTST_SystemColorVal: array[0..29] of AxUCString = ('scrollBar','background','activeCaption','inactiveCaption','menu','window','windowFrame','menuText','windowText','captionText','activeBorder','inactiveBorder','appWorkspace','highlight','highlightText','btnFace','btnShadow','grayText','btnText','inactiveCaptionText','btnHighlight','3dDkShadow','3dLight','infoText','infoBk','hotLight','gradientActiveCaption','gradientInactiveCaption','menuHighlight','menuBar');
type TST_RectAlignment =  (straTl,straT,straTr,straL,straCtr,straR,straBl,straB,straBr);
const StrTST_RectAlignment: array[0..8] of AxUCString = ('tl','t','tr','l','ctr','r','bl','b','br');
type TST_PathShadeType =  (stpstShape,stpstCircle,stpstRect);
const StrTST_PathShadeType: array[0..2] of AxUCString = ('shape','circle','rect');
type TST_PresetShadowVal =  (stpsvShdw1,stpsvShdw2,stpsvShdw3,stpsvShdw4,stpsvShdw5,stpsvShdw6,stpsvShdw7,stpsvShdw8,stpsvShdw9,stpsvShdw10,stpsvShdw11,stpsvShdw12,stpsvShdw13,stpsvShdw14,stpsvShdw15,stpsvShdw16,stpsvShdw17,stpsvShdw18,stpsvShdw19,stpsvShdw20);
const StrTST_PresetShadowVal: array[0..19] of AxUCString = ('shdw1','shdw2','shdw3','shdw4','shdw5','shdw6','shdw7','shdw8','shdw9','shdw10','shdw11','shdw12','shdw13','shdw14','shdw15','shdw16','shdw17','shdw18','shdw19','shdw20');
type TST_PresetPatternVal =  (stppvPct5,stppvPct10,stppvPct20,stppvPct25,stppvPct30,stppvPct40,stppvPct50,stppvPct60,stppvPct70,stppvPct75,stppvPct80,stppvPct90,stppvHorz,stppvVert,stppvLtHorz,stppvLtVert,stppvDkHorz,stppvDkVert,stppvNarHorz,stppvNarVert,stppvDashHorz,stppvDashVert,stppvCross,stppvDnDiag,stppvUpDiag,stppvLtDnDiag,stppvLtUpDiag,stppvDkDnDiag,stppvDkUpDiag,stppvWdDnDiag,stppvWdUpDiag,stppvDashDnDiag,stppvDashUpDiag,stppvDiagCross,stppvSmCheck,stppvLgCheck,stppvSmGrid,stppvLgGrid,stppvDotGrid,stppvSmConfetti,stppvLgConfetti,stppvHorzBrick,stppvDiagBrick,stppvSolidDmnd,stppvOpenDmnd,stppvDotDmnd,stppvPlaid,stppvSphere,stppvWeave,stppvDivot,stppvShingle,stppvWave,stppvTrellis,stppvZigZag);
const StrTST_PresetPatternVal: array[0..53] of AxUCString = ('pct5','pct10','pct20','pct25','pct30','pct40','pct50','pct60','pct70','pct75','pct80','pct90','horz','vert','ltHorz','ltVert','dkHorz','dkVert','narHorz','narVert','dashHorz','dashVert','cross','dnDiag','upDiag','ltDnDiag','ltUpDiag','dkDnDiag','dkUpDiag','wdDnDiag','wdUpDiag','dashDnDiag','dashUpDiag','diagCross','smCheck','lgCheck','smGrid','lgGrid','dotGrid','smConfetti','lgConfetti','horzBrick','diagBrick','solidDmnd','openDmnd','dotDmnd','plaid','sphere','weave','divot','shingle','wave','trellis','zigZag');
type TST_TileFlipMode =  (sttfmNone,sttfmX,sttfmY,sttfmXy);
const StrTST_TileFlipMode: array[0..3] of AxUCString = ('none','x','y','xy');
type TST_BlipCompression =  (stbcEmail,stbcScreen,stbcPrint,stbcHqprint,stbcNone);
const StrTST_BlipCompression: array[0..4] of AxUCString = ('email','screen','print','hqprint','none');
type TST_EffectContainerType =  (stectSib,stectTree);
const StrTST_EffectContainerType: array[0..1] of AxUCString = ('sib','tree');
type TST_BlendMode =  (stbmOver,stbmMult,stbmScreen,stbmDarken,stbmLighten);
const StrTST_BlendMode: array[0..4] of AxUCString = ('over','mult','screen','darken','lighten');
type TST_ShapeType =  (ststLine,ststLineInv,ststTriangle,ststRtTriangle,ststRect,ststDiamond,ststParallelogram,ststTrapezoid,ststNonIsoscelesTrapezoid,ststPentagon,ststHexagon,ststHeptagon,ststOctagon,ststDecagon,ststDodecagon,ststStar4,ststStar5,ststStar6,ststStar7,ststStar8,ststStar10,ststStar12,ststStar16,ststStar24,ststStar32,ststRoundRect,ststRound1Rect,ststRound2SameRect,ststRound2DiagRect,ststSnipRoundRect,ststSnip1Rect,ststSnip2SameRect,ststSnip2DiagRect,ststPlaque,ststEllipse,ststTeardrop,ststHomePlate,ststChevron,ststPieWedge,ststPie,ststBlockArc,ststDonut,
ststNoSmoking,ststRightArrow,ststLeftArrow,ststUpArrow,ststDownArrow,ststStripedRightArrow,ststNotchedRightArrow,ststBentUpArrow,ststLeftRightArrow,ststUpDownArrow,ststLeftUpArrow,ststLeftRightUpArrow,ststQuadArrow,ststLeftArrowCallout,ststRightArrowCallout,ststUpArrowCallout,ststDownArrowCallout,ststLeftRightArrowCallout,ststUpDownArrowCallout,ststQuadArrowCallout,ststBentArrow,ststUturnArrow,ststCircularArrow,ststLeftCircularArrow,ststLeftRightCircularArrow,ststCurvedRightArrow,ststCurvedLeftArrow,ststCurvedUpArrow,ststCurvedDownArrow,ststSwooshArrow,ststCube,
ststCan,ststLightningBolt,ststHeart,ststSun,ststMoon,ststSmileyFace,ststIrregularSeal1,ststIrregularSeal2,ststFoldedCorner,ststBevel,ststFrame,ststHalfFrame,ststCorner,ststDiagStripe,ststChord,ststArc,ststLeftBracket,ststRightBracket,ststLeftBrace,ststRightBrace,ststBracketPair,ststBracePair,ststStraightConnector1,ststBentConnector2,ststBentConnector3,ststBentConnector4,ststBentConnector5,ststCurvedConnector2,ststCurvedConnector3,ststCurvedConnector4,ststCurvedConnector5,ststCallout1,ststCallout2,ststCallout3,ststAccentCallout1,ststAccentCallout2,ststAccentCallout3,
ststBorderCallout1,ststBorderCallout2,ststBorderCallout3,ststAccentBorderCallout1,ststAccentBorderCallout2,ststAccentBorderCallout3,ststWedgeRectCallout,ststWedgeRoundRectCallout,ststWedgeEllipseCallout,ststCloudCallout,ststCloud,ststRibbon,ststRibbon2,ststEllipseRibbon,ststEllipseRibbon2,ststLeftRightRibbon,ststVerticalScroll,ststHorizontalScroll,ststWave,ststDoubleWave,ststPlus,ststFlowChartProcess,ststFlowChartDecision,ststFlowChartInputOutput,ststFlowChartPredefinedProcess,ststFlowChartInternalStorage,ststFlowChartDocument,ststFlowChartMultidocument,ststFlowChartTerminator,
ststFlowChartPreparation,ststFlowChartManualInput,ststFlowChartManualOperation,ststFlowChartConnector,ststFlowChartPunchedCard,ststFlowChartPunchedTape,ststFlowChartSummingJunction,ststFlowChartOr,ststFlowChartCollate,ststFlowChartSort,ststFlowChartExtract,ststFlowChartMerge,ststFlowChartOfflineStorage,ststFlowChartOnlineStorage,ststFlowChartMagneticTape,ststFlowChartMagneticDisk,ststFlowChartMagneticDrum,ststFlowChartDisplay,ststFlowChartDelay,ststFlowChartAlternateProcess,ststFlowChartOffpageConnector,ststActionButtonBlank,ststActionButtonHome,ststActionButtonHelp,
ststActionButtonInformation,ststActionButtonForwardNext,ststActionButtonBackPrevious,ststActionButtonEnd,ststActionButtonBeginning,ststActionButtonReturn,ststActionButtonDocument,ststActionButtonSound,ststActionButtonMovie,ststGear6,ststGear9,ststFunnel,ststMathPlus,ststMathMinus,ststMathMultiply,ststMathDivide,ststMathEqual,ststMathNotEqual,ststCornerTabs,ststSquareTabs,ststPlaqueTabs,ststChartX,ststChartStar,ststChartPlus);
const StrTST_ShapeType: array[0..186] of AxUCString = ('line','lineInv','triangle','rtTriangle','rect','diamond','parallelogram','trapezoid','nonIsoscelesTrapezoid','pentagon','hexagon','heptagon','octagon','decagon','dodecagon','star4','star5','star6','star7','star8','star10','star12','star16','star24','star32','roundRect','round1Rect','round2SameRect','round2DiagRect','snipRoundRect','snip1Rect','snip2SameRect','snip2DiagRect','plaque','ellipse','teardrop','homePlate','chevron','pieWedge','pie','blockArc','donut','noSmoking','rightArrow','leftArrow','upArrow',
'downArrow','stripedRightArrow','notchedRightArrow','bentUpArrow','leftRightArrow','upDownArrow','leftUpArrow','leftRightUpArrow','quadArrow','leftArrowCallout','rightArrowCallout','upArrowCallout','downArrowCallout','leftRightArrowCallout','upDownArrowCallout','quadArrowCallout','bentArrow','uturnArrow','circularArrow','leftCircularArrow','leftRightCircularArrow','curvedRightArrow','curvedLeftArrow','curvedUpArrow','curvedDownArrow','swooshArrow','cube','can','lightningBolt','heart','sun','moon','smileyFace',
'irregularSeal1','irregularSeal2','foldedCorner','bevel','frame','halfFrame','corner','diagStripe','chord','arc','leftBracket','rightBracket','leftBrace','rightBrace','bracketPair','bracePair','straightConnector1','bentConnector2','bentConnector3','bentConnector4','bentConnector5','curvedConnector2','curvedConnector3','curvedConnector4','curvedConnector5','callout1','callout2','callout3','accentCallout1','accentCallout2','accentCallout3','borderCallout1','borderCallout2','borderCallout3','accentBorderCallout1',
'accentBorderCallout2','accentBorderCallout3','wedgeRectCallout','wedgeRoundRectCallout','wedgeEllipseCallout','cloudCallout','cloud','ribbon','ribbon2','ellipseRibbon','ellipseRibbon2','leftRightRibbon','verticalScroll','horizontalScroll','wave','doubleWave','plus','flowChartProcess','flowChartDecision','flowChartInputOutput','flowChartPredefinedProcess','flowChartInternalStorage','flowChartDocument','flowChartMultidocument','flowChartTerminator','flowChartPreparation','flowChartManualInput','flowChartManualOperation',
'flowChartConnector','flowChartPunchedCard','flowChartPunchedTape','flowChartSummingJunction','flowChartOr','flowChartCollate','flowChartSort','flowChartExtract','flowChartMerge','flowChartOfflineStorage','flowChartOnlineStorage','flowChartMagneticTape','flowChartMagneticDisk','flowChartMagneticDrum','flowChartDisplay','flowChartDelay','flowChartAlternateProcess','flowChartOffpageConnector','actionButtonBlank','actionButtonHome','actionButtonHelp','actionButtonInformation','actionButtonForwardNext','actionButtonBackPrevious',
'actionButtonEnd','actionButtonBeginning','actionButtonReturn','actionButtonDocument','actionButtonSound','actionButtonMovie','gear6','gear9','funnel','mathPlus','mathMinus','mathMultiply','mathDivide','mathEqual','mathNotEqual','cornerTabs','squareTabs','plaqueTabs','chartX','chartStar','chartPlus');
type TST_PathFillMode =  (stpfmNone,stpfmNorm,stpfmLighten,stpfmLightenLess,stpfmDarken,stpfmDarkenLess);
const StrTST_PathFillMode: array[0..5] of AxUCString = ('none','norm','lighten','lightenLess','darken','darkenLess');
type TST_TextShapeType =  (sttstTextNoShape,sttstTextPlain,sttstTextStop,sttstTextTriangle,sttstTextTriangleInverted,sttstTextChevron,sttstTextChevronInverted,sttstTextRingInside,sttstTextRingOutside,sttstTextArchUp,sttstTextArchDown,sttstTextCircle,sttstTextButton,sttstTextArchUpPour,sttstTextArchDownPour,sttstTextCirclePour,sttstTextButtonPour,sttstTextCurveUp,sttstTextCurveDown,sttstTextCanUp,sttstTextCanDown,sttstTextWave1,sttstTextWave2,sttstTextDoubleWave1,sttstTextWave4,sttstTextInflate,sttstTextDeflate,sttstTextInflateBottom,sttstTextDeflateBottom,sttstTextInflateTop,sttstTextDeflateTop,sttstTextDeflateInflate,sttstTextDeflateInflateDeflate,sttstTextFadeRight,sttstTextFadeLeft,sttstTextFadeUp,sttstTextFadeDown,sttstTextSlantUp,sttstTextSlantDown,sttstTextCascadeUp,sttstTextCascadeDown);
const StrTST_TextShapeType: array[0..40] of AxUCString = ('textNoShape','textPlain','textStop','textTriangle','textTriangleInverted','textChevron','textChevronInverted','textRingInside','textRingOutside','textArchUp','textArchDown','textCircle','textButton','textArchUpPour','textArchDownPour','textCirclePour','textButtonPour','textCurveUp','textCurveDown','textCanUp','textCanDown','textWave1','textWave2','textDoubleWave1','textWave4','textInflate','textDeflate','textInflateBottom','textDeflateBottom','textInflateTop','textDeflateTop','textDeflateInflate','textDeflateInflateDeflate',
'textFadeRight','textFadeLeft','textFadeUp','textFadeDown','textSlantUp','textSlantDown','textCascadeUp','textCascadeDown');
type TST_PresetCameraType =  (stpctLegacyObliqueTopLeft,stpctLegacyObliqueTop,stpctLegacyObliqueTopRight,stpctLegacyObliqueLeft,stpctLegacyObliqueFront,stpctLegacyObliqueRight,stpctLegacyObliqueBottomLeft,stpctLegacyObliqueBottom,stpctLegacyObliqueBottomRight,stpctLegacyPerspectiveTopLeft,stpctLegacyPerspectiveTop,stpctLegacyPerspectiveTopRight,stpctLegacyPerspectiveLeft,stpctLegacyPerspectiveFront,stpctLegacyPerspectiveRight,stpctLegacyPerspectiveBottomLeft,stpctLegacyPerspectiveBottom,stpctLegacyPerspectiveBottomRight,stpctOrthographicFront,stpctIsometricTopUp,stpctIsometricTopDown,
stpctIsometricBottomUp,stpctIsometricBottomDown,stpctIsometricLeftUp,stpctIsometricLeftDown,stpctIsometricRightUp,stpctIsometricRightDown,stpctIsometricOffAxis1Left,stpctIsometricOffAxis1Right,stpctIsometricOffAxis1Top,stpctIsometricOffAxis2Left,stpctIsometricOffAxis2Right,stpctIsometricOffAxis2Top,stpctIsometricOffAxis3Left,stpctIsometricOffAxis3Right,stpctIsometricOffAxis3Bottom,stpctIsometricOffAxis4Left,stpctIsometricOffAxis4Right,stpctIsometricOffAxis4Bottom,stpctObliqueTopLeft,stpctObliqueTop,stpctObliqueTopRight,stpctObliqueLeft,stpctObliqueRight,stpctObliqueBottomLeft,
stpctObliqueBottom,stpctObliqueBottomRight,stpctPerspectiveFront,stpctPerspectiveLeft,stpctPerspectiveRight,stpctPerspectiveAbove,stpctPerspectiveBelow,stpctPerspectiveAboveLeftFacing,stpctPerspectiveAboveRightFacing,stpctPerspectiveContrastingLeftFacing,stpctPerspectiveContrastingRightFacing,stpctPerspectiveHeroicLeftFacing,stpctPerspectiveHeroicRightFacing,stpctPerspectiveHeroicExtremeLeftFacing,stpctPerspectiveHeroicExtremeRightFacing,stpctPerspectiveRelaxed,stpctPerspectiveRelaxedModerately);
const StrTST_PresetCameraType: array[0..61] of AxUCString = ('legacyObliqueTopLeft','legacyObliqueTop','legacyObliqueTopRight','legacyObliqueLeft','legacyObliqueFront','legacyObliqueRight','legacyObliqueBottomLeft','legacyObliqueBottom','legacyObliqueBottomRight','legacyPerspectiveTopLeft','legacyPerspectiveTop','legacyPerspectiveTopRight','legacyPerspectiveLeft','legacyPerspectiveFront','legacyPerspectiveRight','legacyPerspectiveBottomLeft','legacyPerspectiveBottom','legacyPerspectiveBottomRight','orthographicFront','isometricTopUp','isometricTopDown','isometricBottomUp',
'isometricBottomDown','isometricLeftUp','isometricLeftDown','isometricRightUp','isometricRightDown','isometricOffAxis1Left','isometricOffAxis1Right','isometricOffAxis1Top','isometricOffAxis2Left','isometricOffAxis2Right','isometricOffAxis2Top','isometricOffAxis3Left','isometricOffAxis3Right','isometricOffAxis3Bottom','isometricOffAxis4Left','isometricOffAxis4Right','isometricOffAxis4Bottom','obliqueTopLeft','obliqueTop','obliqueTopRight','obliqueLeft','obliqueRight','obliqueBottomLeft','obliqueBottom','obliqueBottomRight',
'perspectiveFront','perspectiveLeft','perspectiveRight','perspectiveAbove','perspectiveBelow','perspectiveAboveLeftFacing','perspectiveAboveRightFacing','perspectiveContrastingLeftFacing','perspectiveContrastingRightFacing','perspectiveHeroicLeftFacing','perspectiveHeroicRightFacing','perspectiveHeroicExtremeLeftFacing','perspectiveHeroicExtremeRightFacing','perspectiveRelaxed','perspectiveRelaxedModerately');
type TST_LightRigDirection =  (stlrdTl,stlrdT,stlrdTr,stlrdL,stlrdR,stlrdBl,stlrdB,stlrdBr);
const StrTST_LightRigDirection: array[0..7] of AxUCString = ('tl','t','tr','l','r','bl','b','br');
type TST_LightRigType =  (stlrtLegacyFlat1,stlrtLegacyFlat2,stlrtLegacyFlat3,stlrtLegacyFlat4,stlrtLegacyNormal1,stlrtLegacyNormal2,stlrtLegacyNormal3,stlrtLegacyNormal4,stlrtLegacyHarsh1,stlrtLegacyHarsh2,stlrtLegacyHarsh3,stlrtLegacyHarsh4,stlrtThreePt,stlrtBalanced,stlrtSoft,stlrtHarsh,stlrtFlood,stlrtContrasting,stlrtMorning,stlrtSunrise,stlrtSunset,stlrtChilly,stlrtFreezing,stlrtFlat,stlrtTwoPt,stlrtGlow,stlrtBrightRoom);
const StrTST_LightRigType: array[0..26] of AxUCString = ('legacyFlat1','legacyFlat2','legacyFlat3','legacyFlat4','legacyNormal1','legacyNormal2','legacyNormal3','legacyNormal4','legacyHarsh1','legacyHarsh2','legacyHarsh3','legacyHarsh4','threePt','balanced','soft','harsh','flood','contrasting','morning','sunrise','sunset','chilly','freezing','flat','twoPt','glow','brightRoom');
type TST_PenAlignment =  (stpaCtr,stpaIn);
const StrTST_PenAlignment: array[0..1] of AxUCString = ('ctr','in');
type TST_LineEndType =  (stletNone,stletTriangle,stletStealth,stletDiamond,stletOval,stletArrow);
const StrTST_LineEndType: array[0..5] of AxUCString = ('none','triangle','stealth','diamond','oval','arrow');
type TST_LineEndWidth =  (stlewSm,stlewMed,stlewLg);
const StrTST_LineEndWidth: array[0..2] of AxUCString = ('sm','med','lg');
type TST_LineCap =  (stlcRnd,stlcSq,stlcFlat);
const StrTST_LineCap: array[0..2] of AxUCString = ('rnd','sq','flat');
type TST_PresetLineDashVal =  (stpldvSolid,stpldvDot,stpldvDash,stpldvLgDash,stpldvDashDot,stpldvLgDashDot,stpldvLgDashDotDot,stpldvSysDash,stpldvSysDot,stpldvSysDashDot,stpldvSysDashDotDot);
const StrTST_PresetLineDashVal: array[0..10] of AxUCString = ('solid','dot','dash','lgDash','dashDot','lgDashDot','lgDashDotDot','sysDash','sysDot','sysDashDot','sysDashDotDot');
type TST_LineEndLength =  (stlelSm,stlelMed,stlelLg);
const StrTST_LineEndLength: array[0..2] of AxUCString = ('sm','med','lg');
type TST_CompoundLine =  (stclSng,stclDbl,stclThickThin,stclThinThick,stclTri);
const StrTST_CompoundLine: array[0..4] of AxUCString = ('sng','dbl','thickThin','thinThick','tri');
type TST_BevelPresetType =  (stbptRelaxedInset,stbptCircle,stbptSlope,stbptCross,stbptAngle,stbptSoftRound,stbptConvex,stbptCoolSlant,stbptDivot,stbptRiblet,stbptHardEdge,stbptArtDeco);
const StrTST_BevelPresetType: array[0..11] of AxUCString = ('relaxedInset','circle','slope','cross','angle','softRound','convex','coolSlant','divot','riblet','hardEdge','artDeco');
type TST_PresetMaterialType =  (stpmtLegacyMatte,stpmtLegacyPlastic,stpmtLegacyMetal,stpmtLegacyWireframe,stpmtMatte,stpmtPlastic,stpmtMetal,stpmtWarmMatte,stpmtTranslucentPowder,stpmtPowder,stpmtDkEdge,stpmtSoftEdge,stpmtClear,stpmtFlat,stpmtSoftmetal);
const StrTST_PresetMaterialType: array[0..14] of AxUCString = ('legacyMatte','legacyPlastic','legacyMetal','legacyWireframe','matte','plastic','metal','warmMatte','translucentPowder','powder','dkEdge','softEdge','clear','flat','softmetal');
type TST_TextCapsType =  (sttctNone,sttctSmall,sttctAll);
const StrTST_TextCapsType: array[0..2] of AxUCString = ('none','small','all');
type TST_TextStrikeType =  (sttstNoStrike,sttstSngStrike,sttstDblStrike);
const StrTST_TextStrikeType: array[0..2] of AxUCString = ('noStrike','sngStrike','dblStrike');
type TST_TextUnderlineType =  (sttutNone,sttutWords,sttutSng,sttutDbl,sttutHeavy,sttutDotted,sttutDottedHeavy,sttutDash,sttutDashHeavy,sttutDashLong,sttutDashLongHeavy,sttutDotDash,sttutDotDashHeavy,sttutDotDotDash,sttutDotDotDashHeavy,sttutWavy,sttutWavyHeavy,sttutWavyDbl);
const StrTST_TextUnderlineType: array[0..17] of AxUCString = ('none','words','sng','dbl','heavy','dotted','dottedHeavy','dash','dashHeavy','dashLong','dashLongHeavy','dotDash','dotDashHeavy','dotDotDash','dotDotDashHeavy','wavy','wavyHeavy','wavyDbl');
type TST_TextAutonumberScheme =  (sttasAlphaLcParenBoth,sttasAlphaUcParenBoth,sttasAlphaLcParenR,sttasAlphaUcParenR,sttasAlphaLcPeriod,sttasAlphaUcPeriod,sttasArabicParenBoth,sttasArabicParenR,sttasArabicPeriod,sttasArabicPlain,sttasRomanLcParenBoth,sttasRomanUcParenBoth,sttasRomanLcParenR,sttasRomanUcParenR,sttasRomanLcPeriod,sttasRomanUcPeriod,sttasCircleNumDbPlain,sttasCircleNumWdBlackPlain,sttasCircleNumWdWhitePlain,sttasArabicDbPeriod,sttasArabicDbPlain,sttasEa1ChsPeriod,sttasEa1ChsPlain,sttasEa1ChtPeriod,sttasEa1ChtPlain,sttasEa1JpnChsDbPeriod,sttasEa1JpnKorPlain,sttasEa1JpnKorPeriod,
sttasArabic1Minus,sttasArabic2Minus,sttasHebrew2Minus,sttasThaiAlphaPeriod,sttasThaiAlphaParenR,sttasThaiAlphaParenBoth,sttasThaiNumPeriod,sttasThaiNumParenR,sttasThaiNumParenBoth,sttasHindiAlphaPeriod,sttasHindiNumPeriod,sttasHindiNumParenR,sttasHindiAlpha1Period);
const StrTST_TextAutonumberScheme: array[0..40] of AxUCString = ('alphaLcParenBoth','alphaUcParenBoth','alphaLcParenR','alphaUcParenR','alphaLcPeriod','alphaUcPeriod','arabicParenBoth','arabicParenR','arabicPeriod','arabicPlain','romanLcParenBoth','romanUcParenBoth','romanLcParenR','romanUcParenR','romanLcPeriod','romanUcPeriod','circleNumDbPlain','circleNumWdBlackPlain','circleNumWdWhitePlain','arabicDbPeriod','arabicDbPlain','ea1ChsPeriod','ea1ChsPlain','ea1ChtPeriod','ea1ChtPlain','ea1JpnChsDbPeriod','ea1JpnKorPlain','ea1JpnKorPeriod','arabic1Minus','arabic2Minus','hebrew2Minus',
'thaiAlphaPeriod','thaiAlphaParenR','thaiAlphaParenBoth','thaiNumPeriod','thaiNumParenR','thaiNumParenBoth','hindiAlphaPeriod','hindiNumPeriod','hindiNumParenR','hindiAlpha1Period');
type TST_ColorSchemeIndex =  (stcsiDk1,stcsiLt1,stcsiDk2,stcsiLt2,stcsiAccent1,stcsiAccent2,stcsiAccent3,stcsiAccent4,stcsiAccent5,stcsiAccent6,stcsiHlink,stcsiFolHlink);
const StrTST_ColorSchemeIndex: array[0..11] of AxUCString = ('dk1','lt1','dk2','lt2','accent1','accent2','accent3','accent4','accent5','accent6','hlink','folHlink');
type TST_FontCollectionIndex =  (stfciMajor,stfciMinor,stfciNone);
const StrTST_FontCollectionIndex: array[0..2] of AxUCString = ('major','minor','none');
type TST_TextAlignType =  (sttatL,sttatCtr,sttatR,sttatJust,sttatJustLow,sttatDist,sttatThaiDist);
const StrTST_TextAlignType: array[0..6] of AxUCString = ('l','ctr','r','just','justLow','dist','thaiDist');
type TST_TextFontAlignType =  (sttfatAuto,sttfatT,sttfatCtr,sttfatBase,sttfatB);
const StrTST_TextFontAlignType: array[0..4] of AxUCString = ('auto','t','ctr','base','b');
type TST_TextTabAlignType =  (stttatL,stttatCtr,stttatR,stttatDec);
const StrTST_TextTabAlignType: array[0..3] of AxUCString = ('l','ctr','r','dec');
type TST_TextAnchoringType =  (sttatT,sttatCtr1,sttatB,sttatJust1,sttatDist1);
const StrTST_TextAnchoringType: array[0..4] of AxUCString = ('t','ctr','b','just','dist');
type TST_TextVerticalType =  (sttvtHorz,sttvtVert,sttvtVert270,sttvtWordArtVert,sttvtEaVert,sttvtMongolianVert,sttvtWordArtVertRtl);
const StrTST_TextVerticalType: array[0..6] of AxUCString = ('horz','vert','vert270','wordArtVert','eaVert','mongolianVert','wordArtVertRtl');
type TST_TextWrappingType =  (sttwtNone,sttwtSquare);
const StrTST_TextWrappingType: array[0..1] of AxUCString = ('none','square');
type TST_TextHorzOverflowType =  (stthotOverflow,stthotClip);
const StrTST_TextHorzOverflowType: array[0..1] of AxUCString = ('overflow','clip');
type TST_TextVertOverflowType =  (sttvotOverflow,sttvotEllipsis,sttvotClip);
const StrTST_TextVertOverflowType: array[0..2] of AxUCString = ('overflow','ellipsis','clip');
type PST_AdjAngle = ^TST_AdjAngle;
     TST_AdjAngle = record
     Val2: AxUCString;
     case nVal: integer of
       0: (Val1: integer;);
       1: (Dummy2: byte;);
     end;

type PST_AdjCoordinate = ^TST_AdjCoordinate;
     TST_AdjCoordinate = record
     Val2: AxUCString;
     case nVal: integer of
       0: (Val1: integer;);
       1: (Dummy2: byte;);
     end;


type TXPGDocBase = class(TObject)
protected
     FErrors   : TXpgPErrors;
     FGrManager: TXc12GraphicManager;
     FNS       : AxUCString;
public
     constructor Create; overload;
     constructor Create(AGrManager: TXc12GraphicManager); overload;

     property Errors: TXpgPErrors read FErrors;
     property GrManager: TXc12GraphicManager read FGrManager;
     property NS       : AxUCString read FNS;
     end;

     TXPGBase = class(TObject)
protected
     FOwner: TXPGDocBase;
     FElementCount: integer;
     FAttributeCount: integer;
     FAssigneds: TXpgAssigneds;
     FTag      : integer;

     function  CheckAssigned: integer; virtual;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; virtual;
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); virtual;
     procedure AfterTag; virtual;

     class function  StrToEnum(AValue: AxUCString): integer;
     class function  StrToEnumDef(AValue: AxUCString; ADefault: integer): integer;
     class function  TryStrToEnum(AValue: AxUCString; AText: AxUCString; AEnumNames: array of AxUCString; APtrInt: PInteger): boolean;
public
     function  Available: boolean;
     function  Assigned: boolean;

     property ElementCount: integer read FElementCount write FElementCount;
     property AttributeCount: integer read FAttributeCount write FAttributeCount;
     property Assigneds: TXpgAssigneds read FAssigneds write FAssigneds;
     property Owner: TXPGDocBase read FOwner;
     property Tag: integer read FTag write FTag;
     end;

     TXPGBaseList = class(TObjectList)
private
     function GetItems(Index: integer): TXPGBase;
protected
public
     constructor Create;

     function  Add(AItem: TXPGBase): TXPGBase;
     procedure CheckAssigned;

     property Items[Index: integer]: TXPGBase read GetItems; default;
     end;

     TXPGAnyElement = class(TXPGBase)
protected
     FElementName: AxUCString;
     FContent: AxUCString;
     FAttributes: TXpgXMLAttributeList;
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create;
     destructor Destroy; override;

     property ElementName: AxUCString read FElementName write FElementName;
     property Content: AxUCString read FContent write FContent;
     property Attributes: TXpgXMLAttributeList read FAttributes;
     end;

     TXPGAnyElements = class(TObjectList)
protected
     function  GetItems(Index: integer): TXPGAnyElement;
public
     function  Add(AElementName: AxUCString; AContent: AxUCString): TXPGAnyElement;
     procedure Write(AWriter: TXpgWriteXML);
     procedure Assign(AItem: TXPGAnyElements);

     property Items[Index: integer]: TXPGAnyElement read GetItems; default;
     end;

     TXPGBaseObjectList = class(TObjectList)
protected
     FOwner: TXPGDocBase;
     FAssigned: boolean;
     function  GetItems(Index: integer): TXPGBase;
public
     constructor Create(AOwner: TXPGDocBase);
     property Items[Index: integer]: TXPGBase read GetItems;
     end;

     TXPSAlternateContent = class(TXPGBase)
protected
     FCurrent: TXPGBase;
     FInFallback: boolean;
public
     procedure BeginHandle(ACurrent: TXPGBase);

     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     end;

     TXPGReader = class(TXpgReadXML)
protected
     FCurrent: TXPGBase;
     FStack: TObjectStack;
     FErrors: TXpgPErrors;
     FAlternateContent: TXPSAlternateContent;
public
     constructor Create(AErrors: TXpgPErrors; ARoot: TXPGBase);
     destructor Destroy; override;

     procedure BeginTag; override;
     procedure EndTag; override;

     property Errors: TXpgPErrors read FErrors;
     end;

     TCT_EffectContainer = class;
     TCT_EffectContainerXpgList = class;
     TCT_AlphaModulateEffect = class;
     TCT_AlphaModulateEffectXpgList = class;
     TEG_Effect = class;
     TCT_BlipFillProperties = class;
     TCT_BlendEffect = class;
     TCT_BlendEffectXpgList = class;
     TCT_FillOverlayEffect = class;
     TCT_FillOverlayEffectXpgList = class;
     TCT_FillEffect = class;
     TCT_FillEffectXpgList = class;

     TEG_ColorTransform = class(TXPGBase)
private
     function  GetAsDouble(const AName: AxUCString): double;
     procedure SetAsDouble(const AName: AxUCString; const Value: double);
     function  GetAsInteger(const AName: AxUCString): integer;
     procedure SetAsInteger(const AName: AxUCString; const Value: integer);
protected
     FItems: TStringList;

     procedure ReadAddValue(const AName: AxUCString; AValue: integer);
     procedure GetNameValue(const AIndex: integer; out AName: AxUCString; out AValue: integer);
     function  HashA(const AStr: AxUCString): integer;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;

     function  Find(const AName: AxUCString): integer;

     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure Clear;
     procedure Assign(AItem: TEG_ColorTransform);
     procedure CopyTo(AItem: TEG_ColorTransform);

     function  Apply(ARGB: longword): longword;

     property AsDouble[const AName: AxUCString]: double read GetAsDouble write SetAsDouble;
     property AsInteger[const AName: AxUCString]: integer read GetAsInteger write SetAsInteger;
     end;

     TCT_ScRgbColor = class(TXPGBase)
private
     function  GetRGB: longword;
     procedure SetRGB(const Value: longword);
protected
     FR: integer;
     FG: integer;
     FB: integer;
     FA_EG_ColorTransform: TEG_ColorTransform;
public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_ScRgbColor);
     procedure CopyTo(AItem: TCT_ScRgbColor);

     property RGB: longword read GetRGB write SetRGB;

     property R: integer read FR write FR;
     property G: integer read FG write FG;
     property B: integer read FB write FB;
     property ColorTransform: TEG_ColorTransform read FA_EG_ColorTransform;
     end;

     TCT_SRgbColor = class(TXPGBase)
protected
     FVal: integer;
     FA_EG_ColorTransform: TEG_ColorTransform;

public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     procedure Clear;
     procedure Assign(AItem: TCT_SRgbColor);
     procedure CopyTo(AItem: TCT_SRgbColor);

     property Val: integer read FVal write FVal;
     property ColorTransform: TEG_ColorTransform read FA_EG_ColorTransform;
     end;

     TCT_HslColor = class(TXPGBase)
private
     function  GetHSL: longword;
     procedure SetHSL(const Value: longword);
protected
     FHue: integer;
     FSat: integer;
     FLum: integer;
     FA_EG_ColorTransform: TEG_ColorTransform;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_HslColor);
     procedure CopyTo(AItem: TCT_HslColor);

     property Hue: integer read FHue write FHue;
     property Sat: integer read FSat write FSat;
     property Lum: integer read FLum write FLum;
     property HSL: longword read GetHSL write SetHSL;
     property ColorTransform: TEG_ColorTransform read FA_EG_ColorTransform;
     end;

     TCT_SystemColor = class(TXPGBase)
protected
     FVal: TST_SystemColorVal;
     FLastClr: integer;
     FA_EG_ColorTransform: TEG_ColorTransform;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_SystemColor);
     procedure CopyTo(AItem: TCT_SystemColor);

     property Val: TST_SystemColorVal read FVal write FVal;
     property LastClr: integer read FLastClr write FLastClr;
     property ColorTransform: TEG_ColorTransform read FA_EG_ColorTransform;
     end;

     TCT_SchemeColor = class(TXPGBase)
protected
     FVal: TST_SchemeColorVal;
     FA_EG_ColorTransform: TEG_ColorTransform;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_SchemeColor);
     procedure CopyTo(AItem: TCT_SchemeColor);

     property Val: TST_SchemeColorVal read FVal write FVal;
     property ColorTransform: TEG_ColorTransform read FA_EG_ColorTransform;
     end;

     TCT_PresetColor = class(TXPGBase)
protected
     FVal: TST_PresetColorVal;
     FA_EG_ColorTransform: TEG_ColorTransform;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_PresetColor);
     procedure CopyTo(AItem: TCT_PresetColor);

     property Val: TST_PresetColorVal read FVal write FVal;
     property ColorTransform: TEG_ColorTransform read FA_EG_ColorTransform;
     end;

     TEG_ColorChoice = class(TXPGBase)
protected
     FA_ScrgbClr: TCT_ScRgbColor;
     FA_SrgbClr: TCT_SRgbColor;
     FA_HslClr: TCT_HslColor;
     FA_SysClr: TCT_SystemColor;
     FA_SchemeClr: TCT_SchemeColor;
     FA_PrstClr: TCT_PresetColor;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TEG_ColorChoice);
     procedure CopyTo(AItem: TEG_ColorChoice);
     function  Create_ScrgbClr: TCT_ScRgbColor;
     function  Create_SrgbClr: TCT_SRgbColor;
     function  Create_HslClr: TCT_HslColor;
     function  Create_SysClr: TCT_SystemColor;
     function  Create_SchemeClr: TCT_SchemeColor;
     function  Create_PrstClr: TCT_PresetColor;

     property ScrgbClr: TCT_ScRgbColor read FA_ScrgbClr;
     property SrgbClr: TCT_SRgbColor read FA_SrgbClr;
     property HslClr: TCT_HslColor read FA_HslClr;
     property SysClr: TCT_SystemColor read FA_SysClr;
     property SchemeClr: TCT_SchemeColor read FA_SchemeClr;
     property PrstClr: TCT_PresetColor read FA_PrstClr;
     end;

     TCT_OfficeArtExtension = class(TXPGBase)
protected
     FUri: AxUCString;
     FAnyElements: TXPGAnyElements;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_OfficeArtExtension);
     procedure CopyTo(AItem: TCT_OfficeArtExtension);

     property Uri: AxUCString read FUri write FUri;
     property AnyElements: TXPGAnyElements read FAnyElements;
     end;

     TCT_OfficeArtExtensionXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_OfficeArtExtension;
public
     function  Add: TCT_OfficeArtExtension;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_OfficeArtExtensionXpgList);
     procedure CopyTo(AItem: TCT_OfficeArtExtensionXpgList);
     property Items[Index: integer]: TCT_OfficeArtExtension read GetItems; default;
     end;

     TCT_RelativeRect = class(TXPGBase)
protected
     FL: integer;
     FT: integer;
     FR: integer;
     FB: integer;
public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_RelativeRect);
     procedure CopyTo(AItem: TCT_RelativeRect);

     function  Equal(const ALeft,ARight,ATop,ABottom: integer): boolean;
     function  SetValues(const ALeft,ARight,ATop,ABottom: integer): boolean;

     property L: integer read FL write FL;
     property T: integer read FT write FT;
     property R: integer read FR write FR;
     property B: integer read FB write FB;
     end;

     TCT_Color = class(TXPGBase)
protected
     FA_EG_ColorChoice: TEG_ColorChoice;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Color);
     procedure CopyTo(AItem: TCT_Color);

     property A_EG_ColorChoice: TEG_ColorChoice read FA_EG_ColorChoice;
     end;

     TEG_OfficeArtExtensionList = class(TXPGBase)
protected
     FA_Ext: TCT_OfficeArtExtension;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TEG_OfficeArtExtensionList);
     procedure CopyTo(AItem: TEG_OfficeArtExtensionList);
     function  Create_A_Ext: TCT_OfficeArtExtension;

     property A_Ext: TCT_OfficeArtExtension read FA_Ext;
     end;

     TCT_GradientStop = class(TXPGBase)
protected
     FPos: integer;
     FA_EG_ColorChoice: TEG_ColorChoice;
public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_GradientStop);
     procedure CopyTo(AItem: TCT_GradientStop);

     property Pos: integer read FPos write FPos;
     property ColorChoice: TEG_ColorChoice read FA_EG_ColorChoice;
     end;

     TCT_GradientStopXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_GradientStop;
public
     function  Add: TCT_GradientStop;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_GradientStopXpgList);
     procedure CopyTo(AItem: TCT_GradientStopXpgList);
     property Items[Index: integer]: TCT_GradientStop read GetItems; default;
     end;

     TCT_LinearShadeProperties = class(TXPGBase)
protected
     FAng: integer;
     FScaled: boolean;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_LinearShadeProperties);
     procedure CopyTo(AItem: TCT_LinearShadeProperties);

     property Ang: integer read FAng write FAng;
     property Scaled: boolean read FScaled write FScaled;
     end;

     TCT_PathShadeProperties = class(TXPGBase)
protected
     FPath: TST_PathShadeType;
     FA_FillToRect: TCT_RelativeRect;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_PathShadeProperties);
     procedure CopyTo(AItem: TCT_PathShadeProperties);
     function  Create_FillToRect: TCT_RelativeRect;
     procedure Free_FillToRect;

     property Path: TST_PathShadeType read FPath write FPath;
     property FillToRect: TCT_RelativeRect read FA_FillToRect;
     end;

     TCT_AlphaBiLevelEffect = class(TXPGBase)
protected
     FThresh: integer;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_AlphaBiLevelEffect);
     procedure CopyTo(AItem: TCT_AlphaBiLevelEffect);

     property Thresh: integer read FThresh write FThresh;
     end;

     TCT_AlphaBiLevelEffectXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_AlphaBiLevelEffect;
public
     function  Add: TCT_AlphaBiLevelEffect;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_AlphaBiLevelEffectXpgList);
     procedure CopyTo(AItem: TCT_AlphaBiLevelEffectXpgList);
     property Items[Index: integer]: TCT_AlphaBiLevelEffect read GetItems; default;
     end;

     TCT_AlphaCeilingEffect = class(TXPGBase)
public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_AlphaCeilingEffect);
     procedure CopyTo(AItem: TCT_AlphaCeilingEffect);

     end;

     TCT_AlphaCeilingEffectXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_AlphaCeilingEffect;
public
     function  Add: TCT_AlphaCeilingEffect;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_AlphaCeilingEffectXpgList);
     procedure CopyTo(AItem: TCT_AlphaCeilingEffectXpgList);
     property Items[Index: integer]: TCT_AlphaCeilingEffect read GetItems; default;
     end;

     TCT_AlphaFloorEffect = class(TXPGBase)
public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_AlphaFloorEffect);
     procedure CopyTo(AItem: TCT_AlphaFloorEffect);

     end;

     TCT_AlphaFloorEffectXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_AlphaFloorEffect;
public
     function  Add: TCT_AlphaFloorEffect;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_AlphaFloorEffectXpgList);
     procedure CopyTo(AItem: TCT_AlphaFloorEffectXpgList);
     property Items[Index: integer]: TCT_AlphaFloorEffect read GetItems; default;
     end;

     TCT_AlphaInverseEffect = class(TXPGBase)
protected
     FA_EG_ColorChoice: TEG_ColorChoice;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_AlphaInverseEffect);
     procedure CopyTo(AItem: TCT_AlphaInverseEffect);

     property ColorChoice: TEG_ColorChoice read FA_EG_ColorChoice;
     end;

     TCT_AlphaInverseEffectXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_AlphaInverseEffect;
public
     function  Add: TCT_AlphaInverseEffect;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_AlphaInverseEffectXpgList);
     procedure CopyTo(AItem: TCT_AlphaInverseEffectXpgList);
     property Items[Index: integer]: TCT_AlphaInverseEffect read GetItems; default;
     end;

     TCT_AlphaModulateFixedEffect = class(TXPGBase)
protected
     FAmt: integer;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_AlphaModulateFixedEffect);
     procedure CopyTo(AItem: TCT_AlphaModulateFixedEffect);

     property Amt: integer read FAmt write FAmt;
     end;

     TCT_AlphaModulateFixedEffectXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_AlphaModulateFixedEffect;
public
     function  Add: TCT_AlphaModulateFixedEffect;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_AlphaModulateFixedEffectXpgList);
     procedure CopyTo(AItem: TCT_AlphaModulateFixedEffectXpgList);
     property Items[Index: integer]: TCT_AlphaModulateFixedEffect read GetItems; default;
     end;

     TCT_AlphaReplaceEffect = class(TXPGBase)
protected
     FA: integer;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_AlphaReplaceEffect);
     procedure CopyTo(AItem: TCT_AlphaReplaceEffect);

     property A: integer read FA write FA;
     end;

     TCT_AlphaReplaceEffectXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_AlphaReplaceEffect;
public
     function  Add: TCT_AlphaReplaceEffect;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_AlphaReplaceEffectXpgList);
     procedure CopyTo(AItem: TCT_AlphaReplaceEffectXpgList);
     property Items[Index: integer]: TCT_AlphaReplaceEffect read GetItems; default;
     end;

     TCT_BiLevelEffect = class(TXPGBase)
protected
     FThresh: integer;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_BiLevelEffect);
     procedure CopyTo(AItem: TCT_BiLevelEffect);

     property Thresh: integer read FThresh write FThresh;
     end;

     TCT_BiLevelEffectXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_BiLevelEffect;
public
     function  Add: TCT_BiLevelEffect;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_BiLevelEffectXpgList);
     procedure CopyTo(AItem: TCT_BiLevelEffectXpgList);
     property Items[Index: integer]: TCT_BiLevelEffect read GetItems; default;
     end;

     TCT_BlurEffect = class(TXPGBase)
protected
     FRad: integer;
     FGrow: boolean;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_BlurEffect);
     procedure CopyTo(AItem: TCT_BlurEffect);

     property Rad: integer read FRad write FRad;
     property Grow: boolean read FGrow write FGrow;
     end;

     TCT_BlurEffectXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_BlurEffect;
public
     function  Add: TCT_BlurEffect;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_BlurEffectXpgList);
     procedure CopyTo(AItem: TCT_BlurEffectXpgList);
     property Items[Index: integer]: TCT_BlurEffect read GetItems; default;
     end;

     TCT_ColorChangeEffect = class(TXPGBase)
protected
     FUseA: boolean;
     FA_ClrFrom: TCT_Color;
     FA_ClrTo: TCT_Color;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_ColorChangeEffect);
     procedure CopyTo(AItem: TCT_ColorChangeEffect);
     function  Create_A_ClrFrom: TCT_Color;
     function  Create_A_ClrTo: TCT_Color;

     property UseA: boolean read FUseA write FUseA;
     property A_ClrFrom: TCT_Color read FA_ClrFrom;
     property A_ClrTo: TCT_Color read FA_ClrTo;
     end;

     TCT_ColorChangeEffectXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_ColorChangeEffect;
public
     function  Add: TCT_ColorChangeEffect;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_ColorChangeEffectXpgList);
     procedure CopyTo(AItem: TCT_ColorChangeEffectXpgList);
     property Items[Index: integer]: TCT_ColorChangeEffect read GetItems; default;
     end;

     TCT_ColorReplaceEffect = class(TXPGBase)
protected
     FA_EG_ColorChoice: TEG_ColorChoice;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_ColorReplaceEffect);
     procedure CopyTo(AItem: TCT_ColorReplaceEffect);

     property ColorChoice: TEG_ColorChoice read FA_EG_ColorChoice;
     end;

     TCT_ColorReplaceEffectXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_ColorReplaceEffect;
public
     function  Add: TCT_ColorReplaceEffect;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_ColorReplaceEffectXpgList);
     procedure CopyTo(AItem: TCT_ColorReplaceEffectXpgList);
     property Items[Index: integer]: TCT_ColorReplaceEffect read GetItems; default;
     end;

     TCT_DuotoneEffect = class(TXPGBase)
protected
     FA_EG_ColorChoice: TEG_ColorChoice;
public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_DuotoneEffect);
     procedure CopyTo(AItem: TCT_DuotoneEffect);

     property ColorChoice: TEG_ColorChoice read FA_EG_ColorChoice;
     end;

     TCT_DuotoneEffectXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_DuotoneEffect;
public
     function  Add: TCT_DuotoneEffect;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_DuotoneEffectXpgList);
     procedure CopyTo(AItem: TCT_DuotoneEffectXpgList);
     property Items[Index: integer]: TCT_DuotoneEffect read GetItems; default;
     end;

     TCT_GrayscaleEffect = class(TXPGBase)
public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_GrayscaleEffect);
     procedure CopyTo(AItem: TCT_GrayscaleEffect);

     end;

     TCT_GrayscaleEffectXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_GrayscaleEffect;
public
     function  Add: TCT_GrayscaleEffect;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_GrayscaleEffectXpgList);
     procedure CopyTo(AItem: TCT_GrayscaleEffectXpgList);
     property Items[Index: integer]: TCT_GrayscaleEffect read GetItems; default;
     end;

     TCT_HSLEffect = class(TXPGBase)
protected
     FHue: integer;
     FSat: integer;
     FLum: integer;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_HSLEffect);
     procedure CopyTo(AItem: TCT_HSLEffect);

     property Hue: integer read FHue write FHue;
     property Sat: integer read FSat write FSat;
     property Lum: integer read FLum write FLum;
     end;

     TCT_HSLEffectXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_HSLEffect;
public
     function  Add: TCT_HSLEffect;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_HSLEffectXpgList);
     procedure CopyTo(AItem: TCT_HSLEffectXpgList);
     property Items[Index: integer]: TCT_HSLEffect read GetItems; default;
     end;

     TCT_LuminanceEffect = class(TXPGBase)
protected
     FBright: integer;
     FContrast: integer;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_LuminanceEffect);
     procedure CopyTo(AItem: TCT_LuminanceEffect);

     property Bright: integer read FBright write FBright;
     property Contrast: integer read FContrast write FContrast;
     end;

     TCT_LuminanceEffectXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_LuminanceEffect;
public
     function  Add: TCT_LuminanceEffect;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_LuminanceEffectXpgList);
     procedure CopyTo(AItem: TCT_LuminanceEffectXpgList);
     property Items[Index: integer]: TCT_LuminanceEffect read GetItems; default;
     end;

     TCT_TintEffect = class(TXPGBase)
protected
     FHue: integer;
     FAmt: integer;
public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_TintEffect);
     procedure CopyTo(AItem: TCT_TintEffect);

     property Hue: integer read FHue write FHue;
     property Amt: integer read FAmt write FAmt;
     end;

     TCT_TintEffectXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_TintEffect;
public
     function  Add: TCT_TintEffect;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_TintEffectXpgList);
     procedure CopyTo(AItem: TCT_TintEffectXpgList);
     property Items[Index: integer]: TCT_TintEffect read GetItems; default;
     end;

     TCT_OfficeArtExtensionList = class(TXPGBase)
protected
     FA_EG_OfficeArtExtensionList: TEG_OfficeArtExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_OfficeArtExtensionList);
     procedure CopyTo(AItem: TCT_OfficeArtExtensionList);

     property A_EG_OfficeArtExtensionList: TEG_OfficeArtExtensionList read FA_EG_OfficeArtExtensionList;
     end;

     TCT_TileInfoProperties = class(TXPGBase)
protected
     FTx: integer;
     FTy: integer;
     FSx: integer;
     FSy: integer;
     FFlip: TST_TileFlipMode;
     FAlgn: TST_RectAlignment;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_TileInfoProperties);
     procedure CopyTo(AItem: TCT_TileInfoProperties);

     property Tx: integer read FTx write FTx;
     property Ty: integer read FTy write FTy;
     property Sx: integer read FSx write FSx;
     property Sy: integer read FSy write FSy;
     property Flip: TST_TileFlipMode read FFlip write FFlip;
     property Algn: TST_RectAlignment read FAlgn write FAlgn;
     end;

     TCT_StretchInfoProperties = class(TXPGBase)
protected
     FA_FillRect: TCT_RelativeRect;
public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_StretchInfoProperties);
     procedure CopyTo(AItem: TCT_StretchInfoProperties);
     function  Create_A_FillRect: TCT_RelativeRect;

     property FillRect: TCT_RelativeRect read FA_FillRect;
     end;

     TCT_GradientStopList = class(TXPGBase)
protected
     FA_GsXpgList: TCT_GradientStopXpgList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_GradientStopList);
     procedure CopyTo(AItem: TCT_GradientStopList);
     function  Create_A_GsXpgList: TCT_GradientStopXpgList;

     property Gs: TCT_GradientStopXpgList read FA_GsXpgList;
     end;

     TEG_ShadeProperties = class(TXPGBase)
protected
     FA_Lin: TCT_LinearShadeProperties;
     FA_Path: TCT_PathShadeProperties;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TEG_ShadeProperties);
     procedure CopyTo(AItem: TEG_ShadeProperties);
     function  Create_Lin: TCT_LinearShadeProperties;
     function  Create_Path: TCT_PathShadeProperties;

     property Lin: TCT_LinearShadeProperties read FA_Lin;
     property Path: TCT_PathShadeProperties read FA_Path;
     end;

     TCT_Blip = class(TXPGBase)
protected
     FImage: TXc12GraphicImage;

     FR_Embed: AxUCString;
     FR_Link: AxUCString;
     FCstate: TST_BlipCompression;
     FA_AlphaBiLevel: TCT_AlphaBiLevelEffect;
     FA_AlphaCeiling: TCT_AlphaCeilingEffect;
     FA_AlphaFloor: TCT_AlphaFloorEffect;
     FA_AlphaInv: TCT_AlphaInverseEffect;
     FA_AlphaMod: TCT_AlphaModulateEffect;
     FA_AlphaModFix: TCT_AlphaModulateFixedEffect;
     FA_AlphaRepl: TCT_AlphaReplaceEffect;
     FA_BiLevel: TCT_BiLevelEffect;
     FA_Blur: TCT_BlurEffect;
     FA_ClrChange: TCT_ColorChangeEffect;
     FA_ClrRepl: TCT_ColorReplaceEffect;
     FA_Duotone: TCT_DuotoneEffect;
     FA_FillOverlay: TCT_FillOverlayEffect;
     FA_Grayscl: TCT_GrayscaleEffect;
     FA_Hsl: TCT_HSLEffect;
     FA_Lum: TCT_LuminanceEffect;
     FA_Tint: TCT_TintEffect;
     FA_ExtLst: TCT_OfficeArtExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;

     procedure Clear;
     procedure Assign(AItem: TCT_Blip);
     procedure CopyTo(AItem: TCT_Blip);

     function  Create_AlphaBiLevel: TCT_AlphaBiLevelEffect;
     function  Create_AlphaCeiling: TCT_AlphaCeilingEffect;
     function  Create_AlphaFloor: TCT_AlphaFloorEffect;
     function  Create_AlphaInv: TCT_AlphaInverseEffect;
     function  Create_AlphaMod: TCT_AlphaModulateEffect;
     function  Create_AlphaModFix: TCT_AlphaModulateFixedEffect;
     function  Create_AlphaRepl: TCT_AlphaReplaceEffect;
     function  Create_BiLevel: TCT_BiLevelEffect;
     function  Create_Blur: TCT_BlurEffect;
     function  Create_ClrChange: TCT_ColorChangeEffect;
     function  Create_ClrRepl: TCT_ColorReplaceEffect;
     function  Create_Duotone: TCT_DuotoneEffect;
     function  Create_FillOverlay: TCT_FillOverlayEffect;
     function  Create_Grayscl: TCT_GrayscaleEffect;
     function  Create_Hsl: TCT_HSLEffect;
     function  Create_Lum: TCT_LuminanceEffect;
     function  Create_Tint: TCT_TintEffect;
     function  Create_A_ExtLst: TCT_OfficeArtExtensionList;

     procedure Free_Lum;

     procedure Free_AlphaModFix;

     property Image: TXc12GraphicImage read FImage write FImage;

     property R_Embed: AxUCString read FR_Embed write FR_Embed;
     property R_Link: AxUCString read FR_Link write FR_Link;
     property Cstate: TST_BlipCompression read FCstate write FCstate;
     property AlphaBiLevel: TCT_AlphaBiLevelEffect read FA_AlphaBiLevel;
     property AlphaCeiling: TCT_AlphaCeilingEffect read FA_AlphaCeiling;
     property AlphaFloor: TCT_AlphaFloorEffect read FA_AlphaFloor;
     property AlphaInv: TCT_AlphaInverseEffect read FA_AlphaInv;
     property AlphaMod: TCT_AlphaModulateEffect read FA_AlphaMod write FA_AlphaMod;
     property AlphaModFix: TCT_AlphaModulateFixedEffect read FA_AlphaModFix;
     property AlphaRepl: TCT_AlphaReplaceEffect read FA_AlphaRepl;
     property BiLevel: TCT_BiLevelEffect read FA_BiLevel;
     property Blur: TCT_BlurEffect read FA_Blur;
     property ClrChange: TCT_ColorChangeEffect read FA_ClrChange;
     property ClrRepl: TCT_ColorReplaceEffect read FA_ClrRepl;
     property Duotone: TCT_DuotoneEffect read FA_Duotone;
     property FillOverlay: TCT_FillOverlayEffect read FA_FillOverlay write FA_FillOverlay;
     property Grayscl: TCT_GrayscaleEffect read FA_Grayscl;
     property Hsl: TCT_HSLEffect read FA_Hsl;
     property Lum: TCT_LuminanceEffect read FA_Lum;
     property Tint: TCT_TintEffect read FA_Tint;
     property A_ExtLst: TCT_OfficeArtExtensionList read FA_ExtLst;
     end;

     TEG_FillModeProperties = class(TXPGBase)
protected
     FA_Tile: TCT_TileInfoProperties;
     FA_Stretch: TCT_StretchInfoProperties;
public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TEG_FillModeProperties);
     procedure CopyTo(AItem: TEG_FillModeProperties);
     function  Create_Tile: TCT_TileInfoProperties;
     function  Create_Stretch: TCT_StretchInfoProperties;

     property Tile: TCT_TileInfoProperties read FA_Tile;
     property Stretch: TCT_StretchInfoProperties read FA_Stretch;
     end;

     TCT_NoFillProperties = class(TXPGBase)
public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_NoFillProperties);
     procedure CopyTo(AItem: TCT_NoFillProperties);

     end;

     TCT_SolidColorFillProperties = class(TXPGBase)
protected
     FA_EG_ColorChoice: TEG_ColorChoice;
public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_SolidColorFillProperties);
     procedure CopyTo(AItem: TCT_SolidColorFillProperties);

     property ColorChoice: TEG_ColorChoice read FA_EG_ColorChoice;
     end;

     TCT_GradientFillProperties = class(TXPGBase)
protected
     FFlip: TST_TileFlipMode;
     FRotWithShape: boolean;
     FA_GsLst: TCT_GradientStopList;
     FA_EG_ShadeProperties: TEG_ShadeProperties;
     FA_TileRect: TCT_RelativeRect;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     procedure Clear;
     procedure Assign(AItem: TCT_GradientFillProperties);
     procedure CopyTo(AItem: TCT_GradientFillProperties);
     function  Create_GstLst: TCT_GradientStopList;
     function  Create_TileRect: TCT_RelativeRect;
     procedure Free_TileRect;

     property Flip: TST_TileFlipMode read FFlip write FFlip;
     property RotWithShape: boolean read FRotWithShape write FRotWithShape;
     property GsLst: TCT_GradientStopList read FA_GsLst;
     property ShadeProperties: TEG_ShadeProperties read FA_EG_ShadeProperties;
     property TileRect: TCT_RelativeRect read FA_TileRect;
     end;

     TCT_BlipFillProperties = class(TXPGBase)
protected
     FDpi: integer;
     FRotWithShape: boolean;
     FA_Blip: TCT_Blip;
     FA_SrcRect: TCT_RelativeRect;
     FA_EG_FillModeProperties: TEG_FillModeProperties;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_BlipFillProperties);
     procedure CopyTo(AItem: TCT_BlipFillProperties);
     function  Create_Blip: TCT_Blip;
     function  Create_SrcRect: TCT_RelativeRect;

     property Dpi: integer read FDpi write FDpi;
     property RotWithShape: boolean read FRotWithShape write FRotWithShape;
     property Blip: TCT_Blip read FA_Blip;
     property SrcRect: TCT_RelativeRect read FA_SrcRect;
     property FillModeProperties: TEG_FillModeProperties read FA_EG_FillModeProperties;
     end;

     TCT_PatternFillProperties = class(TXPGBase)
protected
     FPrst: TST_PresetPatternVal;
     FA_FgClr: TCT_Color;
     FA_BgClr: TCT_Color;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_PatternFillProperties);
     procedure CopyTo(AItem: TCT_PatternFillProperties);
     function  Create_A_FgClr: TCT_Color;
     function  Create_A_BgClr: TCT_Color;

     property Prst: TST_PresetPatternVal read FPrst write FPrst;
     property A_FgClr: TCT_Color read FA_FgClr;
     property A_BgClr: TCT_Color read FA_BgClr;
     end;

     TCT_GroupFillProperties = class(TXPGBase)
public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_GroupFillProperties);
     procedure CopyTo(AItem: TCT_GroupFillProperties);

     end;

     TEG_FillProperties = class(TXPGBase)
private
     function  GetAutomatic: boolean;
     procedure SetAutomatic(const Value: boolean);
protected
     FA_NoFill: TCT_NoFillProperties;
     FA_SolidFill: TCT_SolidColorFillProperties;
     FA_GradFill: TCT_GradientFillProperties;
     FA_BlipFill: TCT_BlipFillProperties;
     FA_PattFill: TCT_PatternFillProperties;
     FA_GrpFill: TCT_GroupFillProperties;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TEG_FillProperties);
     procedure CopyTo(AItem: TEG_FillProperties);
     function  Create_NoFill: TCT_NoFillProperties;
     function  Create_SolidFill: TCT_SolidColorFillProperties;
     function  Create_GradFill: TCT_GradientFillProperties;
     function  Create_BlipFill: TCT_BlipFillProperties;
     function  Create_PattFill: TCT_PatternFillProperties;
     function  Create_GrpFill: TCT_GroupFillProperties;

     procedure Remove_NoFill;

     property NoFill: TCT_NoFillProperties read FA_NoFill;
     property SolidFill: TCT_SolidColorFillProperties read FA_SolidFill;
     property GradFill: TCT_GradientFillProperties read FA_GradFill;
     property BlipFill: TCT_BlipFillProperties read FA_BlipFill;
     property PattFill: TCT_PatternFillProperties read FA_PattFill;
     property GrpFill: TCT_GroupFillProperties read FA_GrpFill;
     property Automatic: boolean read GetAutomatic write SetAutomatic;
     end;

     TCT_FillOverlayEffect = class(TXPGBase)
protected
     FBlend: TST_BlendMode;
     FA_EG_FillProperties: TEG_FillProperties;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_FillOverlayEffect);
     procedure CopyTo(AItem: TCT_FillOverlayEffect);

     property Blend: TST_BlendMode read FBlend write FBlend;
     property A_EG_FillProperties: TEG_FillProperties read FA_EG_FillProperties;
     end;

     TCT_FillOverlayEffectXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_FillOverlayEffect;
public
     function  Add: TCT_FillOverlayEffect;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_FillOverlayEffectXpgList);
     procedure CopyTo(AItem: TCT_FillOverlayEffectXpgList);
     property Items[Index: integer]: TCT_FillOverlayEffect read GetItems; default;
     end;

     TCT_EffectReference = class(TXPGBase)
protected
     FRef: AxUCString;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_EffectReference);
     procedure CopyTo(AItem: TCT_EffectReference);

     property Ref: AxUCString read FRef write FRef;
     end;

     TCT_EffectReferenceXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_EffectReference;
public
     function  Add: TCT_EffectReference;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_EffectReferenceXpgList);
     procedure CopyTo(AItem: TCT_EffectReferenceXpgList);
     property Items[Index: integer]: TCT_EffectReference read GetItems; default;
     end;

     TCT_AlphaOutsetEffect = class(TXPGBase)
protected
     FRad: integer;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_AlphaOutsetEffect);
     procedure CopyTo(AItem: TCT_AlphaOutsetEffect);

     property Rad: integer read FRad write FRad;
     end;

     TCT_AlphaOutsetEffectXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_AlphaOutsetEffect;
public
     function  Add: TCT_AlphaOutsetEffect;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_AlphaOutsetEffectXpgList);
     procedure CopyTo(AItem: TCT_AlphaOutsetEffectXpgList);
     property Items[Index: integer]: TCT_AlphaOutsetEffect read GetItems; default;
     end;

     TCT_FillEffect = class(TXPGBase)
protected
     FA_EG_FillProperties: TEG_FillProperties;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_FillEffect);
     procedure CopyTo(AItem: TCT_FillEffect);

     property A_EG_FillProperties: TEG_FillProperties read FA_EG_FillProperties;
     end;

     TCT_FillEffectXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_FillEffect;
public
     function  Add: TCT_FillEffect;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_FillEffectXpgList);
     procedure CopyTo(AItem: TCT_FillEffectXpgList);
     property Items[Index: integer]: TCT_FillEffect read GetItems; default;
     end;

     TCT_GlowEffect = class(TXPGBase)
protected
     FRad: integer;
     FA_EG_ColorChoice: TEG_ColorChoice;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_GlowEffect);
     procedure CopyTo(AItem: TCT_GlowEffect);

     property Rad: integer read FRad write FRad;
     property A_EG_ColorChoice: TEG_ColorChoice read FA_EG_ColorChoice;
     end;

     TCT_GlowEffectXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_GlowEffect;
public
     function  Add: TCT_GlowEffect;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_GlowEffectXpgList);
     procedure CopyTo(AItem: TCT_GlowEffectXpgList);
     property Items[Index: integer]: TCT_GlowEffect read GetItems; default;
     end;

     TCT_InnerShadowEffect = class(TXPGBase)
protected
     FBlurRad: integer;
     FDist: integer;
     FDir: integer;
     FA_EG_ColorChoice: TEG_ColorChoice;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_InnerShadowEffect);
     procedure CopyTo(AItem: TCT_InnerShadowEffect);

     property BlurRad: integer read FBlurRad write FBlurRad;
     property Dist: integer read FDist write FDist;
     property Dir: integer read FDir write FDir;
     property A_EG_ColorChoice: TEG_ColorChoice read FA_EG_ColorChoice;
     end;

     TCT_InnerShadowEffectXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_InnerShadowEffect;
public
     function  Add: TCT_InnerShadowEffect;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_InnerShadowEffectXpgList);
     procedure CopyTo(AItem: TCT_InnerShadowEffectXpgList);
     property Items[Index: integer]: TCT_InnerShadowEffect read GetItems; default;
     end;

     TCT_OuterShadowEffect = class(TXPGBase)
protected
     FBlurRad: integer;
     FDist: integer;
     FDir: integer;
     FSx: integer;
     FSy: integer;
     FKx: integer;
     FKy: integer;
     FAlgn: TST_RectAlignment;
     FRotWithShape: boolean;
     FA_EG_ColorChoice: TEG_ColorChoice;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_OuterShadowEffect);
     procedure CopyTo(AItem: TCT_OuterShadowEffect);

     property BlurRad: integer read FBlurRad write FBlurRad;
     property Dist: integer read FDist write FDist;
     property Dir: integer read FDir write FDir;
     property Sx: integer read FSx write FSx;
     property Sy: integer read FSy write FSy;
     property Kx: integer read FKx write FKx;
     property Ky: integer read FKy write FKy;
     property Algn: TST_RectAlignment read FAlgn write FAlgn;
     property RotWithShape: boolean read FRotWithShape write FRotWithShape;
     property A_EG_ColorChoice: TEG_ColorChoice read FA_EG_ColorChoice;
     end;

     TCT_OuterShadowEffectXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_OuterShadowEffect;
public
     function  Add: TCT_OuterShadowEffect;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_OuterShadowEffectXpgList);
     procedure CopyTo(AItem: TCT_OuterShadowEffectXpgList);
     property Items[Index: integer]: TCT_OuterShadowEffect read GetItems; default;
     end;

     TCT_PresetShadowEffect = class(TXPGBase)
protected
     FPrst: TST_PresetShadowVal;
     FDist: integer;
     FDir: integer;
     FA_EG_ColorChoice: TEG_ColorChoice;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_PresetShadowEffect);
     procedure CopyTo(AItem: TCT_PresetShadowEffect);

     property Prst: TST_PresetShadowVal read FPrst write FPrst;
     property Dist: integer read FDist write FDist;
     property Dir: integer read FDir write FDir;
     property A_EG_ColorChoice: TEG_ColorChoice read FA_EG_ColorChoice;
     end;

     TCT_PresetShadowEffectXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_PresetShadowEffect;
public
     function  Add: TCT_PresetShadowEffect;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_PresetShadowEffectXpgList);
     procedure CopyTo(AItem: TCT_PresetShadowEffectXpgList);
     property Items[Index: integer]: TCT_PresetShadowEffect read GetItems; default;
     end;

     TCT_ReflectionEffect = class(TXPGBase)
protected
     FBlurRad: integer;
     FStA: integer;
     FStPos: integer;
     FEndA: integer;
     FEndPos: integer;
     FDist: integer;
     FDir: integer;
     FFadeDir: integer;
     FSx: integer;
     FSy: integer;
     FKx: integer;
     FKy: integer;
     FAlgn: TST_RectAlignment;
     FRotWithShape: boolean;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_ReflectionEffect);
     procedure CopyTo(AItem: TCT_ReflectionEffect);

     property BlurRad: integer read FBlurRad write FBlurRad;
     property StA: integer read FStA write FStA;
     property StPos: integer read FStPos write FStPos;
     property EndA: integer read FEndA write FEndA;
     property EndPos: integer read FEndPos write FEndPos;
     property Dist: integer read FDist write FDist;
     property Dir: integer read FDir write FDir;
     property FadeDir: integer read FFadeDir write FFadeDir;
     property Sx: integer read FSx write FSx;
     property Sy: integer read FSy write FSy;
     property Kx: integer read FKx write FKx;
     property Ky: integer read FKy write FKy;
     property Algn: TST_RectAlignment read FAlgn write FAlgn;
     property RotWithShape: boolean read FRotWithShape write FRotWithShape;
     end;

     TCT_ReflectionEffectXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_ReflectionEffect;
public
     function  Add: TCT_ReflectionEffect;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_ReflectionEffectXpgList);
     procedure CopyTo(AItem: TCT_ReflectionEffectXpgList);
     property Items[Index: integer]: TCT_ReflectionEffect read GetItems; default;
     end;

     TCT_RelativeOffsetEffect = class(TXPGBase)
protected
     FTx: integer;
     FTy: integer;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_RelativeOffsetEffect);
     procedure CopyTo(AItem: TCT_RelativeOffsetEffect);

     property Tx: integer read FTx write FTx;
     property Ty: integer read FTy write FTy;
     end;

     TCT_RelativeOffsetEffectXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_RelativeOffsetEffect;
public
     function  Add: TCT_RelativeOffsetEffect;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_RelativeOffsetEffectXpgList);
     procedure CopyTo(AItem: TCT_RelativeOffsetEffectXpgList);
     property Items[Index: integer]: TCT_RelativeOffsetEffect read GetItems; default;
     end;

     TCT_SoftEdgesEffect = class(TXPGBase)
protected
     FRad: integer;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_SoftEdgesEffect);
     procedure CopyTo(AItem: TCT_SoftEdgesEffect);

     property Rad: integer read FRad write FRad;
     end;

     TCT_SoftEdgesEffectXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_SoftEdgesEffect;
public
     function  Add: TCT_SoftEdgesEffect;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_SoftEdgesEffectXpgList);
     procedure CopyTo(AItem: TCT_SoftEdgesEffectXpgList);
     property Items[Index: integer]: TCT_SoftEdgesEffect read GetItems; default;
     end;

     TCT_TransformEffect = class(TXPGBase)
protected
     FSx: integer;
     FSy: integer;
     FKx: integer;
     FKy: integer;
     FTx: integer;
     FTy: integer;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_TransformEffect);
     procedure CopyTo(AItem: TCT_TransformEffect);

     property Sx: integer read FSx write FSx;
     property Sy: integer read FSy write FSy;
     property Kx: integer read FKx write FKx;
     property Ky: integer read FKy write FKy;
     property Tx: integer read FTx write FTx;
     property Ty: integer read FTy write FTy;
     end;

     TCT_TransformEffectXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_TransformEffect;
public
     function  Add: TCT_TransformEffect;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_TransformEffectXpgList);
     procedure CopyTo(AItem: TCT_TransformEffectXpgList);
     property Items[Index: integer]: TCT_TransformEffect read GetItems; default;
     end;

     TEG_Effect = class(TXPGBase)
protected
     FA_ContXpgList: TCT_EffectContainerXpgList;
     FA_EffectXpgList: TCT_EffectReferenceXpgList;
     FA_AlphaBiLevelXpgList: TCT_AlphaBiLevelEffectXpgList;
     FA_AlphaCeilingXpgList: TCT_AlphaCeilingEffectXpgList;
     FA_AlphaFloorXpgList: TCT_AlphaFloorEffectXpgList;
     FA_AlphaInvXpgList: TCT_AlphaInverseEffectXpgList;
     FA_AlphaModXpgList: TCT_AlphaModulateEffectXpgList;
     FA_AlphaModFixXpgList: TCT_AlphaModulateFixedEffectXpgList;
     FA_AlphaOutsetXpgList: TCT_AlphaOutsetEffectXpgList;
     FA_AlphaReplXpgList: TCT_AlphaReplaceEffectXpgList;
     FA_BiLevelXpgList: TCT_BiLevelEffectXpgList;
     FA_BlendXpgList: TCT_BlendEffectXpgList;
     FA_BlurXpgList: TCT_BlurEffectXpgList;
     FA_ClrChangeXpgList: TCT_ColorChangeEffectXpgList;
     FA_ClrReplXpgList: TCT_ColorReplaceEffectXpgList;
     FA_DuotoneXpgList: TCT_DuotoneEffectXpgList;
     FA_FillXpgList: TCT_FillEffectXpgList;
     FA_FillOverlayXpgList: TCT_FillOverlayEffectXpgList;
     FA_GlowXpgList: TCT_GlowEffectXpgList;
     FA_GraysclXpgList: TCT_GrayscaleEffectXpgList;
     FA_HslXpgList: TCT_HSLEffectXpgList;
     FA_InnerShdwXpgList: TCT_InnerShadowEffectXpgList;
     FA_LumXpgList: TCT_LuminanceEffectXpgList;
     FA_OuterShdwXpgList: TCT_OuterShadowEffectXpgList;
     FA_PrstShdwXpgList: TCT_PresetShadowEffectXpgList;
     FA_ReflectionXpgList: TCT_ReflectionEffectXpgList;
     FA_RelOffXpgList: TCT_RelativeOffsetEffectXpgList;
     FA_SoftEdgeXpgList: TCT_SoftEdgesEffectXpgList;
     FA_TintXpgList: TCT_TintEffectXpgList;
     FA_XfrmXpgList: TCT_TransformEffectXpgList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TEG_Effect);
     procedure CopyTo(AItem: TEG_Effect);
     function  Create_A_ContXpgList: TCT_EffectContainerXpgList;
     function  Create_A_EffectXpgList: TCT_EffectReferenceXpgList;
     function  Create_A_AlphaBiLevelXpgList: TCT_AlphaBiLevelEffectXpgList;
     function  Create_A_AlphaCeilingXpgList: TCT_AlphaCeilingEffectXpgList;
     function  Create_A_AlphaFloorXpgList: TCT_AlphaFloorEffectXpgList;
     function  Create_A_AlphaInvXpgList: TCT_AlphaInverseEffectXpgList;
     function  Create_A_AlphaModXpgList: TCT_AlphaModulateEffectXpgList;
     function  Create_A_AlphaModFixXpgList: TCT_AlphaModulateFixedEffectXpgList;
     function  Create_A_AlphaOutsetXpgList: TCT_AlphaOutsetEffectXpgList;
     function  Create_A_AlphaReplXpgList: TCT_AlphaReplaceEffectXpgList;
     function  Create_A_BiLevelXpgList: TCT_BiLevelEffectXpgList;
     function  Create_A_BlendXpgList: TCT_BlendEffectXpgList;
     function  Create_A_BlurXpgList: TCT_BlurEffectXpgList;
     function  Create_A_ClrChangeXpgList: TCT_ColorChangeEffectXpgList;
     function  Create_A_ClrReplXpgList: TCT_ColorReplaceEffectXpgList;
     function  Create_A_DuotoneXpgList: TCT_DuotoneEffectXpgList;
     function  Create_A_FillXpgList: TCT_FillEffectXpgList;
     function  Create_A_FillOverlayXpgList: TCT_FillOverlayEffectXpgList;
     function  Create_A_GlowXpgList: TCT_GlowEffectXpgList;
     function  Create_A_GraysclXpgList: TCT_GrayscaleEffectXpgList;
     function  Create_A_HslXpgList: TCT_HSLEffectXpgList;
     function  Create_A_InnerShdwXpgList: TCT_InnerShadowEffectXpgList;
     function  Create_A_LumXpgList: TCT_LuminanceEffectXpgList;
     function  Create_A_OuterShdwXpgList: TCT_OuterShadowEffectXpgList;
     function  Create_A_PrstShdwXpgList: TCT_PresetShadowEffectXpgList;
     function  Create_A_ReflectionXpgList: TCT_ReflectionEffectXpgList;
     function  Create_A_RelOffXpgList: TCT_RelativeOffsetEffectXpgList;
     function  Create_A_SoftEdgeXpgList: TCT_SoftEdgesEffectXpgList;
     function  Create_A_TintXpgList: TCT_TintEffectXpgList;
     function  Create_A_XfrmXpgList: TCT_TransformEffectXpgList;

     property A_ContXpgList: TCT_EffectContainerXpgList read FA_ContXpgList write FA_ContXpgList;
     property A_EffectXpgList: TCT_EffectReferenceXpgList read FA_EffectXpgList;
     property A_AlphaBiLevelXpgList: TCT_AlphaBiLevelEffectXpgList read FA_AlphaBiLevelXpgList;
     property A_AlphaCeilingXpgList: TCT_AlphaCeilingEffectXpgList read FA_AlphaCeilingXpgList;
     property A_AlphaFloorXpgList: TCT_AlphaFloorEffectXpgList read FA_AlphaFloorXpgList;
     property A_AlphaInvXpgList: TCT_AlphaInverseEffectXpgList read FA_AlphaInvXpgList;
     property A_AlphaModXpgList: TCT_AlphaModulateEffectXpgList read FA_AlphaModXpgList write FA_AlphaModXpgList;
     property A_AlphaModFixXpgList: TCT_AlphaModulateFixedEffectXpgList read FA_AlphaModFixXpgList;
     property A_AlphaOutsetXpgList: TCT_AlphaOutsetEffectXpgList read FA_AlphaOutsetXpgList;
     property A_AlphaReplXpgList: TCT_AlphaReplaceEffectXpgList read FA_AlphaReplXpgList;
     property A_BiLevelXpgList: TCT_BiLevelEffectXpgList read FA_BiLevelXpgList;
     property A_BlendXpgList: TCT_BlendEffectXpgList read FA_BlendXpgList write FA_BlendXpgList;
     property A_BlurXpgList: TCT_BlurEffectXpgList read FA_BlurXpgList;
     property A_ClrChangeXpgList: TCT_ColorChangeEffectXpgList read FA_ClrChangeXpgList;
     property A_ClrReplXpgList: TCT_ColorReplaceEffectXpgList read FA_ClrReplXpgList;
     property A_DuotoneXpgList: TCT_DuotoneEffectXpgList read FA_DuotoneXpgList;
     property A_FillXpgList: TCT_FillEffectXpgList read FA_FillXpgList;
     property A_FillOverlayXpgList: TCT_FillOverlayEffectXpgList read FA_FillOverlayXpgList;
     property A_GlowXpgList: TCT_GlowEffectXpgList read FA_GlowXpgList;
     property A_GraysclXpgList: TCT_GrayscaleEffectXpgList read FA_GraysclXpgList;
     property A_HslXpgList: TCT_HSLEffectXpgList read FA_HslXpgList;
     property A_InnerShdwXpgList: TCT_InnerShadowEffectXpgList read FA_InnerShdwXpgList;
     property A_LumXpgList: TCT_LuminanceEffectXpgList read FA_LumXpgList;
     property A_OuterShdwXpgList: TCT_OuterShadowEffectXpgList read FA_OuterShdwXpgList;
     property A_PrstShdwXpgList: TCT_PresetShadowEffectXpgList read FA_PrstShdwXpgList;
     property A_ReflectionXpgList: TCT_ReflectionEffectXpgList read FA_ReflectionXpgList;
     property A_RelOffXpgList: TCT_RelativeOffsetEffectXpgList read FA_RelOffXpgList;
     property A_SoftEdgeXpgList: TCT_SoftEdgesEffectXpgList read FA_SoftEdgeXpgList;
     property A_TintXpgList: TCT_TintEffectXpgList read FA_TintXpgList;
     property A_XfrmXpgList: TCT_TransformEffectXpgList read FA_XfrmXpgList;
     end;

     TCT_EffectContainer = class(TXPGBase)
protected
     FType: TST_EffectContainerType;
     FName: AxUCString;
     FA_EG_Effect: TEG_Effect;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_EffectContainer);
     procedure CopyTo(AItem: TCT_EffectContainer);

     property Type_: TST_EffectContainerType read FType write FType;
     property Name: AxUCString read FName write FName;
     property A_EG_Effect: TEG_Effect read FA_EG_Effect;
     end;

     TCT_EffectContainerXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_EffectContainer;
public
     function  Add: TCT_EffectContainer;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_EffectContainerXpgList);
     procedure CopyTo(AItem: TCT_EffectContainerXpgList);
     property Items[Index: integer]: TCT_EffectContainer read GetItems; default;
     end;

     TCT_BlendEffect = class(TXPGBase)
protected
     FBlend: TST_BlendMode;
     FA_Cont: TCT_EffectContainer;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_BlendEffect);
     procedure CopyTo(AItem: TCT_BlendEffect);
     function  Create_A_Cont: TCT_EffectContainer;

     property Blend: TST_BlendMode read FBlend write FBlend;
     property A_Cont: TCT_EffectContainer read FA_Cont;
     end;

     TCT_BlendEffectXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_BlendEffect;
public
     function  Add: TCT_BlendEffect;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_BlendEffectXpgList);
     procedure CopyTo(AItem: TCT_BlendEffectXpgList);
     property Items[Index: integer]: TCT_BlendEffect read GetItems; default;
     end;

     TCT_AlphaModulateEffect = class(TXPGBase)
protected
     FA_Cont: TCT_EffectContainer;
public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_AlphaModulateEffect);
     procedure CopyTo(AItem: TCT_AlphaModulateEffect);
     function  Create_A_Cont: TCT_EffectContainer;

     property Container: TCT_EffectContainer read FA_Cont;
     end;

     TCT_AlphaModulateEffectXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_AlphaModulateEffect;
public
     function  Add: TCT_AlphaModulateEffect;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_AlphaModulateEffectXpgList);
     procedure CopyTo(AItem: TCT_AlphaModulateEffectXpgList);
     property Items[Index: integer]: TCT_AlphaModulateEffect read GetItems; default;
     end;

     TCT_DashStop = class(TXPGBase)
protected
     FD: integer;
     FSp: integer;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_DashStop);
     procedure CopyTo(AItem: TCT_DashStop);

     property D: integer read FD write FD;
     property Sp: integer read FSp write FSp;
     end;

     TCT_DashStopXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_DashStop;
public
     function  Add: TCT_DashStop;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_DashStopXpgList);
     procedure CopyTo(AItem: TCT_DashStopXpgList);
     property Items[Index: integer]: TCT_DashStop read GetItems; default;
     end;

     TCT_PresetLineDashProperties = class(TXPGBase)
protected
     FVal: TST_PresetLineDashVal;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_PresetLineDashProperties);
     procedure CopyTo(AItem: TCT_PresetLineDashProperties);

     property Val: TST_PresetLineDashVal read FVal write FVal;
     end;

     TCT_DashStopList = class(TXPGBase)
protected
     FA_DsXpgList: TCT_DashStopXpgList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_DashStopList);
     procedure CopyTo(AItem: TCT_DashStopList);
     function  Create_A_DsXpgList: TCT_DashStopXpgList;

     property A_DsXpgList: TCT_DashStopXpgList read FA_DsXpgList;
     end;

     TCT_LineJoinRound = class(TXPGBase)
public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_LineJoinRound);
     procedure CopyTo(AItem: TCT_LineJoinRound);

     end;

     TCT_LineJoinBevel = class(TXPGBase)
public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_LineJoinBevel);
     procedure CopyTo(AItem: TCT_LineJoinBevel);

     end;

     TCT_LineJoinMiterProperties = class(TXPGBase)
protected
     FLim: integer;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_LineJoinMiterProperties);
     procedure CopyTo(AItem: TCT_LineJoinMiterProperties);

     property Lim: integer read FLim write FLim;
     end;

     TEG_LineDashProperties = class(TXPGBase)
protected
     FA_PrstDash: TCT_PresetLineDashProperties;
     FA_CustDash: TCT_DashStopList;
public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TEG_LineDashProperties);
     procedure CopyTo(AItem: TEG_LineDashProperties);
     function  Create_PrstDash: TCT_PresetLineDashProperties;
     function  Create_CustDash: TCT_DashStopList;

     property PrstDash: TCT_PresetLineDashProperties read FA_PrstDash;
     property CustDash: TCT_DashStopList read FA_CustDash;
     end;

     TEG_LineJoinProperties = class(TXPGBase)
protected
     FA_Round: TCT_LineJoinRound;
     FA_Bevel: TCT_LineJoinBevel;
     FA_Miter: TCT_LineJoinMiterProperties;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TEG_LineJoinProperties);
     procedure CopyTo(AItem: TEG_LineJoinProperties);
     function  Create_Round: TCT_LineJoinRound;
     function  Create_Bevel: TCT_LineJoinBevel;
     function  Create_Miter: TCT_LineJoinMiterProperties;

     property Round: TCT_LineJoinRound read FA_Round;
     property Bevel: TCT_LineJoinBevel read FA_Bevel;
     property Miter: TCT_LineJoinMiterProperties read FA_Miter;
     end;

     TCT_LineEndProperties = class(TXPGBase)
protected
     FType: TST_LineEndType;
     FW: TST_LineEndWidth;
     FLen: TST_LineEndLength;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_LineEndProperties);
     procedure CopyTo(AItem: TCT_LineEndProperties);

     property Type_: TST_LineEndType read FType write FType;
     property W: TST_LineEndWidth read FW write FW;
     property Len: TST_LineEndLength read FLen write FLen;
     end;

     TCT_EffectList = class(TXPGBase)
protected
     FA_Blur: TCT_BlurEffect;
     FA_FillOverlay: TCT_FillOverlayEffect;
     FA_Glow: TCT_GlowEffect;
     FA_InnerShdw: TCT_InnerShadowEffect;
     FA_OuterShdw: TCT_OuterShadowEffect;
     FA_PrstShdw: TCT_PresetShadowEffect;
     FA_Reflection: TCT_ReflectionEffect;
     FA_SoftEdge: TCT_SoftEdgesEffect;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_EffectList);
     procedure CopyTo(AItem: TCT_EffectList);
     function  Create_A_Blur: TCT_BlurEffect;
     function  Create_A_FillOverlay: TCT_FillOverlayEffect;
     function  Create_A_Glow: TCT_GlowEffect;
     function  Create_A_InnerShdw: TCT_InnerShadowEffect;
     function  Create_A_OuterShdw: TCT_OuterShadowEffect;
     function  Create_A_PrstShdw: TCT_PresetShadowEffect;
     function  Create_A_Reflection: TCT_ReflectionEffect;
     function  Create_A_SoftEdge: TCT_SoftEdgesEffect;

     property A_Blur: TCT_BlurEffect read FA_Blur;
     property A_FillOverlay: TCT_FillOverlayEffect read FA_FillOverlay;
     property A_Glow: TCT_GlowEffect read FA_Glow;
     property A_InnerShdw: TCT_InnerShadowEffect read FA_InnerShdw;
     property A_OuterShdw: TCT_OuterShadowEffect read FA_OuterShdw;
     property A_PrstShdw: TCT_PresetShadowEffect read FA_PrstShdw;
     property A_Reflection: TCT_ReflectionEffect read FA_Reflection;
     property A_SoftEdge: TCT_SoftEdgesEffect read FA_SoftEdge;
     end;

     TCT_TextUnderlineLineFollowText = class(TXPGBase)
public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_TextUnderlineLineFollowText);
     procedure CopyTo(AItem: TCT_TextUnderlineLineFollowText);

     end;

     TCT_LineProperties = class(TXPGBase)
protected
     FW: integer;
     FCap: TST_LineCap;
     FCmpd: TST_CompoundLine;
     FAlgn: TST_PenAlignment;
     FA_EG_LineFillProperties: TEG_FillProperties;
     FA_EG_LineDashProperties: TEG_LineDashProperties;
     FA_EG_LineJoinProperties: TEG_LineJoinProperties;
     FA_HeadEnd: TCT_LineEndProperties;
     FA_TailEnd: TCT_LineEndProperties;
     FA_ExtLst: TCT_OfficeArtExtensionList;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     procedure Assign(AItem: TCT_LineProperties);
     procedure CopyTo(AItem: TCT_LineProperties);
     function  Create_HeadEnd: TCT_LineEndProperties;
     function  Create_TailEnd: TCT_LineEndProperties;
     function  Create_ExtLst: TCT_OfficeArtExtensionList;
     procedure Free_HeadEnd;
     procedure Free_TailEnd;

     property W: integer read FW write FW;
     property Cap: TST_LineCap read FCap write FCap;
     property Cmpd: TST_CompoundLine read FCmpd write FCmpd;
     property Algn: TST_PenAlignment read FAlgn write FAlgn;
     property LineFillProperties: TEG_FillProperties read FA_EG_LineFillProperties;
     property LineDashProperties: TEG_LineDashProperties read FA_EG_LineDashProperties;
     property LineJoinProperties: TEG_LineJoinProperties read FA_EG_LineJoinProperties;
     property HeadEnd: TCT_LineEndProperties read FA_HeadEnd;
     property TailEnd: TCT_LineEndProperties read FA_TailEnd;
     property ExtLst: TCT_OfficeArtExtensionList read FA_ExtLst;
     end;

     TCT_LinePropertiesXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_LineProperties;
public
     function  Add: TCT_LineProperties;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_LinePropertiesXpgList);
     procedure CopyTo(AItem: TCT_LinePropertiesXpgList);
     property Items[Index: integer]: TCT_LineProperties read GetItems; default;
     end;

     TCT_TextUnderlineFillFollowText = class(TXPGBase)
public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_TextUnderlineFillFollowText);
     procedure CopyTo(AItem: TCT_TextUnderlineFillFollowText);

     end;

     TCT_TextUnderlineFillGroupWrapper = class(TXPGBase)
protected
     FA_EG_FillProperties: TEG_FillProperties;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_TextUnderlineFillGroupWrapper);
     procedure CopyTo(AItem: TCT_TextUnderlineFillGroupWrapper);

     property A_EG_FillProperties: TEG_FillProperties read FA_EG_FillProperties;
     end;

     TCT_EmbeddedWAVAudioFile = class(TXPGBase)
protected
     FR_Embed: AxUCString;
     FName: AxUCString;
     FBuiltIn: boolean;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_EmbeddedWAVAudioFile);
     procedure CopyTo(AItem: TCT_EmbeddedWAVAudioFile);

     property R_Embed: AxUCString read FR_Embed write FR_Embed;
     property Name: AxUCString read FName write FName;
     property BuiltIn: boolean read FBuiltIn write FBuiltIn;
     end;

     TCT_AdjPoint2D = class(TXPGBase)
protected
     FX: TST_AdjCoordinate;
     PFX: PST_AdjCoordinate;
     FY: TST_AdjCoordinate;
     PFY: PST_AdjCoordinate;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_AdjPoint2D);
     procedure CopyTo(AItem: TCT_AdjPoint2D);

     property PX: PST_AdjCoordinate read PFX write PFX;
     property PY: PST_AdjCoordinate read PFY write PFY;
     end;

     TCT_AdjPoint2DXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_AdjPoint2D;
public
     function  Add: TCT_AdjPoint2D;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_AdjPoint2DXpgList);
     procedure CopyTo(AItem: TCT_AdjPoint2DXpgList);
     property Items[Index: integer]: TCT_AdjPoint2D read GetItems; default;
     end;

     TCT_TextSpacingPercent = class(TXPGBase)
protected
     FVal: integer;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_TextSpacingPercent);
     procedure CopyTo(AItem: TCT_TextSpacingPercent);

     property Val: integer read FVal write FVal;
     end;

     TCT_TextSpacingPoint = class(TXPGBase)
protected
     FVal: integer;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_TextSpacingPoint);
     procedure CopyTo(AItem: TCT_TextSpacingPoint);

     property Val: integer read FVal write FVal;
     end;

     TCT_TextBulletColorFollowText = class(TXPGBase)
public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_TextBulletColorFollowText);
     procedure CopyTo(AItem: TCT_TextBulletColorFollowText);

     end;

     TCT_TextBulletSizeFollowText = class(TXPGBase)
public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_TextBulletSizeFollowText);
     procedure CopyTo(AItem: TCT_TextBulletSizeFollowText);

     end;

     TCT_TextBulletSizePercent = class(TXPGBase)
protected
     FVal: integer;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_TextBulletSizePercent);
     procedure CopyTo(AItem: TCT_TextBulletSizePercent);

     property Val: integer read FVal write FVal;
     end;

     TCT_TextBulletSizePoint = class(TXPGBase)
protected
     FVal: integer;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_TextBulletSizePoint);
     procedure CopyTo(AItem: TCT_TextBulletSizePoint);

     property Val: integer read FVal write FVal;
     end;

     TCT_TextBulletTypefaceFollowText = class(TXPGBase)
public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_TextBulletTypefaceFollowText);
     procedure CopyTo(AItem: TCT_TextBulletTypefaceFollowText);

     end;

     TCT_TextFont = class(TXPGBase)
protected
     FTypeface: AxUCString;
     FPanose: integer;
     FPitchFamily: integer;
     FCharset: integer;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_TextFont);
     procedure CopyTo(AItem: TCT_TextFont);

     property Typeface: AxUCString read FTypeface write FTypeface;
     property Panose: integer read FPanose write FPanose;
     property PitchFamily: integer read FPitchFamily write FPitchFamily;
     property Charset: integer read FCharset write FCharset;
     end;

     TCT_TextNoBullet = class(TXPGBase)
public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_TextNoBullet);
     procedure CopyTo(AItem: TCT_TextNoBullet);

     end;

     TCT_TextAutonumberBullet = class(TXPGBase)
protected
     FType: TST_TextAutonumberScheme;
     FStartAt: integer;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_TextAutonumberBullet);
     procedure CopyTo(AItem: TCT_TextAutonumberBullet);

     property Type_: TST_TextAutonumberScheme read FType write FType;
     property StartAt: integer read FStartAt write FStartAt;
     end;

     TCT_TextCharBullet = class(TXPGBase)
protected
     FChar: AxUCString;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_TextCharBullet);
     procedure CopyTo(AItem: TCT_TextCharBullet);

     property Char: AxUCString read FChar write FChar;
     end;

     TCT_TextBlipBullet = class(TXPGBase)
protected
     FA_Blip: TCT_Blip;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_TextBlipBullet);
     procedure CopyTo(AItem: TCT_TextBlipBullet);
     function  Create_A_Blip: TCT_Blip;

     property A_Blip: TCT_Blip read FA_Blip;
     end;

     TCT_TextTabStop = class(TXPGBase)
protected
     FPos: integer;
     FAlgn: TST_TextTabAlignType;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_TextTabStop);
     procedure CopyTo(AItem: TCT_TextTabStop);

     property Pos: integer read FPos write FPos;
     property Algn: TST_TextTabAlignType read FAlgn write FAlgn;
     end;

     TCT_TextTabStopXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_TextTabStop;
public
     function  Add: TCT_TextTabStop;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_TextTabStopXpgList);
     procedure CopyTo(AItem: TCT_TextTabStopXpgList);
     property Items[Index: integer]: TCT_TextTabStop read GetItems; default;
     end;

     TEG_EffectProperties = class(TXPGBase)
protected
     FA_EffectLst: TCT_EffectList;
     FA_EffectDag: TCT_EffectContainer;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TEG_EffectProperties);
     procedure CopyTo(AItem: TEG_EffectProperties);
     function  Create_A_EffectLst: TCT_EffectList;
     function  Create_A_EffectDag: TCT_EffectContainer;

     property A_EffectLst: TCT_EffectList read FA_EffectLst;
     property A_EffectDag: TCT_EffectContainer read FA_EffectDag;
     end;

     TEG_TextUnderlineLine = class(TXPGBase)
protected
     FA_ULnTx: TCT_TextUnderlineLineFollowText;
     FA_ULn: TCT_LineProperties;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TEG_TextUnderlineLine);
     procedure CopyTo(AItem: TEG_TextUnderlineLine);
     function  Create_A_ULnTx: TCT_TextUnderlineLineFollowText;
     function  Create_A_ULn: TCT_LineProperties;

     property A_ULnTx: TCT_TextUnderlineLineFollowText read FA_ULnTx;
     property A_ULn: TCT_LineProperties read FA_ULn;
     end;

     TEG_TextUnderlineFill = class(TXPGBase)
protected
     FA_UFillTx: TCT_TextUnderlineFillFollowText;
     FA_UFill: TCT_TextUnderlineFillGroupWrapper;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TEG_TextUnderlineFill);
     procedure CopyTo(AItem: TEG_TextUnderlineFill);
     function  Create_A_UFillTx: TCT_TextUnderlineFillFollowText;
     function  Create_A_UFill: TCT_TextUnderlineFillGroupWrapper;

     property A_UFillTx: TCT_TextUnderlineFillFollowText read FA_UFillTx;
     property A_UFill: TCT_TextUnderlineFillGroupWrapper read FA_UFill;
     end;

     TCT_Hyperlink = class(TXPGBase)
protected
     FR_Id: AxUCString;
     FInvalidUrl: AxUCString;
     FAction: AxUCString;
     FTgtFrame: AxUCString;
     FTooltip: AxUCString;
     FHistory: boolean;
     FHighlightClick: boolean;
     FEndSnd: boolean;
     FA_Snd: TCT_EmbeddedWAVAudioFile;
     FA_ExtLst: TCT_OfficeArtExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Hyperlink);
     procedure CopyTo(AItem: TCT_Hyperlink);
     function  Create_A_Snd: TCT_EmbeddedWAVAudioFile;
     function  Create_A_ExtLst: TCT_OfficeArtExtensionList;

     property R_Id: AxUCString read FR_Id write FR_Id;
     property InvalidUrl: AxUCString read FInvalidUrl write FInvalidUrl;
     property Action: AxUCString read FAction write FAction;
     property TgtFrame: AxUCString read FTgtFrame write FTgtFrame;
     property Tooltip: AxUCString read FTooltip write FTooltip;
     property History: boolean read FHistory write FHistory;
     property HighlightClick: boolean read FHighlightClick write FHighlightClick;
     property EndSnd: boolean read FEndSnd write FEndSnd;
     property A_Snd: TCT_EmbeddedWAVAudioFile read FA_Snd;
     property A_ExtLst: TCT_OfficeArtExtensionList read FA_ExtLst;
     end;

     TCT_Path2DClose = class(TXPGBase)
public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Path2DClose);
     procedure CopyTo(AItem: TCT_Path2DClose);

     end;

     TCT_Path2DCloseXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_Path2DClose;
public
     function  Add: TCT_Path2DClose;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_Path2DCloseXpgList);
     procedure CopyTo(AItem: TCT_Path2DCloseXpgList);
     property Items[Index: integer]: TCT_Path2DClose read GetItems; default;
     end;

     TCT_Path2DMoveTo = class(TXPGBase)
protected
     FA_Pt: TCT_AdjPoint2D;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Path2DMoveTo);
     procedure CopyTo(AItem: TCT_Path2DMoveTo);
     function  Create_A_Pt: TCT_AdjPoint2D;

     property A_Pt: TCT_AdjPoint2D read FA_Pt;
     end;

     TCT_Path2DMoveToXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_Path2DMoveTo;
public
     function  Add: TCT_Path2DMoveTo;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_Path2DMoveToXpgList);
     procedure CopyTo(AItem: TCT_Path2DMoveToXpgList);
     property Items[Index: integer]: TCT_Path2DMoveTo read GetItems; default;
     end;

     TCT_Path2DLineTo = class(TXPGBase)
protected
     FA_Pt: TCT_AdjPoint2D;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Path2DLineTo);
     procedure CopyTo(AItem: TCT_Path2DLineTo);
     function  Create_A_Pt: TCT_AdjPoint2D;

     property A_Pt: TCT_AdjPoint2D read FA_Pt;
     end;

     TCT_Path2DLineToXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_Path2DLineTo;
public
     function  Add: TCT_Path2DLineTo;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_Path2DLineToXpgList);
     procedure CopyTo(AItem: TCT_Path2DLineToXpgList);
     property Items[Index: integer]: TCT_Path2DLineTo read GetItems; default;
     end;

     TCT_Path2DArcTo = class(TXPGBase)
protected
     FWR: TST_AdjCoordinate;
     PFWR: PST_AdjCoordinate;
     FHR: TST_AdjCoordinate;
     PFHR: PST_AdjCoordinate;
     FStAng: TST_AdjAngle;
     PFStAng: PST_AdjAngle;
     FSwAng: TST_AdjAngle;
     PFSwAng: PST_AdjAngle;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Path2DArcTo);
     procedure CopyTo(AItem: TCT_Path2DArcTo);

     property PWR: PST_AdjCoordinate read PFWR write PFWR;
     property PHR: PST_AdjCoordinate read PFHR write PFHR;
     property PStAng: PST_AdjAngle read PFStAng write PFStAng;
     property PSwAng: PST_AdjAngle read PFSwAng write PFSwAng;
     end;

     TCT_Path2DArcToXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_Path2DArcTo;
public
     function  Add: TCT_Path2DArcTo;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_Path2DArcToXpgList);
     procedure CopyTo(AItem: TCT_Path2DArcToXpgList);
     property Items[Index: integer]: TCT_Path2DArcTo read GetItems; default;
     end;

     TCT_Path2DQuadBezierTo = class(TXPGBase)
protected
     FA_PtXpgList: TCT_AdjPoint2DXpgList;
public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Path2DQuadBezierTo);
     procedure CopyTo(AItem: TCT_Path2DQuadBezierTo);

     property A_PtXpgList: TCT_AdjPoint2DXpgList read FA_PtXpgList;
     end;

     TCT_Path2DQuadBezierToXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_Path2DQuadBezierTo;
public
     function  Add: TCT_Path2DQuadBezierTo;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_Path2DQuadBezierToXpgList);
     procedure CopyTo(AItem: TCT_Path2DQuadBezierToXpgList);
     property Items[Index: integer]: TCT_Path2DQuadBezierTo read GetItems; default;
     end;

     TCT_Path2DCubicBezierTo = class(TXPGBase)
protected
     FA_PtXpgList: TCT_AdjPoint2DXpgList;
public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Path2DCubicBezierTo);
     procedure CopyTo(AItem: TCT_Path2DCubicBezierTo);

     property A_PtXpgList: TCT_AdjPoint2DXpgList read FA_PtXpgList;
     end;

     TCT_Path2DCubicBezierToXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_Path2DCubicBezierTo;
public
     function  Add: TCT_Path2DCubicBezierTo;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_Path2DCubicBezierToXpgList);
     procedure CopyTo(AItem: TCT_Path2DCubicBezierToXpgList);
     property Items[Index: integer]: TCT_Path2DCubicBezierTo read GetItems; default;
     end;

     TCT_TextSpacing = class(TXPGBase)
protected
     FA_SpcPct: TCT_TextSpacingPercent;
     FA_SpcPts: TCT_TextSpacingPoint;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_TextSpacing);
     procedure CopyTo(AItem: TCT_TextSpacing);
     function  Create_A_SpcPct: TCT_TextSpacingPercent;
     function  Create_A_SpcPts: TCT_TextSpacingPoint;

     property A_SpcPct: TCT_TextSpacingPercent read FA_SpcPct;
     property A_SpcPts: TCT_TextSpacingPoint read FA_SpcPts;
     end;

     TEG_TextBulletColor = class(TXPGBase)
protected
     FA_BuClrTx: TCT_TextBulletColorFollowText;
     FA_BuClr: TCT_Color;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TEG_TextBulletColor);
     procedure CopyTo(AItem: TEG_TextBulletColor);
     function  Create_A_BuClrTx: TCT_TextBulletColorFollowText;
     function  Create_A_BuClr: TCT_Color;

     property A_BuClrTx: TCT_TextBulletColorFollowText read FA_BuClrTx;
     property A_BuClr: TCT_Color read FA_BuClr;
     end;

     TEG_TextBulletSize = class(TXPGBase)
protected
     FA_BuSzTx: TCT_TextBulletSizeFollowText;
     FA_BuSzPct: TCT_TextBulletSizePercent;
     FA_BuSzPts: TCT_TextBulletSizePoint;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TEG_TextBulletSize);
     procedure CopyTo(AItem: TEG_TextBulletSize);
     function  Create_A_BuSzTx: TCT_TextBulletSizeFollowText;
     function  Create_A_BuSzPct: TCT_TextBulletSizePercent;
     function  Create_A_BuSzPts: TCT_TextBulletSizePoint;

     property A_BuSzTx: TCT_TextBulletSizeFollowText read FA_BuSzTx;
     property A_BuSzPct: TCT_TextBulletSizePercent read FA_BuSzPct;
     property A_BuSzPts: TCT_TextBulletSizePoint read FA_BuSzPts;
     end;

     TEG_TextBulletTypeface = class(TXPGBase)
protected
     FA_BuFontTx: TCT_TextBulletTypefaceFollowText;
     FA_BuFont: TCT_TextFont;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TEG_TextBulletTypeface);
     procedure CopyTo(AItem: TEG_TextBulletTypeface);
     function  Create_A_BuFontTx: TCT_TextBulletTypefaceFollowText;
     function  Create_A_BuFont: TCT_TextFont;

     property A_BuFontTx: TCT_TextBulletTypefaceFollowText read FA_BuFontTx;
     property A_BuFont: TCT_TextFont read FA_BuFont;
     end;

     TEG_TextBullet = class(TXPGBase)
protected
     FA_BuNone: TCT_TextNoBullet;
     FA_BuAutoNum: TCT_TextAutonumberBullet;
     FA_BuChar: TCT_TextCharBullet;
     FA_BuBlip: TCT_TextBlipBullet;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TEG_TextBullet);
     procedure CopyTo(AItem: TEG_TextBullet);
     function  Create_A_BuNone: TCT_TextNoBullet;
     function  Create_A_BuAutoNum: TCT_TextAutonumberBullet;
     function  Create_A_BuChar: TCT_TextCharBullet;
     function  Create_A_BuBlip: TCT_TextBlipBullet;

     property A_BuNone: TCT_TextNoBullet read FA_BuNone;
     property A_BuAutoNum: TCT_TextAutonumberBullet read FA_BuAutoNum;
     property A_BuChar: TCT_TextCharBullet read FA_BuChar;
     property A_BuBlip: TCT_TextBlipBullet read FA_BuBlip;
     end;

     TCT_TextTabStopList = class(TXPGBase)
protected
     FA_TabXpgList: TCT_TextTabStopXpgList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_TextTabStopList);
     procedure CopyTo(AItem: TCT_TextTabStopList);
     function  Create_A_TabXpgList: TCT_TextTabStopXpgList;

     property A_TabXpgList: TCT_TextTabStopXpgList read FA_TabXpgList;
     end;

     TCT_TextCharacterProperties = class(TXPGBase)
protected
     FKumimoji: boolean;
     FLang: AxUCString;
     FAltLang: AxUCString;
     FSz: integer;
     FB: boolean;
     FI: boolean;
     FU: TST_TextUnderlineType;
     FStrike: TST_TextStrikeType;
     FKern: integer;
     FCap: TST_TextCapsType;
     FSpc: integer;
     FNormalizeH: boolean;
     FBaseline: integer;
     FNoProof: boolean;
     FDirty: boolean;
     FErr: boolean;
     FSmtClean: boolean;
     FSmtId: integer;
     FBmk: AxUCString;
     FA_Ln: TCT_LineProperties;
     FA_EG_FillProperties: TEG_FillProperties;
     FA_EG_EffectProperties: TEG_EffectProperties;
     FA_Highlight: TCT_Color;
     FA_EG_TextUnderlineLine: TEG_TextUnderlineLine;
     FA_EG_TextUnderlineFill: TEG_TextUnderlineFill;
     FA_Latin: TCT_TextFont;
     FA_Ea: TCT_TextFont;
     FA_Cs: TCT_TextFont;
     FA_Sym: TCT_TextFont;
     FA_HlinkClick: TCT_Hyperlink;
     FA_HlinkMouseOver: TCT_Hyperlink;
     FA_ExtLst: TCT_OfficeArtExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_TextCharacterProperties);
     // Used when adding new runs with the previous run font propertied;
     procedure AssignFont(AItem: TCT_TextCharacterProperties);

     procedure CopyTo(AItem: TCT_TextCharacterProperties);
     function  Create_Ln: TCT_LineProperties;
     function  Create_Highlight: TCT_Color;
     function  Create_Latin: TCT_TextFont;
     function  Create_Ea: TCT_TextFont;
     function  Create_Cs: TCT_TextFont;
     function  Create_Sym: TCT_TextFont;
     function  Create_HlinkClick: TCT_Hyperlink;
     function  Create_HlinkMouseOver: TCT_Hyperlink;
     function  Create_ExtLst: TCT_OfficeArtExtensionList;

     property Kumimoji: boolean read FKumimoji write FKumimoji;
     property Lang: AxUCString read FLang write FLang;
     property AltLang: AxUCString read FAltLang write FAltLang;
     property Sz: integer read FSz write FSz;
     property B: boolean read FB write FB;
     property I: boolean read FI write FI;
     property U: TST_TextUnderlineType read FU write FU;
     property Strike: TST_TextStrikeType read FStrike write FStrike;
     property Kern: integer read FKern write FKern;
     property Cap: TST_TextCapsType read FCap write FCap;
     property Spc: integer read FSpc write FSpc;
     property NormalizeH: boolean read FNormalizeH write FNormalizeH;
     property Baseline: integer read FBaseline write FBaseline;
     property NoProof: boolean read FNoProof write FNoProof;
     property Dirty: boolean read FDirty write FDirty;
     property Err: boolean read FErr write FErr;
     property SmtClean: boolean read FSmtClean write FSmtClean;
     property SmtId: integer read FSmtId write FSmtId;
     property Bmk: AxUCString read FBmk write FBmk;
     property Ln: TCT_LineProperties read FA_Ln;
     property FillProperties: TEG_FillProperties read FA_EG_FillProperties;
     property EffectProperties: TEG_EffectProperties read FA_EG_EffectProperties;
     property Highlight: TCT_Color read FA_Highlight;
     property TextUnderlineLine: TEG_TextUnderlineLine read FA_EG_TextUnderlineLine;
     property TextUnderlineFill: TEG_TextUnderlineFill read FA_EG_TextUnderlineFill;
     property Latin: TCT_TextFont read FA_Latin;
     property Ea: TCT_TextFont read FA_Ea;
     property Cs: TCT_TextFont read FA_Cs;
     property Sym: TCT_TextFont read FA_Sym;
     property HlinkClick: TCT_Hyperlink read FA_HlinkClick;
     property HlinkMouseOver: TCT_Hyperlink read FA_HlinkMouseOver;
     property ExtLst: TCT_OfficeArtExtensionList read FA_ExtLst;
     end;

     TCT_GeomGuide = class(TXPGBase)
protected
     FName: AxUCString;
     FFmla: AxUCString;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_GeomGuide);
     procedure CopyTo(AItem: TCT_GeomGuide);

     property Name: AxUCString read FName write FName;
     property Fmla: AxUCString read FFmla write FFmla;
     end;

     TCT_GeomGuideXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_GeomGuide;
public
     function  Add: TCT_GeomGuide;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_GeomGuideXpgList);
     procedure CopyTo(AItem: TCT_GeomGuideXpgList);
     property Items[Index: integer]: TCT_GeomGuide read GetItems; default;
     end;

     TCT_XYAdjustHandle = class(TXPGBase)
protected
     FGdRefX: AxUCString;
     FMinX: TST_AdjCoordinate;
     PFMinX: PST_AdjCoordinate;
     FMaxX: TST_AdjCoordinate;
     PFMaxX: PST_AdjCoordinate;
     FGdRefY: AxUCString;
     FMinY: TST_AdjCoordinate;
     PFMinY: PST_AdjCoordinate;
     FMaxY: TST_AdjCoordinate;
     PFMaxY: PST_AdjCoordinate;
     FA_Pos: TCT_AdjPoint2D;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_XYAdjustHandle);
     procedure CopyTo(AItem: TCT_XYAdjustHandle);
     function  Create_A_Pos: TCT_AdjPoint2D;

     property GdRefX: AxUCString read FGdRefX write FGdRefX;
     property PMinX: PST_AdjCoordinate read PFMinX write PFMinX;
     property PMaxX: PST_AdjCoordinate read PFMaxX write PFMaxX;
     property GdRefY: AxUCString read FGdRefY write FGdRefY;
     property PMinY: PST_AdjCoordinate read PFMinY write PFMinY;
     property PMaxY: PST_AdjCoordinate read PFMaxY write PFMaxY;
     property A_Pos: TCT_AdjPoint2D read FA_Pos;
     end;

     TCT_XYAdjustHandleXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_XYAdjustHandle;
public
     function  Add: TCT_XYAdjustHandle;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_XYAdjustHandleXpgList);
     procedure CopyTo(AItem: TCT_XYAdjustHandleXpgList);
     property Items[Index: integer]: TCT_XYAdjustHandle read GetItems; default;
     end;

     TCT_PolarAdjustHandle = class(TXPGBase)
protected
     FGdRefR: AxUCString;
     FMinR: TST_AdjCoordinate;
     PFMinR: PST_AdjCoordinate;
     FMaxR: TST_AdjCoordinate;
     PFMaxR: PST_AdjCoordinate;
     FGdRefAng: AxUCString;
     FMinAng: TST_AdjAngle;
     PFMinAng: PST_AdjAngle;
     FMaxAng: TST_AdjAngle;
     PFMaxAng: PST_AdjAngle;
     FA_Pos: TCT_AdjPoint2D;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_PolarAdjustHandle);
     procedure CopyTo(AItem: TCT_PolarAdjustHandle);
     function  Create_A_Pos: TCT_AdjPoint2D;

     property GdRefR: AxUCString read FGdRefR write FGdRefR;
     property PMinR: PST_AdjCoordinate read PFMinR write PFMinR;
     property PMaxR: PST_AdjCoordinate read PFMaxR write PFMaxR;
     property GdRefAng: AxUCString read FGdRefAng write FGdRefAng;
     property PMinAng: PST_AdjAngle read PFMinAng write PFMinAng;
     property PMaxAng: PST_AdjAngle read PFMaxAng write PFMaxAng;
     property A_Pos: TCT_AdjPoint2D read FA_Pos;
     end;

     TCT_PolarAdjustHandleXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_PolarAdjustHandle;
public
     function  Add: TCT_PolarAdjustHandle;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_PolarAdjustHandleXpgList);
     procedure CopyTo(AItem: TCT_PolarAdjustHandleXpgList);
     property Items[Index: integer]: TCT_PolarAdjustHandle read GetItems; default;
     end;

     TCT_ConnectionSite = class(TXPGBase)
protected
     FAng: TST_AdjAngle;
     PFAng: PST_AdjAngle;
     FA_Pos: TCT_AdjPoint2D;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_ConnectionSite);
     procedure CopyTo(AItem: TCT_ConnectionSite);
     function  Create_A_Pos: TCT_AdjPoint2D;

     property PAng: PST_AdjAngle read PFAng write PFAng;
     property A_Pos: TCT_AdjPoint2D read FA_Pos;
     end;

     TCT_ConnectionSiteXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_ConnectionSite;
public
     function  Add: TCT_ConnectionSite;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_ConnectionSiteXpgList);
     procedure CopyTo(AItem: TCT_ConnectionSiteXpgList);
     property Items[Index: integer]: TCT_ConnectionSite read GetItems; default;
     end;

     TCT_Path2D = class(TXPGBase)
protected
     FW: integer;
     FH: integer;
     FFill: TST_PathFillMode;
     FStroke: boolean;
     FExtrusionOk: boolean;
     FObjects: TXPGBaseObjectList;
public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;

     procedure Clear;
     procedure Assign(AItem: TCT_Path2D);
     procedure CopyTo(AItem: TCT_Path2D);

     property W: integer read FW write FW;
     property H: integer read FH write FH;
     property Fill: TST_PathFillMode read FFill write FFill;
     property Stroke: boolean read FStroke write FStroke;
     property ExtrusionOk: boolean read FExtrusionOk write FExtrusionOk;
     end;

     TCT_Path2DXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_Path2D;
public
     function  Add: TCT_Path2D;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_Path2DXpgList);
     procedure CopyTo(AItem: TCT_Path2DXpgList);
     property Items[Index: integer]: TCT_Path2D read GetItems; default;
     end;

     TCT_SphereCoords = class(TXPGBase)
protected
     FLat: integer;
     FLon: integer;
     FRev: integer;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_SphereCoords);
     procedure CopyTo(AItem: TCT_SphereCoords);

     property Lat: integer read FLat write FLat;
     property Lon: integer read FLon write FLon;
     property Rev: integer read FRev write FRev;
     end;

     TCT_Point3D = class(TXPGBase)
protected
     FX: integer;
     FY: integer;
     FZ: integer;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Point3D);
     procedure CopyTo(AItem: TCT_Point3D);

     property X: integer read FX write FX;
     property Y: integer read FY write FY;
     property Z: integer read FZ write FZ;
     end;

     TCT_Vector3D = class(TXPGBase)
protected
     FDx: integer;
     FDy: integer;
     FDz: integer;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Vector3D);
     procedure CopyTo(AItem: TCT_Vector3D);

     property Dx: integer read FDx write FDx;
     property Dy: integer read FDy write FDy;
     property Dz: integer read FDz write FDz;
     end;

     TCT_Bevel = class(TXPGBase)
protected
     FW: integer;
     FH: integer;
     FPrst: TST_BevelPresetType;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Bevel);
     procedure CopyTo(AItem: TCT_Bevel);

     property W: integer read FW write FW;
     property H: integer read FH write FH;
     property Prst: TST_BevelPresetType read FPrst write FPrst;
     end;

     TCT_TextParagraphProperties = class(TXPGBase)
protected
     FMarL: integer;
     FMarR: integer;
     FLvl: integer;
     FIndent: integer;
     FAlgn: TST_TextAlignType;
     FDefTabSz: integer;
     FRtl: boolean;
     FEaLnBrk: boolean;
     FFontAlgn: TST_TextFontAlignType;
     FLatinLnBrk: boolean;
     FHangingPunct: boolean;
     FA_LnSpc: TCT_TextSpacing;
     FA_SpcBef: TCT_TextSpacing;
     FA_SpcAft: TCT_TextSpacing;
     FA_EG_TextBulletColor: TEG_TextBulletColor;
     FA_EG_TextBulletSize: TEG_TextBulletSize;
     FA_EG_TextBulletTypeface: TEG_TextBulletTypeface;
     FA_EG_TextBullet: TEG_TextBullet;
     FA_TabLst: TCT_TextTabStopList;
     FA_DefRPr: TCT_TextCharacterProperties;
     FA_ExtLst: TCT_OfficeArtExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_TextParagraphProperties);
     procedure CopyTo(AItem: TCT_TextParagraphProperties);
     function  Create_LnSpc: TCT_TextSpacing;
     function  Create_SpcBef: TCT_TextSpacing;
     function  Create_SpcAft: TCT_TextSpacing;
     function  Create_TabLst: TCT_TextTabStopList;
     function  Create_DefRPr: TCT_TextCharacterProperties;
     function  Create_ExtLst: TCT_OfficeArtExtensionList;

     property MarL: integer read FMarL write FMarL;
     property MarR: integer read FMarR write FMarR;
     property Lvl: integer read FLvl write FLvl;
     property Indent: integer read FIndent write FIndent;
     property Algn: TST_TextAlignType read FAlgn write FAlgn;
     property DefTabSz: integer read FDefTabSz write FDefTabSz;
     property Rtl: boolean read FRtl write FRtl;
     property EaLnBrk: boolean read FEaLnBrk write FEaLnBrk;
     property FontAlgn: TST_TextFontAlignType read FFontAlgn write FFontAlgn;
     property LatinLnBrk: boolean read FLatinLnBrk write FLatinLnBrk;
     property HangingPunct: boolean read FHangingPunct write FHangingPunct;
     property LnSpc: TCT_TextSpacing read FA_LnSpc;
     property SpcBef: TCT_TextSpacing read FA_SpcBef;
     property SpcAft: TCT_TextSpacing read FA_SpcAft;
     property TextBulletColor: TEG_TextBulletColor read FA_EG_TextBulletColor;
     property TextBulletSize: TEG_TextBulletSize read FA_EG_TextBulletSize;
     property TextBulletTypeface: TEG_TextBulletTypeface read FA_EG_TextBulletTypeface;
     property TextBullet: TEG_TextBullet read FA_EG_TextBullet;
     property TabLst: TCT_TextTabStopList read FA_TabLst;
     property DefRPr: TCT_TextCharacterProperties read FA_DefRPr;
     property ExtLst: TCT_OfficeArtExtensionList read FA_ExtLst;
     end;

     TCT_GeomGuideList = class(TXPGBase)
protected
     FA_GdXpgList: TCT_GeomGuideXpgList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_GeomGuideList);
     procedure CopyTo(AItem: TCT_GeomGuideList);
     function  Create_A_GdXpgList: TCT_GeomGuideXpgList;

     property A_GdXpgList: TCT_GeomGuideXpgList read FA_GdXpgList;
     end;

     TCT_AdjustHandleList = class(TXPGBase)
protected
     FA_AhXYXpgList: TCT_XYAdjustHandleXpgList;
     FA_AhPolarXpgList: TCT_PolarAdjustHandleXpgList;
public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_AdjustHandleList);
     procedure CopyTo(AItem: TCT_AdjustHandleList);
     function  Create_A_AhXYXpgList: TCT_XYAdjustHandleXpgList;
     function  Create_A_AhPolarXpgList: TCT_PolarAdjustHandleXpgList;

     property A_AhXYXpgList: TCT_XYAdjustHandleXpgList read FA_AhXYXpgList;
     property A_AhPolarXpgList: TCT_PolarAdjustHandleXpgList read FA_AhPolarXpgList;
     end;

     TCT_ConnectionSiteList = class(TXPGBase)
protected
     FA_CxnXpgList: TCT_ConnectionSiteXpgList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_ConnectionSiteList);
     procedure CopyTo(AItem: TCT_ConnectionSiteList);
     function  Create_A_CxnXpgList: TCT_ConnectionSiteXpgList;

     property A_CxnXpgList: TCT_ConnectionSiteXpgList read FA_CxnXpgList;
     end;

     TCT_GeomRect = class(TXPGBase)
protected
     FL: TST_AdjCoordinate;
     PFL: PST_AdjCoordinate;
     FT: TST_AdjCoordinate;
     PFT: PST_AdjCoordinate;
     FR: TST_AdjCoordinate;
     PFR: PST_AdjCoordinate;
     FB: TST_AdjCoordinate;
     PFB: PST_AdjCoordinate;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_GeomRect);
     procedure CopyTo(AItem: TCT_GeomRect);

     property PL: PST_AdjCoordinate read PFL write PFL;
     property PT: PST_AdjCoordinate read PFT write PFT;
     property PR: PST_AdjCoordinate read PFR write PFR;
     property PB: PST_AdjCoordinate read PFB write PFB;
     end;

     TCT_Path2DList = class(TXPGBase)
protected
     FA_PathXpgList: TCT_Path2DXpgList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Path2DList);
     procedure CopyTo(AItem: TCT_Path2DList);
     function  Create_A_PathXpgList: TCT_Path2DXpgList;

     property A_PathXpgList: TCT_Path2DXpgList read FA_PathXpgList;
     end;

     TCT_TextNoAutofit = class(TXPGBase)
public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_TextNoAutofit);
     procedure CopyTo(AItem: TCT_TextNoAutofit);

     end;

     TCT_TextNormalAutofit = class(TXPGBase)
protected
     FFontScale: integer;
     FLnSpcReduction: integer;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_TextNormalAutofit);
     procedure CopyTo(AItem: TCT_TextNormalAutofit);

     property FontScale: integer read FFontScale write FFontScale;
     property LnSpcReduction: integer read FLnSpcReduction write FLnSpcReduction;
     end;

     TCT_TextShapeAutofit = class(TXPGBase)
public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_TextShapeAutofit);
     procedure CopyTo(AItem: TCT_TextShapeAutofit);

     end;

     TCT_Camera = class(TXPGBase)
protected
     FPrst: TST_PresetCameraType;
     FFov: integer;
     FZoom: integer;
     FA_Rot: TCT_SphereCoords;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Camera);
     procedure CopyTo(AItem: TCT_Camera);
     function  Create_A_Rot: TCT_SphereCoords;

     property Prst: TST_PresetCameraType read FPrst write FPrst;
     property Fov: integer read FFov write FFov;
     property Zoom: integer read FZoom write FZoom;
     property A_Rot: TCT_SphereCoords read FA_Rot;
     end;

     TCT_LightRig = class(TXPGBase)
protected
     FRig: TST_LightRigType;
     FDir: TST_LightRigDirection;
     FA_Rot: TCT_SphereCoords;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_LightRig);
     procedure CopyTo(AItem: TCT_LightRig);
     function  Create_A_Rot: TCT_SphereCoords;

     property Rig: TST_LightRigType read FRig write FRig;
     property Dir: TST_LightRigDirection read FDir write FDir;
     property A_Rot: TCT_SphereCoords read FA_Rot;
     end;

     TCT_Backdrop = class(TXPGBase)
protected
     FA_Anchor: TCT_Point3D;
     FA_Norm: TCT_Vector3D;
     FA_Up: TCT_Vector3D;
     FA_ExtLst: TCT_OfficeArtExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Backdrop);
     procedure CopyTo(AItem: TCT_Backdrop);
     function  Create_A_Anchor: TCT_Point3D;
     function  Create_A_Norm: TCT_Vector3D;
     function  Create_A_Up: TCT_Vector3D;
     function  Create_A_ExtLst: TCT_OfficeArtExtensionList;

     property A_Anchor: TCT_Point3D read FA_Anchor;
     property A_Norm: TCT_Vector3D read FA_Norm;
     property A_Up: TCT_Vector3D read FA_Up;
     property A_ExtLst: TCT_OfficeArtExtensionList read FA_ExtLst;
     end;

     TCT_Shape3D = class(TXPGBase)
protected
     FZ: integer;
     FExtrusionH: integer;
     FContourW: integer;
     FPrstMaterial: TST_PresetMaterialType;
     FA_BevelT: TCT_Bevel;
     FA_BevelB: TCT_Bevel;
     FA_ExtrusionClr: TCT_Color;
     FA_ContourClr: TCT_Color;
     FA_ExtLst: TCT_OfficeArtExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Shape3D);
     procedure CopyTo(AItem: TCT_Shape3D);
     function  Create_A_BevelT: TCT_Bevel;
     function  Create_A_BevelB: TCT_Bevel;
     function  Create_A_ExtrusionClr: TCT_Color;
     function  Create_A_ContourClr: TCT_Color;
     function  Create_A_ExtLst: TCT_OfficeArtExtensionList;

     property Z: integer read FZ write FZ;
     property ExtrusionH: integer read FExtrusionH write FExtrusionH;
     property ContourW: integer read FContourW write FContourW;
     property PrstMaterial: TST_PresetMaterialType read FPrstMaterial write FPrstMaterial;
     property A_BevelT: TCT_Bevel read FA_BevelT;
     property A_BevelB: TCT_Bevel read FA_BevelB;
     property A_ExtrusionClr: TCT_Color read FA_ExtrusionClr;
     property A_ContourClr: TCT_Color read FA_ContourClr;
     property A_ExtLst: TCT_OfficeArtExtensionList read FA_ExtLst;
     end;

     TCT_FlatText = class(TXPGBase)
protected
     FZ: integer;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_FlatText);
     procedure CopyTo(AItem: TCT_FlatText);

     property Z: integer read FZ write FZ;
     end;

     TCT_RegularTextRun = class(TXPGBase)
protected
     FA_RPr: TCT_TextCharacterProperties;
     FA_T: AxUCString;
public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_RegularTextRun);
     procedure CopyTo(AItem: TCT_RegularTextRun);
     function  Create_RPr: TCT_TextCharacterProperties;

     property RPr: TCT_TextCharacterProperties read FA_RPr;
     property T: AxUCString read FA_T write FA_T;
     end;

     TCT_TextLineBreak = class(TXPGBase)
protected
     FA_RPr: TCT_TextCharacterProperties;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_TextLineBreak);
     procedure CopyTo(AItem: TCT_TextLineBreak);
     function  Create_RPr: TCT_TextCharacterProperties;

     property RPr: TCT_TextCharacterProperties read FA_RPr;
     end;

     TCT_TextField = class(TXPGBase)
protected
     FId: AxUCString;
     FType: AxUCString;
     FA_RPr: TCT_TextCharacterProperties;
     FA_PPr: TCT_TextParagraphProperties;
     FA_T: AxUCString;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     procedure Assign(AItem: TCT_TextField);
     procedure CopyTo(AItem: TCT_TextField);
     function  Create_A_RPr: TCT_TextCharacterProperties;
     function  Create_A_PPr: TCT_TextParagraphProperties;

     property Id: AxUCString read FId write FId;
     property Type_: AxUCString read FType write FType;
     property A_RPr: TCT_TextCharacterProperties read FA_RPr;
     property A_PPr: TCT_TextParagraphProperties read FA_PPr;
     property A_T: AxUCString read FA_T write FA_T;
     end;

     TCT_Point2D = class(TXPGBase)
protected
     FX: integer;
     FY: integer;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Point2D);
     procedure CopyTo(AItem: TCT_Point2D);

     property X: integer read FX write FX;
     property Y: integer read FY write FY;
     end;

     TCT_PositiveSize2D = class(TXPGBase)
protected
     FCx: integer;
     FCy: integer;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_PositiveSize2D);
     procedure CopyTo(AItem: TCT_PositiveSize2D);

     property Cx: integer read FCx write FCx;
     property Cy: integer read FCy write FCy;
     end;

     TCT_CustomGeometry2D = class(TXPGBase)
protected
     FA_AvLst: TCT_GeomGuideList;
     FA_GdLst: TCT_GeomGuideList;
     FA_AhLst: TCT_AdjustHandleList;
     FA_CxnLst: TCT_ConnectionSiteList;
     FA_Rect: TCT_GeomRect;
     FA_PathLst: TCT_Path2DList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_CustomGeometry2D);
     procedure CopyTo(AItem: TCT_CustomGeometry2D);
     function  Create_A_AvLst: TCT_GeomGuideList;
     function  Create_A_GdLst: TCT_GeomGuideList;
     function  Create_A_AhLst: TCT_AdjustHandleList;
     function  Create_A_CxnLst: TCT_ConnectionSiteList;
     function  Create_A_Rect: TCT_GeomRect;
     function  Create_A_PathLst: TCT_Path2DList;

     property A_AvLst: TCT_GeomGuideList read FA_AvLst;
     property A_GdLst: TCT_GeomGuideList read FA_GdLst;
     property A_AhLst: TCT_AdjustHandleList read FA_AhLst;
     property A_CxnLst: TCT_ConnectionSiteList read FA_CxnLst;
     property A_Rect: TCT_GeomRect read FA_Rect;
     property A_PathLst: TCT_Path2DList read FA_PathLst;
     end;

     TCT_PresetGeometry2D = class(TXPGBase)
protected
     FPrst: TST_ShapeType;
     FA_AvLst: TCT_GeomGuideList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_PresetGeometry2D);
     procedure CopyTo(AItem: TCT_PresetGeometry2D);
     function  Create_A_AvLst: TCT_GeomGuideList;

     property Prst: TST_ShapeType read FPrst write FPrst;
     property A_AvLst: TCT_GeomGuideList read FA_AvLst;
     end;

     TCT_PresetTextShape = class(TXPGBase)
protected
     FPrst: TST_TextShapeType;
     FA_AvLst: TCT_GeomGuideList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_PresetTextShape);
     procedure CopyTo(AItem: TCT_PresetTextShape);
     function  Create_A_AvLst: TCT_GeomGuideList;

     property Prst: TST_TextShapeType read FPrst write FPrst;
     property A_AvLst: TCT_GeomGuideList read FA_AvLst;
     end;

     TEG_TextAutofit = class(TXPGBase)
protected
     FA_NoAutofit: TCT_TextNoAutofit;
     FA_NormAutofit: TCT_TextNormalAutofit;
     FA_SpAutoFit: TCT_TextShapeAutofit;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TEG_TextAutofit);
     procedure CopyTo(AItem: TEG_TextAutofit);
     function  Create_A_NoAutofit: TCT_TextNoAutofit;
     function  Create_A_NormAutofit: TCT_TextNormalAutofit;
     function  Create_A_SpAutoFit: TCT_TextShapeAutofit;

     property A_NoAutofit: TCT_TextNoAutofit read FA_NoAutofit;
     property A_NormAutofit: TCT_TextNormalAutofit read FA_NormAutofit;
     property A_SpAutoFit: TCT_TextShapeAutofit read FA_SpAutoFit;
     end;

     TCT_Scene3D = class(TXPGBase)
protected
     FA_Camera: TCT_Camera;
     FA_LightRig: TCT_LightRig;
     FA_Backdrop: TCT_Backdrop;
     FA_ExtLst: TCT_OfficeArtExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Scene3D);
     procedure CopyTo(AItem: TCT_Scene3D);
     function  Create_A_Camera: TCT_Camera;
     function  Create_A_LightRig: TCT_LightRig;
     function  Create_A_Backdrop: TCT_Backdrop;
     function  Create_A_ExtLst: TCT_OfficeArtExtensionList;

     property A_Camera: TCT_Camera read FA_Camera;
     property A_LightRig: TCT_LightRig read FA_LightRig;
     property A_Backdrop: TCT_Backdrop read FA_Backdrop;
     property A_ExtLst: TCT_OfficeArtExtensionList read FA_ExtLst;
     end;

     TEG_Text3D = class(TXPGBase)
protected
     FA_Sp3d: TCT_Shape3D;
     FA_FlatTx: TCT_FlatText;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TEG_Text3D);
     procedure CopyTo(AItem: TEG_Text3D);
     function  Create_A_Sp3d: TCT_Shape3D;
     function  Create_A_FlatTx: TCT_FlatText;

     property A_Sp3d: TCT_Shape3D read FA_Sp3d;
     property A_FlatTx: TCT_FlatText read FA_FlatTx;
     end;

     TEG_TextRun = class(TXPGBase)
protected
     FA_R: TCT_RegularTextRun;
     FA_Br: TCT_TextLineBreak;
     FA_Fld: TCT_TextField;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure Clear;
     procedure Assign(AItem: TEG_TextRun);
     procedure CopyTo(AItem: TEG_TextRun);
     function  Create_R: TCT_RegularTextRun;
     function  Create_Br: TCT_TextLineBreak;
     function  Create_Fldt: TCT_TextField;
     procedure ConvertToRun;
     procedure ConvertToBreak;

     property Run: TCT_RegularTextRun read FA_R;
     property Br: TCT_TextLineBreak read FA_Br;
     property Fld: TCT_TextField read FA_Fld;
     end;

     TEG_TextRunXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TEG_TextRun;
public
     function  Add: TEG_TextRun;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML);
     procedure Assign(AItem: TEG_TextRunXpgList);
     procedure CopyTo(AItem: TEG_TextRunXpgList);
     property Items[Index: integer]: TEG_TextRun read GetItems; default;
     end;

     TCT_Connection = class(TXPGBase)
protected
     FId: integer;
     FIdx: integer;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Connection);
     procedure CopyTo(AItem: TCT_Connection);

     property Id: integer read FId write FId;
     property Idx: integer read FIdx write FIdx;
     end;

     TCT_Transform2D = class(TXPGBase)
protected
     FRot: integer;
     FFlipH: boolean;
     FFlipV: boolean;
     FA_Off: TCT_Point2D;
     FA_Ext: TCT_PositiveSize2D;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Transform2D);
     procedure CopyTo(AItem: TCT_Transform2D);
     function  Create_A_Off: TCT_Point2D;
     function  Create_A_Ext: TCT_PositiveSize2D;

     property Rot: integer read FRot write FRot;
     property FlipH: boolean read FFlipH write FFlipH;
     property FlipV: boolean read FFlipV write FFlipV;
     property A_Off: TCT_Point2D read FA_Off;
     property A_Ext: TCT_PositiveSize2D read FA_Ext;
     end;

     TEG_Geometry = class(TXPGBase)
protected
     FA_CustGeom: TCT_CustomGeometry2D;
     FA_PrstGeom: TCT_PresetGeometry2D;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TEG_Geometry);
     procedure CopyTo(AItem: TEG_Geometry);
     function  Create_CustGeom: TCT_CustomGeometry2D;
     function  Create_PrstGeom: TCT_PresetGeometry2D;

     property CustGeom: TCT_CustomGeometry2D read FA_CustGeom;
     property PrstGeom: TCT_PresetGeometry2D read FA_PrstGeom;
     end;

     TCT_StyleMatrixReference = class(TXPGBase)
protected
     FIdx: integer;
     FA_EG_ColorChoice: TEG_ColorChoice;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_StyleMatrixReference);
     procedure CopyTo(AItem: TCT_StyleMatrixReference);

     property Idx: integer read FIdx write FIdx;
     property A_EG_ColorChoice: TEG_ColorChoice read FA_EG_ColorChoice;
     end;

     TCT_FontReference = class(TXPGBase)
protected
     FIdx: TST_FontCollectionIndex;
     FA_EG_ColorChoice: TEG_ColorChoice;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_FontReference);
     procedure CopyTo(AItem: TCT_FontReference);

     property Idx: TST_FontCollectionIndex read FIdx write FIdx;
     property A_EG_ColorChoice: TEG_ColorChoice read FA_EG_ColorChoice;
     end;

     TCT_TextBodyProperties = class(TXPGBase)
protected
     FRot: integer;
     FSpcFirstLastPara: boolean;
     FVertOverflow: TST_TextVertOverflowType;
     FHorzOverflow: TST_TextHorzOverflowType;
     FVert: TST_TextVerticalType;
     FWrap: TST_TextWrappingType;
     FLIns: integer;
     FTIns: integer;
     FRIns: integer;
     FBIns: integer;
     FNumCol: integer;
     FSpcCol: integer;
     FRtlCol: boolean;
     FFromWordArt: boolean;
     FAnchor: TST_TextAnchoringType;
     FAnchorCtr: boolean;
     FForceAA: boolean;
     FUpright: boolean;
     FCompatLnSpc: boolean;
     FA_PrstTxWarp: TCT_PresetTextShape;
     FA_EG_TextAutofit: TEG_TextAutofit;
     FA_Scene3d: TCT_Scene3D;
     FA_EG_Text3D: TEG_Text3D;
     FA_ExtLst: TCT_OfficeArtExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_TextBodyProperties);
     procedure CopyTo(AItem: TCT_TextBodyProperties);
     function  Create_A_PrstTxWarp: TCT_PresetTextShape;
     function  Create_A_Scene3d: TCT_Scene3D;
     function  Create_A_ExtLst: TCT_OfficeArtExtensionList;

     property Rot: integer read FRot write FRot;
     property SpcFirstLastPara: boolean read FSpcFirstLastPara write FSpcFirstLastPara;
     property VertOverflow: TST_TextVertOverflowType read FVertOverflow write FVertOverflow;
     property HorzOverflow: TST_TextHorzOverflowType read FHorzOverflow write FHorzOverflow;
     property Vert: TST_TextVerticalType read FVert write FVert;
     property Wrap: TST_TextWrappingType read FWrap write FWrap;
     property LIns: integer read FLIns write FLIns;
     property TIns: integer read FTIns write FTIns;
     property RIns: integer read FRIns write FRIns;
     property BIns: integer read FBIns write FBIns;
     property NumCol: integer read FNumCol write FNumCol;
     property SpcCol: integer read FSpcCol write FSpcCol;
     property RtlCol: boolean read FRtlCol write FRtlCol;
     property FromWordArt: boolean read FFromWordArt write FFromWordArt;
     property Anchor: TST_TextAnchoringType read FAnchor write FAnchor;
     property AnchorCtr: boolean read FAnchorCtr write FAnchorCtr;
     property ForceAA: boolean read FForceAA write FForceAA;
     property Upright: boolean read FUpright write FUpright;
     property CompatLnSpc: boolean read FCompatLnSpc write FCompatLnSpc;
     property PrstTxWarp: TCT_PresetTextShape read FA_PrstTxWarp;
     property EG_TextAutofit: TEG_TextAutofit read FA_EG_TextAutofit;
     property Scene3d: TCT_Scene3D read FA_Scene3d;
     property Text3D: TEG_Text3D read FA_EG_Text3D;
     property ExtLst: TCT_OfficeArtExtensionList read FA_ExtLst;
     end;

     TCT_TextListStyle = class(TXPGBase)
protected
     FA_DefPPr: TCT_TextParagraphProperties;
     FA_Lvl1pPr: TCT_TextParagraphProperties;
     FA_Lvl2pPr: TCT_TextParagraphProperties;
     FA_Lvl3pPr: TCT_TextParagraphProperties;
     FA_Lvl4pPr: TCT_TextParagraphProperties;
     FA_Lvl5pPr: TCT_TextParagraphProperties;
     FA_Lvl6pPr: TCT_TextParagraphProperties;
     FA_Lvl7pPr: TCT_TextParagraphProperties;
     FA_Lvl8pPr: TCT_TextParagraphProperties;
     FA_Lvl9pPr: TCT_TextParagraphProperties;
     FA_ExtLst: TCT_OfficeArtExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_TextListStyle);
     procedure CopyTo(AItem: TCT_TextListStyle);
     function  Create_A_DefPPr: TCT_TextParagraphProperties;
     function  Create_A_Lvl1pPr: TCT_TextParagraphProperties;
     function  Create_A_Lvl2pPr: TCT_TextParagraphProperties;
     function  Create_A_Lvl3pPr: TCT_TextParagraphProperties;
     function  Create_A_Lvl4pPr: TCT_TextParagraphProperties;
     function  Create_A_Lvl5pPr: TCT_TextParagraphProperties;
     function  Create_A_Lvl6pPr: TCT_TextParagraphProperties;
     function  Create_A_Lvl7pPr: TCT_TextParagraphProperties;
     function  Create_A_Lvl8pPr: TCT_TextParagraphProperties;
     function  Create_A_Lvl9pPr: TCT_TextParagraphProperties;
     function  Create_A_ExtLst: TCT_OfficeArtExtensionList;

     property A_DefPPr: TCT_TextParagraphProperties read FA_DefPPr;
     property A_Lvl1pPr: TCT_TextParagraphProperties read FA_Lvl1pPr;
     property A_Lvl2pPr: TCT_TextParagraphProperties read FA_Lvl2pPr;
     property A_Lvl3pPr: TCT_TextParagraphProperties read FA_Lvl3pPr;
     property A_Lvl4pPr: TCT_TextParagraphProperties read FA_Lvl4pPr;
     property A_Lvl5pPr: TCT_TextParagraphProperties read FA_Lvl5pPr;
     property A_Lvl6pPr: TCT_TextParagraphProperties read FA_Lvl6pPr;
     property A_Lvl7pPr: TCT_TextParagraphProperties read FA_Lvl7pPr;
     property A_Lvl8pPr: TCT_TextParagraphProperties read FA_Lvl8pPr;
     property A_Lvl9pPr: TCT_TextParagraphProperties read FA_Lvl9pPr;
     property A_ExtLst: TCT_OfficeArtExtensionList read FA_ExtLst;
     end;

     TCT_TextParagraph = class(TXPGBase)
protected
     FA_PPr: TCT_TextParagraphProperties;
     FA_EG_TextRuns: TEG_TextRunXpgList;
     FA_EndParaRPr: TCT_TextCharacterProperties;
public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_TextParagraph);
     procedure CopyTo(AItem: TCT_TextParagraph);
     function  Create_PPr: TCT_TextParagraphProperties;
     function  Create_EndParaRPr: TCT_TextCharacterProperties;

     function  AppendText(AText: AxUCString): TCT_RegularTextRun; overload;
     function  AppendText(AText: AxUCString; ASize: double; ABold,AItalic,AUnderline: boolean): TCT_RegularTextRun; overload;

     property PPr: TCT_TextParagraphProperties read FA_PPr;
     property TextRuns: TEG_TextRunXpgList read FA_EG_TextRuns;
     property EndParaRPr: TCT_TextCharacterProperties read FA_EndParaRPr;
     end;

     TCT_TextParagraphXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_TextParagraph;
public
     function  Add: TCT_TextParagraph;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_TextParagraphXpgList);
     procedure CopyTo(AItem: TCT_TextParagraphXpgList);
     property Items[Index: integer]: TCT_TextParagraph read GetItems; default;
     end;

     TCT_GroupTransform2D = class(TXPGBase)
protected
     FRot: integer;
     FFlipH: boolean;
     FFlipV: boolean;
     FA_Off: TCT_Point2D;
     FA_Ext: TCT_PositiveSize2D;
     FA_ChOff: TCT_Point2D;
     FA_ChExt: TCT_PositiveSize2D;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_GroupTransform2D);
     procedure CopyTo(AItem: TCT_GroupTransform2D);
     function  Create_A_Off: TCT_Point2D;
     function  Create_A_Ext: TCT_PositiveSize2D;
     function  Create_A_ChOff: TCT_Point2D;
     function  Create_A_ChExt: TCT_PositiveSize2D;

     property Rot: integer read FRot write FRot;
     property FlipH: boolean read FFlipH write FFlipH;
     property FlipV: boolean read FFlipV write FFlipV;
     property A_Off: TCT_Point2D read FA_Off;
     property A_Ext: TCT_PositiveSize2D read FA_Ext;
     property A_ChOff: TCT_Point2D read FA_ChOff;
     property A_ChExt: TCT_PositiveSize2D read FA_ChExt;
     end;

     TCT_ShapeProperties = class(TXPGBase)
protected
     FBwMode: TST_BlackWhiteMode;
     FA_Xfrm: TCT_Transform2D;
     FA_EG_Geometry: TEG_Geometry;
     FA_EG_FillProperties: TEG_FillProperties;
     FA_Ln: TCT_LineProperties;
     FA_EG_EffectProperties: TEG_EffectProperties;
     FA_Scene3d: TCT_Scene3D;
     FA_Sp3d: TCT_Shape3D;
     FA_ExtLst: TCT_OfficeArtExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_ShapeProperties);
     procedure CopyTo(AItem: TCT_ShapeProperties);
     function  Create_Xfrm: TCT_Transform2D;
     function  Create_Ln: TCT_LineProperties;
     function  Create_Scene3d: TCT_Scene3D;
     function  Create_Sp3d: TCT_Shape3D;
     function  Create_ExtLst: TCT_OfficeArtExtensionList;

     procedure  Free_Ln;

     property BwMode: TST_BlackWhiteMode read FBwMode write FBwMode;
     property Xfrm: TCT_Transform2D read FA_Xfrm;
     property Geometry: TEG_Geometry read FA_EG_Geometry;
     property FillProperties: TEG_FillProperties read FA_EG_FillProperties;
     property Ln: TCT_LineProperties read FA_Ln;
     property EffectProperties: TEG_EffectProperties read FA_EG_EffectProperties;
     property Scene3d: TCT_Scene3D read FA_Scene3d;
     property Sp3d: TCT_Shape3D read FA_Sp3d;
     property ExtLst: TCT_OfficeArtExtensionList read FA_ExtLst;
     end;

     TCT_ShapeStyle = class(TXPGBase)
protected
     FA_LnRef: TCT_StyleMatrixReference;
     FA_FillRef: TCT_StyleMatrixReference;
     FA_EffectRef: TCT_StyleMatrixReference;
     FA_FontRef: TCT_FontReference;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_ShapeStyle);
     procedure CopyTo(AItem: TCT_ShapeStyle);
     function  Create_A_LnRef: TCT_StyleMatrixReference;
     function  Create_A_FillRef: TCT_StyleMatrixReference;
     function  Create_A_EffectRef: TCT_StyleMatrixReference;
     function  Create_A_FontRef: TCT_FontReference;

     property A_LnRef: TCT_StyleMatrixReference read FA_LnRef;
     property A_FillRef: TCT_StyleMatrixReference read FA_FillRef;
     property A_EffectRef: TCT_StyleMatrixReference read FA_EffectRef;
     property A_FontRef: TCT_FontReference read FA_FontRef;
     end;

     TCT_TextBody = class(TXPGBase)
private
     function  GetPLainText: AxUCString;
     procedure SetPlainText(const Value: AxUCString);
protected
     FA_BodyPr: TCT_TextBodyProperties;
     FA_LstStyle: TCT_TextListStyle;
     FA_PXpgList: TCT_TextParagraphXpgList;

public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure Assign(AItem: TCT_TextBody);
     procedure CopyTo(AItem: TCT_TextBody);
     function  Create_BodyPr: TCT_TextBodyProperties;
     function  Create_LstStyle: TCT_TextListStyle;
     function  Create_Paras: TCT_TextParagraphXpgList;

     function  MaxFontHeight: double;

     property BodyPr: TCT_TextBodyProperties read FA_BodyPr;
     property LstStyle: TCT_TextListStyle read FA_LstStyle;
     property Paras: TCT_TextParagraphXpgList read FA_PXpgList;
     property PlainText: AxUCString read GetPLainText write SetPlainText;
     end;

     TCT_GroupShapeProperties = class(TXPGBase)
protected
     FBwMode: TST_BlackWhiteMode;
     FA_Xfrm: TCT_GroupTransform2D;
     FA_EG_FillProperties: TEG_FillProperties;
     FA_EG_EffectProperties: TEG_EffectProperties;
     FA_Scene3d: TCT_Scene3D;
     FA_ExtLst: TCT_OfficeArtExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_GroupShapeProperties);
     procedure CopyTo(AItem: TCT_GroupShapeProperties);
     function  Create_A_Xfrm: TCT_GroupTransform2D;
     function  Create_A_Scene3d: TCT_Scene3D;
     function  Create_A_ExtLst: TCT_OfficeArtExtensionList;

     property BwMode: TST_BlackWhiteMode read FBwMode write FBwMode;
     property A_Xfrm: TCT_GroupTransform2D read FA_Xfrm;
     property A_EG_FillProperties: TEG_FillProperties read FA_EG_FillProperties;
     property A_EG_EffectProperties: TEG_EffectProperties read FA_EG_EffectProperties;
     property A_Scene3d: TCT_Scene3D read FA_Scene3d;
     property A_ExtLst: TCT_OfficeArtExtensionList read FA_ExtLst;
     end;

     TCT_NonVisualDrawingProps = class(TXPGBase)
protected
     FId: integer;
     FName: AxUCString;
     FDescr: AxUCString;
     FHidden: boolean;
     FA_HlinkClick: TCT_Hyperlink;
     FA_HlinkHover: TCT_Hyperlink;
     FA_ExtLst: TCT_OfficeArtExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_NonVisualDrawingProps);
     procedure CopyTo(AItem: TCT_NonVisualDrawingProps);
     function  Create_A_HlinkClick: TCT_Hyperlink;
     function  Create_A_HlinkHover: TCT_Hyperlink;
     function  Create_A_ExtLst: TCT_OfficeArtExtensionList;

     property Id: integer read FId write FId;
     property Name: AxUCString read FName write FName;
     property Descr: AxUCString read FDescr write FDescr;
     property Hidden: boolean read FHidden write FHidden;
     property A_HlinkClick: TCT_Hyperlink read FA_HlinkClick;
     property A_HlinkHover: TCT_Hyperlink read FA_HlinkHover;
     property A_ExtLst: TCT_OfficeArtExtensionList read FA_ExtLst;
     end;

     TCT_ShapeLocking = class(TXPGBase)
protected
     FNoGrp: boolean;
     FNoSelect: boolean;
     FNoRot: boolean;
     FNoChangeAspect: boolean;
     FNoMove: boolean;
     FNoResize: boolean;
     FNoEditPoints: boolean;
     FNoAdjustHandles: boolean;
     FNoChangeArrowheads: boolean;
     FNoChangeShapeType: boolean;
     FNoTextEdit: boolean;
     FA_ExtLst: TCT_OfficeArtExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_ShapeLocking);
     procedure CopyTo(AItem: TCT_ShapeLocking);
     function  Create_A_ExtLst: TCT_OfficeArtExtensionList;

     property NoGrp: boolean read FNoGrp write FNoGrp;
     property NoSelect: boolean read FNoSelect write FNoSelect;
     property NoRot: boolean read FNoRot write FNoRot;
     property NoChangeAspect: boolean read FNoChangeAspect write FNoChangeAspect;
     property NoMove: boolean read FNoMove write FNoMove;
     property NoResize: boolean read FNoResize write FNoResize;
     property NoEditPoints: boolean read FNoEditPoints write FNoEditPoints;
     property NoAdjustHandles: boolean read FNoAdjustHandles write FNoAdjustHandles;
     property NoChangeArrowheads: boolean read FNoChangeArrowheads write FNoChangeArrowheads;
     property NoChangeShapeType: boolean read FNoChangeShapeType write FNoChangeShapeType;
     property NoTextEdit: boolean read FNoTextEdit write FNoTextEdit;
     property A_ExtLst: TCT_OfficeArtExtensionList read FA_ExtLst;
     end;

     TCT_NonVisualDrawingShapeProps = class(TXPGBase)
protected
     FTxBox: boolean;
     FA_SpLocks: TCT_ShapeLocking;
     FA_ExtLst: TCT_OfficeArtExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_NonVisualDrawingShapeProps);
     procedure CopyTo(AItem: TCT_NonVisualDrawingShapeProps);
     function  Create_A_SpLocks: TCT_ShapeLocking;
     function  Create_A_ExtLst: TCT_OfficeArtExtensionList;

     property TxBox: boolean read FTxBox write FTxBox;
     property A_SpLocks: TCT_ShapeLocking read FA_SpLocks;
     property A_ExtLst: TCT_OfficeArtExtensionList read FA_ExtLst;
     end;

     TCT_ShapeNonVisual = class(TXPGBase)
protected
     FCNvPr: TCT_NonVisualDrawingProps;
     FCNvSpPr: TCT_NonVisualDrawingShapeProps;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_ShapeNonVisual);
     procedure CopyTo(AItem: TCT_ShapeNonVisual);
     function  Create_CNvPr: TCT_NonVisualDrawingProps;
     function  Create_CNvSpPr: TCT_NonVisualDrawingShapeProps;

     property CNvPr: TCT_NonVisualDrawingProps read FCNvPr;
     property CNvSpPr: TCT_NonVisualDrawingShapeProps read FCNvSpPr;
     end;


     TCT_Shape = class(TXPGBase)
protected
     FMacro: AxUCString;
     FTextlink: AxUCString;
     FFLocksText: boolean;
     FFPublished: boolean;
     FNvSpPr: TCT_ShapeNonVisual;
     FSpPr: TCT_ShapeProperties;
     FStyle: TCT_ShapeStyle;
     FTxBody: TCT_TextBody;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Shape);
     procedure CopyTo(AItem: TCT_Shape);
     function  Create_NvSpPr: TCT_ShapeNonVisual;
     function  Create_SpPr: TCT_ShapeProperties;
     function  Create_Style: TCT_ShapeStyle;
     function  Create_TxBody: TCT_TextBody;

     property Macro: AxUCString read FMacro write FMacro;
     property Textlink: AxUCString read FTextlink write FTextlink;
     property FLocksText: boolean read FFLocksText write FFLocksText;
     property FPublished: boolean read FFPublished write FFPublished;
     property NvSpPr: TCT_ShapeNonVisual read FNvSpPr;
     property SpPr: TCT_ShapeProperties read FSpPr;
     property Style: TCT_ShapeStyle read FStyle;
     property TxBody: TCT_TextBody read FTxBody;
     end;

     TCT_SupplementalFont = class(TXPGBase)
protected
     FScript: AxUCString;
     FTypeface: AxUCString;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_SupplementalFont);
     procedure CopyTo(AItem: TCT_SupplementalFont);

     property Script: AxUCString read FScript write FScript;
     property Typeface: AxUCString read FTypeface write FTypeface;
     end;

     TCT_SupplementalFontXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_SupplementalFont;
public
     function  Add: TCT_SupplementalFont;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_SupplementalFontXpgList);
     procedure CopyTo(AItem: TCT_SupplementalFontXpgList);
     property Items[Index: integer]: TCT_SupplementalFont read GetItems; default;
     end;

     TCT_EffectStyleItem = class(TXPGBase)
protected
     FA_EG_EffectProperties: TEG_EffectProperties;
     FA_Scene3d: TCT_Scene3D;
     FA_Sp3d: TCT_Shape3D;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_EffectStyleItem);
     procedure CopyTo(AItem: TCT_EffectStyleItem);
     function  Create_A_Scene3d: TCT_Scene3D;
     function  Create_A_Sp3d: TCT_Shape3D;

     property A_EG_EffectProperties: TEG_EffectProperties read FA_EG_EffectProperties;
     property A_Scene3d: TCT_Scene3D read FA_Scene3d;
     property A_Sp3d: TCT_Shape3D read FA_Sp3d;
     end;

     TCT_EffectStyleItemXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_EffectStyleItem;
public
     function  Add: TCT_EffectStyleItem;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_EffectStyleItemXpgList);
     procedure CopyTo(AItem: TCT_EffectStyleItemXpgList);
     property Items[Index: integer]: TCT_EffectStyleItem read GetItems; default;
     end;

     TCT_FontCollection = class(TXPGBase)
protected
     FA_Latin: TCT_TextFont;
     FA_Ea: TCT_TextFont;
     FA_Cs: TCT_TextFont;
     FA_FontXpgList: TCT_SupplementalFontXpgList;
     FA_ExtLst: TCT_OfficeArtExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_FontCollection);
     procedure CopyTo(AItem: TCT_FontCollection);
     function  Create_A_Latin: TCT_TextFont;
     function  Create_A_Ea: TCT_TextFont;
     function  Create_A_Cs: TCT_TextFont;
     function  Create_A_FontXpgList: TCT_SupplementalFontXpgList;
     function  Create_A_ExtLst: TCT_OfficeArtExtensionList;

     property A_Latin: TCT_TextFont read FA_Latin;
     property A_Ea: TCT_TextFont read FA_Ea;
     property A_Cs: TCT_TextFont read FA_Cs;
     property A_FontXpgList: TCT_SupplementalFontXpgList read FA_FontXpgList;
     property A_ExtLst: TCT_OfficeArtExtensionList read FA_ExtLst;
     end;

     TCT_FillStyleList = class(TXPGBase)
protected
     FA_EG_FillProperties: TEG_FillProperties;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_FillStyleList);
     procedure CopyTo(AItem: TCT_FillStyleList);

     property A_EG_FillProperties: TEG_FillProperties read FA_EG_FillProperties;
     end;

     TCT_LineStyleList = class(TXPGBase)
protected
     FA_LnXpgList: TCT_LinePropertiesXpgList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_LineStyleList);
     procedure CopyTo(AItem: TCT_LineStyleList);
     function  Create_A_LnXpgList: TCT_LinePropertiesXpgList;

     property A_LnXpgList: TCT_LinePropertiesXpgList read FA_LnXpgList;
     end;

     TCT_EffectStyleList = class(TXPGBase)
protected
     FA_EffectStyleXpgList: TCT_EffectStyleItemXpgList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_EffectStyleList);
     procedure CopyTo(AItem: TCT_EffectStyleList);
     function  Create_A_EffectStyleXpgList: TCT_EffectStyleItemXpgList;

     property A_EffectStyleXpgList: TCT_EffectStyleItemXpgList read FA_EffectStyleXpgList;
     end;

     TCT_BackgroundFillStyleList = class(TXPGBase)
protected
     FA_EG_FillProperties: TEG_FillProperties;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_BackgroundFillStyleList);
     procedure CopyTo(AItem: TCT_BackgroundFillStyleList);

     property A_EG_FillProperties: TEG_FillProperties read FA_EG_FillProperties;
     end;

     TCT_Ratio = class(TXPGBase)
protected
     FN: integer;
     FD: integer;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Ratio);
     procedure CopyTo(AItem: TCT_Ratio);

     property N: integer read FN write FN;
     property D: integer read FD write FD;
     end;

     TCT_CustomColor = class(TXPGBase)
protected
     FName: AxUCString;
     FA_EG_ColorChoice: TEG_ColorChoice;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_CustomColor);
     procedure CopyTo(AItem: TCT_CustomColor);

     property Name: AxUCString read FName write FName;
     property A_EG_ColorChoice: TEG_ColorChoice read FA_EG_ColorChoice;
     end;

     TCT_CustomColorXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_CustomColor;
public
     function  Add: TCT_CustomColor;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_CustomColorXpgList);
     procedure CopyTo(AItem: TCT_CustomColorXpgList);
     property Items[Index: integer]: TCT_CustomColor read GetItems; default;
     end;

     TCT_ColorScheme = class(TXPGBase)
protected
     FName: AxUCString;
     FA_Dk1: TCT_Color;
     FA_Lt1: TCT_Color;
     FA_Dk2: TCT_Color;
     FA_Lt2: TCT_Color;
     FA_Accent1: TCT_Color;
     FA_Accent2: TCT_Color;
     FA_Accent3: TCT_Color;
     FA_Accent4: TCT_Color;
     FA_Accent5: TCT_Color;
     FA_Accent6: TCT_Color;
     FA_Hlink: TCT_Color;
     FA_FolHlink: TCT_Color;
     FA_ExtLst: TCT_OfficeArtExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_ColorScheme);
     procedure CopyTo(AItem: TCT_ColorScheme);
     function  Create_A_Dk1: TCT_Color;
     function  Create_A_Lt1: TCT_Color;
     function  Create_A_Dk2: TCT_Color;
     function  Create_A_Lt2: TCT_Color;
     function  Create_A_Accent1: TCT_Color;
     function  Create_A_Accent2: TCT_Color;
     function  Create_A_Accent3: TCT_Color;
     function  Create_A_Accent4: TCT_Color;
     function  Create_A_Accent5: TCT_Color;
     function  Create_A_Accent6: TCT_Color;
     function  Create_A_Hlink: TCT_Color;
     function  Create_A_FolHlink: TCT_Color;
     function  Create_A_ExtLst: TCT_OfficeArtExtensionList;

     property Name: AxUCString read FName write FName;
     property A_Dk1: TCT_Color read FA_Dk1;
     property A_Lt1: TCT_Color read FA_Lt1;
     property A_Dk2: TCT_Color read FA_Dk2;
     property A_Lt2: TCT_Color read FA_Lt2;
     property A_Accent1: TCT_Color read FA_Accent1;
     property A_Accent2: TCT_Color read FA_Accent2;
     property A_Accent3: TCT_Color read FA_Accent3;
     property A_Accent4: TCT_Color read FA_Accent4;
     property A_Accent5: TCT_Color read FA_Accent5;
     property A_Accent6: TCT_Color read FA_Accent6;
     property A_Hlink: TCT_Color read FA_Hlink;
     property A_FolHlink: TCT_Color read FA_FolHlink;
     property A_ExtLst: TCT_OfficeArtExtensionList read FA_ExtLst;
     end;

     TCT_FontScheme = class(TXPGBase)
protected
     FName: AxUCString;
     FA_MajorFont: TCT_FontCollection;
     FA_MinorFont: TCT_FontCollection;
     FA_ExtLst: TCT_OfficeArtExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_FontScheme);
     procedure CopyTo(AItem: TCT_FontScheme);
     function  Create_A_MajorFont: TCT_FontCollection;
     function  Create_A_MinorFont: TCT_FontCollection;
     function  Create_A_ExtLst: TCT_OfficeArtExtensionList;

     property Name: AxUCString read FName write FName;
     property A_MajorFont: TCT_FontCollection read FA_MajorFont;
     property A_MinorFont: TCT_FontCollection read FA_MinorFont;
     property A_ExtLst: TCT_OfficeArtExtensionList read FA_ExtLst;
     end;

     TCT_StyleMatrix = class(TXPGBase)
protected
     FName: AxUCString;
     FA_FillStyleLst: TCT_FillStyleList;
     FA_LnStyleLst: TCT_LineStyleList;
     FA_EffectStyleLst: TCT_EffectStyleList;
     FA_BgFillStyleLst: TCT_BackgroundFillStyleList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_StyleMatrix);
     procedure CopyTo(AItem: TCT_StyleMatrix);
     function  Create_A_FillStyleLst: TCT_FillStyleList;
     function  Create_A_LnStyleLst: TCT_LineStyleList;
     function  Create_A_EffectStyleLst: TCT_EffectStyleList;
     function  Create_A_BgFillStyleLst: TCT_BackgroundFillStyleList;

     property Name: AxUCString read FName write FName;
     property A_FillStyleLst: TCT_FillStyleList read FA_FillStyleLst;
     property A_LnStyleLst: TCT_LineStyleList read FA_LnStyleLst;
     property A_EffectStyleLst: TCT_EffectStyleList read FA_EffectStyleLst;
     property A_BgFillStyleLst: TCT_BackgroundFillStyleList read FA_BgFillStyleLst;
     end;

     TEG_TextGeometry = class(TXPGBase)
protected
     FA_CustGeom: TCT_CustomGeometry2D;
     FA_PrstTxWarp: TCT_PresetTextShape;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TEG_TextGeometry);
     procedure CopyTo(AItem: TEG_TextGeometry);
     function  Create_A_CustGeom: TCT_CustomGeometry2D;
     function  Create_A_PrstTxWarp: TCT_PresetTextShape;

     property A_CustGeom: TCT_CustomGeometry2D read FA_CustGeom;
     property A_PrstTxWarp: TCT_PresetTextShape read FA_PrstTxWarp;
     end;

     TCT_Scale2D = class(TXPGBase)
protected
     FA_Sx: TCT_Ratio;
     FA_Sy: TCT_Ratio;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Scale2D);
     procedure CopyTo(AItem: TCT_Scale2D);
     function  Create_A_Sx: TCT_Ratio;
     function  Create_A_Sy: TCT_Ratio;

     property A_Sx: TCT_Ratio read FA_Sx;
     property A_Sy: TCT_Ratio read FA_Sy;
     end;

     TCT_FillProperties = class(TXPGBase)
protected
     FA_EG_FillProperties: TEG_FillProperties;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_FillProperties);
     procedure CopyTo(AItem: TCT_FillProperties);

     property A_EG_FillProperties: TEG_FillProperties read FA_EG_FillProperties;
     end;

     TCT_EffectProperties = class(TXPGBase)
protected
     FA_EG_EffectProperties: TEG_EffectProperties;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_EffectProperties);
     procedure CopyTo(AItem: TCT_EffectProperties);

     property A_EG_EffectProperties: TEG_EffectProperties read FA_EG_EffectProperties;
     end;

     TCT_CustomColorList = class(TXPGBase)
protected
     FA_CustClrXpgList: TCT_CustomColorXpgList;
public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_CustomColorList);
     procedure CopyTo(AItem: TCT_CustomColorList);
     function  Create_A_CustClrXpgList: TCT_CustomColorXpgList;

     property A_CustClrXpgList: TCT_CustomColorXpgList read FA_CustClrXpgList;
     end;

     TCT_ColorMRU = class(TXPGBase)
protected
     FA_EG_ColorChoice: TEG_ColorChoice;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_ColorMRU);
     procedure CopyTo(AItem: TCT_ColorMRU);

     property A_EG_ColorChoice: TEG_ColorChoice read FA_EG_ColorChoice;
     end;

     TCT_BaseStyles = class(TXPGBase)
protected
     FA_ClrScheme: TCT_ColorScheme;
     FA_FontScheme: TCT_FontScheme;
     FA_FmtScheme: TCT_StyleMatrix;
     FA_ExtLst: TCT_OfficeArtExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_BaseStyles);
     procedure CopyTo(AItem: TCT_BaseStyles);
     function  Create_A_ClrScheme: TCT_ColorScheme;
     function  Create_A_FontScheme: TCT_FontScheme;
     function  Create_A_FmtScheme: TCT_StyleMatrix;
     function  Create_A_ExtLst: TCT_OfficeArtExtensionList;

     property A_ClrScheme: TCT_ColorScheme read FA_ClrScheme;
     property A_FontScheme: TCT_FontScheme read FA_FontScheme;
     property A_FmtScheme: TCT_StyleMatrix read FA_FmtScheme;
     property A_ExtLst: TCT_OfficeArtExtensionList read FA_ExtLst;
     end;

     TCT_PositiveFixedPercentage = class(TXPGBase)
protected
     FVal: integer;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_PositiveFixedPercentage);
     procedure CopyTo(AItem: TCT_PositiveFixedPercentage);

     property Val: integer read FVal write FVal;
     end;

     TCT_InverseGammaTransform = class(TXPGBase)
public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_InverseGammaTransform);
     procedure CopyTo(AItem: TCT_InverseGammaTransform);

     end;

     TCT_Angle = class(TXPGBase)
protected
     FVal: integer;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Angle);
     procedure CopyTo(AItem: TCT_Angle);

     property Val: integer read FVal write FVal;
     end;

     TCT_Percentage = class(TXPGBase)
protected
     FVal: integer;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Percentage);
     procedure CopyTo(AItem: TCT_Percentage);

     property Val: integer read FVal write FVal;
     end;

     TCT_ComplementTransform = class(TXPGBase)
public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_ComplementTransform);
     procedure CopyTo(AItem: TCT_ComplementTransform);

     end;

     TCT_PositivePercentage = class(TXPGBase)
protected
     FVal: integer;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_PositivePercentage);
     procedure CopyTo(AItem: TCT_PositivePercentage);

     property Val: integer read FVal write FVal;
     end;

     TCT_GrayscaleTransform = class(TXPGBase)
public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_GrayscaleTransform);
     procedure CopyTo(AItem: TCT_GrayscaleTransform);

     end;

     TCT_GammaTransform = class(TXPGBase)
public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_GammaTransform);
     procedure CopyTo(AItem: TCT_GammaTransform);

     end;

     TCT_FixedPercentage = class(TXPGBase)
protected
     FVal: integer;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_FixedPercentage);
     procedure CopyTo(AItem: TCT_FixedPercentage);

     property Val: integer read FVal write FVal;
     end;

     TCT_PositiveFixedAngle = class(TXPGBase)
protected
     FVal: integer;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_PositiveFixedAngle);
     procedure CopyTo(AItem: TCT_PositiveFixedAngle);

     property Val: integer read FVal write FVal;
     end;

     TCT_InverseTransform = class(TXPGBase)
public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_InverseTransform);
     procedure CopyTo(AItem: TCT_InverseTransform);

     end;

procedure ReadUnionTST_AdjAngle(AValue: AxUCString; APtr: PST_AdjAngle);
function  WriteUnionTST_AdjAngle(APtr: PST_AdjAngle): AxUCString;
procedure ReadUnionTST_AdjCoordinate(AValue: AxUCString; APtr: PST_AdjCoordinate);
function  WriteUnionTST_AdjCoordinate(APtr: PST_AdjCoordinate): AxUCString;

function  SolidFillToRGB(AFill: TCT_SolidColorFillProperties): longword; overload;
function  SpPrFillToRGB(ASpPr: TCT_ShapeProperties): longword; overload;

implementation

function SolidFillToRGB(AFill: TCT_SolidColorFillProperties): longword;
var
  RGB: TRGBRec;
  HLS: THLSRec;
  V: double;

function Lighten(Lum: integer; Value: double): integer;
begin
  Result := Round(Lum * (1.0 - Value) + (255 - 255 * (1.0 - Value)));
end;

begin
  Result := $00FFFFFF;
  if AFill <> Nil then begin
    if AFill.ColorChoice.SchemeClr <> Nil then begin
      case AFill.ColorChoice.SchemeClr.Val of
        stscvBg1     : RGB.Value := Xc12DefColorScheme[cscLt1];
        stscvTx1     : RGB.Value := Xc12DefColorScheme[cscDk1];
        stscvBg2     : RGB.Value := Xc12DefColorScheme[cscLt2];
        stscvTx2     : RGB.Value := Xc12DefColorScheme[cscDk2];
        stscvAccent1 : RGB.Value := Xc12DefColorScheme[cscAccent1];
        stscvAccent2 : RGB.Value := Xc12DefColorScheme[cscAccent2];
        stscvAccent3 : RGB.Value := Xc12DefColorScheme[cscAccent3];
        stscvAccent4 : RGB.Value := Xc12DefColorScheme[cscAccent4];
        stscvAccent5 : RGB.Value := Xc12DefColorScheme[cscAccent5];
        stscvAccent6 : RGB.Value := Xc12DefColorScheme[cscAccent6];
        stscvHlink   : RGB.Value := Xc12DefColorScheme[cscHlink];
        stscvFolHlink: RGB.Value := Xc12DefColorScheme[cscFolHlink];
        stscvPhClr   : RGB.Value := Xc12DefColorScheme[cscExtList];
        stscvDk1     : RGB.Value := Xc12DefColorScheme[cscDk1];
        stscvLt1     : RGB.Value := Xc12DefColorScheme[cscLt1];
        stscvDk2     : RGB.Value := Xc12DefColorScheme[cscDk2];
        stscvLt2     : RGB.Value := Xc12DefColorScheme[cscLt2];
      end;
      HLS := RGBToHLS(RGB);
      V := AFill.ColorChoice.SchemeClr.ColorTransform.AsDouble['lumMod'];
      if not IsNaN(V) then begin
        V := V / 100000;

        HLS.L := Lighten(HLS.L,V);
      end;
      V := AFill.ColorChoice.SchemeClr.ColorTransform.AsDouble['lumOff'];
      if not IsNaN(V) then begin
        V := V / 100000;

        if V >= 0.5 then
          HLS.L := HLS.L + Round(128 * V)
        else
          HLS.L := HLS.L - Round(128 * V);
      end;
      RGB := HLSToRGB(HLS);
      Result := RevRGB(RGB.Value);
    end
    else if AFill.ColorChoice.SrgbClr <> Nil then
      Result := AFill.ColorChoice.SrgbClr.Val;
  end;
end;

function SpPrFillToRGB(ASpPr: TCT_ShapeProperties): longword;
begin
  if ASpPr.FillProperties.SolidFill <> Nil then
    Result := SolidFillToRGB(ASpPr.FillProperties.SolidFill)
  else
    Result := $00FFFFFF;
end;

{$ifdef DELPHI_5}
var L_Enums: TStringList;
{$else}
var L_Enums: THashedStringList;
{$endif}

procedure ReadUnionTST_AdjAngle(AValue: AxUCString; APtr: PST_AdjAngle);
begin
  if XmlTryStrToInt(AValue,APtr.Val1) then
  begin
    APtr.nVal := 0;
    Exit;
  end;
  if AValue <> '' then
  begin
    APtr.Val2 := AValue;
    APtr.nVal := 1;
    Exit;
  end;
end;

function  WriteUnionTST_AdjAngle(APtr: PST_AdjAngle): AxUCString;
begin
  case APtr.nVal of
    0: Result := XmlIntToStr(APtr.Val1);
    1: Result := APtr.Val2;
  end;
end;

procedure ReadUnionTST_AdjCoordinate(AValue: AxUCString; APtr: PST_AdjCoordinate);
begin
  if XmlTryStrToInt(AValue,APtr.Val1) then 
  begin
    APtr.nVal := 0;
    Exit;
  end;
  if AValue <> '' then 
  begin
    APtr.Val2 := AValue;
    APtr.nVal := 1;
    Exit;
  end;
end;

function  WriteUnionTST_AdjCoordinate(APtr: PST_AdjCoordinate): AxUCString;
begin
  case APtr.nVal of
    0: Result := XmlIntToStr(APtr.Val1);
    1: Result := APtr.Val2;
  end;
end;

{ TXPGBase }

function  TXPGBase.CheckAssigned: integer;
begin
  Result := 1;
end;

function  TXPGBase.Assigned: boolean;
begin
  Result := FAssigneds <> [];
end;

function  TXPGBase.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
end;

procedure TXPGBase.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
end;

procedure TXPGBase.AfterTag;
begin
end;

procedure AddEnums;
begin
  L_Enums.AddObject('stscvBg1',TObject(0));
  L_Enums.AddObject('stscvTx1',TObject(1));
  L_Enums.AddObject('stscvBg2',TObject(2));
  L_Enums.AddObject('stscvTx2',TObject(3));
  L_Enums.AddObject('stscvAccent1',TObject(4));
  L_Enums.AddObject('stscvAccent2',TObject(5));
  L_Enums.AddObject('stscvAccent3',TObject(6));
  L_Enums.AddObject('stscvAccent4',TObject(7));
  L_Enums.AddObject('stscvAccent5',TObject(8));
  L_Enums.AddObject('stscvAccent6',TObject(9));
  L_Enums.AddObject('stscvHlink',TObject(10));
  L_Enums.AddObject('stscvFolHlink',TObject(11));
  L_Enums.AddObject('stscvPhClr',TObject(12));
  L_Enums.AddObject('stscvDk1',TObject(13));
  L_Enums.AddObject('stscvLt1',TObject(14));
  L_Enums.AddObject('stscvDk2',TObject(15));
  L_Enums.AddObject('stscvLt2',TObject(16));
  L_Enums.AddObject('stbwmClr',TObject(0));
  L_Enums.AddObject('stbwmAuto',TObject(1));
  L_Enums.AddObject('stbwmGray',TObject(2));
  L_Enums.AddObject('stbwmLtGray',TObject(3));
  L_Enums.AddObject('stbwmInvGray',TObject(4));
  L_Enums.AddObject('stbwmGrayWhite',TObject(5));
  L_Enums.AddObject('stbwmBlackGray',TObject(6));
  L_Enums.AddObject('stbwmBlackWhite',TObject(7));
  L_Enums.AddObject('stbwmBlack',TObject(8));
  L_Enums.AddObject('stbwmWhite',TObject(9));
  L_Enums.AddObject('stbwmHidden',TObject(10));
  L_Enums.AddObject('stpcvAliceBlue',TObject(0));
  L_Enums.AddObject('stpcvAntiqueWhite',TObject(1));
  L_Enums.AddObject('stpcvAqua',TObject(2));
  L_Enums.AddObject('stpcvAquamarine',TObject(3));
  L_Enums.AddObject('stpcvAzure',TObject(4));
  L_Enums.AddObject('stpcvBeige',TObject(5));
  L_Enums.AddObject('stpcvBisque',TObject(6));
  L_Enums.AddObject('stpcvBlack',TObject(7));
  L_Enums.AddObject('stpcvBlanchedAlmond',TObject(8));
  L_Enums.AddObject('stpcvBlue',TObject(9));
  L_Enums.AddObject('stpcvBlueViolet',TObject(10));
  L_Enums.AddObject('stpcvBrown',TObject(11));
  L_Enums.AddObject('stpcvBurlyWood',TObject(12));
  L_Enums.AddObject('stpcvCadetBlue',TObject(13));
  L_Enums.AddObject('stpcvChartreuse',TObject(14));
  L_Enums.AddObject('stpcvChocolate',TObject(15));
  L_Enums.AddObject('stpcvCoral',TObject(16));
  L_Enums.AddObject('stpcvCornflowerBlue',TObject(17));
  L_Enums.AddObject('stpcvCornsilk',TObject(18));
  L_Enums.AddObject('stpcvCrimson',TObject(19));
  L_Enums.AddObject('stpcvCyan',TObject(20));
  L_Enums.AddObject('stpcvDkBlue',TObject(21));
  L_Enums.AddObject('stpcvDkCyan',TObject(22));
  L_Enums.AddObject('stpcvDkGoldenrod',TObject(23));
  L_Enums.AddObject('stpcvDkGray',TObject(24));
  L_Enums.AddObject('stpcvDkGreen',TObject(25));
  L_Enums.AddObject('stpcvDkKhaki',TObject(26));
  L_Enums.AddObject('stpcvDkMagenta',TObject(27));
  L_Enums.AddObject('stpcvDkOliveGreen',TObject(28));
  L_Enums.AddObject('stpcvDkOrange',TObject(29));
  L_Enums.AddObject('stpcvDkOrchid',TObject(30));
  L_Enums.AddObject('stpcvDkRed',TObject(31));
  L_Enums.AddObject('stpcvDkSalmon',TObject(32));
  L_Enums.AddObject('stpcvDkSeaGreen',TObject(33));
  L_Enums.AddObject('stpcvDkSlateBlue',TObject(34));
  L_Enums.AddObject('stpcvDkSlateGray',TObject(35));
  L_Enums.AddObject('stpcvDkTurquoise',TObject(36));
  L_Enums.AddObject('stpcvDkViolet',TObject(37));
  L_Enums.AddObject('stpcvDeepPink',TObject(38));
  L_Enums.AddObject('stpcvDeepSkyBlue',TObject(39));
  L_Enums.AddObject('stpcvDimGray',TObject(40));
  L_Enums.AddObject('stpcvDodgerBlue',TObject(41));
  L_Enums.AddObject('stpcvFirebrick',TObject(42));
  L_Enums.AddObject('stpcvFloralWhite',TObject(43));
  L_Enums.AddObject('stpcvForestGreen',TObject(44));
  L_Enums.AddObject('stpcvFuchsia',TObject(45));
  L_Enums.AddObject('stpcvGainsboro',TObject(46));
  L_Enums.AddObject('stpcvGhostWhite',TObject(47));
  L_Enums.AddObject('stpcvGold',TObject(48));
  L_Enums.AddObject('stpcvGoldenrod',TObject(49));
  L_Enums.AddObject('stpcvGray',TObject(50));
  L_Enums.AddObject('stpcvGreen',TObject(51));
  L_Enums.AddObject('stpcvGreenYellow',TObject(52));
  L_Enums.AddObject('stpcvHoneydew',TObject(53));
  L_Enums.AddObject('stpcvHotPink',TObject(54));
  L_Enums.AddObject('stpcvIndianRed',TObject(55));
  L_Enums.AddObject('stpcvIndigo',TObject(56));
  L_Enums.AddObject('stpcvIvory',TObject(57));
  L_Enums.AddObject('stpcvKhaki',TObject(58));
  L_Enums.AddObject('stpcvLavender',TObject(59));
  L_Enums.AddObject('stpcvLavenderBlush',TObject(60));
  L_Enums.AddObject('stpcvLawnGreen',TObject(61));
  L_Enums.AddObject('stpcvLemonChiffon',TObject(62));
  L_Enums.AddObject('stpcvLtBlue',TObject(63));
  L_Enums.AddObject('stpcvLtCoral',TObject(64));
  L_Enums.AddObject('stpcvLtCyan',TObject(65));
  L_Enums.AddObject('stpcvLtGoldenrodYellow',TObject(66));
  L_Enums.AddObject('stpcvLtGray',TObject(67));
  L_Enums.AddObject('stpcvLtGreen',TObject(68));
  L_Enums.AddObject('stpcvLtPink',TObject(69));
  L_Enums.AddObject('stpcvLtSalmon',TObject(70));
  L_Enums.AddObject('stpcvLtSeaGreen',TObject(71));
  L_Enums.AddObject('stpcvLtSkyBlue',TObject(72));
  L_Enums.AddObject('stpcvLtSlateGray',TObject(73));
  L_Enums.AddObject('stpcvLtSteelBlue',TObject(74));
  L_Enums.AddObject('stpcvLtYellow',TObject(75));
  L_Enums.AddObject('stpcvLime',TObject(76));
  L_Enums.AddObject('stpcvLimeGreen',TObject(77));
  L_Enums.AddObject('stpcvLinen',TObject(78));
  L_Enums.AddObject('stpcvMagenta',TObject(79));
  L_Enums.AddObject('stpcvMaroon',TObject(80));
  L_Enums.AddObject('stpcvMedAquamarine',TObject(81));
  L_Enums.AddObject('stpcvMedBlue',TObject(82));
  L_Enums.AddObject('stpcvMedOrchid',TObject(83));
  L_Enums.AddObject('stpcvMedPurple',TObject(84));
  L_Enums.AddObject('stpcvMedSeaGreen',TObject(85));
  L_Enums.AddObject('stpcvMedSlateBlue',TObject(86));
  L_Enums.AddObject('stpcvMedSpringGreen',TObject(87));
  L_Enums.AddObject('stpcvMedTurquoise',TObject(88));
  L_Enums.AddObject('stpcvMedVioletRed',TObject(89));
  L_Enums.AddObject('stpcvMidnightBlue',TObject(90));
  L_Enums.AddObject('stpcvMintCream',TObject(91));
  L_Enums.AddObject('stpcvMistyRose',TObject(92));
  L_Enums.AddObject('stpcvMoccasin',TObject(93));
  L_Enums.AddObject('stpcvNavajoWhite',TObject(94));
  L_Enums.AddObject('stpcvNavy',TObject(95));
  L_Enums.AddObject('stpcvOldLace',TObject(96));
  L_Enums.AddObject('stpcvOlive',TObject(97));
  L_Enums.AddObject('stpcvOliveDrab',TObject(98));
  L_Enums.AddObject('stpcvOrange',TObject(99));
  L_Enums.AddObject('stpcvOrangeRed',TObject(100));
  L_Enums.AddObject('stpcvOrchid',TObject(101));
  L_Enums.AddObject('stpcvPaleGoldenrod',TObject(102));
  L_Enums.AddObject('stpcvPaleGreen',TObject(103));
  L_Enums.AddObject('stpcvPaleTurquoise',TObject(104));
  L_Enums.AddObject('stpcvPaleVioletRed',TObject(105));
  L_Enums.AddObject('stpcvPapayaWhip',TObject(106));
  L_Enums.AddObject('stpcvPeachPuff',TObject(107));
  L_Enums.AddObject('stpcvPeru',TObject(108));
  L_Enums.AddObject('stpcvPink',TObject(109));
  L_Enums.AddObject('stpcvPlum',TObject(110));
  L_Enums.AddObject('stpcvPowderBlue',TObject(111));
  L_Enums.AddObject('stpcvPurple',TObject(112));
  L_Enums.AddObject('stpcvRed',TObject(113));
  L_Enums.AddObject('stpcvRosyBrown',TObject(114));
  L_Enums.AddObject('stpcvRoyalBlue',TObject(115));
  L_Enums.AddObject('stpcvSaddleBrown',TObject(116));
  L_Enums.AddObject('stpcvSalmon',TObject(117));
  L_Enums.AddObject('stpcvSandyBrown',TObject(118));
  L_Enums.AddObject('stpcvSeaGreen',TObject(119));
  L_Enums.AddObject('stpcvSeaShell',TObject(120));
  L_Enums.AddObject('stpcvSienna',TObject(121));
  L_Enums.AddObject('stpcvSilver',TObject(122));
  L_Enums.AddObject('stpcvSkyBlue',TObject(123));
  L_Enums.AddObject('stpcvSlateBlue',TObject(124));
  L_Enums.AddObject('stpcvSlateGray',TObject(125));
  L_Enums.AddObject('stpcvSnow',TObject(126));
  L_Enums.AddObject('stpcvSpringGreen',TObject(127));
  L_Enums.AddObject('stpcvSteelBlue',TObject(128));
  L_Enums.AddObject('stpcvTan',TObject(129));
  L_Enums.AddObject('stpcvTeal',TObject(130));
  L_Enums.AddObject('stpcvThistle',TObject(131));
  L_Enums.AddObject('stpcvTomato',TObject(132));
  L_Enums.AddObject('stpcvTurquoise',TObject(133));
  L_Enums.AddObject('stpcvViolet',TObject(134));
  L_Enums.AddObject('stpcvWheat',TObject(135));
  L_Enums.AddObject('stpcvWhite',TObject(136));
  L_Enums.AddObject('stpcvWhiteSmoke',TObject(137));
  L_Enums.AddObject('stpcvYellow',TObject(138));
  L_Enums.AddObject('stpcvYellowGreen',TObject(139));
  L_Enums.AddObject('stscvScrollBar',TObject(0));
  L_Enums.AddObject('stscvBackground',TObject(1));
  L_Enums.AddObject('stscvActiveCaption',TObject(2));
  L_Enums.AddObject('stscvInactiveCaption',TObject(3));
  L_Enums.AddObject('stscvMenu',TObject(4));
  L_Enums.AddObject('stscvWindow',TObject(5));
  L_Enums.AddObject('stscvWindowFrame',TObject(6));
  L_Enums.AddObject('stscvMenuText',TObject(7));
  L_Enums.AddObject('stscvWindowText',TObject(8));
  L_Enums.AddObject('stscvCaptionText',TObject(9));
  L_Enums.AddObject('stscvActiveBorder',TObject(10));
  L_Enums.AddObject('stscvInactiveBorder',TObject(11));
  L_Enums.AddObject('stscvAppWorkspace',TObject(12));
  L_Enums.AddObject('stscvHighlight',TObject(13));
  L_Enums.AddObject('stscvHighlightText',TObject(14));
  L_Enums.AddObject('stscvBtnFace',TObject(15));
  L_Enums.AddObject('stscvBtnShadow',TObject(16));
  L_Enums.AddObject('stscvGrayText',TObject(17));
  L_Enums.AddObject('stscvBtnText',TObject(18));
  L_Enums.AddObject('stscvInactiveCaptionText',TObject(19));
  L_Enums.AddObject('stscvBtnHighlight',TObject(20));
  L_Enums.AddObject('stscv3dDkShadow',TObject(21));
  L_Enums.AddObject('stscv3dLight',TObject(22));
  L_Enums.AddObject('stscvInfoText',TObject(23));
  L_Enums.AddObject('stscvInfoBk',TObject(24));
  L_Enums.AddObject('stscvHotLight',TObject(25));
  L_Enums.AddObject('stscvGradientActiveCaption',TObject(26));
  L_Enums.AddObject('stscvGradientInactiveCaption',TObject(27));
  L_Enums.AddObject('stscvMenuHighlight',TObject(28));
  L_Enums.AddObject('stscvMenuBar',TObject(29));
  L_Enums.AddObject('straTl',TObject(0));
  L_Enums.AddObject('straT',TObject(1));
  L_Enums.AddObject('straTr',TObject(2));
  L_Enums.AddObject('straL',TObject(3));
  L_Enums.AddObject('straCtr',TObject(4));
  L_Enums.AddObject('straR',TObject(5));
  L_Enums.AddObject('straBl',TObject(6));
  L_Enums.AddObject('straB',TObject(7));
  L_Enums.AddObject('straBr',TObject(8));
  L_Enums.AddObject('stpstShape',TObject(0));
  L_Enums.AddObject('stpstCircle',TObject(1));
  L_Enums.AddObject('stpstRect',TObject(2));
  L_Enums.AddObject('stpsvShdw1',TObject(0));
  L_Enums.AddObject('stpsvShdw2',TObject(1));
  L_Enums.AddObject('stpsvShdw3',TObject(2));
  L_Enums.AddObject('stpsvShdw4',TObject(3));
  L_Enums.AddObject('stpsvShdw5',TObject(4));
  L_Enums.AddObject('stpsvShdw6',TObject(5));
  L_Enums.AddObject('stpsvShdw7',TObject(6));
  L_Enums.AddObject('stpsvShdw8',TObject(7));
  L_Enums.AddObject('stpsvShdw9',TObject(8));
  L_Enums.AddObject('stpsvShdw10',TObject(9));
  L_Enums.AddObject('stpsvShdw11',TObject(10));
  L_Enums.AddObject('stpsvShdw12',TObject(11));
  L_Enums.AddObject('stpsvShdw13',TObject(12));
  L_Enums.AddObject('stpsvShdw14',TObject(13));
  L_Enums.AddObject('stpsvShdw15',TObject(14));
  L_Enums.AddObject('stpsvShdw16',TObject(15));
  L_Enums.AddObject('stpsvShdw17',TObject(16));
  L_Enums.AddObject('stpsvShdw18',TObject(17));
  L_Enums.AddObject('stpsvShdw19',TObject(18));
  L_Enums.AddObject('stpsvShdw20',TObject(19));
  L_Enums.AddObject('stppvPct5',TObject(0));
  L_Enums.AddObject('stppvPct10',TObject(1));
  L_Enums.AddObject('stppvPct20',TObject(2));
  L_Enums.AddObject('stppvPct25',TObject(3));
  L_Enums.AddObject('stppvPct30',TObject(4));
  L_Enums.AddObject('stppvPct40',TObject(5));
  L_Enums.AddObject('stppvPct50',TObject(6));
  L_Enums.AddObject('stppvPct60',TObject(7));
  L_Enums.AddObject('stppvPct70',TObject(8));
  L_Enums.AddObject('stppvPct75',TObject(9));
  L_Enums.AddObject('stppvPct80',TObject(10));
  L_Enums.AddObject('stppvPct90',TObject(11));
  L_Enums.AddObject('stppvHorz',TObject(12));
  L_Enums.AddObject('stppvVert',TObject(13));
  L_Enums.AddObject('stppvLtHorz',TObject(14));
  L_Enums.AddObject('stppvLtVert',TObject(15));
  L_Enums.AddObject('stppvDkHorz',TObject(16));
  L_Enums.AddObject('stppvDkVert',TObject(17));
  L_Enums.AddObject('stppvNarHorz',TObject(18));
  L_Enums.AddObject('stppvNarVert',TObject(19));
  L_Enums.AddObject('stppvDashHorz',TObject(20));
  L_Enums.AddObject('stppvDashVert',TObject(21));
  L_Enums.AddObject('stppvCross',TObject(22));
  L_Enums.AddObject('stppvDnDiag',TObject(23));
  L_Enums.AddObject('stppvUpDiag',TObject(24));
  L_Enums.AddObject('stppvLtDnDiag',TObject(25));
  L_Enums.AddObject('stppvLtUpDiag',TObject(26));
  L_Enums.AddObject('stppvDkDnDiag',TObject(27));
  L_Enums.AddObject('stppvDkUpDiag',TObject(28));
  L_Enums.AddObject('stppvWdDnDiag',TObject(29));
  L_Enums.AddObject('stppvWdUpDiag',TObject(30));
  L_Enums.AddObject('stppvDashDnDiag',TObject(31));
  L_Enums.AddObject('stppvDashUpDiag',TObject(32));
  L_Enums.AddObject('stppvDiagCross',TObject(33));
  L_Enums.AddObject('stppvSmCheck',TObject(34));
  L_Enums.AddObject('stppvLgCheck',TObject(35));
  L_Enums.AddObject('stppvSmGrid',TObject(36));
  L_Enums.AddObject('stppvLgGrid',TObject(37));
  L_Enums.AddObject('stppvDotGrid',TObject(38));
  L_Enums.AddObject('stppvSmConfetti',TObject(39));
  L_Enums.AddObject('stppvLgConfetti',TObject(40));
  L_Enums.AddObject('stppvHorzBrick',TObject(41));
  L_Enums.AddObject('stppvDiagBrick',TObject(42));
  L_Enums.AddObject('stppvSolidDmnd',TObject(43));
  L_Enums.AddObject('stppvOpenDmnd',TObject(44));
  L_Enums.AddObject('stppvDotDmnd',TObject(45));
  L_Enums.AddObject('stppvPlaid',TObject(46));
  L_Enums.AddObject('stppvSphere',TObject(47));
  L_Enums.AddObject('stppvWeave',TObject(48));
  L_Enums.AddObject('stppvDivot',TObject(49));
  L_Enums.AddObject('stppvShingle',TObject(50));
  L_Enums.AddObject('stppvWave',TObject(51));
  L_Enums.AddObject('stppvTrellis',TObject(52));
  L_Enums.AddObject('stppvZigZag',TObject(53));
  L_Enums.AddObject('sttfmNone',TObject(0));
  L_Enums.AddObject('sttfmX',TObject(1));
  L_Enums.AddObject('sttfmY',TObject(2));
  L_Enums.AddObject('sttfmXy',TObject(3));
  L_Enums.AddObject('stbcEmail',TObject(0));
  L_Enums.AddObject('stbcScreen',TObject(1));
  L_Enums.AddObject('stbcPrint',TObject(2));
  L_Enums.AddObject('stbcHqprint',TObject(3));
  L_Enums.AddObject('stbcNone',TObject(4));
  L_Enums.AddObject('stectSib',TObject(0));
  L_Enums.AddObject('stectTree',TObject(1));
  L_Enums.AddObject('stbmOver',TObject(0));
  L_Enums.AddObject('stbmMult',TObject(1));
  L_Enums.AddObject('stbmScreen',TObject(2));
  L_Enums.AddObject('stbmDarken',TObject(3));
  L_Enums.AddObject('stbmLighten',TObject(4));
  L_Enums.AddObject('ststLine',TObject(0));
  L_Enums.AddObject('ststLineInv',TObject(1));
  L_Enums.AddObject('ststTriangle',TObject(2));
  L_Enums.AddObject('ststRtTriangle',TObject(3));
  L_Enums.AddObject('ststRect',TObject(4));
  L_Enums.AddObject('ststDiamond',TObject(5));
  L_Enums.AddObject('ststParallelogram',TObject(6));
  L_Enums.AddObject('ststTrapezoid',TObject(7));
  L_Enums.AddObject('ststNonIsoscelesTrapezoid',TObject(8));
  L_Enums.AddObject('ststPentagon',TObject(9));
  L_Enums.AddObject('ststHexagon',TObject(10));
  L_Enums.AddObject('ststHeptagon',TObject(11));
  L_Enums.AddObject('ststOctagon',TObject(12));
  L_Enums.AddObject('ststDecagon',TObject(13));
  L_Enums.AddObject('ststDodecagon',TObject(14));
  L_Enums.AddObject('ststStar4',TObject(15));
  L_Enums.AddObject('ststStar5',TObject(16));
  L_Enums.AddObject('ststStar6',TObject(17));
  L_Enums.AddObject('ststStar7',TObject(18));
  L_Enums.AddObject('ststStar8',TObject(19));
  L_Enums.AddObject('ststStar10',TObject(20));
  L_Enums.AddObject('ststStar12',TObject(21));
  L_Enums.AddObject('ststStar16',TObject(22));
  L_Enums.AddObject('ststStar24',TObject(23));
  L_Enums.AddObject('ststStar32',TObject(24));
  L_Enums.AddObject('ststRoundRect',TObject(25));
  L_Enums.AddObject('ststRound1Rect',TObject(26));
  L_Enums.AddObject('ststRound2SameRect',TObject(27));
  L_Enums.AddObject('ststRound2DiagRect',TObject(28));
  L_Enums.AddObject('ststSnipRoundRect',TObject(29));
  L_Enums.AddObject('ststSnip1Rect',TObject(30));
  L_Enums.AddObject('ststSnip2SameRect',TObject(31));
  L_Enums.AddObject('ststSnip2DiagRect',TObject(32));
  L_Enums.AddObject('ststPlaque',TObject(33));
  L_Enums.AddObject('ststEllipse',TObject(34));
  L_Enums.AddObject('ststTeardrop',TObject(35));
  L_Enums.AddObject('ststHomePlate',TObject(36));
  L_Enums.AddObject('ststChevron',TObject(37));
  L_Enums.AddObject('ststPieWedge',TObject(38));
  L_Enums.AddObject('ststPie',TObject(39));
  L_Enums.AddObject('ststBlockArc',TObject(40));
  L_Enums.AddObject('ststDonut',TObject(41));
  L_Enums.AddObject('ststNoSmoking',TObject(42));
  L_Enums.AddObject('ststRightArrow',TObject(43));
  L_Enums.AddObject('ststLeftArrow',TObject(44));
  L_Enums.AddObject('ststUpArrow',TObject(45));
  L_Enums.AddObject('ststDownArrow',TObject(46));
  L_Enums.AddObject('ststStripedRightArrow',TObject(47));
  L_Enums.AddObject('ststNotchedRightArrow',TObject(48));
  L_Enums.AddObject('ststBentUpArrow',TObject(49));
  L_Enums.AddObject('ststLeftRightArrow',TObject(50));
  L_Enums.AddObject('ststUpDownArrow',TObject(51));
  L_Enums.AddObject('ststLeftUpArrow',TObject(52));
  L_Enums.AddObject('ststLeftRightUpArrow',TObject(53));
  L_Enums.AddObject('ststQuadArrow',TObject(54));
  L_Enums.AddObject('ststLeftArrowCallout',TObject(55));
  L_Enums.AddObject('ststRightArrowCallout',TObject(56));
  L_Enums.AddObject('ststUpArrowCallout',TObject(57));
  L_Enums.AddObject('ststDownArrowCallout',TObject(58));
  L_Enums.AddObject('ststLeftRightArrowCallout',TObject(59));
  L_Enums.AddObject('ststUpDownArrowCallout',TObject(60));
  L_Enums.AddObject('ststQuadArrowCallout',TObject(61));
  L_Enums.AddObject('ststBentArrow',TObject(62));
  L_Enums.AddObject('ststUturnArrow',TObject(63));
  L_Enums.AddObject('ststCircularArrow',TObject(64));
  L_Enums.AddObject('ststLeftCircularArrow',TObject(65));
  L_Enums.AddObject('ststLeftRightCircularArrow',TObject(66));
  L_Enums.AddObject('ststCurvedRightArrow',TObject(67));
  L_Enums.AddObject('ststCurvedLeftArrow',TObject(68));
  L_Enums.AddObject('ststCurvedUpArrow',TObject(69));
  L_Enums.AddObject('ststCurvedDownArrow',TObject(70));
  L_Enums.AddObject('ststSwooshArrow',TObject(71));
  L_Enums.AddObject('ststCube',TObject(72));
  L_Enums.AddObject('ststCan',TObject(73));
  L_Enums.AddObject('ststLightningBolt',TObject(74));
  L_Enums.AddObject('ststHeart',TObject(75));
  L_Enums.AddObject('ststSun',TObject(76));
  L_Enums.AddObject('ststMoon',TObject(77));
  L_Enums.AddObject('ststSmileyFace',TObject(78));
  L_Enums.AddObject('ststIrregularSeal1',TObject(79));
  L_Enums.AddObject('ststIrregularSeal2',TObject(80));
  L_Enums.AddObject('ststFoldedCorner',TObject(81));
  L_Enums.AddObject('ststBevel',TObject(82));
  L_Enums.AddObject('ststFrame',TObject(83));
  L_Enums.AddObject('ststHalfFrame',TObject(84));
  L_Enums.AddObject('ststCorner',TObject(85));
  L_Enums.AddObject('ststDiagStripe',TObject(86));
  L_Enums.AddObject('ststChord',TObject(87));
  L_Enums.AddObject('ststArc',TObject(88));
  L_Enums.AddObject('ststLeftBracket',TObject(89));
  L_Enums.AddObject('ststRightBracket',TObject(90));
  L_Enums.AddObject('ststLeftBrace',TObject(91));
  L_Enums.AddObject('ststRightBrace',TObject(92));
  L_Enums.AddObject('ststBracketPair',TObject(93));
  L_Enums.AddObject('ststBracePair',TObject(94));
  L_Enums.AddObject('ststStraightConnector1',TObject(95));
  L_Enums.AddObject('ststBentConnector2',TObject(96));
  L_Enums.AddObject('ststBentConnector3',TObject(97));
  L_Enums.AddObject('ststBentConnector4',TObject(98));
  L_Enums.AddObject('ststBentConnector5',TObject(99));
  L_Enums.AddObject('ststCurvedConnector2',TObject(100));
  L_Enums.AddObject('ststCurvedConnector3',TObject(101));
  L_Enums.AddObject('ststCurvedConnector4',TObject(102));
  L_Enums.AddObject('ststCurvedConnector5',TObject(103));
  L_Enums.AddObject('ststCallout1',TObject(104));
  L_Enums.AddObject('ststCallout2',TObject(105));
  L_Enums.AddObject('ststCallout3',TObject(106));
  L_Enums.AddObject('ststAccentCallout1',TObject(107));
  L_Enums.AddObject('ststAccentCallout2',TObject(108));
  L_Enums.AddObject('ststAccentCallout3',TObject(109));
  L_Enums.AddObject('ststBorderCallout1',TObject(110));
  L_Enums.AddObject('ststBorderCallout2',TObject(111));
  L_Enums.AddObject('ststBorderCallout3',TObject(112));
  L_Enums.AddObject('ststAccentBorderCallout1',TObject(113));
  L_Enums.AddObject('ststAccentBorderCallout2',TObject(114));
  L_Enums.AddObject('ststAccentBorderCallout3',TObject(115));
  L_Enums.AddObject('ststWedgeRectCallout',TObject(116));
  L_Enums.AddObject('ststWedgeRoundRectCallout',TObject(117));
  L_Enums.AddObject('ststWedgeEllipseCallout',TObject(118));
  L_Enums.AddObject('ststCloudCallout',TObject(119));
  L_Enums.AddObject('ststCloud',TObject(120));
  L_Enums.AddObject('ststRibbon',TObject(121));
  L_Enums.AddObject('ststRibbon2',TObject(122));
  L_Enums.AddObject('ststEllipseRibbon',TObject(123));
  L_Enums.AddObject('ststEllipseRibbon2',TObject(124));
  L_Enums.AddObject('ststLeftRightRibbon',TObject(125));
  L_Enums.AddObject('ststVerticalScroll',TObject(126));
  L_Enums.AddObject('ststHorizontalScroll',TObject(127));
  L_Enums.AddObject('ststWave',TObject(128));
  L_Enums.AddObject('ststDoubleWave',TObject(129));
  L_Enums.AddObject('ststPlus',TObject(130));
  L_Enums.AddObject('ststFlowChartProcess',TObject(131));
  L_Enums.AddObject('ststFlowChartDecision',TObject(132));
  L_Enums.AddObject('ststFlowChartInputOutput',TObject(133));
  L_Enums.AddObject('ststFlowChartPredefinedProcess',TObject(134));
  L_Enums.AddObject('ststFlowChartInternalStorage',TObject(135));
  L_Enums.AddObject('ststFlowChartDocument',TObject(136));
  L_Enums.AddObject('ststFlowChartMultidocument',TObject(137));
  L_Enums.AddObject('ststFlowChartTerminator',TObject(138));
  L_Enums.AddObject('ststFlowChartPreparation',TObject(139));
  L_Enums.AddObject('ststFlowChartManualInput',TObject(140));
  L_Enums.AddObject('ststFlowChartManualOperation',TObject(141));
  L_Enums.AddObject('ststFlowChartConnector',TObject(142));
  L_Enums.AddObject('ststFlowChartPunchedCard',TObject(143));
  L_Enums.AddObject('ststFlowChartPunchedTape',TObject(144));
  L_Enums.AddObject('ststFlowChartSummingJunction',TObject(145));
  L_Enums.AddObject('ststFlowChartOr',TObject(146));
  L_Enums.AddObject('ststFlowChartCollate',TObject(147));
  L_Enums.AddObject('ststFlowChartSort',TObject(148));
  L_Enums.AddObject('ststFlowChartExtract',TObject(149));
  L_Enums.AddObject('ststFlowChartMerge',TObject(150));
  L_Enums.AddObject('ststFlowChartOfflineStorage',TObject(151));
  L_Enums.AddObject('ststFlowChartOnlineStorage',TObject(152));
  L_Enums.AddObject('ststFlowChartMagneticTape',TObject(153));
  L_Enums.AddObject('ststFlowChartMagneticDisk',TObject(154));
  L_Enums.AddObject('ststFlowChartMagneticDrum',TObject(155));
  L_Enums.AddObject('ststFlowChartDisplay',TObject(156));
  L_Enums.AddObject('ststFlowChartDelay',TObject(157));
  L_Enums.AddObject('ststFlowChartAlternateProcess',TObject(158));
  L_Enums.AddObject('ststFlowChartOffpageConnector',TObject(159));
  L_Enums.AddObject('ststActionButtonBlank',TObject(160));
  L_Enums.AddObject('ststActionButtonHome',TObject(161));
  L_Enums.AddObject('ststActionButtonHelp',TObject(162));
  L_Enums.AddObject('ststActionButtonInformation',TObject(163));
  L_Enums.AddObject('ststActionButtonForwardNext',TObject(164));
  L_Enums.AddObject('ststActionButtonBackPrevious',TObject(165));
  L_Enums.AddObject('ststActionButtonEnd',TObject(166));
  L_Enums.AddObject('ststActionButtonBeginning',TObject(167));
  L_Enums.AddObject('ststActionButtonReturn',TObject(168));
  L_Enums.AddObject('ststActionButtonDocument',TObject(169));
  L_Enums.AddObject('ststActionButtonSound',TObject(170));
  L_Enums.AddObject('ststActionButtonMovie',TObject(171));
  L_Enums.AddObject('ststGear6',TObject(172));
  L_Enums.AddObject('ststGear9',TObject(173));
  L_Enums.AddObject('ststFunnel',TObject(174));
  L_Enums.AddObject('ststMathPlus',TObject(175));
  L_Enums.AddObject('ststMathMinus',TObject(176));
  L_Enums.AddObject('ststMathMultiply',TObject(177));
  L_Enums.AddObject('ststMathDivide',TObject(178));
  L_Enums.AddObject('ststMathEqual',TObject(179));
  L_Enums.AddObject('ststMathNotEqual',TObject(180));
  L_Enums.AddObject('ststCornerTabs',TObject(181));
  L_Enums.AddObject('ststSquareTabs',TObject(182));
  L_Enums.AddObject('ststPlaqueTabs',TObject(183));
  L_Enums.AddObject('ststChartX',TObject(184));
  L_Enums.AddObject('ststChartStar',TObject(185));
  L_Enums.AddObject('ststChartPlus',TObject(186));
  L_Enums.AddObject('stpfmNone',TObject(0));
  L_Enums.AddObject('stpfmNorm',TObject(1));
  L_Enums.AddObject('stpfmLighten',TObject(2));
  L_Enums.AddObject('stpfmLightenLess',TObject(3));
  L_Enums.AddObject('stpfmDarken',TObject(4));
  L_Enums.AddObject('stpfmDarkenLess',TObject(5));
  L_Enums.AddObject('sttstTextNoShape',TObject(0));
  L_Enums.AddObject('sttstTextPlain',TObject(1));
  L_Enums.AddObject('sttstTextStop',TObject(2));
  L_Enums.AddObject('sttstTextTriangle',TObject(3));
  L_Enums.AddObject('sttstTextTriangleInverted',TObject(4));
  L_Enums.AddObject('sttstTextChevron',TObject(5));
  L_Enums.AddObject('sttstTextChevronInverted',TObject(6));
  L_Enums.AddObject('sttstTextRingInside',TObject(7));
  L_Enums.AddObject('sttstTextRingOutside',TObject(8));
  L_Enums.AddObject('sttstTextArchUp',TObject(9));
  L_Enums.AddObject('sttstTextArchDown',TObject(10));
  L_Enums.AddObject('sttstTextCircle',TObject(11));
  L_Enums.AddObject('sttstTextButton',TObject(12));
  L_Enums.AddObject('sttstTextArchUpPour',TObject(13));
  L_Enums.AddObject('sttstTextArchDownPour',TObject(14));
  L_Enums.AddObject('sttstTextCirclePour',TObject(15));
  L_Enums.AddObject('sttstTextButtonPour',TObject(16));
  L_Enums.AddObject('sttstTextCurveUp',TObject(17));
  L_Enums.AddObject('sttstTextCurveDown',TObject(18));
  L_Enums.AddObject('sttstTextCanUp',TObject(19));
  L_Enums.AddObject('sttstTextCanDown',TObject(20));
  L_Enums.AddObject('sttstTextWave1',TObject(21));
  L_Enums.AddObject('sttstTextWave2',TObject(22));
  L_Enums.AddObject('sttstTextDoubleWave1',TObject(23));
  L_Enums.AddObject('sttstTextWave4',TObject(24));
  L_Enums.AddObject('sttstTextInflate',TObject(25));
  L_Enums.AddObject('sttstTextDeflate',TObject(26));
  L_Enums.AddObject('sttstTextInflateBottom',TObject(27));
  L_Enums.AddObject('sttstTextDeflateBottom',TObject(28));
  L_Enums.AddObject('sttstTextInflateTop',TObject(29));
  L_Enums.AddObject('sttstTextDeflateTop',TObject(30));
  L_Enums.AddObject('sttstTextDeflateInflate',TObject(31));
  L_Enums.AddObject('sttstTextDeflateInflateDeflate',TObject(32));
  L_Enums.AddObject('sttstTextFadeRight',TObject(33));
  L_Enums.AddObject('sttstTextFadeLeft',TObject(34));
  L_Enums.AddObject('sttstTextFadeUp',TObject(35));
  L_Enums.AddObject('sttstTextFadeDown',TObject(36));
  L_Enums.AddObject('sttstTextSlantUp',TObject(37));
  L_Enums.AddObject('sttstTextSlantDown',TObject(38));
  L_Enums.AddObject('sttstTextCascadeUp',TObject(39));
  L_Enums.AddObject('sttstTextCascadeDown',TObject(40));
  L_Enums.AddObject('stpctLegacyObliqueTopLeft',TObject(0));
  L_Enums.AddObject('stpctLegacyObliqueTop',TObject(1));
  L_Enums.AddObject('stpctLegacyObliqueTopRight',TObject(2));
  L_Enums.AddObject('stpctLegacyObliqueLeft',TObject(3));
  L_Enums.AddObject('stpctLegacyObliqueFront',TObject(4));
  L_Enums.AddObject('stpctLegacyObliqueRight',TObject(5));
  L_Enums.AddObject('stpctLegacyObliqueBottomLeft',TObject(6));
  L_Enums.AddObject('stpctLegacyObliqueBottom',TObject(7));
  L_Enums.AddObject('stpctLegacyObliqueBottomRight',TObject(8));
  L_Enums.AddObject('stpctLegacyPerspectiveTopLeft',TObject(9));
  L_Enums.AddObject('stpctLegacyPerspectiveTop',TObject(10));
  L_Enums.AddObject('stpctLegacyPerspectiveTopRight',TObject(11));
  L_Enums.AddObject('stpctLegacyPerspectiveLeft',TObject(12));
  L_Enums.AddObject('stpctLegacyPerspectiveFront',TObject(13));
  L_Enums.AddObject('stpctLegacyPerspectiveRight',TObject(14));
  L_Enums.AddObject('stpctLegacyPerspectiveBottomLeft',TObject(15));
  L_Enums.AddObject('stpctLegacyPerspectiveBottom',TObject(16));
  L_Enums.AddObject('stpctLegacyPerspectiveBottomRight',TObject(17));
  L_Enums.AddObject('stpctOrthographicFront',TObject(18));
  L_Enums.AddObject('stpctIsometricTopUp',TObject(19));
  L_Enums.AddObject('stpctIsometricTopDown',TObject(20));
  L_Enums.AddObject('stpctIsometricBottomUp',TObject(21));
  L_Enums.AddObject('stpctIsometricBottomDown',TObject(22));
  L_Enums.AddObject('stpctIsometricLeftUp',TObject(23));
  L_Enums.AddObject('stpctIsometricLeftDown',TObject(24));
  L_Enums.AddObject('stpctIsometricRightUp',TObject(25));
  L_Enums.AddObject('stpctIsometricRightDown',TObject(26));
  L_Enums.AddObject('stpctIsometricOffAxis1Left',TObject(27));
  L_Enums.AddObject('stpctIsometricOffAxis1Right',TObject(28));
  L_Enums.AddObject('stpctIsometricOffAxis1Top',TObject(29));
  L_Enums.AddObject('stpctIsometricOffAxis2Left',TObject(30));
  L_Enums.AddObject('stpctIsometricOffAxis2Right',TObject(31));
  L_Enums.AddObject('stpctIsometricOffAxis2Top',TObject(32));
  L_Enums.AddObject('stpctIsometricOffAxis3Left',TObject(33));
  L_Enums.AddObject('stpctIsometricOffAxis3Right',TObject(34));
  L_Enums.AddObject('stpctIsometricOffAxis3Bottom',TObject(35));
  L_Enums.AddObject('stpctIsometricOffAxis4Left',TObject(36));
  L_Enums.AddObject('stpctIsometricOffAxis4Right',TObject(37));
  L_Enums.AddObject('stpctIsometricOffAxis4Bottom',TObject(38));
  L_Enums.AddObject('stpctObliqueTopLeft',TObject(39));
  L_Enums.AddObject('stpctObliqueTop',TObject(40));
  L_Enums.AddObject('stpctObliqueTopRight',TObject(41));
  L_Enums.AddObject('stpctObliqueLeft',TObject(42));
  L_Enums.AddObject('stpctObliqueRight',TObject(43));
  L_Enums.AddObject('stpctObliqueBottomLeft',TObject(44));
  L_Enums.AddObject('stpctObliqueBottom',TObject(45));
  L_Enums.AddObject('stpctObliqueBottomRight',TObject(46));
  L_Enums.AddObject('stpctPerspectiveFront',TObject(47));
  L_Enums.AddObject('stpctPerspectiveLeft',TObject(48));
  L_Enums.AddObject('stpctPerspectiveRight',TObject(49));
  L_Enums.AddObject('stpctPerspectiveAbove',TObject(50));
  L_Enums.AddObject('stpctPerspectiveBelow',TObject(51));
  L_Enums.AddObject('stpctPerspectiveAboveLeftFacing',TObject(52));
  L_Enums.AddObject('stpctPerspectiveAboveRightFacing',TObject(53));
  L_Enums.AddObject('stpctPerspectiveContrastingLeftFacing',TObject(54));
  L_Enums.AddObject('stpctPerspectiveContrastingRightFacing',TObject(55));
  L_Enums.AddObject('stpctPerspectiveHeroicLeftFacing',TObject(56));
  L_Enums.AddObject('stpctPerspectiveHeroicRightFacing',TObject(57));
  L_Enums.AddObject('stpctPerspectiveHeroicExtremeLeftFacing',TObject(58));
  L_Enums.AddObject('stpctPerspectiveHeroicExtremeRightFacing',TObject(59));
  L_Enums.AddObject('stpctPerspectiveRelaxed',TObject(60));
  L_Enums.AddObject('stpctPerspectiveRelaxedModerately',TObject(61));
  L_Enums.AddObject('stlrdTl',TObject(0));
  L_Enums.AddObject('stlrdT',TObject(1));
  L_Enums.AddObject('stlrdTr',TObject(2));
  L_Enums.AddObject('stlrdL',TObject(3));
  L_Enums.AddObject('stlrdR',TObject(4));
  L_Enums.AddObject('stlrdBl',TObject(5));
  L_Enums.AddObject('stlrdB',TObject(6));
  L_Enums.AddObject('stlrdBr',TObject(7));
  L_Enums.AddObject('stlrtLegacyFlat1',TObject(0));
  L_Enums.AddObject('stlrtLegacyFlat2',TObject(1));
  L_Enums.AddObject('stlrtLegacyFlat3',TObject(2));
  L_Enums.AddObject('stlrtLegacyFlat4',TObject(3));
  L_Enums.AddObject('stlrtLegacyNormal1',TObject(4));
  L_Enums.AddObject('stlrtLegacyNormal2',TObject(5));
  L_Enums.AddObject('stlrtLegacyNormal3',TObject(6));
  L_Enums.AddObject('stlrtLegacyNormal4',TObject(7));
  L_Enums.AddObject('stlrtLegacyHarsh1',TObject(8));
  L_Enums.AddObject('stlrtLegacyHarsh2',TObject(9));
  L_Enums.AddObject('stlrtLegacyHarsh3',TObject(10));
  L_Enums.AddObject('stlrtLegacyHarsh4',TObject(11));
  L_Enums.AddObject('stlrtThreePt',TObject(12));
  L_Enums.AddObject('stlrtBalanced',TObject(13));
  L_Enums.AddObject('stlrtSoft',TObject(14));
  L_Enums.AddObject('stlrtHarsh',TObject(15));
  L_Enums.AddObject('stlrtFlood',TObject(16));
  L_Enums.AddObject('stlrtContrasting',TObject(17));
  L_Enums.AddObject('stlrtMorning',TObject(18));
  L_Enums.AddObject('stlrtSunrise',TObject(19));
  L_Enums.AddObject('stlrtSunset',TObject(20));
  L_Enums.AddObject('stlrtChilly',TObject(21));
  L_Enums.AddObject('stlrtFreezing',TObject(22));
  L_Enums.AddObject('stlrtFlat',TObject(23));
  L_Enums.AddObject('stlrtTwoPt',TObject(24));
  L_Enums.AddObject('stlrtGlow',TObject(25));
  L_Enums.AddObject('stlrtBrightRoom',TObject(26));
  L_Enums.AddObject('stpaCtr',TObject(0));
  L_Enums.AddObject('stpaIn',TObject(1));
  L_Enums.AddObject('stletNone',TObject(0));
  L_Enums.AddObject('stletTriangle',TObject(1));
  L_Enums.AddObject('stletStealth',TObject(2));
  L_Enums.AddObject('stletDiamond',TObject(3));
  L_Enums.AddObject('stletOval',TObject(4));
  L_Enums.AddObject('stletArrow',TObject(5));
  L_Enums.AddObject('stlewSm',TObject(0));
  L_Enums.AddObject('stlewMed',TObject(1));
  L_Enums.AddObject('stlewLg',TObject(2));
  L_Enums.AddObject('stlcRnd',TObject(0));
  L_Enums.AddObject('stlcSq',TObject(1));
  L_Enums.AddObject('stlcFlat',TObject(2));
  L_Enums.AddObject('stpldvSolid',TObject(0));
  L_Enums.AddObject('stpldvDot',TObject(1));
  L_Enums.AddObject('stpldvDash',TObject(2));
  L_Enums.AddObject('stpldvLgDash',TObject(3));
  L_Enums.AddObject('stpldvDashDot',TObject(4));
  L_Enums.AddObject('stpldvLgDashDot',TObject(5));
  L_Enums.AddObject('stpldvLgDashDotDot',TObject(6));
  L_Enums.AddObject('stpldvSysDash',TObject(7));
  L_Enums.AddObject('stpldvSysDot',TObject(8));
  L_Enums.AddObject('stpldvSysDashDot',TObject(9));
  L_Enums.AddObject('stpldvSysDashDotDot',TObject(10));
  L_Enums.AddObject('stlelSm',TObject(0));
  L_Enums.AddObject('stlelMed',TObject(1));
  L_Enums.AddObject('stlelLg',TObject(2));
  L_Enums.AddObject('stclSng',TObject(0));
  L_Enums.AddObject('stclDbl',TObject(1));
  L_Enums.AddObject('stclThickThin',TObject(2));
  L_Enums.AddObject('stclThinThick',TObject(3));
  L_Enums.AddObject('stclTri',TObject(4));
  L_Enums.AddObject('stbptRelaxedInset',TObject(0));
  L_Enums.AddObject('stbptCircle',TObject(1));
  L_Enums.AddObject('stbptSlope',TObject(2));
  L_Enums.AddObject('stbptCross',TObject(3));
  L_Enums.AddObject('stbptAngle',TObject(4));
  L_Enums.AddObject('stbptSoftRound',TObject(5));
  L_Enums.AddObject('stbptConvex',TObject(6));
  L_Enums.AddObject('stbptCoolSlant',TObject(7));
  L_Enums.AddObject('stbptDivot',TObject(8));
  L_Enums.AddObject('stbptRiblet',TObject(9));
  L_Enums.AddObject('stbptHardEdge',TObject(10));
  L_Enums.AddObject('stbptArtDeco',TObject(11));
  L_Enums.AddObject('stpmtLegacyMatte',TObject(0));
  L_Enums.AddObject('stpmtLegacyPlastic',TObject(1));
  L_Enums.AddObject('stpmtLegacyMetal',TObject(2));
  L_Enums.AddObject('stpmtLegacyWireframe',TObject(3));
  L_Enums.AddObject('stpmtMatte',TObject(4));
  L_Enums.AddObject('stpmtPlastic',TObject(5));
  L_Enums.AddObject('stpmtMetal',TObject(6));
  L_Enums.AddObject('stpmtWarmMatte',TObject(7));
  L_Enums.AddObject('stpmtTranslucentPowder',TObject(8));
  L_Enums.AddObject('stpmtPowder',TObject(9));
  L_Enums.AddObject('stpmtDkEdge',TObject(10));
  L_Enums.AddObject('stpmtSoftEdge',TObject(11));
  L_Enums.AddObject('stpmtClear',TObject(12));
  L_Enums.AddObject('stpmtFlat',TObject(13));
  L_Enums.AddObject('stpmtSoftmetal',TObject(14));
  L_Enums.AddObject('sttctNone',TObject(0));
  L_Enums.AddObject('sttctSmall',TObject(1));
  L_Enums.AddObject('sttctAll',TObject(2));
  L_Enums.AddObject('sttstNoStrike',TObject(0));
  L_Enums.AddObject('sttstSngStrike',TObject(1));
  L_Enums.AddObject('sttstDblStrike',TObject(2));
  L_Enums.AddObject('sttutNone',TObject(0));
  L_Enums.AddObject('sttutWords',TObject(1));
  L_Enums.AddObject('sttutSng',TObject(2));
  L_Enums.AddObject('sttutDbl',TObject(3));
  L_Enums.AddObject('sttutHeavy',TObject(4));
  L_Enums.AddObject('sttutDotted',TObject(5));
  L_Enums.AddObject('sttutDottedHeavy',TObject(6));
  L_Enums.AddObject('sttutDash',TObject(7));
  L_Enums.AddObject('sttutDashHeavy',TObject(8));
  L_Enums.AddObject('sttutDashLong',TObject(9));
  L_Enums.AddObject('sttutDashLongHeavy',TObject(10));
  L_Enums.AddObject('sttutDotDash',TObject(11));
  L_Enums.AddObject('sttutDotDashHeavy',TObject(12));
  L_Enums.AddObject('sttutDotDotDash',TObject(13));
  L_Enums.AddObject('sttutDotDotDashHeavy',TObject(14));
  L_Enums.AddObject('sttutWavy',TObject(15));
  L_Enums.AddObject('sttutWavyHeavy',TObject(16));
  L_Enums.AddObject('sttutWavyDbl',TObject(17));
  L_Enums.AddObject('sttasAlphaLcParenBoth',TObject(0));
  L_Enums.AddObject('sttasAlphaUcParenBoth',TObject(1));
  L_Enums.AddObject('sttasAlphaLcParenR',TObject(2));
  L_Enums.AddObject('sttasAlphaUcParenR',TObject(3));
  L_Enums.AddObject('sttasAlphaLcPeriod',TObject(4));
  L_Enums.AddObject('sttasAlphaUcPeriod',TObject(5));
  L_Enums.AddObject('sttasArabicParenBoth',TObject(6));
  L_Enums.AddObject('sttasArabicParenR',TObject(7));
  L_Enums.AddObject('sttasArabicPeriod',TObject(8));
  L_Enums.AddObject('sttasArabicPlain',TObject(9));
  L_Enums.AddObject('sttasRomanLcParenBoth',TObject(10));
  L_Enums.AddObject('sttasRomanUcParenBoth',TObject(11));
  L_Enums.AddObject('sttasRomanLcParenR',TObject(12));
  L_Enums.AddObject('sttasRomanUcParenR',TObject(13));
  L_Enums.AddObject('sttasRomanLcPeriod',TObject(14));
  L_Enums.AddObject('sttasRomanUcPeriod',TObject(15));
  L_Enums.AddObject('sttasCircleNumDbPlain',TObject(16));
  L_Enums.AddObject('sttasCircleNumWdBlackPlain',TObject(17));
  L_Enums.AddObject('sttasCircleNumWdWhitePlain',TObject(18));
  L_Enums.AddObject('sttasArabicDbPeriod',TObject(19));
  L_Enums.AddObject('sttasArabicDbPlain',TObject(20));
  L_Enums.AddObject('sttasEa1ChsPeriod',TObject(21));
  L_Enums.AddObject('sttasEa1ChsPlain',TObject(22));
  L_Enums.AddObject('sttasEa1ChtPeriod',TObject(23));
  L_Enums.AddObject('sttasEa1ChtPlain',TObject(24));
  L_Enums.AddObject('sttasEa1JpnChsDbPeriod',TObject(25));
  L_Enums.AddObject('sttasEa1JpnKorPlain',TObject(26));
  L_Enums.AddObject('sttasEa1JpnKorPeriod',TObject(27));
  L_Enums.AddObject('sttasArabic1Minus',TObject(28));
  L_Enums.AddObject('sttasArabic2Minus',TObject(29));
  L_Enums.AddObject('sttasHebrew2Minus',TObject(30));
  L_Enums.AddObject('sttasThaiAlphaPeriod',TObject(31));
  L_Enums.AddObject('sttasThaiAlphaParenR',TObject(32));
  L_Enums.AddObject('sttasThaiAlphaParenBoth',TObject(33));
  L_Enums.AddObject('sttasThaiNumPeriod',TObject(34));
  L_Enums.AddObject('sttasThaiNumParenR',TObject(35));
  L_Enums.AddObject('sttasThaiNumParenBoth',TObject(36));
  L_Enums.AddObject('sttasHindiAlphaPeriod',TObject(37));
  L_Enums.AddObject('sttasHindiNumPeriod',TObject(38));
  L_Enums.AddObject('sttasHindiNumParenR',TObject(39));
  L_Enums.AddObject('sttasHindiAlpha1Period',TObject(40));
  L_Enums.AddObject('stcsiDk1',TObject(0));
  L_Enums.AddObject('stcsiLt1',TObject(1));
  L_Enums.AddObject('stcsiDk2',TObject(2));
  L_Enums.AddObject('stcsiLt2',TObject(3));
  L_Enums.AddObject('stcsiAccent1',TObject(4));
  L_Enums.AddObject('stcsiAccent2',TObject(5));
  L_Enums.AddObject('stcsiAccent3',TObject(6));
  L_Enums.AddObject('stcsiAccent4',TObject(7));
  L_Enums.AddObject('stcsiAccent5',TObject(8));
  L_Enums.AddObject('stcsiAccent6',TObject(9));
  L_Enums.AddObject('stcsiHlink',TObject(10));
  L_Enums.AddObject('stcsiFolHlink',TObject(11));
  L_Enums.AddObject('stfciMajor',TObject(0));
  L_Enums.AddObject('stfciMinor',TObject(1));
  L_Enums.AddObject('stfciNone',TObject(2));
  L_Enums.AddObject('sttatL',TObject(0));
  L_Enums.AddObject('sttatCtr',TObject(1));
  L_Enums.AddObject('sttatR',TObject(2));
  L_Enums.AddObject('sttatJust',TObject(3));
  L_Enums.AddObject('sttatJustLow',TObject(4));
  L_Enums.AddObject('sttatDist',TObject(5));
  L_Enums.AddObject('sttatThaiDist',TObject(6));
  L_Enums.AddObject('sttfatAuto',TObject(0));
  L_Enums.AddObject('sttfatT',TObject(1));
  L_Enums.AddObject('sttfatCtr',TObject(2));
  L_Enums.AddObject('sttfatBase',TObject(3));
  L_Enums.AddObject('sttfatB',TObject(4));
  L_Enums.AddObject('stttatL',TObject(0));
  L_Enums.AddObject('stttatCtr',TObject(1));
  L_Enums.AddObject('stttatR',TObject(2));
  L_Enums.AddObject('stttatDec',TObject(3));
  L_Enums.AddObject('sttatT',TObject(0));
  L_Enums.AddObject('sttatCtr1',TObject(1));
  L_Enums.AddObject('sttatB',TObject(2));
  L_Enums.AddObject('sttatJust1',TObject(3));
  L_Enums.AddObject('sttatDist1',TObject(4));
  L_Enums.AddObject('sttvtHorz',TObject(0));
  L_Enums.AddObject('sttvtVert',TObject(1));
  L_Enums.AddObject('sttvtVert270',TObject(2));
  L_Enums.AddObject('sttvtWordArtVert',TObject(3));
  L_Enums.AddObject('sttvtEaVert',TObject(4));
  L_Enums.AddObject('sttvtMongolianVert',TObject(5));
  L_Enums.AddObject('sttvtWordArtVertRtl',TObject(6));
  L_Enums.AddObject('sttwtNone',TObject(0));
  L_Enums.AddObject('sttwtSquare',TObject(1));
  L_Enums.AddObject('stthotOverflow',TObject(0));
  L_Enums.AddObject('stthotClip',TObject(1));
  L_Enums.AddObject('sttvotOverflow',TObject(0));
  L_Enums.AddObject('sttvotEllipsis',TObject(1));
  L_Enums.AddObject('sttvotClip',TObject(2));
  L_Enums.AddObject('strsStandard',TObject(0));
  L_Enums.AddObject('strsMarker',TObject(1));
  L_Enums.AddObject('strsFilled',TObject(2));
  L_Enums.AddObject('stbdBar',TObject(0));
  L_Enums.AddObject('stbdCol',TObject(1));
  L_Enums.AddObject('ststAuto',TObject(0));
  L_Enums.AddObject('ststCust',TObject(1));
  L_Enums.AddObject('ststPercent',TObject(2));
  L_Enums.AddObject('ststPos',TObject(3));
  L_Enums.AddObject('ststVal',TObject(4));
  L_Enums.AddObject('sttmCross',TObject(0));
  L_Enums.AddObject('sttmIn',TObject(1));
  L_Enums.AddObject('sttmNone',TObject(2));
  L_Enums.AddObject('sttmOut',TObject(3));
  L_Enums.AddObject('sttuDays',TObject(0));
  L_Enums.AddObject('sttuMonths',TObject(1));
  L_Enums.AddObject('sttuYears',TObject(2));
  L_Enums.AddObject('stgPercentStacked',TObject(0));
  L_Enums.AddObject('stgStandard',TObject(1));
  L_Enums.AddObject('stgStacked',TObject(2));
  L_Enums.AddObject('stoMaxMin',TObject(0));
  L_Enums.AddObject('stoMinMax',TObject(1));
  L_Enums.AddObject('stlaCtr',TObject(0));
  L_Enums.AddObject('stlaL',TObject(1));
  L_Enums.AddObject('stlaR',TObject(2));
  L_Enums.AddObject('stmsCircle',TObject(0));
  L_Enums.AddObject('stmsDash',TObject(1));
  L_Enums.AddObject('stmsDiamond',TObject(2));
  L_Enums.AddObject('stmsDot',TObject(3));
  L_Enums.AddObject('stmsNone',TObject(4));
  L_Enums.AddObject('stmsPicture',TObject(5));
  L_Enums.AddObject('stmsPlus',TObject(6));
  L_Enums.AddObject('stmsSquare',TObject(7));
  L_Enums.AddObject('stmsStar',TObject(8));
  L_Enums.AddObject('stmsTriangle',TObject(9));
  L_Enums.AddObject('stmsX',TObject(10));
  L_Enums.AddObject('stcAutoZero',TObject(0));
  L_Enums.AddObject('stcMax',TObject(1));
  L_Enums.AddObject('stcMin',TObject(2));
  L_Enums.AddObject('stpfStretch',TObject(0));
  L_Enums.AddObject('stpfStack',TObject(1));
  L_Enums.AddObject('stpfStackScale',TObject(2));
  L_Enums.AddObject('stebtBoth',TObject(0));
  L_Enums.AddObject('stebtMinus',TObject(1));
  L_Enums.AddObject('stebtPlus',TObject(2));
  L_Enums.AddObject('sttlpHigh',TObject(0));
  L_Enums.AddObject('sttlpLow',TObject(1));
  L_Enums.AddObject('sttlpNextTo',TObject(2));
  L_Enums.AddObject('sttlpNone',TObject(3));
  L_Enums.AddObject('stapB',TObject(0));
  L_Enums.AddObject('stapL',TObject(1));
  L_Enums.AddObject('stapR',TObject(2));
  L_Enums.AddObject('stapT',TObject(3));
  L_Enums.AddObject('stsCone',TObject(0));
  L_Enums.AddObject('stsConeToMax',TObject(1));
  L_Enums.AddObject('stsBox',TObject(2));
  L_Enums.AddObject('stsCylinder',TObject(3));
  L_Enums.AddObject('stsPyramid',TObject(4));
  L_Enums.AddObject('stsPyramidToMax',TObject(5));
  L_Enums.AddObject('stcbBetween',TObject(0));
  L_Enums.AddObject('stcbMidCat',TObject(1));
  L_Enums.AddObject('stdbaSpan',TObject(0));
  L_Enums.AddObject('stdbaGap',TObject(1));
  L_Enums.AddObject('stdbaZero',TObject(2));
  L_Enums.AddObject('stevtCust',TObject(0));
  L_Enums.AddObject('stevtFixedVal',TObject(1));
  L_Enums.AddObject('stevtPercentage',TObject(2));
  L_Enums.AddObject('stevtStdDev',TObject(3));
  L_Enums.AddObject('stevtStdErr',TObject(4));
  L_Enums.AddObject('stlpB',TObject(0));
  L_Enums.AddObject('stlpTr',TObject(1));
  L_Enums.AddObject('stlpL',TObject(2));
  L_Enums.AddObject('stlpR',TObject(3));
  L_Enums.AddObject('stlpT',TObject(4));
  L_Enums.AddObject('stedX',TObject(0));
  L_Enums.AddObject('stedY',TObject(1));
  L_Enums.AddObject('stlmEdge',TObject(0));
  L_Enums.AddObject('stlmFactor',TObject(1));
  L_Enums.AddObject('stbgPercentStacked',TObject(0));
  L_Enums.AddObject('stbgClustered',TObject(1));
  L_Enums.AddObject('stbgStandard',TObject(2));
  L_Enums.AddObject('stbgStacked',TObject(3));
  L_Enums.AddObject('stdlpBestFit',TObject(0));
  L_Enums.AddObject('stdlpB',TObject(1));
  L_Enums.AddObject('stdlpCtr',TObject(2));
  L_Enums.AddObject('stdlpInBase',TObject(3));
  L_Enums.AddObject('stdlpInEnd',TObject(4));
  L_Enums.AddObject('stdlpL',TObject(5));
  L_Enums.AddObject('stdlpOutEnd',TObject(6));
  L_Enums.AddObject('stdlpR',TObject(7));
  L_Enums.AddObject('stdlpT',TObject(8));
  L_Enums.AddObject('stsrArea',TObject(0));
  L_Enums.AddObject('stsrW',TObject(1));
  L_Enums.AddObject('stpsoDefault',TObject(0));
  L_Enums.AddObject('stpsoPortrait',TObject(1));
  L_Enums.AddObject('stpsoLandscape',TObject(2));
  L_Enums.AddObject('stoptPie',TObject(0));
  L_Enums.AddObject('stoptBar',TObject(1));
  L_Enums.AddObject('stssNone',TObject(0));
  L_Enums.AddObject('stssLine',TObject(1));
  L_Enums.AddObject('stssLineMarker',TObject(2));
  L_Enums.AddObject('stssMarker',TObject(3));
  L_Enums.AddObject('stssSmooth',TObject(4));
  L_Enums.AddObject('stssSmoothMarker',TObject(5));
  L_Enums.AddObject('stltInner',TObject(0));
  L_Enums.AddObject('stltOuter',TObject(1));
  L_Enums.AddObject('stbiuHundreds',TObject(0));
  L_Enums.AddObject('stbiuThousands',TObject(1));
  L_Enums.AddObject('stbiuTenThousands',TObject(2));
  L_Enums.AddObject('stbiuHundredThousands',TObject(3));
  L_Enums.AddObject('stbiuMillions',TObject(4));
  L_Enums.AddObject('stbiuTenMillions',TObject(5));
  L_Enums.AddObject('stbiuHundredMillions',TObject(6));
  L_Enums.AddObject('stbiuBillions',TObject(7));
  L_Enums.AddObject('stbiuTrillions',TObject(8));
  L_Enums.AddObject('stttExp',TObject(0));
  L_Enums.AddObject('stttLinear',TObject(1));
  L_Enums.AddObject('stttLog',TObject(2));
  L_Enums.AddObject('stttMovingAvg',TObject(3));
  L_Enums.AddObject('stttPoly',TObject(4));
  L_Enums.AddObject('stttPower',TObject(5));

  L_Enums.AddObject('steatwoCell',TObject(0));
  L_Enums.AddObject('steaoneCell',TObject(1));
  L_Enums.AddObject('steaabsolute',TObject(2));
end;

class function  TXPGBase.StrToEnum(AValue: AxUCString): integer;
var
  i: integer;
begin
  i := L_Enums.IndexOf(AValue);
  if i >= 0 then
    Result := Integer(L_Enums.Objects[i])
  else
    raise XLSRWException.CreateFmt('Can not find enum "%s"',[AValue]);
end;

class function  TXPGBase.StrToEnumDef(AValue: AxUCString; ADefault: integer): integer;
var
  i: integer;
begin
  i := L_Enums.IndexOf(AValue);
  if i >= 0 then
    Result := Integer(L_Enums.Objects[i])
  else
    Result := ADefault;
end;

class function  TXPGBase.TryStrToEnum(AValue: AxUCString; AText: AxUCString; AEnumNames: array of AxUCString; APtrInt: PInteger): boolean;
var
  i: integer;
begin
  i := L_Enums.IndexOf(AValue);
  if i >= 0 then
  begin
    i := Integer(L_Enums.Objects[i]);
    Result := (i <= High(AEnumNames)) and (AText = AEnumNames[i]);
    if Result then
      APtrInt^ := i;
  end
  else
    Result := False;
end;

function  TXPGBase.Available: boolean;
begin
  Result := xaRead in FAssigneds;
end;

{ TXPGAnyElement }

procedure TXPGAnyElement.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  FAttributes.Assign(AAttributes.FAttributes,AAttributes.Count);
end;

constructor TXPGAnyElement.Create;
begin
  FAttributes := TXpgXMLAttributeList.Create;
end;

destructor TXPGAnyElement.Destroy;
begin
  FAttributes.Free;
end;

{ TXPGAnyElements }

procedure TXPGAnyElements.Assign(AItem: TXPGAnyElements);
var
  i: integer;
begin
  for i := 0 to AItem.Count - 1 do
    Add(AItem[i].FElementName,AItem[i].FContent);
end;

function  TXPGAnyElements.GetItems(Index: integer): TXPGAnyElement;
begin
  Result := TXPGAnyElement(inherited Items[Index]);
end;

function  TXPGAnyElements.Add(AElementName: AxUCString; AContent: AxUCString): TXPGAnyElement;
begin
  Result := TXPGAnyElement.Create;
  Result.FElementName := AElementName;
  Result.FContent := AContent;
  inherited Add(Result);
end;

procedure TXPGAnyElements.Write(AWriter: TXpgWriteXML);
var
  i: integer;
  j: integer;
  Elem: TXPGAnyElement;
begin
  for i := 0 to Count - 1 do 
  begin
    Elem := GetItems(i);
    for j := 0 to Elem.Attributes.Count - 1 do 
      AWriter.AddAttribute(Elem.Attributes.Attributes[j],Elem.Attributes.Values[j]);
    AWriter.Text := Elem.Content;
    AWriter.SimpleTag(Elem.ElementName);
  end
end;

{ TXPGBaseObjectList }

function  TXPGBaseObjectList.GetItems(Index: integer): TXPGBase;
begin
  Result := TXPGBase(inherited Items[Index]);
end;

constructor TXPGBaseObjectList.Create(AOwner: TXPGDocBase);
begin
  inherited Create;
  FOwner := AOwner;
end;

{ TXPGReader }

constructor TXPGReader.Create(AErrors: TXpgPErrors; ARoot: TXPGBase);
begin
  inherited Create;
  FErrors := AErrors;
  FErrors.NoDuplicates := True;
  FCurrent := ARoot;
  FStack := TObjectStack.Create;
  FAlternateContent := TXPSAlternateContent.Create;
end;

destructor TXPGReader.Destroy;
begin
  FStack.Free;
  FAlternateContent.Free;
  inherited Destroy;
end;

procedure TXPGReader.BeginTag;
begin
  FStack.Push(FCurrent);
  //                             "Alte"
  if (PLongword(FTag.pStart)^ = $65746C41) and (Tag = 'AlternateContent') then begin
    FAlternateContent.BeginHandle(FCurrent);
    FCurrent := FAlternateContent;
  end;
  FCurrent := FCurrent.HandleElement(Self);
  if Attributes.Count > 0 then
    FCurrent.AssignAttributes(Attributes);
end;

procedure TXPGReader.EndTag;
begin
  FCurrent := TXPGBase(FStack.Pop);
  FCurrent.AfterTag;
end;

{ TEG_ColorTransform }

function  TEG_ColorTransform.CheckAssigned: integer;
begin
  Result := FItems.Count;
  if Result > 0 then
    FAssigneds := [xaElements]
  else
    FAssigneds := [];
end;

procedure TEG_ColorTransform.Write(AWriter: TXpgWriteXML);
var
  i: integer;
begin
  for i := 0 to FItems.Count - 1 do begin
    AWriter.AddAttribute('val',IntToStr(Integer(FItems.Objects[i])));
    AWriter.SimpleTag('a:' + FItems[i]);
  end;
end;

constructor TEG_ColorTransform.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 0;
  FItems := TStringList.Create;
end;

destructor TEG_ColorTransform.Destroy;
begin
  FItems.Free;
end;

function TEG_ColorTransform.Find(const AName: AxUCString): integer;
begin
  for Result := 0 to FItems.Count - 1 do begin
    if FItems[Result] = AName then
      Exit;
  end;
  Result := -1;
end;

function TEG_ColorTransform.GetAsDouble(const AName: AxUCString): double;
var
  i: integer;
  Name: AxUCString;
  Value: integer;
begin
  i := Find(AName);
  if i >= 0 then begin
    GetNameValue(i,Name,Value);
    Result := Value;
  end
  else
    Result := NaN;
end;

function TEG_ColorTransform.GetAsInteger(const AName: AxUCString): integer;
var
  i: integer;
  Name: AxUCString;
begin
  i := Find(AName);
  if i >= 0 then
    GetNameValue(i,Name,Result)
  else
    Result := 0;
end;

procedure TEG_ColorTransform.GetNameValue(const AIndex: integer; out AName: AxUCString; out AValue: integer);
begin
  AName := FItems[AIndex];
  AValue := Integer(FItems.Objects[AIndex]);
end;

function TEG_ColorTransform.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  Value: integer;
begin
  if AReader.Attributes.Count <= 0 then begin
    Result := Self;
    Exit;
  end;

  Value := StrToIntDef(AReader.Attributes.Values[0],0);

  AReader.Attributes.Clear;

  case AReader.QNameHashA of
    $0000025A: ReadAddValue('tint',Value);
    $000002A0: ReadAddValue('shade',Value);
    $0000024A: ReadAddValue('comp',Value);
    $000001E8: ReadAddValue('inv',Value);
    $0000024E: ReadAddValue('gray',Value);
    $000002A1: ReadAddValue('alpha',Value);
    $000003BC: ReadAddValue('alphaOff',Value);
    $000003C1: ReadAddValue('alphaMod',Value);
    $000001DD: ReadAddValue('hue',Value);
    $000002F8: ReadAddValue('hueOff',Value);
    $000002FD: ReadAddValue('hueMod',Value);
    $000001E3: ReadAddValue('sat',Value);
    $000002FE: ReadAddValue('satOff',Value);
    $00000303: ReadAddValue('satMod',Value);
    $000001E9: ReadAddValue('lum',Value);
    $00000304: ReadAddValue('lumOff',Value);
    $00000309: ReadAddValue('lumMod',Value);
    $000001D6: ReadAddValue('red',Value);
    $000002F1: ReadAddValue('redOff',Value);
    $000002F6: ReadAddValue('redMod',Value);
    $000002AC: ReadAddValue('green',Value);
    $000003C7: ReadAddValue('greenOff',Value);
    $000003CC: ReadAddValue('greenMod',Value);
    $00000243: ReadAddValue('blue',Value);
    $0000035E: ReadAddValue('blueOff',Value);
    $00000363: ReadAddValue('blueMod',Value);
    $0000029E: ReadAddValue('gamma',Value);
    $000003CB: ReadAddValue('invGamma',Value);
    else
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  Result := Self;
end;

function TEG_ColorTransform.HashA(const AStr: AxUCString): integer;
var
  i: integer;
begin
  Result := 0;
  for i := 1 to Length(AStr) do
    Inc(Result,Word(AStr[i]));
end;

procedure TEG_ColorTransform.ReadAddValue(const AName: AxUCString; AValue: integer);
begin
  FItems.AddObject(AName,TObject(AValue));
end;

procedure TEG_ColorTransform.SetAsDouble(const AName: AxUCString; const Value: double);
var
  i: integer;
begin
  i := Find(AName);
  if i >= 0 then
    FItems.Objects[i] := TObject(Round(Value))
  else
    FItems.AddObject(AName,TObject(Round(Value)));
end;

procedure TEG_ColorTransform.SetAsInteger(const AName: AxUCString; const Value: integer);
var
  i: integer;
begin
  i := Find(AName);
  if i >= 0 then
    FItems.Objects[i] := TObject(IntToStr(Value))
  else
    FItems.AddObject(AName,TObject(Value));
end;

procedure TEG_ColorTransform.Clear;
begin
  FAssigneds := [];
  FItems.Clear;
end;

// https://social.msdn.microsoft.com/Forums/en-US/eb151eb0-e539-4f12-a453-1ff46dc28578/satmod-satoff-documentation?forum=os_openXML-ecma
function TEG_ColorTransform.Apply(ARGB: longword): longword;
var
  i    : integer;
  H,L,S: double;
  V    : double;
begin
  RGBToHLS(ARGB,H,L,S);

  for i := 0 to FItems.Count - 1 do begin
    V := Integer(FItems.Objects[i]) / 100000;

    case HashA('a:' + FItems[i]) of
      $0000025A: ; // tint
      $000002A0: ; // shade
      $0000024A: ; // comp
      $000001E8: ; // inv
      $0000024E: ; // gray
      $000002A1: ; // alpha
      $000003BC: ; // alphaOff
      $000003C1: ; // alphaMod

      $000001DD: H := V; // hue
      $000002F8: H := H + V; // hueOff
      $000002FD: H := H * V; // hueMod

      $000001E3: S := V; // sat
      $000002FE: S := S + V; // satOff
      $00000303: S := S * V; // satMod

      $000001E9: L := V; // lum
      $00000304: L := L + V; // lumOff
      $00000309: L := L * V; // lumMod

      $000001D6: ; // red
      $000002F1: ; // redOff
      $000002F6: ; // redMod
      $000002AC: ; // green
      $000003C7: ; // greenOff
      $000003CC: ; // greenMod
      $00000243: ; // blue
      $0000035E: ; // blueOff
      $00000363: ; // blueMod
      $0000029E: ; // gamma
      $000003CB: ; // invGamma
    end;
  end;

  Result := HLSToRGB(H,L,S);
end;

procedure TEG_ColorTransform.Assign(AItem: TEG_ColorTransform);
begin
  FItems.Assign(AItem.FItems);
end;

procedure TEG_ColorTransform.CopyTo(AItem: TEG_ColorTransform);
begin
end;

{ TCT_ScRgbColor }

function  TCT_ScRgbColor.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [xaAttributes];
  Inc(ElemsAssigned,FA_EG_ColorTransform.CheckAssigned);
  Result := 1;
  if ElemsAssigned > 0 then
    FAssigneds := FAssigneds + [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_ScRgbColor.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := FA_EG_ColorTransform.HandleElement(AReader);
  if Result = Nil then
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then
    Result.Assigneds := [xaRead];
end;

procedure TCT_ScRgbColor.SetRGB(const Value: longword);
begin
  FR := ((Value and $00FF0000) shr 16) * 100000;
  FG := ((Value and $0000FF00) shr 8) * 100000;
  FB := (Value and $000000FF) * 100000;
end;

procedure TCT_ScRgbColor.Write(AWriter: TXpgWriteXML);
begin
  FA_EG_ColorTransform.Write(AWriter);
end;

procedure TCT_ScRgbColor.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('r',XmlIntToStr(FR));
  AWriter.AddAttribute('g',XmlIntToStr(FG));
  AWriter.AddAttribute('b',XmlIntToStr(FB));
end;

procedure TCT_ScRgbColor.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $00000072: FR := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000067: FG := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000062: FB := XmlStrToIntDef(AAttributes.Values[i],0);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_ScRgbColor.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 3;
  FA_EG_ColorTransform := TEG_ColorTransform.Create(FOwner);
  FR := 2147483632;
  FG := 2147483632;
  FB := 2147483632;
end;

destructor TCT_ScRgbColor.Destroy;
begin
  FA_EG_ColorTransform.Free;
end;

function TCT_ScRgbColor.GetRGB: longword;
begin
  Result := (Round((FR / 100000) * $FF) shl 16) + (Round((FG / 100000) * $FF) shl 8) + Round((FB / 100000) * $FF)
end;

procedure TCT_ScRgbColor.Clear;
begin
  FAssigneds := [];
  FA_EG_ColorTransform.Clear;
  FR := 2147483632;
  FG := 2147483632;
  FB := 2147483632;
end;

procedure TCT_ScRgbColor.Assign(AItem: TCT_ScRgbColor);
begin
  FR := AItem.FR;
  FG := AItem.FG;
  FB := AItem.FB;
  FA_EG_ColorTransform.Assign(AItem.FA_EG_ColorTransform);
end;

procedure TCT_ScRgbColor.CopyTo(AItem: TCT_ScRgbColor);
begin
end;

{ TCT_SRgbColor }

function  TCT_SRgbColor.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FVal <> -16 then begin
    FAssigneds := [xaAttributes];
    Inc(ElemsAssigned,FA_EG_ColorTransform.CheckAssigned);
    Result := 1;
  end
  else
    Result := 0;
  if ElemsAssigned > 0 then
    FAssigneds := FAssigneds + [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_SRgbColor.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := FA_EG_ColorTransform.HandleElement(AReader);
  if Result = Nil then 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_SRgbColor.Write(AWriter: TXpgWriteXML);
begin
  FA_EG_ColorTransform.Write(AWriter);
end;

procedure TCT_SRgbColor.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('val',XmlIntToHexStr(FVal,6));
end;

procedure TCT_SRgbColor.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'val' then 
    FVal := XmlStrToIntDef('$' + AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_SRgbColor.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 1;
  FA_EG_ColorTransform := TEG_ColorTransform.Create(FOwner);
  FVal := -16;
end;

destructor TCT_SRgbColor.Destroy;
begin
  FA_EG_ColorTransform.Free;
end;

procedure TCT_SRgbColor.Clear;
begin
  FAssigneds := [];
  FA_EG_ColorTransform.Clear;
  FVal := -16;
end;

procedure TCT_SRgbColor.Assign(AItem: TCT_SRgbColor);
begin
  FVal := AItem.FVal;
  FA_EG_ColorTransform.Assign(AItem.FA_EG_ColorTransform);
end;

procedure TCT_SRgbColor.CopyTo(AItem: TCT_SRgbColor);
begin
end;

{ TCT_HslColor }

function  TCT_HslColor.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [xaAttributes];
  Inc(ElemsAssigned,FA_EG_ColorTransform.CheckAssigned);
  Result := 1;
  if ElemsAssigned > 0 then
    FAssigneds := FAssigneds + [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_HslColor.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := FA_EG_ColorTransform.HandleElement(AReader);
  if Result = Nil then
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then
    Result.Assigneds := [xaRead];
end;

procedure TCT_HslColor.SetHSL(const Value: longword);
begin
  FHue := (Value and $00FF0000) shr 16;
  FSat := (Value and $0000FF00) shr 8;
  FLum := Value and $000000FF;
end;

procedure TCT_HslColor.Write(AWriter: TXpgWriteXML);
begin
  FA_EG_ColorTransform.Write(AWriter);
end;

procedure TCT_HslColor.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('hue',XmlIntToStr(FHue));
  AWriter.AddAttribute('sat',XmlIntToStr(FSat));
  AWriter.AddAttribute('lum',XmlIntToStr(FLum));
end;

procedure TCT_HslColor.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $00000142: FHue := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000148: FSat := XmlStrToIntDef(AAttributes.Values[i],0);
      $0000014E: FLum := XmlStrToIntDef(AAttributes.Values[i],0);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_HslColor.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 3;
  FA_EG_ColorTransform := TEG_ColorTransform.Create(FOwner);
  FHue := 2147483632;
  FSat := 2147483632;
  FLum := 2147483632;
end;

destructor TCT_HslColor.Destroy;
begin
  FA_EG_ColorTransform.Free;
end;

function TCT_HslColor.GetHSL: longword;
begin
  Result := (FHue shl 16) + (FSat shl 8) + FLum;
end;

procedure TCT_HslColor.Clear;
begin
  FAssigneds := [];
  FA_EG_ColorTransform.Clear;
  FHue := 2147483632;
  FSat := 2147483632;
  FLum := 2147483632;
end;

procedure TCT_HslColor.Assign(AItem: TCT_HslColor);
begin
  FHue := AItem.FHue;
  FSat := AItem.FSat;
  FLum := AItem.FLum;
  FA_EG_ColorTransform.Assign(AItem.FA_EG_ColorTransform);
end;

procedure TCT_HslColor.CopyTo(AItem: TCT_HslColor);
begin
end;

{ TCT_SystemColor }

function  TCT_SystemColor.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [xaAttributes];
  Inc(ElemsAssigned,FA_EG_ColorTransform.CheckAssigned);
  Result := 1;
  if ElemsAssigned > 0 then
    FAssigneds := FAssigneds + [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_SystemColor.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := FA_EG_ColorTransform.HandleElement(AReader);
  if Result = Nil then 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_SystemColor.Write(AWriter: TXpgWriteXML);
begin
  FA_EG_ColorTransform.Write(AWriter);
end;

procedure TCT_SystemColor.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FVal <> TST_SystemColorVal(XPG_UNKNOWN_ENUM) then
    AWriter.AddAttribute('val',StrTST_SystemColorVal[Integer(FVal)]);
  if FLastClr <> -16 then 
    AWriter.AddAttribute('lastClr',XmlIntToHexStr(FLastClr,6));
end;

procedure TCT_SystemColor.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $00000143: FVal := TST_SystemColorVal(StrToEnum('stscv' + AAttributes.Values[i]));
      $000002D5: FLastClr := XmlStrToIntDef('$' + AAttributes.Values[i],0);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_SystemColor.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 2;
  FA_EG_ColorTransform := TEG_ColorTransform.Create(FOwner);
  Clear;
end;

destructor TCT_SystemColor.Destroy;
begin
  FA_EG_ColorTransform.Free;
end;

procedure TCT_SystemColor.Clear;
begin
  FAssigneds := [];
  FA_EG_ColorTransform.Clear;
  FVal := TST_SystemColorVal(XPG_UNKNOWN_ENUM);
  FLastClr := -16;
end;

procedure TCT_SystemColor.Assign(AItem: TCT_SystemColor);
begin
  FVal := AItem.FVal;
  FLastClr := AItem.FLastClr;
  FA_EG_ColorTransform.Assign(AItem.FA_EG_ColorTransform);
end;

procedure TCT_SystemColor.CopyTo(AItem: TCT_SystemColor);
begin
end;

{ TCT_SchemeColor }

function  TCT_SchemeColor.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [xaAttributes];
  Inc(ElemsAssigned,FA_EG_ColorTransform.CheckAssigned);
  Result := 1;
  if ElemsAssigned > 0 then
    FAssigneds := FAssigneds + [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_SchemeColor.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := FA_EG_ColorTransform.HandleElement(AReader);
  if Result = Nil then 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_SchemeColor.Write(AWriter: TXpgWriteXML);
begin
  FA_EG_ColorTransform.Write(AWriter);
end;

procedure TCT_SchemeColor.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FVal <> TST_SchemeColorVal(XPG_UNKNOWN_ENUM) then
    AWriter.AddAttribute('val',StrTST_SchemeColorVal[Integer(FVal)]);
end;

procedure TCT_SchemeColor.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'val' then 
    FVal := TST_SchemeColorVal(StrToEnum('stscv' + AAttributes.Values[0]))
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_SchemeColor.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 1;
  FA_EG_ColorTransform := TEG_ColorTransform.Create(FOwner);
  FVal := TST_SchemeColorVal(XPG_UNKNOWN_ENUM);
end;

destructor TCT_SchemeColor.Destroy;
begin
  FA_EG_ColorTransform.Free;
end;

procedure TCT_SchemeColor.Clear;
begin
  FAssigneds := [];
  FA_EG_ColorTransform.Clear;
  FVal := TST_SchemeColorVal(XPG_UNKNOWN_ENUM);
end;

procedure TCT_SchemeColor.Assign(AItem: TCT_SchemeColor);
begin
  FVal := AItem.FVal;
  FA_EG_ColorTransform.Assign(AItem.FA_EG_ColorTransform);
end;

procedure TCT_SchemeColor.CopyTo(AItem: TCT_SchemeColor);
begin
end;

{ TCT_PresetColor }

function  TCT_PresetColor.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if Integer(FVal) <> XPG_UNKNOWN_ENUM then
    Inc(AttrsAssigned);
  Inc(ElemsAssigned,FA_EG_ColorTransform.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_PresetColor.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := FA_EG_ColorTransform.HandleElement(AReader);
  if Result = Nil then 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_PresetColor.Write(AWriter: TXpgWriteXML);
begin
  FA_EG_ColorTransform.Write(AWriter);
end;

procedure TCT_PresetColor.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if Integer(FVal) <> XPG_UNKNOWN_ENUM then
    AWriter.AddAttribute('val',StrTST_PresetColorVal[Integer(FVal)]);
end;

procedure TCT_PresetColor.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'val' then 
    FVal := TST_PresetColorVal(StrToEnum('stpcv' + AAttributes.Values[0]))
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_PresetColor.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 1;
  FA_EG_ColorTransform := TEG_ColorTransform.Create(FOwner);
  FVal := TST_PresetColorVal(XPG_UNKNOWN_ENUM);
end;

destructor TCT_PresetColor.Destroy;
begin
  FA_EG_ColorTransform.Free;
end;

procedure TCT_PresetColor.Clear;
begin
  FAssigneds := [];
  FA_EG_ColorTransform.Clear;
  FVal := TST_PresetColorVal(XPG_UNKNOWN_ENUM);
end;

procedure TCT_PresetColor.Assign(AItem: TCT_PresetColor);
begin
  FVal := AItem.FVal;
  FA_EG_ColorTransform.Assign(AItem.FA_EG_ColorTransform);
end;

procedure TCT_PresetColor.CopyTo(AItem: TCT_PresetColor);
begin
end;

{ TEG_ColorChoice }

function  TEG_ColorChoice.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FA_ScrgbClr <> Nil then
    Inc(ElemsAssigned,FA_ScrgbClr.CheckAssigned);
  if FA_SrgbClr <> Nil then
    Inc(ElemsAssigned,FA_SrgbClr.CheckAssigned);
  if FA_HslClr <> Nil then
    Inc(ElemsAssigned,FA_HslClr.CheckAssigned);
  if FA_SysClr <> Nil then
    Inc(ElemsAssigned,FA_SysClr.CheckAssigned);
  if FA_SchemeClr <> Nil then
    Inc(ElemsAssigned,FA_SchemeClr.CheckAssigned);
  if FA_PrstClr <> Nil then
    Inc(ElemsAssigned,FA_PrstClr.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TEG_ColorChoice.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  case AReader.QNameHashA of
    $000003CD: begin
      if FA_ScrgbClr = Nil then
        FA_ScrgbClr := TCT_ScRgbColor.Create(FOwner);
      Result := FA_ScrgbClr;
    end;
    $0000036A: begin
      if FA_SrgbClr = Nil then
        FA_SrgbClr := TCT_SRgbColor.Create(FOwner);
      Result := FA_SrgbClr;
    end;
    $00000303: begin
      if FA_HslClr = Nil then
        FA_HslClr := TCT_HslColor.Create(FOwner);
      Result := FA_HslClr;
    end;
    $0000031B: begin
      if FA_SysClr = Nil then
        FA_SysClr := TCT_SystemColor.Create(FOwner);
      Result := FA_SysClr;
    end;
    $00000431: begin
      if FA_SchemeClr = Nil then
        FA_SchemeClr := TCT_SchemeColor.Create(FOwner);
      Result := FA_SchemeClr;
    end;
    $00000385: begin
      if FA_PrstClr = Nil then
        FA_PrstClr := TCT_PresetColor.Create(FOwner);
      Result := FA_PrstClr;
    end;
    else
    begin
      Result := Nil;
      Exit;
    end
  end;
  if Result <> Self then
    Result.Assigneds := [xaRead];
end;

procedure TEG_ColorChoice.Write(AWriter: TXpgWriteXML);
begin
  if (FA_ScrgbClr <> Nil) and FA_ScrgbClr.Assigned then begin
    FA_ScrgbClr.WriteAttributes(AWriter);
    if xaElements in FA_ScrgbClr.FAssigneds then begin
      AWriter.BeginTag('a:scrgbClr');
      FA_ScrgbClr.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('a:scrgbClr');
  end;
  if (FA_SrgbClr <> Nil) and FA_SrgbClr.Assigned then begin
    FA_SrgbClr.WriteAttributes(AWriter);
    if xaElements in FA_SrgbClr.FAssigneds then begin
      AWriter.BeginTag('a:srgbClr');
      FA_SrgbClr.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('a:srgbClr');
  end;
  if (FA_HslClr <> Nil) and FA_HslClr.Assigned then begin
    FA_HslClr.WriteAttributes(AWriter);
    if xaElements in FA_HslClr.FAssigneds then begin
      AWriter.BeginTag('a:hslClr');
      FA_HslClr.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('a:hslClr');
  end;
  if (FA_SysClr <> Nil) and FA_SysClr.Assigned then begin
    FA_SysClr.WriteAttributes(AWriter);
    if xaElements in FA_SysClr.FAssigneds then begin
      AWriter.BeginTag('a:sysClr');
      FA_SysClr.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('a:sysClr');
  end;
  if (FA_SchemeClr <> Nil) and FA_SchemeClr.Assigned then begin
    FA_SchemeClr.WriteAttributes(AWriter);
    if xaElements in FA_SchemeClr.FAssigneds then begin
      AWriter.BeginTag('a:schemeClr');
      FA_SchemeClr.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('a:schemeClr');
  end;
  if (FA_PrstClr <> Nil) and FA_PrstClr.Assigned then begin
    FA_PrstClr.WriteAttributes(AWriter);
    if xaElements in FA_PrstClr.FAssigneds then begin
      AWriter.BeginTag('a:prstClr');
      FA_PrstClr.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('a:prstClr');
  end;
end;

constructor TEG_ColorChoice.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 6;
  FAttributeCount := 0;
end;

destructor TEG_ColorChoice.Destroy;
begin
  if FA_ScrgbClr <> Nil then
    FA_ScrgbClr.Free;
  if FA_SrgbClr <> Nil then
    FA_SrgbClr.Free;
  if FA_HslClr <> Nil then
    FA_HslClr.Free;
  if FA_SysClr <> Nil then
    FA_SysClr.Free;
  if FA_SchemeClr <> Nil then
    FA_SchemeClr.Free;
  if FA_PrstClr <> Nil then
    FA_PrstClr.Free;
end;

procedure TEG_ColorChoice.Clear;
begin
  FAssigneds := [];
  if FA_ScrgbClr <> Nil then
    FreeAndNil(FA_ScrgbClr);
  if FA_SrgbClr <> Nil then
    FreeAndNil(FA_SrgbClr);
  if FA_HslClr <> Nil then
    FreeAndNil(FA_HslClr);
  if FA_SysClr <> Nil then
    FreeAndNil(FA_SysClr);
  if FA_SchemeClr <> Nil then
    FreeAndNil(FA_SchemeClr);
  if FA_PrstClr <> Nil then
    FreeAndNil(FA_PrstClr);
end;

procedure TEG_ColorChoice.Assign(AItem: TEG_ColorChoice);
begin
  if AItem.FA_ScrgbClr <> Nil then begin
    Create_ScrgbClr;
    FA_ScrgbClr.Assign(AItem.FA_ScrgbClr);
  end;
  if AItem.FA_SrgbClr <> Nil then begin
    Create_SrgbClr;
    FA_SrgbClr.Assign(AItem.FA_SrgbClr);
  end;
  if AItem.FA_HslClr <> Nil then begin
    Create_HslClr;
    FA_HslClr.Assign(AItem.FA_HslClr);
  end;
  if AItem.FA_SysClr <> Nil then begin
    Create_SysClr;
    FA_SysClr.Assign(AItem.FA_SysClr);
  end;
  if AItem.FA_SchemeClr <> Nil then begin
    Create_SchemeClr;
    FA_SchemeClr.Assign(AItem.FA_SchemeClr);
  end;
  if AItem.FA_PrstClr <> Nil then begin
    Create_PrstClr;
    FA_PrstClr.Assign(AItem.FA_PrstClr);
  end;
end;

procedure TEG_ColorChoice.CopyTo(AItem: TEG_ColorChoice);
begin
end;

function  TEG_ColorChoice.Create_ScrgbClr: TCT_ScRgbColor;
begin
  if FA_ScrgbClr = Nil then begin
    Clear;
    FA_ScrgbClr := TCT_ScRgbColor.Create(FOwner);
  end;
  Result := FA_ScrgbClr;
end;

function  TEG_ColorChoice.Create_SrgbClr: TCT_SRgbColor;
begin
  if FA_SrgbClr = Nil then begin
    Clear;
    FA_SrgbClr := TCT_SRgbColor.Create(FOwner);
  end;
  Result := FA_SrgbClr;
end;

function  TEG_ColorChoice.Create_HslClr: TCT_HslColor;
begin
  if FA_HslClr = Nil then begin
    Clear;
    FA_HslClr := TCT_HslColor.Create(FOwner);
  end;
  Result := FA_HslClr;
end;

function  TEG_ColorChoice.Create_SysClr: TCT_SystemColor;
begin
  if FA_SysClr = Nil then begin
    Clear;
    FA_SysClr := TCT_SystemColor.Create(FOwner);
  end;
  Result := FA_SysClr;
end;

function  TEG_ColorChoice.Create_SchemeClr: TCT_SchemeColor;
begin
  if FA_SchemeClr = Nil then begin
    Clear;
    FA_SchemeClr := TCT_SchemeColor.Create(FOwner);
  end;
  Result := FA_SchemeClr;
end;

function  TEG_ColorChoice.Create_PrstClr: TCT_PresetColor;
begin
  if FA_PrstClr = Nil then begin
    Clear;
    FA_PrstClr := TCT_PresetColor.Create(FOwner);
  end;
  Result := FA_PrstClr;
end;

{ TCT_OfficeArtExtension }

function  TCT_OfficeArtExtension.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FUri <> '' then
    Inc(AttrsAssigned);
  Inc(ElemsAssigned,FAnyElements.Count);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_OfficeArtExtension.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := FAnyElements.Add(AReader.QName,AReader.Text);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_OfficeArtExtension.Write(AWriter: TXpgWriteXML);
begin
  FAnyElements.Write(AWriter);
end;

procedure TCT_OfficeArtExtension.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FUri <> '' then 
    AWriter.AddAttribute('uri',FUri);
end;

procedure TCT_OfficeArtExtension.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'uri' then 
    FUri := AAttributes.Values[0]
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_OfficeArtExtension.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 1;
  FAnyElements := TXPGAnyElements.Create;
  FUri := '';
end;

destructor TCT_OfficeArtExtension.Destroy;
begin
  FAnyElements.Free;
end;

procedure TCT_OfficeArtExtension.Clear;
begin
  FAssigneds := [];
  FAnyElements.Clear;
  FUri := '';
end;

procedure TCT_OfficeArtExtension.Assign(AItem: TCT_OfficeArtExtension);
begin
  FUri := AItem.FUri;
  FAnyElements.Assign(AItem.FAnyElements);
end;

procedure TCT_OfficeArtExtension.CopyTo(AItem: TCT_OfficeArtExtension);
begin
end;

{ TCT_OfficeArtExtensionXpgList }

function  TCT_OfficeArtExtensionXpgList.GetItems(Index: integer): TCT_OfficeArtExtension;
begin
  Result := TCT_OfficeArtExtension(inherited Items[Index]);
end;

function  TCT_OfficeArtExtensionXpgList.Add: TCT_OfficeArtExtension;
begin
  Result := TCT_OfficeArtExtension.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_OfficeArtExtensionXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_OfficeArtExtensionXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_OfficeArtExtensionXpgList.Assign(AItem: TCT_OfficeArtExtensionXpgList);
begin
end;

procedure TCT_OfficeArtExtensionXpgList.CopyTo(AItem: TCT_OfficeArtExtensionXpgList);
begin
end;

{ TCT_RelativeRect }

function  TCT_RelativeRect.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FL <> 0 then 
    Inc(AttrsAssigned);
  if FT <> 0 then 
    Inc(AttrsAssigned);
  if FR <> 0 then 
    Inc(AttrsAssigned);
  if FB <> 0 then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_RelativeRect.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_RelativeRect.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FL <> 0 then 
    AWriter.AddAttribute('l',XmlIntToStr(FL));
  if FT <> 0 then 
    AWriter.AddAttribute('t',XmlIntToStr(FT));
  if FR <> 0 then 
    AWriter.AddAttribute('r',XmlIntToStr(FR));
  if FB <> 0 then 
    AWriter.AddAttribute('b',XmlIntToStr(FB));
end;

procedure TCT_RelativeRect.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $0000006C: FL := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000074: FT := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000072: FR := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000062: FB := XmlStrToIntDef(AAttributes.Values[i],0);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_RelativeRect.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 4;
  FL := 0;
  FT := 0;
  FR := 0;
  FB := 0;
end;

destructor TCT_RelativeRect.Destroy;
begin
end;

function TCT_RelativeRect.Equal(const ALeft, ARight, ATop, ABottom: integer): boolean;
begin
  Result := (FL = ALeft) and (FR = ARight) and (FT = ATop) and (FB = ABottom);
end;

function TCT_RelativeRect.SetValues(const ALeft, ARight, ATop, ABottom: integer): boolean;
begin
  FL := ALeft;
  FR := ARight;
  FT := ATop;
  FB := ABottom;
  Result := True;
end;

procedure TCT_RelativeRect.Clear;
begin
  FAssigneds := [];
  FL := 0;
  FT := 0;
  FR := 0;
  FB := 0;
end;

procedure TCT_RelativeRect.Assign(AItem: TCT_RelativeRect);
begin
  FL := AItem.FL;
  FT := AItem.FT;
  FR := AItem.FR;
  FB := AItem.FB;
end;

procedure TCT_RelativeRect.CopyTo(AItem: TCT_RelativeRect);
begin
end;

{ TCT_Color }

function  TCT_Color.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FA_EG_ColorChoice.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_Color.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := FA_EG_ColorChoice.HandleElement(AReader);
  if Result = Nil then 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_Color.Write(AWriter: TXpgWriteXML);
begin
  FA_EG_ColorChoice.Write(AWriter);
end;

constructor TCT_Color.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 0;
  FA_EG_ColorChoice := TEG_ColorChoice.Create(FOwner);
end;

destructor TCT_Color.Destroy;
begin
  FA_EG_ColorChoice.Free;
end;

procedure TCT_Color.Clear;
begin
  FAssigneds := [];
  FA_EG_ColorChoice.Clear;
end;

procedure TCT_Color.Assign(AItem: TCT_Color);
begin
  FA_EG_ColorChoice.Assign(AItem.FA_EG_ColorChoice);
end;

procedure TCT_Color.CopyTo(AItem: TCT_Color);
begin
end;

{ TEG_OfficeArtExtensionList }

function  TEG_OfficeArtExtensionList.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FA_Ext <> Nil then 
    Inc(ElemsAssigned,FA_Ext.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TEG_OfficeArtExtensionList.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'a:ext' then 
  begin
    if FA_Ext = Nil then 
      FA_Ext := TCT_OfficeArtExtension.Create(FOwner);
    Result := FA_Ext;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TEG_OfficeArtExtensionList.Write(AWriter: TXpgWriteXML);
begin
  if (FA_Ext <> Nil) and FA_Ext.Assigned then 
  begin
    FA_Ext.WriteAttributes(AWriter);
    if xaElements in FA_Ext.FAssigneds then 
    begin
      AWriter.BeginTag('a:ext');
      FA_Ext.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:ext');
  end;
end;

constructor TEG_OfficeArtExtensionList.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 0;
end;

destructor TEG_OfficeArtExtensionList.Destroy;
begin
  if FA_Ext <> Nil then 
    FA_Ext.Free;
end;

procedure TEG_OfficeArtExtensionList.Clear;
begin
  FAssigneds := [];
  if FA_Ext <> Nil then 
    FreeAndNil(FA_Ext);
end;

procedure TEG_OfficeArtExtensionList.Assign(AItem: TEG_OfficeArtExtensionList);
begin
  if AItem.FA_Ext <> Nil then begin
    Create_A_Ext;
    FA_Ext.Assign(AItem.FA_Ext);
  end;
end;

procedure TEG_OfficeArtExtensionList.CopyTo(AItem: TEG_OfficeArtExtensionList);
begin
end;

function  TEG_OfficeArtExtensionList.Create_A_Ext: TCT_OfficeArtExtension;
begin
  Result := TCT_OfficeArtExtension.Create(FOwner);
  FA_Ext := Result;
end;

{ TCT_GradientStop }

function  TCT_GradientStop.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [xaAttributes];
  Inc(ElemsAssigned,FA_EG_ColorChoice.CheckAssigned);
  Result := 1;
  if ElemsAssigned > 0 then
    FAssigneds := FAssigneds + [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_GradientStop.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := FA_EG_ColorChoice.HandleElement(AReader);
  if Result = Nil then 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_GradientStop.Write(AWriter: TXpgWriteXML);
begin
  FA_EG_ColorChoice.Write(AWriter);
end;

procedure TCT_GradientStop.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('pos',XmlIntToStr(FPos));
end;

procedure TCT_GradientStop.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'pos' then 
    FPos := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_GradientStop.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 1;
  FA_EG_ColorChoice := TEG_ColorChoice.Create(FOwner);
  FPos := 2147483632;
end;

destructor TCT_GradientStop.Destroy;
begin
  FA_EG_ColorChoice.Free;
end;

procedure TCT_GradientStop.Clear;
begin
  FAssigneds := [];
  FA_EG_ColorChoice.Clear;
  FPos := 2147483632;
end;

procedure TCT_GradientStop.Assign(AItem: TCT_GradientStop);
begin
  FPos := AItem.FPos;
  FA_EG_ColorChoice.Assign(AItem.FA_EG_ColorChoice);
end;

procedure TCT_GradientStop.CopyTo(AItem: TCT_GradientStop);
begin
end;

{ TCT_GradientStopXpgList }

function  TCT_GradientStopXpgList.GetItems(Index: integer): TCT_GradientStop;
begin
  Result := TCT_GradientStop(inherited Items[Index]);
end;

function  TCT_GradientStopXpgList.Add: TCT_GradientStop;
begin
  Result := TCT_GradientStop.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_GradientStopXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_GradientStopXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_GradientStopXpgList.Assign(AItem: TCT_GradientStopXpgList);
var
  i: integer;
begin
  for i := 0 to AItem.Count - 1 do
    Add.Assign(AItem[i]);
end;

procedure TCT_GradientStopXpgList.CopyTo(AItem: TCT_GradientStopXpgList);
begin
end;

{ TCT_LinearShadeProperties }

function  TCT_LinearShadeProperties.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FAng <> 2147483632 then 
    Inc(AttrsAssigned);
  if Byte(FScaled) <> 2 then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_LinearShadeProperties.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_LinearShadeProperties.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FAng <> 2147483632 then 
    AWriter.AddAttribute('ang',XmlIntToStr(FAng));
  if Byte(FScaled) <> 2 then 
    AWriter.AddAttribute('scaled',XmlBoolToStr(FScaled));
end;

procedure TCT_LinearShadeProperties.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $00000136: FAng := XmlStrToIntDef(AAttributes.Values[i],0);
      $0000026C: FScaled := XmlStrToBoolDef(AAttributes.Values[i],False);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_LinearShadeProperties.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 2;
  FAng := 2147483632;
  Byte(FScaled) := 2;
end;

destructor TCT_LinearShadeProperties.Destroy;
begin
end;

procedure TCT_LinearShadeProperties.Clear;
begin
  FAssigneds := [];
  FAng := 2147483632;
  Byte(FScaled) := 2;
end;

procedure TCT_LinearShadeProperties.Assign(AItem: TCT_LinearShadeProperties);
begin
end;

procedure TCT_LinearShadeProperties.CopyTo(AItem: TCT_LinearShadeProperties);
begin
end;

{ TCT_PathShadeProperties }

function  TCT_PathShadeProperties.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if Integer(FPath) <> XPG_UNKNOWN_ENUM then 
    Inc(AttrsAssigned);
  if FA_FillToRect <> Nil then 
    Inc(ElemsAssigned,FA_FillToRect.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_PathShadeProperties.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'a:fillToRect' then 
  begin
    if FA_FillToRect = Nil then 
      FA_FillToRect := TCT_RelativeRect.Create(FOwner);
    Result := FA_FillToRect;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_PathShadeProperties.Write(AWriter: TXpgWriteXML);
begin
  if (FA_FillToRect <> Nil) and FA_FillToRect.Assigned then 
  begin
    FA_FillToRect.WriteAttributes(AWriter);
    AWriter.SimpleTag('a:fillToRect');
  end;
end;

procedure TCT_PathShadeProperties.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if Integer(FPath) <> XPG_UNKNOWN_ENUM then 
    AWriter.AddAttribute('path',StrTST_PathShadeType[Integer(FPath)]);
end;

procedure TCT_PathShadeProperties.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'path' then
    FPath := TST_PathShadeType(StrToEnum('stpst' + AAttributes.Values[0]))
  else
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_PathShadeProperties.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 1;
  FPath := TST_PathShadeType(XPG_UNKNOWN_ENUM);
end;

destructor TCT_PathShadeProperties.Destroy;
begin
  if FA_FillToRect <> Nil then
    FA_FillToRect.Free;
end;

procedure TCT_PathShadeProperties.Free_FillToRect;
begin
  if FA_FillToRect <> Nil then
    FreeAndNil(FA_FillToRect);
end;

procedure TCT_PathShadeProperties.Clear;
begin
  FAssigneds := [];
  if FA_FillToRect <> Nil then
    FreeAndNil(FA_FillToRect);
  FPath := TST_PathShadeType(XPG_UNKNOWN_ENUM);
end;

procedure TCT_PathShadeProperties.Assign(AItem: TCT_PathShadeProperties);
begin
end;

procedure TCT_PathShadeProperties.CopyTo(AItem: TCT_PathShadeProperties);
begin
end;

function  TCT_PathShadeProperties.Create_FillToRect: TCT_RelativeRect;
begin
  if FA_FillToRect = Nil then
    FA_FillToRect := TCT_RelativeRect.Create(FOwner);
  Result := FA_FillToRect;
end;

{ TCT_AlphaBiLevelEffect }

function  TCT_AlphaBiLevelEffect.CheckAssigned: integer;
begin
  FAssigneds := [xaAttributes];
  Result := 1;
end;

procedure TCT_AlphaBiLevelEffect.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_AlphaBiLevelEffect.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('thresh',XmlIntToStr(FThresh));
end;

procedure TCT_AlphaBiLevelEffect.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'thresh' then 
    FThresh := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_AlphaBiLevelEffect.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
  FThresh := 2147483632;
end;

destructor TCT_AlphaBiLevelEffect.Destroy;
begin
end;

procedure TCT_AlphaBiLevelEffect.Clear;
begin
  FAssigneds := [];
  FThresh := 2147483632;
end;

procedure TCT_AlphaBiLevelEffect.Assign(AItem: TCT_AlphaBiLevelEffect);
begin
  FThresh := AItem.FThresh;
end;

procedure TCT_AlphaBiLevelEffect.CopyTo(AItem: TCT_AlphaBiLevelEffect);
begin
end;

{ TCT_AlphaBiLevelEffectXpgList }

function  TCT_AlphaBiLevelEffectXpgList.GetItems(Index: integer): TCT_AlphaBiLevelEffect;
begin
  Result := TCT_AlphaBiLevelEffect(inherited Items[Index]);
end;

function  TCT_AlphaBiLevelEffectXpgList.Add: TCT_AlphaBiLevelEffect;
begin
  Result := TCT_AlphaBiLevelEffect.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_AlphaBiLevelEffectXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_AlphaBiLevelEffectXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_AlphaBiLevelEffectXpgList.Assign(AItem: TCT_AlphaBiLevelEffectXpgList);
var
  i: integer;
begin
  for i := 0 to AItem.Count - 1 do
    Add.Assign(AItem[i]);
end;

procedure TCT_AlphaBiLevelEffectXpgList.CopyTo(AItem: TCT_AlphaBiLevelEffectXpgList);
begin
end;

{ TCT_AlphaCeilingEffect }

function  TCT_AlphaCeilingEffect.CheckAssigned: integer;
begin
  FAssigneds := [];
  Result := 0;
end;

procedure TCT_AlphaCeilingEffect.Write(AWriter: TXpgWriteXML);
begin
end;

constructor TCT_AlphaCeilingEffect.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 0;
end;

destructor TCT_AlphaCeilingEffect.Destroy;
begin
end;

procedure TCT_AlphaCeilingEffect.Clear;
begin
  FAssigneds := [];
end;

procedure TCT_AlphaCeilingEffect.Assign(AItem: TCT_AlphaCeilingEffect);
begin
end;

procedure TCT_AlphaCeilingEffect.CopyTo(AItem: TCT_AlphaCeilingEffect);
begin
end;

{ TCT_AlphaCeilingEffectXpgList }

function  TCT_AlphaCeilingEffectXpgList.GetItems(Index: integer): TCT_AlphaCeilingEffect;
begin
  Result := TCT_AlphaCeilingEffect(inherited Items[Index]);
end;

function  TCT_AlphaCeilingEffectXpgList.Add: TCT_AlphaCeilingEffect;
begin
  Result := TCT_AlphaCeilingEffect.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_AlphaCeilingEffectXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_AlphaCeilingEffectXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
end;

procedure TCT_AlphaCeilingEffectXpgList.Assign(AItem: TCT_AlphaCeilingEffectXpgList);
var
  i: integer;
begin
  for i := 0 to AItem.Count - 1 do
    Add.Assign(AItem[i]);
end;

procedure TCT_AlphaCeilingEffectXpgList.CopyTo(AItem: TCT_AlphaCeilingEffectXpgList);
begin
end;

{ TCT_AlphaFloorEffect }

function  TCT_AlphaFloorEffect.CheckAssigned: integer;
begin
  FAssigneds := [];
  Result := 0;
end;

procedure TCT_AlphaFloorEffect.Write(AWriter: TXpgWriteXML);
begin
end;

constructor TCT_AlphaFloorEffect.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 0;
end;

destructor TCT_AlphaFloorEffect.Destroy;
begin
end;

procedure TCT_AlphaFloorEffect.Clear;
begin
  FAssigneds := [];
end;

procedure TCT_AlphaFloorEffect.Assign(AItem: TCT_AlphaFloorEffect);
begin
end;

procedure TCT_AlphaFloorEffect.CopyTo(AItem: TCT_AlphaFloorEffect);
begin
end;

{ TCT_AlphaFloorEffectXpgList }

function  TCT_AlphaFloorEffectXpgList.GetItems(Index: integer): TCT_AlphaFloorEffect;
begin
  Result := TCT_AlphaFloorEffect(inherited Items[Index]);
end;

function  TCT_AlphaFloorEffectXpgList.Add: TCT_AlphaFloorEffect;
begin
  Result := TCT_AlphaFloorEffect.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_AlphaFloorEffectXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_AlphaFloorEffectXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
end;

procedure TCT_AlphaFloorEffectXpgList.Assign(AItem: TCT_AlphaFloorEffectXpgList);
var
  i: integer;
begin
  for i := 0 to AItem.Count - 1 do
    Add.Assign(AItem[i]);
end;

procedure TCT_AlphaFloorEffectXpgList.CopyTo(AItem: TCT_AlphaFloorEffectXpgList);
begin
end;

{ TCT_AlphaInverseEffect }

function  TCT_AlphaInverseEffect.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FA_EG_ColorChoice.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_AlphaInverseEffect.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := FA_EG_ColorChoice.HandleElement(AReader);
  if Result = Nil then 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_AlphaInverseEffect.Write(AWriter: TXpgWriteXML);
begin
  FA_EG_ColorChoice.Write(AWriter);
end;

constructor TCT_AlphaInverseEffect.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 0;
  FA_EG_ColorChoice := TEG_ColorChoice.Create(FOwner);
end;

destructor TCT_AlphaInverseEffect.Destroy;
begin
  FA_EG_ColorChoice.Free;
end;

procedure TCT_AlphaInverseEffect.Clear;
begin
  FAssigneds := [];
  FA_EG_ColorChoice.Clear;
end;

procedure TCT_AlphaInverseEffect.Assign(AItem: TCT_AlphaInverseEffect);
begin
  FA_EG_ColorChoice.Assign(AItem.FA_EG_ColorChoice);
end;

procedure TCT_AlphaInverseEffect.CopyTo(AItem: TCT_AlphaInverseEffect);
begin
end;

{ TCT_AlphaInverseEffectXpgList }

function  TCT_AlphaInverseEffectXpgList.GetItems(Index: integer): TCT_AlphaInverseEffect;
begin
  Result := TCT_AlphaInverseEffect(inherited Items[Index]);
end;

function  TCT_AlphaInverseEffectXpgList.Add: TCT_AlphaInverseEffect;
begin
  Result := TCT_AlphaInverseEffect.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_AlphaInverseEffectXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_AlphaInverseEffectXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
end;

procedure TCT_AlphaInverseEffectXpgList.Assign(AItem: TCT_AlphaInverseEffectXpgList);
var
  i: integer;
begin
  for i := 0 to AItem.Count - 1 do
    Add.Assign(AItem[i]);
end;

procedure TCT_AlphaInverseEffectXpgList.CopyTo(AItem: TCT_AlphaInverseEffectXpgList);
begin
end;

{ TCT_AlphaModulateFixedEffect }

function  TCT_AlphaModulateFixedEffect.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FAmt <> 100000 then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_AlphaModulateFixedEffect.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_AlphaModulateFixedEffect.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FAmt <> 100000 then 
    AWriter.AddAttribute('amt',XmlIntToStr(FAmt));
end;

procedure TCT_AlphaModulateFixedEffect.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'amt' then 
    FAmt := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_AlphaModulateFixedEffect.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
  FAmt := 100000;
end;

destructor TCT_AlphaModulateFixedEffect.Destroy;
begin
end;

procedure TCT_AlphaModulateFixedEffect.Clear;
begin
  FAssigneds := [];
  FAmt := 100000;
end;

procedure TCT_AlphaModulateFixedEffect.Assign(AItem: TCT_AlphaModulateFixedEffect);
begin
  FAmt := AITem.FAmt;
end;

procedure TCT_AlphaModulateFixedEffect.CopyTo(AItem: TCT_AlphaModulateFixedEffect);
begin
end;

{ TCT_AlphaModulateFixedEffectXpgList }

function  TCT_AlphaModulateFixedEffectXpgList.GetItems(Index: integer): TCT_AlphaModulateFixedEffect;
begin
  Result := TCT_AlphaModulateFixedEffect(inherited Items[Index]);
end;

function  TCT_AlphaModulateFixedEffectXpgList.Add: TCT_AlphaModulateFixedEffect;
begin
  Result := TCT_AlphaModulateFixedEffect.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_AlphaModulateFixedEffectXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_AlphaModulateFixedEffectXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_AlphaModulateFixedEffectXpgList.Assign(AItem: TCT_AlphaModulateFixedEffectXpgList);
var
  i: integer;
begin
  for i := 0 to AItem.Count - 1 do
    Add.Assign(AItem[i]);
end;

procedure TCT_AlphaModulateFixedEffectXpgList.CopyTo(AItem: TCT_AlphaModulateFixedEffectXpgList);
begin
end;

{ TCT_AlphaReplaceEffect }

function  TCT_AlphaReplaceEffect.CheckAssigned: integer;
begin
  FAssigneds := [xaAttributes];
  Result := 1;
end;

procedure TCT_AlphaReplaceEffect.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_AlphaReplaceEffect.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('a',XmlIntToStr(FA));
end;

procedure TCT_AlphaReplaceEffect.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'a' then 
    FA := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_AlphaReplaceEffect.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
  FA := 2147483632;
end;

destructor TCT_AlphaReplaceEffect.Destroy;
begin
end;

procedure TCT_AlphaReplaceEffect.Clear;
begin
  FAssigneds := [];
  FA := 2147483632;
end;

procedure TCT_AlphaReplaceEffect.Assign(AItem: TCT_AlphaReplaceEffect);
begin
  FA := AItem.FA;
end;

procedure TCT_AlphaReplaceEffect.CopyTo(AItem: TCT_AlphaReplaceEffect);
begin
end;

{ TCT_AlphaReplaceEffectXpgList }

function  TCT_AlphaReplaceEffectXpgList.GetItems(Index: integer): TCT_AlphaReplaceEffect;
begin
  Result := TCT_AlphaReplaceEffect(inherited Items[Index]);
end;

function  TCT_AlphaReplaceEffectXpgList.Add: TCT_AlphaReplaceEffect;
begin
  Result := TCT_AlphaReplaceEffect.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_AlphaReplaceEffectXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_AlphaReplaceEffectXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_AlphaReplaceEffectXpgList.Assign(AItem: TCT_AlphaReplaceEffectXpgList);
var
  i: integer;
begin
  for i := 0 to AItem.Count - 1 do
    Add.Assign(AItem[i]);
end;

procedure TCT_AlphaReplaceEffectXpgList.CopyTo(AItem: TCT_AlphaReplaceEffectXpgList);
begin
end;

{ TCT_BiLevelEffect }

function  TCT_BiLevelEffect.CheckAssigned: integer;
begin
  FAssigneds := [xaAttributes];
  Result := 1;
end;

procedure TCT_BiLevelEffect.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_BiLevelEffect.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('thresh',XmlIntToStr(FThresh));
end;

procedure TCT_BiLevelEffect.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'thresh' then 
    FThresh := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_BiLevelEffect.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
  FThresh := 2147483632;
end;

destructor TCT_BiLevelEffect.Destroy;
begin
end;

procedure TCT_BiLevelEffect.Clear;
begin
  FAssigneds := [];
  FThresh := 2147483632;
end;

procedure TCT_BiLevelEffect.Assign(AItem: TCT_BiLevelEffect);
begin
  FThresh := AItem.FThresh;
end;

procedure TCT_BiLevelEffect.CopyTo(AItem: TCT_BiLevelEffect);
begin
end;

{ TCT_BiLevelEffectXpgList }

function  TCT_BiLevelEffectXpgList.GetItems(Index: integer): TCT_BiLevelEffect;
begin
  Result := TCT_BiLevelEffect(inherited Items[Index]);
end;

function  TCT_BiLevelEffectXpgList.Add: TCT_BiLevelEffect;
begin
  Result := TCT_BiLevelEffect.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_BiLevelEffectXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_BiLevelEffectXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_BiLevelEffectXpgList.Assign(AItem: TCT_BiLevelEffectXpgList);
var
  i: integer;
begin
  for i := 0 to AItem.Count - 1 do
    Add.Assign(AItem[i]);
end;

procedure TCT_BiLevelEffectXpgList.CopyTo(AItem: TCT_BiLevelEffectXpgList);
begin
end;

{ TCT_BlurEffect }

function  TCT_BlurEffect.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FRad <> 0 then 
    Inc(AttrsAssigned);
  if FGrow <> True then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_BlurEffect.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_BlurEffect.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FRad <> 0 then 
    AWriter.AddAttribute('rad',XmlIntToStr(FRad));
  if FGrow <> True then 
    AWriter.AddAttribute('grow',XmlBoolToStr(FGrow));
end;

procedure TCT_BlurEffect.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $00000137: FRad := XmlStrToIntDef(AAttributes.Values[i],0);
      $000001BF: FGrow := XmlStrToBoolDef(AAttributes.Values[i],True);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_BlurEffect.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 2;
  FRad := 0;
  FGrow := True;
end;

destructor TCT_BlurEffect.Destroy;
begin
end;

procedure TCT_BlurEffect.Clear;
begin
  FAssigneds := [];
  FRad := 0;
  FGrow := True;
end;

procedure TCT_BlurEffect.Assign(AItem: TCT_BlurEffect);
begin
  FRad := AItem.FRad;
  FGrow := AItem.FGrow;
end;

procedure TCT_BlurEffect.CopyTo(AItem: TCT_BlurEffect);
begin
end;

{ TCT_BlurEffectXpgList }

function  TCT_BlurEffectXpgList.GetItems(Index: integer): TCT_BlurEffect;
begin
  Result := TCT_BlurEffect(inherited Items[Index]);
end;

function  TCT_BlurEffectXpgList.Add: TCT_BlurEffect;
begin
  Result := TCT_BlurEffect.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_BlurEffectXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_BlurEffectXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_BlurEffectXpgList.Assign(AItem: TCT_BlurEffectXpgList);
var
  i: integer;
begin
  for i := 0 to AItem.Count - 1 do
    Add.Assign(AItem[i]);
end;

procedure TCT_BlurEffectXpgList.CopyTo(AItem: TCT_BlurEffectXpgList);
begin
end;

{ TCT_ColorChangeEffect }

function  TCT_ColorChangeEffect.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FUseA <> True then 
    Inc(AttrsAssigned);
  if FA_ClrFrom <> Nil then 
    Inc(ElemsAssigned,FA_ClrFrom.CheckAssigned);
  if FA_ClrTo <> Nil then 
    Inc(ElemsAssigned,FA_ClrTo.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_ColorChangeEffect.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $00000370: begin
      if FA_ClrFrom = Nil then 
        FA_ClrFrom := TCT_Color.Create(FOwner);
      Result := FA_ClrFrom;
    end;
    $0000029F: begin
      if FA_ClrTo = Nil then 
        FA_ClrTo := TCT_Color.Create(FOwner);
      Result := FA_ClrTo;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_ColorChangeEffect.Write(AWriter: TXpgWriteXML);
begin
  if (FA_ClrFrom <> Nil) and FA_ClrFrom.Assigned then 
    if xaElements in FA_ClrFrom.FAssigneds then 
    begin
      AWriter.BeginTag('a:clrFrom');
      FA_ClrFrom.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:clrFrom')
  else 
    AWriter.SimpleTag('a:clrFrom');
  if (FA_ClrTo <> Nil) and FA_ClrTo.Assigned then 
    if xaElements in FA_ClrTo.FAssigneds then 
    begin
      AWriter.BeginTag('a:clrTo');
      FA_ClrTo.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:clrTo')
  else 
    AWriter.SimpleTag('a:clrTo');
end;

procedure TCT_ColorChangeEffect.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FUseA <> True then 
    AWriter.AddAttribute('useA',XmlBoolToStr(FUseA));
end;

procedure TCT_ColorChangeEffect.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'useA' then 
    FUseA := XmlStrToBoolDef(AAttributes.Values[0],True)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_ColorChangeEffect.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 2;
  FAttributeCount := 1;
  FUseA := True;
end;

destructor TCT_ColorChangeEffect.Destroy;
begin
  if FA_ClrFrom <> Nil then 
    FA_ClrFrom.Free;
  if FA_ClrTo <> Nil then 
    FA_ClrTo.Free;
end;

procedure TCT_ColorChangeEffect.Clear;
begin
  FAssigneds := [];
  if FA_ClrFrom <> Nil then 
    FreeAndNil(FA_ClrFrom);
  if FA_ClrTo <> Nil then 
    FreeAndNil(FA_ClrTo);
  FUseA := True;
end;

procedure TCT_ColorChangeEffect.Assign(AItem: TCT_ColorChangeEffect);
begin
  FUseA := AItem.FUseA;
  if AITem.FA_ClrFrom <> Nil then begin
    Create_A_ClrFrom;
    FA_ClrFrom.Assign(AItem.FA_ClrFrom);
  end;
  if AITem.FA_ClrTo <> Nil then begin
    Create_A_ClrTo;
    FA_ClrTo.Assign(AItem.FA_ClrTo);
  end;
end;

procedure TCT_ColorChangeEffect.CopyTo(AItem: TCT_ColorChangeEffect);
begin
end;

function  TCT_ColorChangeEffect.Create_A_ClrFrom: TCT_Color;
begin
  if FA_ClrFrom = Nil then
    FA_ClrFrom := TCT_Color.Create(FOwner);
  Result := FA_ClrFrom;
end;

function  TCT_ColorChangeEffect.Create_A_ClrTo: TCT_Color;
begin
  if FA_ClrTo = Nil then
    FA_ClrTo := TCT_Color.Create(FOwner);
  Result := FA_ClrTo;
end;

{ TCT_ColorChangeEffectXpgList }

function  TCT_ColorChangeEffectXpgList.GetItems(Index: integer): TCT_ColorChangeEffect;
begin
  Result := TCT_ColorChangeEffect(inherited Items[Index]);
end;

function  TCT_ColorChangeEffectXpgList.Add: TCT_ColorChangeEffect;
begin
  Result := TCT_ColorChangeEffect.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_ColorChangeEffectXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_ColorChangeEffectXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_ColorChangeEffectXpgList.Assign(AItem: TCT_ColorChangeEffectXpgList);
var
  i: integer;
begin
  for i := 0 to AItem.Count - 1 do
    Add.Assign(AItem[i]);
end;

procedure TCT_ColorChangeEffectXpgList.CopyTo(AItem: TCT_ColorChangeEffectXpgList);
begin
end;

{ TCT_ColorReplaceEffect }

function  TCT_ColorReplaceEffect.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FA_EG_ColorChoice.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_ColorReplaceEffect.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := FA_EG_ColorChoice.HandleElement(AReader);
  if Result = Nil then 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_ColorReplaceEffect.Write(AWriter: TXpgWriteXML);
begin
  FA_EG_ColorChoice.Write(AWriter);
end;

constructor TCT_ColorReplaceEffect.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 0;
  FA_EG_ColorChoice := TEG_ColorChoice.Create(FOwner);
end;

destructor TCT_ColorReplaceEffect.Destroy;
begin
  FA_EG_ColorChoice.Free;
end;

procedure TCT_ColorReplaceEffect.Clear;
begin
  FAssigneds := [];
  FA_EG_ColorChoice.Clear;
end;

procedure TCT_ColorReplaceEffect.Assign(AItem: TCT_ColorReplaceEffect);
begin
  FA_EG_ColorChoice.Assign(AItem.FA_EG_ColorChoice);
end;

procedure TCT_ColorReplaceEffect.CopyTo(AItem: TCT_ColorReplaceEffect);
begin
end;

{ TCT_ColorReplaceEffectXpgList }

function  TCT_ColorReplaceEffectXpgList.GetItems(Index: integer): TCT_ColorReplaceEffect;
begin
  Result := TCT_ColorReplaceEffect(inherited Items[Index]);
end;

function  TCT_ColorReplaceEffectXpgList.Add: TCT_ColorReplaceEffect;
begin
  Result := TCT_ColorReplaceEffect.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_ColorReplaceEffectXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_ColorReplaceEffectXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
end;

procedure TCT_ColorReplaceEffectXpgList.Assign(AItem: TCT_ColorReplaceEffectXpgList);
var
  i: integer;
begin
  for i := 0 to AItem.Count - 1 do
    Add.Assign(AItem[i]);
end;

procedure TCT_ColorReplaceEffectXpgList.CopyTo(AItem: TCT_ColorReplaceEffectXpgList);
begin
end;

{ TCT_DuotoneEffect }

function  TCT_DuotoneEffect.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FA_EG_ColorChoice.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_DuotoneEffect.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := FA_EG_ColorChoice.HandleElement(AReader);
  if Result = Nil then 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_DuotoneEffect.Write(AWriter: TXpgWriteXML);
begin
  FA_EG_ColorChoice.Write(AWriter);
end;

constructor TCT_DuotoneEffect.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 0;
  FA_EG_ColorChoice := TEG_ColorChoice.Create(FOwner);
end;

destructor TCT_DuotoneEffect.Destroy;
begin
  FA_EG_ColorChoice.Free;
end;

procedure TCT_DuotoneEffect.Clear;
begin
  FAssigneds := [];
  FA_EG_ColorChoice.Clear;
end;

procedure TCT_DuotoneEffect.Assign(AItem: TCT_DuotoneEffect);
begin
  FA_EG_ColorChoice.Assign(AItem.FA_EG_ColorChoice);
end;

procedure TCT_DuotoneEffect.CopyTo(AItem: TCT_DuotoneEffect);
begin
end;

{ TCT_DuotoneEffectXpgList }

function  TCT_DuotoneEffectXpgList.GetItems(Index: integer): TCT_DuotoneEffect;
begin
  Result := TCT_DuotoneEffect(inherited Items[Index]);
end;

function  TCT_DuotoneEffectXpgList.Add: TCT_DuotoneEffect;
begin
  Result := TCT_DuotoneEffect.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_DuotoneEffectXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_DuotoneEffectXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
end;

procedure TCT_DuotoneEffectXpgList.Assign(AItem: TCT_DuotoneEffectXpgList);
var
  i: integer;
begin
  for i := 0 to AItem.Count - 1 do
    Add.Assign(AItem[i]);
end;

procedure TCT_DuotoneEffectXpgList.CopyTo(AItem: TCT_DuotoneEffectXpgList);
begin
end;

{ TCT_GrayscaleEffect }

function  TCT_GrayscaleEffect.CheckAssigned: integer;
begin
  FAssigneds := [];
  Result := 0;
end;

procedure TCT_GrayscaleEffect.Write(AWriter: TXpgWriteXML);
begin
end;

constructor TCT_GrayscaleEffect.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 0;
end;

destructor TCT_GrayscaleEffect.Destroy;
begin
end;

procedure TCT_GrayscaleEffect.Clear;
begin
  FAssigneds := [];
end;

procedure TCT_GrayscaleEffect.Assign(AItem: TCT_GrayscaleEffect);
begin
end;

procedure TCT_GrayscaleEffect.CopyTo(AItem: TCT_GrayscaleEffect);
begin
end;

{ TCT_GrayscaleEffectXpgList }

function  TCT_GrayscaleEffectXpgList.GetItems(Index: integer): TCT_GrayscaleEffect;
begin
  Result := TCT_GrayscaleEffect(inherited Items[Index]);
end;

function  TCT_GrayscaleEffectXpgList.Add: TCT_GrayscaleEffect;
begin
  Result := TCT_GrayscaleEffect.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_GrayscaleEffectXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_GrayscaleEffectXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
end;

procedure TCT_GrayscaleEffectXpgList.Assign(AItem: TCT_GrayscaleEffectXpgList);
var
  i: integer;
begin
  for i := 0 to AItem.Count - 1 do
    Add.Assign(AItem[i]);
end;

procedure TCT_GrayscaleEffectXpgList.CopyTo(AItem: TCT_GrayscaleEffectXpgList);
begin
end;

{ TCT_HSLEffect }

function  TCT_HSLEffect.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FHue <> 0 then 
    Inc(AttrsAssigned);
  if FSat <> 0 then 
    Inc(AttrsAssigned);
  if FLum <> 0 then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_HSLEffect.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_HSLEffect.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FHue <> 0 then 
    AWriter.AddAttribute('hue',XmlIntToStr(FHue));
  if FSat <> 0 then 
    AWriter.AddAttribute('sat',XmlIntToStr(FSat));
  if FLum <> 0 then 
    AWriter.AddAttribute('lum',XmlIntToStr(FLum));
end;

procedure TCT_HSLEffect.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $00000142: FHue := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000148: FSat := XmlStrToIntDef(AAttributes.Values[i],0);
      $0000014E: FLum := XmlStrToIntDef(AAttributes.Values[i],0);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_HSLEffect.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 3;
  FHue := 0;
  FSat := 0;
  FLum := 0;
end;

destructor TCT_HSLEffect.Destroy;
begin
end;

procedure TCT_HSLEffect.Clear;
begin
  FAssigneds := [];
  FHue := 0;
  FSat := 0;
  FLum := 0;
end;

procedure TCT_HSLEffect.Assign(AItem: TCT_HSLEffect);
begin
  FHue := AItem.FHue;
  FSat := AItem.FSat;
  FLum := AItem.FLum;
end;

procedure TCT_HSLEffect.CopyTo(AItem: TCT_HSLEffect);
begin
end;

{ TCT_HSLEffectXpgList }

function  TCT_HSLEffectXpgList.GetItems(Index: integer): TCT_HSLEffect;
begin
  Result := TCT_HSLEffect(inherited Items[Index]);
end;

function  TCT_HSLEffectXpgList.Add: TCT_HSLEffect;
begin
  Result := TCT_HSLEffect.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_HSLEffectXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_HSLEffectXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_HSLEffectXpgList.Assign(AItem: TCT_HSLEffectXpgList);
var
  i: integer;
begin
  for i := 0 to AItem.Count - 1 do
    Add.Assign(AItem[i]);
end;

procedure TCT_HSLEffectXpgList.CopyTo(AItem: TCT_HSLEffectXpgList);
begin
end;

{ TCT_LuminanceEffect }

function  TCT_LuminanceEffect.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FBright <> 0 then 
    Inc(AttrsAssigned);
  if FContrast <> 0 then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_LuminanceEffect.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_LuminanceEffect.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FBright <> 0 then 
    AWriter.AddAttribute('bright',XmlIntToStr(FBright));
  if FContrast <> 0 then 
    AWriter.AddAttribute('contrast',XmlIntToStr(FContrast));
end;

procedure TCT_LuminanceEffect.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $00000280: FBright := XmlStrToIntDef(AAttributes.Values[i],0);
      $0000036E: FContrast := XmlStrToIntDef(AAttributes.Values[i],0);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_LuminanceEffect.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 2;
  FBright := 0;
  FContrast := 0;
end;

destructor TCT_LuminanceEffect.Destroy;
begin
end;

procedure TCT_LuminanceEffect.Clear;
begin
  FAssigneds := [];
  FBright := 0;
  FContrast := 0;
end;

procedure TCT_LuminanceEffect.Assign(AItem: TCT_LuminanceEffect);
begin
  FBright := AItem.FBright;
  FContrast := AItem.FContrast;
end;

procedure TCT_LuminanceEffect.CopyTo(AItem: TCT_LuminanceEffect);
begin
end;

{ TCT_LuminanceEffectXpgList }

function  TCT_LuminanceEffectXpgList.GetItems(Index: integer): TCT_LuminanceEffect;
begin
  Result := TCT_LuminanceEffect(inherited Items[Index]);
end;

function  TCT_LuminanceEffectXpgList.Add: TCT_LuminanceEffect;
begin
  Result := TCT_LuminanceEffect.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_LuminanceEffectXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_LuminanceEffectXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_LuminanceEffectXpgList.Assign(AItem: TCT_LuminanceEffectXpgList);
var
  i: integer;
begin
  for i := 0 to AItem.Count - 1 do
    Add.Assign(AItem[i]);
end;

procedure TCT_LuminanceEffectXpgList.CopyTo(AItem: TCT_LuminanceEffectXpgList);
begin
end;

{ TCT_TintEffect }

function  TCT_TintEffect.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FHue <> 0 then 
    Inc(AttrsAssigned);
  if FAmt <> 0 then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_TintEffect.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_TintEffect.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FHue <> 0 then 
    AWriter.AddAttribute('hue',XmlIntToStr(FHue));
  if FAmt <> 0 then 
    AWriter.AddAttribute('amt',XmlIntToStr(FAmt));
end;

procedure TCT_TintEffect.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashB[i] of
      $80187F2A: FHue := XmlStrToIntDef(AAttributes.Values[i],0);
      $E489E1D0: FAmt := XmlStrToIntDef(AAttributes.Values[i],0);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_TintEffect.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 2;
  FHue := 0;
  FAmt := 0;
end;

destructor TCT_TintEffect.Destroy;
begin
end;

procedure TCT_TintEffect.Clear;
begin
  FAssigneds := [];
  FHue := 0;
  FAmt := 0;
end;

procedure TCT_TintEffect.Assign(AItem: TCT_TintEffect);
begin
  FHue := AItem.FHue;
  FAmt := AItem.FAmt;
end;

procedure TCT_TintEffect.CopyTo(AItem: TCT_TintEffect);
begin
end;

{ TCT_TintEffectXpgList }

function  TCT_TintEffectXpgList.GetItems(Index: integer): TCT_TintEffect;
begin
  Result := TCT_TintEffect(inherited Items[Index]);
end;

function  TCT_TintEffectXpgList.Add: TCT_TintEffect;
begin
  Result := TCT_TintEffect.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_TintEffectXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_TintEffectXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_TintEffectXpgList.Assign(AItem: TCT_TintEffectXpgList);
var
  i: integer;
begin
  for i := 0 to AItem.Count - 1 do
    Add.Assign(AItem[i]);
end;

procedure TCT_TintEffectXpgList.CopyTo(AItem: TCT_TintEffectXpgList);
begin
end;

{ TCT_OfficeArtExtensionList }

function  TCT_OfficeArtExtensionList.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FA_EG_OfficeArtExtensionList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_OfficeArtExtensionList.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := FA_EG_OfficeArtExtensionList.HandleElement(AReader);
  if Result = Nil then 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_OfficeArtExtensionList.Write(AWriter: TXpgWriteXML);
begin
//  FA_EG_OfficeArtExtensionList.Write(AWriter);
end;

constructor TCT_OfficeArtExtensionList.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 0;
  FA_EG_OfficeArtExtensionList := TEG_OfficeArtExtensionList.Create(FOwner);
end;

destructor TCT_OfficeArtExtensionList.Destroy;
begin
  FA_EG_OfficeArtExtensionList.Free;
end;

procedure TCT_OfficeArtExtensionList.Clear;
begin
  FAssigneds := [];
  FA_EG_OfficeArtExtensionList.Clear;
end;

procedure TCT_OfficeArtExtensionList.Assign(AItem: TCT_OfficeArtExtensionList);
begin
  FA_EG_OfficeArtExtensionList.Assign(AItem.FA_EG_OfficeArtExtensionList);
end;

procedure TCT_OfficeArtExtensionList.CopyTo(AItem: TCT_OfficeArtExtensionList);
begin
end;

{ TCT_TileInfoProperties }

function  TCT_TileInfoProperties.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FTx <> 2147483632 then 
    Inc(AttrsAssigned);
  if FTy <> 2147483632 then 
    Inc(AttrsAssigned);
  if FSx <> 2147483632 then 
    Inc(AttrsAssigned);
  if FSy <> 2147483632 then 
    Inc(AttrsAssigned);
  if Integer(FFlip) <> XPG_UNKNOWN_ENUM then
    Inc(AttrsAssigned);
  if Integer(FAlgn) <> XPG_UNKNOWN_ENUM then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_TileInfoProperties.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_TileInfoProperties.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FTx <> 2147483632 then 
    AWriter.AddAttribute('tx',XmlIntToStr(FTx));
  if FTy <> 2147483632 then 
    AWriter.AddAttribute('ty',XmlIntToStr(FTy));
  if FSx <> 2147483632 then 
    AWriter.AddAttribute('sx',XmlIntToStr(FSx));
  if FSy <> 2147483632 then 
    AWriter.AddAttribute('sy',XmlIntToStr(FSy));
  if Integer(FFlip) <> XPG_UNKNOWN_ENUM then 
    AWriter.AddAttribute('flip',StrTST_TileFlipMode[Integer(FFlip)]);
  if Integer(FAlgn) <> XPG_UNKNOWN_ENUM then 
    AWriter.AddAttribute('algn',StrTST_RectAlignment[Integer(FAlgn)]);
end;

procedure TCT_TileInfoProperties.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashB[i] of
      $28AB33C4: FTx := XmlStrToIntDef(AAttributes.Values[i],0);
      $28AB33C5: FTy := XmlStrToIntDef(AAttributes.Values[i],0);
      $8BA0E615: FSx := XmlStrToIntDef(AAttributes.Values[i],0);
      $8BA0E616: FSy := XmlStrToIntDef(AAttributes.Values[i],0);
      $E517E271: FFlip := TST_TileFlipMode(StrToEnum('sttfm' + AAttributes.Values[i]));
      $1F316784: FAlgn := TST_RectAlignment(StrToEnum('stra' + AAttributes.Values[i]));
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_TileInfoProperties.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 6;
  FTx := 2147483632;
  FTy := 2147483632;
  FSx := 2147483632;
  FSy := 2147483632;
  FFlip := TST_TileFlipMode(XPG_UNKNOWN_ENUM);
  FAlgn := TST_RectAlignment(XPG_UNKNOWN_ENUM);
end;

destructor TCT_TileInfoProperties.Destroy;
begin
end;

procedure TCT_TileInfoProperties.Clear;
begin
  FAssigneds := [];
  FTx := 2147483632;
  FTy := 2147483632;
  FSx := 2147483632;
  FSy := 2147483632;
  FFlip := TST_TileFlipMode(XPG_UNKNOWN_ENUM);
  FAlgn := TST_RectAlignment(XPG_UNKNOWN_ENUM);
end;

procedure TCT_TileInfoProperties.Assign(AItem: TCT_TileInfoProperties);
begin
  FTx := AItem.FTx;
  FTy := AItem.FTy;
  FSx := AItem.FSx;
  FSy := AItem.FSy;
  FFlip := AItem.FFlip;
  FAlgn := AItem.FAlgn;
end;

procedure TCT_TileInfoProperties.CopyTo(AItem: TCT_TileInfoProperties);
begin
end;

{ TCT_StretchInfoProperties }

function  TCT_StretchInfoProperties.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FA_FillRect.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_StretchInfoProperties.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'a:fillRect' then
  begin
    Result := FA_FillRect;
  end
  else
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then
    Result.Assigneds := [xaRead];
end;

procedure TCT_StretchInfoProperties.Write(AWriter: TXpgWriteXML);
begin
  if FA_FillRect.Assigned then
  begin
    FA_FillRect.WriteAttributes(AWriter);
    AWriter.SimpleTag('a:fillRect');
  end;
end;

constructor TCT_StretchInfoProperties.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 0;
  FA_FillRect := TCT_RelativeRect.Create(FOwner);
end;

destructor TCT_StretchInfoProperties.Destroy;
begin
  FA_FillRect.Free;
end;

procedure TCT_StretchInfoProperties.Clear;
begin
  FAssigneds := [];
  FA_FillRect.Clear;
end;

procedure TCT_StretchInfoProperties.Assign(AItem: TCT_StretchInfoProperties);
begin
  if AItem.FA_FillRect <> Nil then begin
    Create_A_FillRect;
    FA_FillRect.Assign(AItem.FA_FillRect);
  end;
end;

procedure TCT_StretchInfoProperties.CopyTo(AItem: TCT_StretchInfoProperties);
begin
end;

function  TCT_StretchInfoProperties.Create_A_FillRect: TCT_RelativeRect;
begin
  if FA_FillRect = Nil then
    FA_FillRect := TCT_RelativeRect.Create(FOwner);
  Result := FA_FillRect;
end;

{ TCT_GradientStopList }

function  TCT_GradientStopList.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FA_GsXpgList <> Nil then 
    Inc(ElemsAssigned,FA_GsXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_GradientStopList.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'a:gs' then 
  begin
    Result := FA_GsXpgList.Add;
  end
  else
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then
    Result.Assigneds := [xaRead];
end;

procedure TCT_GradientStopList.Write(AWriter: TXpgWriteXML);
begin
  if FA_GsXpgList <> Nil then
    FA_GsXpgList.Write(AWriter,'a:gs');
end;

constructor TCT_GradientStopList.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 0;
  FA_GsXpgList := TCT_GradientStopXpgList.Create(FOwner);
end;

destructor TCT_GradientStopList.Destroy;
begin
  FA_GsXpgList.Free;
end;

procedure TCT_GradientStopList.Clear;
begin
  FAssigneds := [];
  FA_GsXpgList.Clear;
end;

procedure TCT_GradientStopList.Assign(AItem: TCT_GradientStopList);
begin
  if AItem.FA_GsXpgList <> Nil then begin
    Create_A_GsXpgList;
    FA_GsXpgList.Assign(AItem.FA_GsXpgList);
  end;
end;

procedure TCT_GradientStopList.CopyTo(AItem: TCT_GradientStopList);
begin
end;

function  TCT_GradientStopList.Create_A_GsXpgList: TCT_GradientStopXpgList;
begin
  if FA_GsXpgList = Nil then
    FA_GsXpgList := TCT_GradientStopXpgList.Create(FOwner);
  Result := FA_GsXpgList;
end;

{ TEG_ShadeProperties }

function  TEG_ShadeProperties.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FA_Lin <> Nil then 
    Inc(ElemsAssigned,FA_Lin.CheckAssigned);
  if FA_Path <> Nil then 
    Inc(ElemsAssigned,FA_Path.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TEG_ShadeProperties.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  case AReader.QNameHashA of
    $000001DE: begin
      if FA_Lin = Nil then 
        FA_Lin := TCT_LinearShadeProperties.Create(FOwner);
      Result := FA_Lin;
    end;
    $00000248: begin
      if FA_Path = Nil then 
        FA_Path := TCT_PathShadeProperties.Create(FOwner);
      Result := FA_Path;
    end;
    else 
    begin
      Result := Nil;
      Exit;
    end
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TEG_ShadeProperties.Write(AWriter: TXpgWriteXML);
begin
  if (FA_Lin <> Nil) and FA_Lin.Assigned then 
  begin
    FA_Lin.WriteAttributes(AWriter);
    AWriter.SimpleTag('a:lin');
  end;
  if (FA_Path <> Nil) and FA_Path.Assigned then 
  begin
    FA_Path.WriteAttributes(AWriter);
    if xaElements in FA_Path.FAssigneds then 
    begin
      AWriter.BeginTag('a:path');
      FA_Path.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:path');
  end;
end;

constructor TEG_ShadeProperties.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 2;
  FAttributeCount := 0;
end;

destructor TEG_ShadeProperties.Destroy;
begin
  if FA_Lin <> Nil then 
    FA_Lin.Free;
  if FA_Path <> Nil then 
    FA_Path.Free;
end;

procedure TEG_ShadeProperties.Clear;
begin
  FAssigneds := [];
  if FA_Lin <> Nil then 
    FreeAndNil(FA_Lin);
  if FA_Path <> Nil then 
    FreeAndNil(FA_Path);
end;

procedure TEG_ShadeProperties.Assign(AItem: TEG_ShadeProperties);
begin
end;

procedure TEG_ShadeProperties.CopyTo(AItem: TEG_ShadeProperties);
begin
end;

function  TEG_ShadeProperties.Create_Lin: TCT_LinearShadeProperties;
begin
  if FA_Lin = Nil then
    FA_Lin := TCT_LinearShadeProperties.Create(FOwner);
  Result := FA_Lin;
end;

function  TEG_ShadeProperties.Create_Path: TCT_PathShadeProperties;
begin
  if FA_Path = Nil then
    FA_Path := TCT_PathShadeProperties.Create(FOwner);
  Result := FA_Path;
end;

{ TCT_Blip }

function  TCT_Blip.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];

  if FImage <> Nil then
    Inc(AttrsAssigned);

  if FA_AlphaBiLevel <> Nil then
    Inc(ElemsAssigned,FA_AlphaBiLevel.CheckAssigned);
  if FA_AlphaCeiling <> Nil then
    Inc(ElemsAssigned,FA_AlphaCeiling.CheckAssigned);
  if FA_AlphaFloor <> Nil then
    Inc(ElemsAssigned,FA_AlphaFloor.CheckAssigned);
  if FA_AlphaInv <> Nil then
    Inc(ElemsAssigned,FA_AlphaInv.CheckAssigned);
  if FA_AlphaMod <> Nil then
    Inc(ElemsAssigned,FA_AlphaMod.CheckAssigned);
  if FA_AlphaModFix <> Nil then
    Inc(ElemsAssigned,FA_AlphaModFix.CheckAssigned);
  if FA_AlphaRepl <> Nil then
    Inc(ElemsAssigned,FA_AlphaRepl.CheckAssigned);
  if FA_BiLevel <> Nil then
    Inc(ElemsAssigned,FA_BiLevel.CheckAssigned);
  if FA_Blur <> Nil then
    Inc(ElemsAssigned,FA_Blur.CheckAssigned);
  if FA_ClrChange <> Nil then
    Inc(ElemsAssigned,FA_ClrChange.CheckAssigned);
  if FA_ClrRepl <> Nil then
    Inc(ElemsAssigned,FA_ClrRepl.CheckAssigned);
  if FA_Duotone <> Nil then
    Inc(ElemsAssigned,FA_Duotone.CheckAssigned);
  if FA_FillOverlay <> Nil then
    Inc(ElemsAssigned,FA_FillOverlay.CheckAssigned);
  if FA_Grayscl <> Nil then
    Inc(ElemsAssigned,FA_Grayscl.CheckAssigned);
  if FA_Hsl <> Nil then
    Inc(ElemsAssigned,FA_Hsl.CheckAssigned);
  if FA_Lum <> Nil then
    Inc(ElemsAssigned,FA_Lum.CheckAssigned);
  if FA_Tint <> Nil then
    Inc(ElemsAssigned,FA_Tint.CheckAssigned);
  if FA_ExtLst <> Nil then
    Inc(ElemsAssigned,FA_ExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_Blip.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $00000544: begin
      if FA_AlphaBiLevel = Nil then
        FA_AlphaBiLevel := TCT_AlphaBiLevelEffect.Create(FOwner);
      Result := FA_AlphaBiLevel;
    end;
    $0000055C: begin
      if FA_AlphaCeiling = Nil then
        FA_AlphaCeiling := TCT_AlphaCeilingEffect.Create(FOwner);
      Result := FA_AlphaCeiling;
    end;
    $000004A3: begin
      if FA_AlphaFloor = Nil then
        FA_AlphaFloor := TCT_AlphaFloorEffect.Create(FOwner);
      Result := FA_AlphaFloor;
    end;
    $000003CE: begin
      if FA_AlphaInv = Nil then
        FA_AlphaInv := TCT_AlphaInverseEffect.Create(FOwner);
      Result := FA_AlphaInv;
    end;
    $000003C1: begin
      if FA_AlphaMod = Nil then
        FA_AlphaMod := TCT_AlphaModulateEffect.Create(FOwner);
      Result := FA_AlphaMod;
    end;
    $000004E8: begin
      if FA_AlphaModFix = Nil then
        FA_AlphaModFix := TCT_AlphaModulateFixedEffect.Create(FOwner);
      Result := FA_AlphaModFix;
    end;
    $00000434: begin
      if FA_AlphaRepl = Nil then
        FA_AlphaRepl := TCT_AlphaReplaceEffect.Create(FOwner);
      Result := FA_AlphaRepl;
    end;
    $0000035E: begin
      if FA_BiLevel = Nil then
        FA_BiLevel := TCT_BiLevelEffect.Create(FOwner);
      Result := FA_BiLevel;
    end;
    $00000250: begin
      if FA_Blur = Nil then
        FA_Blur := TCT_BlurEffect.Create(FOwner);
      Result := FA_Blur;
    end;
    $00000422: begin
      if FA_ClrChange = Nil then
        FA_ClrChange := TCT_ColorChangeEffect.Create(FOwner);
      Result := FA_ClrChange;
    end;
    $0000036F: begin
      if FA_ClrRepl = Nil then
        FA_ClrRepl := TCT_ColorReplaceEffect.Create(FOwner);
      Result := FA_ClrRepl;
    end;
    $00000399: begin
      if FA_Duotone = Nil then
        FA_Duotone := TCT_DuotoneEffect.Create(FOwner);
      Result := FA_Duotone;
    end;
    $00000524: begin
      if FA_FillOverlay = Nil then
        FA_FillOverlay := TCT_FillOverlayEffect.Create(FOwner);
      Result := FA_FillOverlay;
    end;
    $00000390: begin
      if FA_Grayscl = Nil then
        FA_Grayscl := TCT_GrayscaleEffect.Create(FOwner);
      Result := FA_Grayscl;
    end;
    $000001E2: begin
      if FA_Hsl = Nil then
        FA_Hsl := TCT_HSLEffect.Create(FOwner);
      Result := FA_Hsl;
    end;
    $000001E9: begin
      if FA_Lum = Nil then
        FA_Lum := TCT_LuminanceEffect.Create(FOwner);
      Result := FA_Lum;
    end;
    $0000025A: begin
      if FA_Tint = Nil then
        FA_Tint := TCT_TintEffect.Create(FOwner);
      Result := FA_Tint;
    end;
    $0000031F: begin
      if FA_ExtLst = Nil then
        FA_ExtLst := TCT_OfficeArtExtensionList.Create(FOwner);
      Result := FA_ExtLst;
    end;
    else
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then
    Result.Assigneds := [xaRead];
end;

procedure TCT_Blip.Write(AWriter: TXpgWriteXML);
begin
  if (FA_AlphaBiLevel <> Nil) and FA_AlphaBiLevel.Assigned then begin
    FA_AlphaBiLevel.WriteAttributes(AWriter);
    AWriter.SimpleTag('a:alphaBiLevel');
  end;
  if FA_AlphaCeiling <> Nil then
    AWriter.SimpleTag('a:alphaCeiling');
  if FA_AlphaFloor <> Nil then
    AWriter.SimpleTag('a:alphaFloor');
  if (FA_AlphaInv <> Nil) and FA_AlphaInv.Assigned then begin
    AWriter.BeginTag('a:alphaInv');
    FA_AlphaInv.Write(AWriter);
    AWriter.EndTag;
  end;
  if (FA_AlphaMod <> Nil) and FA_AlphaMod.Assigned then begin
    AWriter.BeginTag('a:alphaMod');
    FA_AlphaMod.Write(AWriter);
    AWriter.EndTag;
  end;
  if (FA_AlphaModFix <> Nil) and FA_AlphaModFix.Assigned then begin
    FA_AlphaModFix.WriteAttributes(AWriter);
    AWriter.SimpleTag('a:alphaModFix');
  end;
  if (FA_AlphaRepl <> Nil) and FA_AlphaRepl.Assigned then begin
    FA_AlphaRepl.WriteAttributes(AWriter);
    AWriter.SimpleTag('a:alphaRepl');
  end;
  if (FA_BiLevel <> Nil) and FA_BiLevel.Assigned then begin
    FA_BiLevel.WriteAttributes(AWriter);
    AWriter.SimpleTag('a:biLevel');
  end;
  if (FA_Blur <> Nil) and FA_Blur.Assigned then begin
    FA_Blur.WriteAttributes(AWriter);
    AWriter.SimpleTag('a:blur');
  end;
  if (FA_ClrChange <> Nil) and FA_ClrChange.Assigned then begin
    FA_ClrChange.WriteAttributes(AWriter);
    AWriter.SimpleTag('a:clrChange');
  end;
  if (FA_ClrRepl <> Nil) and FA_ClrRepl.Assigned then begin
    AWriter.BeginTag('a:clrRepl');
    FA_ClrRepl.Write(AWriter);
    AWriter.EndTag;
  end;
  if (FA_Duotone <> Nil) and FA_Duotone.Assigned then begin
    AWriter.BeginTag('a:duotone');
    FA_Duotone.Write(AWriter);
    AWriter.EndTag;
  end;
  if (FA_FillOverlay <> Nil) and FA_FillOverlay.Assigned then begin
    FA_FillOverlay.WriteAttributes(AWriter);
    AWriter.SimpleTag('a:fillOverlay');
  end;
  if FA_Grayscl <> Nil then
    AWriter.SimpleTag('a:grayscl');
  if (FA_Hsl <> Nil) and FA_Hsl.Assigned then begin
    FA_Hsl.WriteAttributes(AWriter);
    AWriter.SimpleTag('a:hsl');
  end;
  if (FA_Lum <> Nil) and FA_Lum.Assigned then begin
    FA_Lum.WriteAttributes(AWriter);
    AWriter.SimpleTag('a:lum');
  end;
  if (FA_Tint <> Nil) and FA_Tint.Assigned then begin
    FA_Tint.WriteAttributes(AWriter);
    AWriter.SimpleTag('a:tint');
  end;
  if (FA_ExtLst <> Nil) and FA_ExtLst.Assigned then
    if xaElements in FA_ExtLst.FAssigneds then
    begin
      AWriter.BeginTag('a:extLst');
      FA_ExtLst.Write(AWriter);
      AWriter.EndTag;
    end;
end;

procedure TCT_Blip.WriteAttributes(AWriter: TXpgWriteXML);
var
  RId: AxUCString;
begin
  if FImage <> Nil then begin
    RId := FOwner.GrManager.SaveImageToOPC(FImage);

    AWriter.AddAttribute('xmlns:r',OOXML_URI_OFFICEDOC_RELATIONSHIPS);
    AWriter.AddAttribute('r:embed',RId);

    if FR_Link <> '' then
      AWriter.AddAttribute('r:link',FR_Link);
    if FCstate <> stbcNone then
      AWriter.AddAttribute('cstate',StrTST_BlipCompression[Integer(FCstate)]);
  end;
end;

procedure TCT_Blip.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do begin
    case AAttributes.HashA[i] of
      $000002A9: FR_Embed := AAttributes.Values[i];
      $0000025A: FR_Link := AAttributes.Values[i];
      $00000284: FCstate := TST_BlipCompression(StrToEnum('stbc' + AAttributes.Values[i]));
      else begin
        if AAttributes[i] <> 'xmlns:r' then
          FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
      end;
    end;
  end;
  if FR_Embed <> '' then begin
    FImage := FOwner.GrManager.Images.FindByRId(FOwner.GrManager.XLSOPC.CurrSheet.Id,FR_Embed);
  end;
end;

constructor TCT_Blip.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 18;
  FAttributeCount := 3;
  FCstate := stbcNone;
end;

destructor TCT_Blip.Destroy;
begin
  if FA_AlphaBiLevel <> Nil then
    FA_AlphaBiLevel.Free;
  if FA_AlphaCeiling <> Nil then
    FA_AlphaCeiling.Free;
  if FA_AlphaFloor <> Nil then
    FA_AlphaFloor.Free;
  if FA_AlphaInv <> Nil then
    FA_AlphaInv.Free;
  if FA_AlphaMod <> Nil then
    FA_AlphaMod.Free;
  if FA_AlphaModFix <> Nil then
    FA_AlphaModFix.Free;
  if FA_AlphaRepl <> Nil then
    FA_AlphaRepl.Free;
  if FA_BiLevel <> Nil then
    FA_BiLevel.Free;
  if FA_Blur <> Nil then
    FA_Blur.Free;
  if FA_ClrChange <> Nil then
    FA_ClrChange.Free;
  if FA_ClrRepl <> Nil then
    FA_ClrRepl.Free;
  if FA_Duotone <> Nil then
    FA_Duotone.Free;
  if FA_FillOverlay <> Nil then
    FA_FillOverlay.Free;
  if FA_Grayscl <> Nil then
    FA_Grayscl.Free;
  if FA_Hsl <> Nil then
    FA_Hsl.Free;
  if FA_Lum <> Nil then
    FA_Lum.Free;
  if FA_Tint <> Nil then
    FA_Tint.Free;
  if FA_ExtLst <> Nil then
    FA_ExtLst.Free;
end;

procedure TCT_Blip.Free_AlphaModFix;
begin
  if FA_AlphaModFix <> Nil then
    FreeAndNil(FA_AlphaModFix);
end;

procedure TCT_Blip.Free_Lum;
begin
  if FA_Lum <> Nil then
    FreeAndNil(FA_Lum);
end;

procedure TCT_Blip.Clear;
begin
  FAssigneds := [];

  if FA_AlphaBiLevel <> Nil then
    FreeAndNil(FA_AlphaBiLevel);
  if FA_AlphaCeiling <> Nil then
    FreeAndNil(FA_AlphaCeiling);
  if FA_AlphaFloor <> Nil then
    FreeAndNil(FA_AlphaFloor);
  if FA_AlphaInv <> Nil then
    FreeAndNil(FA_AlphaInv);
  if FA_AlphaMod <> Nil then
    FreeAndNil(FA_AlphaMod);
  if FA_AlphaModFix <> Nil then
    FreeAndNil(FA_AlphaModFix);
  if FA_AlphaRepl <> Nil then
    FreeAndNil(FA_AlphaRepl);
  if FA_BiLevel <> Nil then
    FreeAndNil(FA_BiLevel);
  if FA_Blur <> Nil then
    FreeAndNil(FA_Blur);
  if FA_ClrChange <> Nil then
    FreeAndNil(FA_ClrChange);
  if FA_ClrRepl <> Nil then
    FreeAndNil(FA_ClrRepl);
  if FA_Duotone <> Nil then
    FreeAndNil(FA_Duotone);
  if FA_FillOverlay <> Nil then
    FreeAndNil(FA_FillOverlay);
  if FA_Grayscl <> Nil then
    FreeAndNil(FA_Grayscl);
  if FA_Hsl <> Nil then
    FreeAndNil(FA_Hsl);
  if FA_Lum <> Nil then
    FreeAndNil(FA_Lum);
  if FA_Tint <> Nil then
    FreeAndNil(FA_Tint);
  if FA_ExtLst <> Nil then
    FreeAndNil(FA_ExtLst);
  FR_Embed := '';
  FR_Link := '';
  FCstate := stbcNone;
end;

procedure TCT_Blip.Assign(AItem: TCT_Blip);
begin
  FImage := AItem.FImage;
  FImage.IncRefCount;

  FR_Embed := AItem.FR_Embed;
  FR_Link := AItem.FR_Link;
  FCstate := AItem.FCstate;
  if AITem.FA_AlphaBiLevel <> Nil then begin
    Create_AlphaBiLevel;
    FA_AlphaBiLevel.Assign(AItem.FA_AlphaBiLevel);
  end;
  if AITem.FA_AlphaCeiling <> Nil then begin
    Create_AlphaCeiling;
    FA_AlphaCeiling.Assign(AItem.FA_AlphaCeiling);
  end;
  if AITem.FA_AlphaFloor <> Nil then begin
    Create_AlphaFloor;
    FA_AlphaFloor.Assign(AItem.FA_AlphaFloor);
  end;
  if AITem.FA_AlphaInv <> Nil then begin
    Create_AlphaInv;
    FA_AlphaInv.Assign(AItem.FA_AlphaInv);
  end;
  if AITem.FA_AlphaMod <> Nil then begin
    Create_AlphaMod;
    FA_AlphaMod.Assign(AItem.FA_AlphaMod);
  end;
  if AITem.FA_AlphaModFix <> Nil then begin
    Create_AlphaModFix;
    FA_AlphaModFix.Assign(AItem.FA_AlphaModFix);
  end;
  if AITem.FA_AlphaRepl <> Nil then begin
    Create_AlphaRepl;
    FA_AlphaRepl.Assign(AItem.FA_AlphaRepl);
  end;
  if AITem.FA_BiLevel <> Nil then begin
    Create_BiLevel;
    FA_BiLevel.Assign(AItem.FA_BiLevel);
  end;
  if AITem.FA_Blur <> Nil then begin
    Create_Blur;
    FA_Blur.Assign(AItem.FA_Blur);
  end;
  if AITem.FA_ClrChange <> Nil then begin
    Create_ClrChange;
    FA_ClrChange.Assign(AItem.FA_ClrChange);
  end;
  if AITem.FA_ClrRepl <> Nil then begin
    Create_ClrRepl;
    FA_ClrRepl.Assign(AItem.FA_ClrRepl);
  end;
  if AITem.FA_Duotone <> Nil then begin
    Create_Duotone;
    FA_Duotone.Assign(AItem.FA_Duotone);
  end;
  if AITem.FA_FillOverlay <> Nil then begin
    Create_FillOverlay;
    FA_FillOverlay.Assign(AItem.FA_FillOverlay);
  end;
  if AITem.FA_Grayscl <> Nil then begin
    Create_Grayscl;
    FA_Grayscl.Assign(AItem.FA_Grayscl);
  end;
  if AITem.FA_Hsl <> Nil then begin
    Create_Hsl;
    FA_Hsl.Assign(AItem.FA_Hsl);
  end;
  if AITem.FA_Lum <> Nil then begin
    Create_Lum;
    FA_Lum.Assign(AItem.FA_Lum);
  end;
  if AITem.FA_Tint <> Nil then begin
    Create_Tint;
    FA_Tint.Assign(AItem.FA_Tint);
  end;
  if AITem.FA_ExtLst <> Nil then begin
    Create_A_ExtLst;
    FA_ExtLst.Assign(AItem.FA_ExtLst);
  end;
end;

procedure TCT_Blip.CopyTo(AItem: TCT_Blip);
begin
end;

function  TCT_Blip.Create_AlphaBiLevel: TCT_AlphaBiLevelEffect;
begin
  if FA_AlphaBiLevel = Nil then
    FA_AlphaBiLevel := TCT_AlphaBiLevelEffect.Create(FOwner);
  Result := FA_AlphaBiLevel;
end;

function  TCT_Blip.Create_AlphaCeiling: TCT_AlphaCeilingEffect;
begin
  if FA_AlphaCeiling = Nil then
    FA_AlphaCeiling := TCT_AlphaCeilingEffect.Create(FOwner);
  Result := FA_AlphaCeiling;
end;

function  TCT_Blip.Create_AlphaFloor: TCT_AlphaFloorEffect;
begin
  if FA_AlphaFloor = Nil then
    FA_AlphaFloor := TCT_AlphaFloorEffect.Create(FOwner);
  Result := FA_AlphaFloor;
end;

function  TCT_Blip.Create_AlphaInv: TCT_AlphaInverseEffect;
begin
  if FA_AlphaInv = Nil then
    FA_AlphaInv := TCT_AlphaInverseEffect.Create(FOwner);
  Result := FA_AlphaInv;
end;

function  TCT_Blip.Create_AlphaMod: TCT_AlphaModulateEffect;
begin
  if FA_AlphaMod = Nil then
    FA_AlphaMod := TCT_AlphaModulateEffect.Create(FOwner);
  Result := FA_AlphaMod;
end;

function  TCT_Blip.Create_AlphaModFix: TCT_AlphaModulateFixedEffect;
begin
  if FA_AlphaModFix = Nil then
    FA_AlphaModFix := TCT_AlphaModulateFixedEffect.Create(FOwner);
  Result := FA_AlphaModFix;
end;

function  TCT_Blip.Create_AlphaRepl: TCT_AlphaReplaceEffect;
begin
  if FA_AlphaRepl = Nil then
    FA_AlphaRepl := TCT_AlphaReplaceEffect.Create(FOwner);
  Result := FA_AlphaRepl;
end;

function  TCT_Blip.Create_BiLevel: TCT_BiLevelEffect;
begin
  if FA_BiLevel = Nil then
    FA_BiLevel := TCT_BiLevelEffect.Create(FOwner);
  Result := FA_BiLevel;
end;

function  TCT_Blip.Create_Blur: TCT_BlurEffect;
begin
  if FA_Blur = Nil then
    FA_Blur := TCT_BlurEffect.Create(FOwner);
  Result := FA_Blur;
end;

function  TCT_Blip.Create_ClrChange: TCT_ColorChangeEffect;
begin
  if FA_ClrChange = Nil then
    FA_ClrChange := TCT_ColorChangeEffect.Create(FOwner);
  Result := FA_ClrChange;
end;

function  TCT_Blip.Create_ClrRepl: TCT_ColorReplaceEffect;
begin
  if FA_ClrRepl = Nil then
    FA_ClrRepl := TCT_ColorReplaceEffect.Create(FOwner);
  Result := FA_ClrRepl;
end;

function  TCT_Blip.Create_Duotone: TCT_DuotoneEffect;
begin
  if FA_Duotone = Nil then
    FA_Duotone := TCT_DuotoneEffect.Create(FOwner);
  Result := FA_Duotone;
end;

function  TCT_Blip.Create_FillOverlay: TCT_FillOverlayEffect;
begin
  if FA_FillOverlay = Nil then
    FA_FillOverlay := TCT_FillOverlayEffect.Create(FOwner);
  Result := FA_FillOverlay;
end;

function  TCT_Blip.Create_Grayscl: TCT_GrayscaleEffect;
begin
  if FA_Grayscl = Nil then
    FA_Grayscl := TCT_GrayscaleEffect.Create(FOwner);
  Result := FA_Grayscl;
end;

function  TCT_Blip.Create_Hsl: TCT_HSLEffect;
begin
  if FA_Hsl = Nil then
    FA_Hsl := TCT_HSLEffect.Create(FOwner);
  Result := FA_Hsl;
end;

function  TCT_Blip.Create_Lum: TCT_LuminanceEffect;
begin
  if FA_Lum = Nil then
    FA_Lum := TCT_LuminanceEffect.Create(FOwner);
  Result := FA_Lum;
end;

function  TCT_Blip.Create_Tint: TCT_TintEffect;
begin
  if FA_Tint = Nil then
    FA_Tint := TCT_TintEffect.Create(FOwner);
  Result := FA_Tint;
end;

function  TCT_Blip.Create_A_ExtLst: TCT_OfficeArtExtensionList;
begin
  if FA_ExtLst = Nil then
    FA_ExtLst := TCT_OfficeArtExtensionList.Create(FOwner);
  Result := FA_ExtLst;
end;

{ TEG_FillModeProperties }

function  TEG_FillModeProperties.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FA_Tile <> Nil then 
    Inc(ElemsAssigned,FA_Tile.CheckAssigned);
  if FA_Stretch <> Nil then 
    Inc(ElemsAssigned,FA_Stretch.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TEG_FillModeProperties.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  case AReader.QNameHashA of
    $00000249: begin
      if FA_Tile = Nil then 
        FA_Tile := TCT_TileInfoProperties.Create(FOwner);
      Result := FA_Tile;
    end;
    $00000398: begin
      if FA_Stretch = Nil then 
        FA_Stretch := TCT_StretchInfoProperties.Create(FOwner);
      Result := FA_Stretch;
    end;
    else 
    begin
      Result := Nil;
      Exit;
    end
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TEG_FillModeProperties.Write(AWriter: TXpgWriteXML);
begin
  if (FA_Tile <> Nil) and FA_Tile.Assigned then 
  begin
    FA_Tile.WriteAttributes(AWriter);
    AWriter.SimpleTag('a:tile');
  end;
  if FA_Stretch <> Nil then begin
    if FA_Stretch.Assigned and (xaElements in FA_Stretch.FAssigneds) then
    begin
      AWriter.BeginTag('a:stretch');
      FA_Stretch.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('a:stretch');
  end;
end;

constructor TEG_FillModeProperties.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 2;
  FAttributeCount := 0;
end;

destructor TEG_FillModeProperties.Destroy;
begin
  if FA_Tile <> Nil then 
    FA_Tile.Free;
  if FA_Stretch <> Nil then 
    FA_Stretch.Free;
end;

procedure TEG_FillModeProperties.Clear;
begin
  FAssigneds := [];
  if FA_Tile <> Nil then
    FreeAndNil(FA_Tile);
  if FA_Stretch <> Nil then
    FreeAndNil(FA_Stretch);
end;

procedure TEG_FillModeProperties.Assign(AItem: TEG_FillModeProperties);
begin
  if AITem.FA_Tile <> Nil then begin
    Create_Tile;
    FA_Tile.Assign(AItem.FA_Tile);
  end;
  if AItem.FA_Stretch <> Nil then begin
    Create_Stretch;
    FA_Stretch.Assign(AItem.FA_Stretch);
  end;
end;

procedure TEG_FillModeProperties.CopyTo(AItem: TEG_FillModeProperties);
begin
end;

function  TEG_FillModeProperties.Create_Tile: TCT_TileInfoProperties;
begin
  if FA_Stretch <> Nil then
    FreeAndNil(FA_Stretch);
  if FA_Tile = Nil then
    FA_Tile := TCT_TileInfoProperties.Create(FOwner);
  Result := FA_Tile;
end;

function  TEG_FillModeProperties.Create_Stretch: TCT_StretchInfoProperties;
begin
  if FA_Tile <> Nil then
    FreeAndNil(FA_Tile);
  if FA_Stretch = Nil then
    FA_Stretch := TCT_StretchInfoProperties.Create(FOwner);
  Result := FA_Stretch;
end;

{ TCT_NoFillProperties }

function  TCT_NoFillProperties.CheckAssigned: integer;
begin
  FAssigneds := [];
  Result := 0;
end;

procedure TCT_NoFillProperties.Write(AWriter: TXpgWriteXML);
begin
end;

constructor TCT_NoFillProperties.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 0;
end;

destructor TCT_NoFillProperties.Destroy;
begin
end;

procedure TCT_NoFillProperties.Clear;
begin
  FAssigneds := [];
end;

procedure TCT_NoFillProperties.Assign(AItem: TCT_NoFillProperties);
begin
end;

procedure TCT_NoFillProperties.CopyTo(AItem: TCT_NoFillProperties);
begin
end;

{ TCT_SolidColorFillProperties }

function  TCT_SolidColorFillProperties.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FA_EG_ColorChoice.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_SolidColorFillProperties.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := FA_EG_ColorChoice.HandleElement(AReader);
  if Result = Nil then 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_SolidColorFillProperties.Write(AWriter: TXpgWriteXML);
begin
  FA_EG_ColorChoice.Write(AWriter);
end;

constructor TCT_SolidColorFillProperties.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 0;
  FA_EG_ColorChoice := TEG_ColorChoice.Create(FOwner);
end;

destructor TCT_SolidColorFillProperties.Destroy;
begin
  FA_EG_ColorChoice.Free;
end;

procedure TCT_SolidColorFillProperties.Clear;
begin
  FAssigneds := [];
  FA_EG_ColorChoice.Clear;
end;

procedure TCT_SolidColorFillProperties.Assign(AItem: TCT_SolidColorFillProperties);
begin
  FA_EG_ColorChoice.Assign(AItem.FA_EG_ColorChoice);
end;

procedure TCT_SolidColorFillProperties.CopyTo(AItem: TCT_SolidColorFillProperties);
begin
end;

{ TCT_GradientFillProperties }

function  TCT_GradientFillProperties.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if Integer(FFlip) <> XPG_UNKNOWN_ENUM then
    Inc(AttrsAssigned);
  if Byte(FRotWithShape) <> 2 then
    Inc(AttrsAssigned);
  Inc(ElemsAssigned,FA_GsLst.CheckAssigned);
  Inc(ElemsAssigned,FA_EG_ShadeProperties.CheckAssigned);
  if FA_TileRect <> Nil then
    Inc(ElemsAssigned,FA_TileRect.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_GradientFillProperties.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  case AReader.QNameHashA of
    $000002A8:  Result := FA_GsLst;
    $000003D7: begin
      if FA_TileRect = Nil then
        FA_TileRect := TCT_RelativeRect.Create(FOwner);
      Result := FA_TileRect;
    end;
    else
    begin
      Result := FA_EG_ShadeProperties.HandleElement(AReader);
      if Result = Nil then
        FOwner.Errors.Error(xemUnknownElement,AReader.QName);
    end
  end;
  if Result <> Self then
    Result.Assigneds := [xaRead];
end;

procedure TCT_GradientFillProperties.Write(AWriter: TXpgWriteXML);
begin
  if (FA_GsLst <> Nil) and FA_GsLst.Assigned then
    if xaElements in FA_GsLst.FAssigneds then
    begin
      AWriter.BeginTag('a:gsLst');
      FA_GsLst.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('a:gsLst');
  FA_EG_ShadeProperties.Write(AWriter);
  if (FA_TileRect <> Nil) and FA_TileRect.Assigned then
  begin
    FA_TileRect.WriteAttributes(AWriter);
    AWriter.SimpleTag('a:tileRect');
  end;
end;

procedure TCT_GradientFillProperties.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if Integer(FFlip) <> XPG_UNKNOWN_ENUM then
    AWriter.AddAttribute('flip',StrTST_TileFlipMode[Integer(FFlip)]);
  if Byte(FRotWithShape) <> 2 then
    AWriter.AddAttribute('rotWithShape',XmlBoolToStr(FRotWithShape));
end;

procedure TCT_GradientFillProperties.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
    case AAttributes.HashA[i] of
      $000001AB: FFlip := TST_TileFlipMode(StrToEnum('sttfm' + AAttributes.Values[i]));
      $000004E2: FRotWithShape := XmlStrToBoolDef(AAttributes.Values[i],False);
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_GradientFillProperties.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 3;
  FAttributeCount := 2;
  FA_EG_ShadeProperties := TEG_ShadeProperties.Create(FOwner);
  FA_GsLst := TCT_GradientStopList.Create(FOwner);
  FFlip := TST_TileFlipMode(XPG_UNKNOWN_ENUM);
  Byte(FRotWithShape) := 2;
end;

destructor TCT_GradientFillProperties.Destroy;
begin
  FA_GsLst.Free;
  FA_EG_ShadeProperties.Free;
  if FA_TileRect <> Nil then
    FA_TileRect.Free;
end;

procedure TCT_GradientFillProperties.Free_TileRect;
begin
  if FA_TileRect <> Nil then
    FreeAndNil(FA_TileRect);
end;

procedure TCT_GradientFillProperties.Clear;
begin
  FAssigneds := [];
  FA_GsLst.Clear;
  FA_EG_ShadeProperties.Clear;
  if FA_TileRect <> Nil then
    FreeAndNil(FA_TileRect);
  FFlip := TST_TileFlipMode(XPG_UNKNOWN_ENUM);
  Byte(FRotWithShape) := 2;
end;

procedure TCT_GradientFillProperties.Assign(AItem: TCT_GradientFillProperties);
begin
  FFlip := AItem.FFlip;
  FRotWithShape := AItem.FRotWithShape;
  if AItem.FA_GsLst <> Nil then begin
    Create_GstLst;
    FA_GsLst.Assign(AItem.FA_GsLst);
  end;
  FA_EG_ShadeProperties.Assign(AItem.FA_EG_ShadeProperties);
  if AItem.FA_TileRect <> Nil then begin
    Create_TileRect;
    FA_TileRect.Assign(AItem.FA_TileRect);
  end;
end;

procedure TCT_GradientFillProperties.CopyTo(AItem: TCT_GradientFillProperties);
begin
end;

function TCT_GradientFillProperties.Create_GstLst: TCT_GradientStopList;
begin
  if FA_GsLst = Nil then
    FA_GsLst := TCT_GradientStopList.Create(FOwner);
  Result := FA_GsLst;
end;

function  TCT_GradientFillProperties.Create_TileRect: TCT_RelativeRect;
begin
  if FA_TileRect = Nil then
    FA_TileRect := TCT_RelativeRect.Create(FOwner);
  Result := FA_TileRect;
end;

{ TCT_BlipFillProperties }

function  TCT_BlipFillProperties.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FDpi <> 2147483632 then 
    Inc(AttrsAssigned);
  if Byte(FRotWithShape) <> 2 then 
    Inc(AttrsAssigned);
  if FA_Blip <> Nil then 
    Inc(ElemsAssigned,FA_Blip.CheckAssigned);
  if FA_SrcRect <> Nil then 
    Inc(ElemsAssigned,FA_SrcRect.CheckAssigned);
  Inc(ElemsAssigned,FA_EG_FillModeProperties.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_BlipFillProperties.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  case AReader.QNameHashA of
    $00000242: begin
      if FA_Blip = Nil then 
        FA_Blip := TCT_Blip.Create(FOwner);
      Result := FA_Blip;
    end;
    $00000371: begin
      if FA_SrcRect = Nil then 
        FA_SrcRect := TCT_RelativeRect.Create(FOwner);
      Result := FA_SrcRect;
    end;
    else 
    begin
      Result := FA_EG_FillModeProperties.HandleElement(AReader);
      if Result = Nil then 
        FOwner.Errors.Error(xemUnknownElement,AReader.QName);
    end
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_BlipFillProperties.Write(AWriter: TXpgWriteXML);
begin
  if (FA_Blip <> Nil) and FA_Blip.Assigned then 
  begin
    FA_Blip.WriteAttributes(AWriter);
    if xaElements in FA_Blip.FAssigneds then 
    begin
      AWriter.BeginTag('a:blip');
      FA_Blip.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:blip');
  end;
  if (FA_SrcRect <> Nil) and FA_SrcRect.Assigned then 
  begin
    FA_SrcRect.WriteAttributes(AWriter);
    AWriter.SimpleTag('a:srcRect');
  end;
  FA_EG_FillModeProperties.Write(AWriter);
end;

procedure TCT_BlipFillProperties.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FDpi <> 2147483632 then 
    AWriter.AddAttribute('dpi',XmlIntToStr(FDpi));
  if Byte(FRotWithShape) <> 2 then 
    AWriter.AddAttribute('rotWithShape',XmlBoolToStr(FRotWithShape));
end;

procedure TCT_BlipFillProperties.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $0000013D: FDpi := XmlStrToIntDef(AAttributes.Values[i],0);
      $000004E2: FRotWithShape := XmlStrToBoolDef(AAttributes.Values[i],False);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_BlipFillProperties.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 3;
  FAttributeCount := 2;
  FA_EG_FillModeProperties := TEG_FillModeProperties.Create(FOwner);
  FDpi := 2147483632;
  Byte(FRotWithShape) := 2;
end;

destructor TCT_BlipFillProperties.Destroy;
begin
  if FA_Blip <> Nil then 
    FA_Blip.Free;
  if FA_SrcRect <> Nil then 
    FA_SrcRect.Free;
  FA_EG_FillModeProperties.Free;
end;

procedure TCT_BlipFillProperties.Clear;
begin
  FAssigneds := [];
  if FA_Blip <> Nil then 
    FreeAndNil(FA_Blip);
  if FA_SrcRect <> Nil then 
    FreeAndNil(FA_SrcRect);
  FA_EG_FillModeProperties.Clear;
  FDpi := 2147483632;
  Byte(FRotWithShape) := 2;
end;

procedure TCT_BlipFillProperties.Assign(AItem: TCT_BlipFillProperties);
begin
  FDpi := AItem.FDpi;
  FRotWithShape := AItem.FRotWithShape;
  if AItem.FA_Blip <> Nil then begin
    Create_Blip;
    FA_Blip.Assign(AItem.FA_Blip);
  end;
  if AItem.FA_SrcRect <> Nil then begin
    Create_SrcRect;
    FA_SrcRect.Assign(AItem.FA_SrcRect);
  end;
  FA_EG_FillModeProperties.Assign(AItem.FA_EG_FillModeProperties);
end;

procedure TCT_BlipFillProperties.CopyTo(AItem: TCT_BlipFillProperties);
begin
end;

function  TCT_BlipFillProperties.Create_Blip: TCT_Blip;
begin
  if FA_Blip = Nil then
    FA_Blip := TCT_Blip.Create(FOwner);
  Result := FA_Blip;
end;

function  TCT_BlipFillProperties.Create_SrcRect: TCT_RelativeRect;
begin
  if FA_SrcRect = Nil then
    FA_SrcRect := TCT_RelativeRect.Create(FOwner);
  Result := FA_SrcRect;
end;

{ TCT_PatternFillProperties }

function  TCT_PatternFillProperties.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if Integer(FPrst) <> XPG_UNKNOWN_ENUM then
    Inc(AttrsAssigned);
  if FA_FgClr <> Nil then 
    Inc(ElemsAssigned,FA_FgClr.CheckAssigned);
  if FA_BgClr <> Nil then 
    Inc(ElemsAssigned,FA_BgClr.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_PatternFillProperties.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $00000289: begin
      if FA_FgClr = Nil then 
        FA_FgClr := TCT_Color.Create(FOwner);
      Result := FA_FgClr;
    end;
    $00000285: begin
      if FA_BgClr = Nil then 
        FA_BgClr := TCT_Color.Create(FOwner);
      Result := FA_BgClr;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_PatternFillProperties.Write(AWriter: TXpgWriteXML);
begin
  if (FA_FgClr <> Nil) and FA_FgClr.Assigned then 
    if xaElements in FA_FgClr.FAssigneds then 
    begin
      AWriter.BeginTag('a:fgClr');
      FA_FgClr.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:fgClr');
  if (FA_BgClr <> Nil) and FA_BgClr.Assigned then 
    if xaElements in FA_BgClr.FAssigneds then 
    begin
      AWriter.BeginTag('a:bgClr');
      FA_BgClr.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:bgClr');
end;

procedure TCT_PatternFillProperties.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if Integer(FPrst) <> XPG_UNKNOWN_ENUM then 
    AWriter.AddAttribute('prst',StrTST_PresetPatternVal[Integer(FPrst)]);
end;

procedure TCT_PatternFillProperties.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'prst' then 
    FPrst := TST_PresetPatternVal(StrToEnum('stppv' + AAttributes.Values[0]))
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_PatternFillProperties.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 2;
  FAttributeCount := 1;
  FPrst := TST_PresetPatternVal(XPG_UNKNOWN_ENUM);
end;

destructor TCT_PatternFillProperties.Destroy;
begin
  if FA_FgClr <> Nil then
    FA_FgClr.Free;
  if FA_BgClr <> Nil then
    FA_BgClr.Free;
end;

procedure TCT_PatternFillProperties.Clear;
begin
  FAssigneds := [];
  if FA_FgClr <> Nil then
    FreeAndNil(FA_FgClr);
  if FA_BgClr <> Nil then
    FreeAndNil(FA_BgClr);
  FPrst := TST_PresetPatternVal(XPG_UNKNOWN_ENUM);
end;

procedure TCT_PatternFillProperties.Assign(AItem: TCT_PatternFillProperties);
begin
  FPrst := AItem.FPrst;
  if AItem.FA_FgClr <> Nil then begin
    Create_A_FgClr;
    FA_FgClr.Assign(AItem.FA_FgClr);
  end;
  if AItem.FA_BgClr <> Nil then begin
    Create_A_BgClr;
    FA_BgClr.Assign(AItem.FA_BgClr);
  end;
end;

procedure TCT_PatternFillProperties.CopyTo(AItem: TCT_PatternFillProperties);
begin
end;

function  TCT_PatternFillProperties.Create_A_FgClr: TCT_Color;
begin
  if FA_FgClr = Nil then
    FA_FgClr := TCT_Color.Create(FOwner);
  Result := FA_FgClr;
end;

function  TCT_PatternFillProperties.Create_A_BgClr: TCT_Color;
begin
  if FA_BgClr = Nil then
    FA_BgClr := TCT_Color.Create(FOwner);
  Result := FA_BgClr;
end;

{ TCT_GroupFillProperties }

function  TCT_GroupFillProperties.CheckAssigned: integer;
begin
  FAssigneds := [];
  Result := 0;
end;

procedure TCT_GroupFillProperties.Write(AWriter: TXpgWriteXML);
begin
end;

constructor TCT_GroupFillProperties.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 0;
end;

destructor TCT_GroupFillProperties.Destroy;
begin
end;

procedure TCT_GroupFillProperties.Clear;
begin
  FAssigneds := [];
end;

procedure TCT_GroupFillProperties.Assign(AItem: TCT_GroupFillProperties);
begin
end;

procedure TCT_GroupFillProperties.CopyTo(AItem: TCT_GroupFillProperties);
begin
end;

{ TEG_FillProperties }

function  TEG_FillProperties.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FA_NoFill <> Nil then 
    Inc(ElemsAssigned);
  if FA_SolidFill <> Nil then 
    Inc(ElemsAssigned,FA_SolidFill.CheckAssigned);
  if FA_GradFill <> Nil then 
    Inc(ElemsAssigned,FA_GradFill.CheckAssigned);
  if FA_BlipFill <> Nil then 
    Inc(ElemsAssigned,FA_BlipFill.CheckAssigned);
  if FA_PattFill <> Nil then 
    Inc(ElemsAssigned,FA_PattFill.CheckAssigned);
  if FA_GrpFill <> Nil then 
    Inc(ElemsAssigned,FA_GrpFill.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TEG_FillProperties.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  case AReader.QNameHashA of
    $000002FF: begin
      if FA_NoFill = Nil then 
        FA_NoFill := TCT_NoFillProperties.Create(FOwner);
      Result := FA_NoFill;
    end;
    $0000043D: begin
      if FA_SolidFill = Nil then 
        FA_SolidFill := TCT_SolidColorFillProperties.Create(FOwner);
      Result := FA_SolidFill;
    end;
    $000003C0: begin
      if FA_GradFill = Nil then 
        FA_GradFill := TCT_GradientFillProperties.Create(FOwner);
      Result := FA_GradFill;
    end;
    $000003C9: begin
      if FA_BlipFill = Nil then 
        FA_BlipFill := TCT_BlipFillProperties.Create(FOwner);
      Result := FA_BlipFill;
    end;
    $000003DB: begin
      if FA_PattFill = Nil then 
        FA_PattFill := TCT_PatternFillProperties.Create(FOwner);
      Result := FA_PattFill;
    end;
    $0000036B: begin
      if FA_GrpFill = Nil then 
        FA_GrpFill := TCT_GroupFillProperties.Create(FOwner);
      Result := FA_GrpFill;
    end;
    else 
    begin
      Result := Nil;
      Exit;
    end
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TEG_FillProperties.Remove_NoFill;
begin
  if FA_NoFill <> Nil then begin
    FA_NoFill.Free;
    FA_NoFill := Nil;
  end;
end;

procedure TEG_FillProperties.SetAutomatic(const Value: boolean);
begin
  if Value then

end;

procedure TEG_FillProperties.Write(AWriter: TXpgWriteXML);
begin
  if (FA_NoFill <> Nil) and FA_NoFill.Assigned then 
    AWriter.SimpleTag('a:noFill');
  if (FA_SolidFill <> Nil) and FA_SolidFill.Assigned then 
    if xaElements in FA_SolidFill.FAssigneds then 
    begin
      AWriter.BeginTag('a:solidFill');
      FA_SolidFill.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:solidFill');
  if (FA_GradFill <> Nil) and FA_GradFill.Assigned then 
  begin
    FA_GradFill.WriteAttributes(AWriter);
    if xaElements in FA_GradFill.FAssigneds then 
    begin
      AWriter.BeginTag('a:gradFill');
      FA_GradFill.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:gradFill');
  end;
  if (FA_BlipFill <> Nil) and FA_BlipFill.Assigned then 
  begin
    FA_BlipFill.WriteAttributes(AWriter);
    if xaElements in FA_BlipFill.FAssigneds then 
    begin
      AWriter.BeginTag('a:blipFill');
      FA_BlipFill.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:blipFill');
  end;
  if (FA_PattFill <> Nil) and FA_PattFill.Assigned then 
  begin
    FA_PattFill.WriteAttributes(AWriter);
    if xaElements in FA_PattFill.FAssigneds then 
    begin
      AWriter.BeginTag('a:pattFill');
      FA_PattFill.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:pattFill');
  end;
  if (FA_GrpFill <> Nil) and FA_GrpFill.Assigned then 
    AWriter.SimpleTag('a:grpFill');
end;

constructor TEG_FillProperties.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 6;
  FAttributeCount := 0;
end;

destructor TEG_FillProperties.Destroy;
begin
  if FA_NoFill <> Nil then 
    FA_NoFill.Free;
  if FA_SolidFill <> Nil then 
    FA_SolidFill.Free;
  if FA_GradFill <> Nil then 
    FA_GradFill.Free;
  if FA_BlipFill <> Nil then 
    FA_BlipFill.Free;
  if FA_PattFill <> Nil then
    FA_PattFill.Free;
  if FA_GrpFill <> Nil then
    FA_GrpFill.Free;
end;

function TEG_FillProperties.GetAutomatic: boolean;
begin
  Result := (FA_NoFill = Nil) and (FA_SolidFill = Nil) and (FA_GradFill = Nil) and (FA_BlipFill = Nil) and (FA_PattFill = Nil) and (FA_GrpFill = Nil);
end;

procedure TEG_FillProperties.Clear;
begin
  FAssigneds := [];
  if FA_NoFill <> Nil then
    FreeAndNil(FA_NoFill);
  if FA_SolidFill <> Nil then
    FreeAndNil(FA_SolidFill);
  if FA_GradFill <> Nil then
    FreeAndNil(FA_GradFill);
  if FA_BlipFill <> Nil then
    FreeAndNil(FA_BlipFill);
  if FA_PattFill <> Nil then
    FreeAndNil(FA_PattFill);
  if FA_GrpFill <> Nil then
    FreeAndNil(FA_GrpFill);
end;

procedure TEG_FillProperties.Assign(AItem: TEG_FillProperties);
begin
  if AItem.FA_NoFill <> Nil then begin
    Create_NoFill;
    FA_NoFill.Assign(AItem.FA_NoFill);
  end;
  if AItem.FA_SolidFill <> Nil then begin
    Create_SolidFill;
    FA_SolidFill.Assign(AItem.FA_SolidFill);
  end;
  if AItem.FA_GradFill <> Nil then begin
    Create_GradFill;
    FA_GradFill.Assign(AItem.FA_GradFill);
  end;
  if AItem.FA_BlipFill <> Nil then begin
    Create_BlipFill;
    FA_BlipFill.Assign(AItem.FA_BlipFill);
  end;
  if AItem.FA_PattFill <> Nil then begin
    Create_PattFill;
    FA_PattFill.Assign(AItem.FA_PattFill);
  end;
  if AItem.FA_GrpFill <> Nil then begin
    Create_GrpFill;
    FA_GrpFill.Assign(AItem.FA_GrpFill);
  end;
end;

procedure TEG_FillProperties.CopyTo(AItem: TEG_FillProperties);
begin
end;

function  TEG_FillProperties.Create_NoFill: TCT_NoFillProperties;
begin
  if FA_NoFill = Nil then begin
    Clear;
    FA_NoFill := TCT_NoFillProperties.Create(FOwner);
  end;
  Result := FA_NoFill;
end;

function  TEG_FillProperties.Create_SolidFill: TCT_SolidColorFillProperties;
begin
  if FA_SolidFill = Nil then begin
    Clear;
    FA_SolidFill := TCT_SolidColorFillProperties.Create(FOwner);
  end;
  Result := FA_SolidFill;
end;

function  TEG_FillProperties.Create_GradFill: TCT_GradientFillProperties;
begin
  if FA_GradFill = Nil then begin
    Clear;
    FA_GradFill := TCT_GradientFillProperties.Create(FOwner);
  end;
  Result := FA_GradFill;
end;

function  TEG_FillProperties.Create_BlipFill: TCT_BlipFillProperties;
begin
  if FA_BlipFill = Nil then begin
    Clear;
    FA_BlipFill := TCT_BlipFillProperties.Create(FOwner);
  end;
  Result := FA_BlipFill;
end;

function  TEG_FillProperties.Create_PattFill: TCT_PatternFillProperties;
begin
  if FA_PattFill = Nil then begin
    Clear;
    FA_PattFill := TCT_PatternFillProperties.Create(FOwner);
  end;
  Result := FA_PattFill;
end;

function  TEG_FillProperties.Create_GrpFill: TCT_GroupFillProperties;
begin
  if FA_GrpFill = Nil then begin
    Clear;
    FA_GrpFill := TCT_GroupFillProperties.Create(FOwner);
  end;
  Result := FA_GrpFill;
end;

{ TCT_FillOverlayEffect }

function  TCT_FillOverlayEffect.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [xaAttributes];
  Inc(ElemsAssigned,FA_EG_FillProperties.CheckAssigned);
  Result := 1;
  if ElemsAssigned > 0 then
    FAssigneds := FAssigneds + [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_FillOverlayEffect.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := FA_EG_FillProperties.HandleElement(AReader);
  if Result = Nil then 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_FillOverlayEffect.Write(AWriter: TXpgWriteXML);
begin
  FA_EG_FillProperties.Write(AWriter);
end;

procedure TCT_FillOverlayEffect.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FBlend <> TST_BlendMode(XPG_UNKNOWN_ENUM) then
    AWriter.AddAttribute('blend',StrTST_BlendMode[Integer(FBlend)]);
end;

procedure TCT_FillOverlayEffect.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'blend' then
    FBlend := TST_BlendMode(StrToEnum('stbm' + AAttributes.Values[0]))
  else
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_FillOverlayEffect.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 1;
  FA_EG_FillProperties := TEG_FillProperties.Create(FOwner);
  FBlend := TST_BlendMode(XPG_UNKNOWN_ENUM);
end;

destructor TCT_FillOverlayEffect.Destroy;
begin
  FA_EG_FillProperties.Free;
end;

procedure TCT_FillOverlayEffect.Clear;
begin
  FAssigneds := [];
  FA_EG_FillProperties.Clear;
  FBlend := TST_BlendMode(XPG_UNKNOWN_ENUM);
end;

procedure TCT_FillOverlayEffect.Assign(AItem: TCT_FillOverlayEffect);
begin
  FBlend := AItem.FBlend;
  FA_EG_FillProperties.Assign(AItem.FA_EG_FillProperties);
end;

procedure TCT_FillOverlayEffect.CopyTo(AItem: TCT_FillOverlayEffect);
begin
end;

{ TCT_FillOverlayEffectXpgList }

function  TCT_FillOverlayEffectXpgList.GetItems(Index: integer): TCT_FillOverlayEffect;
begin
  Result := TCT_FillOverlayEffect(inherited Items[Index]);
end;

function  TCT_FillOverlayEffectXpgList.Add: TCT_FillOverlayEffect;
begin
  Result := TCT_FillOverlayEffect.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_FillOverlayEffectXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_FillOverlayEffectXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_FillOverlayEffectXpgList.Assign(AItem: TCT_FillOverlayEffectXpgList);
var
  i: integer;
begin
  for i := 0 to AItem.Count - 1 do
    Add.Assign(AItem[i]);
end;

procedure TCT_FillOverlayEffectXpgList.CopyTo(AItem: TCT_FillOverlayEffectXpgList);
begin
end;

{ TCT_EffectReference }

function  TCT_EffectReference.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FRef <> '' then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_EffectReference.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_EffectReference.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FRef <> '' then 
    AWriter.AddAttribute('ref',FRef);
end;

procedure TCT_EffectReference.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'ref' then 
    FRef := AAttributes.Values[0]
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_EffectReference.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
  FRef := '';
end;

destructor TCT_EffectReference.Destroy;
begin
end;

procedure TCT_EffectReference.Clear;
begin
  FAssigneds := [];
  FRef := '';
end;

procedure TCT_EffectReference.Assign(AItem: TCT_EffectReference);
begin
end;

procedure TCT_EffectReference.CopyTo(AItem: TCT_EffectReference);
begin
  FRef := AItem.FRef;
end;

{ TCT_EffectReferenceXpgList }

function  TCT_EffectReferenceXpgList.GetItems(Index: integer): TCT_EffectReference;
begin
  Result := TCT_EffectReference(inherited Items[Index]);
end;

function  TCT_EffectReferenceXpgList.Add: TCT_EffectReference;
begin
  Result := TCT_EffectReference.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_EffectReferenceXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_EffectReferenceXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_EffectReferenceXpgList.Assign(AItem: TCT_EffectReferenceXpgList);
var
  i: integer;
begin
  for i := 0 to AItem.Count - 1 do
    Add.Assign(AItem[i]);
end;

procedure TCT_EffectReferenceXpgList.CopyTo(AItem: TCT_EffectReferenceXpgList);
begin
end;

{ TCT_AlphaOutsetEffect }

function  TCT_AlphaOutsetEffect.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FRad <> 0 then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_AlphaOutsetEffect.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_AlphaOutsetEffect.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FRad <> 0 then 
    AWriter.AddAttribute('rad',XmlIntToStr(FRad));
end;

procedure TCT_AlphaOutsetEffect.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'rad' then 
    FRad := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_AlphaOutsetEffect.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
  FRad := 0;
end;

destructor TCT_AlphaOutsetEffect.Destroy;
begin
end;

procedure TCT_AlphaOutsetEffect.Clear;
begin
  FAssigneds := [];
  FRad := 0;
end;

procedure TCT_AlphaOutsetEffect.Assign(AItem: TCT_AlphaOutsetEffect);
begin
  FRad := AItem.FRad;
end;

procedure TCT_AlphaOutsetEffect.CopyTo(AItem: TCT_AlphaOutsetEffect);
begin
end;

{ TCT_AlphaOutsetEffectXpgList }

function  TCT_AlphaOutsetEffectXpgList.GetItems(Index: integer): TCT_AlphaOutsetEffect;
begin
  Result := TCT_AlphaOutsetEffect(inherited Items[Index]);
end;

function  TCT_AlphaOutsetEffectXpgList.Add: TCT_AlphaOutsetEffect;
begin
  Result := TCT_AlphaOutsetEffect.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_AlphaOutsetEffectXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_AlphaOutsetEffectXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_AlphaOutsetEffectXpgList.Assign(AItem: TCT_AlphaOutsetEffectXpgList);
var
  i: integer;
begin
  for i := 0 to AItem.Count - 1 do
    Add.Assign(AItem[i]);
end;

procedure TCT_AlphaOutsetEffectXpgList.CopyTo(AItem: TCT_AlphaOutsetEffectXpgList);
begin
end;

{ TCT_FillEffect }

function  TCT_FillEffect.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FA_EG_FillProperties.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_FillEffect.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := FA_EG_FillProperties.HandleElement(AReader);
  if Result = Nil then 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_FillEffect.Write(AWriter: TXpgWriteXML);
begin
  FA_EG_FillProperties.Write(AWriter);
end;

constructor TCT_FillEffect.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 0;
  FA_EG_FillProperties := TEG_FillProperties.Create(FOwner);
end;

destructor TCT_FillEffect.Destroy;
begin
  FA_EG_FillProperties.Free;
end;

procedure TCT_FillEffect.Clear;
begin
  FAssigneds := [];
  FA_EG_FillProperties.Clear;
end;

procedure TCT_FillEffect.Assign(AItem: TCT_FillEffect);
begin
  FA_EG_FillProperties.Assign(AITem.FA_EG_FillProperties);
end;

procedure TCT_FillEffect.CopyTo(AItem: TCT_FillEffect);
begin
end;

{ TCT_FillEffectXpgList }

function  TCT_FillEffectXpgList.GetItems(Index: integer): TCT_FillEffect;
begin
  Result := TCT_FillEffect(inherited Items[Index]);
end;

function  TCT_FillEffectXpgList.Add: TCT_FillEffect;
begin
  Result := TCT_FillEffect.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_FillEffectXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_FillEffectXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
end;

procedure TCT_FillEffectXpgList.Assign(AItem: TCT_FillEffectXpgList);
var
  i: integer;
begin
  for i := 0 to AItem.Count - 1 do
    Add.Assign(AItem[i]);
end;

procedure TCT_FillEffectXpgList.CopyTo(AItem: TCT_FillEffectXpgList);
begin
end;

{ TCT_GlowEffect }

function  TCT_GlowEffect.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FRad <> 0 then 
    Inc(AttrsAssigned);
  Inc(ElemsAssigned,FA_EG_ColorChoice.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_GlowEffect.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := FA_EG_ColorChoice.HandleElement(AReader);
  if Result = Nil then 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_GlowEffect.Write(AWriter: TXpgWriteXML);
begin
  FA_EG_ColorChoice.Write(AWriter);
end;

procedure TCT_GlowEffect.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FRad <> 0 then 
    AWriter.AddAttribute('rad',XmlIntToStr(FRad));
end;

procedure TCT_GlowEffect.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'rad' then 
    FRad := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_GlowEffect.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 1;
  FA_EG_ColorChoice := TEG_ColorChoice.Create(FOwner);
  FRad := 0;
end;

destructor TCT_GlowEffect.Destroy;
begin
  FA_EG_ColorChoice.Free;
end;

procedure TCT_GlowEffect.Clear;
begin
  FAssigneds := [];
  FA_EG_ColorChoice.Clear;
  FRad := 0;
end;

procedure TCT_GlowEffect.Assign(AItem: TCT_GlowEffect);
begin
end;

procedure TCT_GlowEffect.CopyTo(AItem: TCT_GlowEffect);
begin
end;

{ TCT_GlowEffectXpgList }

function  TCT_GlowEffectXpgList.GetItems(Index: integer): TCT_GlowEffect;
begin
  Result := TCT_GlowEffect(inherited Items[Index]);
end;

function  TCT_GlowEffectXpgList.Add: TCT_GlowEffect;
begin
  Result := TCT_GlowEffect.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_GlowEffectXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_GlowEffectXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_GlowEffectXpgList.Assign(AItem: TCT_GlowEffectXpgList);
var
  i: integer;
begin
  for i := 0 to AItem.Count - 1 do
    Add.Assign(AItem[i]);
end;

procedure TCT_GlowEffectXpgList.CopyTo(AItem: TCT_GlowEffectXpgList);
begin
end;

{ TCT_InnerShadowEffect }

function  TCT_InnerShadowEffect.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FBlurRad <> 0 then 
    Inc(AttrsAssigned);
  if FDist <> 0 then 
    Inc(AttrsAssigned);
  if FDir <> 0 then 
    Inc(AttrsAssigned);
  Inc(ElemsAssigned,FA_EG_ColorChoice.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_InnerShadowEffect.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := FA_EG_ColorChoice.HandleElement(AReader);
  if Result = Nil then 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_InnerShadowEffect.Write(AWriter: TXpgWriteXML);
begin
  FA_EG_ColorChoice.Write(AWriter);
end;

procedure TCT_InnerShadowEffect.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FBlurRad <> 0 then 
    AWriter.AddAttribute('blurRad',XmlIntToStr(FBlurRad));
  if FDist <> 0 then 
    AWriter.AddAttribute('dist',XmlIntToStr(FDist));
  if FDir <> 0 then 
    AWriter.AddAttribute('dir',XmlIntToStr(FDir));
end;

procedure TCT_InnerShadowEffect.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $000002CC: FBlurRad := XmlStrToIntDef(AAttributes.Values[i],0);
      $000001B4: FDist := XmlStrToIntDef(AAttributes.Values[i],0);
      $0000013F: FDir := XmlStrToIntDef(AAttributes.Values[i],0);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_InnerShadowEffect.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 3;
  FA_EG_ColorChoice := TEG_ColorChoice.Create(FOwner);
  FBlurRad := 0;
  FDist := 0;
  FDir := 0;
end;

destructor TCT_InnerShadowEffect.Destroy;
begin
  FA_EG_ColorChoice.Free;
end;

procedure TCT_InnerShadowEffect.Clear;
begin
  FAssigneds := [];
  FA_EG_ColorChoice.Clear;
  FBlurRad := 0;
  FDist := 0;
  FDir := 0;
end;

procedure TCT_InnerShadowEffect.Assign(AItem: TCT_InnerShadowEffect);
begin
  FBlurRad := AItem.FBlurRad;
  FDist := AItem.FDist;
  FDir := AItem.FDir;
  FA_EG_ColorChoice.Assign(AItem.FA_EG_ColorChoice);
end;

procedure TCT_InnerShadowEffect.CopyTo(AItem: TCT_InnerShadowEffect);
begin
end;

{ TCT_InnerShadowEffectXpgList }

function  TCT_InnerShadowEffectXpgList.GetItems(Index: integer): TCT_InnerShadowEffect;
begin
  Result := TCT_InnerShadowEffect(inherited Items[Index]);
end;

function  TCT_InnerShadowEffectXpgList.Add: TCT_InnerShadowEffect;
begin
  Result := TCT_InnerShadowEffect.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_InnerShadowEffectXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_InnerShadowEffectXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_InnerShadowEffectXpgList.Assign(AItem: TCT_InnerShadowEffectXpgList);
var
  i: integer;
begin
  for i := 0 to AItem.Count - 1 do
    Add.Assign(AItem[i]);
end;

procedure TCT_InnerShadowEffectXpgList.CopyTo(AItem: TCT_InnerShadowEffectXpgList);
begin
end;

{ TCT_OuterShadowEffect }

function  TCT_OuterShadowEffect.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FBlurRad <> 0 then 
    Inc(AttrsAssigned);
  if FDist <> 0 then 
    Inc(AttrsAssigned);
  if FDir <> 0 then 
    Inc(AttrsAssigned);
  if FSx <> 100000 then 
    Inc(AttrsAssigned);
  if FSy <> 100000 then 
    Inc(AttrsAssigned);
  if FKx <> 0 then 
    Inc(AttrsAssigned);
  if FKy <> 0 then 
    Inc(AttrsAssigned);
  if FAlgn <> straB then 
    Inc(AttrsAssigned);
  if FRotWithShape <> True then 
    Inc(AttrsAssigned);
  Inc(ElemsAssigned,FA_EG_ColorChoice.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_OuterShadowEffect.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := FA_EG_ColorChoice.HandleElement(AReader);
  if Result = Nil then 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_OuterShadowEffect.Write(AWriter: TXpgWriteXML);
begin
  FA_EG_ColorChoice.Write(AWriter);
end;

procedure TCT_OuterShadowEffect.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FBlurRad <> 0 then 
    AWriter.AddAttribute('blurRad',XmlIntToStr(FBlurRad));
  if FDist <> 0 then 
    AWriter.AddAttribute('dist',XmlIntToStr(FDist));
  if FDir <> 0 then 
    AWriter.AddAttribute('dir',XmlIntToStr(FDir));
  if FSx <> 100000 then 
    AWriter.AddAttribute('sx',XmlIntToStr(FSx));
  if FSy <> 100000 then 
    AWriter.AddAttribute('sy',XmlIntToStr(FSy));
  if FKx <> 0 then 
    AWriter.AddAttribute('kx',XmlIntToStr(FKx));
  if FKy <> 0 then 
    AWriter.AddAttribute('ky',XmlIntToStr(FKy));
  if FAlgn <> straB then 
    AWriter.AddAttribute('algn',StrTST_RectAlignment[Integer(FAlgn)]);
  if FRotWithShape <> True then 
    AWriter.AddAttribute('rotWithShape',XmlBoolToStr(FRotWithShape));
end;

procedure TCT_OuterShadowEffect.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $000002CC: FBlurRad := XmlStrToIntDef(AAttributes.Values[i],0);
      $000001B4: FDist := XmlStrToIntDef(AAttributes.Values[i],0);
      $0000013F: FDir := XmlStrToIntDef(AAttributes.Values[i],0);
      $000000EB: FSx := XmlStrToIntDef(AAttributes.Values[i],0);
      $000000EC: FSy := XmlStrToIntDef(AAttributes.Values[i],0);
      $000000E3: FKx := XmlStrToIntDef(AAttributes.Values[i],0);
      $000000E4: FKy := XmlStrToIntDef(AAttributes.Values[i],0);
      $000001A2: FAlgn := TST_RectAlignment(StrToEnum('stra' + AAttributes.Values[i]));
      $000004E2: FRotWithShape := XmlStrToBoolDef(AAttributes.Values[i],True);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_OuterShadowEffect.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 9;
  FA_EG_ColorChoice := TEG_ColorChoice.Create(FOwner);
  FBlurRad := 0;
  FDist := 0;
  FDir := 0;
  FSx := 100000;
  FSy := 100000;
  FKx := 0;
  FKy := 0;
  FAlgn := straB;
  FRotWithShape := True;
end;

destructor TCT_OuterShadowEffect.Destroy;
begin
  FA_EG_ColorChoice.Free;
end;

procedure TCT_OuterShadowEffect.Clear;
begin
  FAssigneds := [];
  FA_EG_ColorChoice.Clear;
  FBlurRad := 0;
  FDist := 0;
  FDir := 0;
  FSx := 100000;
  FSy := 100000;
  FKx := 0;
  FKy := 0;
  FAlgn := straB;
  FRotWithShape := True;
end;

procedure TCT_OuterShadowEffect.Assign(AItem: TCT_OuterShadowEffect);
begin
  FBlurRad := AItem.FBlurRad;
  FDist := AItem.FDist;
  FDir := AItem.FDir;
  FSx := AItem.FSx;
  FSy := AItem.FSy;
  FKx := AItem.FKx;
  FKy := AItem.FKy;
  FAlgn := AItem.FAlgn;
  FRotWithShape := AItem.FRotWithShape;
  FA_EG_ColorChoice.Assign(AItem.FA_EG_ColorChoice);
end;

procedure TCT_OuterShadowEffect.CopyTo(AItem: TCT_OuterShadowEffect);
begin
end;

{ TCT_OuterShadowEffectXpgList }

function  TCT_OuterShadowEffectXpgList.GetItems(Index: integer): TCT_OuterShadowEffect;
begin
  Result := TCT_OuterShadowEffect(inherited Items[Index]);
end;

function  TCT_OuterShadowEffectXpgList.Add: TCT_OuterShadowEffect;
begin
  Result := TCT_OuterShadowEffect.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_OuterShadowEffectXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_OuterShadowEffectXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_OuterShadowEffectXpgList.Assign(AItem: TCT_OuterShadowEffectXpgList);
var
  i: integer;
begin
  for i := 0 to AItem.Count - 1 do
    Add.Assign(AItem[i]);
end;

procedure TCT_OuterShadowEffectXpgList.CopyTo(AItem: TCT_OuterShadowEffectXpgList);
begin
end;

{ TCT_PresetShadowEffect }

function  TCT_PresetShadowEffect.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [xaAttributes];
  Inc(ElemsAssigned,FA_EG_ColorChoice.CheckAssigned);
  Result := 1;
  if ElemsAssigned > 0 then
    FAssigneds := FAssigneds + [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_PresetShadowEffect.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := FA_EG_ColorChoice.HandleElement(AReader);
  if Result = Nil then 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_PresetShadowEffect.Write(AWriter: TXpgWriteXML);
begin
  FA_EG_ColorChoice.Write(AWriter);
end;

procedure TCT_PresetShadowEffect.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FPrst <> TST_PresetShadowVal(XPG_UNKNOWN_ENUM) then
    AWriter.AddAttribute('prst',StrTST_PresetShadowVal[Integer(FPrst)]);
  if FDist <> 0 then
    AWriter.AddAttribute('dist',XmlIntToStr(FDist));
  if FDir <> 0 then
    AWriter.AddAttribute('dir',XmlIntToStr(FDir));
end;

procedure TCT_PresetShadowEffect.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
    case AAttributes.HashA[i] of
      $000001C9: FPrst := TST_PresetShadowVal(StrToEnum('stpsv' + AAttributes.Values[i]));
      $000001B4: FDist := XmlStrToIntDef(AAttributes.Values[i],0);
      $0000013F: FDir := XmlStrToIntDef(AAttributes.Values[i],0);
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_PresetShadowEffect.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 3;
  FA_EG_ColorChoice := TEG_ColorChoice.Create(FOwner);
  FPrst := TST_PresetShadowVal(XPG_UNKNOWN_ENUM);
  FDist := 0;
  FDir := 0;
end;

destructor TCT_PresetShadowEffect.Destroy;
begin
  FA_EG_ColorChoice.Free;
end;

procedure TCT_PresetShadowEffect.Clear;
begin
  FAssigneds := [];
  FA_EG_ColorChoice.Clear;
  FPrst := TST_PresetShadowVal(XPG_UNKNOWN_ENUM);
  FDist := 0;
  FDir := 0;
end;

procedure TCT_PresetShadowEffect.Assign(AItem: TCT_PresetShadowEffect);
begin
  FPrst := AItem.FPrst;
  FDist := AItem.FDist;
  FDir := AItem.FDir;
  FA_EG_ColorChoice.Assign(AItem.FA_EG_ColorChoice);
end;

procedure TCT_PresetShadowEffect.CopyTo(AItem: TCT_PresetShadowEffect);
begin
end;

{ TCT_PresetShadowEffectXpgList }

function  TCT_PresetShadowEffectXpgList.GetItems(Index: integer): TCT_PresetShadowEffect;
begin
  Result := TCT_PresetShadowEffect(inherited Items[Index]);
end;

function  TCT_PresetShadowEffectXpgList.Add: TCT_PresetShadowEffect;
begin
  Result := TCT_PresetShadowEffect.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_PresetShadowEffectXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_PresetShadowEffectXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_PresetShadowEffectXpgList.Assign(AItem: TCT_PresetShadowEffectXpgList);
var
  i: integer;
begin
  for i := 0 to AItem.Count - 1 do
    Add.Assign(AItem[i]);
end;

procedure TCT_PresetShadowEffectXpgList.CopyTo(AItem: TCT_PresetShadowEffectXpgList);
begin
end;

{ TCT_ReflectionEffect }

function  TCT_ReflectionEffect.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FBlurRad <> 0 then 
    Inc(AttrsAssigned);
  if FStA <> 100000 then 
    Inc(AttrsAssigned);
  if FStPos <> 0 then 
    Inc(AttrsAssigned);
  if FEndA <> 0 then 
    Inc(AttrsAssigned);
  if FEndPos <> 100000 then 
    Inc(AttrsAssigned);
  if FDist <> 0 then 
    Inc(AttrsAssigned);
  if FDir <> 0 then 
    Inc(AttrsAssigned);
  if FFadeDir <> 5400000 then 
    Inc(AttrsAssigned);
  if FSx <> 100000 then 
    Inc(AttrsAssigned);
  if FSy <> 100000 then 
    Inc(AttrsAssigned);
  if FKx <> 0 then 
    Inc(AttrsAssigned);
  if FKy <> 0 then 
    Inc(AttrsAssigned);
  if FAlgn <> straB then 
    Inc(AttrsAssigned);
  if FRotWithShape <> True then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_ReflectionEffect.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_ReflectionEffect.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FBlurRad <> 0 then 
    AWriter.AddAttribute('blurRad',XmlIntToStr(FBlurRad));
  if FStA <> 100000 then 
    AWriter.AddAttribute('stA',XmlIntToStr(FStA));
  if FStPos <> 0 then 
    AWriter.AddAttribute('stPos',XmlIntToStr(FStPos));
  if FEndA <> 0 then 
    AWriter.AddAttribute('endA',XmlIntToStr(FEndA));
  if FEndPos <> 100000 then 
    AWriter.AddAttribute('endPos',XmlIntToStr(FEndPos));
  if FDist <> 0 then 
    AWriter.AddAttribute('dist',XmlIntToStr(FDist));
  if FDir <> 0 then 
    AWriter.AddAttribute('dir',XmlIntToStr(FDir));
  if FFadeDir <> 5400000 then 
    AWriter.AddAttribute('fadeDir',XmlIntToStr(FFadeDir));
  if FSx <> 100000 then 
    AWriter.AddAttribute('sx',XmlIntToStr(FSx));
  if FSy <> 100000 then 
    AWriter.AddAttribute('sy',XmlIntToStr(FSy));
  if FKx <> 0 then 
    AWriter.AddAttribute('kx',XmlIntToStr(FKx));
  if FKy <> 0 then 
    AWriter.AddAttribute('ky',XmlIntToStr(FKy));
  if FAlgn <> straB then 
    AWriter.AddAttribute('algn',StrTST_RectAlignment[Integer(FAlgn)]);
  if FRotWithShape <> True then 
    AWriter.AddAttribute('rotWithShape',XmlBoolToStr(FRotWithShape));
end;

procedure TCT_ReflectionEffect.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $000002CC: FBlurRad := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000128: FStA := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000219: FStPos := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000178: FEndA := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000269: FEndPos := XmlStrToIntDef(AAttributes.Values[i],0);
      $000001B4: FDist := XmlStrToIntDef(AAttributes.Values[i],0);
      $0000013F: FDir := XmlStrToIntDef(AAttributes.Values[i],0);
      $000002AF: FFadeDir := XmlStrToIntDef(AAttributes.Values[i],0);
      $000000EB: FSx := XmlStrToIntDef(AAttributes.Values[i],0);
      $000000EC: FSy := XmlStrToIntDef(AAttributes.Values[i],0);
      $000000E3: FKx := XmlStrToIntDef(AAttributes.Values[i],0);
      $000000E4: FKy := XmlStrToIntDef(AAttributes.Values[i],0);
      $000001A2: FAlgn := TST_RectAlignment(StrToEnum('stra' + AAttributes.Values[i]));
      $000004E2: FRotWithShape := XmlStrToBoolDef(AAttributes.Values[i],True);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_ReflectionEffect.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 14;
  FBlurRad := 0;
  FStA := 100000;
  FStPos := 0;
  FEndA := 0;
  FEndPos := 100000;
  FDist := 0;
  FDir := 0;
  FFadeDir := 5400000;
  FSx := 100000;
  FSy := 100000;
  FKx := 0;
  FKy := 0;
  FAlgn := straB;
  FRotWithShape := True;
end;

destructor TCT_ReflectionEffect.Destroy;
begin
end;

procedure TCT_ReflectionEffect.Clear;
begin
  FAssigneds := [];
  FBlurRad := 0;
  FStA := 100000;
  FStPos := 0;
  FEndA := 0;
  FEndPos := 100000;
  FDist := 0;
  FDir := 0;
  FFadeDir := 5400000;
  FSx := 100000;
  FSy := 100000;
  FKx := 0;
  FKy := 0;
  FAlgn := straB;
  FRotWithShape := True;
end;

procedure TCT_ReflectionEffect.Assign(AItem: TCT_ReflectionEffect);
begin
  FBlurRad := AItem.FBlurRad;
  FStA := AItem.FStA;
  FStPos := AItem.FStPos;
  FEndA := AItem.FEndA;
  FEndPos := AItem.FEndPos;
  FDist := AItem.FDist;
  FDir := AItem.FDir;
  FFadeDir := AItem.FFadeDir;
  FSx := AItem.FSx;
  FSy := AItem.FSy;
  FKx := AItem.FKx;
  FKy := AItem.FKy;
  FAlgn := AItem.FAlgn;
  FRotWithShape := AItem.FRotWithShape;
end;

procedure TCT_ReflectionEffect.CopyTo(AItem: TCT_ReflectionEffect);
begin
end;

{ TCT_ReflectionEffectXpgList }

function  TCT_ReflectionEffectXpgList.GetItems(Index: integer): TCT_ReflectionEffect;
begin
  Result := TCT_ReflectionEffect(inherited Items[Index]);
end;

function  TCT_ReflectionEffectXpgList.Add: TCT_ReflectionEffect;
begin
  Result := TCT_ReflectionEffect.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_ReflectionEffectXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_ReflectionEffectXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_ReflectionEffectXpgList.Assign(AItem: TCT_ReflectionEffectXpgList);
var
  i: integer;
begin
  for i := 0 to AItem.Count - 1 do
    Add.Assign(AItem[i]);
end;

procedure TCT_ReflectionEffectXpgList.CopyTo(AItem: TCT_ReflectionEffectXpgList);
begin
end;

{ TCT_RelativeOffsetEffect }

function  TCT_RelativeOffsetEffect.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FTx <> 0 then 
    Inc(AttrsAssigned);
  if FTy <> 0 then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_RelativeOffsetEffect.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_RelativeOffsetEffect.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FTx <> 0 then 
    AWriter.AddAttribute('tx',XmlIntToStr(FTx));
  if FTy <> 0 then 
    AWriter.AddAttribute('ty',XmlIntToStr(FTy));
end;

procedure TCT_RelativeOffsetEffect.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $000000EC: FTx := XmlStrToIntDef(AAttributes.Values[i],0);
      $000000ED: FTy := XmlStrToIntDef(AAttributes.Values[i],0);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_RelativeOffsetEffect.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 2;
  FTx := 0;
  FTy := 0;
end;

destructor TCT_RelativeOffsetEffect.Destroy;
begin
end;

procedure TCT_RelativeOffsetEffect.Clear;
begin
  FAssigneds := [];
  FTx := 0;
  FTy := 0;
end;

procedure TCT_RelativeOffsetEffect.Assign(AItem: TCT_RelativeOffsetEffect);
begin
  FTx := AItem.FTx;
  FTy := AItem.FTy;
end;

procedure TCT_RelativeOffsetEffect.CopyTo(AItem: TCT_RelativeOffsetEffect);
begin
end;

{ TCT_RelativeOffsetEffectXpgList }

function  TCT_RelativeOffsetEffectXpgList.GetItems(Index: integer): TCT_RelativeOffsetEffect;
begin
  Result := TCT_RelativeOffsetEffect(inherited Items[Index]);
end;

function  TCT_RelativeOffsetEffectXpgList.Add: TCT_RelativeOffsetEffect;
begin
  Result := TCT_RelativeOffsetEffect.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_RelativeOffsetEffectXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_RelativeOffsetEffectXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_RelativeOffsetEffectXpgList.Assign(AItem: TCT_RelativeOffsetEffectXpgList);
var
  i: integer;
begin
  for i := 0 to AItem.Count - 1 do
    Add.Assign(AItem[i]);
end;

procedure TCT_RelativeOffsetEffectXpgList.CopyTo(AItem: TCT_RelativeOffsetEffectXpgList);
begin
end;

{ TCT_SoftEdgesEffect }

function  TCT_SoftEdgesEffect.CheckAssigned: integer;
begin
  FAssigneds := [xaAttributes];
  Result := 1;
end;

procedure TCT_SoftEdgesEffect.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_SoftEdgesEffect.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('rad',XmlIntToStr(FRad));
end;

procedure TCT_SoftEdgesEffect.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'rad' then 
    FRad := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_SoftEdgesEffect.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
  FRad := 2147483632;
end;

destructor TCT_SoftEdgesEffect.Destroy;
begin
end;

procedure TCT_SoftEdgesEffect.Clear;
begin
  FAssigneds := [];
  FRad := 2147483632;
end;

procedure TCT_SoftEdgesEffect.Assign(AItem: TCT_SoftEdgesEffect);
begin
  FRad := AItem.FRad;
end;

procedure TCT_SoftEdgesEffect.CopyTo(AItem: TCT_SoftEdgesEffect);
begin
end;

{ TCT_SoftEdgesEffectXpgList }

function  TCT_SoftEdgesEffectXpgList.GetItems(Index: integer): TCT_SoftEdgesEffect;
begin
  Result := TCT_SoftEdgesEffect(inherited Items[Index]);
end;

function  TCT_SoftEdgesEffectXpgList.Add: TCT_SoftEdgesEffect;
begin
  Result := TCT_SoftEdgesEffect.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_SoftEdgesEffectXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_SoftEdgesEffectXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_SoftEdgesEffectXpgList.Assign(AItem: TCT_SoftEdgesEffectXpgList);
var
  i: integer;
begin
  for i := 0 to AItem.Count - 1 do
    Add.Assign(AItem[i]);
end;

procedure TCT_SoftEdgesEffectXpgList.CopyTo(AItem: TCT_SoftEdgesEffectXpgList);
begin
end;

{ TCT_TransformEffect }

function  TCT_TransformEffect.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FSx <> 100000 then 
    Inc(AttrsAssigned);
  if FSy <> 100000 then 
    Inc(AttrsAssigned);
  if FKx <> 0 then 
    Inc(AttrsAssigned);
  if FKy <> 0 then 
    Inc(AttrsAssigned);
  if FTx <> 0 then 
    Inc(AttrsAssigned);
  if FTy <> 0 then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_TransformEffect.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_TransformEffect.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FSx <> 100000 then 
    AWriter.AddAttribute('sx',XmlIntToStr(FSx));
  if FSy <> 100000 then 
    AWriter.AddAttribute('sy',XmlIntToStr(FSy));
  if FKx <> 0 then 
    AWriter.AddAttribute('kx',XmlIntToStr(FKx));
  if FKy <> 0 then 
    AWriter.AddAttribute('ky',XmlIntToStr(FKy));
  if FTx <> 0 then 
    AWriter.AddAttribute('tx',XmlIntToStr(FTx));
  if FTy <> 0 then 
    AWriter.AddAttribute('ty',XmlIntToStr(FTy));
end;

procedure TCT_TransformEffect.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashB[i] of
      $8BA0E615: FSx := XmlStrToIntDef(AAttributes.Values[i],0);
      $8BA0E616: FSy := XmlStrToIntDef(AAttributes.Values[i],0);
      $A34E789D: FKx := XmlStrToIntDef(AAttributes.Values[i],0);
      $A34E789E: FKy := XmlStrToIntDef(AAttributes.Values[i],0);
      $28AB33C4: FTx := XmlStrToIntDef(AAttributes.Values[i],0);
      $28AB33C5: FTy := XmlStrToIntDef(AAttributes.Values[i],0);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_TransformEffect.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 6;
  FSx := 100000;
  FSy := 100000;
  FKx := 0;
  FKy := 0;
  FTx := 0;
  FTy := 0;
end;

destructor TCT_TransformEffect.Destroy;
begin
end;

procedure TCT_TransformEffect.Clear;
begin
  FAssigneds := [];
  FSx := 100000;
  FSy := 100000;
  FKx := 0;
  FKy := 0;
  FTx := 0;
  FTy := 0;
end;

procedure TCT_TransformEffect.Assign(AItem: TCT_TransformEffect);
begin
  FSx := AItem.FSx;
  FSy := AItem.FSy;
  FKx := AItem.FKx;
  FKy := AItem.FKy;
  FTx := AItem.FTx;
  FTy := AItem.FTy;
end;

procedure TCT_TransformEffect.CopyTo(AItem: TCT_TransformEffect);
begin
end;

{ TCT_TransformEffectXpgList }

function  TCT_TransformEffectXpgList.GetItems(Index: integer): TCT_TransformEffect;
begin
  Result := TCT_TransformEffect(inherited Items[Index]);
end;

function  TCT_TransformEffectXpgList.Add: TCT_TransformEffect;
begin
  Result := TCT_TransformEffect.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_TransformEffectXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_TransformEffectXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_TransformEffectXpgList.Assign(AItem: TCT_TransformEffectXpgList);
var
  i: integer;
begin
  for i := 0 to AItem.Count - 1 do
    Add.Assign(AItem[i]);
end;

procedure TCT_TransformEffectXpgList.CopyTo(AItem: TCT_TransformEffectXpgList);
begin
end;

{ TEG_Effect }

function  TEG_Effect.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FA_ContXpgList <> Nil then 
    Inc(ElemsAssigned,FA_ContXpgList.CheckAssigned);
  if FA_EffectXpgList <> Nil then 
    Inc(ElemsAssigned,FA_EffectXpgList.CheckAssigned);
  if FA_AlphaBiLevelXpgList <> Nil then 
    Inc(ElemsAssigned,FA_AlphaBiLevelXpgList.CheckAssigned);
  if FA_AlphaCeilingXpgList <> Nil then 
    Inc(ElemsAssigned,FA_AlphaCeilingXpgList.CheckAssigned);
  if FA_AlphaFloorXpgList <> Nil then 
    Inc(ElemsAssigned,FA_AlphaFloorXpgList.CheckAssigned);
  if FA_AlphaInvXpgList <> Nil then 
    Inc(ElemsAssigned,FA_AlphaInvXpgList.CheckAssigned);
  if FA_AlphaModXpgList <> Nil then 
    Inc(ElemsAssigned,FA_AlphaModXpgList.CheckAssigned);
  if FA_AlphaModFixXpgList <> Nil then 
    Inc(ElemsAssigned,FA_AlphaModFixXpgList.CheckAssigned);
  if FA_AlphaOutsetXpgList <> Nil then 
    Inc(ElemsAssigned,FA_AlphaOutsetXpgList.CheckAssigned);
  if FA_AlphaReplXpgList <> Nil then 
    Inc(ElemsAssigned,FA_AlphaReplXpgList.CheckAssigned);
  if FA_BiLevelXpgList <> Nil then 
    Inc(ElemsAssigned,FA_BiLevelXpgList.CheckAssigned);
  if FA_BlendXpgList <> Nil then 
    Inc(ElemsAssigned,FA_BlendXpgList.CheckAssigned);
  if FA_BlurXpgList <> Nil then 
    Inc(ElemsAssigned,FA_BlurXpgList.CheckAssigned);
  if FA_ClrChangeXpgList <> Nil then 
    Inc(ElemsAssigned,FA_ClrChangeXpgList.CheckAssigned);
  if FA_ClrReplXpgList <> Nil then 
    Inc(ElemsAssigned,FA_ClrReplXpgList.CheckAssigned);
  if FA_DuotoneXpgList <> Nil then 
    Inc(ElemsAssigned,FA_DuotoneXpgList.CheckAssigned);
  if FA_FillXpgList <> Nil then 
    Inc(ElemsAssigned,FA_FillXpgList.CheckAssigned);
  if FA_FillOverlayXpgList <> Nil then 
    Inc(ElemsAssigned,FA_FillOverlayXpgList.CheckAssigned);
  if FA_GlowXpgList <> Nil then 
    Inc(ElemsAssigned,FA_GlowXpgList.CheckAssigned);
  if FA_GraysclXpgList <> Nil then 
    Inc(ElemsAssigned,FA_GraysclXpgList.CheckAssigned);
  if FA_HslXpgList <> Nil then 
    Inc(ElemsAssigned,FA_HslXpgList.CheckAssigned);
  if FA_InnerShdwXpgList <> Nil then 
    Inc(ElemsAssigned,FA_InnerShdwXpgList.CheckAssigned);
  if FA_LumXpgList <> Nil then 
    Inc(ElemsAssigned,FA_LumXpgList.CheckAssigned);
  if FA_OuterShdwXpgList <> Nil then 
    Inc(ElemsAssigned,FA_OuterShdwXpgList.CheckAssigned);
  if FA_PrstShdwXpgList <> Nil then 
    Inc(ElemsAssigned,FA_PrstShdwXpgList.CheckAssigned);
  if FA_ReflectionXpgList <> Nil then 
    Inc(ElemsAssigned,FA_ReflectionXpgList.CheckAssigned);
  if FA_RelOffXpgList <> Nil then 
    Inc(ElemsAssigned,FA_RelOffXpgList.CheckAssigned);
  if FA_SoftEdgeXpgList <> Nil then 
    Inc(ElemsAssigned,FA_SoftEdgeXpgList.CheckAssigned);
  if FA_TintXpgList <> Nil then 
    Inc(ElemsAssigned,FA_TintXpgList.CheckAssigned);
  if FA_XfrmXpgList <> Nil then 
    Inc(ElemsAssigned,FA_XfrmXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TEG_Effect.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  case AReader.QNameHashA of
    $0000024F: begin
      if FA_ContXpgList = Nil then 
        FA_ContXpgList := TCT_EffectContainerXpgList.Create(FOwner);
      Result := FA_ContXpgList.Add;
    end;
    $00000308: begin
      if FA_EffectXpgList = Nil then 
        FA_EffectXpgList := TCT_EffectReferenceXpgList.Create(FOwner);
      Result := FA_EffectXpgList.Add;
    end;
    $00000544: begin
      if FA_AlphaBiLevelXpgList = Nil then 
        FA_AlphaBiLevelXpgList := TCT_AlphaBiLevelEffectXpgList.Create(FOwner);
      Result := FA_AlphaBiLevelXpgList.Add;
    end;
    $0000055C: begin
      if FA_AlphaCeilingXpgList = Nil then 
        FA_AlphaCeilingXpgList := TCT_AlphaCeilingEffectXpgList.Create(FOwner);
      Result := FA_AlphaCeilingXpgList.Add;
    end;
    $000004A3: begin
      if FA_AlphaFloorXpgList = Nil then 
        FA_AlphaFloorXpgList := TCT_AlphaFloorEffectXpgList.Create(FOwner);
      Result := FA_AlphaFloorXpgList.Add;
    end;
    $000003CE: begin
      if FA_AlphaInvXpgList = Nil then 
        FA_AlphaInvXpgList := TCT_AlphaInverseEffectXpgList.Create(FOwner);
      Result := FA_AlphaInvXpgList.Add;
    end;
    $000003C1: begin
      if FA_AlphaModXpgList = Nil then 
        FA_AlphaModXpgList := TCT_AlphaModulateEffectXpgList.Create(FOwner);
      Result := FA_AlphaModXpgList.Add;
    end;
    $000004E8: begin
      if FA_AlphaModFixXpgList = Nil then 
        FA_AlphaModFixXpgList := TCT_AlphaModulateFixedEffectXpgList.Create(FOwner);
      Result := FA_AlphaModFixXpgList.Add;
    end;
    $00000525: begin
      if FA_AlphaOutsetXpgList = Nil then 
        FA_AlphaOutsetXpgList := TCT_AlphaOutsetEffectXpgList.Create(FOwner);
      Result := FA_AlphaOutsetXpgList.Add;
    end;
    $00000434: begin
      if FA_AlphaReplXpgList = Nil then 
        FA_AlphaReplXpgList := TCT_AlphaReplaceEffectXpgList.Create(FOwner);
      Result := FA_AlphaReplXpgList.Add;
    end;
    $0000035E: begin
      if FA_BiLevelXpgList = Nil then 
        FA_BiLevelXpgList := TCT_BiLevelEffectXpgList.Create(FOwner);
      Result := FA_BiLevelXpgList.Add;
    end;
    $000002A0: begin
      if FA_BlendXpgList = Nil then 
        FA_BlendXpgList := TCT_BlendEffectXpgList.Create(FOwner);
      Result := FA_BlendXpgList.Add;
    end;
    $00000250: begin
      if FA_BlurXpgList = Nil then 
        FA_BlurXpgList := TCT_BlurEffectXpgList.Create(FOwner);
      Result := FA_BlurXpgList.Add;
    end;
    $00000422: begin
      if FA_ClrChangeXpgList = Nil then 
        FA_ClrChangeXpgList := TCT_ColorChangeEffectXpgList.Create(FOwner);
      Result := FA_ClrChangeXpgList.Add;
    end;
    $0000036F: begin
      if FA_ClrReplXpgList = Nil then 
        FA_ClrReplXpgList := TCT_ColorReplaceEffectXpgList.Create(FOwner);
      Result := FA_ClrReplXpgList.Add;
    end;
    $00000399: begin
      if FA_DuotoneXpgList = Nil then 
        FA_DuotoneXpgList := TCT_DuotoneEffectXpgList.Create(FOwner);
      Result := FA_DuotoneXpgList.Add;
    end;
    $00000242: begin
      if FA_FillXpgList = Nil then 
        FA_FillXpgList := TCT_FillEffectXpgList.Create(FOwner);
      Result := FA_FillXpgList.Add;
    end;
    $00000524: begin
      if FA_FillOverlayXpgList = Nil then 
        FA_FillOverlayXpgList := TCT_FillOverlayEffectXpgList.Create(FOwner);
      Result := FA_FillOverlayXpgList.Add;
    end;
    $00000254: begin
      if FA_GlowXpgList = Nil then 
        FA_GlowXpgList := TCT_GlowEffectXpgList.Create(FOwner);
      Result := FA_GlowXpgList.Add;
    end;
    $00000390: begin
      if FA_GraysclXpgList = Nil then 
        FA_GraysclXpgList := TCT_GrayscaleEffectXpgList.Create(FOwner);
      Result := FA_GraysclXpgList.Add;
    end;
    $000001E2: begin
      if FA_HslXpgList = Nil then 
        FA_HslXpgList := TCT_HSLEffectXpgList.Create(FOwner);
      Result := FA_HslXpgList.Add;
    end;
    $0000044D: begin
      if FA_InnerShdwXpgList = Nil then 
        FA_InnerShdwXpgList := TCT_InnerShadowEffectXpgList.Create(FOwner);
      Result := FA_InnerShdwXpgList.Add;
    end;
    $000001E9: begin
      if FA_LumXpgList = Nil then 
        FA_LumXpgList := TCT_LuminanceEffectXpgList.Create(FOwner);
      Result := FA_LumXpgList.Add;
    end;
    $00000460: begin
      if FA_OuterShdwXpgList = Nil then 
        FA_OuterShdwXpgList := TCT_OuterShadowEffectXpgList.Create(FOwner);
      Result := FA_OuterShdwXpgList.Add;
    end;
    $000003FA: begin
      if FA_PrstShdwXpgList = Nil then 
        FA_PrstShdwXpgList := TCT_PresetShadowEffectXpgList.Create(FOwner);
      Result := FA_PrstShdwXpgList.Add;
    end;
    $000004C6: begin
      if FA_ReflectionXpgList = Nil then 
        FA_ReflectionXpgList := TCT_ReflectionEffectXpgList.Create(FOwner);
      Result := FA_ReflectionXpgList.Add;
    end;
    $000002F9: begin
      if FA_RelOffXpgList = Nil then 
        FA_RelOffXpgList := TCT_RelativeOffsetEffectXpgList.Create(FOwner);
      Result := FA_RelOffXpgList.Add;
    end;
    $000003CC: begin
      if FA_SoftEdgeXpgList = Nil then 
        FA_SoftEdgeXpgList := TCT_SoftEdgesEffectXpgList.Create(FOwner);
      Result := FA_SoftEdgeXpgList.Add;
    end;
    $0000025A: begin
      if FA_TintXpgList = Nil then 
        FA_TintXpgList := TCT_TintEffectXpgList.Create(FOwner);
      Result := FA_TintXpgList.Add;
    end;
    $00000258: begin
      if FA_XfrmXpgList = Nil then 
        FA_XfrmXpgList := TCT_TransformEffectXpgList.Create(FOwner);
      Result := FA_XfrmXpgList.Add;
    end;
    else 
    begin
      Result := Nil;
      Exit;
    end
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TEG_Effect.Write(AWriter: TXpgWriteXML);
begin
  if FA_ContXpgList <> Nil then 
    FA_ContXpgList.Write(AWriter,'a:cont');
  if FA_EffectXpgList <> Nil then 
    FA_EffectXpgList.Write(AWriter,'a:effect');
  if FA_AlphaBiLevelXpgList <> Nil then 
    FA_AlphaBiLevelXpgList.Write(AWriter,'a:alphaBiLevel');
  if FA_AlphaCeilingXpgList <> Nil then 
    FA_AlphaCeilingXpgList.Write(AWriter,'a:alphaCeiling');
  if FA_AlphaFloorXpgList <> Nil then 
    FA_AlphaFloorXpgList.Write(AWriter,'a:alphaFloor');
  if FA_AlphaInvXpgList <> Nil then 
    FA_AlphaInvXpgList.Write(AWriter,'a:alphaInv');
  if FA_AlphaModXpgList <> Nil then 
    FA_AlphaModXpgList.Write(AWriter,'a:alphaMod');
  if FA_AlphaModFixXpgList <> Nil then 
    FA_AlphaModFixXpgList.Write(AWriter,'a:alphaModFix');
  if FA_AlphaOutsetXpgList <> Nil then 
    FA_AlphaOutsetXpgList.Write(AWriter,'a:alphaOutset');
  if FA_AlphaReplXpgList <> Nil then 
    FA_AlphaReplXpgList.Write(AWriter,'a:alphaRepl');
  if FA_BiLevelXpgList <> Nil then 
    FA_BiLevelXpgList.Write(AWriter,'a:biLevel');
  if FA_BlendXpgList <> Nil then 
    FA_BlendXpgList.Write(AWriter,'a:blend');
  if FA_BlurXpgList <> Nil then 
    FA_BlurXpgList.Write(AWriter,'a:blur');
  if FA_ClrChangeXpgList <> Nil then 
    FA_ClrChangeXpgList.Write(AWriter,'a:clrChange');
  if FA_ClrReplXpgList <> Nil then 
    FA_ClrReplXpgList.Write(AWriter,'a:clrRepl');
  if FA_DuotoneXpgList <> Nil then 
    FA_DuotoneXpgList.Write(AWriter,'a:duotone');
  if FA_FillXpgList <> Nil then 
    FA_FillXpgList.Write(AWriter,'a:fill');
  if FA_FillOverlayXpgList <> Nil then 
    FA_FillOverlayXpgList.Write(AWriter,'a:fillOverlay');
  if FA_GlowXpgList <> Nil then 
    FA_GlowXpgList.Write(AWriter,'a:glow');
  if FA_GraysclXpgList <> Nil then 
    FA_GraysclXpgList.Write(AWriter,'a:grayscl');
  if FA_HslXpgList <> Nil then 
    FA_HslXpgList.Write(AWriter,'a:hsl');
  if FA_InnerShdwXpgList <> Nil then 
    FA_InnerShdwXpgList.Write(AWriter,'a:innerShdw');
  if FA_LumXpgList <> Nil then 
    FA_LumXpgList.Write(AWriter,'a:lum');
  if FA_OuterShdwXpgList <> Nil then 
    FA_OuterShdwXpgList.Write(AWriter,'a:outerShdw');
  if FA_PrstShdwXpgList <> Nil then 
    FA_PrstShdwXpgList.Write(AWriter,'a:prstShdw');
  if FA_ReflectionXpgList <> Nil then 
    FA_ReflectionXpgList.Write(AWriter,'a:reflection');
  if FA_RelOffXpgList <> Nil then 
    FA_RelOffXpgList.Write(AWriter,'a:relOff');
  if FA_SoftEdgeXpgList <> Nil then 
    FA_SoftEdgeXpgList.Write(AWriter,'a:softEdge');
  if FA_TintXpgList <> Nil then 
    FA_TintXpgList.Write(AWriter,'a:tint');
  if FA_XfrmXpgList <> Nil then 
    FA_XfrmXpgList.Write(AWriter,'a:xfrm');
end;

constructor TEG_Effect.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 30;
  FAttributeCount := 0;
end;

destructor TEG_Effect.Destroy;
begin
  if FA_ContXpgList <> Nil then 
    FA_ContXpgList.Free;
  if FA_EffectXpgList <> Nil then 
    FA_EffectXpgList.Free;
  if FA_AlphaBiLevelXpgList <> Nil then 
    FA_AlphaBiLevelXpgList.Free;
  if FA_AlphaCeilingXpgList <> Nil then 
    FA_AlphaCeilingXpgList.Free;
  if FA_AlphaFloorXpgList <> Nil then 
    FA_AlphaFloorXpgList.Free;
  if FA_AlphaInvXpgList <> Nil then 
    FA_AlphaInvXpgList.Free;
  if FA_AlphaModXpgList <> Nil then 
    FA_AlphaModXpgList.Free;
  if FA_AlphaModFixXpgList <> Nil then 
    FA_AlphaModFixXpgList.Free;
  if FA_AlphaOutsetXpgList <> Nil then 
    FA_AlphaOutsetXpgList.Free;
  if FA_AlphaReplXpgList <> Nil then 
    FA_AlphaReplXpgList.Free;
  if FA_BiLevelXpgList <> Nil then 
    FA_BiLevelXpgList.Free;
  if FA_BlendXpgList <> Nil then 
    FA_BlendXpgList.Free;
  if FA_BlurXpgList <> Nil then 
    FA_BlurXpgList.Free;
  if FA_ClrChangeXpgList <> Nil then 
    FA_ClrChangeXpgList.Free;
  if FA_ClrReplXpgList <> Nil then 
    FA_ClrReplXpgList.Free;
  if FA_DuotoneXpgList <> Nil then 
    FA_DuotoneXpgList.Free;
  if FA_FillXpgList <> Nil then 
    FA_FillXpgList.Free;
  if FA_FillOverlayXpgList <> Nil then 
    FA_FillOverlayXpgList.Free;
  if FA_GlowXpgList <> Nil then 
    FA_GlowXpgList.Free;
  if FA_GraysclXpgList <> Nil then 
    FA_GraysclXpgList.Free;
  if FA_HslXpgList <> Nil then 
    FA_HslXpgList.Free;
  if FA_InnerShdwXpgList <> Nil then 
    FA_InnerShdwXpgList.Free;
  if FA_LumXpgList <> Nil then 
    FA_LumXpgList.Free;
  if FA_OuterShdwXpgList <> Nil then 
    FA_OuterShdwXpgList.Free;
  if FA_PrstShdwXpgList <> Nil then 
    FA_PrstShdwXpgList.Free;
  if FA_ReflectionXpgList <> Nil then 
    FA_ReflectionXpgList.Free;
  if FA_RelOffXpgList <> Nil then 
    FA_RelOffXpgList.Free;
  if FA_SoftEdgeXpgList <> Nil then 
    FA_SoftEdgeXpgList.Free;
  if FA_TintXpgList <> Nil then 
    FA_TintXpgList.Free;
  if FA_XfrmXpgList <> Nil then 
    FA_XfrmXpgList.Free;
end;

procedure TEG_Effect.Clear;
begin
  FAssigneds := [];
  if FA_ContXpgList <> Nil then 
    FreeAndNil(FA_ContXpgList);
  if FA_EffectXpgList <> Nil then 
    FreeAndNil(FA_EffectXpgList);
  if FA_AlphaBiLevelXpgList <> Nil then 
    FreeAndNil(FA_AlphaBiLevelXpgList);
  if FA_AlphaCeilingXpgList <> Nil then 
    FreeAndNil(FA_AlphaCeilingXpgList);
  if FA_AlphaFloorXpgList <> Nil then 
    FreeAndNil(FA_AlphaFloorXpgList);
  if FA_AlphaInvXpgList <> Nil then 
    FreeAndNil(FA_AlphaInvXpgList);
  if FA_AlphaModXpgList <> Nil then 
    FreeAndNil(FA_AlphaModXpgList);
  if FA_AlphaModFixXpgList <> Nil then 
    FreeAndNil(FA_AlphaModFixXpgList);
  if FA_AlphaOutsetXpgList <> Nil then 
    FreeAndNil(FA_AlphaOutsetXpgList);
  if FA_AlphaReplXpgList <> Nil then 
    FreeAndNil(FA_AlphaReplXpgList);
  if FA_BiLevelXpgList <> Nil then 
    FreeAndNil(FA_BiLevelXpgList);
  if FA_BlendXpgList <> Nil then 
    FreeAndNil(FA_BlendXpgList);
  if FA_BlurXpgList <> Nil then 
    FreeAndNil(FA_BlurXpgList);
  if FA_ClrChangeXpgList <> Nil then 
    FreeAndNil(FA_ClrChangeXpgList);
  if FA_ClrReplXpgList <> Nil then 
    FreeAndNil(FA_ClrReplXpgList);
  if FA_DuotoneXpgList <> Nil then 
    FreeAndNil(FA_DuotoneXpgList);
  if FA_FillXpgList <> Nil then 
    FreeAndNil(FA_FillXpgList);
  if FA_FillOverlayXpgList <> Nil then 
    FreeAndNil(FA_FillOverlayXpgList);
  if FA_GlowXpgList <> Nil then 
    FreeAndNil(FA_GlowXpgList);
  if FA_GraysclXpgList <> Nil then 
    FreeAndNil(FA_GraysclXpgList);
  if FA_HslXpgList <> Nil then 
    FreeAndNil(FA_HslXpgList);
  if FA_InnerShdwXpgList <> Nil then 
    FreeAndNil(FA_InnerShdwXpgList);
  if FA_LumXpgList <> Nil then 
    FreeAndNil(FA_LumXpgList);
  if FA_OuterShdwXpgList <> Nil then 
    FreeAndNil(FA_OuterShdwXpgList);
  if FA_PrstShdwXpgList <> Nil then 
    FreeAndNil(FA_PrstShdwXpgList);
  if FA_ReflectionXpgList <> Nil then 
    FreeAndNil(FA_ReflectionXpgList);
  if FA_RelOffXpgList <> Nil then
    FreeAndNil(FA_RelOffXpgList);
  if FA_SoftEdgeXpgList <> Nil then
    FreeAndNil(FA_SoftEdgeXpgList);
  if FA_TintXpgList <> Nil then
    FreeAndNil(FA_TintXpgList);
  if FA_XfrmXpgList <> Nil then
    FreeAndNil(FA_XfrmXpgList);
end;

procedure TEG_Effect.Assign(AItem: TEG_Effect);
begin
  if AITem.FA_ContXpgList <> Nil then begin
    Create_A_ContXpgList;
    FA_ContXpgList.Assign(AItem.FA_ContXpgList);
  end;
  if AITem.FA_EffectXpgList <> Nil then begin
    Create_A_EffectXpgList;
    FA_EffectXpgList.Assign(AItem.FA_EffectXpgList);
  end;
  if AITem.FA_AlphaBiLevelXpgList <> Nil then begin
    Create_A_AlphaBiLevelXpgList;
    FA_AlphaBiLevelXpgList.Assign(AItem.FA_AlphaBiLevelXpgList);
  end;
  if AITem.FA_AlphaCeilingXpgList <> Nil then begin
    Create_A_AlphaCeilingXpgList;
    FA_AlphaCeilingXpgList.Assign(AItem.FA_AlphaCeilingXpgList);
  end;
  if AITem.FA_AlphaFloorXpgList <> Nil then begin
    Create_A_AlphaFloorXpgList;
    FA_AlphaFloorXpgList.Assign(AItem.FA_AlphaFloorXpgList);
  end;
  if AITem.FA_AlphaInvXpgList <> Nil then begin
    Create_A_AlphaInvXpgList;
    FA_AlphaInvXpgList.Assign(AItem.FA_AlphaInvXpgList);
  end;
  if AITem.FA_AlphaModXpgList <> Nil then begin
    Create_A_AlphaModXpgList;
    FA_AlphaModXpgList.Assign(AItem.FA_AlphaModXpgList);
  end;
  if AITem.FA_AlphaModFixXpgList <> Nil then begin
    Create_A_AlphaModFixXpgList;
    FA_AlphaModFixXpgList.Assign(AItem.FA_AlphaModFixXpgList);
  end;
  if AITem.FA_AlphaOutsetXpgList <> Nil then begin
    Create_A_AlphaOutsetXpgList;
    FA_AlphaOutsetXpgList.Assign(AItem.FA_AlphaOutsetXpgList);
  end;
  if AITem.FA_AlphaReplXpgList <> Nil then begin
    Create_A_AlphaReplXpgList;
    FA_AlphaReplXpgList.Assign(AItem.FA_AlphaReplXpgList);
  end;
  if AITem.FA_BiLevelXpgList <> Nil then begin
    Create_A_BiLevelXpgList;
    FA_BiLevelXpgList.Assign(AItem.FA_BiLevelXpgList);
  end;
  if AITem.FA_BlendXpgList <> Nil then begin
    Create_A_BlendXpgList;
    FA_BlendXpgList.Assign(AItem.FA_BlendXpgList);
  end;
  if AITem.FA_BlurXpgList <> Nil then begin
    Create_A_BlurXpgList;
    FA_BlurXpgList.Assign(AItem.FA_BlurXpgList);
  end;
  if AITem.FA_ClrChangeXpgList <> Nil then begin
    Create_A_ClrChangeXpgList;
    FA_ClrChangeXpgList.Assign(AItem.FA_ClrChangeXpgList);
  end;
  if AITem.FA_ClrReplXpgList <> Nil then begin
    Create_A_ClrReplXpgList;
    FA_ClrReplXpgList.Assign(AItem.FA_ClrReplXpgList);
  end;
  if AITem.FA_DuotoneXpgList <> Nil then begin
    Create_A_DuotoneXpgList;
    FA_DuotoneXpgList.Assign(AItem.FA_DuotoneXpgList);
  end;
  if AITem.FA_FillXpgList <> Nil then begin
    Create_A_FillXpgList;
    FA_FillXpgList.Assign(AItem.FA_FillXpgList);
  end;
  if AITem.FA_FillOverlayXpgList <> Nil then begin
    Create_A_FillOverlayXpgList;
    FA_FillOverlayXpgList.Assign(AItem.FA_FillOverlayXpgList);
  end;
  if AITem.FA_GlowXpgList <> Nil then begin
    Create_A_GlowXpgList;
    FA_GlowXpgList.Assign(AItem.FA_GlowXpgList);
  end;
  if AITem.FA_GraysclXpgList <> Nil then begin
    Create_A_GraysclXpgList;
    FA_GraysclXpgList.Assign(AItem.FA_GraysclXpgList);
  end;
  if AITem.FA_HslXpgList <> Nil then begin
    Create_A_HslXpgList;
    FA_HslXpgList.Assign(AItem.FA_HslXpgList);
  end;
  if AITem.FA_InnerShdwXpgList <> Nil then begin
    Create_A_InnerShdwXpgList;
    FA_InnerShdwXpgList.Assign(AItem.FA_InnerShdwXpgList);
  end;
  if AITem.FA_LumXpgList <> Nil then begin
    Create_A_LumXpgList;
    FA_LumXpgList.Assign(AItem.FA_LumXpgList);
  end;
  if AITem.FA_OuterShdwXpgList <> Nil then begin
    Create_A_OuterShdwXpgList;
    FA_OuterShdwXpgList.Assign(AItem.FA_OuterShdwXpgList);
  end;
  if AITem.FA_PrstShdwXpgList <> Nil then begin
    Create_A_PrstShdwXpgList;
    FA_PrstShdwXpgList.Assign(AItem.FA_PrstShdwXpgList);
  end;
  if AITem.FA_ReflectionXpgList <> Nil then begin
    Create_A_ReflectionXpgList;
    FA_ReflectionXpgList.Assign(AItem.FA_ReflectionXpgList);
  end;
  if AITem.FA_RelOffXpgList <> Nil then begin
    Create_A_RelOffXpgList;
    FA_RelOffXpgList.Assign(AItem.FA_RelOffXpgList);
  end;
  if AITem.FA_SoftEdgeXpgList <> Nil then begin
    Create_A_SoftEdgeXpgList;
    FA_SoftEdgeXpgList.Assign(AItem.FA_SoftEdgeXpgList);
  end;
  if AITem.FA_TintXpgList <> Nil then begin
    Create_A_TintXpgList;
    FA_TintXpgList.Assign(AItem.FA_TintXpgList);
  end;
  if AITem.FA_XfrmXpgList <> Nil then begin
    Create_A_XfrmXpgList;
    FA_XfrmXpgList.Assign(AItem.FA_XfrmXpgList);
  end;
end;

procedure TEG_Effect.CopyTo(AItem: TEG_Effect);
begin
end;

function  TEG_Effect.Create_A_ContXpgList: TCT_EffectContainerXpgList;
begin
  if FA_ContXpgList = Nil then
    FA_ContXpgList := TCT_EffectContainerXpgList.Create(FOwner);
  Result := FA_ContXpgList;
end;

function  TEG_Effect.Create_A_EffectXpgList: TCT_EffectReferenceXpgList;
begin
  if FA_EffectXpgList = Nil then
    FA_EffectXpgList := TCT_EffectReferenceXpgList.Create(FOwner);
  Result := FA_EffectXpgList;
end;

function  TEG_Effect.Create_A_AlphaBiLevelXpgList: TCT_AlphaBiLevelEffectXpgList;
begin
  if FA_AlphaBiLevelXpgList = Nil then
    FA_AlphaBiLevelXpgList := TCT_AlphaBiLevelEffectXpgList.Create(FOwner);
  Result := FA_AlphaBiLevelXpgList;
end;

function  TEG_Effect.Create_A_AlphaCeilingXpgList: TCT_AlphaCeilingEffectXpgList;
begin
  if FA_AlphaCeilingXpgList = Nil then
    FA_AlphaCeilingXpgList := TCT_AlphaCeilingEffectXpgList.Create(FOwner);
  Result := FA_AlphaCeilingXpgList;
end;

function  TEG_Effect.Create_A_AlphaFloorXpgList: TCT_AlphaFloorEffectXpgList;
begin
  if FA_AlphaFloorXpgList = Nil then
    FA_AlphaFloorXpgList := TCT_AlphaFloorEffectXpgList.Create(FOwner);
  Result := FA_AlphaFloorXpgList;
end;

function  TEG_Effect.Create_A_AlphaInvXpgList: TCT_AlphaInverseEffectXpgList;
begin
  if FA_AlphaInvXpgList = Nil then
    FA_AlphaInvXpgList := TCT_AlphaInverseEffectXpgList.Create(FOwner);
  Result := FA_AlphaInvXpgList;
end;

function  TEG_Effect.Create_A_AlphaModXpgList: TCT_AlphaModulateEffectXpgList;
begin
  if FA_AlphaModXpgList = Nil then
    FA_AlphaModXpgList := TCT_AlphaModulateEffectXpgList.Create(FOwner);
  Result := FA_AlphaModXpgList;
end;

function  TEG_Effect.Create_A_AlphaModFixXpgList: TCT_AlphaModulateFixedEffectXpgList;
begin
  if FA_AlphaModFixXpgList = Nil then
    FA_AlphaModFixXpgList := TCT_AlphaModulateFixedEffectXpgList.Create(FOwner);
  Result := FA_AlphaModFixXpgList;
end;

function  TEG_Effect.Create_A_AlphaOutsetXpgList: TCT_AlphaOutsetEffectXpgList;
begin
  if FA_AlphaOutsetXpgList = Nil then
    FA_AlphaOutsetXpgList := TCT_AlphaOutsetEffectXpgList.Create(FOwner);
  Result := FA_AlphaOutsetXpgList;
end;

function  TEG_Effect.Create_A_AlphaReplXpgList: TCT_AlphaReplaceEffectXpgList;
begin
  if FA_AlphaReplXpgList = Nil then
    FA_AlphaReplXpgList := TCT_AlphaReplaceEffectXpgList.Create(FOwner);
  Result := FA_AlphaReplXpgList;
end;

function  TEG_Effect.Create_A_BiLevelXpgList: TCT_BiLevelEffectXpgList;
begin
  if FA_BiLevelXpgList = Nil then
    FA_BiLevelXpgList := TCT_BiLevelEffectXpgList.Create(FOwner);
  Result := FA_BiLevelXpgList;
end;

function  TEG_Effect.Create_A_BlendXpgList: TCT_BlendEffectXpgList;
begin
  if FA_BlendXpgList = Nil then
    FA_BlendXpgList := TCT_BlendEffectXpgList.Create(FOwner);
  Result := FA_BlendXpgList;
end;

function  TEG_Effect.Create_A_BlurXpgList: TCT_BlurEffectXpgList;
begin
  if FA_BlurXpgList = Nil then
    FA_BlurXpgList := TCT_BlurEffectXpgList.Create(FOwner);
  Result := FA_BlurXpgList;
end;

function  TEG_Effect.Create_A_ClrChangeXpgList: TCT_ColorChangeEffectXpgList;
begin
  if FA_ClrChangeXpgList = Nil then
    FA_ClrChangeXpgList := TCT_ColorChangeEffectXpgList.Create(FOwner);
  Result := FA_ClrChangeXpgList;
end;

function  TEG_Effect.Create_A_ClrReplXpgList: TCT_ColorReplaceEffectXpgList;
begin
  if FA_ClrReplXpgList = Nil then
    FA_ClrReplXpgList := TCT_ColorReplaceEffectXpgList.Create(FOwner);
  Result := FA_ClrReplXpgList;
end;

function  TEG_Effect.Create_A_DuotoneXpgList: TCT_DuotoneEffectXpgList;
begin
  if FA_DuotoneXpgList = Nil then
    FA_DuotoneXpgList := TCT_DuotoneEffectXpgList.Create(FOwner);
  Result := FA_DuotoneXpgList;
end;

function  TEG_Effect.Create_A_FillXpgList: TCT_FillEffectXpgList;
begin
  if FA_FillXpgList = Nil then
    FA_FillXpgList := TCT_FillEffectXpgList.Create(FOwner);
  Result := FA_FillXpgList;
end;

function  TEG_Effect.Create_A_FillOverlayXpgList: TCT_FillOverlayEffectXpgList;
begin
  if FA_FillOverlayXpgList = Nil then
    FA_FillOverlayXpgList := TCT_FillOverlayEffectXpgList.Create(FOwner);
  Result := FA_FillOverlayXpgList;
end;

function  TEG_Effect.Create_A_GlowXpgList: TCT_GlowEffectXpgList;
begin
  if FA_GlowXpgList = Nil then
    FA_GlowXpgList := TCT_GlowEffectXpgList.Create(FOwner);
  Result := FA_GlowXpgList;
end;

function  TEG_Effect.Create_A_GraysclXpgList: TCT_GrayscaleEffectXpgList;
begin
  if FA_GraysclXpgList = Nil then
    FA_GraysclXpgList := TCT_GrayscaleEffectXpgList.Create(FOwner);
  Result := FA_GraysclXpgList;
end;

function  TEG_Effect.Create_A_HslXpgList: TCT_HSLEffectXpgList;
begin
  if FA_HslXpgList = Nil then
    FA_HslXpgList := TCT_HSLEffectXpgList.Create(FOwner);
  Result := FA_HslXpgList;
end;

function  TEG_Effect.Create_A_InnerShdwXpgList: TCT_InnerShadowEffectXpgList;
begin
  if FA_InnerShdwXpgList = Nil then
    FA_InnerShdwXpgList := TCT_InnerShadowEffectXpgList.Create(FOwner);
  Result := FA_InnerShdwXpgList;
end;

function  TEG_Effect.Create_A_LumXpgList: TCT_LuminanceEffectXpgList;
begin
  if FA_LumXpgList = Nil then
    FA_LumXpgList := TCT_LuminanceEffectXpgList.Create(FOwner);
  Result := FA_LumXpgList;
end;

function  TEG_Effect.Create_A_OuterShdwXpgList: TCT_OuterShadowEffectXpgList;
begin
  if FA_OuterShdwXpgList = Nil then
    FA_OuterShdwXpgList := TCT_OuterShadowEffectXpgList.Create(FOwner);
  Result := FA_OuterShdwXpgList;
end;

function  TEG_Effect.Create_A_PrstShdwXpgList: TCT_PresetShadowEffectXpgList;
begin
  if FA_PrstShdwXpgList = Nil then
    FA_PrstShdwXpgList := TCT_PresetShadowEffectXpgList.Create(FOwner);
  Result := FA_PrstShdwXpgList;
end;

function  TEG_Effect.Create_A_ReflectionXpgList: TCT_ReflectionEffectXpgList;
begin
  if FA_ReflectionXpgList = Nil then
    FA_ReflectionXpgList := TCT_ReflectionEffectXpgList.Create(FOwner);
  Result := FA_ReflectionXpgList;
end;

function  TEG_Effect.Create_A_RelOffXpgList: TCT_RelativeOffsetEffectXpgList;
begin
  if FA_RelOffXpgList = Nil then
    FA_RelOffXpgList := TCT_RelativeOffsetEffectXpgList.Create(FOwner);
  Result := FA_RelOffXpgList;
end;

function  TEG_Effect.Create_A_SoftEdgeXpgList: TCT_SoftEdgesEffectXpgList;
begin
  if FA_SoftEdgeXpgList = Nil then
    FA_SoftEdgeXpgList := TCT_SoftEdgesEffectXpgList.Create(FOwner);
  Result := FA_SoftEdgeXpgList;
end;

function  TEG_Effect.Create_A_TintXpgList: TCT_TintEffectXpgList;
begin
  if FA_TintXpgList = Nil then
    FA_TintXpgList := TCT_TintEffectXpgList.Create(FOwner);
  Result := FA_TintXpgList;
end;

function  TEG_Effect.Create_A_XfrmXpgList: TCT_TransformEffectXpgList;
begin
  if FA_XfrmXpgList = Nil then
    FA_XfrmXpgList := TCT_TransformEffectXpgList.Create(FOwner);
  Result := FA_XfrmXpgList;
end;

{ TCT_EffectContainer }

function  TCT_EffectContainer.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FType <> stectSib then 
    Inc(AttrsAssigned);
  if FName <> '' then 
    Inc(AttrsAssigned);
  Inc(ElemsAssigned,FA_EG_Effect.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_EffectContainer.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := FA_EG_Effect.HandleElement(AReader);
  if Result = Nil then 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_EffectContainer.Write(AWriter: TXpgWriteXML);
begin
  FA_EG_Effect.Write(AWriter);
end;

procedure TCT_EffectContainer.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FType <> stectSib then 
    AWriter.AddAttribute('type',StrTST_EffectContainerType[Integer(FType)]);
  if FName <> '' then 
    AWriter.AddAttribute('name',FName);
end;

procedure TCT_EffectContainer.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $000001C2: FType := TST_EffectContainerType(StrToEnum('stect' + AAttributes.Values[i]));
      $000001A1: FName := AAttributes.Values[i];
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_EffectContainer.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 2;
  FA_EG_Effect := TEG_Effect.Create(FOwner);
  FType := stectSib;
  FName := '';
end;

destructor TCT_EffectContainer.Destroy;
begin
  FA_EG_Effect.Free;
end;

procedure TCT_EffectContainer.Clear;
begin
  FAssigneds := [];
  FA_EG_Effect.Clear;
  FType := stectSib;
  FName := '';
end;

procedure TCT_EffectContainer.Assign(AItem: TCT_EffectContainer);
begin
  FType := AItem.FType;
  FName := AItem.FName;
  FA_EG_Effect.Assign(AItem.FA_EG_Effect);
end;

procedure TCT_EffectContainer.CopyTo(AItem: TCT_EffectContainer);
begin
end;

{ TCT_EffectContainerXpgList }

function  TCT_EffectContainerXpgList.GetItems(Index: integer): TCT_EffectContainer;
begin
  Result := TCT_EffectContainer(inherited Items[Index]);
end;

function  TCT_EffectContainerXpgList.Add: TCT_EffectContainer;
begin
  Result := TCT_EffectContainer.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_EffectContainerXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_EffectContainerXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_EffectContainerXpgList.Assign(AItem: TCT_EffectContainerXpgList);
var
  i: integer;
begin
  for i := 0 to AItem.Count - 1 do
    Add.Assign(AItem[i]);
end;

procedure TCT_EffectContainerXpgList.CopyTo(AItem: TCT_EffectContainerXpgList);
begin
end;

{ TCT_BlendEffect }

function  TCT_BlendEffect.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [xaAttributes];
  if FA_Cont <> Nil then
    Inc(ElemsAssigned,FA_Cont.CheckAssigned);
  Result := 1;
  if ElemsAssigned > 0 then
    FAssigneds := FAssigneds + [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_BlendEffect.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'a:cont' then 
  begin
    if FA_Cont = Nil then 
      FA_Cont := TCT_EffectContainer.Create(FOwner);
    Result := FA_Cont;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_BlendEffect.Write(AWriter: TXpgWriteXML);
begin
  if (FA_Cont <> Nil) and FA_Cont.Assigned then 
  begin
    FA_Cont.WriteAttributes(AWriter);
    if xaElements in FA_Cont.FAssigneds then 
    begin
      AWriter.BeginTag('a:cont');
      FA_Cont.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:cont');
  end
  else 
    AWriter.SimpleTag('a:cont');
end;

procedure TCT_BlendEffect.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FBlend <> TST_BlendMode(XPG_UNKNOWN_ENUM) then
    AWriter.AddAttribute('blend',StrTST_BlendMode[Integer(FBlend)]);
end;

procedure TCT_BlendEffect.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'blend' then
    FBlend := TST_BlendMode(StrToEnum('stbm' + AAttributes.Values[0]))
  else
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_BlendEffect.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 1;
  FBlend := TST_BlendMode(XPG_UNKNOWN_ENUM);
end;

destructor TCT_BlendEffect.Destroy;
begin
  if FA_Cont <> Nil then
    FA_Cont.Free;
end;

procedure TCT_BlendEffect.Clear;
begin
  FAssigneds := [];
  if FA_Cont <> Nil then
    FreeAndNil(FA_Cont);
  FBlend := TST_BlendMode(XPG_UNKNOWN_ENUM);
end;

procedure TCT_BlendEffect.Assign(AItem: TCT_BlendEffect);
begin
  FBlend := AItem.FBlend;
  FA_Cont.Assign(AItem.FA_Cont);
end;

procedure TCT_BlendEffect.CopyTo(AItem: TCT_BlendEffect);
begin
end;

function  TCT_BlendEffect.Create_A_Cont: TCT_EffectContainer;
begin
  if FA_Cont = Nil then
    FA_Cont := TCT_EffectContainer.Create(FOwner);
  Result := FA_Cont;
end;

{ TCT_BlendEffectXpgList }

function  TCT_BlendEffectXpgList.GetItems(Index: integer): TCT_BlendEffect;
begin
  Result := TCT_BlendEffect(inherited Items[Index]);
end;

function  TCT_BlendEffectXpgList.Add: TCT_BlendEffect;
begin
  Result := TCT_BlendEffect.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_BlendEffectXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_BlendEffectXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_BlendEffectXpgList.Assign(AItem: TCT_BlendEffectXpgList);
var
  i: integer;
begin
  for i := 0 to AItem.Count - 1 do
    Add.Assign(AItem[i]);
end;

procedure TCT_BlendEffectXpgList.CopyTo(AItem: TCT_BlendEffectXpgList);
begin
end;

{ TCT_AlphaModulateEffect }

function  TCT_AlphaModulateEffect.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FA_Cont <> Nil then 
    Inc(ElemsAssigned,FA_Cont.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_AlphaModulateEffect.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'a:cont' then 
  begin
    if FA_Cont = Nil then 
      FA_Cont := TCT_EffectContainer.Create(FOwner);
    Result := FA_Cont;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_AlphaModulateEffect.Write(AWriter: TXpgWriteXML);
begin
  if (FA_Cont <> Nil) and FA_Cont.Assigned then 
  begin
    FA_Cont.WriteAttributes(AWriter);
    if xaElements in FA_Cont.FAssigneds then 
    begin
      AWriter.BeginTag('a:cont');
      FA_Cont.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:cont');
  end
  else 
    AWriter.SimpleTag('a:cont');
end;

constructor TCT_AlphaModulateEffect.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 0;
end;

destructor TCT_AlphaModulateEffect.Destroy;
begin
  if FA_Cont <> Nil then 
    FA_Cont.Free;
end;

procedure TCT_AlphaModulateEffect.Clear;
begin
  FAssigneds := [];
  if FA_Cont <> Nil then 
    FreeAndNil(FA_Cont);
end;

procedure TCT_AlphaModulateEffect.Assign(AItem: TCT_AlphaModulateEffect);
begin
  if AITem.FA_Cont <> Nil then begin
    Create_A_Cont;
    FA_Cont.Assign(AItem.FA_Cont);
  end;
end;

procedure TCT_AlphaModulateEffect.CopyTo(AItem: TCT_AlphaModulateEffect);
begin
end;

function  TCT_AlphaModulateEffect.Create_A_Cont: TCT_EffectContainer;
begin
  if FA_Cont = Nil then
    FA_Cont := TCT_EffectContainer.Create(FOwner);
  Result := FA_Cont;
end;

{ TCT_AlphaModulateEffectXpgList }

function  TCT_AlphaModulateEffectXpgList.GetItems(Index: integer): TCT_AlphaModulateEffect;
begin
  Result := TCT_AlphaModulateEffect(inherited Items[Index]);
end;

function  TCT_AlphaModulateEffectXpgList.Add: TCT_AlphaModulateEffect;
begin
  Result := TCT_AlphaModulateEffect.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_AlphaModulateEffectXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_AlphaModulateEffectXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
end;

procedure TCT_AlphaModulateEffectXpgList.Assign(AItem: TCT_AlphaModulateEffectXpgList);
var
  i: integer;
begin
  for i := 0 to AItem.Count - 1 do
    Add.Assign(AItem[i]);
end;

procedure TCT_AlphaModulateEffectXpgList.CopyTo(AItem: TCT_AlphaModulateEffectXpgList);
begin
end;

{ TCT_DashStop }

function  TCT_DashStop.CheckAssigned: integer;
begin
  FAssigneds := [xaAttributes];
  Result := 1;
end;

procedure TCT_DashStop.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_DashStop.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('d',XmlIntToStr(FD));
  AWriter.AddAttribute('sp',XmlIntToStr(FSp));
end;

procedure TCT_DashStop.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $00000064: FD := XmlStrToIntDef(AAttributes.Values[i],0);
      $000000E3: FSp := XmlStrToIntDef(AAttributes.Values[i],0);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_DashStop.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 2;
  FD := 2147483632;
  FSp := 2147483632;
end;

destructor TCT_DashStop.Destroy;
begin
end;

procedure TCT_DashStop.Clear;
begin
  FAssigneds := [];
  FD := 2147483632;
  FSp := 2147483632;
end;

procedure TCT_DashStop.Assign(AItem: TCT_DashStop);
begin
  FD := AItem.FD;
  FSp := AItem.FSp;
end;

procedure TCT_DashStop.CopyTo(AItem: TCT_DashStop);
begin
end;

{ TCT_DashStopXpgList }

function  TCT_DashStopXpgList.GetItems(Index: integer): TCT_DashStop;
begin
  Result := TCT_DashStop(inherited Items[Index]);
end;

function  TCT_DashStopXpgList.Add: TCT_DashStop;
begin
  Result := TCT_DashStop.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_DashStopXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_DashStopXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_DashStopXpgList.Assign(AItem: TCT_DashStopXpgList);
var
  i: integer;
begin
  for i := 0 to AItem.Count - 1 do
    Add.Assign(AItem[i]);
end;

procedure TCT_DashStopXpgList.CopyTo(AItem: TCT_DashStopXpgList);
begin
end;

{ TCT_PresetLineDashProperties }

function  TCT_PresetLineDashProperties.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if Integer(FVal) <> XPG_UNKNOWN_ENUM then
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_PresetLineDashProperties.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_PresetLineDashProperties.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if Integer(FVal) <> XPG_UNKNOWN_ENUM then 
    AWriter.AddAttribute('val',StrTST_PresetLineDashVal[Integer(FVal)]);
end;

procedure TCT_PresetLineDashProperties.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'val' then 
    FVal := TST_PresetLineDashVal(StrToEnum('stpldv' + AAttributes.Values[0]))
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_PresetLineDashProperties.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
  FVal := TST_PresetLineDashVal(XPG_UNKNOWN_ENUM);
end;

destructor TCT_PresetLineDashProperties.Destroy;
begin
end;

procedure TCT_PresetLineDashProperties.Clear;
begin
  FAssigneds := [];
  FVal := TST_PresetLineDashVal(XPG_UNKNOWN_ENUM);
end;

procedure TCT_PresetLineDashProperties.Assign(AItem: TCT_PresetLineDashProperties);
begin
  FVal := AItem.FVal;
end;

procedure TCT_PresetLineDashProperties.CopyTo(AItem: TCT_PresetLineDashProperties);
begin
end;

{ TCT_DashStopList }

function  TCT_DashStopList.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FA_DsXpgList <> Nil then 
    Inc(ElemsAssigned,FA_DsXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_DashStopList.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'a:ds' then 
  begin
    if FA_DsXpgList = Nil then 
      FA_DsXpgList := TCT_DashStopXpgList.Create(FOwner);
    Result := FA_DsXpgList.Add;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_DashStopList.Write(AWriter: TXpgWriteXML);
begin
  if FA_DsXpgList <> Nil then 
    FA_DsXpgList.Write(AWriter,'a:ds');
end;

constructor TCT_DashStopList.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 0;
end;

destructor TCT_DashStopList.Destroy;
begin
  if FA_DsXpgList <> Nil then 
    FA_DsXpgList.Free;
end;

procedure TCT_DashStopList.Clear;
begin
  FAssigneds := [];
  if FA_DsXpgList <> Nil then 
    FreeAndNil(FA_DsXpgList);
end;

procedure TCT_DashStopList.Assign(AItem: TCT_DashStopList);
begin
  if AItem.FA_DsXpgList <> Nil then begin
    Create_A_DsXpgList;
    FA_DsXpgList.Assign(AItem.FA_DsXpgList);
  end;
end;

procedure TCT_DashStopList.CopyTo(AItem: TCT_DashStopList);
begin
end;

function  TCT_DashStopList.Create_A_DsXpgList: TCT_DashStopXpgList;
begin
  if FA_DsXpgList = Nil then
    FA_DsXpgList := TCT_DashStopXpgList.Create(FOwner);
  Result := FA_DsXpgList;
end;

{ TCT_LineJoinRound }

function  TCT_LineJoinRound.CheckAssigned: integer;
begin
  FAssigneds := [];
  Result := 0;
end;

procedure TCT_LineJoinRound.Write(AWriter: TXpgWriteXML);
begin
end;

constructor TCT_LineJoinRound.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 0;
end;

destructor TCT_LineJoinRound.Destroy;
begin
end;

procedure TCT_LineJoinRound.Clear;
begin
  FAssigneds := [];
end;

procedure TCT_LineJoinRound.Assign(AItem: TCT_LineJoinRound);
begin
end;

procedure TCT_LineJoinRound.CopyTo(AItem: TCT_LineJoinRound);
begin
end;

{ TCT_LineJoinBevel }

function  TCT_LineJoinBevel.CheckAssigned: integer;
begin
  FAssigneds := [];
  Result := 0;
end;

procedure TCT_LineJoinBevel.Write(AWriter: TXpgWriteXML);
begin
end;

constructor TCT_LineJoinBevel.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 0;
end;

destructor TCT_LineJoinBevel.Destroy;
begin
end;

procedure TCT_LineJoinBevel.Clear;
begin
  FAssigneds := [];
end;

procedure TCT_LineJoinBevel.Assign(AItem: TCT_LineJoinBevel);
begin
end;

procedure TCT_LineJoinBevel.CopyTo(AItem: TCT_LineJoinBevel);
begin
end;

{ TCT_LineJoinMiterProperties }

function  TCT_LineJoinMiterProperties.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FLim <> 2147483632 then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_LineJoinMiterProperties.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_LineJoinMiterProperties.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FLim <> 2147483632 then 
    AWriter.AddAttribute('lim',XmlIntToStr(FLim));
end;

procedure TCT_LineJoinMiterProperties.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'lim' then 
    FLim := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_LineJoinMiterProperties.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
  FLim := 2147483632;
end;

destructor TCT_LineJoinMiterProperties.Destroy;
begin
end;

procedure TCT_LineJoinMiterProperties.Clear;
begin
  FAssigneds := [];
  FLim := 2147483632;
end;

procedure TCT_LineJoinMiterProperties.Assign(AItem: TCT_LineJoinMiterProperties);
begin
  FLim := AItem.FLim;
end;

procedure TCT_LineJoinMiterProperties.CopyTo(AItem: TCT_LineJoinMiterProperties);
begin
end;

{ TEG_LineDashProperties }

function  TEG_LineDashProperties.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FA_PrstDash <> Nil then
    Inc(ElemsAssigned,FA_PrstDash.CheckAssigned);
  if FA_CustDash <> Nil then
    Inc(ElemsAssigned,FA_CustDash.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TEG_LineDashProperties.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  case AReader.QNameHashA of
    $000003E4: begin
      if FA_PrstDash = Nil then 
        FA_PrstDash := TCT_PresetLineDashProperties.Create(FOwner);
      Result := FA_PrstDash;
    end;
    $000003DA: begin
      if FA_CustDash = Nil then 
        FA_CustDash := TCT_DashStopList.Create(FOwner);
      Result := FA_CustDash;
    end;
    else 
    begin
      Result := Nil;
      Exit;
    end
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TEG_LineDashProperties.Write(AWriter: TXpgWriteXML);
begin
  if (FA_PrstDash <> Nil) and FA_PrstDash.Assigned then 
  begin
    FA_PrstDash.WriteAttributes(AWriter);
    AWriter.SimpleTag('a:prstDash');
  end;
  if (FA_CustDash <> Nil) and FA_CustDash.Assigned then 
    if xaElements in FA_CustDash.FAssigneds then 
    begin
      AWriter.BeginTag('a:custDash');
      FA_CustDash.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:custDash');
end;

constructor TEG_LineDashProperties.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 2;
  FAttributeCount := 0;
end;

destructor TEG_LineDashProperties.Destroy;
begin
  if FA_PrstDash <> Nil then 
    FA_PrstDash.Free;
  if FA_CustDash <> Nil then 
    FA_CustDash.Free;
end;

procedure TEG_LineDashProperties.Clear;
begin
  FAssigneds := [];
  if FA_PrstDash <> Nil then 
    FreeAndNil(FA_PrstDash);
  if FA_CustDash <> Nil then 
    FreeAndNil(FA_CustDash);
end;

procedure TEG_LineDashProperties.Assign(AItem: TEG_LineDashProperties);
begin
  if AItem.FA_PrstDash <> Nil then begin
    Create_PrstDash;
    FA_PrstDash.Assign(AItem.FA_PrstDash);
  end;
  if AItem.FA_CustDash <> Nil then begin
    Create_CustDash;
    FA_CustDash.Assign(AItem.FA_CustDash);
  end;
end;

procedure TEG_LineDashProperties.CopyTo(AItem: TEG_LineDashProperties);
begin
end;

function  TEG_LineDashProperties.Create_PrstDash: TCT_PresetLineDashProperties;
begin
  if FA_PrstDash = Nil then begin
    Clear;
    FA_PrstDash := TCT_PresetLineDashProperties.Create(FOwner);
  end;
  Result := FA_PrstDash;
end;

function  TEG_LineDashProperties.Create_CustDash: TCT_DashStopList;
begin
  if FA_CustDash = Nil then begin
    Clear;
    FA_CustDash := TCT_DashStopList.Create(FOwner);
  end;
  Result := FA_CustDash;
end;

{ TEG_LineJoinProperties }

function  TEG_LineJoinProperties.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FA_Round <> Nil then 
    Inc(ElemsAssigned,FA_Round.CheckAssigned);
  if FA_Bevel <> Nil then 
    Inc(ElemsAssigned,FA_Bevel.CheckAssigned);
  if FA_Miter <> Nil then 
    Inc(ElemsAssigned,FA_Miter.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TEG_LineJoinProperties.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  case AReader.QNameHashA of
    $000002C3: begin
      if FA_Round = Nil then 
        FA_Round := TCT_LineJoinRound.Create(FOwner);
      Result := FA_Round;
    end;
    $000002A9: begin
      if FA_Bevel = Nil then 
        FA_Bevel := TCT_LineJoinBevel.Create(FOwner);
      Result := FA_Bevel;
    end;
    $000002BC: begin
      if FA_Miter = Nil then 
        FA_Miter := TCT_LineJoinMiterProperties.Create(FOwner);
      Result := FA_Miter;
    end;
    else 
    begin
      Result := Nil;
      Exit;
    end
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TEG_LineJoinProperties.Write(AWriter: TXpgWriteXML);
begin
  if FA_Round <> Nil  then
    AWriter.SimpleTag('a:round');
  if FA_Bevel <> Nil then
    AWriter.SimpleTag('a:bevel');
  if FA_Miter <> Nil then
  begin
    FA_Miter.WriteAttributes(AWriter);
    AWriter.SimpleTag('a:miter');
  end;
end;

constructor TEG_LineJoinProperties.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 3;
  FAttributeCount := 0;
end;

destructor TEG_LineJoinProperties.Destroy;
begin
  if FA_Round <> Nil then 
    FA_Round.Free;
  if FA_Bevel <> Nil then 
    FA_Bevel.Free;
  if FA_Miter <> Nil then 
    FA_Miter.Free;
end;

procedure TEG_LineJoinProperties.Clear;
begin
  FAssigneds := [];
  if FA_Round <> Nil then 
    FreeAndNil(FA_Round);
  if FA_Bevel <> Nil then 
    FreeAndNil(FA_Bevel);
  if FA_Miter <> Nil then 
    FreeAndNil(FA_Miter);
end;

procedure TEG_LineJoinProperties.Assign(AItem: TEG_LineJoinProperties);
begin
  if AItem.FA_Round <> Nil then begin
    Create_Round;
    FA_Round.Assign(AItem.FA_Round);
  end;
  if AItem.FA_Bevel <> Nil then begin
    Create_Bevel;
    FA_Bevel.Assign(AItem.FA_Bevel);
  end;
  if AItem.FA_Miter <> Nil then begin
    Create_Miter;
    FA_Miter.Assign(AItem.FA_Miter);
  end;
end;

procedure TEG_LineJoinProperties.CopyTo(AItem: TEG_LineJoinProperties);
begin
end;

function  TEG_LineJoinProperties.Create_Round: TCT_LineJoinRound;
begin
  if FA_Round = Nil then begin
    Clear;
    FA_Round := TCT_LineJoinRound.Create(FOwner);
  end;
  Result := FA_Round;
end;

function  TEG_LineJoinProperties.Create_Bevel: TCT_LineJoinBevel;
begin
  if FA_Bevel = Nil then begin
    Clear;
    FA_Bevel := TCT_LineJoinBevel.Create(FOwner);
  end;
  Result := FA_Bevel;
end;

function  TEG_LineJoinProperties.Create_Miter: TCT_LineJoinMiterProperties;
begin
  if FA_Miter = Nil then begin
    Clear;
    FA_Miter := TCT_LineJoinMiterProperties.Create(FOwner);
  end;
  Result := FA_Miter;
end;

{ TCT_LineEndProperties }

function  TCT_LineEndProperties.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if Integer(FType) <> XPG_UNKNOWN_ENUM then
    Inc(AttrsAssigned);
  if Integer(FW) <> XPG_UNKNOWN_ENUM then 
    Inc(AttrsAssigned);
  if Integer(FLen) <> XPG_UNKNOWN_ENUM then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_LineEndProperties.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_LineEndProperties.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if Integer(FType) <> XPG_UNKNOWN_ENUM then 
    AWriter.AddAttribute('type',StrTST_LineEndType[Integer(FType)]);
  if Integer(FW) <> XPG_UNKNOWN_ENUM then 
    AWriter.AddAttribute('w',StrTST_LineEndWidth[Integer(FW)]);
  if Integer(FLen) <> XPG_UNKNOWN_ENUM then 
    AWriter.AddAttribute('len',StrTST_LineEndLength[Integer(FLen)]);
end;

procedure TCT_LineEndProperties.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $000001C2: FType := TST_LineEndType(StrToEnum('stlet' + AAttributes.Values[i]));
      $00000077: FW := TST_LineEndWidth(StrToEnum('stlew' + AAttributes.Values[i]));
      $0000013F: FLen := TST_LineEndLength(StrToEnum('stlel' + AAttributes.Values[i]));
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_LineEndProperties.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 3;
  Clear;
end;

destructor TCT_LineEndProperties.Destroy;
begin
end;

procedure TCT_LineEndProperties.Clear;
begin
  FAssigneds := [];
  FType := TST_LineEndType(XPG_UNKNOWN_ENUM);
  FW := TST_LineEndWidth(XPG_UNKNOWN_ENUM);
  FLen := TST_LineEndLength(XPG_UNKNOWN_ENUM);
end;

procedure TCT_LineEndProperties.Assign(AItem: TCT_LineEndProperties);
begin
  FType := AItem.FType;
  FW := AItem.FW;
  FLen := AItem.FLen;
end;

procedure TCT_LineEndProperties.CopyTo(AItem: TCT_LineEndProperties);
begin
end;

{ TCT_EffectList }

function  TCT_EffectList.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FA_Blur <> Nil then 
    Inc(ElemsAssigned,FA_Blur.CheckAssigned);
  if FA_FillOverlay <> Nil then 
    Inc(ElemsAssigned,FA_FillOverlay.CheckAssigned);
  if FA_Glow <> Nil then 
    Inc(ElemsAssigned,FA_Glow.CheckAssigned);
  if FA_InnerShdw <> Nil then 
    Inc(ElemsAssigned,FA_InnerShdw.CheckAssigned);
  if FA_OuterShdw <> Nil then 
    Inc(ElemsAssigned,FA_OuterShdw.CheckAssigned);
  if FA_PrstShdw <> Nil then 
    Inc(ElemsAssigned,FA_PrstShdw.CheckAssigned);
  if FA_Reflection <> Nil then 
    Inc(ElemsAssigned,FA_Reflection.CheckAssigned);
  if FA_SoftEdge <> Nil then 
    Inc(ElemsAssigned,FA_SoftEdge.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_EffectList.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $00000250: begin
      if FA_Blur = Nil then 
        FA_Blur := TCT_BlurEffect.Create(FOwner);
      Result := FA_Blur;
    end;
    $00000524: begin
      if FA_FillOverlay = Nil then 
        FA_FillOverlay := TCT_FillOverlayEffect.Create(FOwner);
      Result := FA_FillOverlay;
    end;
    $00000254: begin
      if FA_Glow = Nil then 
        FA_Glow := TCT_GlowEffect.Create(FOwner);
      Result := FA_Glow;
    end;
    $0000044D: begin
      if FA_InnerShdw = Nil then 
        FA_InnerShdw := TCT_InnerShadowEffect.Create(FOwner);
      Result := FA_InnerShdw;
    end;
    $00000460: begin
      if FA_OuterShdw = Nil then 
        FA_OuterShdw := TCT_OuterShadowEffect.Create(FOwner);
      Result := FA_OuterShdw;
    end;
    $000003FA: begin
      if FA_PrstShdw = Nil then 
        FA_PrstShdw := TCT_PresetShadowEffect.Create(FOwner);
      Result := FA_PrstShdw;
    end;
    $000004C6: begin
      if FA_Reflection = Nil then 
        FA_Reflection := TCT_ReflectionEffect.Create(FOwner);
      Result := FA_Reflection;
    end;
    $000003CC: begin
      if FA_SoftEdge = Nil then 
        FA_SoftEdge := TCT_SoftEdgesEffect.Create(FOwner);
      Result := FA_SoftEdge;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_EffectList.Write(AWriter: TXpgWriteXML);
begin
  if (FA_Blur <> Nil) and FA_Blur.Assigned then 
  begin
    FA_Blur.WriteAttributes(AWriter);
    AWriter.SimpleTag('a:blur');
  end;
  if (FA_FillOverlay <> Nil) and FA_FillOverlay.Assigned then 
  begin
    FA_FillOverlay.WriteAttributes(AWriter);
    if xaElements in FA_FillOverlay.FAssigneds then 
    begin
      AWriter.BeginTag('a:fillOverlay');
      FA_FillOverlay.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:fillOverlay');
  end;
  if (FA_Glow <> Nil) and FA_Glow.Assigned then 
  begin
    FA_Glow.WriteAttributes(AWriter);
    if xaElements in FA_Glow.FAssigneds then 
    begin
      AWriter.BeginTag('a:glow');
      FA_Glow.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:glow');
  end;
  if (FA_InnerShdw <> Nil) and FA_InnerShdw.Assigned then 
  begin
    FA_InnerShdw.WriteAttributes(AWriter);
    if xaElements in FA_InnerShdw.FAssigneds then 
    begin
      AWriter.BeginTag('a:innerShdw');
      FA_InnerShdw.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:innerShdw');
  end;
  if (FA_OuterShdw <> Nil) and FA_OuterShdw.Assigned then 
  begin
    FA_OuterShdw.WriteAttributes(AWriter);
    if xaElements in FA_OuterShdw.FAssigneds then 
    begin
      AWriter.BeginTag('a:outerShdw');
      FA_OuterShdw.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:outerShdw');
  end;
  if (FA_PrstShdw <> Nil) and FA_PrstShdw.Assigned then 
  begin
    FA_PrstShdw.WriteAttributes(AWriter);
    if xaElements in FA_PrstShdw.FAssigneds then 
    begin
      AWriter.BeginTag('a:prstShdw');
      FA_PrstShdw.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:prstShdw');
  end;
  if (FA_Reflection <> Nil) and FA_Reflection.Assigned then 
  begin
    FA_Reflection.WriteAttributes(AWriter);
    AWriter.SimpleTag('a:reflection');
  end;
  if (FA_SoftEdge <> Nil) and FA_SoftEdge.Assigned then 
  begin
    FA_SoftEdge.WriteAttributes(AWriter);
    AWriter.SimpleTag('a:softEdge');
  end;
end;

constructor TCT_EffectList.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 8;
  FAttributeCount := 0;
end;

destructor TCT_EffectList.Destroy;
begin
  if FA_Blur <> Nil then 
    FA_Blur.Free;
  if FA_FillOverlay <> Nil then 
    FA_FillOverlay.Free;
  if FA_Glow <> Nil then 
    FA_Glow.Free;
  if FA_InnerShdw <> Nil then 
    FA_InnerShdw.Free;
  if FA_OuterShdw <> Nil then 
    FA_OuterShdw.Free;
  if FA_PrstShdw <> Nil then 
    FA_PrstShdw.Free;
  if FA_Reflection <> Nil then 
    FA_Reflection.Free;
  if FA_SoftEdge <> Nil then 
    FA_SoftEdge.Free;
end;

procedure TCT_EffectList.Clear;
begin
  FAssigneds := [];
  if FA_Blur <> Nil then 
    FreeAndNil(FA_Blur);
  if FA_FillOverlay <> Nil then 
    FreeAndNil(FA_FillOverlay);
  if FA_Glow <> Nil then 
    FreeAndNil(FA_Glow);
  if FA_InnerShdw <> Nil then 
    FreeAndNil(FA_InnerShdw);
  if FA_OuterShdw <> Nil then 
    FreeAndNil(FA_OuterShdw);
  if FA_PrstShdw <> Nil then 
    FreeAndNil(FA_PrstShdw);
  if FA_Reflection <> Nil then 
    FreeAndNil(FA_Reflection);
  if FA_SoftEdge <> Nil then 
    FreeAndNil(FA_SoftEdge);
end;

procedure TCT_EffectList.Assign(AItem: TCT_EffectList);
begin
end;

procedure TCT_EffectList.CopyTo(AItem: TCT_EffectList);
begin
end;

function  TCT_EffectList.Create_A_Blur: TCT_BlurEffect;
begin
  if FA_Blur = Nil then
    FA_Blur := TCT_BlurEffect.Create(FOwner);
  Result := FA_Blur;
end;

function  TCT_EffectList.Create_A_FillOverlay: TCT_FillOverlayEffect;
begin
  if FA_FillOverlay = Nil then
    FA_FillOverlay := TCT_FillOverlayEffect.Create(FOwner);
  Result := FA_FillOverlay;
end;

function  TCT_EffectList.Create_A_Glow: TCT_GlowEffect;
begin
  if FA_Glow = Nil then
    FA_Glow := TCT_GlowEffect.Create(FOwner);
  Result := FA_Glow;
end;

function  TCT_EffectList.Create_A_InnerShdw: TCT_InnerShadowEffect;
begin
  if FA_InnerShdw = Nil then
    FA_InnerShdw := TCT_InnerShadowEffect.Create(FOwner);
  Result := FA_InnerShdw;
end;

function  TCT_EffectList.Create_A_OuterShdw: TCT_OuterShadowEffect;
begin
  if FA_OuterShdw = Nil then
    FA_OuterShdw := TCT_OuterShadowEffect.Create(FOwner);
  Result := FA_OuterShdw;
end;

function  TCT_EffectList.Create_A_PrstShdw: TCT_PresetShadowEffect;
begin
  if FA_PrstShdw = Nil then
    FA_PrstShdw := TCT_PresetShadowEffect.Create(FOwner);
  Result := FA_PrstShdw;
end;

function  TCT_EffectList.Create_A_Reflection: TCT_ReflectionEffect;
begin
  if FA_Reflection = Nil then
    FA_Reflection := TCT_ReflectionEffect.Create(FOwner);
  Result := FA_Reflection;
end;

function  TCT_EffectList.Create_A_SoftEdge: TCT_SoftEdgesEffect;
begin
  if FA_SoftEdge = Nil then
    FA_SoftEdge := TCT_SoftEdgesEffect.Create(FOwner);
  Result := FA_SoftEdge;
end;

{ TCT_TextUnderlineLineFollowText }

function  TCT_TextUnderlineLineFollowText.CheckAssigned: integer;
begin
  FAssigneds := [];
  Result := 0;
end;

procedure TCT_TextUnderlineLineFollowText.Write(AWriter: TXpgWriteXML);
begin
end;

constructor TCT_TextUnderlineLineFollowText.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 0;
end;

destructor TCT_TextUnderlineLineFollowText.Destroy;
begin
end;

procedure TCT_TextUnderlineLineFollowText.Clear;
begin
  FAssigneds := [];
end;

procedure TCT_TextUnderlineLineFollowText.Assign(AItem: TCT_TextUnderlineLineFollowText);
begin
end;

procedure TCT_TextUnderlineLineFollowText.CopyTo(AItem: TCT_TextUnderlineLineFollowText);
begin
end;

{ TCT_LineProperties }

function  TCT_LineProperties.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FW <> 2147483632 then 
    Inc(AttrsAssigned);
  if Integer(FCap) <> XPG_UNKNOWN_ENUM then
    Inc(AttrsAssigned);
  if Integer(FCmpd) <> XPG_UNKNOWN_ENUM then 
    Inc(AttrsAssigned);
  if Integer(FAlgn) <> XPG_UNKNOWN_ENUM then 
    Inc(AttrsAssigned);
  Inc(ElemsAssigned,FA_EG_LineFillProperties.CheckAssigned);
  Inc(ElemsAssigned,FA_EG_LineDashProperties.CheckAssigned);
  Inc(ElemsAssigned,FA_EG_LineJoinProperties.CheckAssigned);
  if FA_HeadEnd <> Nil then 
    Inc(ElemsAssigned,FA_HeadEnd.CheckAssigned);
  if FA_TailEnd <> Nil then 
    Inc(ElemsAssigned,FA_TailEnd.CheckAssigned);
  if FA_ExtLst <> Nil then 
    Inc(ElemsAssigned,FA_ExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_LineProperties.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  case AReader.QNameHashA of
    $00000344: begin
      if FA_HeadEnd = Nil then 
        FA_HeadEnd := TCT_LineEndProperties.Create(FOwner);
      Result := FA_HeadEnd;
    end;
    $0000035C: begin
      if FA_TailEnd = Nil then 
        FA_TailEnd := TCT_LineEndProperties.Create(FOwner);
      Result := FA_TailEnd;
    end;
    $0000031F: begin
      if FA_ExtLst = Nil then 
        FA_ExtLst := TCT_OfficeArtExtensionList.Create(FOwner);
      Result := FA_ExtLst;
    end;
    else 
    begin
      Result := FA_EG_LineFillProperties.HandleElement(AReader);
      if Result = Nil then 
      begin
        Result := FA_EG_LineDashProperties.HandleElement(AReader);
        if Result = Nil then 
        begin
          Result := FA_EG_LineJoinProperties.HandleElement(AReader);
          if Result = Nil then 
            FOwner.Errors.Error(xemUnknownElement,AReader.QName);
        end;
      end;
    end
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_LineProperties.Write(AWriter: TXpgWriteXML);
begin
  FA_EG_LineFillProperties.Write(AWriter);
  FA_EG_LineDashProperties.Write(AWriter);
  FA_EG_LineJoinProperties.Write(AWriter);
  if FA_HeadEnd <> Nil then
  begin
    FA_HeadEnd.WriteAttributes(AWriter);
    AWriter.SimpleTag('a:headEnd');
  end;
  if FA_TailEnd <> Nil then
  begin
    FA_TailEnd.WriteAttributes(AWriter);
    AWriter.SimpleTag('a:tailEnd');
  end;
  if (FA_ExtLst <> Nil) and FA_ExtLst.Assigned then 
    if xaElements in FA_ExtLst.FAssigneds then 
    begin
      AWriter.BeginTag('a:extLst');
      FA_ExtLst.Write(AWriter);
      AWriter.EndTag;
    end;
end;

procedure TCT_LineProperties.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FW <> 2147483632 then 
    AWriter.AddAttribute('w',XmlIntToStr(FW));
  if Integer(FCap) <> XPG_UNKNOWN_ENUM then 
    AWriter.AddAttribute('cap',StrTST_LineCap[Integer(FCap)]);
  if Integer(FCmpd) <> XPG_UNKNOWN_ENUM then 
    AWriter.AddAttribute('cmpd',StrTST_CompoundLine[Integer(FCmpd)]);
  if Integer(FAlgn) <> XPG_UNKNOWN_ENUM then 
    AWriter.AddAttribute('algn',StrTST_PenAlignment[Integer(FAlgn)]);
end;

procedure TCT_LineProperties.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $00000077: FW := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000134: FCap := TST_LineCap(StrToEnum('stlc' + AAttributes.Values[i]));
      $000001A4: FCmpd := TST_CompoundLine(StrToEnum('stcl' + AAttributes.Values[i]));
      $000001A2: FAlgn := TST_PenAlignment(StrToEnum('stpa' + AAttributes.Values[i]));
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_LineProperties.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 6;
  FAttributeCount := 4;
  FA_EG_LineFillProperties := TEG_FillProperties.Create(FOwner);
  FA_EG_LineDashProperties := TEG_LineDashProperties.Create(FOwner);
  FA_EG_LineJoinProperties := TEG_LineJoinProperties.Create(FOwner);
  FW := 2147483632;
  FCap := TST_LineCap(XPG_UNKNOWN_ENUM);
  FCmpd := TST_CompoundLine(XPG_UNKNOWN_ENUM);
  FAlgn := TST_PenAlignment(XPG_UNKNOWN_ENUM);
end;

destructor TCT_LineProperties.Destroy;
begin
  FA_EG_LineFillProperties.Free;
  FA_EG_LineDashProperties.Free;
  FA_EG_LineJoinProperties.Free;
  if FA_HeadEnd <> Nil then
    FA_HeadEnd.Free;
  if FA_TailEnd <> Nil then
    FA_TailEnd.Free;
  if FA_ExtLst <> Nil then
    FA_ExtLst.Free;
end;

procedure TCT_LineProperties.Free_HeadEnd;
begin
  if FA_HeadEnd <> Nil then
    FreeAndNil(FA_HeadEnd);
end;

procedure TCT_LineProperties.Free_TailEnd;
begin
  if FA_TailEnd <> Nil then
    FreeAndNil(FA_TailEnd);
end;

procedure TCT_LineProperties.Clear;
begin
  FAssigneds := [];
  FA_EG_LineFillProperties.Clear;
  FA_EG_LineDashProperties.Clear;
  FA_EG_LineJoinProperties.Clear;
  if FA_HeadEnd <> Nil then
    FreeAndNil(FA_HeadEnd);
  if FA_TailEnd <> Nil then
    FreeAndNil(FA_TailEnd);
  if FA_ExtLst <> Nil then
    FreeAndNil(FA_ExtLst);
  FW := 2147483632;
  FCap := TST_LineCap(XPG_UNKNOWN_ENUM);
  FCmpd := TST_CompoundLine(XPG_UNKNOWN_ENUM);
  FAlgn := TST_PenAlignment(XPG_UNKNOWN_ENUM);
end;

procedure TCT_LineProperties.Assign(AItem: TCT_LineProperties);
begin
  FW := AItem.FW;
  FCap := AItem.FCap;
  FCmpd := AItem.FCmpd;
  FAlgn := AItem.FAlgn;
  FA_EG_LineFillProperties.Assign(AItem.FA_EG_LineFillProperties);
  FA_EG_LineDashProperties.Assign(AItem.FA_EG_LineDashProperties);
  FA_EG_LineJoinProperties.Assign(AItem.FA_EG_LineJoinProperties);
  if AItem.FA_HeadEnd <> Nil then begin
    Create_HeadEnd;
    FA_HeadEnd.Assign(AItem.FA_HeadEnd);
  end;
  if AItem.FA_TailEnd <> Nil then begin
    Create_TailEnd;
    FA_TailEnd.Assign(AItem.FA_TailEnd);
  end;
  if AItem.FA_ExtLst <> Nil then begin
    Create_ExtLst;
    FA_ExtLst.Assign(AItem.FA_ExtLst);
  end;
end;

procedure TCT_LineProperties.CopyTo(AItem: TCT_LineProperties);
begin
end;

function  TCT_LineProperties.Create_HeadEnd: TCT_LineEndProperties;
begin
  if FA_HeadEnd = Nil then
    FA_HeadEnd := TCT_LineEndProperties.Create(FOwner);
  Result := FA_HeadEnd;
end;

function  TCT_LineProperties.Create_TailEnd: TCT_LineEndProperties;
begin
  if FA_TailEnd = Nil then
    FA_TailEnd := TCT_LineEndProperties.Create(FOwner);
  Result := FA_TailEnd;
end;

function  TCT_LineProperties.Create_ExtLst: TCT_OfficeArtExtensionList;
begin
  Result := TCT_OfficeArtExtensionList.Create(FOwner);
  FA_ExtLst := Result;
end;

{ TCT_LinePropertiesXpgList }

function  TCT_LinePropertiesXpgList.GetItems(Index: integer): TCT_LineProperties;
begin
  Result := TCT_LineProperties(inherited Items[Index]);
end;

function  TCT_LinePropertiesXpgList.Add: TCT_LineProperties;
begin
  Result := TCT_LineProperties.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_LinePropertiesXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_LinePropertiesXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_LinePropertiesXpgList.Assign(AItem: TCT_LinePropertiesXpgList);
begin
end;

procedure TCT_LinePropertiesXpgList.CopyTo(AItem: TCT_LinePropertiesXpgList);
begin
end;

{ TCT_TextUnderlineFillFollowText }

function  TCT_TextUnderlineFillFollowText.CheckAssigned: integer;
begin
  FAssigneds := [];
  Result := 0;
end;

procedure TCT_TextUnderlineFillFollowText.Write(AWriter: TXpgWriteXML);
begin
end;

constructor TCT_TextUnderlineFillFollowText.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 0;
end;

destructor TCT_TextUnderlineFillFollowText.Destroy;
begin
end;

procedure TCT_TextUnderlineFillFollowText.Clear;
begin
  FAssigneds := [];
end;

procedure TCT_TextUnderlineFillFollowText.Assign(AItem: TCT_TextUnderlineFillFollowText);
begin
end;

procedure TCT_TextUnderlineFillFollowText.CopyTo(AItem: TCT_TextUnderlineFillFollowText);
begin
end;

{ TCT_TextUnderlineFillGroupWrapper }

function  TCT_TextUnderlineFillGroupWrapper.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FA_EG_FillProperties.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_TextUnderlineFillGroupWrapper.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := FA_EG_FillProperties.HandleElement(AReader);
  if Result = Nil then 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_TextUnderlineFillGroupWrapper.Write(AWriter: TXpgWriteXML);
begin
  FA_EG_FillProperties.Write(AWriter);
end;

constructor TCT_TextUnderlineFillGroupWrapper.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 0;
  FA_EG_FillProperties := TEG_FillProperties.Create(FOwner);
end;

destructor TCT_TextUnderlineFillGroupWrapper.Destroy;
begin
  FA_EG_FillProperties.Free;
end;

procedure TCT_TextUnderlineFillGroupWrapper.Clear;
begin
  FAssigneds := [];
  FA_EG_FillProperties.Clear;
end;

procedure TCT_TextUnderlineFillGroupWrapper.Assign(AItem: TCT_TextUnderlineFillGroupWrapper);
begin
end;

procedure TCT_TextUnderlineFillGroupWrapper.CopyTo(AItem: TCT_TextUnderlineFillGroupWrapper);
begin
end;

{ TCT_EmbeddedWAVAudioFile }

function  TCT_EmbeddedWAVAudioFile.CheckAssigned: integer;
begin
  FAssigneds := [xaAttributes];
  Result := 1;
end;

procedure TCT_EmbeddedWAVAudioFile.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_EmbeddedWAVAudioFile.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('r:embed',FR_Embed);
  if FName <> '' then 
    AWriter.AddAttribute('name',FName);
  if FBuiltIn <> False then 
    AWriter.AddAttribute('builtIn',XmlBoolToStr(FBuiltIn));
end;

procedure TCT_EmbeddedWAVAudioFile.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $000002A9: FR_Embed := AAttributes.Values[i];
      $000001A1: FName := AAttributes.Values[i];
      $000002D7: FBuiltIn := XmlStrToBoolDef(AAttributes.Values[i],False);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_EmbeddedWAVAudioFile.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 3;
  FBuiltIn := False;
end;

destructor TCT_EmbeddedWAVAudioFile.Destroy;
begin
end;

procedure TCT_EmbeddedWAVAudioFile.Clear;
begin
  FAssigneds := [];
  FR_Embed := '';
  FName := '';
  FBuiltIn := False;
end;

procedure TCT_EmbeddedWAVAudioFile.Assign(AItem: TCT_EmbeddedWAVAudioFile);
begin
end;

procedure TCT_EmbeddedWAVAudioFile.CopyTo(AItem: TCT_EmbeddedWAVAudioFile);
begin
end;

{ TCT_AdjPoint2D }

function  TCT_AdjPoint2D.CheckAssigned: integer;
begin
  FAssigneds := [xaAttributes];
  Result := 1;
end;

procedure TCT_AdjPoint2D.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_AdjPoint2D.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('x',WriteUnionTST_AdjCoordinate(PFX));
  AWriter.AddAttribute('y',WriteUnionTST_AdjCoordinate(PFY));
end;

procedure TCT_AdjPoint2D.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $00000078: ReadUnionTST_AdjCoordinate(AAttributes.Values[i],PFX);
      $00000079: ReadUnionTST_AdjCoordinate(AAttributes.Values[i],PFY);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_AdjPoint2D.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 2;
  PFX := @FX;
  PX.nVal := -1;
  PFY := @FY;
  PY.nVal := -1;
end;

destructor TCT_AdjPoint2D.Destroy;
begin
end;

procedure TCT_AdjPoint2D.Clear;
begin
  FAssigneds := [];
end;

procedure TCT_AdjPoint2D.Assign(AItem: TCT_AdjPoint2D);
begin
end;

procedure TCT_AdjPoint2D.CopyTo(AItem: TCT_AdjPoint2D);
begin
end;

{ TCT_AdjPoint2DXpgList }

function  TCT_AdjPoint2DXpgList.GetItems(Index: integer): TCT_AdjPoint2D;
begin
  Result := TCT_AdjPoint2D(inherited Items[Index]);
end;

function  TCT_AdjPoint2DXpgList.Add: TCT_AdjPoint2D;
begin
  Result := TCT_AdjPoint2D.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_AdjPoint2DXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_AdjPoint2DXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_AdjPoint2DXpgList.Assign(AItem: TCT_AdjPoint2DXpgList);
begin
end;

procedure TCT_AdjPoint2DXpgList.CopyTo(AItem: TCT_AdjPoint2DXpgList);
begin
end;

{ TCT_TextSpacingPercent }

function  TCT_TextSpacingPercent.CheckAssigned: integer;
begin
  FAssigneds := [xaAttributes];
  Result := 1;
end;

procedure TCT_TextSpacingPercent.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_TextSpacingPercent.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('val',XmlIntToStr(FVal));
end;

procedure TCT_TextSpacingPercent.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'val' then 
    FVal := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_TextSpacingPercent.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
  FVal := 2147483632;
end;

destructor TCT_TextSpacingPercent.Destroy;
begin
end;

procedure TCT_TextSpacingPercent.Clear;
begin
  FAssigneds := [];
  FVal := 2147483632;
end;

procedure TCT_TextSpacingPercent.Assign(AItem: TCT_TextSpacingPercent);
begin
end;

procedure TCT_TextSpacingPercent.CopyTo(AItem: TCT_TextSpacingPercent);
begin
end;

{ TCT_TextSpacingPoint }

function  TCT_TextSpacingPoint.CheckAssigned: integer;
begin
  FAssigneds := [xaAttributes];
  Result := 1;
end;

procedure TCT_TextSpacingPoint.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_TextSpacingPoint.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('val',XmlIntToStr(FVal));
end;

procedure TCT_TextSpacingPoint.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'val' then 
    FVal := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_TextSpacingPoint.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
  FVal := 2147483632;
end;

destructor TCT_TextSpacingPoint.Destroy;
begin
end;

procedure TCT_TextSpacingPoint.Clear;
begin
  FAssigneds := [];
  FVal := 2147483632;
end;

procedure TCT_TextSpacingPoint.Assign(AItem: TCT_TextSpacingPoint);
begin
end;

procedure TCT_TextSpacingPoint.CopyTo(AItem: TCT_TextSpacingPoint);
begin
end;

{ TCT_TextBulletColorFollowText }

function  TCT_TextBulletColorFollowText.CheckAssigned: integer;
begin
  FAssigneds := [];
  Result := 0;
end;

procedure TCT_TextBulletColorFollowText.Write(AWriter: TXpgWriteXML);
begin
end;

constructor TCT_TextBulletColorFollowText.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 0;
end;

destructor TCT_TextBulletColorFollowText.Destroy;
begin
end;

procedure TCT_TextBulletColorFollowText.Clear;
begin
  FAssigneds := [];
end;

procedure TCT_TextBulletColorFollowText.Assign(AItem: TCT_TextBulletColorFollowText);
begin
end;

procedure TCT_TextBulletColorFollowText.CopyTo(AItem: TCT_TextBulletColorFollowText);
begin
end;

{ TCT_TextBulletSizeFollowText }

function  TCT_TextBulletSizeFollowText.CheckAssigned: integer;
begin
  FAssigneds := [];
  Result := 0;
end;

procedure TCT_TextBulletSizeFollowText.Write(AWriter: TXpgWriteXML);
begin
end;

constructor TCT_TextBulletSizeFollowText.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 0;
end;

destructor TCT_TextBulletSizeFollowText.Destroy;
begin
end;

procedure TCT_TextBulletSizeFollowText.Clear;
begin
  FAssigneds := [];
end;

procedure TCT_TextBulletSizeFollowText.Assign(AItem: TCT_TextBulletSizeFollowText);
begin
end;

procedure TCT_TextBulletSizeFollowText.CopyTo(AItem: TCT_TextBulletSizeFollowText);
begin
end;

{ TCT_TextBulletSizePercent }

function  TCT_TextBulletSizePercent.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FVal <> 2147483632 then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_TextBulletSizePercent.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_TextBulletSizePercent.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FVal <> 2147483632 then 
    AWriter.AddAttribute('val',XmlIntToStr(FVal));
end;

procedure TCT_TextBulletSizePercent.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'val' then 
    FVal := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_TextBulletSizePercent.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
  FVal := 2147483632;
end;

destructor TCT_TextBulletSizePercent.Destroy;
begin
end;

procedure TCT_TextBulletSizePercent.Clear;
begin
  FAssigneds := [];
  FVal := 2147483632;
end;

procedure TCT_TextBulletSizePercent.Assign(AItem: TCT_TextBulletSizePercent);
begin
end;

procedure TCT_TextBulletSizePercent.CopyTo(AItem: TCT_TextBulletSizePercent);
begin
end;

{ TCT_TextBulletSizePoint }

function  TCT_TextBulletSizePoint.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FVal <> 2147483632 then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_TextBulletSizePoint.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_TextBulletSizePoint.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FVal <> 2147483632 then 
    AWriter.AddAttribute('val',XmlIntToStr(FVal));
end;

procedure TCT_TextBulletSizePoint.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'val' then 
    FVal := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_TextBulletSizePoint.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
  FVal := 2147483632;
end;

destructor TCT_TextBulletSizePoint.Destroy;
begin
end;

procedure TCT_TextBulletSizePoint.Clear;
begin
  FAssigneds := [];
  FVal := 2147483632;
end;

procedure TCT_TextBulletSizePoint.Assign(AItem: TCT_TextBulletSizePoint);
begin
end;

procedure TCT_TextBulletSizePoint.CopyTo(AItem: TCT_TextBulletSizePoint);
begin
end;

{ TCT_TextBulletTypefaceFollowText }

function  TCT_TextBulletTypefaceFollowText.CheckAssigned: integer;
begin
  FAssigneds := [];
  Result := 0;
end;

procedure TCT_TextBulletTypefaceFollowText.Write(AWriter: TXpgWriteXML);
begin
end;

constructor TCT_TextBulletTypefaceFollowText.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 0;
end;

destructor TCT_TextBulletTypefaceFollowText.Destroy;
begin
end;

procedure TCT_TextBulletTypefaceFollowText.Clear;
begin
  FAssigneds := [];
end;

procedure TCT_TextBulletTypefaceFollowText.Assign(AItem: TCT_TextBulletTypefaceFollowText);
begin
end;

procedure TCT_TextBulletTypefaceFollowText.CopyTo(AItem: TCT_TextBulletTypefaceFollowText);
begin
end;

{ TCT_TextFont }

function  TCT_TextFont.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FTypeface <> '' then 
    Inc(AttrsAssigned);
  if FPanose <> -16 then 
    Inc(AttrsAssigned);
  if FPitchFamily <> 0 then 
    Inc(AttrsAssigned);
  if FCharset <> 1 then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_TextFont.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_TextFont.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FTypeface <> '' then 
    AWriter.AddAttribute('typeface',FTypeface);
  if FPanose <> -16 then 
    AWriter.AddAttribute('panose',XmlIntToHexStr(FPanose,6));
  if FPitchFamily <> 0 then 
    AWriter.AddAttribute('pitchFamily',XmlIntToStr(FPitchFamily));
  if FCharset <> 1 then 
    AWriter.AddAttribute('charset',XmlIntToStr(FCharset));
end;

procedure TCT_TextFont.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $00000351: FTypeface := AAttributes.Values[i];
      $00000286: FPanose := XmlStrToIntDef('$' + AAttributes.Values[i],0);
      $0000047A: FPitchFamily := XmlStrToIntDef(AAttributes.Values[i],0);
      $000002EA: FCharset := XmlStrToIntDef(AAttributes.Values[i],0);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_TextFont.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 4;
  FPanose := -16;
  FPitchFamily := 0;
  FCharset := 1;
end;

destructor TCT_TextFont.Destroy;
begin
end;

procedure TCT_TextFont.Clear;
begin
  FAssigneds := [];
  FTypeface := '';
  FPanose := -16;
  FPitchFamily := 0;
  FCharset := 1;
end;

procedure TCT_TextFont.Assign(AItem: TCT_TextFont);
begin
end;

procedure TCT_TextFont.CopyTo(AItem: TCT_TextFont);
begin
end;

{ TCT_TextNoBullet }

function  TCT_TextNoBullet.CheckAssigned: integer;
begin
  FAssigneds := [];
  Result := 0;
end;

procedure TCT_TextNoBullet.Write(AWriter: TXpgWriteXML);
begin
end;

constructor TCT_TextNoBullet.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 0;
end;

destructor TCT_TextNoBullet.Destroy;
begin
end;

procedure TCT_TextNoBullet.Clear;
begin
  FAssigneds := [];
end;

procedure TCT_TextNoBullet.Assign(AItem: TCT_TextNoBullet);
begin
end;

procedure TCT_TextNoBullet.CopyTo(AItem: TCT_TextNoBullet);
begin
end;

{ TCT_TextAutonumberBullet }

function  TCT_TextAutonumberBullet.CheckAssigned: integer;
begin
  FAssigneds := [xaAttributes];
  Result := 1;
end;

procedure TCT_TextAutonumberBullet.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_TextAutonumberBullet.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FType <> TST_TextAutonumberScheme(XPG_UNKNOWN_ENUM) then
    AWriter.AddAttribute('type',StrTST_TextAutonumberScheme[Integer(FType)]);
  if FStartAt <> 1 then
    AWriter.AddAttribute('startAt',XmlIntToStr(FStartAt));
end;

procedure TCT_TextAutonumberBullet.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
    case AAttributes.HashA[i] of
      $000001C2: FType := TST_TextAutonumberScheme(StrToEnum('sttas' + AAttributes.Values[i]));
      $000002E3: FStartAt := XmlStrToIntDef(AAttributes.Values[i],0);
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_TextAutonumberBullet.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 2;
  FType := TST_TextAutonumberScheme(XPG_UNKNOWN_ENUM);
  FStartAt := 1;
end;

destructor TCT_TextAutonumberBullet.Destroy;
begin
end;

procedure TCT_TextAutonumberBullet.Clear;
begin
  FAssigneds := [];
  FStartAt := 1;
  FType := TST_TextAutonumberScheme(XPG_UNKNOWN_ENUM);
end;

procedure TCT_TextAutonumberBullet.Assign(AItem: TCT_TextAutonumberBullet);
begin
end;

procedure TCT_TextAutonumberBullet.CopyTo(AItem: TCT_TextAutonumberBullet);
begin
end;

{ TCT_TextCharBullet }

function  TCT_TextCharBullet.CheckAssigned: integer;
begin
  FAssigneds := [xaAttributes];
  Result := 1;
end;

procedure TCT_TextCharBullet.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_TextCharBullet.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('char',FChar);
end;

procedure TCT_TextCharBullet.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'char' then 
    FChar := AAttributes.Values[0]
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_TextCharBullet.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
end;

destructor TCT_TextCharBullet.Destroy;
begin
end;

procedure TCT_TextCharBullet.Clear;
begin
  FAssigneds := [];
  FChar := '';
end;

procedure TCT_TextCharBullet.Assign(AItem: TCT_TextCharBullet);
begin
end;

procedure TCT_TextCharBullet.CopyTo(AItem: TCT_TextCharBullet);
begin
end;

{ TCT_TextBlipBullet }

function  TCT_TextBlipBullet.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FA_Blip <> Nil then 
    Inc(ElemsAssigned,FA_Blip.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_TextBlipBullet.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'a:blip' then 
  begin
    if FA_Blip = Nil then 
      FA_Blip := TCT_Blip.Create(FOwner);
    Result := FA_Blip;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_TextBlipBullet.Write(AWriter: TXpgWriteXML);
begin
  if (FA_Blip <> Nil) and FA_Blip.Assigned then 
  begin
    FA_Blip.WriteAttributes(AWriter);
    if xaElements in FA_Blip.FAssigneds then 
    begin
      AWriter.BeginTag('a:blip');
      FA_Blip.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:blip');
  end
  else 
    AWriter.SimpleTag('a:blip');
end;

constructor TCT_TextBlipBullet.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 0;
end;

destructor TCT_TextBlipBullet.Destroy;
begin
  if FA_Blip <> Nil then 
    FA_Blip.Free;
end;

procedure TCT_TextBlipBullet.Clear;
begin
  FAssigneds := [];
  if FA_Blip <> Nil then 
    FreeAndNil(FA_Blip);
end;

procedure TCT_TextBlipBullet.Assign(AItem: TCT_TextBlipBullet);
begin
end;

procedure TCT_TextBlipBullet.CopyTo(AItem: TCT_TextBlipBullet);
begin
end;

function  TCT_TextBlipBullet.Create_A_Blip: TCT_Blip;
begin
  if FA_Blip = Nil then
    FA_Blip := TCT_Blip.Create(FOwner);
  Result := FA_Blip;
end;

{ TCT_TextTabStop }

function  TCT_TextTabStop.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FPos <> 2147483632 then 
    Inc(AttrsAssigned);
  if Integer(FAlgn) <> XPG_UNKNOWN_ENUM then
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_TextTabStop.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_TextTabStop.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FPos <> 2147483632 then 
    AWriter.AddAttribute('pos',XmlIntToStr(FPos));
  if Integer(FAlgn) <> XPG_UNKNOWN_ENUM then 
    AWriter.AddAttribute('algn',StrTST_TextTabAlignType[Integer(FAlgn)]);
end;

procedure TCT_TextTabStop.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $00000152: FPos := XmlStrToIntDef(AAttributes.Values[i],0);
      $000001A2: FAlgn := TST_TextTabAlignType(StrToEnum('stttat' + AAttributes.Values[i]));
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_TextTabStop.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 2;
  Clear;
end;

destructor TCT_TextTabStop.Destroy;
begin
end;

procedure TCT_TextTabStop.Clear;
begin
  FAssigneds := [];
  FPos := 2147483632;
  FAlgn := TST_TextTabAlignType(XPG_UNKNOWN_ENUM);
end;

procedure TCT_TextTabStop.Assign(AItem: TCT_TextTabStop);
begin
end;

procedure TCT_TextTabStop.CopyTo(AItem: TCT_TextTabStop);
begin
end;

{ TCT_TextTabStopXpgList }

function  TCT_TextTabStopXpgList.GetItems(Index: integer): TCT_TextTabStop;
begin
  Result := TCT_TextTabStop(inherited Items[Index]);
end;

function  TCT_TextTabStopXpgList.Add: TCT_TextTabStop;
begin
  Result := TCT_TextTabStop.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_TextTabStopXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_TextTabStopXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_TextTabStopXpgList.Assign(AItem: TCT_TextTabStopXpgList);
begin
end;

procedure TCT_TextTabStopXpgList.CopyTo(AItem: TCT_TextTabStopXpgList);
begin
end;

{ TEG_EffectProperties }

function  TEG_EffectProperties.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FA_EffectLst <> Nil then 
    Inc(ElemsAssigned,FA_EffectLst.CheckAssigned);
  if FA_EffectDag <> Nil then 
    Inc(ElemsAssigned,FA_EffectDag.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TEG_EffectProperties.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  case AReader.QNameHashA of
    $0000043B: begin
      if FA_EffectLst = Nil then 
        FA_EffectLst := TCT_EffectList.Create(FOwner);
      Result := FA_EffectLst;
    end;
    $00000414: begin
      if FA_EffectDag = Nil then 
        FA_EffectDag := TCT_EffectContainer.Create(FOwner);
      Result := FA_EffectDag;
    end;
    else 
    begin
      Result := Nil;
      Exit;
    end
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TEG_EffectProperties.Write(AWriter: TXpgWriteXML);
begin
  if (FA_EffectLst <> Nil) and FA_EffectLst.Assigned then 
    if xaElements in FA_EffectLst.FAssigneds then 
    begin
      AWriter.BeginTag('a:effectLst');
      FA_EffectLst.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:effectLst');
  if (FA_EffectDag <> Nil) and FA_EffectDag.Assigned then 
  begin
    FA_EffectDag.WriteAttributes(AWriter);
    if xaElements in FA_EffectDag.FAssigneds then 
    begin
      AWriter.BeginTag('a:effectDag');
      FA_EffectDag.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:effectDag');
  end;
end;

constructor TEG_EffectProperties.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 2;
  FAttributeCount := 0;
end;

destructor TEG_EffectProperties.Destroy;
begin
  if FA_EffectLst <> Nil then 
    FA_EffectLst.Free;
  if FA_EffectDag <> Nil then 
    FA_EffectDag.Free;
end;

procedure TEG_EffectProperties.Clear;
begin
  FAssigneds := [];
  if FA_EffectLst <> Nil then 
    FreeAndNil(FA_EffectLst);
  if FA_EffectDag <> Nil then 
    FreeAndNil(FA_EffectDag);
end;

procedure TEG_EffectProperties.Assign(AItem: TEG_EffectProperties);
begin
end;

procedure TEG_EffectProperties.CopyTo(AItem: TEG_EffectProperties);
begin
end;

function  TEG_EffectProperties.Create_A_EffectLst: TCT_EffectList;
begin
  if FA_EffectLst = Nil then
    FA_EffectLst := TCT_EffectList.Create(FOwner);
  Result := FA_EffectLst;
end;

function  TEG_EffectProperties.Create_A_EffectDag: TCT_EffectContainer;
begin
  if FA_EffectDag = Nil then
    FA_EffectDag := TCT_EffectContainer.Create(FOwner);
  Result := FA_EffectDag;
end;

{ TEG_TextUnderlineLine }

function  TEG_TextUnderlineLine.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FA_ULnTx <> Nil then 
    Inc(ElemsAssigned,FA_ULnTx.CheckAssigned);
  if FA_ULn <> Nil then 
    Inc(ElemsAssigned,FA_ULn.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TEG_TextUnderlineLine.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  case AReader.QNameHashA of
    $00000296: begin
      if FA_ULnTx = Nil then 
        FA_ULnTx := TCT_TextUnderlineLineFollowText.Create(FOwner);
      Result := FA_ULnTx;
    end;
    $000001CA: begin
      if FA_ULn = Nil then 
        FA_ULn := TCT_LineProperties.Create(FOwner);
      Result := FA_ULn;
    end;
    else 
    begin
      Result := Nil;
      Exit;
    end
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TEG_TextUnderlineLine.Write(AWriter: TXpgWriteXML);
begin
  if (FA_ULnTx <> Nil) and FA_ULnTx.Assigned then 
    AWriter.SimpleTag('a:uLnTx');
  if (FA_ULn <> Nil) and FA_ULn.Assigned then 
  begin
    FA_ULn.WriteAttributes(AWriter);
    if xaElements in FA_ULn.FAssigneds then 
    begin
      AWriter.BeginTag('a:uLn');
      FA_ULn.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:uLn');
  end;
end;

constructor TEG_TextUnderlineLine.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 2;
  FAttributeCount := 0;
end;

destructor TEG_TextUnderlineLine.Destroy;
begin
  if FA_ULnTx <> Nil then 
    FA_ULnTx.Free;
  if FA_ULn <> Nil then 
    FA_ULn.Free;
end;

procedure TEG_TextUnderlineLine.Clear;
begin
  FAssigneds := [];
  if FA_ULnTx <> Nil then 
    FreeAndNil(FA_ULnTx);
  if FA_ULn <> Nil then 
    FreeAndNil(FA_ULn);
end;

procedure TEG_TextUnderlineLine.Assign(AItem: TEG_TextUnderlineLine);
begin
end;

procedure TEG_TextUnderlineLine.CopyTo(AItem: TEG_TextUnderlineLine);
begin
end;

function  TEG_TextUnderlineLine.Create_A_ULnTx: TCT_TextUnderlineLineFollowText;
begin
  if FA_ULnTx = Nil then
    FA_ULnTx := TCT_TextUnderlineLineFollowText.Create(FOwner);
  Result := FA_ULnTx;
end;

function  TEG_TextUnderlineLine.Create_A_ULn: TCT_LineProperties;
begin
  if FA_ULn = Nil then
    FA_ULn := TCT_LineProperties.Create(FOwner);
  Result := FA_ULn;
end;

{ TEG_TextUnderlineFill }

function  TEG_TextUnderlineFill.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FA_UFillTx <> Nil then 
    Inc(ElemsAssigned,FA_UFillTx.CheckAssigned);
  if FA_UFill <> Nil then 
    Inc(ElemsAssigned,FA_UFill.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TEG_TextUnderlineFill.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  case AReader.QNameHashA of
    $00000363: begin
      if FA_UFillTx = Nil then 
        FA_UFillTx := TCT_TextUnderlineFillFollowText.Create(FOwner);
      Result := FA_UFillTx;
    end;
    $00000297: begin
      if FA_UFill = Nil then 
        FA_UFill := TCT_TextUnderlineFillGroupWrapper.Create(FOwner);
      Result := FA_UFill;
    end;
    else 
    begin
      Result := Nil;
      Exit;
    end
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TEG_TextUnderlineFill.Write(AWriter: TXpgWriteXML);
begin
  if (FA_UFillTx <> Nil) and FA_UFillTx.Assigned then 
    AWriter.SimpleTag('a:uFillTx');
  if (FA_UFill <> Nil) and FA_UFill.Assigned then 
    if xaElements in FA_UFill.FAssigneds then 
    begin
      AWriter.BeginTag('a:uFill');
      FA_UFill.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:uFill');
end;

constructor TEG_TextUnderlineFill.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 2;
  FAttributeCount := 0;
end;

destructor TEG_TextUnderlineFill.Destroy;
begin
  if FA_UFillTx <> Nil then 
    FA_UFillTx.Free;
  if FA_UFill <> Nil then 
    FA_UFill.Free;
end;

procedure TEG_TextUnderlineFill.Clear;
begin
  FAssigneds := [];
  if FA_UFillTx <> Nil then 
    FreeAndNil(FA_UFillTx);
  if FA_UFill <> Nil then 
    FreeAndNil(FA_UFill);
end;

procedure TEG_TextUnderlineFill.Assign(AItem: TEG_TextUnderlineFill);
begin
end;

procedure TEG_TextUnderlineFill.CopyTo(AItem: TEG_TextUnderlineFill);
begin
end;

function  TEG_TextUnderlineFill.Create_A_UFillTx: TCT_TextUnderlineFillFollowText;
begin
  if FA_UFillTx = Nil then
    FA_UFillTx := TCT_TextUnderlineFillFollowText.Create(FOwner);
  Result := FA_UFillTx;
end;

function  TEG_TextUnderlineFill.Create_A_UFill: TCT_TextUnderlineFillGroupWrapper;
begin
  if FA_UFill = Nil then
    FA_UFill := TCT_TextUnderlineFillGroupWrapper.Create(FOwner);
  Result := FA_UFill;
end;

{ TCT_Hyperlink }

function  TCT_Hyperlink.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FR_Id <> '' then 
    Inc(AttrsAssigned);
  if FInvalidUrl <> '' then 
    Inc(AttrsAssigned);
  if FAction <> '' then 
    Inc(AttrsAssigned);
  if FTgtFrame <> '' then 
    Inc(AttrsAssigned);
  if FTooltip <> '' then 
    Inc(AttrsAssigned);
  if FHistory <> True then 
    Inc(AttrsAssigned);
  if FHighlightClick <> False then 
    Inc(AttrsAssigned);
  if FEndSnd <> False then 
    Inc(AttrsAssigned);
  if FA_Snd <> Nil then 
    Inc(ElemsAssigned,FA_Snd.CheckAssigned);
  if FA_ExtLst <> Nil then 
    Inc(ElemsAssigned,FA_ExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_Hyperlink.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $000001E0: begin
      if FA_Snd = Nil then 
        FA_Snd := TCT_EmbeddedWAVAudioFile.Create(FOwner);
      Result := FA_Snd;
    end;
    $0000031F: begin
      if FA_ExtLst = Nil then 
        FA_ExtLst := TCT_OfficeArtExtensionList.Create(FOwner);
      Result := FA_ExtLst;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_Hyperlink.Write(AWriter: TXpgWriteXML);
begin
  if (FA_Snd <> Nil) and FA_Snd.Assigned then 
  begin
    FA_Snd.WriteAttributes(AWriter);
    AWriter.SimpleTag('a:snd');
  end;
  if (FA_ExtLst <> Nil) and FA_ExtLst.Assigned then 
    if xaElements in FA_ExtLst.FAssigneds then 
    begin
      AWriter.BeginTag('a:extLst');
      FA_ExtLst.Write(AWriter);
      AWriter.EndTag;
    end;
end;

procedure TCT_Hyperlink.WriteAttributes(AWriter: TXpgWriteXML);
begin
// TODO See comment at top
//  if FR_Id <> '' then
//    AWriter.AddAttribute('r:id',FR_Id);
//  if FInvalidUrl <> '' then
//    AWriter.AddAttribute('invalidUrl',FInvalidUrl);
//  if FAction <> '' then
//    AWriter.AddAttribute('action',FAction);
//  if FTgtFrame <> '' then
//    AWriter.AddAttribute('tgtFrame',FTgtFrame);
//  if FTooltip <> '' then
//    AWriter.AddAttribute('tooltip',FTooltip);
//  if FHistory <> True then
//    AWriter.AddAttribute('history',XmlBoolToStr(FHistory));
//  if FHighlightClick <> False then
//    AWriter.AddAttribute('highlightClick',XmlBoolToStr(FHighlightClick));
//  if FEndSnd <> False then
//    AWriter.AddAttribute('endSnd',XmlBoolToStr(FEndSnd));
end;

procedure TCT_Hyperlink.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $00000179: FR_Id := AAttributes.Values[i];
      $0000041A: FInvalidUrl := AAttributes.Values[i];
      $0000027E: FAction := AAttributes.Values[i];
      $0000033A: FTgtFrame := AAttributes.Values[i];
      $0000030B: FTooltip := AAttributes.Values[i];
      $00000312: FHistory := XmlStrToBoolDef(AAttributes.Values[i],True);
      $0000059E: FHighlightClick := XmlStrToBoolDef(AAttributes.Values[i],False);
      $0000025C: FEndSnd := XmlStrToBoolDef(AAttributes.Values[i],False);
//{$ifdef _AXOLOT_DEBUG}
//{$message warn '_TODO_ RId'}
//{$endif}
//      else
//        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_Hyperlink.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 2;
  FAttributeCount := 8;
  FHistory := True;
  FHighlightClick := False;
  FEndSnd := False;
end;

destructor TCT_Hyperlink.Destroy;
begin
  if FA_Snd <> Nil then 
    FA_Snd.Free;
  if FA_ExtLst <> Nil then 
    FA_ExtLst.Free;
end;

procedure TCT_Hyperlink.Clear;
begin
  FAssigneds := [];
  if FA_Snd <> Nil then 
    FreeAndNil(FA_Snd);
  if FA_ExtLst <> Nil then 
    FreeAndNil(FA_ExtLst);
  FR_Id := '';
  FInvalidUrl := '';
  FAction := '';
  FTgtFrame := '';
  FTooltip := '';
  FHistory := True;
  FHighlightClick := False;
  FEndSnd := False;
end;

procedure TCT_Hyperlink.Assign(AItem: TCT_Hyperlink);
begin
end;

procedure TCT_Hyperlink.CopyTo(AItem: TCT_Hyperlink);
begin
end;

function  TCT_Hyperlink.Create_A_Snd: TCT_EmbeddedWAVAudioFile;
begin
  if FA_Snd = Nil then
    FA_Snd := TCT_EmbeddedWAVAudioFile.Create(FOwner);
  Result := FA_Snd;
end;

function  TCT_Hyperlink.Create_A_ExtLst: TCT_OfficeArtExtensionList;
begin
  if FA_ExtLst = Nil then
    FA_ExtLst := TCT_OfficeArtExtensionList.Create(FOwner);
  Result := FA_ExtLst;
end;

{ TCT_Path2DClose }

function  TCT_Path2DClose.CheckAssigned: integer;
begin
  FAssigneds := [];
  Result := 0;
end;

procedure TCT_Path2DClose.Write(AWriter: TXpgWriteXML);
begin
end;

constructor TCT_Path2DClose.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 0;
end;

destructor TCT_Path2DClose.Destroy;
begin
end;

procedure TCT_Path2DClose.Clear;
begin
  FAssigneds := [];
end;

procedure TCT_Path2DClose.Assign(AItem: TCT_Path2DClose);
begin
end;

procedure TCT_Path2DClose.CopyTo(AItem: TCT_Path2DClose);
begin
end;

{ TCT_Path2DCloseXpgList }

function  TCT_Path2DCloseXpgList.GetItems(Index: integer): TCT_Path2DClose;
begin
  Result := TCT_Path2DClose(inherited Items[Index]);
end;

function  TCT_Path2DCloseXpgList.Add: TCT_Path2DClose;
begin
  Result := TCT_Path2DClose.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_Path2DCloseXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_Path2DCloseXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
end;

procedure TCT_Path2DCloseXpgList.Assign(AItem: TCT_Path2DCloseXpgList);
begin
end;

procedure TCT_Path2DCloseXpgList.CopyTo(AItem: TCT_Path2DCloseXpgList);
begin
end;

{ TCT_Path2DMoveTo }

function  TCT_Path2DMoveTo.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FA_Pt <> Nil then 
    Inc(ElemsAssigned,FA_Pt.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_Path2DMoveTo.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'a:pt' then 
  begin
    if FA_Pt = Nil then 
      FA_Pt := TCT_AdjPoint2D.Create(FOwner);
    Result := FA_Pt;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_Path2DMoveTo.Write(AWriter: TXpgWriteXML);
begin
  if (FA_Pt <> Nil) and FA_Pt.Assigned then 
  begin
    FA_Pt.WriteAttributes(AWriter);
    AWriter.SimpleTag('a:pt');
  end
  else 
    AWriter.SimpleTag('a:pt');
end;

constructor TCT_Path2DMoveTo.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 0;
end;

destructor TCT_Path2DMoveTo.Destroy;
begin
  if FA_Pt <> Nil then 
    FA_Pt.Free;
end;

procedure TCT_Path2DMoveTo.Clear;
begin
  FAssigneds := [];
  if FA_Pt <> Nil then 
    FreeAndNil(FA_Pt);
end;

procedure TCT_Path2DMoveTo.Assign(AItem: TCT_Path2DMoveTo);
begin
end;

procedure TCT_Path2DMoveTo.CopyTo(AItem: TCT_Path2DMoveTo);
begin
end;

function  TCT_Path2DMoveTo.Create_A_Pt: TCT_AdjPoint2D;
begin
  if FA_Pt = Nil then
    FA_Pt := TCT_AdjPoint2D.Create(FOwner);
  Result := FA_Pt;
end;

{ TCT_Path2DMoveToXpgList }

function  TCT_Path2DMoveToXpgList.GetItems(Index: integer): TCT_Path2DMoveTo;
begin
  Result := TCT_Path2DMoveTo(inherited Items[Index]);
end;

function  TCT_Path2DMoveToXpgList.Add: TCT_Path2DMoveTo;
begin
  Result := TCT_Path2DMoveTo.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_Path2DMoveToXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_Path2DMoveToXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
end;

procedure TCT_Path2DMoveToXpgList.Assign(AItem: TCT_Path2DMoveToXpgList);
begin
end;

procedure TCT_Path2DMoveToXpgList.CopyTo(AItem: TCT_Path2DMoveToXpgList);
begin
end;

{ TCT_Path2DLineTo }

function  TCT_Path2DLineTo.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FA_Pt <> Nil then 
    Inc(ElemsAssigned,FA_Pt.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_Path2DLineTo.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'a:pt' then 
  begin
    if FA_Pt = Nil then 
      FA_Pt := TCT_AdjPoint2D.Create(FOwner);
    Result := FA_Pt;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_Path2DLineTo.Write(AWriter: TXpgWriteXML);
begin
  if (FA_Pt <> Nil) and FA_Pt.Assigned then 
  begin
    FA_Pt.WriteAttributes(AWriter);
    AWriter.SimpleTag('a:pt');
  end
  else 
    AWriter.SimpleTag('a:pt');
end;

constructor TCT_Path2DLineTo.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 0;
end;

destructor TCT_Path2DLineTo.Destroy;
begin
  if FA_Pt <> Nil then 
    FA_Pt.Free;
end;

procedure TCT_Path2DLineTo.Clear;
begin
  FAssigneds := [];
  if FA_Pt <> Nil then 
    FreeAndNil(FA_Pt);
end;

procedure TCT_Path2DLineTo.Assign(AItem: TCT_Path2DLineTo);
begin
end;

procedure TCT_Path2DLineTo.CopyTo(AItem: TCT_Path2DLineTo);
begin
end;

function  TCT_Path2DLineTo.Create_A_Pt: TCT_AdjPoint2D;
begin
  if FA_Pt = Nil then
    FA_Pt := TCT_AdjPoint2D.Create(FOwner);
  Result := FA_Pt;
end;

{ TCT_Path2DLineToXpgList }

function  TCT_Path2DLineToXpgList.GetItems(Index: integer): TCT_Path2DLineTo;
begin
  Result := TCT_Path2DLineTo(inherited Items[Index]);
end;

function  TCT_Path2DLineToXpgList.Add: TCT_Path2DLineTo;
begin
  Result := TCT_Path2DLineTo.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_Path2DLineToXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_Path2DLineToXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
end;

procedure TCT_Path2DLineToXpgList.Assign(AItem: TCT_Path2DLineToXpgList);
begin
end;

procedure TCT_Path2DLineToXpgList.CopyTo(AItem: TCT_Path2DLineToXpgList);
begin
end;

{ TCT_Path2DArcTo }

function  TCT_Path2DArcTo.CheckAssigned: integer;
begin
  FAssigneds := [xaAttributes];
  Result := 1;
end;

procedure TCT_Path2DArcTo.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_Path2DArcTo.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('wR',WriteUnionTST_AdjCoordinate(PFWR));
  AWriter.AddAttribute('hR',WriteUnionTST_AdjCoordinate(PFHR));
  AWriter.AddAttribute('stAng',WriteUnionTST_AdjAngle(PFStAng));
  AWriter.AddAttribute('swAng',WriteUnionTST_AdjAngle(PFSwAng));
end;

procedure TCT_Path2DArcTo.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $000000C9: ReadUnionTST_AdjCoordinate(AAttributes.Values[i],PFWR);
      $000000BA: ReadUnionTST_AdjCoordinate(AAttributes.Values[i],PFHR);
      $000001FD: ReadUnionTST_AdjAngle(AAttributes.Values[i],PFStAng);
      $00000200: ReadUnionTST_AdjAngle(AAttributes.Values[i],PFSwAng);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_Path2DArcTo.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 4;
  PFWR := @FWR;
  PWR.nVal := -1;
  PFHR := @FHR;
  PHR.nVal := -1;
  PFStAng := @FStAng;
  PStAng.nVal := -1;
  PFSwAng := @FSwAng;
  PSwAng.nVal := -1;
end;

destructor TCT_Path2DArcTo.Destroy;
begin
end;

procedure TCT_Path2DArcTo.Clear;
begin
  FAssigneds := [];
end;

procedure TCT_Path2DArcTo.Assign(AItem: TCT_Path2DArcTo);
begin
end;

procedure TCT_Path2DArcTo.CopyTo(AItem: TCT_Path2DArcTo);
begin
end;

{ TCT_Path2DArcToXpgList }

function  TCT_Path2DArcToXpgList.GetItems(Index: integer): TCT_Path2DArcTo;
begin
  Result := TCT_Path2DArcTo(inherited Items[Index]);
end;

function  TCT_Path2DArcToXpgList.Add: TCT_Path2DArcTo;
begin
  Result := TCT_Path2DArcTo.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_Path2DArcToXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_Path2DArcToXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_Path2DArcToXpgList.Assign(AItem: TCT_Path2DArcToXpgList);
begin
end;

procedure TCT_Path2DArcToXpgList.CopyTo(AItem: TCT_Path2DArcToXpgList);
begin
end;

{ TCT_Path2DQuadBezierTo }

function  TCT_Path2DQuadBezierTo.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FA_PtXpgList <> Nil then 
    Inc(ElemsAssigned,FA_PtXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_Path2DQuadBezierTo.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'a:pt' then 
  begin
    if FA_PtXpgList = Nil then 
      FA_PtXpgList := TCT_AdjPoint2DXpgList.Create(FOwner);
    Result := FA_PtXpgList.Add;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_Path2DQuadBezierTo.Write(AWriter: TXpgWriteXML);
begin
  if FA_PtXpgList <> Nil then
    FA_PtXpgList.Write(AWriter,'a:pt');
end;

constructor TCT_Path2DQuadBezierTo.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 0;
end;

destructor TCT_Path2DQuadBezierTo.Destroy;
begin
  if FA_PtXpgList <> Nil then 
    FA_PtXpgList.Free;
end;

procedure TCT_Path2DQuadBezierTo.Clear;
begin
  FAssigneds := [];
  if FA_PtXpgList <> Nil then 
    FreeAndNil(FA_PtXpgList);
end;

procedure TCT_Path2DQuadBezierTo.Assign(AItem: TCT_Path2DQuadBezierTo);
begin
end;

procedure TCT_Path2DQuadBezierTo.CopyTo(AItem: TCT_Path2DQuadBezierTo);
begin
end;

{ TCT_Path2DQuadBezierToXpgList }

function  TCT_Path2DQuadBezierToXpgList.GetItems(Index: integer): TCT_Path2DQuadBezierTo;
begin
  Result := TCT_Path2DQuadBezierTo(inherited Items[Index]);
end;

function  TCT_Path2DQuadBezierToXpgList.Add: TCT_Path2DQuadBezierTo;
begin
  Result := TCT_Path2DQuadBezierTo.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_Path2DQuadBezierToXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_Path2DQuadBezierToXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
end;

procedure TCT_Path2DQuadBezierToXpgList.Assign(AItem: TCT_Path2DQuadBezierToXpgList);
begin
end;

procedure TCT_Path2DQuadBezierToXpgList.CopyTo(AItem: TCT_Path2DQuadBezierToXpgList);
begin
end;

{ TCT_Path2DCubicBezierTo }

function  TCT_Path2DCubicBezierTo.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FA_PtXpgList <> Nil then
    Inc(ElemsAssigned,FA_PtXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_Path2DCubicBezierTo.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'a:pt' then 
  begin
    if FA_PtXpgList = Nil then
      FA_PtXpgList := TCT_AdjPoint2DXpgList.Create(FOwner);
    Result := FA_PtXpgList.Add;
  end
  else
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then
    Result.Assigneds := [xaRead];
end;

procedure TCT_Path2DCubicBezierTo.Write(AWriter: TXpgWriteXML);
begin
  if FA_PtXpgList <> Nil then
    FA_PtXpgList.Write(AWriter,'a:pt');
end;

constructor TCT_Path2DCubicBezierTo.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 0;
end;

destructor TCT_Path2DCubicBezierTo.Destroy;
begin
  if FA_PtXpgList <> Nil then
    FA_PtXpgList.Free;
end;

procedure TCT_Path2DCubicBezierTo.Clear;
begin
  FAssigneds := [];
  if FA_PtXpgList <> Nil then
    FreeAndNil(FA_PtXpgList);
end;

procedure TCT_Path2DCubicBezierTo.Assign(AItem: TCT_Path2DCubicBezierTo);
begin
end;

procedure TCT_Path2DCubicBezierTo.CopyTo(AItem: TCT_Path2DCubicBezierTo);
begin
end;

{ TCT_Path2DCubicBezierToXpgList }

function  TCT_Path2DCubicBezierToXpgList.GetItems(Index: integer): TCT_Path2DCubicBezierTo;
begin
  Result := TCT_Path2DCubicBezierTo(inherited Items[Index]);
end;

function  TCT_Path2DCubicBezierToXpgList.Add: TCT_Path2DCubicBezierTo;
begin
  Result := TCT_Path2DCubicBezierTo.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_Path2DCubicBezierToXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_Path2DCubicBezierToXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
end;

procedure TCT_Path2DCubicBezierToXpgList.Assign(AItem: TCT_Path2DCubicBezierToXpgList);
begin
end;

procedure TCT_Path2DCubicBezierToXpgList.CopyTo(AItem: TCT_Path2DCubicBezierToXpgList);
begin
end;

{ TCT_TextSpacing }

function  TCT_TextSpacing.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FA_SpcPct <> Nil then 
    Inc(ElemsAssigned,FA_SpcPct.CheckAssigned);
  if FA_SpcPts <> Nil then 
    Inc(ElemsAssigned,FA_SpcPts.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_TextSpacing.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $00000308: begin
      if FA_SpcPct = Nil then 
        FA_SpcPct := TCT_TextSpacingPercent.Create(FOwner);
      Result := FA_SpcPct;
    end;
    $00000318: begin
      if FA_SpcPts = Nil then 
        FA_SpcPts := TCT_TextSpacingPoint.Create(FOwner);
      Result := FA_SpcPts;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_TextSpacing.Write(AWriter: TXpgWriteXML);
begin
  if (FA_SpcPct <> Nil) and FA_SpcPct.Assigned then 
  begin
    FA_SpcPct.WriteAttributes(AWriter);
    AWriter.SimpleTag('a:spcPct');
  end;
//  else
//    AWriter.SimpleTag('a:spcPct');
  if (FA_SpcPts <> Nil) and FA_SpcPts.Assigned then
  begin
    FA_SpcPts.WriteAttributes(AWriter);
    AWriter.SimpleTag('a:spcPts');
  end
  else 
    AWriter.SimpleTag('a:spcPts');
end;

constructor TCT_TextSpacing.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 2;
  FAttributeCount := 0;
end;

destructor TCT_TextSpacing.Destroy;
begin
  if FA_SpcPct <> Nil then 
    FA_SpcPct.Free;
  if FA_SpcPts <> Nil then 
    FA_SpcPts.Free;
end;

procedure TCT_TextSpacing.Clear;
begin
  FAssigneds := [];
  if FA_SpcPct <> Nil then 
    FreeAndNil(FA_SpcPct);
  if FA_SpcPts <> Nil then 
    FreeAndNil(FA_SpcPts);
end;

procedure TCT_TextSpacing.Assign(AItem: TCT_TextSpacing);
begin
end;

procedure TCT_TextSpacing.CopyTo(AItem: TCT_TextSpacing);
begin
end;

function  TCT_TextSpacing.Create_A_SpcPct: TCT_TextSpacingPercent;
begin
  if FA_SpcPct = Nil then
    FA_SpcPct := TCT_TextSpacingPercent.Create(FOwner);
  Result := FA_SpcPct;
end;

function  TCT_TextSpacing.Create_A_SpcPts: TCT_TextSpacingPoint;
begin
  if FA_SpcPts = Nil then
    FA_SpcPts := TCT_TextSpacingPoint.Create(FOwner);
  Result := FA_SpcPts;
end;

{ TEG_TextBulletColor }

function  TEG_TextBulletColor.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FA_BuClrTx <> Nil then 
    Inc(ElemsAssigned,FA_BuClrTx.CheckAssigned);
  if FA_BuClr <> Nil then 
    Inc(ElemsAssigned,FA_BuClr.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TEG_TextBulletColor.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  case AReader.QNameHashA of
    $0000035F: begin
      if FA_BuClrTx = Nil then 
        FA_BuClrTx := TCT_TextBulletColorFollowText.Create(FOwner);
      Result := FA_BuClrTx;
    end;
    $00000293: begin
      if FA_BuClr = Nil then 
        FA_BuClr := TCT_Color.Create(FOwner);
      Result := FA_BuClr;
    end;
    else 
    begin
      Result := Nil;
      Exit;
    end
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TEG_TextBulletColor.Write(AWriter: TXpgWriteXML);
begin
  if (FA_BuClrTx <> Nil) and FA_BuClrTx.Assigned then 
    AWriter.SimpleTag('a:buClrTx');
  if (FA_BuClr <> Nil) and FA_BuClr.Assigned then 
    if xaElements in FA_BuClr.FAssigneds then 
    begin
      AWriter.BeginTag('a:buClr');
      FA_BuClr.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:buClr');
end;

constructor TEG_TextBulletColor.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 2;
  FAttributeCount := 0;
end;

destructor TEG_TextBulletColor.Destroy;
begin
  if FA_BuClrTx <> Nil then 
    FA_BuClrTx.Free;
  if FA_BuClr <> Nil then 
    FA_BuClr.Free;
end;

procedure TEG_TextBulletColor.Clear;
begin
  FAssigneds := [];
  if FA_BuClrTx <> Nil then 
    FreeAndNil(FA_BuClrTx);
  if FA_BuClr <> Nil then 
    FreeAndNil(FA_BuClr);
end;

procedure TEG_TextBulletColor.Assign(AItem: TEG_TextBulletColor);
begin
end;

procedure TEG_TextBulletColor.CopyTo(AItem: TEG_TextBulletColor);
begin
end;

function  TEG_TextBulletColor.Create_A_BuClrTx: TCT_TextBulletColorFollowText;
begin
  if FA_BuClrTx = Nil then
    FA_BuClrTx := TCT_TextBulletColorFollowText.Create(FOwner);
  Result := FA_BuClrTx;
end;

function  TEG_TextBulletColor.Create_A_BuClr: TCT_Color;
begin
  if FA_BuClr = Nil then
    FA_BuClr := TCT_Color.Create(FOwner);
  Result := FA_BuClr;
end;

{ TEG_TextBulletSize }

function  TEG_TextBulletSize.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FA_BuSzTx <> Nil then 
    Inc(ElemsAssigned,FA_BuSzTx.CheckAssigned);
  if FA_BuSzPct <> Nil then 
    Inc(ElemsAssigned,FA_BuSzPct.CheckAssigned);
  if FA_BuSzPts <> Nil then 
    Inc(ElemsAssigned,FA_BuSzPts.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TEG_TextBulletSize.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  case AReader.QNameHashA of
    $0000030B: begin
      if FA_BuSzTx = Nil then 
        FA_BuSzTx := TCT_TextBulletSizeFollowText.Create(FOwner);
      Result := FA_BuSzTx;
    end;
    $00000366: begin
      if FA_BuSzPct = Nil then 
        FA_BuSzPct := TCT_TextBulletSizePercent.Create(FOwner);
      Result := FA_BuSzPct;
    end;
    $00000376: begin
      if FA_BuSzPts = Nil then 
        FA_BuSzPts := TCT_TextBulletSizePoint.Create(FOwner);
      Result := FA_BuSzPts;
    end;
    else 
    begin
      Result := Nil;
      Exit;
    end
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TEG_TextBulletSize.Write(AWriter: TXpgWriteXML);
begin
  if (FA_BuSzTx <> Nil) and FA_BuSzTx.Assigned then 
    AWriter.SimpleTag('a:buSzTx');
  if (FA_BuSzPct <> Nil) and FA_BuSzPct.Assigned then 
  begin
    FA_BuSzPct.WriteAttributes(AWriter);
    AWriter.SimpleTag('a:buSzPct');
  end;
  if (FA_BuSzPts <> Nil) and FA_BuSzPts.Assigned then 
  begin
    FA_BuSzPts.WriteAttributes(AWriter);
    AWriter.SimpleTag('a:buSzPts');
  end;
end;

constructor TEG_TextBulletSize.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 3;
  FAttributeCount := 0;
end;

destructor TEG_TextBulletSize.Destroy;
begin
  if FA_BuSzTx <> Nil then 
    FA_BuSzTx.Free;
  if FA_BuSzPct <> Nil then 
    FA_BuSzPct.Free;
  if FA_BuSzPts <> Nil then 
    FA_BuSzPts.Free;
end;

procedure TEG_TextBulletSize.Clear;
begin
  FAssigneds := [];
  if FA_BuSzTx <> Nil then 
    FreeAndNil(FA_BuSzTx);
  if FA_BuSzPct <> Nil then 
    FreeAndNil(FA_BuSzPct);
  if FA_BuSzPts <> Nil then 
    FreeAndNil(FA_BuSzPts);
end;

procedure TEG_TextBulletSize.Assign(AItem: TEG_TextBulletSize);
begin
end;

procedure TEG_TextBulletSize.CopyTo(AItem: TEG_TextBulletSize);
begin
end;

function  TEG_TextBulletSize.Create_A_BuSzTx: TCT_TextBulletSizeFollowText;
begin
  if FA_BuSzTx = Nil then
    FA_BuSzTx := TCT_TextBulletSizeFollowText.Create(FOwner);
  Result := FA_BuSzTx;
end;

function  TEG_TextBulletSize.Create_A_BuSzPct: TCT_TextBulletSizePercent;
begin
  if FA_BuSzPct = Nil then
    FA_BuSzPct := TCT_TextBulletSizePercent.Create(FOwner);
  Result := FA_BuSzPct;
end;

function  TEG_TextBulletSize.Create_A_BuSzPts: TCT_TextBulletSizePoint;
begin
  if FA_BuSzPts = Nil then
    FA_BuSzPts := TCT_TextBulletSizePoint.Create(FOwner);
  Result := FA_BuSzPts;
end;

{ TEG_TextBulletTypeface }

function  TEG_TextBulletTypeface.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FA_BuFontTx <> Nil then 
    Inc(ElemsAssigned,FA_BuFontTx.CheckAssigned);
  if FA_BuFont <> Nil then 
    Inc(ElemsAssigned,FA_BuFont.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TEG_TextBulletTypeface.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  case AReader.QNameHashA of
    $000003D5: begin
      if FA_BuFontTx = Nil then 
        FA_BuFontTx := TCT_TextBulletTypefaceFollowText.Create(FOwner);
      Result := FA_BuFontTx;
    end;
    $00000309: begin
      if FA_BuFont = Nil then 
        FA_BuFont := TCT_TextFont.Create(FOwner);
      Result := FA_BuFont;
    end;
    else 
    begin
      Result := Nil;
      Exit;
    end
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TEG_TextBulletTypeface.Write(AWriter: TXpgWriteXML);
begin
  if (FA_BuFontTx <> Nil) and FA_BuFontTx.Assigned then 
    AWriter.SimpleTag('a:buFontTx');
  if (FA_BuFont <> Nil) and FA_BuFont.Assigned then 
  begin
    FA_BuFont.WriteAttributes(AWriter);
    AWriter.SimpleTag('a:buFont');
  end;
end;

constructor TEG_TextBulletTypeface.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 2;
  FAttributeCount := 0;
end;

destructor TEG_TextBulletTypeface.Destroy;
begin
  if FA_BuFontTx <> Nil then 
    FA_BuFontTx.Free;
  if FA_BuFont <> Nil then 
    FA_BuFont.Free;
end;

procedure TEG_TextBulletTypeface.Clear;
begin
  FAssigneds := [];
  if FA_BuFontTx <> Nil then 
    FreeAndNil(FA_BuFontTx);
  if FA_BuFont <> Nil then 
    FreeAndNil(FA_BuFont);
end;

procedure TEG_TextBulletTypeface.Assign(AItem: TEG_TextBulletTypeface);
begin
end;

procedure TEG_TextBulletTypeface.CopyTo(AItem: TEG_TextBulletTypeface);
begin
end;

function  TEG_TextBulletTypeface.Create_A_BuFontTx: TCT_TextBulletTypefaceFollowText;
begin
  if FA_BuFontTx = Nil then
    FA_BuFontTx := TCT_TextBulletTypefaceFollowText.Create(FOwner);
  Result := FA_BuFontTx;
end;

function  TEG_TextBulletTypeface.Create_A_BuFont: TCT_TextFont;
begin
  if FA_BuFont = Nil then
    FA_BuFont := TCT_TextFont.Create(FOwner);
  Result := FA_BuFont;
end;

{ TEG_TextBullet }

function  TEG_TextBullet.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FA_BuNone <> Nil then 
    Inc(ElemsAssigned,FA_BuNone.CheckAssigned);
  if FA_BuAutoNum <> Nil then 
    Inc(ElemsAssigned,FA_BuAutoNum.CheckAssigned);
  if FA_BuChar <> Nil then 
    Inc(ElemsAssigned,FA_BuChar.CheckAssigned);
  if FA_BuBlip <> Nil then 
    Inc(ElemsAssigned,FA_BuBlip.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TEG_TextBullet.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  case AReader.QNameHashA of
    $00000302: begin
      if FA_BuNone = Nil then 
        FA_BuNone := TCT_TextNoBullet.Create(FOwner);
      Result := FA_BuNone;
    end;
    $0000043B: begin
      if FA_BuAutoNum = Nil then 
        FA_BuAutoNum := TCT_TextAutonumberBullet.Create(FOwner);
      Result := FA_BuAutoNum;
    end;
    $000002F0: begin
      if FA_BuChar = Nil then 
        FA_BuChar := TCT_TextCharBullet.Create(FOwner);
      Result := FA_BuChar;
    end;
    $000002F9: begin
      if FA_BuBlip = Nil then 
        FA_BuBlip := TCT_TextBlipBullet.Create(FOwner);
      Result := FA_BuBlip;
    end;
    else 
    begin
      Result := Nil;
      Exit;
    end
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TEG_TextBullet.Write(AWriter: TXpgWriteXML);
begin
  if (FA_BuNone <> Nil) and FA_BuNone.Assigned then 
    AWriter.SimpleTag('a:buNone');
  if (FA_BuAutoNum <> Nil) and FA_BuAutoNum.Assigned then 
  begin
    FA_BuAutoNum.WriteAttributes(AWriter);
    AWriter.SimpleTag('a:buAutoNum');
  end;
  if (FA_BuChar <> Nil) and FA_BuChar.Assigned then 
  begin
    FA_BuChar.WriteAttributes(AWriter);
    AWriter.SimpleTag('a:buChar');
  end;
  if (FA_BuBlip <> Nil) and FA_BuBlip.Assigned then 
    if xaElements in FA_BuBlip.FAssigneds then 
    begin
      AWriter.BeginTag('a:buBlip');
      FA_BuBlip.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:buBlip');
end;

constructor TEG_TextBullet.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 4;
  FAttributeCount := 0;
end;

destructor TEG_TextBullet.Destroy;
begin
  if FA_BuNone <> Nil then 
    FA_BuNone.Free;
  if FA_BuAutoNum <> Nil then 
    FA_BuAutoNum.Free;
  if FA_BuChar <> Nil then 
    FA_BuChar.Free;
  if FA_BuBlip <> Nil then 
    FA_BuBlip.Free;
end;

procedure TEG_TextBullet.Clear;
begin
  FAssigneds := [];
  if FA_BuNone <> Nil then 
    FreeAndNil(FA_BuNone);
  if FA_BuAutoNum <> Nil then 
    FreeAndNil(FA_BuAutoNum);
  if FA_BuChar <> Nil then 
    FreeAndNil(FA_BuChar);
  if FA_BuBlip <> Nil then 
    FreeAndNil(FA_BuBlip);
end;

procedure TEG_TextBullet.Assign(AItem: TEG_TextBullet);
begin
end;

procedure TEG_TextBullet.CopyTo(AItem: TEG_TextBullet);
begin
end;

function  TEG_TextBullet.Create_A_BuNone: TCT_TextNoBullet;
begin
  if FA_BuNone = Nil then
    FA_BuNone := TCT_TextNoBullet.Create(FOwner);
  Result := FA_BuNone;
end;

function  TEG_TextBullet.Create_A_BuAutoNum: TCT_TextAutonumberBullet;
begin
  if FA_BuAutoNum = Nil then
    FA_BuAutoNum := TCT_TextAutonumberBullet.Create(FOwner);
  Result := FA_BuAutoNum;
end;

function  TEG_TextBullet.Create_A_BuChar: TCT_TextCharBullet;
begin
  if FA_BuChar = Nil then
    FA_BuChar := TCT_TextCharBullet.Create(FOwner);
  Result := FA_BuChar;
end;

function  TEG_TextBullet.Create_A_BuBlip: TCT_TextBlipBullet;
begin
  if FA_BuBlip = Nil then
    FA_BuBlip := TCT_TextBlipBullet.Create(FOwner);
  Result := FA_BuBlip;
end;

{ TCT_TextTabStopList }

function  TCT_TextTabStopList.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FA_TabXpgList <> Nil then 
    Inc(ElemsAssigned,FA_TabXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_TextTabStopList.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'a:tab' then 
  begin
    if FA_TabXpgList = Nil then 
      FA_TabXpgList := TCT_TextTabStopXpgList.Create(FOwner);
    Result := FA_TabXpgList.Add;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_TextTabStopList.Write(AWriter: TXpgWriteXML);
begin
  if FA_TabXpgList <> Nil then 
    FA_TabXpgList.Write(AWriter,'a:tab');
end;

constructor TCT_TextTabStopList.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 0;
end;

destructor TCT_TextTabStopList.Destroy;
begin
  if FA_TabXpgList <> Nil then 
    FA_TabXpgList.Free;
end;

procedure TCT_TextTabStopList.Clear;
begin
  FAssigneds := [];
  if FA_TabXpgList <> Nil then 
    FreeAndNil(FA_TabXpgList);
end;

procedure TCT_TextTabStopList.Assign(AItem: TCT_TextTabStopList);
begin
end;

procedure TCT_TextTabStopList.CopyTo(AItem: TCT_TextTabStopList);
begin
end;

function  TCT_TextTabStopList.Create_A_TabXpgList: TCT_TextTabStopXpgList;
begin
  if FA_TabXpgList = Nil then
    FA_TabXpgList := TCT_TextTabStopXpgList.Create(FOwner);
  Result := FA_TabXpgList;
end;

{ TCT_TextCharacterProperties }

function  TCT_TextCharacterProperties.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if Byte(FKumimoji) <> 2 then 
    Inc(AttrsAssigned);
  if FLang <> '' then 
    Inc(AttrsAssigned);
  if FAltLang <> '' then 
    Inc(AttrsAssigned);
  if FSz <> 2147483632 then 
    Inc(AttrsAssigned);
  if Byte(FB) <> 2 then 
    Inc(AttrsAssigned);
  if Byte(FI) <> 2 then 
    Inc(AttrsAssigned);
  if Integer(FU) <> XPG_UNKNOWN_ENUM then 
    Inc(AttrsAssigned);
  if Integer(FStrike) <> XPG_UNKNOWN_ENUM then 
    Inc(AttrsAssigned);
  if FKern <> 2147483632 then 
    Inc(AttrsAssigned);
  if Integer(FCap) <> XPG_UNKNOWN_ENUM then 
    Inc(AttrsAssigned);
  if FSpc <> 2147483632 then 
    Inc(AttrsAssigned);
  if Byte(FNormalizeH) <> 2 then 
    Inc(AttrsAssigned);
  if FBaseline <> 2147483632 then 
    Inc(AttrsAssigned);
  if Byte(FNoProof) <> 2 then 
    Inc(AttrsAssigned);
  if FDirty <> True then 
    Inc(AttrsAssigned);
  if FErr <> False then 
    Inc(AttrsAssigned);
  if FSmtClean <> True then 
    Inc(AttrsAssigned);
  if FSmtId <> 0 then 
    Inc(AttrsAssigned);
  if FBmk <> '' then 
    Inc(AttrsAssigned);
  if FA_Ln <> Nil then 
    Inc(ElemsAssigned,FA_Ln.CheckAssigned);
  Inc(ElemsAssigned,FA_EG_FillProperties.CheckAssigned);
  Inc(ElemsAssigned,FA_EG_EffectProperties.CheckAssigned);
  if FA_Highlight <> Nil then 
    Inc(ElemsAssigned,FA_Highlight.CheckAssigned);
  Inc(ElemsAssigned,FA_EG_TextUnderlineLine.CheckAssigned);
  Inc(ElemsAssigned,FA_EG_TextUnderlineFill.CheckAssigned);
  if FA_Latin <> Nil then 
    Inc(ElemsAssigned,FA_Latin.CheckAssigned);
  if FA_Ea <> Nil then 
    Inc(ElemsAssigned,FA_Ea.CheckAssigned);
  if FA_Cs <> Nil then 
    Inc(ElemsAssigned,FA_Cs.CheckAssigned);
  if FA_Sym <> Nil then 
    Inc(ElemsAssigned,FA_Sym.CheckAssigned);
  if FA_HlinkClick <> Nil then 
    Inc(ElemsAssigned,FA_HlinkClick.CheckAssigned);
  if FA_HlinkMouseOver <> Nil then 
    Inc(ElemsAssigned,FA_HlinkMouseOver.CheckAssigned);
  if FA_ExtLst <> Nil then 
    Inc(ElemsAssigned,FA_ExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_TextCharacterProperties.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  case AReader.QNameHashA of
    $00000175: begin
      if FA_Ln = Nil then 
        FA_Ln := TCT_LineProperties.Create(FOwner);
      Result := FA_Ln;
    end;
    $00000453: begin
      if FA_Highlight = Nil then 
        FA_Highlight := TCT_Color.Create(FOwner);
      Result := FA_Highlight;
    end;
    $000002B3: begin
      if FA_Latin = Nil then 
        FA_Latin := TCT_TextFont.Create(FOwner);
      Result := FA_Latin;
    end;
    $00000161: begin
      if FA_Ea = Nil then 
        FA_Ea := TCT_TextFont.Create(FOwner);
      Result := FA_Ea;
    end;
    $00000171: begin
      if FA_Cs = Nil then 
        FA_Cs := TCT_TextFont.Create(FOwner);
      Result := FA_Cs;
    end;
    $000001F4: begin
      if FA_Sym = Nil then 
        FA_Sym := TCT_TextFont.Create(FOwner);
      Result := FA_Sym;
    end;
    $00000497: begin
      if FA_HlinkClick = Nil then 
        FA_HlinkClick := TCT_Hyperlink.Create(FOwner);
      Result := FA_HlinkClick;
    end;
    $00000656: begin
      if FA_HlinkMouseOver = Nil then 
        FA_HlinkMouseOver := TCT_Hyperlink.Create(FOwner);
      Result := FA_HlinkMouseOver;
    end;
    $0000031F: begin
      if FA_ExtLst = Nil then 
        FA_ExtLst := TCT_OfficeArtExtensionList.Create(FOwner);
      Result := FA_ExtLst;
    end;
    else 
    begin
      Result := FA_EG_FillProperties.HandleElement(AReader);
      if Result = Nil then 
      begin
        Result := FA_EG_EffectProperties.HandleElement(AReader);
        if Result = Nil then 
        begin
          Result := FA_EG_TextUnderlineLine.HandleElement(AReader);
          if Result = Nil then 
          begin
            Result := FA_EG_TextUnderlineFill.HandleElement(AReader);
            if Result = Nil then 
              FOwner.Errors.Error(xemUnknownElement,AReader.QName);
          end;
        end;
      end;
    end
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_TextCharacterProperties.Write(AWriter: TXpgWriteXML);
begin
  if (FA_Ln <> Nil) and FA_Ln.Assigned then 
  begin
    FA_Ln.WriteAttributes(AWriter);
    if xaElements in FA_Ln.FAssigneds then 
    begin
      AWriter.BeginTag('a:ln');
      FA_Ln.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:ln');
  end;
  FA_EG_FillProperties.Write(AWriter);
  FA_EG_EffectProperties.Write(AWriter);
  if (FA_Highlight <> Nil) and FA_Highlight.Assigned then 
    if xaElements in FA_Highlight.FAssigneds then 
    begin
      AWriter.BeginTag('a:highlight');
      FA_Highlight.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:highlight');
  FA_EG_TextUnderlineLine.Write(AWriter);
  FA_EG_TextUnderlineFill.Write(AWriter);
  if (FA_Latin <> Nil) and FA_Latin.Assigned then 
  begin
    FA_Latin.WriteAttributes(AWriter);
    AWriter.SimpleTag('a:latin');
  end;
  if (FA_Ea <> Nil) and FA_Ea.Assigned then 
  begin
    FA_Ea.WriteAttributes(AWriter);
    AWriter.SimpleTag('a:ea');
  end;
  if (FA_Cs <> Nil) and FA_Cs.Assigned then 
  begin
    FA_Cs.WriteAttributes(AWriter);
    AWriter.SimpleTag('a:cs');
  end;
  if (FA_Sym <> Nil) and FA_Sym.Assigned then 
  begin
    FA_Sym.WriteAttributes(AWriter);
    AWriter.SimpleTag('a:sym');
  end;
  if (FA_HlinkClick <> Nil) and FA_HlinkClick.Assigned then 
  begin
    FA_HlinkClick.WriteAttributes(AWriter);
    if xaElements in FA_HlinkClick.FAssigneds then 
    begin
      AWriter.BeginTag('a:hlinkClick');
      FA_HlinkClick.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:hlinkClick');
  end;
  if (FA_HlinkMouseOver <> Nil) and FA_HlinkMouseOver.Assigned then 
  begin
    FA_HlinkMouseOver.WriteAttributes(AWriter);
    if xaElements in FA_HlinkMouseOver.FAssigneds then 
    begin
      AWriter.BeginTag('a:hlinkMouseOver');
      FA_HlinkMouseOver.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:hlinkMouseOver');
  end;
  if (FA_ExtLst <> Nil) and FA_ExtLst.Assigned then 
    if xaElements in FA_ExtLst.FAssigneds then 
    begin
      AWriter.BeginTag('a:extLst');
      FA_ExtLst.Write(AWriter);
      AWriter.EndTag;
    end;
end;

procedure TCT_TextCharacterProperties.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if Byte(FKumimoji) <> 2 then 
    AWriter.AddAttribute('kumimoji',XmlBoolToStr(FKumimoji));
  if FLang <> '' then 
    AWriter.AddAttribute('lang',FLang);
  if FAltLang <> '' then 
    AWriter.AddAttribute('altLang',FAltLang);
  if FSz <> 2147483632 then 
    AWriter.AddAttribute('sz',XmlIntToStr(FSz));
  if Byte(FB) = 1 then
    AWriter.AddAttribute('b',XmlBoolToStr(FB));
  if Byte(FI) = 1 then
    AWriter.AddAttribute('i',XmlBoolToStr(FI));
  if Integer(FU) <> XPG_UNKNOWN_ENUM then 
    AWriter.AddAttribute('u',StrTST_TextUnderlineType[Integer(FU)]);
  if Integer(FStrike) <> XPG_UNKNOWN_ENUM then 
    AWriter.AddAttribute('strike',StrTST_TextStrikeType[Integer(FStrike)]);
  if FKern <> 2147483632 then 
    AWriter.AddAttribute('kern',XmlIntToStr(FKern));
  if Integer(FCap) <> XPG_UNKNOWN_ENUM then 
    AWriter.AddAttribute('cap',StrTST_TextCapsType[Integer(FCap)]);
  if FSpc <> 2147483632 then 
    AWriter.AddAttribute('spc',XmlIntToStr(FSpc));
  if Byte(FNormalizeH) <> 2 then 
    AWriter.AddAttribute('normalizeH',XmlBoolToStr(FNormalizeH));
  if FBaseline <> 2147483632 then 
    AWriter.AddAttribute('baseline',XmlIntToStr(FBaseline));
  if Byte(FNoProof) <> 2 then 
    AWriter.AddAttribute('noProof',XmlBoolToStr(FNoProof));
  if FDirty <> True then 
    AWriter.AddAttribute('dirty',XmlBoolToStr(FDirty));
  if FErr <> False then 
    AWriter.AddAttribute('err',XmlBoolToStr(FErr));
  if FSmtClean <> True then 
    AWriter.AddAttribute('smtClean',XmlBoolToStr(FSmtClean));
  if FSmtId <> 0 then 
    AWriter.AddAttribute('smtId',XmlIntToStr(FSmtId));
  if FBmk <> '' then 
    AWriter.AddAttribute('bmk',FBmk);
end;

procedure TCT_TextCharacterProperties.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $00000365: FKumimoji := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000001A2: FLang := AAttributes.Values[i];
      $000002C3: FAltLang := AAttributes.Values[i];
      $000000ED: FSz := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000062: FB := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000069: FI := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000075: FU := TST_TextUnderlineType(StrToEnum('sttut' + AAttributes.Values[i]));
      $00000292: FStrike := TST_TextStrikeType(StrToEnum('sttst' + AAttributes.Values[i]));
      $000001B0: FKern := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000134: FCap := TST_TextCapsType(StrToEnum('sttct' + AAttributes.Values[i]));
      $00000146: FSpc := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000419: FNormalizeH := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000343: FBaseline := XmlStrToIntDef(AAttributes.Values[i],0);
      $000002E3: FNoProof := XmlStrToBoolDef(AAttributes.Values[i],False);
      $0000022C: FDirty := XmlStrToBoolDef(AAttributes.Values[i],True);
      $00000149: FErr := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000337: FSmtClean := XmlStrToBoolDef(AAttributes.Values[i],True);
      $00000201: FSmtId := XmlStrToIntDef(AAttributes.Values[i],0);
      $0000013A: FBmk := AAttributes.Values[i];
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

procedure TCT_TextCharacterProperties.AssignFont(AItem: TCT_TextCharacterProperties);
begin
  FSz := AItem.FSz;
  FB := AItem.FB;
  FI := AItem.FI;
  FU := AItem.FU;
  FStrike := AItem.FStrike;
  FKern := AItem.FKern;
  FCap := AItem.FCap;
  FSpc := AItem.FSpc;
  FNormalizeH := AItem.FNormalizeH;
  FBaseline := AItem.FBaseline;
  if (AItem.Latin <> Nil) and (AItem.Latin.Typeface <> '') then begin
    Create_Latin;
    FA_Latin.Typeface := AItem.Latin.Typeface;
  end;
end;

constructor TCT_TextCharacterProperties.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 13;
  FAttributeCount := 19;
  FA_EG_FillProperties := TEG_FillProperties.Create(FOwner);
  FA_EG_EffectProperties := TEG_EffectProperties.Create(FOwner);
  FA_EG_TextUnderlineLine := TEG_TextUnderlineLine.Create(FOwner);
  FA_EG_TextUnderlineFill := TEG_TextUnderlineFill.Create(FOwner);
  FKumimoji := False;
  FSz := 2147483632;
  FB := False;
  FI := False;
  FU := TST_TextUnderlineType(XPG_UNKNOWN_ENUM);
  FStrike := TST_TextStrikeType(XPG_UNKNOWN_ENUM);
  FKern := 2147483632;
  FCap := TST_TextCapsType(XPG_UNKNOWN_ENUM);
  FSpc := 2147483632;
  FNormalizeH := False;
  FBaseline := 2147483632;
  FNoProof := False;
  FDirty := True;
  FErr := False;
  FSmtClean := True;
  FSmtId := 0;
end;

destructor TCT_TextCharacterProperties.Destroy;
begin
  if FA_Ln <> Nil then 
    FA_Ln.Free;
  FA_EG_FillProperties.Free;
  FA_EG_EffectProperties.Free;
  if FA_Highlight <> Nil then 
    FA_Highlight.Free;
  FA_EG_TextUnderlineLine.Free;
  FA_EG_TextUnderlineFill.Free;
  if FA_Latin <> Nil then 
    FA_Latin.Free;
  if FA_Ea <> Nil then 
    FA_Ea.Free;
  if FA_Cs <> Nil then 
    FA_Cs.Free;
  if FA_Sym <> Nil then 
    FA_Sym.Free;
  if FA_HlinkClick <> Nil then 
    FA_HlinkClick.Free;
  if FA_HlinkMouseOver <> Nil then 
    FA_HlinkMouseOver.Free;
  if FA_ExtLst <> Nil then 
    FA_ExtLst.Free;
end;

procedure TCT_TextCharacterProperties.Clear;
begin
  FAssigneds := [];
  if FA_Ln <> Nil then 
    FreeAndNil(FA_Ln);
  FA_EG_FillProperties.Clear;
  FA_EG_EffectProperties.Clear;
  if FA_Highlight <> Nil then 
    FreeAndNil(FA_Highlight);
  FA_EG_TextUnderlineLine.Clear;
  FA_EG_TextUnderlineFill.Clear;
  if FA_Latin <> Nil then 
    FreeAndNil(FA_Latin);
  if FA_Ea <> Nil then 
    FreeAndNil(FA_Ea);
  if FA_Cs <> Nil then 
    FreeAndNil(FA_Cs);
  if FA_Sym <> Nil then 
    FreeAndNil(FA_Sym);
  if FA_HlinkClick <> Nil then 
    FreeAndNil(FA_HlinkClick);
  if FA_HlinkMouseOver <> Nil then 
    FreeAndNil(FA_HlinkMouseOver);
  if FA_ExtLst <> Nil then 
    FreeAndNil(FA_ExtLst);
  Byte(FKumimoji) := 2;
  FLang := '';
  FAltLang := '';
  FSz := 2147483632;
  Byte(FB) := 2;
  Byte(FI) := 2;
  FU := TST_TextUnderlineType(XPG_UNKNOWN_ENUM);
  FStrike := TST_TextStrikeType(XPG_UNKNOWN_ENUM);
  FKern := 2147483632;
  FCap := TST_TextCapsType(XPG_UNKNOWN_ENUM);
  FSpc := 2147483632;
  Byte(FNormalizeH) := 2;
  FBaseline := 2147483632;
  Byte(FNoProof) := 2;
  FDirty := True;
  FErr := False;
  FSmtClean := True;
  FSmtId := 0;
  FBmk := '';
end;

procedure TCT_TextCharacterProperties.Assign(AItem: TCT_TextCharacterProperties);
begin
end;

procedure TCT_TextCharacterProperties.CopyTo(AItem: TCT_TextCharacterProperties);
begin
end;

function  TCT_TextCharacterProperties.Create_Ln: TCT_LineProperties;
begin
  if FA_Ln = Nil then
    FA_Ln := TCT_LineProperties.Create(FOwner);
  Result := FA_Ln;
end;

function  TCT_TextCharacterProperties.Create_Highlight: TCT_Color;
begin
  if FA_Highlight = Nil then
    FA_Highlight := TCT_Color.Create(FOwner);
  Result := FA_Highlight;
end;

function  TCT_TextCharacterProperties.Create_Latin: TCT_TextFont;
begin
  if FA_Latin = Nil then
    FA_Latin := TCT_TextFont.Create(FOwner);
  Result := FA_Latin;
end;

function  TCT_TextCharacterProperties.Create_Ea: TCT_TextFont;
begin
  if FA_Ea = Nil then
    FA_Ea := TCT_TextFont.Create(FOwner);
  Result := FA_Ea;
end;

function  TCT_TextCharacterProperties.Create_Cs: TCT_TextFont;
begin
  if FA_Cs = Nil then
    FA_Cs := TCT_TextFont.Create(FOwner);
  Result := FA_Cs;
end;

function  TCT_TextCharacterProperties.Create_Sym: TCT_TextFont;
begin
  if FA_Sym = Nil then
    FA_Sym := TCT_TextFont.Create(FOwner);
  Result := FA_Sym;
end;

function  TCT_TextCharacterProperties.Create_HlinkClick: TCT_Hyperlink;
begin
  if FA_HlinkClick = Nil then
    FA_HlinkClick := TCT_Hyperlink.Create(FOwner);
  Result := FA_HlinkClick;
end;

function  TCT_TextCharacterProperties.Create_HlinkMouseOver: TCT_Hyperlink;
begin
  if FA_HlinkMouseOver = Nil then
    FA_HlinkMouseOver := TCT_Hyperlink.Create(FOwner);
  Result := FA_HlinkMouseOver;
end;

function  TCT_TextCharacterProperties.Create_ExtLst: TCT_OfficeArtExtensionList;
begin
  if FA_ExtLst = Nil then
    FA_ExtLst := TCT_OfficeArtExtensionList.Create(FOwner);
  Result := FA_ExtLst;
end;

{ TCT_GeomGuide }

function  TCT_GeomGuide.CheckAssigned: integer;
begin
  FAssigneds := [xaAttributes];
  Result := 1;
end;

procedure TCT_GeomGuide.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_GeomGuide.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('name',FName);
  AWriter.AddAttribute('fmla',FFmla);
end;

procedure TCT_GeomGuide.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $000001A1: FName := AAttributes.Values[i];
      $000001A0: FFmla := AAttributes.Values[i];
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_GeomGuide.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 2;
  FName := '';
end;

destructor TCT_GeomGuide.Destroy;
begin
end;

procedure TCT_GeomGuide.Clear;
begin
  FAssigneds := [];
  FName := '';
  FFmla := '';
end;

procedure TCT_GeomGuide.Assign(AItem: TCT_GeomGuide);
begin
  FName := AItem.FName;
  FFmla := AItem.FFmla;
end;

procedure TCT_GeomGuide.CopyTo(AItem: TCT_GeomGuide);
begin
end;

{ TCT_GeomGuideXpgList }

function  TCT_GeomGuideXpgList.GetItems(Index: integer): TCT_GeomGuide;
begin
  Result := TCT_GeomGuide(inherited Items[Index]);
end;

function  TCT_GeomGuideXpgList.Add: TCT_GeomGuide;
begin
  Result := TCT_GeomGuide.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_GeomGuideXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_GeomGuideXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_GeomGuideXpgList.Assign(AItem: TCT_GeomGuideXpgList);
begin
end;

procedure TCT_GeomGuideXpgList.CopyTo(AItem: TCT_GeomGuideXpgList);
begin
end;

{ TCT_XYAdjustHandle }

function  TCT_XYAdjustHandle.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FGdRefX <> '' then 
    Inc(AttrsAssigned);
  if FMinX.nVal >= 0 then 
    Inc(AttrsAssigned);
  if FMaxX.nVal >= 0 then 
    Inc(AttrsAssigned);
  if FGdRefY <> '' then 
    Inc(AttrsAssigned);
  if FMinY.nVal >= 0 then 
    Inc(AttrsAssigned);
  if FMaxY.nVal >= 0 then 
    Inc(AttrsAssigned);
  if FA_Pos <> Nil then 
    Inc(ElemsAssigned,FA_Pos.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_XYAdjustHandle.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'a:pos' then 
  begin
    if FA_Pos = Nil then 
      FA_Pos := TCT_AdjPoint2D.Create(FOwner);
    Result := FA_Pos;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_XYAdjustHandle.Write(AWriter: TXpgWriteXML);
begin
  if (FA_Pos <> Nil) and FA_Pos.Assigned then 
  begin
    FA_Pos.WriteAttributes(AWriter);
    AWriter.SimpleTag('a:pos');
  end
  else 
    AWriter.SimpleTag('a:pos');
end;

procedure TCT_XYAdjustHandle.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FGdRefX <> '' then 
    AWriter.AddAttribute('gdRefX',FGdRefX);
  if FMinX.nVal >= 0 then 
    AWriter.AddAttribute('minX',WriteUnionTST_AdjCoordinate(PFMinX));
  if FMaxX.nVal >= 0 then 
    AWriter.AddAttribute('maxX',WriteUnionTST_AdjCoordinate(PFMaxX));
  if FGdRefY <> '' then 
    AWriter.AddAttribute('gdRefY',FGdRefY);
  if FMinY.nVal >= 0 then 
    AWriter.AddAttribute('minY',WriteUnionTST_AdjCoordinate(PFMinY));
  if FMaxY.nVal >= 0 then 
    AWriter.AddAttribute('maxY',WriteUnionTST_AdjCoordinate(PFMaxY));
end;

procedure TCT_XYAdjustHandle.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $00000240: FGdRefX := AAttributes.Values[i];
      $0000019C: ReadUnionTST_AdjCoordinate(AAttributes.Values[i],PFMinX);
      $0000019E: ReadUnionTST_AdjCoordinate(AAttributes.Values[i],PFMaxX);
      $00000241: FGdRefY := AAttributes.Values[i];
      $0000019D: ReadUnionTST_AdjCoordinate(AAttributes.Values[i],PFMinY);
      $0000019F: ReadUnionTST_AdjCoordinate(AAttributes.Values[i],PFMaxY);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_XYAdjustHandle.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 6;
  FGdRefX := '';
  PFMinX := @FMinX;
  PMinX.nVal := -1;
  PFMaxX := @FMaxX;
  PMaxX.nVal := -1;
  FGdRefY := '';
  PFMinY := @FMinY;
  PMinY.nVal := -1;
  PFMaxY := @FMaxY;
  PMaxY.nVal := -1;
end;

destructor TCT_XYAdjustHandle.Destroy;
begin
  if FA_Pos <> Nil then
    FA_Pos.Free;
end;

procedure TCT_XYAdjustHandle.Clear;
begin
  FAssigneds := [];
  if FA_Pos <> Nil then
    FreeAndNil(FA_Pos);
  FGdRefX := '';
  FGdRefY := '';
end;

procedure TCT_XYAdjustHandle.Assign(AItem: TCT_XYAdjustHandle);
begin
  FGdRefX := AItem.FGdRefX;
  FMinX := AItem.FMinX;
  FMaxX := AItem.FMaxX;
  FGdRefY := AItem.FGdRefY;
  FMinY := AItem.FMinY;
  FMaxY := AItem.FMaxY;
  FA_Pos := AItem.FA_Pos;
end;

procedure TCT_XYAdjustHandle.CopyTo(AItem: TCT_XYAdjustHandle);
begin
end;

function  TCT_XYAdjustHandle.Create_A_Pos: TCT_AdjPoint2D;
begin
  if FA_Pos = Nil then
    FA_Pos := TCT_AdjPoint2D.Create(FOwner);
  Result := FA_Pos;
end;

{ TCT_XYAdjustHandleXpgList }

function  TCT_XYAdjustHandleXpgList.GetItems(Index: integer): TCT_XYAdjustHandle;
begin
  Result := TCT_XYAdjustHandle(inherited Items[Index]);
end;

function  TCT_XYAdjustHandleXpgList.Add: TCT_XYAdjustHandle;
begin
  Result := TCT_XYAdjustHandle.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_XYAdjustHandleXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_XYAdjustHandleXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_XYAdjustHandleXpgList.Assign(AItem: TCT_XYAdjustHandleXpgList);
var
  i: integer;
begin
  for i := 0 to AItem.Count - 1 do
    Add.Assign(AItem[i]);
end;

procedure TCT_XYAdjustHandleXpgList.CopyTo(AItem: TCT_XYAdjustHandleXpgList);
begin
end;

{ TCT_PolarAdjustHandle }

function  TCT_PolarAdjustHandle.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FGdRefR <> '' then 
    Inc(AttrsAssigned);
  if FMinR.nVal >= 0 then 
    Inc(AttrsAssigned);
  if FMaxR.nVal >= 0 then 
    Inc(AttrsAssigned);
  if FGdRefAng <> '' then 
    Inc(AttrsAssigned);
  if FMinAng.nVal >= 0 then 
    Inc(AttrsAssigned);
  if FMaxAng.nVal >= 0 then 
    Inc(AttrsAssigned);
  if FA_Pos <> Nil then 
    Inc(ElemsAssigned,FA_Pos.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_PolarAdjustHandle.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'a:pos' then 
  begin
    if FA_Pos = Nil then 
      FA_Pos := TCT_AdjPoint2D.Create(FOwner);
    Result := FA_Pos;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_PolarAdjustHandle.Write(AWriter: TXpgWriteXML);
begin
  if (FA_Pos <> Nil) and FA_Pos.Assigned then 
  begin
    FA_Pos.WriteAttributes(AWriter);
    AWriter.SimpleTag('a:pos');
  end
  else 
    AWriter.SimpleTag('a:pos');
end;

procedure TCT_PolarAdjustHandle.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FGdRefR <> '' then 
    AWriter.AddAttribute('gdRefR',FGdRefR);
  if FMinR.nVal >= 0 then 
    AWriter.AddAttribute('minR',WriteUnionTST_AdjCoordinate(PFMinR));
  if FMaxR.nVal >= 0 then 
    AWriter.AddAttribute('maxR',WriteUnionTST_AdjCoordinate(PFMaxR));
  if FGdRefAng <> '' then 
    AWriter.AddAttribute('gdRefAng',FGdRefAng);
  if FMinAng.nVal >= 0 then 
    AWriter.AddAttribute('minAng',WriteUnionTST_AdjAngle(PFMinAng));
  if FMaxAng.nVal >= 0 then 
    AWriter.AddAttribute('maxAng',WriteUnionTST_AdjAngle(PFMaxAng));
end;

procedure TCT_PolarAdjustHandle.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $0000023A: FGdRefR := AAttributes.Values[i];
      $00000196: ReadUnionTST_AdjCoordinate(AAttributes.Values[i],PFMinR);
      $00000198: ReadUnionTST_AdjCoordinate(AAttributes.Values[i],PFMaxR);
      $000002FE: FGdRefAng := AAttributes.Values[i];
      $0000025A: ReadUnionTST_AdjAngle(AAttributes.Values[i],PFMinAng);
      $0000025C: ReadUnionTST_AdjAngle(AAttributes.Values[i],PFMaxAng);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_PolarAdjustHandle.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 6;
  FGdRefR := '';
  PFMinR := @FMinR;
  PMinR.nVal := -1;
  PFMaxR := @FMaxR;
  PMaxR.nVal := -1;
  FGdRefAng := '';
  PFMinAng := @FMinAng;
  PMinAng.nVal := -1;
  PFMaxAng := @FMaxAng;
  PMaxAng.nVal := -1;
end;

destructor TCT_PolarAdjustHandle.Destroy;
begin
  if FA_Pos <> Nil then 
    FA_Pos.Free;
end;

procedure TCT_PolarAdjustHandle.Clear;
begin
  FAssigneds := [];
  if FA_Pos <> Nil then 
    FreeAndNil(FA_Pos);
  FGdRefR := '';
  FGdRefAng := '';
end;

procedure TCT_PolarAdjustHandle.Assign(AItem: TCT_PolarAdjustHandle);
begin
  FGdRefR := AItem.FGdRefR;
  FMinR := AItem.FMinR;
  FMaxR := AItem.FMaxR;
  FGdRefAng := AItem.FGdRefAng;
  FMinAng := AItem.FMinAng;
  FMaxAng := AItem.FMaxAng;
  FA_Pos := AItem.FA_Pos;
end;

procedure TCT_PolarAdjustHandle.CopyTo(AItem: TCT_PolarAdjustHandle);
begin
end;

function  TCT_PolarAdjustHandle.Create_A_Pos: TCT_AdjPoint2D;
begin
  if FA_Pos = Nil then
    FA_Pos := TCT_AdjPoint2D.Create(FOwner);
  Result := FA_Pos;
end;

{ TCT_PolarAdjustHandleXpgList }

function  TCT_PolarAdjustHandleXpgList.GetItems(Index: integer): TCT_PolarAdjustHandle;
begin
  Result := TCT_PolarAdjustHandle(inherited Items[Index]);
end;

function  TCT_PolarAdjustHandleXpgList.Add: TCT_PolarAdjustHandle;
begin
  Result := TCT_PolarAdjustHandle.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_PolarAdjustHandleXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_PolarAdjustHandleXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_PolarAdjustHandleXpgList.Assign(AItem: TCT_PolarAdjustHandleXpgList);
var
  i: integer;
begin
  for i := 0 to AItem.Count - 1 do
    Add.Assign(AItem[i]);
end;

procedure TCT_PolarAdjustHandleXpgList.CopyTo(AItem: TCT_PolarAdjustHandleXpgList);
begin
end;

{ TCT_ConnectionSite }

function  TCT_ConnectionSite.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [xaAttributes];
  if FA_Pos <> Nil then
    Inc(ElemsAssigned,FA_Pos.CheckAssigned);
  Result := 1;
  if ElemsAssigned > 0 then
    FAssigneds := FAssigneds + [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_ConnectionSite.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'a:pos' then 
  begin
    if FA_Pos = Nil then 
      FA_Pos := TCT_AdjPoint2D.Create(FOwner);
    Result := FA_Pos;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_ConnectionSite.Write(AWriter: TXpgWriteXML);
begin
  if (FA_Pos <> Nil) and FA_Pos.Assigned then 
  begin
    FA_Pos.WriteAttributes(AWriter);
    AWriter.SimpleTag('a:pos');
  end
  else 
    AWriter.SimpleTag('a:pos');
end;

procedure TCT_ConnectionSite.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('ang',WriteUnionTST_AdjAngle(PFAng));
end;

procedure TCT_ConnectionSite.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'ang' then 
    ReadUnionTST_AdjAngle(AAttributes.Values[0],PFAng)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_ConnectionSite.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 1;
  PFAng := @FAng;
  PAng.nVal := -1;
end;

destructor TCT_ConnectionSite.Destroy;
begin
  if FA_Pos <> Nil then 
    FA_Pos.Free;
end;

procedure TCT_ConnectionSite.Clear;
begin
  FAssigneds := [];
  if FA_Pos <> Nil then 
    FreeAndNil(FA_Pos);
end;

procedure TCT_ConnectionSite.Assign(AItem: TCT_ConnectionSite);
begin
  FAng := AItem.FAng;
  FA_Pos := AItem.FA_Pos;
end;

procedure TCT_ConnectionSite.CopyTo(AItem: TCT_ConnectionSite);
begin
end;

function  TCT_ConnectionSite.Create_A_Pos: TCT_AdjPoint2D;
begin
  if FA_Pos = Nil then
    FA_Pos := TCT_AdjPoint2D.Create(FOwner);
  Result := FA_Pos;
end;

{ TCT_ConnectionSiteXpgList }

function  TCT_ConnectionSiteXpgList.GetItems(Index: integer): TCT_ConnectionSite;
begin
  Result := TCT_ConnectionSite(inherited Items[Index]);
end;

function  TCT_ConnectionSiteXpgList.Add: TCT_ConnectionSite;
begin
  Result := TCT_ConnectionSite.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_ConnectionSiteXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_ConnectionSiteXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_ConnectionSiteXpgList.Assign(AItem: TCT_ConnectionSiteXpgList);
var
  i: integer;
begin
  for i := 0 to AItem.Count - 1 do
    Add.Assign(AItem[i]);
end;

procedure TCT_ConnectionSiteXpgList.CopyTo(AItem: TCT_ConnectionSiteXpgList);
begin
end;

{ TCT_Path2D }

function  TCT_Path2D.CheckAssigned: integer;
var
  i: integer;
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FW <> 0 then 
    Inc(AttrsAssigned);
  if FH <> 0 then 
    Inc(AttrsAssigned);
  if FFill <> stpfmNorm then 
    Inc(AttrsAssigned);
  if FStroke <> True then 
    Inc(AttrsAssigned);
  if FExtrusionOk <> True then 
    Inc(AttrsAssigned);

  for i := 0 to FObjects.Count - 1 do
    Inc(ElemsAssigned,TXPGBase(FObjects[i]).CheckAssigned);

  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_Path2D.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;

  case AReader.QNameHashA of
    $000002B1: begin
      Result := TCT_Path2DClose.Create(FOwner);
      FObjects.Add(Result);
    end;
    $00000315: begin
      Result := TCT_Path2DMoveTo.Create(FOwner);
      FObjects.Add(Result);
    end;
    $00000238: begin
      Result := TCT_Path2DLineTo.Create(FOwner);
      FObjects.Add(Result);
    end;
    $00000294: begin
      Result := TCT_Path2DArcTo.Create(FOwner);
      FObjects.Add(Result);
    end;
    $0000042A: begin
      Result := TCT_Path2DQuadBezierTo.Create(FOwner);
      FObjects.Add(Result);
    end;
    $00000485: begin
      Result := TCT_Path2DCubicBezierTo.Create(FOwner);
      FObjects.Add(Result);
    end;
    else
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then
    Result.Assigneds := [xaRead];
end;

procedure TCT_Path2D.Write(AWriter: TXpgWriteXML);
var
  i: integer;
  O: TXPGBase;
  TagName: AxUCString;
begin
  for i := 0 to FObjects.Count - 1 do begin
    O := TXPGBase(FObjects[i]);

    TagName := '';
    if O is TCT_Path2DClose then begin
      AWriter.BeginTag('a:close');
      TCT_Path2DClose(O).Write(AWriter);
      AWriter.EndTag;
    end
    else if O is TCT_Path2DMoveTo then begin
      AWriter.BeginTag('a:moveTo');
      TCT_Path2DMoveTo(O).Write(AWriter);
      AWriter.EndTag;
    end
    else if O is TCT_Path2DLineTo then begin
      AWriter.BeginTag('a:lnTo');
      TCT_Path2DLineTo(O).Write(AWriter);
      AWriter.EndTag;
    end
    else if O is TCT_Path2DArcTo then begin
      AWriter.BeginTag('a:arcTo');
      TCT_Path2DArcTo(O).Write(AWriter);
      AWriter.EndTag;
    end
    else if O is TCT_Path2DQuadBezierTo then begin
      AWriter.BeginTag('a:quadBezTo');
      TCT_Path2DQuadBezierTo(O).Write(AWriter);
      AWriter.EndTag;
    end
    else if O is TCT_Path2DCubicBezierTo then begin
      AWriter.BeginTag('a:cubicBezTo');
      TCT_Path2DCubicBezierTo(O).Write(AWriter);
      AWriter.EndTag;
    end;
  end;
end;

procedure TCT_Path2D.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FW <> 0 then
    AWriter.AddAttribute('w',XmlIntToStr(FW));
  if FH <> 0 then 
    AWriter.AddAttribute('h',XmlIntToStr(FH));
  if FFill <> stpfmNorm then 
    AWriter.AddAttribute('fill',StrTST_PathFillMode[Integer(FFill)]);
  if FStroke <> True then 
    AWriter.AddAttribute('stroke',XmlBoolToStr(FStroke));
  if FExtrusionOk <> True then 
    AWriter.AddAttribute('extrusionOk',XmlBoolToStr(FExtrusionOk));
end;

procedure TCT_Path2D.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $00000077: FW := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000068: FH := XmlStrToIntDef(AAttributes.Values[i],0);
      $000001A7: FFill := TST_PathFillMode(StrToEnum('stpfm' + AAttributes.Values[i]));
      $00000298: FStroke := XmlStrToBoolDef(AAttributes.Values[i],True);
      $000004AB: FExtrusionOk := XmlStrToBoolDef(AAttributes.Values[i],True);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_Path2D.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 6;
  FAttributeCount := 5;
  FW := 0;
  FH := 0;
  FFill := stpfmNorm;
  FStroke := True;
  FExtrusionOk := True;

  FObjects := TXPGBaseObjectList.Create(FOwner);
end;

destructor TCT_Path2D.Destroy;
begin
  FObjects.Free;
end;

procedure TCT_Path2D.Clear;
begin
  FAssigneds := [];
  FObjects.Clear;

  FW := 0;
  FH := 0;
  FFill := stpfmNorm;
  FStroke := True;
  FExtrusionOk := True;
end;

procedure TCT_Path2D.Assign(AItem: TCT_Path2D);
begin
  FW := AItem.FW;
  FH := AItem.FH;
  FFill := AItem.FFill;
  FStroke := AItem.FStroke;
  FExtrusionOk := AItem.FExtrusionOk;

  raise Exception.Create('__TODO__');
//  FObjects: TXPGBaseObjectList;
end;

procedure TCT_Path2D.CopyTo(AItem: TCT_Path2D);
begin
end;

{ TCT_Path2DXpgList }

function  TCT_Path2DXpgList.GetItems(Index: integer): TCT_Path2D;
begin
  Result := TCT_Path2D(inherited Items[Index]);
end;

function  TCT_Path2DXpgList.Add: TCT_Path2D;
begin
  Result := TCT_Path2D.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_Path2DXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_Path2DXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_Path2DXpgList.Assign(AItem: TCT_Path2DXpgList);
var
  i: integer;
begin
  for i := 0 to AItem.Count - 1 do
    Add.Assign(AItem[i]);
end;

procedure TCT_Path2DXpgList.CopyTo(AItem: TCT_Path2DXpgList);
begin
end;

{ TCT_SphereCoords }

function  TCT_SphereCoords.CheckAssigned: integer;
begin
  FAssigneds := [xaAttributes];
  Result := 1;
end;

procedure TCT_SphereCoords.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_SphereCoords.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('lat',XmlIntToStr(FLat));
  AWriter.AddAttribute('lon',XmlIntToStr(FLon));
  AWriter.AddAttribute('rev',XmlIntToStr(FRev));
end;

procedure TCT_SphereCoords.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $00000141: FLat := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000149: FLon := XmlStrToIntDef(AAttributes.Values[i],0);
      $0000014D: FRev := XmlStrToIntDef(AAttributes.Values[i],0);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_SphereCoords.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 3;
  FLat := 2147483632;
  FLon := 2147483632;
  FRev := 2147483632;
end;

destructor TCT_SphereCoords.Destroy;
begin
end;

procedure TCT_SphereCoords.Clear;
begin
  FAssigneds := [];
  FLat := 2147483632;
  FLon := 2147483632;
  FRev := 2147483632;
end;

procedure TCT_SphereCoords.Assign(AItem: TCT_SphereCoords);
begin
  FLat := AItem.FLat;
  FLon := AItem.FLon;
  FRev := AItem.FRev;
end;

procedure TCT_SphereCoords.CopyTo(AItem: TCT_SphereCoords);
begin
end;

{ TCT_Point3D }

function  TCT_Point3D.CheckAssigned: integer;
begin
  FAssigneds := [xaAttributes];
  Result := 1;
end;

procedure TCT_Point3D.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_Point3D.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('x',XmlIntToStr(FX));
  AWriter.AddAttribute('y',XmlIntToStr(FY));
  AWriter.AddAttribute('z',XmlIntToStr(FZ));
end;

procedure TCT_Point3D.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $00000078: FX := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000079: FY := XmlStrToIntDef(AAttributes.Values[i],0);
      $0000007A: FZ := XmlStrToIntDef(AAttributes.Values[i],0);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_Point3D.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 3;
  FX := 2147483632;
  FY := 2147483632;
  FZ := 2147483632;
end;

destructor TCT_Point3D.Destroy;
begin
end;

procedure TCT_Point3D.Clear;
begin
  FAssigneds := [];
  FX := 2147483632;
  FY := 2147483632;
  FZ := 2147483632;
end;

procedure TCT_Point3D.Assign(AItem: TCT_Point3D);
begin
  FX := AItem.FX;
  FY := AItem.FY;
  FZ := AItem.FZ;
end;

procedure TCT_Point3D.CopyTo(AItem: TCT_Point3D);
begin
end;

{ TCT_Vector3D }

function  TCT_Vector3D.CheckAssigned: integer;
begin
  FAssigneds := [xaAttributes];
  Result := 1;
end;

procedure TCT_Vector3D.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_Vector3D.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('dx',XmlIntToStr(FDx));
  AWriter.AddAttribute('dy',XmlIntToStr(FDy));
  AWriter.AddAttribute('dz',XmlIntToStr(FDz));
end;

procedure TCT_Vector3D.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $000000DC: FDx := XmlStrToIntDef(AAttributes.Values[i],0);
      $000000DD: FDy := XmlStrToIntDef(AAttributes.Values[i],0);
      $000000DE: FDz := XmlStrToIntDef(AAttributes.Values[i],0);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_Vector3D.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 3;
  FDx := 2147483632;
  FDy := 2147483632;
  FDz := 2147483632;
end;

destructor TCT_Vector3D.Destroy;
begin
end;

procedure TCT_Vector3D.Clear;
begin
  FAssigneds := [];
  FDx := 2147483632;
  FDy := 2147483632;
  FDz := 2147483632;
end;

procedure TCT_Vector3D.Assign(AItem: TCT_Vector3D);
begin
  FDx := AItem.FDx;
  FDy := AItem.FDy;
  FDz := AItem.FDz;
end;

procedure TCT_Vector3D.CopyTo(AItem: TCT_Vector3D);
begin
end;

{ TCT_Bevel }

function  TCT_Bevel.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FW <> 76200 then
    Inc(AttrsAssigned);
  if FH <> 76200 then
    Inc(AttrsAssigned);
  if FPrst <> stbptCircle then
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then
    FAssigneds := FAssigneds + [xaAttributes];
end;

procedure TCT_Bevel.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_Bevel.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FW <> 76200 then 
    AWriter.AddAttribute('w',XmlIntToStr(FW));
  if FH <> 76200 then 
    AWriter.AddAttribute('h',XmlIntToStr(FH));
  if FPrst <> stbptCircle then 
    AWriter.AddAttribute('prst',StrTST_BevelPresetType[Integer(FPrst)]);
end;

procedure TCT_Bevel.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $00000077: FW := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000068: FH := XmlStrToIntDef(AAttributes.Values[i],0);
      $000001C9: FPrst := TST_BevelPresetType(StrToEnum('stbpt' + AAttributes.Values[i]));
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_Bevel.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 3;
  FW := 76200;
  FH := 76200;
  FPrst := stbptCircle;
end;

destructor TCT_Bevel.Destroy;
begin
end;

procedure TCT_Bevel.Clear;
begin
  FAssigneds := [];
  FW := 76200;
  FH := 76200;
  FPrst := stbptCircle;
end;

procedure TCT_Bevel.Assign(AItem: TCT_Bevel);
begin
  FW := AItem.FW;
  FH := AItem.FH;
  FPrst := AItem.FPrst;
end;

procedure TCT_Bevel.CopyTo(AItem: TCT_Bevel);
begin
end;

{ TCT_TextParagraphProperties }

function  TCT_TextParagraphProperties.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FMarL <> 2147483632 then 
    Inc(AttrsAssigned);
  if FMarR <> 2147483632 then 
    Inc(AttrsAssigned);
  if FLvl <> 2147483632 then 
    Inc(AttrsAssigned);
  if FIndent <> 2147483632 then 
    Inc(AttrsAssigned);
  if Integer(FAlgn) <> XPG_UNKNOWN_ENUM then
    Inc(AttrsAssigned);
  if FDefTabSz <> 2147483632 then 
    Inc(AttrsAssigned);
  if Byte(FRtl) <> 2 then 
    Inc(AttrsAssigned);
  if Byte(FEaLnBrk) <> 2 then 
    Inc(AttrsAssigned);
  if Integer(FFontAlgn) <> XPG_UNKNOWN_ENUM then 
    Inc(AttrsAssigned);
  if Byte(FLatinLnBrk) <> 2 then 
    Inc(AttrsAssigned);
  if Byte(FHangingPunct) <> 2 then 
    Inc(AttrsAssigned);
  if FA_LnSpc <> Nil then 
    Inc(ElemsAssigned,FA_LnSpc.CheckAssigned);
  if FA_SpcBef <> Nil then 
    Inc(ElemsAssigned,FA_SpcBef.CheckAssigned);
  if FA_SpcAft <> Nil then 
    Inc(ElemsAssigned,FA_SpcAft.CheckAssigned);
  Inc(ElemsAssigned,FA_EG_TextBulletColor.CheckAssigned);
  Inc(ElemsAssigned,FA_EG_TextBulletSize.CheckAssigned);
  Inc(ElemsAssigned,FA_EG_TextBulletTypeface.CheckAssigned);
  Inc(ElemsAssigned,FA_EG_TextBullet.CheckAssigned);
  if FA_TabLst <> Nil then 
    Inc(ElemsAssigned,FA_TabLst.CheckAssigned);
  if FA_DefRPr <> Nil then 
    Inc(ElemsAssigned,FA_DefRPr.CheckAssigned);
  if FA_ExtLst <> Nil then 
    Inc(ElemsAssigned,FA_ExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_TextParagraphProperties.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  case AReader.QNameHashA of
    $0000029B: begin
      if FA_LnSpc = Nil then 
        FA_LnSpc := TCT_TextSpacing.Create(FOwner);
      Result := FA_LnSpc;
    end;
    $000002EE: begin
      if FA_SpcBef = Nil then 
        FA_SpcBef := TCT_TextSpacing.Create(FOwner);
      Result := FA_SpcBef;
    end;
    $000002FC: begin
      if FA_SpcAft = Nil then 
        FA_SpcAft := TCT_TextSpacing.Create(FOwner);
      Result := FA_SpcAft;
    end;
    $00000305: begin
      if FA_TabLst = Nil then 
        FA_TabLst := TCT_TextTabStopList.Create(FOwner);
      Result := FA_TabLst;
    end;
    $000002DE: begin
      if FA_DefRPr = Nil then 
        FA_DefRPr := TCT_TextCharacterProperties.Create(FOwner);
      Result := FA_DefRPr;
    end;
    $0000031F: begin
      if FA_ExtLst = Nil then 
        FA_ExtLst := TCT_OfficeArtExtensionList.Create(FOwner);
      Result := FA_ExtLst;
    end;
    else 
    begin
      Result := FA_EG_TextBulletColor.HandleElement(AReader);
      if Result = Nil then 
      begin
        Result := FA_EG_TextBulletSize.HandleElement(AReader);
        if Result = Nil then 
        begin
          Result := FA_EG_TextBulletTypeface.HandleElement(AReader);
          if Result = Nil then 
          begin
            Result := FA_EG_TextBullet.HandleElement(AReader);
            if Result = Nil then 
              FOwner.Errors.Error(xemUnknownElement,AReader.QName);
          end;
        end;
      end;
    end
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_TextParagraphProperties.Write(AWriter: TXpgWriteXML);
begin
  if (FA_LnSpc <> Nil) and FA_LnSpc.Assigned then 
    if xaElements in FA_LnSpc.FAssigneds then 
    begin
      AWriter.BeginTag('a:lnSpc');
      FA_LnSpc.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:lnSpc');
  if (FA_SpcBef <> Nil) and FA_SpcBef.Assigned then 
    if xaElements in FA_SpcBef.FAssigneds then 
    begin
      AWriter.BeginTag('a:spcBef');
      FA_SpcBef.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:spcBef');
  if (FA_SpcAft <> Nil) and FA_SpcAft.Assigned then 
    if xaElements in FA_SpcAft.FAssigneds then 
    begin
      AWriter.BeginTag('a:spcAft');
      FA_SpcAft.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:spcAft');
  FA_EG_TextBulletColor.Write(AWriter);
  FA_EG_TextBulletSize.Write(AWriter);
  FA_EG_TextBulletTypeface.Write(AWriter);
  FA_EG_TextBullet.Write(AWriter);
  if (FA_TabLst <> Nil) and FA_TabLst.Assigned then 
    if xaElements in FA_TabLst.FAssigneds then 
    begin
      AWriter.BeginTag('a:tabLst');
      FA_TabLst.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:tabLst');
  if (FA_DefRPr <> Nil) and FA_DefRPr.Assigned then 
  begin
    FA_DefRPr.WriteAttributes(AWriter);
    if xaElements in FA_DefRPr.FAssigneds then 
    begin
      AWriter.BeginTag('a:defRPr');
      FA_DefRPr.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:defRPr');
  end;
  if (FA_ExtLst <> Nil) and FA_ExtLst.Assigned then 
    if xaElements in FA_ExtLst.FAssigneds then 
    begin
      AWriter.BeginTag('a:extLst');
      FA_ExtLst.Write(AWriter);
      AWriter.EndTag;
    end;
end;

procedure TCT_TextParagraphProperties.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FMarL <> 2147483632 then 
    AWriter.AddAttribute('marL',XmlIntToStr(FMarL));
  if FMarR <> 2147483632 then 
    AWriter.AddAttribute('marR',XmlIntToStr(FMarR));
  if FLvl <> 2147483632 then 
    AWriter.AddAttribute('lvl',XmlIntToStr(FLvl));
  if FIndent <> 2147483632 then 
    AWriter.AddAttribute('indent',XmlIntToStr(FIndent));
  if Integer(FAlgn) <> XPG_UNKNOWN_ENUM then 
    AWriter.AddAttribute('algn',StrTST_TextAlignType[Integer(FAlgn)]);
  if FDefTabSz <> 2147483632 then 
    AWriter.AddAttribute('defTabSz',XmlIntToStr(FDefTabSz));
  if Byte(FRtl) <> 2 then 
    AWriter.AddAttribute('rtl',XmlBoolToStr(FRtl));
  if Byte(FEaLnBrk) <> 2 then 
    AWriter.AddAttribute('eaLnBrk',XmlBoolToStr(FEaLnBrk));
  if Integer(FFontAlgn) <> XPG_UNKNOWN_ENUM then 
    AWriter.AddAttribute('fontAlgn',StrTST_TextFontAlignType[Integer(FFontAlgn)]);
  if Byte(FLatinLnBrk) <> 2 then 
    AWriter.AddAttribute('latinLnBrk',XmlBoolToStr(FLatinLnBrk));
  if Byte(FHangingPunct) <> 2 then 
    AWriter.AddAttribute('hangingPunct',XmlBoolToStr(FHangingPunct));
end;

procedure TCT_TextParagraphProperties.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $0000018C: FMarL := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000192: FMarR := XmlStrToIntDef(AAttributes.Values[i],0);
      $0000014E: FLvl := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000282: FIndent := XmlStrToIntDef(AAttributes.Values[i],0);
      $000001A2: FAlgn := TST_TextAlignType(StrToEnum('sttat' + AAttributes.Values[i]));
      $00000313: FDefTabSz := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000152: FRtl := XmlStrToBoolDef(AAttributes.Values[i],False);
      $0000029F: FEaLnBrk := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000339: FFontAlgn := TST_TextFontAlignType(StrToEnum('sttfat' + AAttributes.Values[i]));
      $000003F1: FLatinLnBrk := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000004E6: FHangingPunct := XmlStrToBoolDef(AAttributes.Values[i],False);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_TextParagraphProperties.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 10;
  FAttributeCount := 11;
  FA_EG_TextBulletColor := TEG_TextBulletColor.Create(FOwner);
  FA_EG_TextBulletSize := TEG_TextBulletSize.Create(FOwner);
  FA_EG_TextBulletTypeface := TEG_TextBulletTypeface.Create(FOwner);
  FA_EG_TextBullet := TEG_TextBullet.Create(FOwner);
  FMarL := 2147483632;
  FMarR := 2147483632;
  FLvl := 2147483632;
  FIndent := 2147483632;
  FAlgn := TST_TextAlignType(XPG_UNKNOWN_ENUM);
  FDefTabSz := 2147483632;
  Byte(FRtl) := 2;
  Byte(FEaLnBrk) := 2;
  FFontAlgn := TST_TextFontAlignType(XPG_UNKNOWN_ENUM);
  Byte(FLatinLnBrk) := 2;
  Byte(FHangingPunct) := 2;
end;

destructor TCT_TextParagraphProperties.Destroy;
begin
  if FA_LnSpc <> Nil then 
    FA_LnSpc.Free;
  if FA_SpcBef <> Nil then 
    FA_SpcBef.Free;
  if FA_SpcAft <> Nil then 
    FA_SpcAft.Free;
  FA_EG_TextBulletColor.Free;
  FA_EG_TextBulletSize.Free;
  FA_EG_TextBulletTypeface.Free;
  FA_EG_TextBullet.Free;
  if FA_TabLst <> Nil then 
    FA_TabLst.Free;
  if FA_DefRPr <> Nil then 
    FA_DefRPr.Free;
  if FA_ExtLst <> Nil then 
    FA_ExtLst.Free;
end;

procedure TCT_TextParagraphProperties.Clear;
begin
  FAssigneds := [];
  if FA_LnSpc <> Nil then 
    FreeAndNil(FA_LnSpc);
  if FA_SpcBef <> Nil then 
    FreeAndNil(FA_SpcBef);
  if FA_SpcAft <> Nil then 
    FreeAndNil(FA_SpcAft);
  FA_EG_TextBulletColor.Clear;
  FA_EG_TextBulletSize.Clear;
  FA_EG_TextBulletTypeface.Clear;
  FA_EG_TextBullet.Clear;
  if FA_TabLst <> Nil then 
    FreeAndNil(FA_TabLst);
  if FA_DefRPr <> Nil then 
    FreeAndNil(FA_DefRPr);
  if FA_ExtLst <> Nil then 
    FreeAndNil(FA_ExtLst);
  FMarL := 2147483632;
  FMarR := 2147483632;
  FLvl := 2147483632;
  FIndent := 2147483632;
  FAlgn := TST_TextAlignType(XPG_UNKNOWN_ENUM);
  FFontAlgn := TST_TextFontAlignType(XPG_UNKNOWN_ENUM);
  FDefTabSz := 2147483632;
  Byte(FRtl) := 2;
  Byte(FEaLnBrk) := 2;
  Byte(FLatinLnBrk) := 2;
  Byte(FHangingPunct) := 2;
end;

procedure TCT_TextParagraphProperties.Assign(AItem: TCT_TextParagraphProperties);
begin
end;

procedure TCT_TextParagraphProperties.CopyTo(AItem: TCT_TextParagraphProperties);
begin
end;

function  TCT_TextParagraphProperties.Create_LnSpc: TCT_TextSpacing;
begin
  if FA_LnSpc = Nil then
    FA_LnSpc := TCT_TextSpacing.Create(FOwner);
  Result := FA_LnSpc;
end;

function  TCT_TextParagraphProperties.Create_SpcBef: TCT_TextSpacing;
begin
  if FA_SpcBef = Nil then
    FA_SpcBef := TCT_TextSpacing.Create(FOwner);
  Result := FA_SpcBef;
end;

function  TCT_TextParagraphProperties.Create_SpcAft: TCT_TextSpacing;
begin
  if FA_SpcAft = Nil then
    FA_SpcAft := TCT_TextSpacing.Create(FOwner);
  Result := FA_SpcAft;
end;

function  TCT_TextParagraphProperties.Create_TabLst: TCT_TextTabStopList;
begin
  if FA_TabLst = Nil then
    FA_TabLst := TCT_TextTabStopList.Create(FOwner);
  Result := FA_TabLst;
end;

function  TCT_TextParagraphProperties.Create_DefRPr: TCT_TextCharacterProperties;
begin
  if FA_DefRPr = Nil then
    FA_DefRPr := TCT_TextCharacterProperties.Create(FOwner);
  Result := FA_DefRPr;
end;

function  TCT_TextParagraphProperties.Create_ExtLst: TCT_OfficeArtExtensionList;
begin
  if FA_ExtLst = Nil then
    FA_ExtLst := TCT_OfficeArtExtensionList.Create(FOwner);
  Result := FA_ExtLst;
end;

{ TCT_GeomGuideList }

function  TCT_GeomGuideList.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 1;
  FAssigneds := [];
  if FA_GdXpgList <> Nil then
    Inc(ElemsAssigned,FA_GdXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_GeomGuideList.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'a:gd' then 
  begin
    if FA_GdXpgList = Nil then 
      FA_GdXpgList := TCT_GeomGuideXpgList.Create(FOwner);
    Result := FA_GdXpgList.Add;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_GeomGuideList.Write(AWriter: TXpgWriteXML);
begin
  if FA_GdXpgList <> Nil then 
    FA_GdXpgList.Write(AWriter,'a:gd');
end;

constructor TCT_GeomGuideList.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 0;
end;

destructor TCT_GeomGuideList.Destroy;
begin
  if FA_GdXpgList <> Nil then 
    FA_GdXpgList.Free;
end;

procedure TCT_GeomGuideList.Clear;
begin
  FAssigneds := [];
  if FA_GdXpgList <> Nil then 
    FreeAndNil(FA_GdXpgList);
end;

procedure TCT_GeomGuideList.Assign(AItem: TCT_GeomGuideList);
var
  i: integer;
begin
 if AItem.FA_GdXpgList <> Nil then begin
   for i := 0 to AItem.FA_GdXpgList.Count - 1 do
     FA_GdXpgList.Add.Assign(AItem.FA_GdXpgList[i]);
 end;
end;

procedure TCT_GeomGuideList.CopyTo(AItem: TCT_GeomGuideList);
begin
end;

function  TCT_GeomGuideList.Create_A_GdXpgList: TCT_GeomGuideXpgList;
begin
  if FA_GdXpgList = Nil then
    FA_GdXpgList := TCT_GeomGuideXpgList.Create(FOwner);
  Result := FA_GdXpgList;
end;

{ TCT_AdjustHandleList }

function  TCT_AdjustHandleList.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FA_AhXYXpgList <> Nil then 
    Inc(ElemsAssigned,FA_AhXYXpgList.CheckAssigned);
  if FA_AhPolarXpgList <> Nil then 
    Inc(ElemsAssigned,FA_AhPolarXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_AdjustHandleList.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $00000215: begin
      if FA_AhXYXpgList = Nil then 
        FA_AhXYXpgList := TCT_XYAdjustHandleXpgList.Create(FOwner);
      Result := FA_AhXYXpgList.Add;
    end;
    $00000362: begin
      if FA_AhPolarXpgList = Nil then 
        FA_AhPolarXpgList := TCT_PolarAdjustHandleXpgList.Create(FOwner);
      Result := FA_AhPolarXpgList.Add;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_AdjustHandleList.Write(AWriter: TXpgWriteXML);
begin
  if FA_AhXYXpgList <> Nil then 
    FA_AhXYXpgList.Write(AWriter,'a:ahXY');
  if FA_AhPolarXpgList <> Nil then 
    FA_AhPolarXpgList.Write(AWriter,'a:ahPolar');
end;

constructor TCT_AdjustHandleList.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 2;
  FAttributeCount := 0;
end;

destructor TCT_AdjustHandleList.Destroy;
begin
  if FA_AhXYXpgList <> Nil then 
    FA_AhXYXpgList.Free;
  if FA_AhPolarXpgList <> Nil then
    FA_AhPolarXpgList.Free;
end;

procedure TCT_AdjustHandleList.Clear;
begin
  FAssigneds := [];
  if FA_AhXYXpgList <> Nil then
    FreeAndNil(FA_AhXYXpgList);
  if FA_AhPolarXpgList <> Nil then
    FreeAndNil(FA_AhPolarXpgList);
end;

procedure TCT_AdjustHandleList.Assign(AItem: TCT_AdjustHandleList);
begin
  if AItem.FA_AhXYXpgList <> Nil then begin
    Create_A_AhXYXpgList;
    FA_AhXYXpgList.Assign(AItem.FA_AhXYXpgList);
  end;
  if AItem.FA_AhPolarXpgList <> Nil then begin
    Create_A_AhPolarXpgList;
    FA_AhPolarXpgList.Assign(AItem.FA_AhPolarXpgList);
  end;
end;

procedure TCT_AdjustHandleList.CopyTo(AItem: TCT_AdjustHandleList);
begin
end;

function  TCT_AdjustHandleList.Create_A_AhXYXpgList: TCT_XYAdjustHandleXpgList;
begin
  if FA_AhXYXpgList = Nil then
    FA_AhXYXpgList := TCT_XYAdjustHandleXpgList.Create(FOwner);
  Result := FA_AhXYXpgList;
end;

function  TCT_AdjustHandleList.Create_A_AhPolarXpgList: TCT_PolarAdjustHandleXpgList;
begin
  if FA_AhPolarXpgList = Nil then
    FA_AhPolarXpgList := TCT_PolarAdjustHandleXpgList.Create(FOwner);
  Result := FA_AhPolarXpgList;
end;

{ TCT_ConnectionSiteList }

function  TCT_ConnectionSiteList.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FA_CxnXpgList <> Nil then 
    Inc(ElemsAssigned,FA_CxnXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_ConnectionSiteList.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'a:cxn' then 
  begin
    if FA_CxnXpgList = Nil then 
      FA_CxnXpgList := TCT_ConnectionSiteXpgList.Create(FOwner);
    Result := FA_CxnXpgList.Add;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_ConnectionSiteList.Write(AWriter: TXpgWriteXML);
begin
  if FA_CxnXpgList <> Nil then 
    FA_CxnXpgList.Write(AWriter,'a:cxn');
end;

constructor TCT_ConnectionSiteList.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 0;
end;

destructor TCT_ConnectionSiteList.Destroy;
begin
  if FA_CxnXpgList <> Nil then 
    FA_CxnXpgList.Free;
end;

procedure TCT_ConnectionSiteList.Clear;
begin
  FAssigneds := [];
  if FA_CxnXpgList <> Nil then 
    FreeAndNil(FA_CxnXpgList);
end;

procedure TCT_ConnectionSiteList.Assign(AItem: TCT_ConnectionSiteList);
begin
  if AItem.FA_CxnXpgList <> Nil then begin
    Create_A_CxnXpgList;
    FA_CxnXpgList.Assign(AItem.FA_CxnXpgList);
  end;
end;

procedure TCT_ConnectionSiteList.CopyTo(AItem: TCT_ConnectionSiteList);
begin
end;

function  TCT_ConnectionSiteList.Create_A_CxnXpgList: TCT_ConnectionSiteXpgList;
begin
  if FA_CxnXpgList = Nil then
    FA_CxnXpgList := TCT_ConnectionSiteXpgList.Create(FOwner);
  Result := FA_CxnXpgList;
end;

{ TCT_GeomRect }

function  TCT_GeomRect.CheckAssigned: integer;
begin
  FAssigneds := [xaAttributes];
  Result := 1;
end;

procedure TCT_GeomRect.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_GeomRect.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('l',WriteUnionTST_AdjCoordinate(PFL));
  AWriter.AddAttribute('t',WriteUnionTST_AdjCoordinate(PFT));
  AWriter.AddAttribute('r',WriteUnionTST_AdjCoordinate(PFR));
  AWriter.AddAttribute('b',WriteUnionTST_AdjCoordinate(PFB));
end;

procedure TCT_GeomRect.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $0000006C: ReadUnionTST_AdjCoordinate(AAttributes.Values[i],PFL);
      $00000074: ReadUnionTST_AdjCoordinate(AAttributes.Values[i],PFT);
      $00000072: ReadUnionTST_AdjCoordinate(AAttributes.Values[i],PFR);
      $00000062: ReadUnionTST_AdjCoordinate(AAttributes.Values[i],PFB);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_GeomRect.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 4;
  PFL := @FL;
  PL.nVal := -1;
  PFT := @FT;
  PT.nVal := -1;
  PFR := @FR;
  PR.nVal := -1;
  PFB := @FB;
  PB.nVal := -1;
end;

destructor TCT_GeomRect.Destroy;
begin
end;

procedure TCT_GeomRect.Clear;
begin
  FAssigneds := [];
end;

procedure TCT_GeomRect.Assign(AItem: TCT_GeomRect);
begin
  FL := AItem.FL;
  FT := AItem.FT;
  FR := AItem.FR;
  FB := AItem.FB;
end;

procedure TCT_GeomRect.CopyTo(AItem: TCT_GeomRect);
begin
end;

{ TCT_Path2DList }

function  TCT_Path2DList.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FA_PathXpgList <> Nil then 
    Inc(ElemsAssigned,FA_PathXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_Path2DList.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'a:path' then 
  begin
    if FA_PathXpgList = Nil then 
      FA_PathXpgList := TCT_Path2DXpgList.Create(FOwner);
    Result := FA_PathXpgList.Add;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_Path2DList.Write(AWriter: TXpgWriteXML);
begin
  if FA_PathXpgList <> Nil then 
    FA_PathXpgList.Write(AWriter,'a:path');
end;

constructor TCT_Path2DList.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 0;
end;

destructor TCT_Path2DList.Destroy;
begin
  if FA_PathXpgList <> Nil then 
    FA_PathXpgList.Free;
end;

procedure TCT_Path2DList.Clear;
begin
  FAssigneds := [];
  if FA_PathXpgList <> Nil then 
    FreeAndNil(FA_PathXpgList);
end;

procedure TCT_Path2DList.Assign(AItem: TCT_Path2DList);
begin
  if AItem.FA_PathXpgList <> Nil then begin
    Create_A_PathXpgList;
    FA_PathXpgList.Assign(AItem.FA_PathXpgList);
  end;
end;

procedure TCT_Path2DList.CopyTo(AItem: TCT_Path2DList);
begin
end;

function  TCT_Path2DList.Create_A_PathXpgList: TCT_Path2DXpgList;
begin
  if FA_PathXpgList = Nil then
    FA_PathXpgList := TCT_Path2DXpgList.Create(FOwner);
  Result := FA_PathXpgList;
end;

{ TCT_TextNoAutofit }

function  TCT_TextNoAutofit.CheckAssigned: integer;
begin
  FAssigneds := [];
  Result := 0;
end;

procedure TCT_TextNoAutofit.Write(AWriter: TXpgWriteXML);
begin
end;

constructor TCT_TextNoAutofit.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 0;
end;

destructor TCT_TextNoAutofit.Destroy;
begin
end;

procedure TCT_TextNoAutofit.Clear;
begin
  FAssigneds := [];
end;

procedure TCT_TextNoAutofit.Assign(AItem: TCT_TextNoAutofit);
begin
end;

procedure TCT_TextNoAutofit.CopyTo(AItem: TCT_TextNoAutofit);
begin
end;

{ TCT_TextNormalAutofit }

function  TCT_TextNormalAutofit.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FFontScale <> 100000 then 
    Inc(AttrsAssigned);
  if FLnSpcReduction <> 0 then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_TextNormalAutofit.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_TextNormalAutofit.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FFontScale <> 100000 then 
    AWriter.AddAttribute('fontScale',XmlIntToStr(FFontScale));
  if FLnSpcReduction <> 0 then 
    AWriter.AddAttribute('lnSpcReduction',XmlIntToStr(FLnSpcReduction));
end;

procedure TCT_TextNormalAutofit.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $0000039F: FFontScale := XmlStrToIntDef(AAttributes.Values[i],0);
      $000005AD: FLnSpcReduction := XmlStrToIntDef(AAttributes.Values[i],0);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_TextNormalAutofit.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 2;
  FFontScale := 100000;
  FLnSpcReduction := 0;
end;

destructor TCT_TextNormalAutofit.Destroy;
begin
end;

procedure TCT_TextNormalAutofit.Clear;
begin
  FAssigneds := [];
  FFontScale := 100000;
  FLnSpcReduction := 0;
end;

procedure TCT_TextNormalAutofit.Assign(AItem: TCT_TextNormalAutofit);
begin
end;

procedure TCT_TextNormalAutofit.CopyTo(AItem: TCT_TextNormalAutofit);
begin
end;

{ TCT_TextShapeAutofit }

function  TCT_TextShapeAutofit.CheckAssigned: integer;
begin
  FAssigneds := [];
  Result := 0;
end;

procedure TCT_TextShapeAutofit.Write(AWriter: TXpgWriteXML);
begin
end;

constructor TCT_TextShapeAutofit.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 0;
end;

destructor TCT_TextShapeAutofit.Destroy;
begin
end;

procedure TCT_TextShapeAutofit.Clear;
begin
  FAssigneds := [];
end;

procedure TCT_TextShapeAutofit.Assign(AItem: TCT_TextShapeAutofit);
begin
end;

procedure TCT_TextShapeAutofit.CopyTo(AItem: TCT_TextShapeAutofit);
begin
end;

{ TCT_Camera }

function  TCT_Camera.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [xaAttributes];
  if FA_Rot <> Nil then
    Inc(ElemsAssigned,FA_Rot.CheckAssigned);
  Result := 1;
  if ElemsAssigned > 0 then
    FAssigneds := FAssigneds + [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_Camera.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'a:rot' then 
  begin
    if FA_Rot = Nil then 
      FA_Rot := TCT_SphereCoords.Create(FOwner);
    Result := FA_Rot;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_Camera.Write(AWriter: TXpgWriteXML);
begin
  if (FA_Rot <> Nil) and FA_Rot.Assigned then 
  begin
    FA_Rot.WriteAttributes(AWriter);
    AWriter.SimpleTag('a:rot');
  end;
end;

procedure TCT_Camera.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FPrst <> TST_PresetCameraType(XPG_UNKNOWN_ENUM) then
    AWriter.AddAttribute('prst',StrTST_PresetCameraType[Integer(FPrst)]);
  if FFov <> 2147483632 then 
    AWriter.AddAttribute('fov',XmlIntToStr(FFov));
  if FZoom <> 100000 then 
    AWriter.AddAttribute('zoom',XmlIntToStr(FZoom));
end;

procedure TCT_Camera.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $000001C9: FPrst := TST_PresetCameraType(StrToEnum('stpct' + AAttributes.Values[i]));
      $0000014B: FFov := XmlStrToIntDef(AAttributes.Values[i],0);
      $000001C5: FZoom := XmlStrToIntDef(AAttributes.Values[i],0);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_Camera.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 3;
  FPrst := TST_PresetCameraType(XPG_UNKNOWN_ENUM);
  FFov := 2147483632;
  FZoom := 100000;
end;

destructor TCT_Camera.Destroy;
begin
  if FA_Rot <> Nil then
    FA_Rot.Free;
end;

procedure TCT_Camera.Clear;
begin
  FAssigneds := [];
  if FA_Rot <> Nil then
    FreeAndNil(FA_Rot);
  FPrst := TST_PresetCameraType(XPG_UNKNOWN_ENUM);
  FFov := 2147483632;
  FZoom := 100000;
end;

procedure TCT_Camera.Assign(AItem: TCT_Camera);
begin
  FPrst := AItem.FPrst;
  FFov := AItem.FFov;
  FZoom := AItem.FZoom;
  if AItem.FA_Rot <> Nil then begin
    Create_A_Rot;
    FA_Rot.Assign(AItem.FA_Rot);
  end;
end;

procedure TCT_Camera.CopyTo(AItem: TCT_Camera);
begin
end;

function  TCT_Camera.Create_A_Rot: TCT_SphereCoords;
begin
  if FA_Rot = Nil then
    FA_Rot := TCT_SphereCoords.Create(FOwner);
  Result := FA_Rot;
end;

{ TCT_LightRig }

function  TCT_LightRig.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [xaAttributes];
  if FA_Rot <> Nil then
    Inc(ElemsAssigned,FA_Rot.CheckAssigned);
  Result := 1;
  if ElemsAssigned > 0 then
    FAssigneds := FAssigneds + [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_LightRig.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'a:rot' then 
  begin
    if FA_Rot = Nil then 
      FA_Rot := TCT_SphereCoords.Create(FOwner);
    Result := FA_Rot;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_LightRig.Write(AWriter: TXpgWriteXML);
begin
  if (FA_Rot <> Nil) and FA_Rot.Assigned then 
  begin
    FA_Rot.WriteAttributes(AWriter);
    AWriter.SimpleTag('a:rot');
  end;
end;

procedure TCT_LightRig.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FRig <> TST_LightRigType(XPG_UNKNOWN_ENUM) then
    AWriter.AddAttribute('rig',StrTST_LightRigType[Integer(FRig)]);
  if FDir <> TST_LightRigDirection(XPG_UNKNOWN_ENUM) then
    AWriter.AddAttribute('dir',StrTST_LightRigDirection[Integer(FDir)]);
end;

procedure TCT_LightRig.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
    case AAttributes.HashA[i] of
      $00000142: FRig := TST_LightRigType(StrToEnum('stlrt' + AAttributes.Values[i]));
      $0000013F: FDir := TST_LightRigDirection(StrToEnum('stlrd' + AAttributes.Values[i]));
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_LightRig.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 2;
  FRig := TST_LightRigType(XPG_UNKNOWN_ENUM);
  FDir := TST_LightRigDirection(XPG_UNKNOWN_ENUM);
end;

destructor TCT_LightRig.Destroy;
begin
  if FA_Rot <> Nil then
    FA_Rot.Free;
end;

procedure TCT_LightRig.Clear;
begin
  FAssigneds := [];
  if FA_Rot <> Nil then
    FreeAndNil(FA_Rot);
  FRig := TST_LightRigType(XPG_UNKNOWN_ENUM);
  FDir := TST_LightRigDirection(XPG_UNKNOWN_ENUM);
end;

procedure TCT_LightRig.Assign(AItem: TCT_LightRig);
begin
  FRig := AItem.FRig;
  FDir := AItem.FDir;
  if AItem.FA_Rot <> Nil then begin
    Create_A_Rot;
    FA_Rot.Assign(AItem.FA_Rot);
  end;
end;

procedure TCT_LightRig.CopyTo(AItem: TCT_LightRig);
begin
end;

function  TCT_LightRig.Create_A_Rot: TCT_SphereCoords;
begin
  if FA_Rot = Nil then
    FA_Rot := TCT_SphereCoords.Create(FOwner);
  Result := FA_Rot;
end;

{ TCT_Backdrop }

function  TCT_Backdrop.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FA_Anchor <> Nil then 
    Inc(ElemsAssigned,FA_Anchor.CheckAssigned);
  if FA_Norm <> Nil then 
    Inc(ElemsAssigned,FA_Norm.CheckAssigned);
  if FA_Up <> Nil then 
    Inc(ElemsAssigned,FA_Up.CheckAssigned);
  if FA_ExtLst <> Nil then 
    Inc(ElemsAssigned,FA_ExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_Backdrop.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $00000316: begin
      if FA_Anchor = Nil then 
        FA_Anchor := TCT_Point3D.Create(FOwner);
      Result := FA_Anchor;
    end;
    $00000257: begin
      if FA_Norm = Nil then 
        FA_Norm := TCT_Vector3D.Create(FOwner);
      Result := FA_Norm;
    end;
    $00000180: begin
      if FA_Up = Nil then 
        FA_Up := TCT_Vector3D.Create(FOwner);
      Result := FA_Up;
    end;
    $0000031F: begin
      if FA_ExtLst = Nil then 
        FA_ExtLst := TCT_OfficeArtExtensionList.Create(FOwner);
      Result := FA_ExtLst;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_Backdrop.Write(AWriter: TXpgWriteXML);
begin
  if (FA_Anchor <> Nil) and FA_Anchor.Assigned then 
  begin
    FA_Anchor.WriteAttributes(AWriter);
    AWriter.SimpleTag('a:anchor');
  end
  else 
    AWriter.SimpleTag('a:anchor');
  if (FA_Norm <> Nil) and FA_Norm.Assigned then 
  begin
    FA_Norm.WriteAttributes(AWriter);
    AWriter.SimpleTag('a:norm');
  end
  else 
    AWriter.SimpleTag('a:norm');
  if (FA_Up <> Nil) and FA_Up.Assigned then 
  begin
    FA_Up.WriteAttributes(AWriter);
    AWriter.SimpleTag('a:up');
  end
  else 
    AWriter.SimpleTag('a:up');
  if (FA_ExtLst <> Nil) and FA_ExtLst.Assigned then 
    if xaElements in FA_ExtLst.FAssigneds then 
    begin
      AWriter.BeginTag('a:extLst');
      FA_ExtLst.Write(AWriter);
      AWriter.EndTag;
    end;
end;

constructor TCT_Backdrop.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 4;
  FAttributeCount := 0;
end;

destructor TCT_Backdrop.Destroy;
begin
  if FA_Anchor <> Nil then 
    FA_Anchor.Free;
  if FA_Norm <> Nil then 
    FA_Norm.Free;
  if FA_Up <> Nil then 
    FA_Up.Free;
  if FA_ExtLst <> Nil then 
    FA_ExtLst.Free;
end;

procedure TCT_Backdrop.Clear;
begin
  FAssigneds := [];
  if FA_Anchor <> Nil then 
    FreeAndNil(FA_Anchor);
  if FA_Norm <> Nil then 
    FreeAndNil(FA_Norm);
  if FA_Up <> Nil then 
    FreeAndNil(FA_Up);
  if FA_ExtLst <> Nil then 
    FreeAndNil(FA_ExtLst);
end;

procedure TCT_Backdrop.Assign(AItem: TCT_Backdrop);
begin
  if AItem.FA_Anchor <> Nil then begin
    Create_A_Anchor;
    FA_Anchor.Assign(AItem.FA_Anchor);
  end;
  if AItem.FA_Norm <> Nil then begin
    Create_A_Norm;
    FA_Norm.Assign(AItem.FA_Norm);
  end;
  if AItem.FA_Up <> Nil then begin
    Create_A_Up;
    FA_Up.Assign(AItem.FA_Up);
  end;
  if AItem.FA_ExtLst <> Nil then begin
    Create_A_ExtLst;
    FA_ExtLst.Assign(AItem.FA_ExtLst);
  end;
end;

procedure TCT_Backdrop.CopyTo(AItem: TCT_Backdrop);
begin
end;

function  TCT_Backdrop.Create_A_Anchor: TCT_Point3D;
begin
  if FA_Anchor = Nil then
    FA_Anchor := TCT_Point3D.Create(FOwner);
  Result := FA_Anchor;
end;

function  TCT_Backdrop.Create_A_Norm: TCT_Vector3D;
begin
  if FA_Norm = Nil then
    FA_Norm := TCT_Vector3D.Create(FOwner);
  Result := FA_Norm;
end;

function  TCT_Backdrop.Create_A_Up: TCT_Vector3D;
begin
  if FA_Up = Nil then
    FA_Up := TCT_Vector3D.Create(FOwner);
  Result := FA_Up;
end;

function  TCT_Backdrop.Create_A_ExtLst: TCT_OfficeArtExtensionList;
begin
  if FA_ExtLst = Nil then
    FA_ExtLst := TCT_OfficeArtExtensionList.Create(FOwner);
  Result := FA_ExtLst;
end;

{ TCT_Shape3D }

function  TCT_Shape3D.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FZ <> 0 then
    Inc(AttrsAssigned);
  if FExtrusionH <> 0 then
    Inc(AttrsAssigned);
  if FContourW <> 0 then
    Inc(AttrsAssigned);
  if FPrstMaterial <> stpmtWarmMatte then
    Inc(AttrsAssigned);
  if FA_BevelT <> Nil then
    Inc(ElemsAssigned);
  if FA_BevelB <> Nil then
    Inc(ElemsAssigned);
  if FA_ExtrusionClr <> Nil then
    Inc(ElemsAssigned,FA_ExtrusionClr.CheckAssigned);
  if FA_ContourClr <> Nil then 
    Inc(ElemsAssigned,FA_ContourClr.CheckAssigned);
  if FA_ExtLst <> Nil then 
    Inc(ElemsAssigned,FA_ExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_Shape3D.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $000002FD: begin
      if FA_BevelT = Nil then 
        FA_BevelT := TCT_Bevel.Create(FOwner);
      Result := FA_BevelT;
    end;
    $000002EB: begin
      if FA_BevelB = Nil then 
        FA_BevelB := TCT_Bevel.Create(FOwner);
      Result := FA_BevelB;
    end;
    $000005AD: begin
      if FA_ExtrusionClr = Nil then 
        FA_ExtrusionClr := TCT_Color.Create(FOwner);
      Result := FA_ExtrusionClr;
    end;
    $000004C6: begin
      if FA_ContourClr = Nil then 
        FA_ContourClr := TCT_Color.Create(FOwner);
      Result := FA_ContourClr;
    end;
    $0000031F: begin
      if FA_ExtLst = Nil then 
        FA_ExtLst := TCT_OfficeArtExtensionList.Create(FOwner);
      Result := FA_ExtLst;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_Shape3D.Write(AWriter: TXpgWriteXML);
begin
  if (FA_BevelT <> Nil) and FA_BevelT.Assigned then 
  begin
    FA_BevelT.WriteAttributes(AWriter);
    AWriter.SimpleTag('a:bevelT');
  end;
  if (FA_BevelB <> Nil) and FA_BevelB.Assigned then 
  begin
    FA_BevelB.WriteAttributes(AWriter);
    AWriter.SimpleTag('a:bevelB');
  end;
  if (FA_ExtrusionClr <> Nil) and FA_ExtrusionClr.Assigned then 
    if xaElements in FA_ExtrusionClr.FAssigneds then 
    begin
      AWriter.BeginTag('a:extrusionClr');
      FA_ExtrusionClr.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:extrusionClr');
  if (FA_ContourClr <> Nil) and FA_ContourClr.Assigned then 
    if xaElements in FA_ContourClr.FAssigneds then 
    begin
      AWriter.BeginTag('a:contourClr');
      FA_ContourClr.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:contourClr');
  if (FA_ExtLst <> Nil) and FA_ExtLst.Assigned then 
    if xaElements in FA_ExtLst.FAssigneds then 
    begin
      AWriter.BeginTag('a:extLst');
      FA_ExtLst.Write(AWriter);
      AWriter.EndTag;
    end;
end;

procedure TCT_Shape3D.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FZ <> 0 then 
    AWriter.AddAttribute('z',XmlIntToStr(FZ));
  if FExtrusionH <> 0 then 
    AWriter.AddAttribute('extrusionH',XmlIntToStr(FExtrusionH));
  if FContourW <> 0 then 
    AWriter.AddAttribute('contourW',XmlIntToStr(FContourW));
  if FPrstMaterial <> stpmtWarmMatte then 
    AWriter.AddAttribute('prstMaterial',StrTST_PresetMaterialType[Integer(FPrstMaterial)]);
end;

procedure TCT_Shape3D.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $0000007A: FZ := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000439: FExtrusionH := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000361: FContourW := XmlStrToIntDef(AAttributes.Values[i],0);
      $000004F8: FPrstMaterial := TST_PresetMaterialType(StrToEnum('stpmt' + AAttributes.Values[i]));
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_Shape3D.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 5;
  FAttributeCount := 4;
  FZ := 0;
  FExtrusionH := 0;
  FContourW := 0;
  FPrstMaterial := stpmtWarmMatte;
end;

destructor TCT_Shape3D.Destroy;
begin
  if FA_BevelT <> Nil then
    FA_BevelT.Free;
  if FA_BevelB <> Nil then 
    FA_BevelB.Free;
  if FA_ExtrusionClr <> Nil then 
    FA_ExtrusionClr.Free;
  if FA_ContourClr <> Nil then 
    FA_ContourClr.Free;
  if FA_ExtLst <> Nil then 
    FA_ExtLst.Free;
end;

procedure TCT_Shape3D.Clear;
begin
  FAssigneds := [];
  if FA_BevelT <> Nil then 
    FreeAndNil(FA_BevelT);
  if FA_BevelB <> Nil then 
    FreeAndNil(FA_BevelB);
  if FA_ExtrusionClr <> Nil then 
    FreeAndNil(FA_ExtrusionClr);
  if FA_ContourClr <> Nil then 
    FreeAndNil(FA_ContourClr);
  if FA_ExtLst <> Nil then 
    FreeAndNil(FA_ExtLst);
  FZ := 0;
  FExtrusionH := 0;
  FContourW := 0;
  FPrstMaterial := stpmtWarmMatte;
end;

procedure TCT_Shape3D.Assign(AItem: TCT_Shape3D);
begin
  FZ := AItem.FZ;
  FExtrusionH := AItem.FExtrusionH;
  FContourW := AItem.FContourW;
  FPrstMaterial := AItem.FPrstMaterial;

  if AItem.FA_BevelT <> Nil then begin
    Create_A_BevelT;
    FA_BevelT.Assign(AItem.FA_BevelT);
  end;
  if AItem.FA_BevelB <> Nil then begin
    Create_A_BevelB;
    FA_BevelB.Assign(AItem.FA_BevelB);
  end;
  if AItem.FA_ExtrusionClr <> Nil then begin
    Create_A_ExtrusionClr;
    FA_ExtrusionClr.Assign(AItem.FA_ExtrusionClr);
  end;
  if AItem.FA_ContourClr <> Nil then begin
    Create_A_ContourClr;
    FA_ContourClr.Assign(AItem.FA_ContourClr);
  end;
  if AItem.FA_ExtLst <> Nil then begin
    Create_A_ExtLst;
    FA_ExtLst.Assign(AItem.FA_ExtLst);
  end;
end;

procedure TCT_Shape3D.CopyTo(AItem: TCT_Shape3D);
begin
end;

function  TCT_Shape3D.Create_A_BevelT: TCT_Bevel;
begin
  if FA_BevelT = Nil then
    FA_BevelT := TCT_Bevel.Create(FOwner);
  Result := FA_BevelT;
end;

function  TCT_Shape3D.Create_A_BevelB: TCT_Bevel;
begin
  if FA_BevelB = Nil then
    FA_BevelB := TCT_Bevel.Create(FOwner);
  Result := FA_BevelB;
end;

function  TCT_Shape3D.Create_A_ExtrusionClr: TCT_Color;
begin
  if FA_ExtrusionClr = Nil then
    FA_ExtrusionClr := TCT_Color.Create(FOwner);
  Result := FA_ExtrusionClr;
end;

function  TCT_Shape3D.Create_A_ContourClr: TCT_Color;
begin
  if FA_ContourClr = Nil then
    FA_ContourClr := TCT_Color.Create(FOwner);
  Result := FA_ContourClr;
end;

function  TCT_Shape3D.Create_A_ExtLst: TCT_OfficeArtExtensionList;
begin
  if FA_ExtLst = Nil then
    FA_ExtLst := TCT_OfficeArtExtensionList.Create(FOwner);
  Result := FA_ExtLst;
end;

{ TCT_FlatText }

function  TCT_FlatText.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FZ <> 0 then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_FlatText.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_FlatText.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FZ <> 0 then 
    AWriter.AddAttribute('z',XmlIntToStr(FZ));
end;

procedure TCT_FlatText.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'z' then 
    FZ := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_FlatText.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
  FZ := 0;
end;

destructor TCT_FlatText.Destroy;
begin
end;

procedure TCT_FlatText.Clear;
begin
  FAssigneds := [];
  FZ := 0;
end;

procedure TCT_FlatText.Assign(AItem: TCT_FlatText);
begin
end;

procedure TCT_FlatText.CopyTo(AItem: TCT_FlatText);
begin
end;

{ TCT_RegularTextRun }

function  TCT_RegularTextRun.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FA_RPr <> Nil then
    Inc(ElemsAssigned,FA_RPr.CheckAssigned);
  if FA_T <> '' then
    Inc(ElemsAssigned);
  Result := 0;
  if ElemsAssigned > 0 then
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_RegularTextRun.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $000001CF: begin
      if FA_RPr = Nil then 
        FA_RPr := TCT_TextCharacterProperties.Create(FOwner);
      Result := FA_RPr;
    end;
    $0000010F: FA_T := AReader.Text;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_RegularTextRun.Write(AWriter: TXpgWriteXML);
begin
  if (FA_RPr <> Nil) and FA_RPr.Assigned then 
  begin
    FA_RPr.WriteAttributes(AWriter);
    if xaElements in FA_RPr.FAssigneds then 
    begin
      AWriter.BeginTag('a:rPr');
      FA_RPr.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:rPr');
  end;
  AWriter.SimpleTextTag('a:t',FA_T);
end;

constructor TCT_RegularTextRun.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 2;
  FAttributeCount := 0;
end;

destructor TCT_RegularTextRun.Destroy;
begin
  if FA_RPr <> Nil then 
    FA_RPr.Free;
end;

procedure TCT_RegularTextRun.Clear;
begin
  FAssigneds := [];
  if FA_RPr <> Nil then 
    FreeAndNil(FA_RPr);
  FA_T := '';
end;

procedure TCT_RegularTextRun.Assign(AItem: TCT_RegularTextRun);
begin
end;

procedure TCT_RegularTextRun.CopyTo(AItem: TCT_RegularTextRun);
begin
end;

function  TCT_RegularTextRun.Create_RPr: TCT_TextCharacterProperties;
begin
  if FA_RPr = Nil then
    FA_RPr := TCT_TextCharacterProperties.Create(FOwner);
  Result := FA_RPr;
end;

{ TCT_TextLineBreak }

function  TCT_TextLineBreak.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FA_RPr <> Nil then
    Inc(ElemsAssigned,FA_RPr.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_TextLineBreak.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'a:rPr' then 
  begin
    if FA_RPr = Nil then 
      FA_RPr := TCT_TextCharacterProperties.Create(FOwner);
    Result := FA_RPr;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_TextLineBreak.Write(AWriter: TXpgWriteXML);
begin
  if (FA_RPr <> Nil) and FA_RPr.Assigned then 
  begin
    FA_RPr.WriteAttributes(AWriter);
    if xaElements in FA_RPr.FAssigneds then 
    begin
      AWriter.BeginTag('a:rPr');
      FA_RPr.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:rPr');
  end;
end;

constructor TCT_TextLineBreak.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 0;
end;

destructor TCT_TextLineBreak.Destroy;
begin
  if FA_RPr <> Nil then 
    FA_RPr.Free;
end;

procedure TCT_TextLineBreak.Clear;
begin
  FAssigneds := [];
  if FA_RPr <> Nil then 
    FreeAndNil(FA_RPr);
end;

procedure TCT_TextLineBreak.Assign(AItem: TCT_TextLineBreak);
begin
end;

procedure TCT_TextLineBreak.CopyTo(AItem: TCT_TextLineBreak);
begin
end;

function  TCT_TextLineBreak.Create_RPr: TCT_TextCharacterProperties;
begin
  if FA_RPr = Nil then
    FA_RPr := TCT_TextCharacterProperties.Create(FOwner);
  Result := FA_RPr;
end;

{ TCT_TextField }

function  TCT_TextField.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [xaAttributes];
  if FA_RPr <> Nil then
    Inc(ElemsAssigned,FA_RPr.CheckAssigned);
  if FA_PPr <> Nil then
    Inc(ElemsAssigned,FA_PPr.CheckAssigned);
  if FA_T <> '' then
    Inc(ElemsAssigned);
  Result := 1;
  if ElemsAssigned > 0 then
    FAssigneds := FAssigneds + [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_TextField.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $000001CF: begin
      if FA_RPr = Nil then 
        FA_RPr := TCT_TextCharacterProperties.Create(FOwner);
      Result := FA_RPr;
    end;
    $000001CD: begin
      if FA_PPr = Nil then 
        FA_PPr := TCT_TextParagraphProperties.Create(FOwner);
      Result := FA_PPr;
    end;
    $0000010F: FA_T := AReader.Text;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_TextField.Write(AWriter: TXpgWriteXML);
begin
  if (FA_RPr <> Nil) and FA_RPr.Assigned then 
  begin
    FA_RPr.WriteAttributes(AWriter);
    if xaElements in FA_RPr.FAssigneds then 
    begin
      AWriter.BeginTag('a:rPr');
      FA_RPr.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:rPr');
  end;
  if (FA_PPr <> Nil) and FA_PPr.Assigned then 
  begin
    FA_PPr.WriteAttributes(AWriter);
    if xaElements in FA_PPr.FAssigneds then 
    begin
      AWriter.BeginTag('a:pPr');
      FA_PPr.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:pPr');
  end;
  if FA_T <> '' then 
    AWriter.SimpleTextTag('a:t',FA_T);
end;

procedure TCT_TextField.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('id',FId);
  if FType <> '' then 
    AWriter.AddAttribute('type',FType);
end;

procedure TCT_TextField.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $000000CD: FId := AAttributes.Values[i];
      $000001C2: FType := AAttributes.Values[i];
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_TextField.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 3;
  FAttributeCount := 2;
  FId := '';
end;

destructor TCT_TextField.Destroy;
begin
  if FA_RPr <> Nil then 
    FA_RPr.Free;
  if FA_PPr <> Nil then 
    FA_PPr.Free;
end;

procedure TCT_TextField.Clear;
begin
  FAssigneds := [];
  if FA_RPr <> Nil then 
    FreeAndNil(FA_RPr);
  if FA_PPr <> Nil then 
    FreeAndNil(FA_PPr);
  FA_T := '';
  FId := '';
  FType := '';
end;

procedure TCT_TextField.Assign(AItem: TCT_TextField);
begin
end;

procedure TCT_TextField.CopyTo(AItem: TCT_TextField);
begin
end;

function  TCT_TextField.Create_A_RPr: TCT_TextCharacterProperties;
begin
  if FA_RPr = Nil then
    FA_RPr := TCT_TextCharacterProperties.Create(FOwner);
  Result := FA_RPr;
end;

function  TCT_TextField.Create_A_PPr: TCT_TextParagraphProperties;
begin
  if FA_PPr = Nil then
    FA_PPr := TCT_TextParagraphProperties.Create(FOwner);
  Result := FA_PPr;
end;

{ TCT_Point2D }

function  TCT_Point2D.CheckAssigned: integer;
begin
  FAssigneds := [xaAttributes];
  Result := 1;
end;

procedure TCT_Point2D.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_Point2D.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('x',XmlIntToStr(FX));
  AWriter.AddAttribute('y',XmlIntToStr(FY));
end;

procedure TCT_Point2D.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $00000078: FX := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000079: FY := XmlStrToIntDef(AAttributes.Values[i],0);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_Point2D.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 2;
  FX := 2147483632;
  FY := 2147483632;
end;

destructor TCT_Point2D.Destroy;
begin
end;

procedure TCT_Point2D.Clear;
begin
  FAssigneds := [];
  FX := 2147483632;
  FY := 2147483632;
end;

procedure TCT_Point2D.Assign(AItem: TCT_Point2D);
begin
  FX := AItem.FX;
  FY := AItem.FY;
end;

procedure TCT_Point2D.CopyTo(AItem: TCT_Point2D);
begin
end;

{ TCT_PositiveSize2D }

function  TCT_PositiveSize2D.CheckAssigned: integer;
begin
  FAssigneds := [xaAttributes];
  Result := 1;
end;

procedure TCT_PositiveSize2D.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_PositiveSize2D.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('cx',XmlIntToStr(FCx));
  AWriter.AddAttribute('cy',XmlIntToStr(FCy));
end;

procedure TCT_PositiveSize2D.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $000000DB: FCx := XmlStrToIntDef(AAttributes.Values[i],0);
      $000000DC: FCy := XmlStrToIntDef(AAttributes.Values[i],0);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_PositiveSize2D.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 2;
  FCx := 2147483632;
  FCy := 2147483632;
end;

destructor TCT_PositiveSize2D.Destroy;
begin
end;

procedure TCT_PositiveSize2D.Clear;
begin
  FAssigneds := [];
  FCx := 2147483632;
  FCy := 2147483632;
end;

procedure TCT_PositiveSize2D.Assign(AItem: TCT_PositiveSize2D);
begin
  FCx := AItem.FCx;
  FCy := AItem.FCy;
end;

procedure TCT_PositiveSize2D.CopyTo(AItem: TCT_PositiveSize2D);
begin
end;

{ TCT_CustomGeometry2D }

function  TCT_CustomGeometry2D.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FA_AvLst <> Nil then 
    Inc(ElemsAssigned,FA_AvLst.CheckAssigned);
  if FA_GdLst <> Nil then 
    Inc(ElemsAssigned,FA_GdLst.CheckAssigned);
  if FA_AhLst <> Nil then 
    Inc(ElemsAssigned,FA_AhLst.CheckAssigned);
  if FA_CxnLst <> Nil then 
    Inc(ElemsAssigned,FA_CxnLst.CheckAssigned);
  if FA_Rect <> Nil then 
    Inc(ElemsAssigned,FA_Rect.CheckAssigned);
  if FA_PathLst <> Nil then 
    Inc(ElemsAssigned,FA_PathLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_CustomGeometry2D.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $000002A5: begin
      if FA_AvLst = Nil then 
        FA_AvLst := TCT_GeomGuideList.Create(FOwner);
      Result := FA_AvLst;
    end;
    $00000299: begin
      if FA_GdLst = Nil then 
        FA_GdLst := TCT_GeomGuideList.Create(FOwner);
      Result := FA_GdLst;
    end;
    $00000297: begin
      if FA_AhLst = Nil then 
        FA_AhLst := TCT_AdjustHandleList.Create(FOwner);
      Result := FA_AhLst;
    end;
    $00000317: begin
      if FA_CxnLst = Nil then 
        FA_CxnLst := TCT_ConnectionSiteList.Create(FOwner);
      Result := FA_CxnLst;
    end;
    $00000249: begin
      if FA_Rect = Nil then 
        FA_Rect := TCT_GeomRect.Create(FOwner);
      Result := FA_Rect;
    end;
    $0000037B: begin
      if FA_PathLst = Nil then 
        FA_PathLst := TCT_Path2DList.Create(FOwner);
      Result := FA_PathLst;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_CustomGeometry2D.Write(AWriter: TXpgWriteXML);
begin
  if (FA_AvLst <> Nil) and FA_AvLst.Assigned then 
    if xaElements in FA_AvLst.FAssigneds then 
    begin
      AWriter.BeginTag('a:avLst');
      FA_AvLst.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:avLst');
  if (FA_GdLst <> Nil) and FA_GdLst.Assigned then 
    if xaElements in FA_GdLst.FAssigneds then 
    begin
      AWriter.BeginTag('a:gdLst');
      FA_GdLst.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:gdLst');
  if (FA_AhLst <> Nil) and FA_AhLst.Assigned then 
    if xaElements in FA_AhLst.FAssigneds then 
    begin
      AWriter.BeginTag('a:ahLst');
      FA_AhLst.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:ahLst');
  if (FA_CxnLst <> Nil) and FA_CxnLst.Assigned then 
    if xaElements in FA_CxnLst.FAssigneds then 
    begin
      AWriter.BeginTag('a:cxnLst');
      FA_CxnLst.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:cxnLst');
  if (FA_Rect <> Nil) and FA_Rect.Assigned then 
  begin
    FA_Rect.WriteAttributes(AWriter);
    AWriter.SimpleTag('a:rect');
  end;
  if (FA_PathLst <> Nil) and FA_PathLst.Assigned then 
    if xaElements in FA_PathLst.FAssigneds then 
    begin
      AWriter.BeginTag('a:pathLst');
      FA_PathLst.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:pathLst')
  else 
    AWriter.SimpleTag('a:pathLst');
end;

constructor TCT_CustomGeometry2D.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 6;
  FAttributeCount := 0;
end;

destructor TCT_CustomGeometry2D.Destroy;
begin
  if FA_AvLst <> Nil then 
    FA_AvLst.Free;
  if FA_GdLst <> Nil then 
    FA_GdLst.Free;
  if FA_AhLst <> Nil then 
    FA_AhLst.Free;
  if FA_CxnLst <> Nil then 
    FA_CxnLst.Free;
  if FA_Rect <> Nil then
    FA_Rect.Free;
  if FA_PathLst <> Nil then
    FA_PathLst.Free;
end;

procedure TCT_CustomGeometry2D.Clear;
begin
  FAssigneds := [];
  if FA_AvLst <> Nil then
    FreeAndNil(FA_AvLst);
  if FA_GdLst <> Nil then
    FreeAndNil(FA_GdLst);
  if FA_AhLst <> Nil then
    FreeAndNil(FA_AhLst);
  if FA_CxnLst <> Nil then
    FreeAndNil(FA_CxnLst);
  if FA_Rect <> Nil then
    FreeAndNil(FA_Rect);
  if FA_PathLst <> Nil then
    FreeAndNil(FA_PathLst);
end;

procedure TCT_CustomGeometry2D.Assign(AItem: TCT_CustomGeometry2D);
begin
  if AItem.FA_AvLst <> Nil then begin
    Create_A_AvLst;
    FA_AvLst.Assign(AItem.FA_AvLst);
  end;
  if AItem.FA_GdLst <> Nil then begin
    Create_A_GdLst;
    FA_GdLst.Assign(AItem.FA_GdLst);
  end;
  if AItem.FA_AhLst <> Nil then begin
    Create_A_AhLst;
    FA_AhLst.Assign(AItem.FA_AhLst);
  end;
  if AItem.FA_CxnLst <> Nil then begin
    Create_A_CxnLst;
    FA_CxnLst.Assign(AItem.FA_CxnLst);
  end;
  if AItem.FA_Rect <> Nil then begin
    Create_A_Rect;
    FA_Rect.Assign(AItem.FA_Rect);
  end;
  if AItem.FA_PathLst <> Nil then begin
    Create_A_PathLst;
    FA_PathLst.Assign(AItem.FA_PathLst);
  end;
end;

procedure TCT_CustomGeometry2D.CopyTo(AItem: TCT_CustomGeometry2D);
begin
end;

function  TCT_CustomGeometry2D.Create_A_AvLst: TCT_GeomGuideList;
begin
  if FA_AvLst = Nil then
    FA_AvLst := TCT_GeomGuideList.Create(FOwner);
  Result := FA_AvLst;
end;

function  TCT_CustomGeometry2D.Create_A_GdLst: TCT_GeomGuideList;
begin
  if FA_GdLst = Nil then
    FA_GdLst := TCT_GeomGuideList.Create(FOwner);
  Result := FA_GdLst;
end;

function  TCT_CustomGeometry2D.Create_A_AhLst: TCT_AdjustHandleList;
begin
  if FA_AhLst = Nil then
    FA_AhLst := TCT_AdjustHandleList.Create(FOwner);
  Result := FA_AhLst;
end;

function  TCT_CustomGeometry2D.Create_A_CxnLst: TCT_ConnectionSiteList;
begin
  if FA_CxnLst = Nil then
    FA_CxnLst := TCT_ConnectionSiteList.Create(FOwner);
  Result := FA_CxnLst;
end;

function  TCT_CustomGeometry2D.Create_A_Rect: TCT_GeomRect;
begin
  if FA_Rect = Nil then
    FA_Rect := TCT_GeomRect.Create(FOwner);
  Result := FA_Rect;
end;

function  TCT_CustomGeometry2D.Create_A_PathLst: TCT_Path2DList;
begin
  if FA_PathLst = Nil then
    FA_PathLst := TCT_Path2DList.Create(FOwner);
  Result := FA_PathLst;
end;

{ TCT_PresetGeometry2D }

function  TCT_PresetGeometry2D.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [xaAttributes];
  if FA_AvLst <> Nil then
    Inc(ElemsAssigned,FA_AvLst.CheckAssigned);
  Result := 1;
  if ElemsAssigned > 0 then
    FAssigneds := FAssigneds + [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_PresetGeometry2D.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'a:avLst' then 
  begin
    if FA_AvLst = Nil then 
      FA_AvLst := TCT_GeomGuideList.Create(FOwner);
    Result := FA_AvLst;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_PresetGeometry2D.Write(AWriter: TXpgWriteXML);
begin
  if FA_AvLst <> Nil then
    if xaElements in FA_AvLst.FAssigneds then 
    begin
      AWriter.BeginTag('a:avLst');
      FA_AvLst.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:avLst');
end;

procedure TCT_PresetGeometry2D.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FPrst <> TST_ShapeType(XPG_UNKNOWN_ENUM) then
    AWriter.AddAttribute('prst',StrTST_ShapeType[Integer(FPrst)]);
end;

procedure TCT_PresetGeometry2D.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'prst' then 
    FPrst := TST_ShapeType(StrToEnum('stst' + AAttributes.Values[0]))
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_PresetGeometry2D.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 1;
  FPrst := TST_ShapeType(XPG_UNKNOWN_ENUM);
end;

destructor TCT_PresetGeometry2D.Destroy;
begin
  if FA_AvLst <> Nil then
    FA_AvLst.Free;
end;

procedure TCT_PresetGeometry2D.Clear;
begin
  FAssigneds := [];
  if FA_AvLst <> Nil then
    FreeAndNil(FA_AvLst);
  FPrst := TST_ShapeType(XPG_UNKNOWN_ENUM);
end;

procedure TCT_PresetGeometry2D.Assign(AItem: TCT_PresetGeometry2D);
begin
end;

procedure TCT_PresetGeometry2D.CopyTo(AItem: TCT_PresetGeometry2D);
begin
end;

function  TCT_PresetGeometry2D.Create_A_AvLst: TCT_GeomGuideList;
begin
  if FA_AvLst = Nil then
    FA_AvLst := TCT_GeomGuideList.Create(FOwner);
  Result := FA_AvLst;
end;

{ TCT_PresetTextShape }

function  TCT_PresetTextShape.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [xaAttributes];
  if FA_AvLst <> Nil then
    Inc(ElemsAssigned,FA_AvLst.CheckAssigned);
  Result := 1;
  if ElemsAssigned > 0 then
    FAssigneds := FAssigneds + [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_PresetTextShape.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'a:avLst' then
  begin
    if FA_AvLst = Nil then
      FA_AvLst := TCT_GeomGuideList.Create(FOwner);
    Result := FA_AvLst;
  end
  else
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then
    Result.Assigneds := [xaRead];
end;

procedure TCT_PresetTextShape.Write(AWriter: TXpgWriteXML);
begin
  if (FA_AvLst <> Nil) and FA_AvLst.Assigned then
    if xaElements in FA_AvLst.FAssigneds then
    begin
      AWriter.BeginTag('a:avLst');
      FA_AvLst.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('a:avLst');
end;

procedure TCT_PresetTextShape.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FPrst <> TST_TextShapeType(XPG_UNKNOWN_ENUM) then
    AWriter.AddAttribute('prst',StrTST_TextShapeType[Integer(FPrst)]);
end;

procedure TCT_PresetTextShape.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'prst' then 
    FPrst := TST_TextShapeType(StrToEnum('sttst' + AAttributes.Values[0]))
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_PresetTextShape.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 1;
  FPrst := TST_TextShapeType(XPG_UNKNOWN_ENUM);
end;

destructor TCT_PresetTextShape.Destroy;
begin
  if FA_AvLst <> Nil then
    FA_AvLst.Free;
end;

procedure TCT_PresetTextShape.Clear;
begin
  FAssigneds := [];
  if FA_AvLst <> Nil then
    FreeAndNil(FA_AvLst);
  FPrst := TST_TextShapeType(XPG_UNKNOWN_ENUM);
end;

procedure TCT_PresetTextShape.Assign(AItem: TCT_PresetTextShape);
begin
end;

procedure TCT_PresetTextShape.CopyTo(AItem: TCT_PresetTextShape);
begin
end;

function  TCT_PresetTextShape.Create_A_AvLst: TCT_GeomGuideList;
begin
  if FA_AvLst = Nil then
    FA_AvLst := TCT_GeomGuideList.Create(FOwner);
  Result := FA_AvLst;
end;

{ TEG_TextAutofit }

function  TEG_TextAutofit.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FA_NoAutofit <> Nil then 
    Inc(ElemsAssigned,FA_NoAutofit.CheckAssigned);
  if FA_NormAutofit <> Nil then 
    Inc(ElemsAssigned,FA_NormAutofit.CheckAssigned);
  if FA_SpAutoFit <> Nil then 
    Inc(ElemsAssigned,FA_SpAutoFit.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TEG_TextAutofit.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  case AReader.QNameHashA of
    $00000454: begin
      if FA_NoAutofit = Nil then 
        FA_NoAutofit := TCT_TextNoAutofit.Create(FOwner);
      Result := FA_NoAutofit;
    end;
    $00000533: begin
      if FA_NormAutofit = Nil then 
        FA_NormAutofit := TCT_TextNormalAutofit.Create(FOwner);
      Result := FA_NormAutofit;
    end;
    $0000043A: begin
      if FA_SpAutoFit = Nil then 
        FA_SpAutoFit := TCT_TextShapeAutofit.Create(FOwner);
      Result := FA_SpAutoFit;
    end;
    else 
    begin
      Result := Nil;
      Exit;
    end
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TEG_TextAutofit.Write(AWriter: TXpgWriteXML);
begin
  if (FA_NoAutofit <> Nil) and FA_NoAutofit.Assigned then 
    AWriter.SimpleTag('a:noAutofit');
  if (FA_NormAutofit <> Nil) and FA_NormAutofit.Assigned then 
  begin
    FA_NormAutofit.WriteAttributes(AWriter);
    AWriter.SimpleTag('a:normAutofit');
  end;
  if (FA_SpAutoFit <> Nil) and FA_SpAutoFit.Assigned then 
    AWriter.SimpleTag('a:spAutoFit');
end;

constructor TEG_TextAutofit.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 3;
  FAttributeCount := 0;
end;

destructor TEG_TextAutofit.Destroy;
begin
  if FA_NoAutofit <> Nil then 
    FA_NoAutofit.Free;
  if FA_NormAutofit <> Nil then 
    FA_NormAutofit.Free;
  if FA_SpAutoFit <> Nil then 
    FA_SpAutoFit.Free;
end;

procedure TEG_TextAutofit.Clear;
begin
  FAssigneds := [];
  if FA_NoAutofit <> Nil then 
    FreeAndNil(FA_NoAutofit);
  if FA_NormAutofit <> Nil then 
    FreeAndNil(FA_NormAutofit);
  if FA_SpAutoFit <> Nil then 
    FreeAndNil(FA_SpAutoFit);
end;

procedure TEG_TextAutofit.Assign(AItem: TEG_TextAutofit);
begin
end;

procedure TEG_TextAutofit.CopyTo(AItem: TEG_TextAutofit);
begin
end;

function  TEG_TextAutofit.Create_A_NoAutofit: TCT_TextNoAutofit;
begin
  if FA_NoAutofit = Nil then
    FA_NoAutofit := TCT_TextNoAutofit.Create(FOwner);
  Result := FA_NoAutofit;
end;

function  TEG_TextAutofit.Create_A_NormAutofit: TCT_TextNormalAutofit;
begin
  if FA_NormAutofit = Nil then
    FA_NormAutofit := TCT_TextNormalAutofit.Create(FOwner);
  Result := FA_NormAutofit;
end;

function  TEG_TextAutofit.Create_A_SpAutoFit: TCT_TextShapeAutofit;
begin
  if FA_SpAutoFit = Nil then
    FA_SpAutoFit := TCT_TextShapeAutofit.Create(FOwner);
  Result := FA_SpAutoFit;
end;

{ TCT_Scene3D }

function  TCT_Scene3D.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FA_Camera <> Nil then 
    Inc(ElemsAssigned,FA_Camera.CheckAssigned);
  if FA_LightRig <> Nil then 
    Inc(ElemsAssigned,FA_LightRig.CheckAssigned);
  if FA_Backdrop <> Nil then 
    Inc(ElemsAssigned,FA_Backdrop.CheckAssigned);
  if FA_ExtLst <> Nil then 
    Inc(ElemsAssigned,FA_ExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_Scene3D.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $00000304: begin
      if FA_Camera = Nil then 
        FA_Camera := TCT_Camera.Create(FOwner);
      Result := FA_Camera;
    end;
    $000003D5: begin
      if FA_LightRig = Nil then 
        FA_LightRig := TCT_LightRig.Create(FOwner);
      Result := FA_LightRig;
    end;
    $000003E1: begin
      if FA_Backdrop = Nil then 
        FA_Backdrop := TCT_Backdrop.Create(FOwner);
      Result := FA_Backdrop;
    end;
    $0000031F: begin
      if FA_ExtLst = Nil then 
        FA_ExtLst := TCT_OfficeArtExtensionList.Create(FOwner);
      Result := FA_ExtLst;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_Scene3D.Write(AWriter: TXpgWriteXML);
begin
  if (FA_Camera <> Nil) and FA_Camera.Assigned then 
  begin
    FA_Camera.WriteAttributes(AWriter);
    if xaElements in FA_Camera.FAssigneds then 
    begin
      AWriter.BeginTag('a:camera');
      FA_Camera.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:camera');
  end
  else 
    AWriter.SimpleTag('a:camera');
  if (FA_LightRig <> Nil) and FA_LightRig.Assigned then 
  begin
    FA_LightRig.WriteAttributes(AWriter);
    if xaElements in FA_LightRig.FAssigneds then 
    begin
      AWriter.BeginTag('a:lightRig');
      FA_LightRig.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:lightRig');
  end
  else 
    AWriter.SimpleTag('a:lightRig');
  if (FA_Backdrop <> Nil) and FA_Backdrop.Assigned then 
    if xaElements in FA_Backdrop.FAssigneds then 
    begin
      AWriter.BeginTag('a:backdrop');
      FA_Backdrop.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:backdrop');
  if (FA_ExtLst <> Nil) and FA_ExtLst.Assigned then 
    if xaElements in FA_ExtLst.FAssigneds then 
    begin
      AWriter.BeginTag('a:extLst');
      FA_ExtLst.Write(AWriter);
      AWriter.EndTag;
    end;
end;

constructor TCT_Scene3D.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 4;
  FAttributeCount := 0;
end;

destructor TCT_Scene3D.Destroy;
begin
  if FA_Camera <> Nil then 
    FA_Camera.Free;
  if FA_LightRig <> Nil then 
    FA_LightRig.Free;
  if FA_Backdrop <> Nil then 
    FA_Backdrop.Free;
  if FA_ExtLst <> Nil then 
    FA_ExtLst.Free;
end;

procedure TCT_Scene3D.Clear;
begin
  FAssigneds := [];
  if FA_Camera <> Nil then 
    FreeAndNil(FA_Camera);
  if FA_LightRig <> Nil then 
    FreeAndNil(FA_LightRig);
  if FA_Backdrop <> Nil then 
    FreeAndNil(FA_Backdrop);
  if FA_ExtLst <> Nil then 
    FreeAndNil(FA_ExtLst);
end;

procedure TCT_Scene3D.Assign(AItem: TCT_Scene3D);
begin
  if AItem.FA_Camera <> Nil then begin
    Create_A_Camera;
    FA_Camera.Assign(AItem.FA_Camera);
  end;
  if AItem.FA_LightRig <> Nil then begin
    Create_A_LightRig;
    FA_LightRig.Assign(AItem.FA_LightRig);
  end;
  if AItem.FA_Backdrop <> Nil then begin
    Create_A_Backdrop;
    FA_Backdrop.Assign(AItem.FA_Backdrop);
  end;
  if AItem.FA_ExtLst <> Nil then begin
    Create_A_ExtLst;
    FA_ExtLst.Assign(AItem.FA_ExtLst);
  end;
end;

procedure TCT_Scene3D.CopyTo(AItem: TCT_Scene3D);
begin
end;

function  TCT_Scene3D.Create_A_Camera: TCT_Camera;
begin
  if FA_Camera = Nil then
    FA_Camera := TCT_Camera.Create(FOwner);
  Result := FA_Camera;
end;

function  TCT_Scene3D.Create_A_LightRig: TCT_LightRig;
begin
  if FA_LightRig = Nil then
    FA_LightRig := TCT_LightRig.Create(FOwner);
  Result := FA_LightRig;
end;

function  TCT_Scene3D.Create_A_Backdrop: TCT_Backdrop;
begin
  if FA_Backdrop = Nil then
    FA_Backdrop := TCT_Backdrop.Create(FOwner);
  Result := FA_Backdrop;
end;

function  TCT_Scene3D.Create_A_ExtLst: TCT_OfficeArtExtensionList;
begin
  if FA_ExtLst = Nil then
    FA_ExtLst := TCT_OfficeArtExtensionList.Create(FOwner);
  Result := FA_ExtLst;
end;

{ TEG_Text3D }

function  TEG_Text3D.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FA_Sp3d <> Nil then 
    Inc(ElemsAssigned,FA_Sp3d.CheckAssigned);
  if FA_FlatTx <> Nil then 
    Inc(ElemsAssigned,FA_FlatTx.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TEG_Text3D.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  case AReader.QNameHashA of
    $00000215: begin
      if FA_Sp3d = Nil then 
        FA_Sp3d := TCT_Shape3D.Create(FOwner);
      Result := FA_Sp3d;
    end;
    $0000030E: begin
      if FA_FlatTx = Nil then 
        FA_FlatTx := TCT_FlatText.Create(FOwner);
      Result := FA_FlatTx;
    end;
    else 
    begin
      Result := Nil;
      Exit;
    end
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TEG_Text3D.Write(AWriter: TXpgWriteXML);
begin
  if (FA_Sp3d <> Nil) and FA_Sp3d.Assigned then 
  begin
    FA_Sp3d.WriteAttributes(AWriter);
    if xaElements in FA_Sp3d.FAssigneds then 
    begin
      AWriter.BeginTag('a:sp3d');
      FA_Sp3d.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:sp3d');
  end;
  if (FA_FlatTx <> Nil) and FA_FlatTx.Assigned then 
  begin
    FA_FlatTx.WriteAttributes(AWriter);
    AWriter.SimpleTag('a:flatTx');
  end;
end;

constructor TEG_Text3D.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 2;
  FAttributeCount := 0;
end;

destructor TEG_Text3D.Destroy;
begin
  if FA_Sp3d <> Nil then 
    FA_Sp3d.Free;
  if FA_FlatTx <> Nil then 
    FA_FlatTx.Free;
end;

procedure TEG_Text3D.Clear;
begin
  FAssigneds := [];
  if FA_Sp3d <> Nil then 
    FreeAndNil(FA_Sp3d);
  if FA_FlatTx <> Nil then 
    FreeAndNil(FA_FlatTx);
end;

procedure TEG_Text3D.Assign(AItem: TEG_Text3D);
begin
end;

procedure TEG_Text3D.CopyTo(AItem: TEG_Text3D);
begin
end;

function  TEG_Text3D.Create_A_Sp3d: TCT_Shape3D;
begin
  if FA_Sp3d = Nil then
    FA_Sp3d := TCT_Shape3D.Create(FOwner);
  Result := FA_Sp3d;
end;

function  TEG_Text3D.Create_A_FlatTx: TCT_FlatText;
begin
  if FA_FlatTx = Nil then
    FA_FlatTx := TCT_FlatText.Create(FOwner);
  Result := FA_FlatTx;
end;

{ TEG_TextRun }

function  TEG_TextRun.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FA_R <> Nil then
    Inc(ElemsAssigned,FA_R.CheckAssigned);
  if FA_Br <> Nil then
    Inc(ElemsAssigned,FA_Br.CheckAssigned);
  if FA_Fld <> Nil then
    Inc(ElemsAssigned,FA_Fld.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TEG_TextRun.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  case AReader.QNameHashA of
    $0000010D: begin
      if FA_R = Nil then
        FA_R := TCT_RegularTextRun.Create(FOwner);
      Result := FA_R;
    end;
    $0000016F: begin
      if FA_Br = Nil then
        FA_Br := TCT_TextLineBreak.Create(FOwner);
      Result := FA_Br;
    end;
    $000001D1: begin
      if FA_Fld = Nil then
        FA_Fld := TCT_TextField.Create(FOwner);
      Result := FA_Fld;
    end;
    else
    begin
      Result := Nil;
      Exit;
    end
  end;
  if Result <> Self then
    Result.Assigneds := [xaRead];
end;

procedure TEG_TextRun.Write(AWriter: TXpgWriteXML);
begin
  if (FA_R <> Nil) and FA_R.Assigned then begin
    if xaElements in FA_R.FAssigneds then begin
      AWriter.BeginTag('a:r');
      FA_R.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('a:r');
  end;
  if (FA_Br <> Nil) and FA_Br.Assigned then begin
    if xaElements in FA_Br.FAssigneds then begin
      AWriter.BeginTag('a:br');
      FA_Br.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('a:br');
  end;
  if (FA_Fld <> Nil) and FA_Fld.Assigned then begin
    FA_Fld.WriteAttributes(AWriter);
    if xaElements in FA_Fld.FAssigneds then begin
      AWriter.BeginTag('a:fld');
      FA_Fld.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('a:fld');
  end
end;

constructor TEG_TextRun.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 3;
  FAttributeCount := 0;
end;

destructor TEG_TextRun.Destroy;
begin
  if FA_R <> Nil then
    FA_R.Free;
  if FA_Br <> Nil then
    FA_Br.Free;
  if FA_Fld <> Nil then
    FA_Fld.Free;
end;

procedure TEG_TextRun.Clear;
begin
  FAssigneds := [];
  if FA_R <> Nil then
    FreeAndNil(FA_R);
  if FA_Br <> Nil then
    FreeAndNil(FA_Br);
  if FA_Fld <> Nil then
    FreeAndNil(FA_Fld);
end;

procedure TEG_TextRun.Assign(AItem: TEG_TextRun);
begin
end;

procedure TEG_TextRun.ConvertToBreak;
var
  RPr: TCT_TextCharacterProperties;
begin
  if FA_Br = Nil then begin
    if FA_R <> Nil then begin
      RPr := FA_R.RPr;
      FA_R.FA_RPr := Nil;
    end
    else
      RPr := Nil;
    Clear;
    Create_Br;
    FA_Br.FA_RPr := RPr;
  end;
end;

procedure TEG_TextRun.ConvertToRun;
var
  RPr: TCT_TextCharacterProperties;
begin
  if FA_R = Nil then begin
    if FA_Br <> Nil then begin
      RPr := FA_Br.RPr;
      FA_Br.FA_RPr := Nil;
    end
    else
      RPr := Nil;
    Clear;
    Create_R;
    FA_R.FA_RPr := RPr;
  end;
end;

procedure TEG_TextRun.CopyTo(AItem: TEG_TextRun);
begin
end;

function TEG_TextRun.Create_R: TCT_RegularTextRun;
begin
  if FA_R = Nil then begin
    Clear;
    FA_R := TCT_RegularTextRun.Create(FOwner);
  end;
  Result := FA_R;
end;

function TEG_TextRun.Create_Br: TCT_TextLineBreak;
begin
  if FA_Br = Nil then begin
    Clear;
    FA_Br := TCT_TextLineBreak.Create(FOwner);
  end;
  Result := FA_Br;
end;

function TEG_TextRun.Create_Fldt: TCT_TextField;
begin
  if FA_Fld = Nil then begin
    Clear;
    FA_Fld := TCT_TextField.Create(FOwner);
  end;
  Result := FA_Fld;
end;

{ TCT_Connection }

function  TCT_Connection.CheckAssigned: integer;
begin
  FAssigneds := [xaAttributes];
  Result := 1;
end;

procedure TCT_Connection.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_Connection.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('id',XmlIntToStr(FId));
  AWriter.AddAttribute('idx',XmlIntToStr(FIdx));
end;

procedure TCT_Connection.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $000000CD: FId := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000145: FIdx := XmlStrToIntDef(AAttributes.Values[i],0);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_Connection.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 2;
  FId := 2147483632;
  FIdx := 2147483632;
end;

destructor TCT_Connection.Destroy;
begin
end;

procedure TCT_Connection.Clear;
begin
  FAssigneds := [];
  FId := 2147483632;
  FIdx := 2147483632;
end;

procedure TCT_Connection.Assign(AItem: TCT_Connection);
begin
end;

procedure TCT_Connection.CopyTo(AItem: TCT_Connection);
begin
end;

{ TCT_Transform2D }

function  TCT_Transform2D.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FRot <> 0 then 
    Inc(AttrsAssigned);
  if FFlipH <> False then 
    Inc(AttrsAssigned);
  if FFlipV <> False then 
    Inc(AttrsAssigned);
  if FA_Off <> Nil then 
    Inc(ElemsAssigned,FA_Off.CheckAssigned);
  if FA_Ext <> Nil then 
    Inc(ElemsAssigned,FA_Ext.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_Transform2D.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $000001D6: begin
      if FA_Off = Nil then 
        FA_Off := TCT_Point2D.Create(FOwner);
      Result := FA_Off;
    end;
    $000001EC: begin
      if FA_Ext = Nil then 
        FA_Ext := TCT_PositiveSize2D.Create(FOwner);
      Result := FA_Ext;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_Transform2D.Write(AWriter: TXpgWriteXML);
begin
  if (FA_Off <> Nil) and FA_Off.Assigned then 
  begin
    FA_Off.WriteAttributes(AWriter);
    AWriter.SimpleTag('a:off');
  end;
  if (FA_Ext <> Nil) and FA_Ext.Assigned then 
  begin
    FA_Ext.WriteAttributes(AWriter);
    AWriter.SimpleTag('a:ext');
  end;
end;

procedure TCT_Transform2D.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FRot <> 0 then 
    AWriter.AddAttribute('rot',XmlIntToStr(FRot));
  if FFlipH <> False then 
    AWriter.AddAttribute('flipH',XmlBoolToStr(FFlipH));
  if FFlipV <> False then 
    AWriter.AddAttribute('flipV',XmlBoolToStr(FFlipV));
end;

procedure TCT_Transform2D.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $00000155: FRot := XmlStrToIntDef(AAttributes.Values[i],0);
      $000001F3: FFlipH := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000201: FFlipV := XmlStrToBoolDef(AAttributes.Values[i],False);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_Transform2D.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 2;
  FAttributeCount := 3;
  FRot := 0;
  FFlipH := False;
  FFlipV := False;
end;

destructor TCT_Transform2D.Destroy;
begin
  if FA_Off <> Nil then 
    FA_Off.Free;
  if FA_Ext <> Nil then 
    FA_Ext.Free;
end;

procedure TCT_Transform2D.Clear;
begin
  FAssigneds := [];
  if FA_Off <> Nil then 
    FreeAndNil(FA_Off);
  if FA_Ext <> Nil then 
    FreeAndNil(FA_Ext);
  FRot := 0;
  FFlipH := False;
  FFlipV := False;
end;

procedure TCT_Transform2D.Assign(AItem: TCT_Transform2D);
begin
  FRot := AItem.FRot;
  FFlipH := AItem.FFlipH;
  FFlipV := AItem.FFlipV;
  if AItem.FA_Off <> Nil then begin
    Create_A_Off;
    A_Off.Assign(AItem.FA_Off);
  end;
  if AItem.FA_Ext <> Nil then begin
    Create_A_Ext;
    A_Ext.Assign(AItem.FA_Ext);
  end;
end;

procedure TCT_Transform2D.CopyTo(AItem: TCT_Transform2D);
begin
end;

function  TCT_Transform2D.Create_A_Off: TCT_Point2D;
begin
  if FA_Off = Nil then
    FA_Off := TCT_Point2D.Create(FOwner);
  Result := FA_Off;
end;

function  TCT_Transform2D.Create_A_Ext: TCT_PositiveSize2D;
begin
  if FA_Ext = Nil then
    FA_Ext := TCT_PositiveSize2D.Create(FOwner);
  Result := FA_Ext;
end;

{ TEG_Geometry }

function  TEG_Geometry.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FA_CustGeom <> Nil then 
    Inc(ElemsAssigned,FA_CustGeom.CheckAssigned);
  if FA_PrstGeom <> Nil then 
    Inc(ElemsAssigned,FA_PrstGeom.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TEG_Geometry.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  case AReader.QNameHashA of
    $000003E2: begin
      if FA_CustGeom = Nil then 
        FA_CustGeom := TCT_CustomGeometry2D.Create(FOwner);
      Result := FA_CustGeom;
    end;
    $000003EC: begin
      if FA_PrstGeom = Nil then 
        FA_PrstGeom := TCT_PresetGeometry2D.Create(FOwner);
      Result := FA_PrstGeom;
    end;
    else 
    begin
      Result := Nil;
      Exit;
    end
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TEG_Geometry.Write(AWriter: TXpgWriteXML);
begin
  if (FA_CustGeom <> Nil) and FA_CustGeom.Assigned then 
    if xaElements in FA_CustGeom.FAssigneds then 
    begin
      AWriter.BeginTag('a:custGeom');
      FA_CustGeom.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:custGeom');
  if (FA_PrstGeom <> Nil) and FA_PrstGeom.Assigned then 
  begin
    FA_PrstGeom.WriteAttributes(AWriter);
    if xaElements in FA_PrstGeom.FAssigneds then 
    begin
      AWriter.BeginTag('a:prstGeom');
      FA_PrstGeom.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:prstGeom');
  end;
end;

constructor TEG_Geometry.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 2;
  FAttributeCount := 0;
end;

destructor TEG_Geometry.Destroy;
begin
  if FA_CustGeom <> Nil then
    FA_CustGeom.Free;
  if FA_PrstGeom <> Nil then
    FA_PrstGeom.Free;
end;

procedure TEG_Geometry.Clear;
begin
  FAssigneds := [];
  if FA_CustGeom <> Nil then
    FreeAndNil(FA_CustGeom);
  if FA_PrstGeom <> Nil then
    FreeAndNil(FA_PrstGeom);
end;

procedure TEG_Geometry.Assign(AItem: TEG_Geometry);
begin
  if AItem.FA_CustGeom <> Nil then begin
    Create_CustGeom;
    FA_CustGeom.Assign(AItem.FA_CustGeom);
  end;
  if AItem.FA_PrstGeom <> Nil then begin
    Create_PrstGeom;
    FA_PrstGeom.Assign(AItem.FA_PrstGeom);
  end;
end;

procedure TEG_Geometry.CopyTo(AItem: TEG_Geometry);
begin
end;

function  TEG_Geometry.Create_CustGeom: TCT_CustomGeometry2D;
begin
  if FA_CustGeom = Nil then begin
    Clear;
    FA_CustGeom := TCT_CustomGeometry2D.Create(FOwner);
  end;
  Result := FA_CustGeom;
end;

function  TEG_Geometry.Create_PrstGeom: TCT_PresetGeometry2D;
begin
  if FA_PrstGeom = Nil then begin
    Clear;
    FA_PrstGeom := TCT_PresetGeometry2D.Create(FOwner);
  end;
  Result := FA_PrstGeom;
end;

{ TCT_StyleMatrixReference }

function  TCT_StyleMatrixReference.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [xaAttributes];
  Inc(ElemsAssigned,FA_EG_ColorChoice.CheckAssigned);
  Result := 1;
  if ElemsAssigned > 0 then
    FAssigneds := FAssigneds + [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_StyleMatrixReference.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := FA_EG_ColorChoice.HandleElement(AReader);
  if Result = Nil then 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_StyleMatrixReference.Write(AWriter: TXpgWriteXML);
begin
  FA_EG_ColorChoice.Write(AWriter);
end;

procedure TCT_StyleMatrixReference.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('idx',XmlIntToStr(FIdx));
end;

procedure TCT_StyleMatrixReference.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'idx' then 
    FIdx := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_StyleMatrixReference.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 1;
  FA_EG_ColorChoice := TEG_ColorChoice.Create(FOwner);
  FIdx := 2147483632;
end;

destructor TCT_StyleMatrixReference.Destroy;
begin
  FA_EG_ColorChoice.Free;
end;

procedure TCT_StyleMatrixReference.Clear;
begin
  FAssigneds := [];
  FA_EG_ColorChoice.Clear;
  FIdx := 2147483632;
end;

procedure TCT_StyleMatrixReference.Assign(AItem: TCT_StyleMatrixReference);
begin
end;

procedure TCT_StyleMatrixReference.CopyTo(AItem: TCT_StyleMatrixReference);
begin
end;

{ TCT_FontReference }

function  TCT_FontReference.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [xaAttributes];
  Inc(ElemsAssigned,FA_EG_ColorChoice.CheckAssigned);
  Result := 1;
  if ElemsAssigned > 0 then
    FAssigneds := FAssigneds + [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_FontReference.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := FA_EG_ColorChoice.HandleElement(AReader);
  if Result = Nil then 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_FontReference.Write(AWriter: TXpgWriteXML);
begin
  FA_EG_ColorChoice.Write(AWriter);
end;

procedure TCT_FontReference.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FIdx <> TST_FontCollectionIndex(XPG_UNKNOWN_ENUM) then
    AWriter.AddAttribute('idx',StrTST_FontCollectionIndex[Integer(FIdx)]);
end;

procedure TCT_FontReference.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'idx' then 
    FIdx := TST_FontCollectionIndex(StrToEnum('stfci' + AAttributes.Values[0]))
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_FontReference.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 1;
  FA_EG_ColorChoice := TEG_ColorChoice.Create(FOwner);
  FIdx := TST_FontCollectionIndex(XPG_UNKNOWN_ENUM);
end;

destructor TCT_FontReference.Destroy;
begin
  FA_EG_ColorChoice.Free;
end;

procedure TCT_FontReference.Clear;
begin
  FAssigneds := [];
  FA_EG_ColorChoice.Clear;
  FIdx := TST_FontCollectionIndex(XPG_UNKNOWN_ENUM);
end;

procedure TCT_FontReference.Assign(AItem: TCT_FontReference);
begin
end;

procedure TCT_FontReference.CopyTo(AItem: TCT_FontReference);
begin
end;

{ TCT_TextBodyProperties }

function  TCT_TextBodyProperties.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FRot <> 2147483632 then 
    Inc(AttrsAssigned);
  if Byte(FSpcFirstLastPara) <> 2 then 
    Inc(AttrsAssigned);
  if Integer(FVertOverflow) <> XPG_UNKNOWN_ENUM then
    Inc(AttrsAssigned);
  if Integer(FHorzOverflow) <> XPG_UNKNOWN_ENUM then 
    Inc(AttrsAssigned);
  if Integer(FVert) <> XPG_UNKNOWN_ENUM then 
    Inc(AttrsAssigned);
  if Integer(FWrap) <> XPG_UNKNOWN_ENUM then 
    Inc(AttrsAssigned);
  if FLIns <> 2147483632 then 
    Inc(AttrsAssigned);
  if FTIns <> 2147483632 then 
    Inc(AttrsAssigned);
  if FRIns <> 2147483632 then 
    Inc(AttrsAssigned);
  if FBIns <> 2147483632 then 
    Inc(AttrsAssigned);
  if FNumCol <> 2147483632 then 
    Inc(AttrsAssigned);
  if FSpcCol <> 2147483632 then 
    Inc(AttrsAssigned);
  if Byte(FRtlCol) <> 2 then 
    Inc(AttrsAssigned);
  if Byte(FFromWordArt) <> 2 then 
    Inc(AttrsAssigned);
  if Integer(FAnchor) <> XPG_UNKNOWN_ENUM then 
    Inc(AttrsAssigned);
  if Byte(FAnchorCtr) <> 2 then 
    Inc(AttrsAssigned);
  if Byte(FForceAA) <> 2 then 
    Inc(AttrsAssigned);
  if FUpright <> False then 
    Inc(AttrsAssigned);
  if Byte(FCompatLnSpc) <> 2 then 
    Inc(AttrsAssigned);
  if FA_PrstTxWarp <> Nil then 
    Inc(ElemsAssigned,FA_PrstTxWarp.CheckAssigned);
  Inc(ElemsAssigned,FA_EG_TextAutofit.CheckAssigned);
  if FA_Scene3d <> Nil then 
    Inc(ElemsAssigned,FA_Scene3d.CheckAssigned);
  Inc(ElemsAssigned,FA_EG_Text3D.CheckAssigned);
  if FA_ExtLst <> Nil then 
    Inc(ElemsAssigned,FA_ExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_TextBodyProperties.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  case AReader.QNameHashA of
    $000004CA: begin
      if FA_PrstTxWarp = Nil then 
        FA_PrstTxWarp := TCT_PresetTextShape.Create(FOwner);
      Result := FA_PrstTxWarp;
    end;
    $00000340: begin
      if FA_Scene3d = Nil then 
        FA_Scene3d := TCT_Scene3D.Create(FOwner);
      Result := FA_Scene3d;
    end;
    $0000031F: begin
      if FA_ExtLst = Nil then 
        FA_ExtLst := TCT_OfficeArtExtensionList.Create(FOwner);
      Result := FA_ExtLst;
    end;
    else 
    begin
      Result := FA_EG_TextAutofit.HandleElement(AReader);
      if Result = Nil then 
      begin
        Result := FA_EG_Text3D.HandleElement(AReader);
        if Result = Nil then 
          FOwner.Errors.Error(xemUnknownElement,AReader.QName);
      end;
    end
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_TextBodyProperties.Write(AWriter: TXpgWriteXML);
begin
  if (FA_PrstTxWarp <> Nil) and FA_PrstTxWarp.Assigned then 
  begin
    FA_PrstTxWarp.WriteAttributes(AWriter);
    if xaElements in FA_PrstTxWarp.FAssigneds then 
    begin
      AWriter.BeginTag('a:prstTxWarp');
      FA_PrstTxWarp.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:prstTxWarp');
  end;
  FA_EG_TextAutofit.Write(AWriter);
  if (FA_Scene3d <> Nil) and FA_Scene3d.Assigned then 
    if xaElements in FA_Scene3d.FAssigneds then 
    begin
      AWriter.BeginTag('a:scene3d');
      FA_Scene3d.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:scene3d');
  FA_EG_Text3D.Write(AWriter);
  if (FA_ExtLst <> Nil) and FA_ExtLst.Assigned then 
    if xaElements in FA_ExtLst.FAssigneds then 
    begin
      AWriter.BeginTag('a:extLst');
      FA_ExtLst.Write(AWriter);
      AWriter.EndTag;
    end;
end;

procedure TCT_TextBodyProperties.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FRot <> 2147483632 then 
    AWriter.AddAttribute('rot',XmlIntToStr(FRot));
  if Byte(FSpcFirstLastPara) <> 2 then 
    AWriter.AddAttribute('spcFirstLastPara',XmlBoolToStr(FSpcFirstLastPara));
  if Integer(FVertOverflow) <> XPG_UNKNOWN_ENUM then 
    AWriter.AddAttribute('vertOverflow',StrTST_TextVertOverflowType[Integer(FVertOverflow)]);
  if Integer(FHorzOverflow) <> XPG_UNKNOWN_ENUM then 
    AWriter.AddAttribute('horzOverflow',StrTST_TextHorzOverflowType[Integer(FHorzOverflow)]);
  if Integer(FVert) <> XPG_UNKNOWN_ENUM then 
    AWriter.AddAttribute('vert',StrTST_TextVerticalType[Integer(FVert)]);
  if Integer(FWrap) <> XPG_UNKNOWN_ENUM then 
    AWriter.AddAttribute('wrap',StrTST_TextWrappingType[Integer(FWrap)]);
  if FLIns <> 2147483632 then 
    AWriter.AddAttribute('lIns',XmlIntToStr(FLIns));
  if FTIns <> 2147483632 then 
    AWriter.AddAttribute('tIns',XmlIntToStr(FTIns));
  if FRIns <> 2147483632 then 
    AWriter.AddAttribute('rIns',XmlIntToStr(FRIns));
  if FBIns <> 2147483632 then 
    AWriter.AddAttribute('bIns',XmlIntToStr(FBIns));
  if FNumCol <> 2147483632 then 
    AWriter.AddAttribute('numCol',XmlIntToStr(FNumCol));
  if FSpcCol <> 2147483632 then 
    AWriter.AddAttribute('spcCol',XmlIntToStr(FSpcCol));
  if Byte(FRtlCol) <> 2 then 
    AWriter.AddAttribute('rtlCol',XmlBoolToStr(FRtlCol));
  if Byte(FFromWordArt) <> 2 then 
    AWriter.AddAttribute('fromWordArt',XmlBoolToStr(FFromWordArt));
  if Integer(FAnchor) <> XPG_UNKNOWN_ENUM then 
    AWriter.AddAttribute('anchor',StrTST_TextAnchoringType[Integer(FAnchor)]);
  if Byte(FAnchorCtr) <> 2 then 
    AWriter.AddAttribute('anchorCtr',XmlBoolToStr(FAnchorCtr));
  if Byte(FForceAA) <> 2 then 
    AWriter.AddAttribute('forceAA',XmlBoolToStr(FForceAA));
  if FUpright <> False then 
    AWriter.AddAttribute('upright',XmlBoolToStr(FUpright));
  if Byte(FCompatLnSpc) <> 2 then 
    AWriter.AddAttribute('compatLnSpc',XmlBoolToStr(FCompatLnSpc));
end;

procedure TCT_TextBodyProperties.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $00000155: FRot := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000666: FSpcFirstLastPara := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000515: FVertOverflow := TST_TextVertOverflowType(StrToEnum('sttvot' + AAttributes.Values[i]));
      $00000517: FHorzOverflow := TST_TextHorzOverflowType(StrToEnum('stthot' + AAttributes.Values[i]));
      $000001C1: FVert := TST_TextVerticalType(StrToEnum('sttvt' + AAttributes.Values[i]));
      $000001BA: FWrap := TST_TextWrappingType(StrToEnum('sttwt' + AAttributes.Values[i]));
      $00000196: FLIns := XmlStrToIntDef(AAttributes.Values[i],0);
      $0000019E: FTIns := XmlStrToIntDef(AAttributes.Values[i],0);
      $0000019C: FRIns := XmlStrToIntDef(AAttributes.Values[i],0);
      $0000018C: FBIns := XmlStrToIntDef(AAttributes.Values[i],0);
      $0000026E: FNumCol := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000264: FSpcCol := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000270: FRtlCol := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000477: FFromWordArt := XmlStrToBoolDef(AAttributes.Values[i],False);
      $0000027B: FAnchor := TST_TextAnchoringType(StrToEnum('sttat' + AAttributes.Values[i]));
      $000003A4: FAnchorCtr := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000291: FForceAA := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000303: FUpright := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000464: FCompatLnSpc := XmlStrToBoolDef(AAttributes.Values[i],False);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_TextBodyProperties.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 5;
  FAttributeCount := 19;
  FA_EG_TextAutofit := TEG_TextAutofit.Create(FOwner);
  FA_EG_Text3D := TEG_Text3D.Create(FOwner);
  FRot := 2147483632;
  Byte(FSpcFirstLastPara) := 2;
  FVertOverflow := TST_TextVertOverflowType(XPG_UNKNOWN_ENUM);
  FHorzOverflow := TST_TextHorzOverflowType(XPG_UNKNOWN_ENUM);
  FVert := TST_TextVerticalType(XPG_UNKNOWN_ENUM);
  FWrap := TST_TextWrappingType(XPG_UNKNOWN_ENUM);
  FLIns := 2147483632;
  FTIns := 2147483632;
  FRIns := 2147483632;
  FBIns := 2147483632;
  FNumCol := 2147483632;
  FSpcCol := 2147483632;
  Byte(FRtlCol) := 2;
  Byte(FFromWordArt) := 2;
  FAnchor := TST_TextAnchoringType(XPG_UNKNOWN_ENUM);
  Byte(FAnchorCtr) := 2;
  Byte(FForceAA) := 2;
  FUpright := False;
  Byte(FCompatLnSpc) := 2;
end;

destructor TCT_TextBodyProperties.Destroy;
begin
  if FA_PrstTxWarp <> Nil then
    FA_PrstTxWarp.Free;
  FA_EG_TextAutofit.Free;
  if FA_Scene3d <> Nil then
    FA_Scene3d.Free;
  FA_EG_Text3D.Free;
  if FA_ExtLst <> Nil then
    FA_ExtLst.Free;
end;

procedure TCT_TextBodyProperties.Clear;
begin
  FAssigneds := [];
  if FA_PrstTxWarp <> Nil then
    FreeAndNil(FA_PrstTxWarp);
  FA_EG_TextAutofit.Clear;
  if FA_Scene3d <> Nil then
    FreeAndNil(FA_Scene3d);
  FA_EG_Text3D.Clear;
  if FA_ExtLst <> Nil then
    FreeAndNil(FA_ExtLst);
  FRot := 2147483632;
  Byte(FSpcFirstLastPara) := 2;
  FVertOverflow := TST_TextVertOverflowType(XPG_UNKNOWN_ENUM);
  FHorzOverflow := TST_TextHorzOverflowType(XPG_UNKNOWN_ENUM);
  FVert := TST_TextVerticalType(XPG_UNKNOWN_ENUM);
  FWrap := TST_TextWrappingType(XPG_UNKNOWN_ENUM);
  FLIns := 2147483632;
  FTIns := 2147483632;
  FRIns := 2147483632;
  FBIns := 2147483632;
  FNumCol := 2147483632;
  FSpcCol := 2147483632;
  Byte(FRtlCol) := 2;
  Byte(FFromWordArt) := 2;
  FAnchor := TST_TextAnchoringType(XPG_UNKNOWN_ENUM);
  Byte(FAnchorCtr) := 2;
  Byte(FForceAA) := 2;
  FUpright := False;
  Byte(FCompatLnSpc) := 2;
end;

procedure TCT_TextBodyProperties.Assign(AItem: TCT_TextBodyProperties);
begin
end;

procedure TCT_TextBodyProperties.CopyTo(AItem: TCT_TextBodyProperties);
begin
end;

function  TCT_TextBodyProperties.Create_A_PrstTxWarp: TCT_PresetTextShape;
begin
  if FA_PrstTxWarp = Nil then
    FA_PrstTxWarp := TCT_PresetTextShape.Create(FOwner);
  Result := FA_PrstTxWarp;
end;

function  TCT_TextBodyProperties.Create_A_Scene3d: TCT_Scene3D;
begin
  if FA_Scene3d = Nil then
    FA_Scene3d := TCT_Scene3D.Create(FOwner);
  Result := FA_Scene3d;
end;

function  TCT_TextBodyProperties.Create_A_ExtLst: TCT_OfficeArtExtensionList;
begin
  if FA_ExtLst = Nil then
    FA_ExtLst := TCT_OfficeArtExtensionList.Create(FOwner);
  Result := FA_ExtLst;
end;

{ TCT_TextListStyle }

function  TCT_TextListStyle.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FA_DefPPr <> Nil then 
    Inc(ElemsAssigned,FA_DefPPr.CheckAssigned);
  if FA_Lvl1pPr <> Nil then 
    Inc(ElemsAssigned,FA_Lvl1pPr.CheckAssigned);
  if FA_Lvl2pPr <> Nil then 
    Inc(ElemsAssigned,FA_Lvl2pPr.CheckAssigned);
  if FA_Lvl3pPr <> Nil then 
    Inc(ElemsAssigned,FA_Lvl3pPr.CheckAssigned);
  if FA_Lvl4pPr <> Nil then 
    Inc(ElemsAssigned,FA_Lvl4pPr.CheckAssigned);
  if FA_Lvl5pPr <> Nil then 
    Inc(ElemsAssigned,FA_Lvl5pPr.CheckAssigned);
  if FA_Lvl6pPr <> Nil then 
    Inc(ElemsAssigned,FA_Lvl6pPr.CheckAssigned);
  if FA_Lvl7pPr <> Nil then 
    Inc(ElemsAssigned,FA_Lvl7pPr.CheckAssigned);
  if FA_Lvl8pPr <> Nil then 
    Inc(ElemsAssigned,FA_Lvl8pPr.CheckAssigned);
  if FA_Lvl9pPr <> Nil then 
    Inc(ElemsAssigned,FA_Lvl9pPr.CheckAssigned);
  if FA_ExtLst <> Nil then 
    Inc(ElemsAssigned,FA_ExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_TextListStyle.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $000002DC: begin
      if FA_DefPPr = Nil then 
        FA_DefPPr := TCT_TextParagraphProperties.Create(FOwner);
      Result := FA_DefPPr;
    end;
    $0000034C: begin
      if FA_Lvl1pPr = Nil then 
        FA_Lvl1pPr := TCT_TextParagraphProperties.Create(FOwner);
      Result := FA_Lvl1pPr;
    end;
    $0000034D: begin
      if FA_Lvl2pPr = Nil then 
        FA_Lvl2pPr := TCT_TextParagraphProperties.Create(FOwner);
      Result := FA_Lvl2pPr;
    end;
    $0000034E: begin
      if FA_Lvl3pPr = Nil then 
        FA_Lvl3pPr := TCT_TextParagraphProperties.Create(FOwner);
      Result := FA_Lvl3pPr;
    end;
    $0000034F: begin
      if FA_Lvl4pPr = Nil then 
        FA_Lvl4pPr := TCT_TextParagraphProperties.Create(FOwner);
      Result := FA_Lvl4pPr;
    end;
    $00000350: begin
      if FA_Lvl5pPr = Nil then 
        FA_Lvl5pPr := TCT_TextParagraphProperties.Create(FOwner);
      Result := FA_Lvl5pPr;
    end;
    $00000351: begin
      if FA_Lvl6pPr = Nil then 
        FA_Lvl6pPr := TCT_TextParagraphProperties.Create(FOwner);
      Result := FA_Lvl6pPr;
    end;
    $00000352: begin
      if FA_Lvl7pPr = Nil then 
        FA_Lvl7pPr := TCT_TextParagraphProperties.Create(FOwner);
      Result := FA_Lvl7pPr;
    end;
    $00000353: begin
      if FA_Lvl8pPr = Nil then 
        FA_Lvl8pPr := TCT_TextParagraphProperties.Create(FOwner);
      Result := FA_Lvl8pPr;
    end;
    $00000354: begin
      if FA_Lvl9pPr = Nil then 
        FA_Lvl9pPr := TCT_TextParagraphProperties.Create(FOwner);
      Result := FA_Lvl9pPr;
    end;
    $0000031F: begin
      if FA_ExtLst = Nil then 
        FA_ExtLst := TCT_OfficeArtExtensionList.Create(FOwner);
      Result := FA_ExtLst;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_TextListStyle.Write(AWriter: TXpgWriteXML);
begin
  if (FA_DefPPr <> Nil) and FA_DefPPr.Assigned then 
  begin
    FA_DefPPr.WriteAttributes(AWriter);
    if xaElements in FA_DefPPr.FAssigneds then 
    begin
      AWriter.BeginTag('a:defPPr');
      FA_DefPPr.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:defPPr');
  end;
  if (FA_Lvl1pPr <> Nil) and FA_Lvl1pPr.Assigned then 
  begin
    FA_Lvl1pPr.WriteAttributes(AWriter);
    if xaElements in FA_Lvl1pPr.FAssigneds then 
    begin
      AWriter.BeginTag('a:lvl1pPr');
      FA_Lvl1pPr.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:lvl1pPr');
  end;
  if (FA_Lvl2pPr <> Nil) and FA_Lvl2pPr.Assigned then 
  begin
    FA_Lvl2pPr.WriteAttributes(AWriter);
    if xaElements in FA_Lvl2pPr.FAssigneds then 
    begin
      AWriter.BeginTag('a:lvl2pPr');
      FA_Lvl2pPr.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:lvl2pPr');
  end;
  if (FA_Lvl3pPr <> Nil) and FA_Lvl3pPr.Assigned then 
  begin
    FA_Lvl3pPr.WriteAttributes(AWriter);
    if xaElements in FA_Lvl3pPr.FAssigneds then 
    begin
      AWriter.BeginTag('a:lvl3pPr');
      FA_Lvl3pPr.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:lvl3pPr');
  end;
  if (FA_Lvl4pPr <> Nil) and FA_Lvl4pPr.Assigned then 
  begin
    FA_Lvl4pPr.WriteAttributes(AWriter);
    if xaElements in FA_Lvl4pPr.FAssigneds then 
    begin
      AWriter.BeginTag('a:lvl4pPr');
      FA_Lvl4pPr.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:lvl4pPr');
  end;
  if (FA_Lvl5pPr <> Nil) and FA_Lvl5pPr.Assigned then 
  begin
    FA_Lvl5pPr.WriteAttributes(AWriter);
    if xaElements in FA_Lvl5pPr.FAssigneds then 
    begin
      AWriter.BeginTag('a:lvl5pPr');
      FA_Lvl5pPr.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:lvl5pPr');
  end;
  if (FA_Lvl6pPr <> Nil) and FA_Lvl6pPr.Assigned then 
  begin
    FA_Lvl6pPr.WriteAttributes(AWriter);
    if xaElements in FA_Lvl6pPr.FAssigneds then 
    begin
      AWriter.BeginTag('a:lvl6pPr');
      FA_Lvl6pPr.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:lvl6pPr');
  end;
  if (FA_Lvl7pPr <> Nil) and FA_Lvl7pPr.Assigned then 
  begin
    FA_Lvl7pPr.WriteAttributes(AWriter);
    if xaElements in FA_Lvl7pPr.FAssigneds then 
    begin
      AWriter.BeginTag('a:lvl7pPr');
      FA_Lvl7pPr.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:lvl7pPr');
  end;
  if (FA_Lvl8pPr <> Nil) and FA_Lvl8pPr.Assigned then 
  begin
    FA_Lvl8pPr.WriteAttributes(AWriter);
    if xaElements in FA_Lvl8pPr.FAssigneds then 
    begin
      AWriter.BeginTag('a:lvl8pPr');
      FA_Lvl8pPr.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:lvl8pPr');
  end;
  if (FA_Lvl9pPr <> Nil) and FA_Lvl9pPr.Assigned then 
  begin
    FA_Lvl9pPr.WriteAttributes(AWriter);
    if xaElements in FA_Lvl9pPr.FAssigneds then 
    begin
      AWriter.BeginTag('a:lvl9pPr');
      FA_Lvl9pPr.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:lvl9pPr');
  end;
  if (FA_ExtLst <> Nil) and FA_ExtLst.Assigned then 
    if xaElements in FA_ExtLst.FAssigneds then 
    begin
      AWriter.BeginTag('a:extLst');
      FA_ExtLst.Write(AWriter);
      AWriter.EndTag;
    end;
end;

constructor TCT_TextListStyle.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 11;
  FAttributeCount := 0;
end;

destructor TCT_TextListStyle.Destroy;
begin
  if FA_DefPPr <> Nil then 
    FA_DefPPr.Free;
  if FA_Lvl1pPr <> Nil then 
    FA_Lvl1pPr.Free;
  if FA_Lvl2pPr <> Nil then 
    FA_Lvl2pPr.Free;
  if FA_Lvl3pPr <> Nil then 
    FA_Lvl3pPr.Free;
  if FA_Lvl4pPr <> Nil then 
    FA_Lvl4pPr.Free;
  if FA_Lvl5pPr <> Nil then 
    FA_Lvl5pPr.Free;
  if FA_Lvl6pPr <> Nil then 
    FA_Lvl6pPr.Free;
  if FA_Lvl7pPr <> Nil then 
    FA_Lvl7pPr.Free;
  if FA_Lvl8pPr <> Nil then 
    FA_Lvl8pPr.Free;
  if FA_Lvl9pPr <> Nil then 
    FA_Lvl9pPr.Free;
  if FA_ExtLst <> Nil then 
    FA_ExtLst.Free;
end;

procedure TCT_TextListStyle.Clear;
begin
  FAssigneds := [];
  if FA_DefPPr <> Nil then 
    FreeAndNil(FA_DefPPr);
  if FA_Lvl1pPr <> Nil then 
    FreeAndNil(FA_Lvl1pPr);
  if FA_Lvl2pPr <> Nil then 
    FreeAndNil(FA_Lvl2pPr);
  if FA_Lvl3pPr <> Nil then 
    FreeAndNil(FA_Lvl3pPr);
  if FA_Lvl4pPr <> Nil then 
    FreeAndNil(FA_Lvl4pPr);
  if FA_Lvl5pPr <> Nil then 
    FreeAndNil(FA_Lvl5pPr);
  if FA_Lvl6pPr <> Nil then 
    FreeAndNil(FA_Lvl6pPr);
  if FA_Lvl7pPr <> Nil then 
    FreeAndNil(FA_Lvl7pPr);
  if FA_Lvl8pPr <> Nil then 
    FreeAndNil(FA_Lvl8pPr);
  if FA_Lvl9pPr <> Nil then 
    FreeAndNil(FA_Lvl9pPr);
  if FA_ExtLst <> Nil then 
    FreeAndNil(FA_ExtLst);
end;

procedure TCT_TextListStyle.Assign(AItem: TCT_TextListStyle);
begin
end;

procedure TCT_TextListStyle.CopyTo(AItem: TCT_TextListStyle);
begin
end;

function  TCT_TextListStyle.Create_A_DefPPr: TCT_TextParagraphProperties;
begin
  if FA_DefPPr = Nil then
    FA_DefPPr := TCT_TextParagraphProperties.Create(FOwner);
  Result := FA_DefPPr;
end;

function  TCT_TextListStyle.Create_A_Lvl1pPr: TCT_TextParagraphProperties;
begin
  if FA_Lvl1pPr = Nil then
    FA_Lvl1pPr := TCT_TextParagraphProperties.Create(FOwner);
  Result := FA_Lvl1pPr;
end;

function  TCT_TextListStyle.Create_A_Lvl2pPr: TCT_TextParagraphProperties;
begin
  if FA_Lvl2pPr = Nil then
    FA_Lvl2pPr := TCT_TextParagraphProperties.Create(FOwner);
  Result := FA_Lvl2pPr;
end;

function  TCT_TextListStyle.Create_A_Lvl3pPr: TCT_TextParagraphProperties;
begin
  if FA_Lvl3pPr = Nil then
    FA_Lvl3pPr := TCT_TextParagraphProperties.Create(FOwner);
  Result := FA_Lvl3pPr;
end;

function  TCT_TextListStyle.Create_A_Lvl4pPr: TCT_TextParagraphProperties;
begin
  if FA_Lvl4pPr = Nil then
    FA_Lvl4pPr := TCT_TextParagraphProperties.Create(FOwner);
  Result := FA_Lvl4pPr;
end;

function  TCT_TextListStyle.Create_A_Lvl5pPr: TCT_TextParagraphProperties;
begin
  if FA_Lvl5pPr = Nil then
    FA_Lvl5pPr := TCT_TextParagraphProperties.Create(FOwner);
  Result := FA_Lvl5pPr;
end;

function  TCT_TextListStyle.Create_A_Lvl6pPr: TCT_TextParagraphProperties;
begin
  if FA_Lvl6pPr = Nil then
    FA_Lvl6pPr := TCT_TextParagraphProperties.Create(FOwner);
  Result := FA_Lvl6pPr;
end;

function  TCT_TextListStyle.Create_A_Lvl7pPr: TCT_TextParagraphProperties;
begin
  if FA_Lvl7pPr = Nil then
    FA_Lvl7pPr := TCT_TextParagraphProperties.Create(FOwner);
  Result := FA_Lvl7pPr;
end;

function  TCT_TextListStyle.Create_A_Lvl8pPr: TCT_TextParagraphProperties;
begin
  if FA_Lvl8pPr = Nil then
    FA_Lvl8pPr := TCT_TextParagraphProperties.Create(FOwner);
  Result := FA_Lvl8pPr;
end;

function  TCT_TextListStyle.Create_A_Lvl9pPr: TCT_TextParagraphProperties;
begin
  if FA_Lvl9pPr = Nil then
    FA_Lvl9pPr := TCT_TextParagraphProperties.Create(FOwner);
  Result := FA_Lvl9pPr;
end;

function  TCT_TextListStyle.Create_A_ExtLst: TCT_OfficeArtExtensionList;
begin
  if FA_ExtLst = Nil then
    FA_ExtLst := TCT_OfficeArtExtensionList.Create(FOwner);
  Result := FA_ExtLst;
end;

{ TCT_TextParagraph }

function  TCT_TextParagraph.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FA_PPr <> Nil then 
    Inc(ElemsAssigned,FA_PPr.CheckAssigned);
  Inc(ElemsAssigned,FA_EG_TextRuns.CheckAssigned);
  if FA_EndParaRPr <> Nil then 
    Inc(ElemsAssigned,FA_EndParaRPr.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_TextParagraph.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  case AReader.QNameHashA of
    $000001CD: begin
      if FA_PPr = Nil then 
        FA_PPr := TCT_TextParagraphProperties.Create(FOwner);
      Result := FA_PPr;
    end;
    $0000046A: begin
      if FA_EndParaRPr = Nil then 
        FA_EndParaRPr := TCT_TextCharacterProperties.Create(FOwner);
      Result := FA_EndParaRPr;
    end;
    else 
    begin
      Result := FA_EG_TextRuns.Add.HandleElement(AReader);
      if Result = Nil then 
        FOwner.Errors.Error(xemUnknownElement,AReader.QName);
    end
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_TextParagraph.Write(AWriter: TXpgWriteXML);
begin
  if (FA_PPr <> Nil) and FA_PPr.Assigned then 
  begin
    FA_PPr.WriteAttributes(AWriter);
    if xaElements in FA_PPr.FAssigneds then 
    begin
      AWriter.BeginTag('a:pPr');
      FA_PPr.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:pPr');
  end;
  FA_EG_TextRuns.Write(AWriter);
  if (FA_EndParaRPr <> Nil) and FA_EndParaRPr.Assigned then 
  begin
    FA_EndParaRPr.WriteAttributes(AWriter);
    if xaElements in FA_EndParaRPr.FAssigneds then 
    begin
      AWriter.BeginTag('a:endParaRPr');
      FA_EndParaRPr.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:endParaRPr');
  end;
end;

constructor TCT_TextParagraph.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 3;
  FAttributeCount := 0;
  FA_EG_TextRuns := TEG_TextRunXpgList.Create(FOwner);
end;

destructor TCT_TextParagraph.Destroy;
begin
  if FA_PPr <> Nil then 
    FA_PPr.Free;
  FA_EG_TextRuns.Free;
  if FA_EndParaRPr <> Nil then
    FA_EndParaRPr.Free;
end;

procedure TCT_TextParagraph.Clear;
begin
  FAssigneds := [];
  if FA_PPr <> Nil then
    FreeAndNil(FA_PPr);
  FA_EG_TextRuns.Clear;
  if FA_EndParaRPr <> Nil then
    FreeAndNil(FA_EndParaRPr);
end;

function TCT_TextParagraph.AppendText(AText: AxUCString): TCT_RegularTextRun;
begin
  Result := FA_EG_TextRuns.Add.Create_R;
  Result.T := AText;
end;

function TCT_TextParagraph.AppendText(AText: AxUCString; ASize: double; ABold, AItalic, AUnderline: boolean): TCT_RegularTextRun;
begin
  Result := AppendText(Atext);

  Result.Create_RPr;

  Result.RPr.Sz := Round(ASize * 100);
  Result.RPr.B := ABold;
  Result.RPr.I := AItalic;
  if AUnderline then
    Result.RPr.U := sttutSng;
end;

procedure TCT_TextParagraph.Assign(AItem: TCT_TextParagraph);
begin
end;

procedure TCT_TextParagraph.CopyTo(AItem: TCT_TextParagraph);
begin
end;

function  TCT_TextParagraph.Create_PPr: TCT_TextParagraphProperties;
begin
  if FA_PPr = Nil then
    FA_PPr := TCT_TextParagraphProperties.Create(FOwner);
  Result := FA_PPr;
end;

function  TCT_TextParagraph.Create_EndParaRPr: TCT_TextCharacterProperties;
begin
  if FA_EndParaRPr = Nil then
    FA_EndParaRPr := TCT_TextCharacterProperties.Create(FOwner);
  Result := FA_EndParaRPr;
end;

{ TCT_TextParagraphXpgList }

function  TCT_TextParagraphXpgList.GetItems(Index: integer): TCT_TextParagraph;
begin
  Result := TCT_TextParagraph(inherited Items[Index]);
end;

function  TCT_TextParagraphXpgList.Add: TCT_TextParagraph;
begin
  Result := TCT_TextParagraph.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_TextParagraphXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_TextParagraphXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
end;

procedure TCT_TextParagraphXpgList.Assign(AItem: TCT_TextParagraphXpgList);
begin
end;

procedure TCT_TextParagraphXpgList.CopyTo(AItem: TCT_TextParagraphXpgList);
begin
end;

{ TCT_GroupTransform2D }

function  TCT_GroupTransform2D.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FRot <> 0 then 
    Inc(AttrsAssigned);
  if FFlipH <> False then 
    Inc(AttrsAssigned);
  if FFlipV <> False then 
    Inc(AttrsAssigned);
  if FA_Off <> Nil then 
    Inc(ElemsAssigned,FA_Off.CheckAssigned);
  if FA_Ext <> Nil then 
    Inc(ElemsAssigned,FA_Ext.CheckAssigned);
  if FA_ChOff <> Nil then 
    Inc(ElemsAssigned,FA_ChOff.CheckAssigned);
  if FA_ChExt <> Nil then 
    Inc(ElemsAssigned,FA_ChExt.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_GroupTransform2D.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $000001D6: begin
      if FA_Off = Nil then 
        FA_Off := TCT_Point2D.Create(FOwner);
      Result := FA_Off;
    end;
    $000001EC: begin
      if FA_Ext = Nil then 
        FA_Ext := TCT_PositiveSize2D.Create(FOwner);
      Result := FA_Ext;
    end;
    $00000281: begin
      if FA_ChOff = Nil then 
        FA_ChOff := TCT_Point2D.Create(FOwner);
      Result := FA_ChOff;
    end;
    $00000297: begin
      if FA_ChExt = Nil then 
        FA_ChExt := TCT_PositiveSize2D.Create(FOwner);
      Result := FA_ChExt;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_GroupTransform2D.Write(AWriter: TXpgWriteXML);
begin
  if (FA_Off <> Nil) and FA_Off.Assigned then 
  begin
    FA_Off.WriteAttributes(AWriter);
    AWriter.SimpleTag('a:off');
  end;
  if (FA_Ext <> Nil) and FA_Ext.Assigned then 
  begin
    FA_Ext.WriteAttributes(AWriter);
    AWriter.SimpleTag('a:ext');
  end;
  if (FA_ChOff <> Nil) and FA_ChOff.Assigned then 
  begin
    FA_ChOff.WriteAttributes(AWriter);
    AWriter.SimpleTag('a:chOff');
  end;
  if (FA_ChExt <> Nil) and FA_ChExt.Assigned then 
  begin
    FA_ChExt.WriteAttributes(AWriter);
    AWriter.SimpleTag('a:chExt');
  end;
end;

procedure TCT_GroupTransform2D.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FRot <> 0 then 
    AWriter.AddAttribute('rot',XmlIntToStr(FRot));
  if FFlipH <> False then 
    AWriter.AddAttribute('flipH',XmlBoolToStr(FFlipH));
  if FFlipV <> False then 
    AWriter.AddAttribute('flipV',XmlBoolToStr(FFlipV));
end;

procedure TCT_GroupTransform2D.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $00000155: FRot := XmlStrToIntDef(AAttributes.Values[i],0);
      $000001F3: FFlipH := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000201: FFlipV := XmlStrToBoolDef(AAttributes.Values[i],False);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_GroupTransform2D.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 4;
  FAttributeCount := 3;
  FRot := 0;
  FFlipH := False;
  FFlipV := False;
end;

destructor TCT_GroupTransform2D.Destroy;
begin
  if FA_Off <> Nil then 
    FA_Off.Free;
  if FA_Ext <> Nil then 
    FA_Ext.Free;
  if FA_ChOff <> Nil then 
    FA_ChOff.Free;
  if FA_ChExt <> Nil then 
    FA_ChExt.Free;
end;

procedure TCT_GroupTransform2D.Clear;
begin
  FAssigneds := [];
  if FA_Off <> Nil then 
    FreeAndNil(FA_Off);
  if FA_Ext <> Nil then 
    FreeAndNil(FA_Ext);
  if FA_ChOff <> Nil then 
    FreeAndNil(FA_ChOff);
  if FA_ChExt <> Nil then 
    FreeAndNil(FA_ChExt);
  FRot := 0;
  FFlipH := False;
  FFlipV := False;
end;

procedure TCT_GroupTransform2D.Assign(AItem: TCT_GroupTransform2D);
begin
end;

procedure TCT_GroupTransform2D.CopyTo(AItem: TCT_GroupTransform2D);
begin
end;

function  TCT_GroupTransform2D.Create_A_Off: TCT_Point2D;
begin
  if FA_Off = Nil then
    FA_Off := TCT_Point2D.Create(FOwner);
  Result := FA_Off;
end;

function  TCT_GroupTransform2D.Create_A_Ext: TCT_PositiveSize2D;
begin
  if FA_Ext = Nil then
    FA_Ext := TCT_PositiveSize2D.Create(FOwner);
  Result := FA_Ext;
end;

function  TCT_GroupTransform2D.Create_A_ChOff: TCT_Point2D;
begin
  if FA_ChOff = Nil then
    FA_ChOff := TCT_Point2D.Create(FOwner);
  Result := FA_ChOff;
end;

function  TCT_GroupTransform2D.Create_A_ChExt: TCT_PositiveSize2D;
begin
  if FA_ChExt = Nil then
    FA_ChExt := TCT_PositiveSize2D.Create(FOwner);
  Result := FA_ChExt;
end;

{ TCT_ShapeProperties }

function  TCT_ShapeProperties.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 1;
  AttrsAssigned := 0;
  if Integer(FBwMode) <> XPG_UNKNOWN_ENUM then
    Inc(AttrsAssigned);
  if FA_Xfrm <> Nil then
    Inc(ElemsAssigned,FA_Xfrm.CheckAssigned);
  Inc(ElemsAssigned,FA_EG_Geometry.CheckAssigned);
  Inc(ElemsAssigned,FA_EG_FillProperties.CheckAssigned);
  if FA_Ln <> Nil then
    Inc(ElemsAssigned,FA_Ln.CheckAssigned);
  Inc(ElemsAssigned,FA_EG_EffectProperties.CheckAssigned);
  if FA_Scene3d <> Nil then
    Inc(ElemsAssigned,FA_Scene3d.CheckAssigned);
  if FA_Sp3d <> Nil then
    Inc(ElemsAssigned,FA_Sp3d.CheckAssigned);
  if FA_ExtLst <> Nil then
    Inc(ElemsAssigned,FA_ExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_ShapeProperties.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  case AReader.QNameHashA of
    $00000258: begin
      if FA_Xfrm = Nil then 
        FA_Xfrm := TCT_Transform2D.Create(FOwner);
      Result := FA_Xfrm;
    end;
    $00000175: begin
      if FA_Ln = Nil then 
        FA_Ln := TCT_LineProperties.Create(FOwner);
      Result := FA_Ln;
    end;
    $00000340: begin
      if FA_Scene3d = Nil then 
        FA_Scene3d := TCT_Scene3D.Create(FOwner);
      Result := FA_Scene3d;
    end;
    $00000215: begin
      if FA_Sp3d = Nil then 
        FA_Sp3d := TCT_Shape3D.Create(FOwner);
      Result := FA_Sp3d;
    end;
    $0000031F: begin
      if FA_ExtLst = Nil then 
        FA_ExtLst := TCT_OfficeArtExtensionList.Create(FOwner);
      Result := FA_ExtLst;
    end;
    else 
    begin
      Result := FA_EG_Geometry.HandleElement(AReader);
      if Result = Nil then 
      begin
        Result := FA_EG_FillProperties.HandleElement(AReader);
        if Result = Nil then 
        begin
          Result := FA_EG_EffectProperties.HandleElement(AReader);
          if Result = Nil then 
            FOwner.Errors.Error(xemUnknownElement,AReader.QName);
        end;
      end;
    end
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_ShapeProperties.Write(AWriter: TXpgWriteXML);
begin
  if (FA_Xfrm <> Nil) and FA_Xfrm.Assigned then 
  begin
    FA_Xfrm.WriteAttributes(AWriter);
    if xaElements in FA_Xfrm.FAssigneds then 
    begin
      AWriter.BeginTag('a:xfrm');
      FA_Xfrm.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:xfrm');
  end
  else
    AWriter.SimpleTag('a:xfrm');

  FA_EG_Geometry.Write(AWriter);
  FA_EG_FillProperties.Write(AWriter);
  if (FA_Ln <> Nil) and FA_Ln.Assigned then 
  begin
    FA_Ln.WriteAttributes(AWriter);
    if xaElements in FA_Ln.FAssigneds then 
    begin
      AWriter.BeginTag('a:ln');
      FA_Ln.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:ln');
  end;
  FA_EG_EffectProperties.Write(AWriter);
  if (FA_Scene3d <> Nil) and FA_Scene3d.Assigned then 
    if xaElements in FA_Scene3d.FAssigneds then 
    begin
      AWriter.BeginTag('a:scene3d');
      FA_Scene3d.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:scene3d');
  if (FA_Sp3d <> Nil) and FA_Sp3d.Assigned then 
  begin
    FA_Sp3d.WriteAttributes(AWriter);
    if xaElements in FA_Sp3d.FAssigneds then 
    begin
      AWriter.BeginTag('a:sp3d');
      FA_Sp3d.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:sp3d');
  end;
  if (FA_ExtLst <> Nil) and FA_ExtLst.Assigned then 
    if xaElements in FA_ExtLst.FAssigneds then 
    begin
      AWriter.BeginTag('a:extLst');
      FA_ExtLst.Write(AWriter);
      AWriter.EndTag;
    end;
end;

procedure TCT_ShapeProperties.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if Integer(FBwMode) <> XPG_UNKNOWN_ENUM then 
    AWriter.AddAttribute('bwMode',StrTST_BlackWhiteMode[Integer(FBwMode)]);
end;

procedure TCT_ShapeProperties.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'bwMode' then 
    FBwMode := TST_BlackWhiteMode(StrToEnum('stbwm' + AAttributes.Values[0]))
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_ShapeProperties.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 8;
  FAttributeCount := 1;
  FA_EG_Geometry := TEG_Geometry.Create(FOwner);
  FA_EG_FillProperties := TEG_FillProperties.Create(FOwner);
  FA_EG_EffectProperties := TEG_EffectProperties.Create(FOwner);
  FBwMode := TST_BlackWhiteMode(XPG_UNKNOWN_ENUM);
end;

destructor TCT_ShapeProperties.Destroy;
begin
  if FA_Xfrm <> Nil then 
    FA_Xfrm.Free;
  FA_EG_Geometry.Free;
  FA_EG_FillProperties.Free;
  if FA_Ln <> Nil then 
    FA_Ln.Free;
  FA_EG_EffectProperties.Free;
  if FA_Scene3d <> Nil then 
    FA_Scene3d.Free;
  if FA_Sp3d <> Nil then 
    FA_Sp3d.Free;
  if FA_ExtLst <> Nil then 
    FA_ExtLst.Free;
end;

procedure TCT_ShapeProperties.Free_Ln;
begin
  if FA_Ln <> Nil then begin
    FA_Ln.Free;
    FA_Ln := Nil;
  end;
end;

procedure TCT_ShapeProperties.Clear;
begin
  FAssigneds := [];
  if FA_Xfrm <> Nil then
    FreeAndNil(FA_Xfrm);
  FA_EG_Geometry.Clear;
  FA_EG_FillProperties.Clear;
  if FA_Ln <> Nil then
    FreeAndNil(FA_Ln);
  FA_EG_EffectProperties.Clear;
  if FA_Scene3d <> Nil then
    FreeAndNil(FA_Scene3d);
  if FA_Sp3d <> Nil then
    FreeAndNil(FA_Sp3d);
  if FA_ExtLst <> Nil then
    FreeAndNil(FA_ExtLst);
  FBwMode := TST_BlackWhiteMode(XPG_UNKNOWN_ENUM);
end;

procedure TCT_ShapeProperties.Assign(AItem: TCT_ShapeProperties);
begin
 FBwMode := AItem.FBwMode;
 if AItem.FA_Xfrm <> Nil then begin
   Create_Xfrm;
   FA_Xfrm.Assign(AItem.FA_Xfrm);
 end;
 FA_EG_Geometry.Assign(AItem.FA_EG_Geometry);
 FA_EG_FillProperties.Assign(AItem.FA_EG_FillProperties);
 if AItem.FA_Ln <> Nil then begin
   Create_Ln;
   FA_Ln.Assign(AItem.FA_Ln);
 end;
 FA_EG_EffectProperties.Assign(AItem.FA_EG_EffectProperties);
 if AItem.FA_Scene3d <> Nil then begin
   Create_Scene3d;
   FA_Scene3d.Assign(AItem.FA_Scene3d);
 end;
 if AItem.FA_Sp3d <> Nil then begin
   Create_Sp3d;
   FA_Sp3d.Assign(AItem.FA_Sp3d);
 end;
end;

procedure TCT_ShapeProperties.CopyTo(AItem: TCT_ShapeProperties);
begin
end;

function  TCT_ShapeProperties.Create_Xfrm: TCT_Transform2D;
begin
  if FA_Xfrm = Nil then
    FA_Xfrm := TCT_Transform2D.Create(FOwner);
  Result := FA_Xfrm;
end;

function  TCT_ShapeProperties.Create_Ln: TCT_LineProperties;
begin
  if FA_Ln = Nil then
    FA_Ln := TCT_LineProperties.Create(FOwner);
  Result := FA_Ln;
end;

function  TCT_ShapeProperties.Create_Scene3d: TCT_Scene3D;
begin
  if FA_Scene3d = Nil then
    FA_Scene3d := TCT_Scene3D.Create(FOwner);
  Result := FA_Scene3d;
end;

function  TCT_ShapeProperties.Create_Sp3d: TCT_Shape3D;
begin
  if FA_Sp3d = Nil then
    FA_Sp3d := TCT_Shape3D.Create(FOwner);
  Result := FA_Sp3d;
end;

function  TCT_ShapeProperties.Create_ExtLst: TCT_OfficeArtExtensionList;
begin
  if FA_ExtLst = Nil then
    FA_ExtLst := TCT_OfficeArtExtensionList.Create(FOwner);
  Result := FA_ExtLst;
end;

{ TCT_ShapeStyle }

function  TCT_ShapeStyle.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FA_LnRef <> Nil then
    Inc(ElemsAssigned,FA_LnRef.CheckAssigned);
  if FA_FillRef <> Nil then
    Inc(ElemsAssigned,FA_FillRef.CheckAssigned);
  if FA_EffectRef <> Nil then 
    Inc(ElemsAssigned,FA_EffectRef.CheckAssigned);
  if FA_FontRef <> Nil then 
    Inc(ElemsAssigned,FA_FontRef.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_ShapeStyle.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $00000292: begin
      if FA_LnRef = Nil then 
        FA_LnRef := TCT_StyleMatrixReference.Create(FOwner);
      Result := FA_LnRef;
    end;
    $0000035F: begin
      if FA_FillRef = Nil then 
        FA_FillRef := TCT_StyleMatrixReference.Create(FOwner);
      Result := FA_FillRef;
    end;
    $00000425: begin
      if FA_EffectRef = Nil then 
        FA_EffectRef := TCT_StyleMatrixReference.Create(FOwner);
      Result := FA_EffectRef;
    end;
    $0000036F: begin
      if FA_FontRef = Nil then 
        FA_FontRef := TCT_FontReference.Create(FOwner);
      Result := FA_FontRef;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_ShapeStyle.Write(AWriter: TXpgWriteXML);
begin
  if (FA_LnRef <> Nil) and FA_LnRef.Assigned then 
  begin
    FA_LnRef.WriteAttributes(AWriter);
    if xaElements in FA_LnRef.FAssigneds then 
    begin
      AWriter.BeginTag('a:lnRef');
      FA_LnRef.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:lnRef');
  end
  else 
    AWriter.SimpleTag('a:lnRef');
  if (FA_FillRef <> Nil) and FA_FillRef.Assigned then 
  begin
    FA_FillRef.WriteAttributes(AWriter);
    if xaElements in FA_FillRef.FAssigneds then 
    begin
      AWriter.BeginTag('a:fillRef');
      FA_FillRef.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:fillRef');
  end
  else 
    AWriter.SimpleTag('a:fillRef');
  if (FA_EffectRef <> Nil) and FA_EffectRef.Assigned then 
  begin
    FA_EffectRef.WriteAttributes(AWriter);
    if xaElements in FA_EffectRef.FAssigneds then 
    begin
      AWriter.BeginTag('a:effectRef');
      FA_EffectRef.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:effectRef');
  end
  else 
    AWriter.SimpleTag('a:effectRef');
  if (FA_FontRef <> Nil) and FA_FontRef.Assigned then 
  begin
    FA_FontRef.WriteAttributes(AWriter);
    if xaElements in FA_FontRef.FAssigneds then 
    begin
      AWriter.BeginTag('a:fontRef');
      FA_FontRef.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:fontRef');
  end
  else 
    AWriter.SimpleTag('a:fontRef');
end;

constructor TCT_ShapeStyle.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 4;
  FAttributeCount := 0;
end;

destructor TCT_ShapeStyle.Destroy;
begin
  if FA_LnRef <> Nil then 
    FA_LnRef.Free;
  if FA_FillRef <> Nil then 
    FA_FillRef.Free;
  if FA_EffectRef <> Nil then 
    FA_EffectRef.Free;
  if FA_FontRef <> Nil then 
    FA_FontRef.Free;
end;

procedure TCT_ShapeStyle.Clear;
begin
  FAssigneds := [];
  if FA_LnRef <> Nil then 
    FreeAndNil(FA_LnRef);
  if FA_FillRef <> Nil then 
    FreeAndNil(FA_FillRef);
  if FA_EffectRef <> Nil then 
    FreeAndNil(FA_EffectRef);
  if FA_FontRef <> Nil then 
    FreeAndNil(FA_FontRef);
end;

procedure TCT_ShapeStyle.Assign(AItem: TCT_ShapeStyle);
begin
end;

procedure TCT_ShapeStyle.CopyTo(AItem: TCT_ShapeStyle);
begin
end;

function  TCT_ShapeStyle.Create_A_LnRef: TCT_StyleMatrixReference;
begin
  if FA_LnRef = Nil then
    FA_LnRef := TCT_StyleMatrixReference.Create(FOwner);
  Result := FA_LnRef;
end;

function  TCT_ShapeStyle.Create_A_FillRef: TCT_StyleMatrixReference;
begin
  if FA_FillRef = Nil then
    FA_FillRef := TCT_StyleMatrixReference.Create(FOwner);
  Result := FA_FillRef;
end;

function  TCT_ShapeStyle.Create_A_EffectRef: TCT_StyleMatrixReference;
begin
  if FA_EffectRef = Nil then
    FA_EffectRef := TCT_StyleMatrixReference.Create(FOwner);
  Result := FA_EffectRef;
end;

function  TCT_ShapeStyle.Create_A_FontRef: TCT_FontReference;
begin
  if FA_FontRef = Nil then
    FA_FontRef := TCT_FontReference.Create(FOwner);
  Result := FA_FontRef;
end;

{ TCT_TextBody }

function  TCT_TextBody.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FA_BodyPr <> Nil then 
    Inc(ElemsAssigned,FA_BodyPr.CheckAssigned);
  if FA_LstStyle <> Nil then 
    Inc(ElemsAssigned,FA_LstStyle.CheckAssigned);
  if FA_PXpgList <> Nil then 
    Inc(ElemsAssigned,FA_PXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_TextBody.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $0000030B: begin
      if FA_BodyPr = Nil then
        FA_BodyPr := TCT_TextBodyProperties.Create(FOwner);
      Result := FA_BodyPr;
    end;
    $000003FF: begin
      if FA_LstStyle = Nil then
        FA_LstStyle := TCT_TextListStyle.Create(FOwner);
      Result := FA_LstStyle;
    end;
    $0000010B: begin
      if FA_PXpgList = Nil then
        FA_PXpgList := TCT_TextParagraphXpgList.Create(FOwner);
      Result := FA_PXpgList.Add;
    end;
    else
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then
    Result.Assigneds := [xaRead];
end;

function TCT_TextBody.MaxFontHeight: double;
var
  i,j: integer;
begin
  Result := 0;
  for i := 0 to FA_PXpgList.Count - 1 do begin
    for j := 0 to FA_PXpgList[i].FA_EG_TextRuns.Count - 1 do begin
      if FA_PXpgList[i].FA_EG_TextRuns[j].Run.RPr.Sz < 100000 then
        Result := Max(Result,FA_PXpgList[i].FA_EG_TextRuns[j].Run.RPr.Sz);
    end;
  end;
  Result := Result / 100;
end;

procedure TCT_TextBody.SetPlainText(const Value: AxUCString);
var
  P  : TCT_TextParagraph;
  TR : TEG_TextRun;
  RTR: TCT_RegularTextRun;
begin
  Create_Paras;

  FA_PXpgList.Clear;

  P := FA_PXpgList.Add;
  TR := P.TextRuns.Add;
  RTR := TR.Create_R;
  RTR.T := Value;
end;

procedure TCT_TextBody.Write(AWriter: TXpgWriteXML);
begin
  if (FA_BodyPr <> Nil) and FA_BodyPr.Assigned then 
  begin
    FA_BodyPr.WriteAttributes(AWriter);
    if xaElements in FA_BodyPr.FAssigneds then 
    begin
      AWriter.BeginTag('a:bodyPr');
      FA_BodyPr.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:bodyPr');
  end
  else 
    AWriter.SimpleTag('a:bodyPr');
// TODO
//  if (FA_LstStyle <> Nil) and FA_LstStyle.Assigned then
    if (FA_LstStyle <> Nil) and (xaElements in FA_LstStyle.FAssigneds) then
    begin
      AWriter.BeginTag('a:lstStyle');
      FA_LstStyle.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('a:lstStyle');
  if FA_PXpgList <> Nil then 
    FA_PXpgList.Write(AWriter,'a:p');
end;

constructor TCT_TextBody.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 3;
  FAttributeCount := 0;
end;

destructor TCT_TextBody.Destroy;
begin
  if FA_BodyPr <> Nil then 
    FA_BodyPr.Free;
  if FA_LstStyle <> Nil then
    FA_LstStyle.Free;
  if FA_PXpgList <> Nil then 
    FA_PXpgList.Free;
end;

function TCT_TextBody.GetPLainText: AxUCString;
var
  i,j: integer;
begin
  Result := '';
  if FA_PXpgList <> Nil then begin
    for i := 0 to FA_PXpgList.Count - 1 do begin
      for j := 0 to FA_PXpgList[i].FA_EG_TextRuns.Count - 1 do begin
        Result := Result + FA_PXpgList[i].FA_EG_TextRuns[j].Run.FA_T;
      end;
    end;
  end;
end;

procedure TCT_TextBody.Clear;
begin
  FAssigneds := [];
  if FA_BodyPr <> Nil then 
    FreeAndNil(FA_BodyPr);
  if FA_LstStyle <> Nil then
    FreeAndNil(FA_LstStyle);
  if FA_PXpgList <> Nil then
    FreeAndNil(FA_PXpgList);
end;

procedure TCT_TextBody.Assign(AItem: TCT_TextBody);
begin
end;

procedure TCT_TextBody.CopyTo(AItem: TCT_TextBody);
begin
end;

function  TCT_TextBody.Create_BodyPr: TCT_TextBodyProperties;
begin
  if FA_BodyPr = Nil then
    FA_BodyPr := TCT_TextBodyProperties.Create(FOwner);
  Result := FA_BodyPr;
end;

function  TCT_TextBody.Create_LstStyle: TCT_TextListStyle;
begin
  if FA_LstStyle = Nil then
    FA_LstStyle := TCT_TextListStyle.Create(FOwner);
  Result := FA_LstStyle;
end;

function  TCT_TextBody.Create_Paras: TCT_TextParagraphXpgList;
begin
  if FA_PXpgList = Nil then
    FA_PXpgList := TCT_TextParagraphXpgList.Create(FOwner);
  Result := FA_PXpgList;
end;

{ TCT_GroupShapeProperties }

function  TCT_GroupShapeProperties.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if Integer(FBwMode) <> XPG_UNKNOWN_ENUM then
    Inc(AttrsAssigned);
  if FA_Xfrm <> Nil then 
    Inc(ElemsAssigned,FA_Xfrm.CheckAssigned);
  Inc(ElemsAssigned,FA_EG_FillProperties.CheckAssigned);
  Inc(ElemsAssigned,FA_EG_EffectProperties.CheckAssigned);
  if FA_Scene3d <> Nil then 
    Inc(ElemsAssigned,FA_Scene3d.CheckAssigned);
  if FA_ExtLst <> Nil then 
    Inc(ElemsAssigned,FA_ExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_GroupShapeProperties.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  case AReader.QNameHashA of
    $00000258: begin
      if FA_Xfrm = Nil then 
        FA_Xfrm := TCT_GroupTransform2D.Create(FOwner);
      Result := FA_Xfrm;
    end;
    $00000340: begin
      if FA_Scene3d = Nil then 
        FA_Scene3d := TCT_Scene3D.Create(FOwner);
      Result := FA_Scene3d;
    end;
    $0000031F: begin
      if FA_ExtLst = Nil then 
        FA_ExtLst := TCT_OfficeArtExtensionList.Create(FOwner);
      Result := FA_ExtLst;
    end;
    else 
    begin
      Result := FA_EG_FillProperties.HandleElement(AReader);
      if Result = Nil then 
      begin
        Result := FA_EG_EffectProperties.HandleElement(AReader);
        if Result = Nil then 
          FOwner.Errors.Error(xemUnknownElement,AReader.QName);
      end;
    end
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_GroupShapeProperties.Write(AWriter: TXpgWriteXML);
begin
  if (FA_Xfrm <> Nil) and FA_Xfrm.Assigned then 
  begin
    FA_Xfrm.WriteAttributes(AWriter);
    if xaElements in FA_Xfrm.FAssigneds then 
    begin
      AWriter.BeginTag('a:xfrm');
      FA_Xfrm.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:xfrm');
  end;
  FA_EG_FillProperties.Write(AWriter);
  FA_EG_EffectProperties.Write(AWriter);
  if (FA_Scene3d <> Nil) and FA_Scene3d.Assigned then 
    if xaElements in FA_Scene3d.FAssigneds then 
    begin
      AWriter.BeginTag('a:scene3d');
      FA_Scene3d.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:scene3d');
  if (FA_ExtLst <> Nil) and FA_ExtLst.Assigned then 
    if xaElements in FA_ExtLst.FAssigneds then 
    begin
      AWriter.BeginTag('a:extLst');
      FA_ExtLst.Write(AWriter);
      AWriter.EndTag;
    end;
end;

procedure TCT_GroupShapeProperties.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if Integer(FBwMode) <> XPG_UNKNOWN_ENUM then 
    AWriter.AddAttribute('bwMode',StrTST_BlackWhiteMode[Integer(FBwMode)]);
end;

procedure TCT_GroupShapeProperties.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'bwMode' then 
    FBwMode := TST_BlackWhiteMode(StrToEnum('stbwm' + AAttributes.Values[0]))
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_GroupShapeProperties.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 5;
  FAttributeCount := 1;
  FA_EG_FillProperties := TEG_FillProperties.Create(FOwner);
  FA_EG_EffectProperties := TEG_EffectProperties.Create(FOwner);
  FBwMode := TST_BlackWhiteMode(XPG_UNKNOWN_ENUM);
end;

destructor TCT_GroupShapeProperties.Destroy;
begin
  if FA_Xfrm <> Nil then
    FA_Xfrm.Free;
  FA_EG_FillProperties.Free;
  FA_EG_EffectProperties.Free;
  if FA_Scene3d <> Nil then
    FA_Scene3d.Free;
  if FA_ExtLst <> Nil then
    FA_ExtLst.Free;
end;

procedure TCT_GroupShapeProperties.Clear;
begin
  FAssigneds := [];
  if FA_Xfrm <> Nil then
    FreeAndNil(FA_Xfrm);
  FA_EG_FillProperties.Clear;
  FA_EG_EffectProperties.Clear;
  if FA_Scene3d <> Nil then
    FreeAndNil(FA_Scene3d);
  if FA_ExtLst <> Nil then
    FreeAndNil(FA_ExtLst);
  FBwMode := TST_BlackWhiteMode(XPG_UNKNOWN_ENUM);
end;

procedure TCT_GroupShapeProperties.Assign(AItem: TCT_GroupShapeProperties);
begin
end;

procedure TCT_GroupShapeProperties.CopyTo(AItem: TCT_GroupShapeProperties);
begin
end;

function  TCT_GroupShapeProperties.Create_A_Xfrm: TCT_GroupTransform2D;
begin
  if FA_Xfrm = Nil then
    FA_Xfrm := TCT_GroupTransform2D.Create(FOwner);
  Result := FA_Xfrm;
end;

function  TCT_GroupShapeProperties.Create_A_Scene3d: TCT_Scene3D;
begin
  if FA_Scene3d = Nil then
    FA_Scene3d := TCT_Scene3D.Create(FOwner);
  Result := FA_Scene3d;
end;

function  TCT_GroupShapeProperties.Create_A_ExtLst: TCT_OfficeArtExtensionList;
begin
  if FA_ExtLst = Nil then
    FA_ExtLst := TCT_OfficeArtExtensionList.Create(FOwner);
  Result := FA_ExtLst;
end;

{ TCT_Shape }

function  TCT_Shape.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FMacro <> '' then 
    Inc(AttrsAssigned);
  if FTextlink <> '' then 
    Inc(AttrsAssigned);
  if FFLocksText <> True then 
    Inc(AttrsAssigned);
  if FFPublished <> False then 
    Inc(AttrsAssigned);
  if FNvSpPr <> Nil then 
    Inc(ElemsAssigned,FNvSpPr.CheckAssigned);
  if FSpPr <> Nil then 
    Inc(ElemsAssigned,FSpPr.CheckAssigned);
  if FStyle <> Nil then 
    Inc(ElemsAssigned,FStyle.CheckAssigned);
  if FTxBody <> Nil then 
    Inc(ElemsAssigned,FTxBody.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_Shape.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $000003F1: begin
      if FNvSpPr = Nil then
        FNvSpPr := TCT_ShapeNonVisual.Create(FOwner);
      Result := FNvSpPr;
    end;
    $0000032D: begin
      if FSpPr = Nil then
        FSpPr := TCT_ShapeProperties.Create(FOwner);
      Result := FSpPr;
    end;
    $000003B9: begin
      if FStyle = Nil then
        FStyle := TCT_ShapeStyle.Create(FOwner);
      Result := FStyle;
    end;
    $00000402: begin
      if FTxBody = Nil then
        FTxBody := TCT_TextBody.Create(FOwner);
      Result := FTxBody;
    end;
    else begin
      if AReader.QName = 'cdr:nvSpPr' then begin
        if FNvSpPr = Nil then
          FNvSpPr := TCT_ShapeNonVisual.Create(FOwner);
        Result := FNvSpPr;
      end
      else if AReader.QName = 'cdr:SpPr' then begin
        if FSpPr = Nil then
          FSpPr := TCT_ShapeProperties.Create(FOwner);
        Result := FSpPr;
      end
      else if AReader.QName = 'cdr:style' then begin
        if FStyle = Nil then
          FStyle := TCT_ShapeStyle.Create(FOwner);
        Result := FNvSpPr;
      end
      else if AReader.QName = 'cdr:txBody' then begin
        if FTxBody = Nil then
          FTxBody := TCT_TextBody.Create(FOwner);
        Result := FTxBody;
      end
      else
        FOwner.Errors.Error(xemUnknownElement,AReader.QName);
    end;
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_Shape.Write(AWriter: TXpgWriteXML);
begin
  if (FNvSpPr <> Nil) and FNvSpPr.Assigned then 
    if xaElements in FNvSpPr.FAssigneds then 
    begin
      AWriter.BeginTag(FOwner.NS + ':nvSpPr');
      FNvSpPr.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag(FOwner.NS + ':nvSpPr')
  else
    AWriter.SimpleTag(FOwner.NS + ':nvSpPr');
  if (FSpPr <> Nil) and FSpPr.Assigned then 
  begin
    FSpPr.WriteAttributes(AWriter);
    if xaElements in FSpPr.FAssigneds then 
    begin
      AWriter.BeginTag(FOwner.NS + ':spPr');
      FSpPr.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(FOwner.NS + ':spPr');
  end
  else 
    AWriter.SimpleTag(FOwner.NS + ':spPr');
  if (FStyle <> Nil) and FStyle.Assigned then 
    if xaElements in FStyle.FAssigneds then 
    begin
      AWriter.BeginTag(FOwner.NS + ':style');
      FStyle.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(FOwner.NS + ':style');
  if (FTxBody <> Nil) and FTxBody.Assigned then 
    if xaElements in FTxBody.FAssigneds then 
    begin
      AWriter.BeginTag(FOwner.NS + ':txBody');
      FTxBody.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(FOwner.NS + ':txBody');
end;

procedure TCT_Shape.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FMacro <> '' then 
    AWriter.AddAttribute('macro',FMacro);
  if FTextlink <> '' then 
    AWriter.AddAttribute('textlink',FTextlink);
  if FFLocksText <> True then 
    AWriter.AddAttribute('fLocksText',XmlBoolToStr(FFLocksText));
  if FFPublished <> False then 
    AWriter.AddAttribute('fPublished',XmlBoolToStr(FFPublished));
end;

procedure TCT_Shape.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $00000212: FMacro := AAttributes.Values[i];
      $00000373: FTextlink := AAttributes.Values[i];
      $00000407: FFLocksText := XmlStrToBoolDef(AAttributes.Values[i],True);
      $00000406: FFPublished := XmlStrToBoolDef(AAttributes.Values[i],False);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_Shape.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 4;
  FAttributeCount := 4;
  FFLocksText := True;
  FFPublished := False;
end;

destructor TCT_Shape.Destroy;
begin
  if FNvSpPr <> Nil then 
    FNvSpPr.Free;
  if FSpPr <> Nil then 
    FSpPr.Free;
  if FStyle <> Nil then 
    FStyle.Free;
  if FTxBody <> Nil then 
    FTxBody.Free;
end;

procedure TCT_Shape.Clear;
begin
  FAssigneds := [];
  if FNvSpPr <> Nil then 
    FreeAndNil(FNvSpPr);
  if FSpPr <> Nil then 
    FreeAndNil(FSpPr);
  if FStyle <> Nil then 
    FreeAndNil(FStyle);
  if FTxBody <> Nil then 
    FreeAndNil(FTxBody);
  FMacro := '';
  FTextlink := '';
  FFLocksText := True;
  FFPublished := False;
end;

procedure TCT_Shape.Assign(AItem: TCT_Shape);
begin
end;

procedure TCT_Shape.CopyTo(AItem: TCT_Shape);
begin
end;

function  TCT_Shape.Create_NvSpPr: TCT_ShapeNonVisual;
begin
  if FNvSpPr = Nil then
    FNvSpPr := TCT_ShapeNonVisual.Create(FOwner);
  Result := FNvSpPr;
end;

function  TCT_Shape.Create_SpPr: TCT_ShapeProperties;
begin
  if FSpPr = Nil then
    FSpPr := TCT_ShapeProperties.Create(FOwner);
  Result := FSpPr;
end;

function  TCT_Shape.Create_Style: TCT_ShapeStyle;
begin
  if FStyle = Nil then
    FStyle := TCT_ShapeStyle.Create(FOwner);
  Result := FStyle;
end;

function  TCT_Shape.Create_TxBody: TCT_TextBody;
begin
  if FTxBody = Nil then
    FTxBody := TCT_TextBody.Create(FOwner);
  Result := FTxBody;
end;

{ TCT_SupplementalFont }

function  TCT_SupplementalFont.CheckAssigned: integer;
begin
  FAssigneds := [xaAttributes];
  Result := 1;
end;

procedure TCT_SupplementalFont.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_SupplementalFont.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('script',FScript);
  AWriter.AddAttribute('typeface',FTypeface);
end;

procedure TCT_SupplementalFont.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $00000295: FScript := AAttributes.Values[i];
      $00000351: FTypeface := AAttributes.Values[i];
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_SupplementalFont.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 2;
end;

destructor TCT_SupplementalFont.Destroy;
begin
end;

procedure TCT_SupplementalFont.Clear;
begin
  FAssigneds := [];
  FScript := '';
  FTypeface := '';
end;

procedure TCT_SupplementalFont.Assign(AItem: TCT_SupplementalFont);
begin
end;

procedure TCT_SupplementalFont.CopyTo(AItem: TCT_SupplementalFont);
begin
end;

{ TCT_SupplementalFontXpgList }

function  TCT_SupplementalFontXpgList.GetItems(Index: integer): TCT_SupplementalFont;
begin
  Result := TCT_SupplementalFont(inherited Items[Index]);
end;

function  TCT_SupplementalFontXpgList.Add: TCT_SupplementalFont;
begin
  Result := TCT_SupplementalFont.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_SupplementalFontXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_SupplementalFontXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_SupplementalFontXpgList.Assign(AItem: TCT_SupplementalFontXpgList);
begin
end;

procedure TCT_SupplementalFontXpgList.CopyTo(AItem: TCT_SupplementalFontXpgList);
begin
end;

{ TCT_EffectStyleItem }

function  TCT_EffectStyleItem.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FA_EG_EffectProperties.CheckAssigned);
  if FA_Scene3d <> Nil then 
    Inc(ElemsAssigned,FA_Scene3d.CheckAssigned);
  if FA_Sp3d <> Nil then 
    Inc(ElemsAssigned,FA_Sp3d.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_EffectStyleItem.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  case AReader.QNameHashA of
    $00000340: begin
      if FA_Scene3d = Nil then 
        FA_Scene3d := TCT_Scene3D.Create(FOwner);
      Result := FA_Scene3d;
    end;
    $00000215: begin
      if FA_Sp3d = Nil then 
        FA_Sp3d := TCT_Shape3D.Create(FOwner);
      Result := FA_Sp3d;
    end;
    else 
    begin
      Result := FA_EG_EffectProperties.HandleElement(AReader);
      if Result = Nil then 
        FOwner.Errors.Error(xemUnknownElement,AReader.QName);
    end
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_EffectStyleItem.Write(AWriter: TXpgWriteXML);
begin
  FA_EG_EffectProperties.Write(AWriter);
  if (FA_Scene3d <> Nil) and FA_Scene3d.Assigned then 
    if xaElements in FA_Scene3d.FAssigneds then 
    begin
      AWriter.BeginTag('a:scene3d');
      FA_Scene3d.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:scene3d');
  if (FA_Sp3d <> Nil) and FA_Sp3d.Assigned then 
  begin
    FA_Sp3d.WriteAttributes(AWriter);
    if xaElements in FA_Sp3d.FAssigneds then 
    begin
      AWriter.BeginTag('a:sp3d');
      FA_Sp3d.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:sp3d');
  end;
end;

constructor TCT_EffectStyleItem.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 3;
  FAttributeCount := 0;
  FA_EG_EffectProperties := TEG_EffectProperties.Create(FOwner);
end;

destructor TCT_EffectStyleItem.Destroy;
begin
  FA_EG_EffectProperties.Free;
  if FA_Scene3d <> Nil then 
    FA_Scene3d.Free;
  if FA_Sp3d <> Nil then 
    FA_Sp3d.Free;
end;

procedure TCT_EffectStyleItem.Clear;
begin
  FAssigneds := [];
  FA_EG_EffectProperties.Clear;
  if FA_Scene3d <> Nil then 
    FreeAndNil(FA_Scene3d);
  if FA_Sp3d <> Nil then 
    FreeAndNil(FA_Sp3d);
end;

procedure TCT_EffectStyleItem.Assign(AItem: TCT_EffectStyleItem);
begin
end;

procedure TCT_EffectStyleItem.CopyTo(AItem: TCT_EffectStyleItem);
begin
end;

function  TCT_EffectStyleItem.Create_A_Scene3d: TCT_Scene3D;
begin
  if FA_Scene3d = Nil then
    FA_Scene3d := TCT_Scene3D.Create(FOwner);
  Result := FA_Scene3d;
end;

function  TCT_EffectStyleItem.Create_A_Sp3d: TCT_Shape3D;
begin
  if FA_Sp3d = Nil then
    FA_Sp3d := TCT_Shape3D.Create(FOwner);
  Result := FA_Sp3d;
end;

{ TCT_EffectStyleItemXpgList }

function  TCT_EffectStyleItemXpgList.GetItems(Index: integer): TCT_EffectStyleItem;
begin
  Result := TCT_EffectStyleItem(inherited Items[Index]);
end;

function  TCT_EffectStyleItemXpgList.Add: TCT_EffectStyleItem;
begin
  Result := TCT_EffectStyleItem.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_EffectStyleItemXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_EffectStyleItemXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
end;

procedure TCT_EffectStyleItemXpgList.Assign(AItem: TCT_EffectStyleItemXpgList);
begin
end;

procedure TCT_EffectStyleItemXpgList.CopyTo(AItem: TCT_EffectStyleItemXpgList);
begin
end;

{ TCT_FontCollection }

function  TCT_FontCollection.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FA_Latin <> Nil then 
    Inc(ElemsAssigned,FA_Latin.CheckAssigned);
  if FA_Ea <> Nil then 
    Inc(ElemsAssigned,FA_Ea.CheckAssigned);
  if FA_Cs <> Nil then 
    Inc(ElemsAssigned,FA_Cs.CheckAssigned);
  if FA_FontXpgList <> Nil then 
    Inc(ElemsAssigned,FA_FontXpgList.CheckAssigned);
  if FA_ExtLst <> Nil then 
    Inc(ElemsAssigned,FA_ExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_FontCollection.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $000002B3: begin
      if FA_Latin = Nil then 
        FA_Latin := TCT_TextFont.Create(FOwner);
      Result := FA_Latin;
    end;
    $00000161: begin
      if FA_Ea = Nil then 
        FA_Ea := TCT_TextFont.Create(FOwner);
      Result := FA_Ea;
    end;
    $00000171: begin
      if FA_Cs = Nil then 
        FA_Cs := TCT_TextFont.Create(FOwner);
      Result := FA_Cs;
    end;
    $00000252: begin
      if FA_FontXpgList = Nil then 
        FA_FontXpgList := TCT_SupplementalFontXpgList.Create(FOwner);
      Result := FA_FontXpgList.Add;
    end;
    $0000031F: begin
      if FA_ExtLst = Nil then 
        FA_ExtLst := TCT_OfficeArtExtensionList.Create(FOwner);
      Result := FA_ExtLst;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_FontCollection.Write(AWriter: TXpgWriteXML);
begin
  if (FA_Latin <> Nil) and FA_Latin.Assigned then 
  begin
    FA_Latin.WriteAttributes(AWriter);
    AWriter.SimpleTag('a:latin');
  end
  else 
    AWriter.SimpleTag('a:latin');
  if (FA_Ea <> Nil) and FA_Ea.Assigned then 
  begin
    FA_Ea.WriteAttributes(AWriter);
    AWriter.SimpleTag('a:ea');
  end
  else 
    AWriter.SimpleTag('a:ea');
  if (FA_Cs <> Nil) and FA_Cs.Assigned then 
  begin
    FA_Cs.WriteAttributes(AWriter);
    AWriter.SimpleTag('a:cs');
  end
  else 
    AWriter.SimpleTag('a:cs');
  if FA_FontXpgList <> Nil then 
    FA_FontXpgList.Write(AWriter,'a:font');
  if (FA_ExtLst <> Nil) and FA_ExtLst.Assigned then 
    if xaElements in FA_ExtLst.FAssigneds then 
    begin
      AWriter.BeginTag('a:extLst');
      FA_ExtLst.Write(AWriter);
      AWriter.EndTag;
    end;
end;

constructor TCT_FontCollection.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 5;
  FAttributeCount := 0;
end;

destructor TCT_FontCollection.Destroy;
begin
  if FA_Latin <> Nil then 
    FA_Latin.Free;
  if FA_Ea <> Nil then 
    FA_Ea.Free;
  if FA_Cs <> Nil then 
    FA_Cs.Free;
  if FA_FontXpgList <> Nil then 
    FA_FontXpgList.Free;
  if FA_ExtLst <> Nil then 
    FA_ExtLst.Free;
end;

procedure TCT_FontCollection.Clear;
begin
  FAssigneds := [];
  if FA_Latin <> Nil then 
    FreeAndNil(FA_Latin);
  if FA_Ea <> Nil then 
    FreeAndNil(FA_Ea);
  if FA_Cs <> Nil then 
    FreeAndNil(FA_Cs);
  if FA_FontXpgList <> Nil then 
    FreeAndNil(FA_FontXpgList);
  if FA_ExtLst <> Nil then 
    FreeAndNil(FA_ExtLst);
end;

procedure TCT_FontCollection.Assign(AItem: TCT_FontCollection);
begin
end;

procedure TCT_FontCollection.CopyTo(AItem: TCT_FontCollection);
begin
end;

function  TCT_FontCollection.Create_A_Latin: TCT_TextFont;
begin
  if FA_Latin = Nil then
    FA_Latin := TCT_TextFont.Create(FOwner);
  Result := FA_Latin;
end;

function  TCT_FontCollection.Create_A_Ea: TCT_TextFont;
begin
  if FA_Ea = Nil then
    FA_Ea := TCT_TextFont.Create(FOwner);
  Result := FA_Ea;
end;

function  TCT_FontCollection.Create_A_Cs: TCT_TextFont;
begin
  if FA_Cs = Nil then
    FA_Cs := TCT_TextFont.Create(FOwner);
  Result := FA_Cs;
end;

function  TCT_FontCollection.Create_A_FontXpgList: TCT_SupplementalFontXpgList;
begin
  if FA_FontXpgList = Nil then
    FA_FontXpgList := TCT_SupplementalFontXpgList.Create(FOwner);
  Result := FA_FontXpgList;
end;

function  TCT_FontCollection.Create_A_ExtLst: TCT_OfficeArtExtensionList;
begin
  if FA_ExtLst = Nil then
    FA_ExtLst := TCT_OfficeArtExtensionList.Create(FOwner);
  Result := FA_ExtLst;
end;

{ TCT_FillStyleList }

function  TCT_FillStyleList.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FA_EG_FillProperties.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_FillStyleList.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := FA_EG_FillProperties.HandleElement(AReader);
  if Result = Nil then 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_FillStyleList.Write(AWriter: TXpgWriteXML);
begin
  FA_EG_FillProperties.Write(AWriter);
end;

constructor TCT_FillStyleList.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 0;
  FA_EG_FillProperties := TEG_FillProperties.Create(FOwner);
end;

destructor TCT_FillStyleList.Destroy;
begin
  FA_EG_FillProperties.Free;
end;

procedure TCT_FillStyleList.Clear;
begin
  FAssigneds := [];
  FA_EG_FillProperties.Clear;
end;

procedure TCT_FillStyleList.Assign(AItem: TCT_FillStyleList);
begin
end;

procedure TCT_FillStyleList.CopyTo(AItem: TCT_FillStyleList);
begin
end;

{ TCT_LineStyleList }

function  TCT_LineStyleList.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FA_LnXpgList <> Nil then 
    Inc(ElemsAssigned,FA_LnXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_LineStyleList.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'a:ln' then 
  begin
    if FA_LnXpgList = Nil then 
      FA_LnXpgList := TCT_LinePropertiesXpgList.Create(FOwner);
    Result := FA_LnXpgList.Add;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_LineStyleList.Write(AWriter: TXpgWriteXML);
begin
  if FA_LnXpgList <> Nil then 
    FA_LnXpgList.Write(AWriter,'a:ln');
end;

constructor TCT_LineStyleList.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 0;
end;

destructor TCT_LineStyleList.Destroy;
begin
  if FA_LnXpgList <> Nil then 
    FA_LnXpgList.Free;
end;

procedure TCT_LineStyleList.Clear;
begin
  FAssigneds := [];
  if FA_LnXpgList <> Nil then 
    FreeAndNil(FA_LnXpgList);
end;

procedure TCT_LineStyleList.Assign(AItem: TCT_LineStyleList);
begin
end;

procedure TCT_LineStyleList.CopyTo(AItem: TCT_LineStyleList);
begin
end;

function  TCT_LineStyleList.Create_A_LnXpgList: TCT_LinePropertiesXpgList;
begin
  if FA_LnXpgList = Nil then
    FA_LnXpgList := TCT_LinePropertiesXpgList.Create(FOwner);
  Result := FA_LnXpgList;
end;

{ TCT_EffectStyleList }

function  TCT_EffectStyleList.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FA_EffectStyleXpgList <> Nil then 
    Inc(ElemsAssigned,FA_EffectStyleXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_EffectStyleList.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'a:effectStyle' then 
  begin
    if FA_EffectStyleXpgList = Nil then 
      FA_EffectStyleXpgList := TCT_EffectStyleItemXpgList.Create(FOwner);
    Result := FA_EffectStyleXpgList.Add;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_EffectStyleList.Write(AWriter: TXpgWriteXML);
begin
  if FA_EffectStyleXpgList <> Nil then 
    FA_EffectStyleXpgList.Write(AWriter,'a:effectStyle');
end;

constructor TCT_EffectStyleList.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 0;
end;

destructor TCT_EffectStyleList.Destroy;
begin
  if FA_EffectStyleXpgList <> Nil then 
    FA_EffectStyleXpgList.Free;
end;

procedure TCT_EffectStyleList.Clear;
begin
  FAssigneds := [];
  if FA_EffectStyleXpgList <> Nil then 
    FreeAndNil(FA_EffectStyleXpgList);
end;

procedure TCT_EffectStyleList.Assign(AItem: TCT_EffectStyleList);
begin
end;

procedure TCT_EffectStyleList.CopyTo(AItem: TCT_EffectStyleList);
begin
end;

function  TCT_EffectStyleList.Create_A_EffectStyleXpgList: TCT_EffectStyleItemXpgList;
begin
  if FA_EffectStyleXpgList = Nil then
    FA_EffectStyleXpgList := TCT_EffectStyleItemXpgList.Create(FOwner);
  Result := FA_EffectStyleXpgList;
end;

{ TCT_BackgroundFillStyleList }

function  TCT_BackgroundFillStyleList.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FA_EG_FillProperties.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_BackgroundFillStyleList.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := FA_EG_FillProperties.HandleElement(AReader);
  if Result = Nil then 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_BackgroundFillStyleList.Write(AWriter: TXpgWriteXML);
begin
  FA_EG_FillProperties.Write(AWriter);
end;

constructor TCT_BackgroundFillStyleList.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 0;
  FA_EG_FillProperties := TEG_FillProperties.Create(FOwner);
end;

destructor TCT_BackgroundFillStyleList.Destroy;
begin
  FA_EG_FillProperties.Free;
end;

procedure TCT_BackgroundFillStyleList.Clear;
begin
  FAssigneds := [];
  FA_EG_FillProperties.Clear;
end;

procedure TCT_BackgroundFillStyleList.Assign(AItem: TCT_BackgroundFillStyleList);
begin
end;

procedure TCT_BackgroundFillStyleList.CopyTo(AItem: TCT_BackgroundFillStyleList);
begin
end;

{ TCT_Ratio }

function  TCT_Ratio.CheckAssigned: integer;
begin
  FAssigneds := [xaAttributes];
  Result := 1;
end;

procedure TCT_Ratio.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_Ratio.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('n',XmlIntToStr(FN));
  AWriter.AddAttribute('d',XmlIntToStr(FD));
end;

procedure TCT_Ratio.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $0000006E: FN := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000064: FD := XmlStrToIntDef(AAttributes.Values[i],0);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_Ratio.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 2;
  FN := 2147483632;
  FD := 2147483632;
end;

destructor TCT_Ratio.Destroy;
begin
end;

procedure TCT_Ratio.Clear;
begin
  FAssigneds := [];
  FN := 2147483632;
  FD := 2147483632;
end;

procedure TCT_Ratio.Assign(AItem: TCT_Ratio);
begin
end;

procedure TCT_Ratio.CopyTo(AItem: TCT_Ratio);
begin
end;

{ TCT_CustomColor }

function  TCT_CustomColor.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FName <> '' then 
    Inc(AttrsAssigned);
  Inc(ElemsAssigned,FA_EG_ColorChoice.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_CustomColor.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := FA_EG_ColorChoice.HandleElement(AReader);
  if Result = Nil then 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_CustomColor.Write(AWriter: TXpgWriteXML);
begin
  FA_EG_ColorChoice.Write(AWriter);
end;

procedure TCT_CustomColor.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FName <> '' then 
    AWriter.AddAttribute('name',FName);
end;

procedure TCT_CustomColor.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'name' then 
    FName := AAttributes.Values[0]
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_CustomColor.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 1;
  FA_EG_ColorChoice := TEG_ColorChoice.Create(FOwner);
end;

destructor TCT_CustomColor.Destroy;
begin
  FA_EG_ColorChoice.Free;
end;

procedure TCT_CustomColor.Clear;
begin
  FAssigneds := [];
  FA_EG_ColorChoice.Clear;
  FName := '';
end;

procedure TCT_CustomColor.Assign(AItem: TCT_CustomColor);
begin
end;

procedure TCT_CustomColor.CopyTo(AItem: TCT_CustomColor);
begin
end;

{ TCT_CustomColorXpgList }

function  TCT_CustomColorXpgList.GetItems(Index: integer): TCT_CustomColor;
begin
  Result := TCT_CustomColor(inherited Items[Index]);
end;

function  TCT_CustomColorXpgList.Add: TCT_CustomColor;
begin
  Result := TCT_CustomColor.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_CustomColorXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_CustomColorXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_CustomColorXpgList.Assign(AItem: TCT_CustomColorXpgList);
begin
end;

procedure TCT_CustomColorXpgList.CopyTo(AItem: TCT_CustomColorXpgList);
begin
end;

{ TCT_ColorScheme }

function  TCT_ColorScheme.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [xaAttributes];
  if FA_Dk1 <> Nil then 
    Inc(ElemsAssigned,FA_Dk1.CheckAssigned);
  if FA_Lt1 <> Nil then 
    Inc(ElemsAssigned,FA_Lt1.CheckAssigned);
  if FA_Dk2 <> Nil then 
    Inc(ElemsAssigned,FA_Dk2.CheckAssigned);
  if FA_Lt2 <> Nil then 
    Inc(ElemsAssigned,FA_Lt2.CheckAssigned);
  if FA_Accent1 <> Nil then 
    Inc(ElemsAssigned,FA_Accent1.CheckAssigned);
  if FA_Accent2 <> Nil then 
    Inc(ElemsAssigned,FA_Accent2.CheckAssigned);
  if FA_Accent3 <> Nil then 
    Inc(ElemsAssigned,FA_Accent3.CheckAssigned);
  if FA_Accent4 <> Nil then 
    Inc(ElemsAssigned,FA_Accent4.CheckAssigned);
  if FA_Accent5 <> Nil then 
    Inc(ElemsAssigned,FA_Accent5.CheckAssigned);
  if FA_Accent6 <> Nil then 
    Inc(ElemsAssigned,FA_Accent6.CheckAssigned);
  if FA_Hlink <> Nil then 
    Inc(ElemsAssigned,FA_Hlink.CheckAssigned);
  if FA_FolHlink <> Nil then 
    Inc(ElemsAssigned,FA_FolHlink.CheckAssigned);
  if FA_ExtLst <> Nil then 
    Inc(ElemsAssigned,FA_ExtLst.CheckAssigned);
  Result := 1;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_ColorScheme.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $0000019B: begin
      if FA_Dk1 = Nil then 
        FA_Dk1 := TCT_Color.Create(FOwner);
      Result := FA_Dk1;
    end;
    $000001AC: begin
      if FA_Lt1 = Nil then 
        FA_Lt1 := TCT_Color.Create(FOwner);
      Result := FA_Lt1;
    end;
    $0000019C: begin
      if FA_Dk2 = Nil then 
        FA_Dk2 := TCT_Color.Create(FOwner);
      Result := FA_Dk2;
    end;
    $000001AD: begin
      if FA_Lt2 = Nil then 
        FA_Lt2 := TCT_Color.Create(FOwner);
      Result := FA_Lt2;
    end;
    $0000033A: begin
      if FA_Accent1 = Nil then 
        FA_Accent1 := TCT_Color.Create(FOwner);
      Result := FA_Accent1;
    end;
    $0000033B: begin
      if FA_Accent2 = Nil then 
        FA_Accent2 := TCT_Color.Create(FOwner);
      Result := FA_Accent2;
    end;
    $0000033C: begin
      if FA_Accent3 = Nil then 
        FA_Accent3 := TCT_Color.Create(FOwner);
      Result := FA_Accent3;
    end;
    $0000033D: begin
      if FA_Accent4 = Nil then 
        FA_Accent4 := TCT_Color.Create(FOwner);
      Result := FA_Accent4;
    end;
    $0000033E: begin
      if FA_Accent5 = Nil then 
        FA_Accent5 := TCT_Color.Create(FOwner);
      Result := FA_Accent5;
    end;
    $0000033F: begin
      if FA_Accent6 = Nil then 
        FA_Accent6 := TCT_Color.Create(FOwner);
      Result := FA_Accent6;
    end;
    $000002B1: begin
      if FA_Hlink = Nil then 
        FA_Hlink := TCT_Color.Create(FOwner);
      Result := FA_Hlink;
    end;
    $000003D2: begin
      if FA_FolHlink = Nil then 
        FA_FolHlink := TCT_Color.Create(FOwner);
      Result := FA_FolHlink;
    end;
    $0000031F: begin
      if FA_ExtLst = Nil then 
        FA_ExtLst := TCT_OfficeArtExtensionList.Create(FOwner);
      Result := FA_ExtLst;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_ColorScheme.Write(AWriter: TXpgWriteXML);
begin
  if (FA_Dk1 <> Nil) and FA_Dk1.Assigned then 
    if xaElements in FA_Dk1.FAssigneds then 
    begin
      AWriter.BeginTag('a:dk1');
      FA_Dk1.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:dk1')
  else 
    AWriter.SimpleTag('a:dk1');
  if (FA_Lt1 <> Nil) and FA_Lt1.Assigned then 
    if xaElements in FA_Lt1.FAssigneds then 
    begin
      AWriter.BeginTag('a:lt1');
      FA_Lt1.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:lt1')
  else 
    AWriter.SimpleTag('a:lt1');
  if (FA_Dk2 <> Nil) and FA_Dk2.Assigned then 
    if xaElements in FA_Dk2.FAssigneds then 
    begin
      AWriter.BeginTag('a:dk2');
      FA_Dk2.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:dk2')
  else 
    AWriter.SimpleTag('a:dk2');
  if (FA_Lt2 <> Nil) and FA_Lt2.Assigned then 
    if xaElements in FA_Lt2.FAssigneds then 
    begin
      AWriter.BeginTag('a:lt2');
      FA_Lt2.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:lt2')
  else 
    AWriter.SimpleTag('a:lt2');
  if (FA_Accent1 <> Nil) and FA_Accent1.Assigned then 
    if xaElements in FA_Accent1.FAssigneds then 
    begin
      AWriter.BeginTag('a:accent1');
      FA_Accent1.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:accent1')
  else 
    AWriter.SimpleTag('a:accent1');
  if (FA_Accent2 <> Nil) and FA_Accent2.Assigned then 
    if xaElements in FA_Accent2.FAssigneds then 
    begin
      AWriter.BeginTag('a:accent2');
      FA_Accent2.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:accent2')
  else 
    AWriter.SimpleTag('a:accent2');
  if (FA_Accent3 <> Nil) and FA_Accent3.Assigned then 
    if xaElements in FA_Accent3.FAssigneds then 
    begin
      AWriter.BeginTag('a:accent3');
      FA_Accent3.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:accent3')
  else 
    AWriter.SimpleTag('a:accent3');
  if (FA_Accent4 <> Nil) and FA_Accent4.Assigned then 
    if xaElements in FA_Accent4.FAssigneds then 
    begin
      AWriter.BeginTag('a:accent4');
      FA_Accent4.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:accent4')
  else 
    AWriter.SimpleTag('a:accent4');
  if (FA_Accent5 <> Nil) and FA_Accent5.Assigned then 
    if xaElements in FA_Accent5.FAssigneds then 
    begin
      AWriter.BeginTag('a:accent5');
      FA_Accent5.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:accent5')
  else 
    AWriter.SimpleTag('a:accent5');
  if (FA_Accent6 <> Nil) and FA_Accent6.Assigned then 
    if xaElements in FA_Accent6.FAssigneds then 
    begin
      AWriter.BeginTag('a:accent6');
      FA_Accent6.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:accent6')
  else 
    AWriter.SimpleTag('a:accent6');
  if (FA_Hlink <> Nil) and FA_Hlink.Assigned then 
    if xaElements in FA_Hlink.FAssigneds then 
    begin
      AWriter.BeginTag('a:hlink');
      FA_Hlink.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:hlink')
  else 
    AWriter.SimpleTag('a:hlink');
  if (FA_FolHlink <> Nil) and FA_FolHlink.Assigned then 
    if xaElements in FA_FolHlink.FAssigneds then 
    begin
      AWriter.BeginTag('a:folHlink');
      FA_FolHlink.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:folHlink')
  else 
    AWriter.SimpleTag('a:folHlink');
  if (FA_ExtLst <> Nil) and FA_ExtLst.Assigned then 
    if xaElements in FA_ExtLst.FAssigneds then 
    begin
      AWriter.BeginTag('a:extLst');
      FA_ExtLst.Write(AWriter);
      AWriter.EndTag;
    end;
end;

procedure TCT_ColorScheme.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('name',FName);
end;

procedure TCT_ColorScheme.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'name' then 
    FName := AAttributes.Values[0]
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_ColorScheme.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 13;
  FAttributeCount := 1;
end;

destructor TCT_ColorScheme.Destroy;
begin
  if FA_Dk1 <> Nil then 
    FA_Dk1.Free;
  if FA_Lt1 <> Nil then 
    FA_Lt1.Free;
  if FA_Dk2 <> Nil then 
    FA_Dk2.Free;
  if FA_Lt2 <> Nil then 
    FA_Lt2.Free;
  if FA_Accent1 <> Nil then 
    FA_Accent1.Free;
  if FA_Accent2 <> Nil then 
    FA_Accent2.Free;
  if FA_Accent3 <> Nil then 
    FA_Accent3.Free;
  if FA_Accent4 <> Nil then 
    FA_Accent4.Free;
  if FA_Accent5 <> Nil then 
    FA_Accent5.Free;
  if FA_Accent6 <> Nil then 
    FA_Accent6.Free;
  if FA_Hlink <> Nil then 
    FA_Hlink.Free;
  if FA_FolHlink <> Nil then 
    FA_FolHlink.Free;
  if FA_ExtLst <> Nil then 
    FA_ExtLst.Free;
end;

procedure TCT_ColorScheme.Clear;
begin
  FAssigneds := [];
  if FA_Dk1 <> Nil then 
    FreeAndNil(FA_Dk1);
  if FA_Lt1 <> Nil then 
    FreeAndNil(FA_Lt1);
  if FA_Dk2 <> Nil then 
    FreeAndNil(FA_Dk2);
  if FA_Lt2 <> Nil then 
    FreeAndNil(FA_Lt2);
  if FA_Accent1 <> Nil then 
    FreeAndNil(FA_Accent1);
  if FA_Accent2 <> Nil then 
    FreeAndNil(FA_Accent2);
  if FA_Accent3 <> Nil then 
    FreeAndNil(FA_Accent3);
  if FA_Accent4 <> Nil then 
    FreeAndNil(FA_Accent4);
  if FA_Accent5 <> Nil then 
    FreeAndNil(FA_Accent5);
  if FA_Accent6 <> Nil then 
    FreeAndNil(FA_Accent6);
  if FA_Hlink <> Nil then 
    FreeAndNil(FA_Hlink);
  if FA_FolHlink <> Nil then 
    FreeAndNil(FA_FolHlink);
  if FA_ExtLst <> Nil then 
    FreeAndNil(FA_ExtLst);
  FName := '';
end;

procedure TCT_ColorScheme.Assign(AItem: TCT_ColorScheme);
begin
end;

procedure TCT_ColorScheme.CopyTo(AItem: TCT_ColorScheme);
begin
end;

function  TCT_ColorScheme.Create_A_Dk1: TCT_Color;
begin
  if FA_Dk1 = Nil then
    FA_Dk1 := TCT_Color.Create(FOwner);
  Result := FA_Dk1;
end;

function  TCT_ColorScheme.Create_A_Lt1: TCT_Color;
begin
  if FA_Lt1 = Nil then
    FA_Lt1 := TCT_Color.Create(FOwner);
  Result := FA_Lt1;
end;

function  TCT_ColorScheme.Create_A_Dk2: TCT_Color;
begin
  if FA_Dk2 = Nil then
    FA_Dk2 := TCT_Color.Create(FOwner);
  Result := FA_Dk2;
end;

function  TCT_ColorScheme.Create_A_Lt2: TCT_Color;
begin
  if FA_Lt2 = Nil then
    FA_Lt2 := TCT_Color.Create(FOwner);
  Result := FA_Lt2;
end;

function  TCT_ColorScheme.Create_A_Accent1: TCT_Color;
begin
  if FA_Accent1 = Nil then
    FA_Accent1 := TCT_Color.Create(FOwner);
  Result := FA_Accent1;
end;

function  TCT_ColorScheme.Create_A_Accent2: TCT_Color;
begin
  if FA_Accent2 = Nil then
    FA_Accent2 := TCT_Color.Create(FOwner);
  Result := FA_Accent2;
end;

function  TCT_ColorScheme.Create_A_Accent3: TCT_Color;
begin
  if FA_Accent3 = Nil then
    FA_Accent3 := TCT_Color.Create(FOwner);
  Result := FA_Accent3;
end;

function  TCT_ColorScheme.Create_A_Accent4: TCT_Color;
begin
  if FA_Accent4 = Nil then
    FA_Accent4 := TCT_Color.Create(FOwner);
  Result := FA_Accent4;
end;

function  TCT_ColorScheme.Create_A_Accent5: TCT_Color;
begin
  if FA_Accent5 = Nil then
    FA_Accent5 := TCT_Color.Create(FOwner);
  Result := FA_Accent5;
end;

function  TCT_ColorScheme.Create_A_Accent6: TCT_Color;
begin
  if FA_Accent6 = Nil then
    FA_Accent6 := TCT_Color.Create(FOwner);
  Result := FA_Accent6;
end;

function  TCT_ColorScheme.Create_A_Hlink: TCT_Color;
begin
  if FA_Hlink = Nil then
    FA_Hlink := TCT_Color.Create(FOwner);
  Result := FA_Hlink;
end;

function  TCT_ColorScheme.Create_A_FolHlink: TCT_Color;
begin
  if FA_FolHlink = Nil then
    FA_FolHlink := TCT_Color.Create(FOwner);
  Result := FA_FolHlink;
end;

function  TCT_ColorScheme.Create_A_ExtLst: TCT_OfficeArtExtensionList;
begin
  if FA_ExtLst = Nil then
    FA_ExtLst := TCT_OfficeArtExtensionList.Create(FOwner);
  Result := FA_ExtLst;
end;

{ TCT_FontScheme }

function  TCT_FontScheme.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [xaAttributes];
  if FA_MajorFont <> Nil then
    Inc(ElemsAssigned,FA_MajorFont.CheckAssigned);
  if FA_MinorFont <> Nil then
    Inc(ElemsAssigned,FA_MinorFont.CheckAssigned);
  if FA_ExtLst <> Nil then
    Inc(ElemsAssigned,FA_ExtLst.CheckAssigned);
  Result := 1;
  if ElemsAssigned > 0 then
    FAssigneds := FAssigneds + [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_FontScheme.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $0000044B: begin
      if FA_MajorFont = Nil then 
        FA_MajorFont := TCT_FontCollection.Create(FOwner);
      Result := FA_MajorFont;
    end;
    $00000457: begin
      if FA_MinorFont = Nil then 
        FA_MinorFont := TCT_FontCollection.Create(FOwner);
      Result := FA_MinorFont;
    end;
    $0000031F: begin
      if FA_ExtLst = Nil then 
        FA_ExtLst := TCT_OfficeArtExtensionList.Create(FOwner);
      Result := FA_ExtLst;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_FontScheme.Write(AWriter: TXpgWriteXML);
begin
  if (FA_MajorFont <> Nil) and FA_MajorFont.Assigned then 
    if xaElements in FA_MajorFont.FAssigneds then 
    begin
      AWriter.BeginTag('a:majorFont');
      FA_MajorFont.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:majorFont')
  else 
    AWriter.SimpleTag('a:majorFont');
  if (FA_MinorFont <> Nil) and FA_MinorFont.Assigned then 
    if xaElements in FA_MinorFont.FAssigneds then 
    begin
      AWriter.BeginTag('a:minorFont');
      FA_MinorFont.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:minorFont')
  else 
    AWriter.SimpleTag('a:minorFont');
  if (FA_ExtLst <> Nil) and FA_ExtLst.Assigned then 
    if xaElements in FA_ExtLst.FAssigneds then 
    begin
      AWriter.BeginTag('a:extLst');
      FA_ExtLst.Write(AWriter);
      AWriter.EndTag;
    end;
end;

procedure TCT_FontScheme.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('name',FName);
end;

procedure TCT_FontScheme.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'name' then 
    FName := AAttributes.Values[0]
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_FontScheme.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 3;
  FAttributeCount := 1;
end;

destructor TCT_FontScheme.Destroy;
begin
  if FA_MajorFont <> Nil then 
    FA_MajorFont.Free;
  if FA_MinorFont <> Nil then 
    FA_MinorFont.Free;
  if FA_ExtLst <> Nil then 
    FA_ExtLst.Free;
end;

procedure TCT_FontScheme.Clear;
begin
  FAssigneds := [];
  if FA_MajorFont <> Nil then 
    FreeAndNil(FA_MajorFont);
  if FA_MinorFont <> Nil then 
    FreeAndNil(FA_MinorFont);
  if FA_ExtLst <> Nil then 
    FreeAndNil(FA_ExtLst);
  FName := '';
end;

procedure TCT_FontScheme.Assign(AItem: TCT_FontScheme);
begin
end;

procedure TCT_FontScheme.CopyTo(AItem: TCT_FontScheme);
begin
end;

function  TCT_FontScheme.Create_A_MajorFont: TCT_FontCollection;
begin
  if FA_MajorFont = Nil then
    FA_MajorFont := TCT_FontCollection.Create(FOwner);
  Result := FA_MajorFont;
end;

function  TCT_FontScheme.Create_A_MinorFont: TCT_FontCollection;
begin
  if FA_MinorFont = Nil then
    FA_MinorFont := TCT_FontCollection.Create(FOwner);
  Result := FA_MinorFont;
end;

function  TCT_FontScheme.Create_A_ExtLst: TCT_OfficeArtExtensionList;
begin
  if FA_ExtLst = Nil then
    FA_ExtLst := TCT_OfficeArtExtensionList.Create(FOwner);
  Result := FA_ExtLst;
end;

{ TCT_StyleMatrix }

function  TCT_StyleMatrix.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FName <> '' then 
    Inc(AttrsAssigned);
  if FA_FillStyleLst <> Nil then 
    Inc(ElemsAssigned,FA_FillStyleLst.CheckAssigned);
  if FA_LnStyleLst <> Nil then 
    Inc(ElemsAssigned,FA_LnStyleLst.CheckAssigned);
  if FA_EffectStyleLst <> Nil then 
    Inc(ElemsAssigned,FA_EffectStyleLst.CheckAssigned);
  if FA_BgFillStyleLst <> Nil then 
    Inc(ElemsAssigned,FA_BgFillStyleLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_StyleMatrix.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $00000586: begin
      if FA_FillStyleLst = Nil then 
        FA_FillStyleLst := TCT_FillStyleList.Create(FOwner);
      Result := FA_FillStyleLst;
    end;
    $000004B9: begin
      if FA_LnStyleLst = Nil then 
        FA_LnStyleLst := TCT_LineStyleList.Create(FOwner);
      Result := FA_LnStyleLst;
    end;
    $0000064C: begin
      if FA_EffectStyleLst = Nil then 
        FA_EffectStyleLst := TCT_EffectStyleList.Create(FOwner);
      Result := FA_EffectStyleLst;
    end;
    $0000062F: begin
      if FA_BgFillStyleLst = Nil then 
        FA_BgFillStyleLst := TCT_BackgroundFillStyleList.Create(FOwner);
      Result := FA_BgFillStyleLst;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_StyleMatrix.Write(AWriter: TXpgWriteXML);
begin
  if (FA_FillStyleLst <> Nil) and FA_FillStyleLst.Assigned then 
    if xaElements in FA_FillStyleLst.FAssigneds then 
    begin
      AWriter.BeginTag('a:fillStyleLst');
      FA_FillStyleLst.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:fillStyleLst')
  else 
    AWriter.SimpleTag('a:fillStyleLst');
  if (FA_LnStyleLst <> Nil) and FA_LnStyleLst.Assigned then 
    if xaElements in FA_LnStyleLst.FAssigneds then 
    begin
      AWriter.BeginTag('a:lnStyleLst');
      FA_LnStyleLst.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:lnStyleLst')
  else 
    AWriter.SimpleTag('a:lnStyleLst');
  if (FA_EffectStyleLst <> Nil) and FA_EffectStyleLst.Assigned then 
    if xaElements in FA_EffectStyleLst.FAssigneds then 
    begin
      AWriter.BeginTag('a:effectStyleLst');
      FA_EffectStyleLst.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:effectStyleLst')
  else 
    AWriter.SimpleTag('a:effectStyleLst');
  if (FA_BgFillStyleLst <> Nil) and FA_BgFillStyleLst.Assigned then 
    if xaElements in FA_BgFillStyleLst.FAssigneds then 
    begin
      AWriter.BeginTag('a:bgFillStyleLst');
      FA_BgFillStyleLst.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:bgFillStyleLst')
  else 
    AWriter.SimpleTag('a:bgFillStyleLst');
end;

procedure TCT_StyleMatrix.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FName <> '' then 
    AWriter.AddAttribute('name',FName);
end;

procedure TCT_StyleMatrix.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'name' then 
    FName := AAttributes.Values[0]
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_StyleMatrix.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 4;
  FAttributeCount := 1;
end;

destructor TCT_StyleMatrix.Destroy;
begin
  if FA_FillStyleLst <> Nil then 
    FA_FillStyleLst.Free;
  if FA_LnStyleLst <> Nil then 
    FA_LnStyleLst.Free;
  if FA_EffectStyleLst <> Nil then 
    FA_EffectStyleLst.Free;
  if FA_BgFillStyleLst <> Nil then 
    FA_BgFillStyleLst.Free;
end;

procedure TCT_StyleMatrix.Clear;
begin
  FAssigneds := [];
  if FA_FillStyleLst <> Nil then 
    FreeAndNil(FA_FillStyleLst);
  if FA_LnStyleLst <> Nil then 
    FreeAndNil(FA_LnStyleLst);
  if FA_EffectStyleLst <> Nil then 
    FreeAndNil(FA_EffectStyleLst);
  if FA_BgFillStyleLst <> Nil then 
    FreeAndNil(FA_BgFillStyleLst);
  FName := '';
end;

procedure TCT_StyleMatrix.Assign(AItem: TCT_StyleMatrix);
begin
end;

procedure TCT_StyleMatrix.CopyTo(AItem: TCT_StyleMatrix);
begin
end;

function  TCT_StyleMatrix.Create_A_FillStyleLst: TCT_FillStyleList;
begin
  if FA_FillStyleLst = Nil then
    FA_FillStyleLst := TCT_FillStyleList.Create(FOwner);
  Result := FA_FillStyleLst;
end;

function  TCT_StyleMatrix.Create_A_LnStyleLst: TCT_LineStyleList;
begin
  if FA_LnStyleLst = Nil then
    FA_LnStyleLst := TCT_LineStyleList.Create(FOwner);
  Result := FA_LnStyleLst;
end;

function  TCT_StyleMatrix.Create_A_EffectStyleLst: TCT_EffectStyleList;
begin
  if FA_EffectStyleLst = Nil then
    FA_EffectStyleLst := TCT_EffectStyleList.Create(FOwner);
  Result := FA_EffectStyleLst;
end;

function  TCT_StyleMatrix.Create_A_BgFillStyleLst: TCT_BackgroundFillStyleList;
begin
  if FA_BgFillStyleLst = Nil then
    FA_BgFillStyleLst := TCT_BackgroundFillStyleList.Create(FOwner);
  Result := FA_BgFillStyleLst;
end;

{ TEG_TextGeometry }

function  TEG_TextGeometry.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FA_CustGeom <> Nil then 
    Inc(ElemsAssigned,FA_CustGeom.CheckAssigned);
  if FA_PrstTxWarp <> Nil then 
    Inc(ElemsAssigned,FA_PrstTxWarp.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TEG_TextGeometry.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  case AReader.QNameHashA of
    $000003E2: begin
      if FA_CustGeom = Nil then 
        FA_CustGeom := TCT_CustomGeometry2D.Create(FOwner);
      Result := FA_CustGeom;
    end;
    $000004CA: begin
      if FA_PrstTxWarp = Nil then 
        FA_PrstTxWarp := TCT_PresetTextShape.Create(FOwner);
      Result := FA_PrstTxWarp;
    end;
    else 
    begin
      Result := Nil;
      Exit;
    end
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TEG_TextGeometry.Write(AWriter: TXpgWriteXML);
begin
  if (FA_CustGeom <> Nil) and FA_CustGeom.Assigned then 
    if xaElements in FA_CustGeom.FAssigneds then 
    begin
      AWriter.BeginTag('a:custGeom');
      FA_CustGeom.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:custGeom');
  if (FA_PrstTxWarp <> Nil) and FA_PrstTxWarp.Assigned then 
  begin
    FA_PrstTxWarp.WriteAttributes(AWriter);
    if xaElements in FA_PrstTxWarp.FAssigneds then 
    begin
      AWriter.BeginTag('a:prstTxWarp');
      FA_PrstTxWarp.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:prstTxWarp');
  end;
end;

constructor TEG_TextGeometry.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 2;
  FAttributeCount := 0;
end;

destructor TEG_TextGeometry.Destroy;
begin
  if FA_CustGeom <> Nil then 
    FA_CustGeom.Free;
  if FA_PrstTxWarp <> Nil then 
    FA_PrstTxWarp.Free;
end;

procedure TEG_TextGeometry.Clear;
begin
  FAssigneds := [];
  if FA_CustGeom <> Nil then 
    FreeAndNil(FA_CustGeom);
  if FA_PrstTxWarp <> Nil then 
    FreeAndNil(FA_PrstTxWarp);
end;

procedure TEG_TextGeometry.Assign(AItem: TEG_TextGeometry);
begin
end;

procedure TEG_TextGeometry.CopyTo(AItem: TEG_TextGeometry);
begin
end;

function  TEG_TextGeometry.Create_A_CustGeom: TCT_CustomGeometry2D;
begin
  if FA_CustGeom = Nil then
    FA_CustGeom := TCT_CustomGeometry2D.Create(FOwner);
  Result := FA_CustGeom;
end;

function  TEG_TextGeometry.Create_A_PrstTxWarp: TCT_PresetTextShape;
begin
  if FA_PrstTxWarp = Nil then
    FA_PrstTxWarp := TCT_PresetTextShape.Create(FOwner);
  Result := FA_PrstTxWarp;
end;

{ TCT_Scale2D }

function  TCT_Scale2D.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FA_Sx <> Nil then 
    Inc(ElemsAssigned,FA_Sx.CheckAssigned);
  if FA_Sy <> Nil then 
    Inc(ElemsAssigned,FA_Sy.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_Scale2D.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $00000186: begin
      if FA_Sx = Nil then 
        FA_Sx := TCT_Ratio.Create(FOwner);
      Result := FA_Sx;
    end;
    $00000187: begin
      if FA_Sy = Nil then 
        FA_Sy := TCT_Ratio.Create(FOwner);
      Result := FA_Sy;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_Scale2D.Write(AWriter: TXpgWriteXML);
begin
  if (FA_Sx <> Nil) and FA_Sx.Assigned then 
  begin
    FA_Sx.WriteAttributes(AWriter);
    AWriter.SimpleTag('a:sx');
  end
  else 
    AWriter.SimpleTag('a:sx');
  if (FA_Sy <> Nil) and FA_Sy.Assigned then 
  begin
    FA_Sy.WriteAttributes(AWriter);
    AWriter.SimpleTag('a:sy');
  end
  else 
    AWriter.SimpleTag('a:sy');
end;

constructor TCT_Scale2D.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 2;
  FAttributeCount := 0;
end;

destructor TCT_Scale2D.Destroy;
begin
  if FA_Sx <> Nil then 
    FA_Sx.Free;
  if FA_Sy <> Nil then 
    FA_Sy.Free;
end;

procedure TCT_Scale2D.Clear;
begin
  FAssigneds := [];
  if FA_Sx <> Nil then 
    FreeAndNil(FA_Sx);
  if FA_Sy <> Nil then 
    FreeAndNil(FA_Sy);
end;

procedure TCT_Scale2D.Assign(AItem: TCT_Scale2D);
begin
end;

procedure TCT_Scale2D.CopyTo(AItem: TCT_Scale2D);
begin
end;

function  TCT_Scale2D.Create_A_Sx: TCT_Ratio;
begin
  if FA_Sx = Nil then
    FA_Sx := TCT_Ratio.Create(FOwner);
  Result := FA_Sx;
end;

function  TCT_Scale2D.Create_A_Sy: TCT_Ratio;
begin
  if FA_Sy = Nil then
    FA_Sy := TCT_Ratio.Create(FOwner);
  Result := FA_Sy;
end;

{ TCT_FillProperties }

function  TCT_FillProperties.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FA_EG_FillProperties.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_FillProperties.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := FA_EG_FillProperties.HandleElement(AReader);
  if Result = Nil then 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_FillProperties.Write(AWriter: TXpgWriteXML);
begin
  FA_EG_FillProperties.Write(AWriter);
end;

constructor TCT_FillProperties.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 0;
  FA_EG_FillProperties := TEG_FillProperties.Create(FOwner);
end;

destructor TCT_FillProperties.Destroy;
begin
  FA_EG_FillProperties.Free;
end;

procedure TCT_FillProperties.Clear;
begin
  FAssigneds := [];
  FA_EG_FillProperties.Clear;
end;

procedure TCT_FillProperties.Assign(AItem: TCT_FillProperties);
begin
end;

procedure TCT_FillProperties.CopyTo(AItem: TCT_FillProperties);
begin
end;

{ TCT_EffectProperties }

function  TCT_EffectProperties.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FA_EG_EffectProperties.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_EffectProperties.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := FA_EG_EffectProperties.HandleElement(AReader);
  if Result = Nil then 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_EffectProperties.Write(AWriter: TXpgWriteXML);
begin
  FA_EG_EffectProperties.Write(AWriter);
end;

constructor TCT_EffectProperties.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 0;
  FA_EG_EffectProperties := TEG_EffectProperties.Create(FOwner);
end;

destructor TCT_EffectProperties.Destroy;
begin
  FA_EG_EffectProperties.Free;
end;

procedure TCT_EffectProperties.Clear;
begin
  FAssigneds := [];
  FA_EG_EffectProperties.Clear;
end;

procedure TCT_EffectProperties.Assign(AItem: TCT_EffectProperties);
begin
end;

procedure TCT_EffectProperties.CopyTo(AItem: TCT_EffectProperties);
begin
end;

{ TCT_CustomColorList }

function  TCT_CustomColorList.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FA_CustClrXpgList <> Nil then 
    Inc(ElemsAssigned,FA_CustClrXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_CustomColorList.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'a:custClr' then 
  begin
    if FA_CustClrXpgList = Nil then 
      FA_CustClrXpgList := TCT_CustomColorXpgList.Create(FOwner);
    Result := FA_CustClrXpgList.Add;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_CustomColorList.Write(AWriter: TXpgWriteXML);
begin
  if FA_CustClrXpgList <> Nil then 
    FA_CustClrXpgList.Write(AWriter,'a:custClr');
end;

constructor TCT_CustomColorList.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 0;
end;

destructor TCT_CustomColorList.Destroy;
begin
  if FA_CustClrXpgList <> Nil then 
    FA_CustClrXpgList.Free;
end;

procedure TCT_CustomColorList.Clear;
begin
  FAssigneds := [];
  if FA_CustClrXpgList <> Nil then 
    FreeAndNil(FA_CustClrXpgList);
end;

procedure TCT_CustomColorList.Assign(AItem: TCT_CustomColorList);
begin
end;

procedure TCT_CustomColorList.CopyTo(AItem: TCT_CustomColorList);
begin
end;

function  TCT_CustomColorList.Create_A_CustClrXpgList: TCT_CustomColorXpgList;
begin
  if FA_CustClrXpgList = Nil then
    FA_CustClrXpgList := TCT_CustomColorXpgList.Create(FOwner);
  Result := FA_CustClrXpgList;
end;

{ TCT_ColorMRU }

function  TCT_ColorMRU.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FA_EG_ColorChoice.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_ColorMRU.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := FA_EG_ColorChoice.HandleElement(AReader);
  if Result = Nil then 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_ColorMRU.Write(AWriter: TXpgWriteXML);
begin
  FA_EG_ColorChoice.Write(AWriter);
end;

constructor TCT_ColorMRU.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 0;
  FA_EG_ColorChoice := TEG_ColorChoice.Create(FOwner);
end;

destructor TCT_ColorMRU.Destroy;
begin
  FA_EG_ColorChoice.Free;
end;

procedure TCT_ColorMRU.Clear;
begin
  FAssigneds := [];
  FA_EG_ColorChoice.Clear;
end;

procedure TCT_ColorMRU.Assign(AItem: TCT_ColorMRU);
begin
end;

procedure TCT_ColorMRU.CopyTo(AItem: TCT_ColorMRU);
begin
end;

{ TCT_BaseStyles }

function  TCT_BaseStyles.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FA_ClrScheme <> Nil then 
    Inc(ElemsAssigned,FA_ClrScheme.CheckAssigned);
  if FA_FontScheme <> Nil then 
    Inc(ElemsAssigned,FA_FontScheme.CheckAssigned);
  if FA_FmtScheme <> Nil then 
    Inc(ElemsAssigned,FA_FmtScheme.CheckAssigned);
  if FA_ExtLst <> Nil then 
    Inc(ElemsAssigned,FA_ExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_BaseStyles.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $00000431: begin
      if FA_ClrScheme = Nil then 
        FA_ClrScheme := TCT_ColorScheme.Create(FOwner);
      Result := FA_ClrScheme;
    end;
    $000004A7: begin
      if FA_FontScheme = Nil then 
        FA_FontScheme := TCT_FontScheme.Create(FOwner);
      Result := FA_FontScheme;
    end;
    $00000437: begin
      if FA_FmtScheme = Nil then 
        FA_FmtScheme := TCT_StyleMatrix.Create(FOwner);
      Result := FA_FmtScheme;
    end;
    $0000031F: begin
      if FA_ExtLst = Nil then 
        FA_ExtLst := TCT_OfficeArtExtensionList.Create(FOwner);
      Result := FA_ExtLst;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_BaseStyles.Write(AWriter: TXpgWriteXML);
begin
  if (FA_ClrScheme <> Nil) and FA_ClrScheme.Assigned then 
  begin
    FA_ClrScheme.WriteAttributes(AWriter);
    if xaElements in FA_ClrScheme.FAssigneds then 
    begin
      AWriter.BeginTag('a:clrScheme');
      FA_ClrScheme.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:clrScheme');
  end
  else 
    AWriter.SimpleTag('a:clrScheme');
  if (FA_FontScheme <> Nil) and FA_FontScheme.Assigned then 
  begin
    FA_FontScheme.WriteAttributes(AWriter);
    if xaElements in FA_FontScheme.FAssigneds then 
    begin
      AWriter.BeginTag('a:fontScheme');
      FA_FontScheme.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:fontScheme');
  end
  else 
    AWriter.SimpleTag('a:fontScheme');
  if (FA_FmtScheme <> Nil) and FA_FmtScheme.Assigned then 
  begin
    FA_FmtScheme.WriteAttributes(AWriter);
    if xaElements in FA_FmtScheme.FAssigneds then 
    begin
      AWriter.BeginTag('a:fmtScheme');
      FA_FmtScheme.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:fmtScheme');
  end
  else 
    AWriter.SimpleTag('a:fmtScheme');
  if (FA_ExtLst <> Nil) and FA_ExtLst.Assigned then 
    if xaElements in FA_ExtLst.FAssigneds then 
    begin
      AWriter.BeginTag('a:extLst');
      FA_ExtLst.Write(AWriter);
      AWriter.EndTag;
    end;
end;

constructor TCT_BaseStyles.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 4;
  FAttributeCount := 0;
end;

destructor TCT_BaseStyles.Destroy;
begin
  if FA_ClrScheme <> Nil then 
    FA_ClrScheme.Free;
  if FA_FontScheme <> Nil then 
    FA_FontScheme.Free;
  if FA_FmtScheme <> Nil then 
    FA_FmtScheme.Free;
  if FA_ExtLst <> Nil then 
    FA_ExtLst.Free;
end;

procedure TCT_BaseStyles.Clear;
begin
  FAssigneds := [];
  if FA_ClrScheme <> Nil then 
    FreeAndNil(FA_ClrScheme);
  if FA_FontScheme <> Nil then 
    FreeAndNil(FA_FontScheme);
  if FA_FmtScheme <> Nil then 
    FreeAndNil(FA_FmtScheme);
  if FA_ExtLst <> Nil then 
    FreeAndNil(FA_ExtLst);
end;

procedure TCT_BaseStyles.Assign(AItem: TCT_BaseStyles);
begin
end;

procedure TCT_BaseStyles.CopyTo(AItem: TCT_BaseStyles);
begin
end;

function  TCT_BaseStyles.Create_A_ClrScheme: TCT_ColorScheme;
begin
  if FA_ClrScheme = Nil then
    FA_ClrScheme := TCT_ColorScheme.Create(FOwner);
  Result := FA_ClrScheme;
end;

function  TCT_BaseStyles.Create_A_FontScheme: TCT_FontScheme;
begin
  if FA_FontScheme = Nil then
    FA_FontScheme := TCT_FontScheme.Create(FOwner);
  Result := FA_FontScheme;
end;

function  TCT_BaseStyles.Create_A_FmtScheme: TCT_StyleMatrix;
begin
  if FA_FmtScheme = Nil then
    FA_FmtScheme := TCT_StyleMatrix.Create(FOwner);
  Result := FA_FmtScheme;
end;

function  TCT_BaseStyles.Create_A_ExtLst: TCT_OfficeArtExtensionList;
begin
  if FA_ExtLst = Nil then
    FA_ExtLst := TCT_OfficeArtExtensionList.Create(FOwner);
  Result := FA_ExtLst;
end;

{ TCT_PositiveFixedPercentage }

function  TCT_PositiveFixedPercentage.CheckAssigned: integer;
begin
  FAssigneds := [xaAttributes];
  Result := 1;
end;

procedure TCT_PositiveFixedPercentage.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_PositiveFixedPercentage.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('val',XmlIntToStr(FVal));
end;

procedure TCT_PositiveFixedPercentage.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'val' then 
    FVal := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_PositiveFixedPercentage.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
  FVal := 2147483632;
end;

destructor TCT_PositiveFixedPercentage.Destroy;
begin
end;

procedure TCT_PositiveFixedPercentage.Clear;
begin
  FAssigneds := [];
  FVal := 2147483632;
end;

procedure TCT_PositiveFixedPercentage.Assign(AItem: TCT_PositiveFixedPercentage);
begin
end;

procedure TCT_PositiveFixedPercentage.CopyTo(AItem: TCT_PositiveFixedPercentage);
begin
end;

{ TCT_InverseGammaTransform }

function  TCT_InverseGammaTransform.CheckAssigned: integer;
begin
  FAssigneds := [];
  Result := 0;
end;

procedure TCT_InverseGammaTransform.Write(AWriter: TXpgWriteXML);
begin
end;

constructor TCT_InverseGammaTransform.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 0;
end;

destructor TCT_InverseGammaTransform.Destroy;
begin
end;

procedure TCT_InverseGammaTransform.Clear;
begin
  FAssigneds := [];
end;

procedure TCT_InverseGammaTransform.Assign(AItem: TCT_InverseGammaTransform);
begin
end;

procedure TCT_InverseGammaTransform.CopyTo(AItem: TCT_InverseGammaTransform);
begin
end;

{ TCT_Angle }

function  TCT_Angle.CheckAssigned: integer;
begin
  FAssigneds := [xaAttributes];
  Result := 1;
end;

procedure TCT_Angle.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_Angle.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('val',XmlIntToStr(FVal));
end;

procedure TCT_Angle.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'val' then 
    FVal := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_Angle.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
  FVal := 2147483632;
end;

destructor TCT_Angle.Destroy;
begin
end;

procedure TCT_Angle.Clear;
begin
  FAssigneds := [];
  FVal := 2147483632;
end;

procedure TCT_Angle.Assign(AItem: TCT_Angle);
begin
end;

procedure TCT_Angle.CopyTo(AItem: TCT_Angle);
begin
end;

{ TCT_Percentage }

function  TCT_Percentage.CheckAssigned: integer;
begin
  FAssigneds := [xaAttributes];
  Result := 1;
end;

procedure TCT_Percentage.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_Percentage.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('val',XmlIntToStr(FVal));
end;

procedure TCT_Percentage.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'val' then 
    FVal := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_Percentage.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
  FVal := 2147483632;
end;

destructor TCT_Percentage.Destroy;
begin
end;

procedure TCT_Percentage.Clear;
begin
  FAssigneds := [];
  FVal := 2147483632;
end;

procedure TCT_Percentage.Assign(AItem: TCT_Percentage);
begin
end;

procedure TCT_Percentage.CopyTo(AItem: TCT_Percentage);
begin
end;

{ TCT_ComplementTransform }

function  TCT_ComplementTransform.CheckAssigned: integer;
begin
  FAssigneds := [];
  Result := 0;
end;

procedure TCT_ComplementTransform.Write(AWriter: TXpgWriteXML);
begin
end;

constructor TCT_ComplementTransform.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 0;
end;

destructor TCT_ComplementTransform.Destroy;
begin
end;

procedure TCT_ComplementTransform.Clear;
begin
  FAssigneds := [];
end;

procedure TCT_ComplementTransform.Assign(AItem: TCT_ComplementTransform);
begin
end;

procedure TCT_ComplementTransform.CopyTo(AItem: TCT_ComplementTransform);
begin
end;

{ TCT_PositivePercentage }

function  TCT_PositivePercentage.CheckAssigned: integer;
begin
  FAssigneds := [xaAttributes];
  Result := 1;
end;

procedure TCT_PositivePercentage.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_PositivePercentage.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('val',XmlIntToStr(FVal));
end;

procedure TCT_PositivePercentage.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'val' then 
    FVal := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_PositivePercentage.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
  FVal := 2147483632;
end;

destructor TCT_PositivePercentage.Destroy;
begin
end;

procedure TCT_PositivePercentage.Clear;
begin
  FAssigneds := [];
  FVal := 2147483632;
end;

procedure TCT_PositivePercentage.Assign(AItem: TCT_PositivePercentage);
begin
end;

procedure TCT_PositivePercentage.CopyTo(AItem: TCT_PositivePercentage);
begin
end;

{ TCT_GrayscaleTransform }

function  TCT_GrayscaleTransform.CheckAssigned: integer;
begin
  FAssigneds := [];
  Result := 0;
end;

procedure TCT_GrayscaleTransform.Write(AWriter: TXpgWriteXML);
begin
end;

constructor TCT_GrayscaleTransform.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 0;
end;

destructor TCT_GrayscaleTransform.Destroy;
begin
end;

procedure TCT_GrayscaleTransform.Clear;
begin
  FAssigneds := [];
end;

procedure TCT_GrayscaleTransform.Assign(AItem: TCT_GrayscaleTransform);
begin
end;

procedure TCT_GrayscaleTransform.CopyTo(AItem: TCT_GrayscaleTransform);
begin
end;

{ TCT_GammaTransform }

function  TCT_GammaTransform.CheckAssigned: integer;
begin
  FAssigneds := [];
  Result := 0;
end;

procedure TCT_GammaTransform.Write(AWriter: TXpgWriteXML);
begin
end;

constructor TCT_GammaTransform.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 0;
end;

destructor TCT_GammaTransform.Destroy;
begin
end;

procedure TCT_GammaTransform.Clear;
begin
  FAssigneds := [];
end;

procedure TCT_GammaTransform.Assign(AItem: TCT_GammaTransform);
begin
end;

procedure TCT_GammaTransform.CopyTo(AItem: TCT_GammaTransform);
begin
end;

{ TCT_FixedPercentage }

function  TCT_FixedPercentage.CheckAssigned: integer;
begin
  FAssigneds := [xaAttributes];
  Result := 1;
end;

procedure TCT_FixedPercentage.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_FixedPercentage.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('val',XmlIntToStr(FVal));
end;

procedure TCT_FixedPercentage.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'val' then 
    FVal := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_FixedPercentage.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
  FVal := 2147483632;
end;

destructor TCT_FixedPercentage.Destroy;
begin
end;

procedure TCT_FixedPercentage.Clear;
begin
  FAssigneds := [];
  FVal := 2147483632;
end;

procedure TCT_FixedPercentage.Assign(AItem: TCT_FixedPercentage);
begin
end;

procedure TCT_FixedPercentage.CopyTo(AItem: TCT_FixedPercentage);
begin
end;

{ TCT_PositiveFixedAngle }

function  TCT_PositiveFixedAngle.CheckAssigned: integer;
begin
  FAssigneds := [xaAttributes];
  Result := 1;
end;

procedure TCT_PositiveFixedAngle.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_PositiveFixedAngle.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('val',XmlIntToStr(FVal));
end;

procedure TCT_PositiveFixedAngle.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'val' then 
    FVal := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_PositiveFixedAngle.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
  FVal := 2147483632;
end;

destructor TCT_PositiveFixedAngle.Destroy;
begin
end;

procedure TCT_PositiveFixedAngle.Clear;
begin
  FAssigneds := [];
  FVal := 2147483632;
end;

procedure TCT_PositiveFixedAngle.Assign(AItem: TCT_PositiveFixedAngle);
begin
end;

procedure TCT_PositiveFixedAngle.CopyTo(AItem: TCT_PositiveFixedAngle);
begin
end;

{ TCT_InverseTransform }

function  TCT_InverseTransform.CheckAssigned: integer;
begin
  FAssigneds := [];
  Result := 0;
end;

procedure TCT_InverseTransform.Write(AWriter: TXpgWriteXML);
begin
end;

constructor TCT_InverseTransform.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 0;
end;

destructor TCT_InverseTransform.Destroy;
begin
end;

procedure TCT_InverseTransform.Clear;
begin
  FAssigneds := [];
end;

procedure TCT_InverseTransform.Assign(AItem: TCT_InverseTransform);
begin
end;

procedure TCT_InverseTransform.CopyTo(AItem: TCT_InverseTransform);
begin
end;

{ TCT_NonVisualDrawingProps }

function  TCT_NonVisualDrawingProps.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [xaAttributes];
  if FA_HlinkClick <> Nil then
    Inc(ElemsAssigned,FA_HlinkClick.CheckAssigned);
  if FA_HlinkHover <> Nil then
    Inc(ElemsAssigned,FA_HlinkHover.CheckAssigned);
  if FA_ExtLst <> Nil then
    Inc(ElemsAssigned,FA_ExtLst.CheckAssigned);
  Result := 1;
  if ElemsAssigned > 0 then
    FAssigneds := FAssigneds + [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_NonVisualDrawingProps.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $00000497: begin
      if FA_HlinkClick = Nil then
        FA_HlinkClick := TCT_Hyperlink.Create(FOwner);
      Result := FA_HlinkClick;
    end;
    $000004B5: begin
      if FA_HlinkHover = Nil then
        FA_HlinkHover := TCT_Hyperlink.Create(FOwner);
      Result := FA_HlinkHover;
    end;
    $0000031F: begin
      if FA_ExtLst = Nil then
        FA_ExtLst := TCT_OfficeArtExtensionList.Create(FOwner);
      Result := FA_ExtLst;
    end;
    else
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then
    Result.Assigneds := [xaRead];
end;

procedure TCT_NonVisualDrawingProps.Write(AWriter: TXpgWriteXML);
begin
  if (FA_HlinkClick <> Nil) and FA_HlinkClick.Assigned then
  begin
    FA_HlinkClick.WriteAttributes(AWriter);
    if xaElements in FA_HlinkClick.FAssigneds then
    begin
      AWriter.BeginTag('a:hlinkClick');
      FA_HlinkClick.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('a:hlinkClick');
  end;
  if (FA_HlinkHover <> Nil) and FA_HlinkHover.Assigned then
  begin
    FA_HlinkHover.WriteAttributes(AWriter);
    if xaElements in FA_HlinkHover.FAssigneds then
    begin
      AWriter.BeginTag('a:hlinkHover');
      FA_HlinkHover.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('a:hlinkHover');
  end;
  if (FA_ExtLst <> Nil) and FA_ExtLst.Assigned then
    if xaElements in FA_ExtLst.FAssigneds then
    begin
      AWriter.BeginTag('a:extLst');
      FA_ExtLst.Write(AWriter);
      AWriter.EndTag;
    end;
end;

procedure TCT_NonVisualDrawingProps.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('id',XmlIntToStr(FId));
  AWriter.AddAttribute('name',FName);
  if FDescr <> '' then
    AWriter.AddAttribute('descr',FDescr);
  if FHidden <> False then
    AWriter.AddAttribute('hidden',XmlBoolToStr(FHidden));
end;

procedure TCT_NonVisualDrawingProps.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
    case AAttributes.HashA[i] of
      $000000CD: FId := XmlStrToIntDef(AAttributes.Values[i],0);
      $000001A1: FName := AAttributes.Values[i];
      $00000211: FDescr := AAttributes.Values[i];
      $0000026C: FHidden := XmlStrToBoolDef(AAttributes.Values[i],False);
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_NonVisualDrawingProps.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 3;
  FAttributeCount := 4;
  FId := 2147483632;
  FHidden := False;
end;

destructor TCT_NonVisualDrawingProps.Destroy;
begin
  if FA_HlinkClick <> Nil then
    FA_HlinkClick.Free;
  if FA_HlinkHover <> Nil then
    FA_HlinkHover.Free;
  if FA_ExtLst <> Nil then
    FA_ExtLst.Free;
end;

procedure TCT_NonVisualDrawingProps.Clear;
begin
  FAssigneds := [];
  if FA_HlinkClick <> Nil then
    FreeAndNil(FA_HlinkClick);
  if FA_HlinkHover <> Nil then
    FreeAndNil(FA_HlinkHover);
  if FA_ExtLst <> Nil then
    FreeAndNil(FA_ExtLst);
  FId := 2147483632;
  FName := '';
  FDescr := '';
  FHidden := False;
end;

procedure TCT_NonVisualDrawingProps.Assign(AItem: TCT_NonVisualDrawingProps);
begin
end;

procedure TCT_NonVisualDrawingProps.CopyTo(AItem: TCT_NonVisualDrawingProps);
begin
end;

function  TCT_NonVisualDrawingProps.Create_A_HlinkClick: TCT_Hyperlink;
begin
  if FA_HlinkClick = Nil then
    FA_HlinkClick := TCT_Hyperlink.Create(FOwner);
  Result := FA_HlinkClick;
end;

function  TCT_NonVisualDrawingProps.Create_A_HlinkHover: TCT_Hyperlink;
begin
  if FA_HlinkHover = Nil then
    FA_HlinkHover := TCT_Hyperlink.Create(FOwner);
  Result := FA_HlinkHover;
end;

function  TCT_NonVisualDrawingProps.Create_A_ExtLst: TCT_OfficeArtExtensionList;
begin
  if FA_ExtLst = Nil then
    FA_ExtLst := TCT_OfficeArtExtensionList.Create(FOwner);
  Result := FA_ExtLst;
end;

{ TCT_NonVisualDrawingShapeProps }

function  TCT_NonVisualDrawingShapeProps.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FTxBox <> False then
    Inc(AttrsAssigned);
  if FA_SpLocks <> Nil then
    Inc(ElemsAssigned);
  if FA_ExtLst <> Nil then
    Inc(ElemsAssigned,FA_ExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_NonVisualDrawingShapeProps.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $0000037A: begin
      if FA_SpLocks = Nil then
        FA_SpLocks := TCT_ShapeLocking.Create(FOwner);
      Result := FA_SpLocks;
    end;
    $0000031F: begin
      if FA_ExtLst = Nil then
        FA_ExtLst := TCT_OfficeArtExtensionList.Create(FOwner);
      Result := FA_ExtLst;
    end;
    else
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then
    Result.Assigneds := [xaRead];
end;

procedure TCT_NonVisualDrawingShapeProps.Write(AWriter: TXpgWriteXML);
begin
  if (FA_SpLocks <> Nil) and FA_SpLocks.Assigned then
  begin
    FA_SpLocks.WriteAttributes(AWriter);
    if xaElements in FA_SpLocks.FAssigneds then
    begin
      AWriter.BeginTag('a:spLocks');
      FA_SpLocks.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('a:spLocks');
  end;
  if (FA_ExtLst <> Nil) and FA_ExtLst.Assigned then
    if xaElements in FA_ExtLst.FAssigneds then
    begin
      AWriter.BeginTag('a:extLst');
      FA_ExtLst.Write(AWriter);
      AWriter.EndTag;
    end;
end;

procedure TCT_NonVisualDrawingShapeProps.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FTxBox <> False then
    AWriter.AddAttribute('txBox',XmlBoolToStr(FTxBox));
end;

procedure TCT_NonVisualDrawingShapeProps.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'txBox' then
    FTxBox := XmlStrToBoolDef(AAttributes.Values[0],False)
  else
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_NonVisualDrawingShapeProps.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 2;
  FAttributeCount := 1;
  FTxBox := False;
end;

destructor TCT_NonVisualDrawingShapeProps.Destroy;
begin
  if FA_SpLocks <> Nil then
    FA_SpLocks.Free;
  if FA_ExtLst <> Nil then
    FA_ExtLst.Free;
end;

procedure TCT_NonVisualDrawingShapeProps.Clear;
begin
  FAssigneds := [];
  if FA_SpLocks <> Nil then
    FreeAndNil(FA_SpLocks);
  if FA_ExtLst <> Nil then
    FreeAndNil(FA_ExtLst);
  FTxBox := False;
end;

procedure TCT_NonVisualDrawingShapeProps.Assign(AItem: TCT_NonVisualDrawingShapeProps);
begin
end;

procedure TCT_NonVisualDrawingShapeProps.CopyTo(AItem: TCT_NonVisualDrawingShapeProps);
begin
end;

function  TCT_NonVisualDrawingShapeProps.Create_A_SpLocks: TCT_ShapeLocking;
begin
  if FA_SpLocks = Nil then
    FA_SpLocks := TCT_ShapeLocking.Create(FOwner);
  Result := FA_SpLocks;
end;

function  TCT_NonVisualDrawingShapeProps.Create_A_ExtLst: TCT_OfficeArtExtensionList;
begin
  if FA_ExtLst = Nil then
    FA_ExtLst := TCT_OfficeArtExtensionList.Create(FOwner);
  Result := FA_ExtLst;
end;

{ TCT_ShapeLocking }

function  TCT_ShapeLocking.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FNoGrp <> False then
    Inc(AttrsAssigned);
  if FNoSelect <> False then
    Inc(AttrsAssigned);
  if FNoRot <> False then
    Inc(AttrsAssigned);
  if FNoChangeAspect <> False then
    Inc(AttrsAssigned);
  if FNoMove <> False then
    Inc(AttrsAssigned);
  if FNoResize <> False then
    Inc(AttrsAssigned);
  if FNoEditPoints <> False then
    Inc(AttrsAssigned);
  if FNoAdjustHandles <> False then
    Inc(AttrsAssigned);
  if FNoChangeArrowheads <> False then
    Inc(AttrsAssigned);
  if FNoChangeShapeType <> False then
    Inc(AttrsAssigned);
  if FNoTextEdit <> False then
    Inc(AttrsAssigned);
  if FA_ExtLst <> Nil then
    Inc(ElemsAssigned,FA_ExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_ShapeLocking.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'a:extLst' then
  begin
    if FA_ExtLst = Nil then
      FA_ExtLst := TCT_OfficeArtExtensionList.Create(FOwner);
    Result := FA_ExtLst;
  end
  else
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then
    Result.Assigneds := [xaRead];
end;

procedure TCT_ShapeLocking.Write(AWriter: TXpgWriteXML);
begin
  if (FA_ExtLst <> Nil) and FA_ExtLst.Assigned then
    if xaElements in FA_ExtLst.FAssigneds then
    begin
      AWriter.BeginTag('a:extLst');
      FA_ExtLst.Write(AWriter);
      AWriter.EndTag;
    end;
end;

procedure TCT_ShapeLocking.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FNoGrp <> False then
    AWriter.AddAttribute('noGrp',XmlBoolToStr(FNoGrp));
  if FNoSelect <> False then
    AWriter.AddAttribute('noSelect',XmlBoolToStr(FNoSelect));
  if FNoRot <> False then
    AWriter.AddAttribute('noRot',XmlBoolToStr(FNoRot));
  if FNoChangeAspect <> False then
    AWriter.AddAttribute('noChangeAspect',XmlBoolToStr(FNoChangeAspect));
  if FNoMove <> False then
    AWriter.AddAttribute('noMove',XmlBoolToStr(FNoMove));
  if FNoResize <> False then
    AWriter.AddAttribute('noResize',XmlBoolToStr(FNoResize));
  if FNoEditPoints <> False then
    AWriter.AddAttribute('noEditPoints',XmlBoolToStr(FNoEditPoints));
  if FNoAdjustHandles <> False then
    AWriter.AddAttribute('noAdjustHandles',XmlBoolToStr(FNoAdjustHandles));
  if FNoChangeArrowheads <> False then
    AWriter.AddAttribute('noChangeArrowheads',XmlBoolToStr(FNoChangeArrowheads));
  if FNoChangeShapeType <> False then
    AWriter.AddAttribute('noChangeShapeType',XmlBoolToStr(FNoChangeShapeType));
  if FNoTextEdit <> False then
    AWriter.AddAttribute('noTextEdit',XmlBoolToStr(FNoTextEdit));
end;

procedure TCT_ShapeLocking.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
    case AAttributes.HashA[i] of
      $00000206: FNoGrp := XmlStrToBoolDef(AAttributes.Values[i],False);
      $0000033D: FNoSelect := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000212: FNoRot := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000583: FNoChangeAspect := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000274: FNoMove := XmlStrToBoolDef(AAttributes.Values[i],False);
      $0000034F: FNoResize := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000004E0: FNoEditPoints := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000607: FNoAdjustHandles := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000733: FNoChangeArrowheads := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000006B6: FNoChangeShapeType := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000408: FNoTextEdit := XmlStrToBoolDef(AAttributes.Values[i],False);
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_ShapeLocking.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 11;
  FNoGrp := False;
  FNoSelect := False;
  FNoRot := False;
  FNoChangeAspect := False;
  FNoMove := False;
  FNoResize := False;
  FNoEditPoints := False;
  FNoAdjustHandles := False;
  FNoChangeArrowheads := False;
  FNoChangeShapeType := False;
  FNoTextEdit := False;
end;

destructor TCT_ShapeLocking.Destroy;
begin
  if FA_ExtLst <> Nil then
    FA_ExtLst.Free;
end;

procedure TCT_ShapeLocking.Clear;
begin
  FAssigneds := [];
  if FA_ExtLst <> Nil then
    FreeAndNil(FA_ExtLst);
  FNoGrp := False;
  FNoSelect := False;
  FNoRot := False;
  FNoChangeAspect := False;
  FNoMove := False;
  FNoResize := False;
  FNoEditPoints := False;
  FNoAdjustHandles := False;
  FNoChangeArrowheads := False;
  FNoChangeShapeType := False;
  FNoTextEdit := False;
end;

procedure TCT_ShapeLocking.Assign(AItem: TCT_ShapeLocking);
begin
end;

procedure TCT_ShapeLocking.CopyTo(AItem: TCT_ShapeLocking);
begin
end;

function  TCT_ShapeLocking.Create_A_ExtLst: TCT_OfficeArtExtensionList;
begin
  if FA_ExtLst = Nil then
    FA_ExtLst := TCT_OfficeArtExtensionList.Create(FOwner);
  Result := FA_ExtLst;
end;

{ TCT_ShapeNonVisual }

function  TCT_ShapeNonVisual.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FCNvPr <> Nil then
    Inc(ElemsAssigned,FCNvPr.CheckAssigned);
  if FCNvSpPr <> Nil then
    Inc(ElemsAssigned,FCNvSpPr.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_ShapeNonVisual.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;

  if AReader.QName = FOwner.NS + ':cNvPr' then begin
    if FCNvPr = Nil then
      FCNvPr := TCT_NonVisualDrawingProps.Create(FOwner);
    Result := FCNvPr;
  end
  else if AReader.QName = FOwner.NS + ':cNvSpPr' then begin
      if FCNvSpPr = Nil then
        FCNvSpPr := TCT_NonVisualDrawingShapeProps.Create(FOwner);
      Result := FCNvSpPr;
  end
  else
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);

  if Result <> Self then
    Result.Assigneds := [xaRead];
end;

procedure TCT_ShapeNonVisual.Write(AWriter: TXpgWriteXML);
begin
  if (FCNvPr <> Nil) and FCNvPr.Assigned then
  begin
    FCNvPr.WriteAttributes(AWriter);
    if xaElements in FCNvPr.FAssigneds then
    begin
      AWriter.BeginTag(FOwner.NS + ':cNvPr');
      FCNvPr.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag(FOwner.NS + ':cNvPr');
  end;
  if (FCNvSpPr <> Nil) and FCNvSpPr.Assigned then begin
    FCNvSpPr.WriteAttributes(AWriter);
    if xaElements in FCNvSpPr.FAssigneds then begin
      AWriter.BeginTag(FOwner.NS + ':cNvSpPr');
      FCNvSpPr.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag(FOwner.NS + ':cNvSpPr');
  end
  else
    AWriter.SimpleTag(FOwner.NS + ':cNvSpPr');  // TODO Must be written
end;

constructor TCT_ShapeNonVisual.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 2;
  FAttributeCount := 0;
end;

destructor TCT_ShapeNonVisual.Destroy;
begin
  if FCNvPr <> Nil then
    FCNvPr.Free;
  if FCNvSpPr <> Nil then
    FCNvSpPr.Free;
end;

procedure TCT_ShapeNonVisual.Clear;
begin
  FAssigneds := [];
  if FCNvPr <> Nil then
    FreeAndNil(FCNvPr);
  if FCNvSpPr <> Nil then
    FreeAndNil(FCNvSpPr);
end;

procedure TCT_ShapeNonVisual.Assign(AItem: TCT_ShapeNonVisual);
begin
end;

procedure TCT_ShapeNonVisual.CopyTo(AItem: TCT_ShapeNonVisual);
begin
end;

function  TCT_ShapeNonVisual.Create_CNvPr: TCT_NonVisualDrawingProps;
begin
  if FCNvPr = Nil then
    FCNvPr := TCT_NonVisualDrawingProps.Create(FOwner);
  Result := FCNvPr;
end;

function  TCT_ShapeNonVisual.Create_CNvSpPr: TCT_NonVisualDrawingShapeProps;
begin
  if FCNvSpPr = Nil then
    FCNvSpPr := TCT_NonVisualDrawingShapeProps.Create(FOwner);
  Result := FCNvSpPr;
end;

{ TEG_TextRunXpgList }

function TEG_TextRunXpgList.Add: TEG_TextRun;
begin
  Result := TEG_TextRun.Create(FOwner);
  inherited Add(Result);
end;

procedure TEG_TextRunXpgList.Assign(AItem: TEG_TextRunXpgList);
begin

end;

function TEG_TextRunXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TEG_TextRunXpgList.CopyTo(AItem: TEG_TextRunXpgList);
begin

end;

function TEG_TextRunXpgList.GetItems(Index: integer): TEG_TextRun;
begin
  Result := TEG_TextRun(inherited Items[Index]);
end;

procedure TEG_TextRunXpgList.Write(AWriter: TXpgWriteXML);
var
  i: integer;
begin
  for i := 0 to Count - 1 do
    GetItems(i).Write(AWriter);
end;

{ TXPSAlternateContent }

procedure TXPSAlternateContent.BeginHandle(ACurrent: TXPGBase);
begin
  FCurrent := ACurrent;
  FInFallback := False;
end;

function TXPSAlternateContent.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  if FInFallback then
    Result := FCurrent.HandleElement(AReader)
  else begin
    FInFallback := AReader.Tag = 'Fallback';
    Result := Self;
  end;
end;

{ TXPGBaseList }

function TXPGBaseList.Add(AItem: TXPGBase): TXPGBase;
begin
  Result := AItem;
  inherited Add(AItem);
end;

procedure TXPGBaseList.CheckAssigned;
var
  i: integer;
begin
  for i := 0 to Count - 1 do
    Items[i].CheckAssigned;
end;

constructor TXPGBaseList.Create;
begin
  inherited Create;
end;

function TXPGBaseList.GetItems(Index: integer): TXPGBase;
begin
  Result := TXPGBase(inherited Items[Index]);
end;

{ TXPGDocBase }

constructor TXPGDocBase.Create;
begin
  FNS := 'xdr';
end;

constructor TXPGDocBase.Create(AGrManager: TXc12GraphicManager);
begin
  FGrManager := AGrManager;
end;

initialization
{$ifdef DELPHI_5}
  L_Enums := TStringList.Create;
{$else}
  L_Enums := THashedStringList.Create;
{$endif}
  AddEnums;

finalization
  L_Enums.Free;

end.
