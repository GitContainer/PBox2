unit BIFF_EscherTypes5;

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

{$ifdef ver130}
uses Windows;
{$else}
uses Types;
{$endif}

type TMSOShapeType = (mstUnknown,mstCtrlLabel,mstCtrlButton,mstCtrlListBox,
                      mstCtrlCheckBox,mstCtrlComboBox,mstCtrlRadioButton,
                      mstCtrlGroupBox,mstCtrlScrollBar,mstTextBox,mstTextShape,
                      mstNote,mstChart);

const MSO_ALIGNRULE            = $F013;
const MSO_ANCHOR               = $F00E;
const MSO_ARCRULE              = $F014;
const MSO_BLIPEND              = $F117;
const MSO_BLIPSTART            = $F018;
const MSO_BSE                  = $F007;
const MSO_BSTORECONTAINER      = $F001;
const MSO_CALLOUTRULE          = $F017;
const MSO_CHILDANCHOR          = $F00F;
const MSO_CLIENTANCHOR         = $F010;
const MSO_CLIENTDATA           = $F011;
const MSO_CLIENTRULE           = $F015;
const MSO_CLIENTTEXTBOX        = $F00D;
const MSO_CLSID                = $F016;
const MSO_COLORSCHEME          = $F120;
const MSO_COLORMRU             = $F11A;
const MSO_CONNECTORRULE        = $F012;
const MSO_DELETEDPSPL          = $F11D;
const MSO_DG                   = $F008;
const MSO_DGCONTAINER          = $F002;
const MSO_DGG                  = $F006;
const MSO_DGGCONTAINER         = $F000;
const MSO_OLEOBJECT            = $F11F;
const MSO_OPT                  = $F00B;
const MSO_REGROUPITEMS         = $F118;
const MSO_SELECTION            = $F119;
const MSO_SOLVERCONTAINER      = $F005;
const MSO_SP                   = $F00A;
const MSO_SPCONTAINER          = $F004;
const MSO_SPGR                 = $F009;
const MSO_SPGRCONTAINER        = $F003;
const MSO_SPLITMENUCOLORS      = $F11E;
const MSO_TEXTBOX              = $F00C;
const MSO_TERTIARY_FOPT        = $F122;

const BLIP_DIB                 = $07A8;
const BLIP_JPEG                = $046A;
const BLIP_PNG                 = $06E0;
const BLIP_EMF                 = $03D4;
const BLIP_WMF                 = $0216;
const BLIP_PICT                = $0542;

const msosptMin                      = 0;
const msosptNotPrimitive             = msosptMin;
const msosptRectangle                = 1;
const msosptRoundRectangle           = 2;
const msosptEllipse                  = 3;
const msosptDiamond                  = 4;
const msosptIsocelesTriangle         = 5;
const msosptRightTriangle            = 6;
const msosptParallelogram            = 7;
const msosptTrapezoid                = 8;
const msosptHexagon                  = 9;
const msosptOctagon                  = 10;
const msosptPlus                     = 11;
const msosptStar                     = 12;
const msosptArrow                    = 13;
const msosptThickArrow               = 14;
const msosptHomePlate                = 15;
const msosptCube                     = 16;
const msosptBalloon                  = 17;
const msosptSeal                     = 18;
const msosptArc                      = 19;
const msosptLine                     = 20;
const msosptPlaque                   = 21;
const msosptCan                      = 22;
const msosptDonut                    = 23;
const msosptTextSimple               = 24;
const msosptTextOctagon              = 25;
const msosptTextHexagon              = 26;
const msosptTextCurve                = 27;
const msosptTextWave                 = 28;
const msosptTextRing                 = 29;
const msosptTextOnCurve              = 30;
const msosptTextOnRing               = 31;
const msosptStraightConnector1       = 32;
const msosptBentConnector2           = 33;
const msosptBentConnector3           = 34;
const msosptBentConnector4           = 35;
const msosptBentConnector5           = 36;
const msosptCurvedConnector2         = 37;
const msosptCurvedConnector3         = 38;
const msosptCurvedConnector4         = 39;
const msosptCurvedConnector5         = 40;
const msosptCallout1                 = 41;
const msosptCallout2                 = 42;
const msosptCallout3                 = 43;
const msosptAccentCallout1           = 44;
const msosptAccentCallout2           = 45;
const msosptAccentCallout3           = 46;
const msosptBorderCallout1           = 47;
const msosptBorderCallout2           = 48;
const msosptBorderCallout3           = 49;
const msosptAccentBorderCallout1     = 50;
const msosptAccentBorderCallout2     = 51;
const msosptAccentBorderCallout3     = 52;
const msosptRibbon                   = 53;
const msosptRibbon2                  = 54;
const msosptChevron                  = 55;
const msosptPentagon                 = 56;
const msosptNoSmoking                = 57;
const msosptSeal8                    = 58;
const msosptSeal16                   = 59;
const msosptSeal32                   = 60;
const msosptWedgeRectCallout         = 61;
const msosptWedgeRRectCallout        = 62;
const msosptWedgeEllipseCallout      = 63;
const msosptWave                     = 64;
const msosptFoldedCorner             = 65;
const msosptLeftArrow                = 66;
const msosptDownArrow                = 67;
const msosptUpArrow                  = 68;
const msosptLeftRightArrow           = 69;
const msosptUpDownArrow              = 70;
const msosptIrregularSeal1           = 71;
const msosptIrregularSeal2           = 72;
const msosptLightningBolt            = 73;
const msosptHeart                    = 74;
const msosptPictureFrame             = 75;
const msosptQuadArrow                = 76;
const msosptLeftArrowCallout         = 77;
const msosptRightArrowCallout        = 78;
const msosptUpArrowCallout           = 79;
const msosptDownArrowCallout         = 80;
const msosptLeftRightArrowCallout    = 81;
const msosptUpDownArrowCallout       = 82;
const msosptQuadArrowCallout         = 83;
const msosptBevel                    = 84;
const msosptLeftBracket              = 85;
const msosptRightBracket             = 86;
const msosptLeftBrace                = 87;
const msosptRightBrace               = 88;
const msosptLeftUpArrow              = 89;
const msosptBentUpArrow              = 90;
const msosptBentArrow                = 91;
const msosptSeal24                   = 92;
const msosptStripedRightArrow        = 93;
const msosptNotchedRightArrow        = 94;
const msosptBlockArc                 = 95;
const msosptSmileyFace               = 96;
const msosptVerticalScroll           = 97;
const msosptHorizontalScroll         = 98;
const msosptCircularArrow            = 99;
const msosptNotchedCircularArrow     = 100;
const msosptUturnArrow               = 101;
const msosptCurvedRightArrow         = 102;
const msosptCurvedLeftArrow          = 103;
const msosptCurvedUpArrow            = 104;
const msosptCurvedDownArrow          = 105;
const msosptCloudCallout             = 106;
const msosptEllipseRibbon            = 107;
const msosptEllipseRibbon2           = 108;
const msosptFlowChartProcess         = 109;
const msosptFlowChartDecision        = 110;
const msosptFlowChartInputOutput     = 111;
const msosptFlowChartPredefinedProcess= 112;
const msosptFlowChartInternalStorage = 113;
const msosptFlowChartDocument        = 114;
const msosptFlowChartMultidocument   = 115;
const msosptFlowChartTerminator      = 116;
const msosptFlowChartPreparation     = 117;
const msosptFlowChartManualInput     = 118;
const msosptFlowChartManualOperation = 119;
const msosptFlowChartConnector       = 120;
const msosptFlowChartPunchedCard     = 121;
const msosptFlowChartPunchedTape     = 122;
const msosptFlowChartSummingJunction = 123;
const msosptFlowChartOr              = 124;
const msosptFlowChartCollate         = 125;
const msosptFlowChartSort            = 126;
const msosptFlowChartExtract         = 127;
const msosptFlowChartMerge           = 128;
const msosptFlowChartOfflineStorage  = 129;
const msosptFlowChartOnlineStorage   = 130;
const msosptFlowChartMagneticTape    = 131;
const msosptFlowChartMagneticDisk    = 132;
const msosptFlowChartMagneticDrum    = 133;
const msosptFlowChartDisplay         = 134;
const msosptFlowChartDelay           = 135;
const msosptTextPlainText            = 136;
const msosptTextStop                 = 137;
const msosptTextTriangle             = 138;
const msosptTextTriangleInverted     = 139;
const msosptTextChevron              = 140;
const msosptTextChevronInverted      = 141;
const msosptTextRingInside           = 142;
const msosptTextRingOutside          = 143;
const msosptTextArchUpCurve          = 144;
const msosptTextArchDownCurve        = 145;
const msosptTextCircleCurve          = 146;
const msosptTextButtonCurve          = 147;
const msosptTextArchUpPour           = 148;
const msosptTextArchDownPour         = 149;
const msosptTextCirclePour           = 150;
const msosptTextButtonPour           = 151;
const msosptTextCurveUp              = 152;
const msosptTextCurveDown            = 153;
const msosptTextCascadeUp            = 154;
const msosptTextCascadeDown          = 155;
const msosptTextWave1                = 156;
const msosptTextWave2                = 157;
const msosptTextWave3                = 158;
const msosptTextWave4                = 159;
const msosptTextInflate              = 160;
const msosptTextDeflate              = 161;
const msosptTextInflateBottom        = 162;
const msosptTextDeflateBottom        = 163;
const msosptTextInflateTop           = 164;
const msosptTextDeflateTop           = 165;
const msosptTextDeflateInflate       = 166;
const msosptTextDeflateInflateDeflate= 167;
const msosptTextFadeRight            = 168;
const msosptTextFadeLeft             = 169;
const msosptTextFadeUp               = 170;
const msosptTextFadeDown             = 171;
const msosptTextSlantUp              = 172;
const msosptTextSlantDown            = 173;
const msosptTextCanUp                = 174;
const msosptTextCanDown              = 175;
const msosptFlowChartAlternateProcess= 176;
const msosptFlowChartOffpageConnector= 177;
const msosptCallout90                = 178;
const msosptAccentCallout90          = 179;
const msosptBorderCallout90          = 180;
const msosptAccentBorderCallout90    = 181;
const msosptLeftRightUpArrow         = 182;
const msosptSun                      = 183;
const msosptMoon                     = 184;
const msosptBracketPair              = 185;
const msosptBracePair                = 186;
const msosptSeal4                    = 187;
const msosptDoubleWave               = 188;
const msosptActionButtonBlank        = 189;
const msosptActionButtonHome         = 190;
const msosptActionButtonHelp         = 191;
const msosptActionButtonInformation  = 192;
const msosptActionButtonForwardNext  = 193;
const msosptActionButtonBackPrevious = 194;
const msosptActionButtonEnd          = 195;
const msosptActionButtonBeginning    = 196;
const msosptActionButtonReturn       = 197;
const msosptActionButtonDocument     = 198;
const msosptActionButtonSound        = 199;
const msosptActionButtonMovie        = 200;
const msosptHostControl              = 201;
const msosptTextBox                  = 202;
const SpOptGroup      = $00000001;
const SpOptChild      = $00000002;
const SpOptPatriarch  = $00000004;
const SpOptDeleted    = $00000008;
const SpOptOleShape   = $00000010;
const SpOptHaveMaster = $00000020;
const SpOptFlipH      = $00000040;
const SpOptFlipV      = $00000080;
const SpOptConnector  = $00000100;
const SpOptHaveAnchor = $00000200;
const SpOptBackground = $00000400;
const SpOptHaveSpt    = $00000800;

const msooptRotation = 4;
const msooptFLockRotation = 119;
const msooptFLockAspectRatio = 120;
const msooptFLockPosition = 121;
const msooptFLockAgainstSelect = 122;
const msooptFLockCropping = 123;
const msooptLockVertices = 124;
const msooptFLockText = 125;
const msooptFLockAdjustHandles = 126;
const msooptFLockAgainstGrouping = 127;
const msooptLTxid = 128;
const msooptDxTextLeft = 129;
const msooptDyTextTop = 130;
const msooptDxTextRight = 131;
const msooptDyTextBottom = 132;
const msooptWrapText = 133;
const msooptScaleText = 134;
const msooptAnchorText = 135;
const msooptXflTextFlow = 136;
const msooptCdirFont = 137;
const msooptHspNext = 138;
const msooptTxdir = 139;
const msooptFSelectText = 187;
const msooptFAutoTextMargin = 188;
const msooptFRotateText = 189;
const msooptFFitShapeToText = 190;
const msooptFFitTextToShape = 191;
const msooptGtextUNICODE = 192;
const msooptTextRTF = 193;
const msooptGtextAlign = 194;
const msooptTextSize = 195;
const msooptGtextSpacing = 196;
const msooptGtextFont = 197;
const msooptGtextFReverseRows = 240;
const msooptFGtext = 241;
const msooptGtextFVertical = 242;
const msooptGtextFKern = 243;
const msooptGtextFTight = 244;
const msooptGtextFStretch = 245;
const msooptGtextFShrinkFit = 246;
const msooptGtextFBestFit = 247;
const msooptGtextFNormalize = 248;
const msooptGtextFDxMeasure = 249;
const msooptGtextFBold = 250;
const msooptGtextFItalic = 251;
const msooptGtextFUnderline = 252;
const msooptGtextFShadow = 253;
const msooptGtextFSmallcaps = 254;
const msooptGtextFStrikethrough = 255;
const msooptCropFromTop = 256;
const msooptRopFromBottom = 257;
const msooptRopFromLeft = 258;
const msooptCropFromRight = 259;
const msooptPib = 260;
const msooptPibName = 261;
const msooptPibFlags = 262;
const msooptPictureTransparent = 263;
const msooptPictureContrast = 264;
const msooptPictureBrightness = 265;
const msooptPictureGamma = 266;
const msooptPictureId = 267;
const msooptPictureDblCrMod = 268;
const msooptPictureFillCrMod = 269;
const msooptPictureLineCrMod = 270;
const msooptPibPrint = 271;
const msooptPibPrintName = 272;
const msooptPibPrintFlags = 273;
const msooptFNoHitTestPicture = 316;
const msooptPictureGray = 317;
const msooptPictureBiLevel = 318;
const msooptPictureActive = 319;
const msooptFillType = 384;
const msooptFillColor = 385;
const msooptFillOpacity = 386;
const msooptFillBackColor = 387;
const msooptFillBackOpacity = 388;
const msooptFillCrMod = 389;
const msooptFillBlip = 390;
const msooptFillBlipName = 391;
const msooptFillBlipFlags = 392;
const msooptFillWidth = 393;
const msooptFillHeight = 394;
const msooptFillAngle = 395;
const msooptFillFocus = 396;
const msooptFillToLeft = 397;
const msooptFillToTop = 398;
const msooptFillToRight = 399;
const msooptFillToBottom = 400;
const msooptFillRectLeft = 401;
const msooptFillRectTop = 402;
const msooptFillRectRight = 403;
const msooptFillRectBottom = 404;
const msooptFillDztype = 405;
const msooptFillShadePreset = 406;
const msooptFillShadeColors = 407;
const msooptFillOriginX = 408;
const msooptFillOriginY = 409;
const msooptFillShapeOriginX = 410;
const msooptFillShapeOriginY = 411;
const msooptFillShadeType = 412;
const msooptFFilled = 443;
const msooptFHitTestFill = 444;
const msooptFillShape = 445;
const msooptFillUseRect = 446;
const msooptFNoFillHitTest = 447;
const msooptLineColor = 448;
const msooptLineOpacity = 449;
const msooptLineBackColor = 450;
const msooptLineCrMod = 451;
const msooptLineType = 452;
const msooptLineFillBlip = 453;
const msooptLineFillBlipName = 454;
const msooptLineFillBlipFlags = 455;
const msooptLineFillWidth = 456;
const msooptLineFillHeight = 457;
const msooptLineFillDztype = 458;
const msooptLineWidth = 459;
const msooptLineMiterLimit = 460;
const msooptLineStyle = 461;
const msooptLineDashing = 462;
const msooptLineDashStyle = 463;
const msooptLineStartArrowhead = 464;
const msooptLineEndArrowhead = 465;
const msooptLineStartArrowWidth = 466;
const msooptLineStartArrowLength = 467;
const msooptLineEndArrowWidth = 468;
const msooptLineEndArrowLength = 469;
const msooptLineJoinStyle = 470;
const msooptLineEndCapStyle = 471;
const msooptFArrowheadsOK = 507;
const msooptLine = 508;
const msooptFHitTestLine = 509;
const msooptLineFillShape = 510;
const msooptFNoLineDrawDash = 511;
const msooptShadowType = 512;
const msooptShadowColor = 513;
const msooptShadowHighlight = 514;
const msooptShadowCrMod = 515;
const msooptShadowOpacity = 516;
const msooptShadowOffsetX = 517;
const msooptShadowOffsetY = 518;
const msooptShadowSecondOffsetX = 519;
const msooptShadowSecondOffsetY = 520;
const msooptShadowScaleXToX = 521;
const msooptShadowScaleYToX = 522;
const msooptShadowScaleXToY = 523;
const msooptShadowScaleYToY = 524;
const msooptShadowPerspectiveX = 525;
const msooptShadowPerspectiveY = 526;
const msooptShadowWeight = 527;
const msooptShadowOriginX = 528;
const msooptShadowOriginY = 529;
const msooptFShadow = 574;
const msooptShadowObscured = 575;
const msooptPerspectiveType = 576;
const msooptPerspectiveOffsetX = 577;
const msooptPerspectiveOffsetY = 578;
const msooptPerspectiveScaleXToX = 579;
const msooptPerspectiveScaleYToX = 580;
const msooptPerspectiveScaleXToY = 581;
const msooptPerspectiveScaleYToY = 582;
const msooptPerspectivePerspectiveX = 583;
const msooptPerspectivePerspectiveY = 584;
const msooptPerspectiveWeight = 585;
const msooptPerspectiveOriginX = 586;
const msooptPerspectiveOriginY = 587;
const msooptFPerspective = 639;
const msooptC3DSpecularAmt = 640;
const msooptC3DDiffuseAmt = 641;
const msooptC3DShininess = 642;
const msooptC3DEdgeThickness = 643;
const msooptC3DExtrudeForward = 644;
const msooptC3DExtrudeBackward = 645;
const msooptC3DExtrudePlane = 646;
const msooptC3DExtrusionColor = 647;
const msooptC3DCrMod = 648;
const msooptF3D = 700;
const msooptFc3DMetallic = 701;
const msooptFc3DUseExtrusionColor = 702;
const msooptFc3DLightFace = 703;
const msooptC3DYRotationAngle = 704;
const msoopt3DXRotationAngle = 705;
const msoopt3DRotationAxisX = 706;
const msooptC3DRotationAxisY = 707;
const msooptC3DRotationAxisZ = 708;
const msooptC3DRotationAngle = 709;
const msooptC3DRotationCenterX = 710;
const msooptC3DRotationCenterY = 711;
const msooptC3DRotationCenterZ = 712;
const msooptC3DRenderMode = 713;
const msooptC3DTolerance = 714;
const msooptC3DXViewpoint = 715;
const msooptC3DYViewpoint = 716;
const msooptC3DZViewpoint = 717;
const msooptC3DOriginX = 718;
const msooptC3DOriginY = 719;
const msooptC3DSkewAngle = 720;
const msooptC3DSkewAmount = 721;
const msooptC3DAmbientIntensity = 722;
const msooptC3DKeyX = 723;
const msooptC3DKeyY = 724;
const msooptC3DKeyZ = 725;
const msooptC3DKeyIntensity = 726;
const msooptC3DFillX = 727;
const msooptC3DFillY = 728;
const msooptC3DFillZ = 729;
const msoopt3DFillIntensity = 730;
const msooptC3DConstrainRotation = 763;
const msooptC3DRotationCenterAuto = 764;
const msooptC3DParallel = 765;
const msooptC3DKeyHarsh = 766;
const msooptC3DFillHarsh = 767;
const msooptHspMaster = 769;
const msooptCxstyle = 771;
const msooptBWMode = 772;
const msooptBWModePureBW = 773;
const msooptBWModeBW = 774;
const msooptFOleIcon = 826;
const msooptFPreferRelativeResize = 827;
const msooptFLockShapeType = 828;
const msooptFDeleteAttachedObject = 830;
const msooptFBackground = 831;
const msooptSpcot = 832;
const msooptDxyCalloutGap = 833;
const msooptSpcoa = 834;
const msooptSpcod = 835;
const msooptDxyCalloutDropSpecified = 836;
const msooptDxyCalloutLengthSpecified = 837;
const msooptFCallout = 889;
const msooptFCalloutAccentBar = 890;
const msooptFCalloutTextBorder = 891;
const msooptFCalloutMinusX = 892;
const msooptFCalloutMinusY = 893;
const msooptFCalloutDropAuto = 894;
const msooptFCalloutLengthSpecified = 895;
const msooptWzName = 896;
const msooptZDescription = 897;
const msooptPihlShape = 898;
const msooptPWrapPolygonVertices = 899;
const msooptDxWrapDistLeft = 900;
const msooptDyWrapDistTop = 901;
const msooptDxWrapDistRight = 902;
const msooptDyWrapDistBottom = 903;
const msooptLidRegroup = 904;
const msooptFEditedWrap = 953;
const msooptFBehindDocument = 954;
const msooptFOnDblClickNotify = 955;
const msooptFIsButton = 956;
const msooptFOneD = 957;
const msooptFHidden = 958;
const msooptFPrint = 959;

const msooptGeoLeft = 320;
const msooptGeoTop = 321;
const msooptGeoRight = 322;
const msooptGeoBottom = 323;
const msooptShapePath = 324;
const msooptPVertices = 325;
const msooptPSegmentInfo = 326;
const msooptAdjustValue = 327;
const msooptAdjust2Value = 328;
const msooptAdjust3Value = 329;
const msooptAdjust4Value = 330;
const msooptAdjust5Value = 331;
const msooptAdjust6Value = 332;
const msooptAdjust7Value = 333;
const msooptAdjust8Value = 334;
const msooptAdjust9Value = 335;
const msooptAdjust10Value = 336;

const msooptSegmentData1 = 337;
const msooptSegmentData2 = 341;
const msooptSegmentData3 = 342;
const msooptSegmentData4 = 343;

const msooptFShadowOK = 378;
const msooptF3DOK = 379;
const msooptFLineOK = 380;
const msooptFGtextOK = 381;
const msooptFFillShadeShapeOK = 382;
const msooptFFillOK = 383;

const msoshapeLines        = 0;  // A line of straight segments
const msoshapeLinesClosed  = 1;  // A closed polygonal object
const msoshapeCurves       = 2;  // A line of Bezier curve segments
const msoshapeCurvesClosed = 3;  // A closed shape with curved edges
const msoshapeComplex      = 4;  // pSegmentInfo must be non-empty

const msolineNoEnd = 0;
const msolineArrowEnd = 1;
const msolineArrowStealthEnd = 2;
const msolineArrowDiamondEnd = 3;
const msolineArrowOvalEnd = 4;
const msolineArrowOpenEnd = 5;

const msofillSolid = 0;            // Fill with a solid color
const msofillPattern = 1;          // Fill with a pattern (bitmap)
const msofillTexture = 2;          // A texture (pattern with its own color map)
const msofillPicture = 3;          // Center a picture in the shape
const msofillShade = 4;            // Shade from start to end points
const msofillShadeCenter = 5;      // Shade from bounding rectangle to end point
const msofillShadeShape = 6;       // Shade from shape outline to end point
const msofillShadeScale = 7;       // Similar to msofillShade, but the fillAngle
                                   // is additionally scaled by the aspect ratio of
                                   // the shape. If shape is square, it is the
                                   // same as msofillShade.
const msofillShadeTitle = 8;       // special type - shade to title ---  for PP
const msofillBackground = 9;       // Use the background fill color/pattern


type PMSOOPTVertices = ^TMSOOPTVertices;
     TMSOOPTVertices = packed record
     Unknown: word;
     VertCount: word;
     VertSize: word;
     Points: array[0..$FFFF] of TPoint;
     end; 

type PMSOHeader = ^TMSOHeader;
     TMSOHeader = packed record
     VerInst: word;
     FBT: word;
     Length: longword;
     end;

type  TMSOBlipType = (msoblipERROR,msoblipUNKNOWN,msoblipEMF,msoblipWMF,msoblipPICT,msoblipJPEG,msoblipPNG,msoblipDIB);
const TMSOBlipTypeStr: array[0..7] of string = ('ERROR','UNKNOWN','EMF','WMF','PICT','JPEG','PNG','BMP');

type PMSOBSEMetafile = ^TMSOBSEMetafile;
     TMSOBSEMetafile = packed record
     CacheSize: integer;
     Rect: TRect;
     PtSize: TPoint;
     Size: integer;
     Compressed: byte;
     Filter: byte;
     end;

type PMSOFileReGroupItem = ^TMSOFileReGroupItem;
     TMSOFileReGroupItem = packed record
     fridNew: word;
     fridOld: word;
     end;

// ******* MSO Records *******

type PMSORecBSE = ^TMSORecBSE;
     TMSORecBSE = packed record
     WinType: TMSOBlipType;
     MacType: TMSOBlipType;
     BlipId: array[0..15] of byte;
     Tag: word;
     Size: longword;
     RefCount: longword;
     DelayOffset: longword;
     Usage: byte;
     NameLen: byte;
     Unused: word;
     end;

type PMSORecCHILDANCHOR = ^TMSORecCHILDANCHOR;
     TMSORecCHILDANCHOR = packed record
     X1,Y1,X2,Y2: longword;
     end;

type PMSORecCLIENTANCHOR = ^TMSORecCLIENTANCHOR;
     TMSORecCLIENTANCHOR = packed record
     Options: word;
     Col1: word;
     Col1Offset: word;
     Row1: word;
     Row1Offset: word;
     Col2: word;
     Col2Offset: word;
     Row2: word;
     Row2Offset: word;
     end;

type PDGGRecFIDCL = ^TDGGRecFIDCL;
     TDGGRecFIDCL = packed record
     DgId: longword;
     SpIdUsed: longword;
     end;

type PMSORecDG = ^TMSORecDG;
     TMSORecDG = packed record
     ShapeCount_: longword;
     CurSpId: longword;
     end;

type PMSORecDGG = ^TMSORecDGG;
     TMSORecDGG = packed record
     MaxSpId: longword;
     IdClusterCount: longword;
     SavedShapes: longword;
     SavedDrawings: longword;
     end;

type PMSORecOPTValue = ^TMSORecOPTValue;
     TMSORecOPTValue = packed record
     Options: word;
     Value: longword;
     end;

type PMSORecSP = ^TMSORecSP;
     TMSORecSP = packed record
     SpId: longword;
     Options: longword;
     end;

type PMSORecSPGR = ^TMSORecSPGR;
     TMSORecSPGR = packed record
     X1,Y1,X2,Y2: longword;
     end;

implementation

end.
