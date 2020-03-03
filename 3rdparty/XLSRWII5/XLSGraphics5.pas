unit XLSGraphics5;

interface

uses Classes, SysUtils;

type TMSOShapeType = (mstUnknown,mstCtrlLabel,mstCtrlButton,mstCtrlListBox,
                      mstCtrlCheckBox,mstCtrlComboBox,mstCtrlRadioButton,
                      mstCtrlGroupBox,mstCtrlScrollBar,mstTextBox,mstTextShape,
                      mstNote,mstChart);

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

type TXLSPicture = class(TObject)
protected
//     FPNG: TPNGObject;
public
     function IsRaster: boolean;
     function Id: integer;

//     property PNG: TPNGObject read FPNG;
     end;

type TXLSGraphicNote = class(TObject)
private
     function GetAutoVisible: boolean;
protected
     FCol,FRow: integer;
public
     property Col: integer read FCol write FCol;
     property Row: integer read FRow write FRow;
     property AutoVisible: boolean read GetAutoVisible;
     end;

implementation

{ TXLSPicture }

function TXLSPicture.Id: integer;
begin
  Result := -1;
end;

function TXLSPicture.IsRaster: boolean;
begin
  Result := False;
end;

{ TXLSGraphicNote }

function TXLSGraphicNote.GetAutoVisible: boolean;
begin
  Result := True;
end;

end.
