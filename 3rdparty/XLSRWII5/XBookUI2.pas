unit XBookUI2;

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

uses Classes, SysUtils, Types, vcl.Controls, vcl.StdCtrls, vcl.ComCtrls, vcl.Forms, vcl.Graphics,
     Windows, vcl.ExtCtrls, Math,
     ExcelColorPicker,
     Xc12Utils5, Xc12DataStyleSheet5,
     XLSUtils5, XLSMMU5, XLSCellMMU5, XLSSheetData5, XLSReadWriteII5, XLSCmdFormat5,
     XLSMask5,
     XBookUtils2, XLSCellAreas5, XBookStdComponents;

// Naming convention: xuic[Excel dialog][Page][(More specific)][Control]
type TXLSBookUIControlId = (                  // Excpected control type.
     xuicFormat,                              // TForm
     xuicFormatNumberSample,                  // TLabel
     xuicFormatNumberCategory,                // TListBox

     xuicFormatNumberNumberDecimals,          // TUpDown
     xuicFormatNumberNumberDecimals_Txt,      // TLabel
     xuicFormatNumberNumberUse1000Sep,        // TCheckBox
     xuicFormatNumberNumberNegNumbers,        // TListBox
     xuicFormatNumberNumberNegNumbers_Txt,    // TLabel

     xuicFormatNumberCurrencyDecimals,        // TUpDown
     xuicFormatNumberCurrencyDecimals_Txt,    // TLabel
     xuicFormatNumberCurrencySymbol,          // TComboBox
     xuicFormatNumberCurrencySymbol_Txt,      // TLabel
     xuicFormatNumberCurrencyNegNumbers,      // TListBox
     xuicFormatNumberCurrencyNegNumbers_Txt,  // TLabel

     xuicFormatNumberAccountingDecimals,      // TUpDown
     xuicFormatNumberAccountingDecimals_Txt,  // TLabel
     xuicFormatNumberAccountingSymbol,        // TComboBox
     xuicFormatNumberAccountingSymbol_Txt,    // TLabel

     xuicFormatNumberDate,                    // TListBox
     xuicFormatNumberDateType_Txt,            // TLabel

     xuicFormatNumberTime,                    // TListBox
     xuicFormatNumberTimeType_Txt,            // TLabel

     xuicFormatNumberPercentageDecimals,      // TUpDown
     xuicFormatNumberPercentageDecimals_Txt,  // TLabel

     xuicFormatNumberFractionType,            // TListBox
     xuicFormatNumberFractionType_Txt,        // TLabel

     xuicFormatNumberScientificDecimals,      // TUpDown
     xuicFormatNumberScientificDecimals_Txt,  // TLabel

     xuicFormatNumberSpecialType,             // TListBox
     xuicFormatNumberSpecialType_Txt,         // TLabel

     xuicFormatNumberCustom,
     xuicFormatNumberCustomType_Txt,          // TLabel

     xuicFormatAlignOrientation,              // TPaintBox
     xuicFormatAlignDegrees,                  // TUpDown
     xuicFormatAlignHoriz,                    // TComboBox
     xuicFormatAlignVert,                     // TComboBox
     xuicFormatAlignIndent,                   // TUpDown
     xuicFormatAlignJustifyDistributed,       // TCheckbox
     xuicFormatAlignWrapText,                 // TCheckbox
     xuicFormatAlignShrinkToFit,              // TCheckbox
     xuicFormatAlignMergeCells,               // TCheckbox
     xuicFormatAlignTextDirection,            // TComboBox

     xuicFormatFontFont,                      // TListBox
     xuicFormatFontFontEd,                    // TEdit
     xuicFormatFontStyle,                     // TListBox
     xuicFormatFontStyleEd,                   // TEdit
     xuicFormatFontSize,                      // TListBox
     xuicFormatFontSizeEd,                    // TEdit
     xuicFormatFontUnderline,                 // TComboBox
     xuicFormatFontColor,                     // TXBookColorComboBox
     xuicFormatFontNormal,                    // TButton
     xuicFormatFontStriketrough,              // TCheckBox
     xuicFormatFontSuperscript,               // TCheckBox
     xuicFormatFontSubscript,                 // TCheckBox
     xuicFormatFontPreview,                   // TPaintBox

     xuicFormatBorderColor,                   // TXBookColorComboBox
     xuicFormatBorderStyle,                   // TXBookCellBorderStylePicker
     xuicFormatBorderBorders,                 // TXBookCellBorderPicker

     xuicFillColor,                           // TExcelColorPicker
     xuicFillNoColor,                         // TButton
     xuicFillSample,                          // TShape

     xuicProtectionLocked,                    // TCheckBox
     xuicProtectionHidden                     // TCheckBox
     );

type TXLSBookUIControlData = record
     Control: TControl;
     ControlClass: TClass;
     Group: integer;
     end;

type TXLSBookUI = class(TObject)
protected
     FXLS        : TXLSReadWriteII5;
     FSheet      : TXLSWorksheet;
     FControls   : array[Low(TXLSBookUIControlId)..High(TXLSBookUIControlId)] of TXLSBookUIControlData;
     FSampleNum  : double;
     FDecimals   : integer;
     FFillChanged: boolean;

     FBMPTrueType: vcl.Graphics.TBitmap;
     FBMPPrinter : vcl.Graphics.TBitmap;

     FAlignOrientationBtnDown: boolean;

     function  SearchStrings(const AText: AxUCString; AStrings: TSTrings): integer;

     procedure AddControls;
     procedure AddControlData(const AId: TXLSBookUIControlId; AClass: TClass; const AGroup: integer = -1);
     procedure ShowControls(const AGroup: integer; const AVisible: boolean);

     procedure UpdateFontStyle;

     procedure UpdateNumberSample;
     procedure UpdateNegNumSample;

     procedure UpdateNumFmt;

     procedure UpdateFontSample;

     procedure DialogActivate(ASender: TObject);
     procedure DialogCloseQuery(ASender: TObject; var ACanClose: boolean);

     procedure NumberCategoryClick(ASender: TObject);
     procedure NumberChanged(ASender: TObject);
     procedure NumberDecimalsChanged(ASender: TObject);
     procedure NumberNegNumbersDrawItem(Control: TWinControl; Index: Integer; Rect: TRect; State: TOwnerDrawState);
     procedure NumberNegNumbersClick(ASender: TObject);
     procedure NumberCurrencyCloseUp(ASender: TObject);
     procedure NumberTypeClick(ASender: TObject);
     procedure NumberCustomChanged(ASender: TObject);

     procedure AlignOrientationPaint(Sender: TObject);
     procedure AlignOrientationMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
     procedure AlignOrientationMouseMove(Sender: TObject; Shift: TShiftState; X, Y: Integer);
     procedure AlignOrientationMouseUp(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
     procedure AlignOrientationDegreesChanged(ASender: TObject);
     procedure AlignChanged(ASender: TObject);

     procedure FontFontDrawItem(Control: TWinControl; Index: Integer; Rect: TRect; State: TOwnerDrawState);
     procedure FontFontClick(ASender: TObject);
     procedure FontFontEdChanged(ASender: TObject);
     procedure FontStyleClick(ASender: TObject);
     procedure FontSizeEdChanged(ASender: TObject);
     procedure FontSizeClick(ASender: TObject);
     procedure FontUnderlineCloseUp(ASender: TObject);
     procedure FontStriketroughClick(ASender: TObject);
     procedure FontSuperscriptClick(ASender: TObject);
     procedure FontSubscriptClick(ASender: TObject);
     procedure FontNormalFontClick(ASender: TObject);
     procedure FontColorChanged(ASender: TObject);
     procedure FontSamplePaint(ASender: TObject);

     procedure BorderColorSelect(ASender: TObject);
     procedure BorderStyleSelect(ASender: TObject);
     procedure BorderBorderSelect(ASender: TObject);

     procedure FillColorClick(ASender: TObject);
     procedure FillNoColorClick(ASender: TObject);

     procedure ProtectionLockedClick(ASender: TObject);
     procedure ProtectionHiddenClick(ASender: TObject);

     procedure SelectNumberType(const AIndex: integer);

     procedure FillCurrencyStrings(AStrings: TStrings);
     function  FindCommonFill: TXc12Fill;
     procedure Prepare;
public
     constructor Create(AXLS: TXLSReadWriteII5);
     destructor Destroy; override;

     procedure RegisterControl(AControl: TControl; const AId: TXLSBookUIControlId);

     property Sheet: TXLSWorksheet read FSheet write FSheet;
     end;

implementation

const UI_NUM_GENERAL       = 0;
const UI_NUM_NUMBER        = 1;
const UI_NUM_CURRENCY      = 2;
const UI_NUM_ACCOUNTING    = 3;
const UI_NUM_DATE          = 4;
const UI_NUM_TIME          = 5;
const UI_NUM_PERCENTAGE    = 6;
const UI_NUM_FRACTION      = 7;
const UI_NUM_SCIENTIFIC    = 8;
const UI_NUM_TEXT          = 9;
const UI_NUM_SPECIAL       = 10;
const UI_NUM_CUSTOM        = 11;

const UI_NUM_FIRST         = UI_NUM_GENERAL;
const UI_NUM_LAST          = UI_NUM_CUSTOM;

const LCID_List: array[0..236] of integer =(
1078,1052,1156,1118,1025,5121,15361,3073,2049,11265,13313,12289,4097,6145,8193,16385,10241,7169,14337,9217,1067,1101,2092,1068,1133,
1069,1059,1093,2117,5146,1150,1026,1109,1027,1116,2052,4100,1028,3076,5124,1155,1050,4122,1029,1030,1164,1125,1043,2067,1126,1033,
2057,3081,10249,4105,9225,15369,16393,14345,6153,8201,17417,5129,13321,18441,7177,11273,12297,1061,1080,1065,1124,1035,1036,2060,11276,
3084,9228,12300,15372,5132,13324,6156,14348,58380,8204,10252,4108,7180,1122,1127,1071,1110,1079,1031,3079,5127,4103,2055,1032,1135,
1140,1095,1128,1141,1037,1081,1038,1129,1039,1136,1057,1117,2108,1040,2064,1041,1158,1099,1137,2144,1120,1087,1107,1159,1111,1042,1088,
1108,1142,1062,1063,1134,1086,2110,1100,1082,1112,1153,1146,1102,1148,1104,2128,1121,2145,1044,2068,1154,1096,1138,1145,1123,1045,1046,
2070,1094,2118,1131,2155,3179,1047,1048,2072,1049,2073,1083,1103,1084,1132,3098,2074,1113,2137,1115,1051,1060,1143,1070,3082,1034,11274,
16394,13322,9226,5130,7178,12298,17418,4106,18442,22538,2058,19466,6154,15370,10250,20490,21514,14346,8202,1072,1089,1053,2077,1114,
1064,1119,2143,1097,1092,1098,1054,2129,1105,2163,1139,1073,1074,1055,1090,1152,1058,1056,2080,2115,1091,1075,1066,1106,1160,1076,1157,
1144,1085,1130,1077);

const DataBMP_Printer: array[0..725] of byte = (
      $42,$4D,$D6,$02,$00,$00,$00,$00,$00,$00,$36,$00,$00,$00,$28,$00,$00,$00,$0E,$00,
      $00,$00,$0C,$00,$00,$00,$01,$00,$20,$00,$00,$00,$00,$00,$A0,$02,$00,$00,$12,$0B,$00,$00,$12,$0B,$00,$00,$00,$00,$00,$00,$00,$00,
      $00,$00,$FF,$FF,$FF,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,
      $00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$FF,$FF,$FF,$00,$00,$00,$00,$00,$80,$80,
      $80,$00,$80,$80,$80,$00,$80,$80,$80,$00,$80,$80,$80,$00,$80,$80,$80,$00,$80,$80,$80,$00,$80,$80,$80,$00,$80,$80,$80,$00,$00,$00,
      $FF,$00,$00,$00,$FF,$00,$80,$80,$80,$00,$80,$80,$80,$00,$00,$00,$00,$00,$00,$00,$00,$00,$C0,$C0,$C0,$00,$C0,$C0,$C0,$00,$C0,$C0,
      $C0,$00,$C0,$C0,$C0,$00,$C0,$C0,$C0,$00,$C0,$C0,$C0,$00,$C0,$C0,$C0,$00,$C0,$C0,$C0,$00,$C0,$C0,$C0,$00,$C0,$C0,$C0,$00,$C0,$C0,
      $C0,$00,$80,$80,$80,$00,$00,$00,$00,$00,$FF,$FF,$FF,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,
      $00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$FF,$FF,
      $FF,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$00,$00,$00,$00,$80,$80,$80,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$FF,$FF,
      $FF,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$80,$80,$80,$00,$00,$00,$00,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$FF,$FF,
      $FF,$00,$00,$00,$00,$00,$FF,$FF,$FF,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,
      $00,$00,$FF,$FF,$FF,$00,$00,$00,$00,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$00,$00,$00,$00,$80,$80,
      $80,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$80,$80,$80,$00,$00,$00,
      $00,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$00,$00,$00,$00,$FF,$FF,$FF,$00,$00,$00,$00,$00,$00,$00,
      $00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$FF,$FF,$FF,$00,$00,$00,$00,$00,$FF,$FF,$FF,$00,$FF,$FF,
      $FF,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$00,$00,$00,$00,$80,$80,$80,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$FF,$FF,
      $FF,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$80,$80,$80,$00,$00,$00,$00,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$FF,$FF,
      $FF,$00,$00,$00,$00,$00,$FF,$FF,$FF,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,
      $00,$00,$FF,$FF,$FF,$00,$00,$00,$00,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$00,$00,$00,$00,$80,$80,
      $80,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$80,$80,$80,$00,$00,$00,
      $00,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,
      $00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$FF,$FF,$FF,$00,$FF,$FF,
      $FF,$00);


const DataBMP_TrueType: array[0..725] of byte = (
      $42,$4D,$D6,$02,$00,$00,$00,$00,$00,$00,$36,$00,$00,$00,$28,$00,$00,$00,$0E,$00,
      $00,$00,$0C,$00,$00,$00,$01,$00,$20,$00,$00,$00,$00,$00,$A0,$02,$00,$00,$12,$0B,$00,$00,$12,$0B,$00,$00,$00,$00,$00,$00,$00,$00,
      $00,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$00,$00,$00,$00,$00,$00,
      $00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$FF,$FF,
      $FF,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$00,$00,$00,$00,$00,$00,
      $00,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$FF,$FF,
      $FF,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$00,$00,$00,$00,$00,$00,$00,$00,$FF,$FF,$FF,$00,$FF,$FF,
      $FF,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$80,$80,$80,$00,$80,$80,$80,$00,$80,$80,$80,$00,$80,$80,
      $80,$00,$80,$80,$80,$00,$80,$80,$80,$00,$00,$00,$00,$00,$00,$00,$00,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$FF,$FF,
      $FF,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$80,$80,$80,$00,$80,$80,$80,$00,$FF,$FF,$FF,$00,$FF,$FF,
      $FF,$00,$00,$00,$00,$00,$00,$00,$00,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$FF,$FF,
      $FF,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$80,$80,$80,$00,$00,$00,$00,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$00,$00,$00,$00,$00,$00,
      $00,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$00,$00,$00,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$FF,$FF,
      $FF,$00,$80,$80,$80,$00,$00,$00,$00,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$00,$00,$00,$00,$00,$00,$00,$00,$FF,$FF,$FF,$00,$FF,$FF,
      $FF,$00,$00,$00,$00,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$80,$80,$80,$00,$00,$00,
      $00,$00,$00,$00,$00,$00,$FF,$FF,$FF,$00,$00,$00,$00,$00,$00,$00,$00,$00,$FF,$FF,$FF,$00,$00,$00,$00,$00,$00,$00,$00,$00,$FF,$FF,
      $FF,$00,$FF,$FF,$FF,$00,$80,$80,$80,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$80,$80,$80,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,
      $00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$80,$80,
      $80,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$80,$80,$80,$00,$80,$80,$80,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$80,$80,$80,$00,$FF,$FF,
      $FF,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$80,$80,$80,$00,$80,$80,$80,$00,$FF,$FF,
      $FF,$00,$80,$80,$80,$00,$80,$80,$80,$00,$FF,$FF,$FF,$00,$80,$80,$80,$00,$80,$80,$80,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$FF,$FF,
      $FF,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$80,$80,$80,$00,$80,$80,$80,$00,$80,$80,$80,$00,$80,$80,$80,$00,$80,$80,
      $80,$00,$80,$80,$80,$00,$80,$80,$80,$00,$80,$80,$80,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$FF,$FF,$FF,$00,$FF,$FF,
      $FF,$00);

{ TXLSBookUI }

procedure TXLSBookUI.AddControlData(const AId: TXLSBookUIControlId; AClass: TClass; const AGroup: integer);
begin
  FControls[AId].Control := Nil;
  FControls[AId].ControlClass := AClass;
  FControls[AId].Group := AGroup
end;

procedure TXLSBookUI.AddControls;
begin
  AddControlData(xuicFormat,TForm);

  AddControlData(xuicFormatNumberSample,TLabel);
  AddControlData(xuicFormatNumberCategory,TListBox);

  AddControlData(xuicFormatNumberNumberDecimals,TUpDown,UI_NUM_NUMBER);
  AddControlData(xuicFormatNumberNumberDecimals_Txt,TLabel,UI_NUM_NUMBER);
  AddControlData(xuicFormatNumberNumberUse1000Sep,TCheckBox,UI_NUM_NUMBER);
  AddControlData(xuicFormatNumberNumberNegNumbers,TListBox,UI_NUM_NUMBER);
  AddControlData(xuicFormatNumberNumberNegNumbers_Txt,TLabel,UI_NUM_NUMBER);

  AddControlData(xuicFormatNumberCurrencyDecimals,TUpDown,UI_NUM_CURRENCY);
  AddControlData(xuicFormatNumberCurrencyDecimals_Txt,TLabel,UI_NUM_CURRENCY);
  AddControlData(xuicFormatNumberCurrencySymbol,TComboBox,UI_NUM_CURRENCY);
  AddControlData(xuicFormatNumberCurrencySymbol_Txt,TLabel,UI_NUM_CURRENCY);
  AddControlData(xuicFormatNumberCurrencyNegNumbers,TListBox,UI_NUM_CURRENCY);
  AddControlData(xuicFormatNumberCurrencyNegNumbers_Txt,TLabel,UI_NUM_CURRENCY);

  AddControlData(xuicFormatNumberAccountingDecimals,TUpDown,UI_NUM_ACCOUNTING);
  AddControlData(xuicFormatNumberAccountingDecimals_Txt,TLabel,UI_NUM_ACCOUNTING);
  AddControlData(xuicFormatNumberAccountingSymbol,TComboBox,UI_NUM_ACCOUNTING);
  AddControlData(xuicFormatNumberAccountingSymbol_Txt,TLabel,UI_NUM_ACCOUNTING);

  AddControlData(xuicFormatNumberDate,TListBox,UI_NUM_DATE);
  AddControlData(xuicFormatNumberDateType_Txt,TLabel,UI_NUM_DATE);

  AddControlData(xuicFormatNumberTime,TListBox,UI_NUM_TIME);
  AddControlData(xuicFormatNumberTimeType_Txt,TLabel,UI_NUM_TIME);

  AddControlData(xuicFormatNumberPercentageDecimals,TUpDown,UI_NUM_PERCENTAGE);
  AddControlData(xuicFormatNumberPercentageDecimals_Txt,TLabel,UI_NUM_PERCENTAGE);

  AddControlData(xuicFormatNumberFractionType,TListBox,UI_NUM_FRACTION);
  AddControlData(xuicFormatNumberFractionType_Txt,TLabel,UI_NUM_FRACTION);

  AddControlData(xuicFormatNumberScientificDecimals,TUpDown,UI_NUM_SCIENTIFIC);
  AddControlData(xuicFormatNumberScientificDecimals_Txt,TLabel,UI_NUM_SCIENTIFIC);

  AddControlData(xuicFormatNumberSpecialType,TListBox,UI_NUM_SPECIAL);
  AddControlData(xuicFormatNumberSpecialType_Txt,TLabel,UI_NUM_SPECIAL);

  AddControlData(xuicFormatNumberCustom,TEdit,UI_NUM_CUSTOM);
  AddControlData(xuicFormatNumberCustomType_Txt,TLabel,UI_NUM_CUSTOM);

  AddControlData(xuicFormatAlignOrientation,TPaintBox);
  AddControlData(xuicFormatAlignDegrees,TUpDown);
  AddControlData(xuicFormatAlignHoriz,TComboBox);
  AddControlData(xuicFormatAlignVert,TComboBox);
  AddControlData(xuicFormatAlignIndent,TUpDown);
  AddControlData(xuicFormatAlignJustifyDistributed,TCheckbox);
  AddControlData(xuicFormatAlignWrapText,TCheckbox);
  AddControlData(xuicFormatAlignShrinkToFit,TCheckbox);
  AddControlData(xuicFormatAlignMergeCells,TCheckbox);
  AddControlData(xuicFormatAlignTextDirection,TComboBox);

  AddControlData(xuicFormatFontFont,TListBox);
  AddControlData(xuicFormatFontFontEd,TEdit);
  AddControlData(xuicFormatFontStyle,TListBox);
  AddControlData(xuicFormatFontStyleEd,TEdit);
  AddControlData(xuicFormatFontSize,TListBox);
  AddControlData(xuicFormatFontSizeEd,TEdit);
  AddControlData(xuicFormatFontUnderline,TComboBox);
  AddControlData(xuicFormatFontColor,TXBookColorComboBox);
  AddControlData(xuicFormatFontNormal,TButton);
  AddControlData(xuicFormatFontStriketrough,TCheckBox);
  AddControlData(xuicFormatFontSuperscript,TCheckBox);
  AddControlData(xuicFormatFontSubscript,TCheckBox);
  AddControlData(xuicFormatFontPreview,TPaintBox);

  AddControlData(xuicFormatBorderColor,TXBookColorComboBox);
  AddControlData(xuicFormatBorderStyle,TXBookCellBorderStylePicker);
  AddControlData(xuicFormatBorderBorders,TXBookCellBorderPicker);

  AddControlData(xuicFillColor,TExcelColorPicker);
  AddControlData(xuicFillNoColor,TButton);
  AddControlData(xuicFillSample,TShape);

  AddControlData(xuicProtectionLocked,TCheckBox);
  AddControlData(xuicProtectionHidden,TCheckBox);
end;

procedure TXLSBookUI.AlignChanged(ASender: TObject);
begin
  FXLS.CmdFormat.Alignment.Horizontal := TXc12HorizAlignment(TComboBox(FControls[xuicFormatAlignHoriz].Control).ItemIndex);
  FXLS.CmdFormat.Alignment.Vertical := TXc12VertAlignment(TComboBox(FControls[xuicFormatAlignVert].Control).ItemIndex);
  FXLS.CmdFormat.Alignment.Indent := TUpDown(FControls[xuicFormatAlignIndent].Control).Position;
  FXLS.CmdFormat.Alignment.JustifyDistributed := TCheckBox(FControls[xuicFormatAlignJustifyDistributed].Control).Checked;
  FXLS.CmdFormat.Alignment.WrapText := TCheckBox(FControls[xuicFormatAlignWrapText].Control).Checked;
  FXLS.CmdFormat.Alignment.ShrinkToFit := TCheckBox(FControls[xuicFormatAlignShrinkToFit].Control).Checked;
  FXLS.CmdFormat.Alignment.MergeCells := TCheckBox(FControls[xuicFormatAlignMergeCells].Control).Checked;
  FXLS.CmdFormat.Alignment.TextDirection := TXc12ReadOrder(TComboBox(FControls[xuicFormatAlignTextDirection].Control).ItemIndex);
end;

procedure TXLSBookUI.AlignOrientationDegreesChanged(ASender: TObject);
begin
  AlignOrientationPaint(FControls[xuicFormatAlignOrientation].Control);

  FXLS.CmdFormat.Alignment.Rotation := TUpDown(FControls[xuicFormatAlignDegrees].Control).Position;
end;

procedure TXLSBookUI.AlignOrientationMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
  V: double;
  D: integer;
begin
  V := RadToDeg(ArcTan2(X - 10,Y - (TPaintBox(Sender).Height div 2))) - 90;
  if Abs(Frac(V / 15)) < 0.2 then
    D := Round(Int(V / 15) * 15)
  else
    D := Round(V);
  if D < -200 then
    D := 90;

  TUpDown(FControls[xuicFormatAlignDegrees].Control).Position := Fork(D,-90,90);

  AlignOrientationPaint(Sender);
  FAlignOrientationBtnDown := True;
end;

procedure TXLSBookUI.AlignOrientationMouseMove(Sender: TObject; Shift: TShiftState; X, Y: Integer);
var
  D: integer;
begin
  if FAlignOrientationBtnDown then begin
    D := Round(RadToDeg(ArcTan2(X - 10,Y - (TPaintBox(Sender).Height div 2))) - 90);
    if D < -200 then
      D := 90;
    TUpDown(FControls[xuicFormatAlignDegrees].Control).Position := Fork(D,-90,90);
    AlignOrientationPaint(Sender);
  end;
end;

procedure TXLSBookUI.AlignOrientationMouseUp(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  FAlignOrientationBtnDown := False;
end;

procedure TXLSBookUI.AlignOrientationPaint(Sender: TObject);
var
  PB: TPaintBox;
  D: integer;
  Deg: integer;
  X,Y: integer;
  W,H: double;
begin
  PB := TPaintBox(Sender);

  PB.Canvas.Brush.Style := bsSolid;
  PB.Canvas.Pen.Color := clBlack;
  PB.Canvas.Brush.Color := clWhite;
  PB.Canvas.Rectangle(0,0,PB.Width,PB.Height);

  D := TUpDown(FControls[xuicFormatAlignDegrees].Control).Position;
  W := PB.Width - 15;
  H := PB.Height / 2;
  Deg := -90;
  while Deg <= 90 do begin
    X := Round(Sin(DegToRad(Deg + 90)) * W) + 10;
    Y := Round(H + Cos(DegToRad(Deg + 90)) * W);

    if Deg = D then begin
      PB.Canvas.Pen.Color := clRed;
      PB.Canvas.Brush.Color := clRed;
    end
    else begin
      PB.Canvas.Pen.Color := clBlack;
      PB.Canvas.Brush.Color := clBlack;
    end;

    if (Deg mod 45) = 0 then
      PB.Canvas.Ellipse(X - 3,Y - 3,X + 3,Y + 3)
    else
      PB.Canvas.Rectangle(X - 1,Y - 1,X + 1,Y + 1);

    Inc(Deg,15);
  end;

  PB.Canvas.Pen.Color := clBlack;
  PB.Canvas.MoveTo(10,Round(H));
  PB.Canvas.LineTo(Round(Sin(DegToRad(D + 90)) * W) + 10,Round(H + Cos(DegToRad(D + 90)) * W));

  PB.Canvas.Brush.Style := bsClear;
{$ifdef DELPHI_2006_OR_LATER}
  PB.Canvas.Font.Orientation := D * 10;
{$endif}
  SetTextAlign(PB.Canvas.Handle,TA_LEFT or TA_BASELINE);
  PB.Canvas.TextOut(10,Round(H),'Text');
end;

procedure TXLSBookUI.BorderBorderSelect(ASender: TObject);
begin

end;

procedure TXLSBookUI.BorderColorSelect(ASender: TObject);
begin
  TXBookCellBorderPicker(FControls[xuicFormatBorderBorders].Control).BorderColor := TXBookColorComboBox(ASender).Color;
  TXBookCellBorderStylePicker(FControls[xuicFormatBorderStyle].Control).LineColor := TXBookColorComboBox(ASender).Color;
end;

procedure TXLSBookUI.BorderStyleSelect(ASender: TObject);
begin
  TXBookCellBorderPicker(FControls[xuicFormatBorderBorders].Control).BorderStyle := TXBookCellBorderStylePicker(ASender).LineStyle;
end;

constructor TXLSBookUI.Create(AXLS: TXLSReadWriteII5);
var
  Stream: TMemoryStream;
  P: Pointer;
begin
  FXLS := AXLS;

  FBMPTrueType := vcl.Graphics.TBitmap.Create;
  FBMPPrinter := vcl.Graphics.TBitmap.Create;

  Stream := TMemoryStream.Create;
  try
    P := @DataBMP_TrueType;
    Stream.Write(P^,Length(DataBMP_TrueType));
    Stream.Seek(0,soFromBeginning);

    FBMPTrueType.LoadFromStream(Stream);

    Stream.Clear;
    P := @DataBMP_Printer;
    Stream.Write(P^,Length(DataBMP_Printer));
    Stream.Seek(0,soFromBeginning);

    FBMPPrinter.LoadFromStream(Stream);
  finally
    Stream.Free;
  end;

  AddControls;
end;

destructor TXLSBookUI.Destroy;
begin
  FBMPTrueType.Free;
  FBMPPrinter.Free;

  inherited;
end;

procedure TXLSBookUI.FillColorClick(ASender: TObject);
begin
  TShape(FControls[xuicFillSample].Control).Brush.Style := bsSolid;
  TShape(FControls[xuicFillSample].Control).Brush.Color := RevRGB(TExcelColorPicker(ASender).ExcelColor.ARGB);

  FFillChanged := True;
end;

procedure TXLSBookUI.FillCurrencyStrings(AStrings: TStrings);
var
  i: integer;
  sCurrency,sLang,sCountry: AxUCString;
begin
  for i := 0 to High(LCID_List) do begin
{$WARN SYMBOL_PLATFORM OFF}
    sCurrency := GetLocaleStr(LCID_List[i],LOCALE_SCURRENCY,'');
    sLang := GetLocaleStr(LCID_List[i],LOCALE_SENGLANGUAGE,'');
    sCountry := GetLocaleStr(LCID_List[i],LOCALE_SENGCOUNTRY,'');
{$WARN SYMBOL_PLATFORM ON}
    if (sCurrency <> '') and (sLang <> '') and (sCountry <> '') then
      AStrings.AddObject(sCurrency + ' ' + sLang + ' (' + sCountry + ')',TObject(LCID_List[i]));
  end;
end;

procedure TXLSBookUI.FillNoColorClick(ASender: TObject);
begin
  TShape(FControls[xuicFillSample].Control).Brush.Style := bsClear;
  TExcelColorPicker(FControls[xuicFillColor].Control).ExcelColor := MakeXc12ColorAuto;

  FFillChanged := True;
end;

function TXLSBookUI.FindCommonFill: TXc12Fill;
var
  i,j: integer;
  C,R: integer;
  Cell: TXLSCellItem;
  Fill: TXc12Fill;
  CA: TCellArea;
begin
  Fill := Nil;
  Result := Nil;
  for i := 0 to FSheet.SelectedAreas.Count - 1 do begin
    CA := FSheet.SelectedAreas[i];
    if CA.IsColumn then begin
      if Fill = Nil then
        Fill := FSheet.Columns[CA.Col1].Style.Fill;
      for C := CA.Col1 to CA.Col2 do begin
        if FSheet.Columns[C].Style.Fill <> Fill then
          Exit;
      end;
    end
    else if CA.IsRow then begin
      if Fill = Nil then
        Fill := FSheet.Rows[CA.Row1].Style.Fill;
      for R := CA.Row1 to CA.Row2 do begin
        if FSheet.Rows[R].Style.Fill <> Fill then
          Exit;
      end;
    end
    else begin
      for R := CA.Row1 to CA.Row2 do begin
        for C := CA.Col1 to CA.Col2 do begin
          Cell := FSheet.MMUCells.FindCell(C,R);
          if Cell.Data <> Nil then begin
            j := FSheet.MMUCells.GetStyle(@Cell);
            if Fill = Nil then
              Fill := FXLS.Manager.StyleSheet.XFs[j].Fill;
            if Fill <> FXLS.Manager.StyleSheet.XFs[j].Fill then
              Exit;
          end;
        end;
      end;
    end
  end;
  Result := Fill;
end;

procedure TXLSBookUI.FontColorChanged(ASender: TObject);
begin
  FXLS.CmdFormat.Font.Color.Color := TXBookColorComboBox(ASender).Color;

  UpdateFontSample;
end;

procedure TXLSBookUI.FontFontClick(ASender: TObject);
begin
  TEdit(FControls[xuicFormatFontFontEd].Control).Text := TListBox(ASender).Items[TListBox(ASender).ItemIndex];
  FXLS.CmdFormat.Font.Name := TListBox(ASender).Items[TListBox(ASender).ItemIndex];

  UpdateFontSample;
end;

procedure TXLSBookUI.FontFontDrawItem(Control: TWinControl; Index: Integer; Rect: TRect; State: TOwnerDrawState);
var
  LB: TListBox;
begin
  LB := TListBox(Control);

  LB.Canvas.FillRect(Rect);
  case Integer(LB.Items.Objects[Index]) of
    DEVICE_FONTTYPE  : LB.Canvas.Draw(Rect.Left + 1, Rect.Top + 1,FBMPPrinter);
    RASTER_FONTTYPE  : ;
    TRUETYPE_FONTTYPE: LB.Canvas.Draw(Rect.Left + 1, Rect.Top + 1,FBMPTrueType);
  end;

  LB.Canvas.TextOut(Rect.Left + 16, Rect.Top, LB.Items[Index]);
  if odFocused In State then begin
    LB.Canvas.Brush.Color := clWindow;
    LB.Canvas.DrawFocusRect(Rect);
  end;
end;

procedure TXLSBookUI.FontFontEdChanged(ASender: TObject);
var
  i: integer;
begin
  i := SearchStrings(TEdit(ASender).Text,TListBox(FControls[xuicFormatFontFont].Control).Items);
  if i >= 0 then begin
    TListBox(FControls[xuicFormatFontFont].Control).ItemIndex := i;
    FXLS.CmdFormat.Font.Name := TListBox(FControls[xuicFormatFontFont].Control).Items[i];
  end;

  UpdateFontSample;
end;

procedure TXLSBookUI.FontNormalFontClick(ASender: TObject);
begin
  FXLS.CmdFormat.Font.SetDefault;

  TEdit(FControls[xuicFormatFontFontEd].Control).Text := FXLS.CmdFormat.Font.Name;
  UpdateFontStyle;
  TEdit(FControls[xuicFormatFontSizeEd].Control).Text := IntToStr(Round(FXLS.CmdFormat.Font.Size));
  TXBookColorComboBox(FControls[xuicFormatFontSizeEd].Control).Color := MakeXc12ColorAuto;
  TComboBox(FControls[xuicFormatFontUnderline].Control).ItemIndex := 0;
  TCheckBox(FControls[xuicFormatFontStriketrough].Control).Checked := False;
  TCheckBox(FControls[xuicFormatFontSuperscript].Control).Checked := False;
  TCheckBox(FControls[xuicFormatFontSubscript].Control).Checked := False;
end;

procedure TXLSBookUI.FontSamplePaint(ASender: TObject);
begin
  UpdateFontSample;
end;

procedure TXLSBookUI.FontSizeClick(ASender: TObject);
begin
  TEdit(FControls[xuicFormatFontSizeEd].Control).Text := TListBox(ASender).Items[TListBox(ASender).ItemIndex];
  FXLS.CmdFormat.Font.Size := StrToIntDef(TListBox(ASender).Items[TListBox(ASender).ItemIndex],11);

  UpdateFontSample;
end;

procedure TXLSBookUI.FontSizeEdChanged(ASender: TObject);
var
  i: integer;
begin
  i := SearchStrings(TEdit(ASender).Text,TListBox(FControls[xuicFormatFontSize].Control).Items);
  if i >= 0 then begin
    TListBox(FControls[xuicFormatFontSize].Control).ItemIndex := i;
    FXLS.CmdFormat.Font.Size := StrToIntDef(TListBox(FControls[xuicFormatFontSize].Control).Items[i],11);
  end;

  UpdateFontSample;
end;

procedure TXLSBookUI.FontStriketroughClick(ASender: TObject);
begin
  FXLS.CmdFormat.Font.Style := FXLS.CmdFormat.Font.Style + [xfsStrikeOut];

  UpdateFontSample;
end;

procedure TXLSBookUI.FontStyleClick(ASender: TObject);
begin
  TEdit(FControls[xuicFormatFontStyleEd].Control).Text := TListBox(ASender).Items[TListBox(ASender).ItemIndex];
  case TListBox(ASender).ItemIndex of
    0 : FXLS.CmdFormat.Font.Style := [];
    1 : FXLS.CmdFormat.Font.Style := [xfsItalic];
    2 : FXLS.CmdFormat.Font.Style := [xfsBold];
    3 : FXLS.CmdFormat.Font.Style := [xfsbold,xfsItalic];
  end;

  UpdateFontSample;
end;

procedure TXLSBookUI.FontSubscriptClick(ASender: TObject);
begin
  if TCheckBox(ASender).Checked then
    FXLS.CmdFormat.Font.SubSuperscript := xssSubscript
  else
    FXLS.CmdFormat.Font.SubSuperscript := xssNone;
  TCheckBox(FControls[xuicFormatFontSuperscript].Control).OnClick := Nil;
  TCheckBox(FControls[xuicFormatFontSuperscript].Control).Checked := False;
  TCheckBox(FControls[xuicFormatFontSuperscript].Control).OnClick := FontSuperscriptClick;

  UpdateFontSample;
end;

procedure TXLSBookUI.FontSuperscriptClick(ASender: TObject);
begin
  if TCheckBox(ASender).Checked then
    FXLS.CmdFormat.Font.SubSuperscript := xssSuperscript
  else
    FXLS.CmdFormat.Font.SubSuperscript := xssNone;
  TCheckBox(FControls[xuicFormatFontSubscript].Control).OnClick := Nil;
  TCheckBox(FControls[xuicFormatFontSubscript].Control).Checked := False;
  TCheckBox(FControls[xuicFormatFontSubscript].Control).OnClick := FontSubscriptClick;

  UpdateFontSample;
end;

procedure TXLSBookUI.FontUnderlineCloseUp(ASender: TObject);
begin
  FXLS.CmdFormat.Font.Underline := TXc12Underline(TComboBox(ASender).ItemIndex);

  UpdateFontSample;
end;

procedure TXLSBookUI.DialogActivate(ASender: TObject);
begin
  FXLS.CmdFormat.SetRowHeight := True;
  FXLS.CmdFormat.BeginEdit(FSheet);

  Prepare;

  SelectNumberType(UI_NUM_GENERAL);

  UpdateNumFmt;
end;

procedure TXLSBookUI.NumberCategoryClick(ASender: TObject);
begin
  SelectNumberType(TListBox(ASender).ItemIndex);

  UpdateNumFmt;
end;

procedure TXLSBookUI.NumberChanged(ASender: TObject);
begin
  UpdateNumFmt;
end;

procedure TXLSBookUI.DialogCloseQuery(ASender: TObject; var ACanClose: boolean);
var
  CBP: TXBookCellBorderPicker;
begin
  if TForm(ASender).ModalResult = mrOk then begin
    CBP := TXBookCellBorderPicker(FControls[xuicFormatBorderBorders].Control);
    if CBP.IsChanged then begin
      if CBP.Sides[xcbpsLeft].Changed then begin
        FXLS.CmdFormat.Border.Color.Color := CBP.Sides[xcbpsLeft].Color;
        FXLS.CmdFormat.Border.Style := CBP.Sides[xcbpsLeft].Style;
        FXLS.CmdFormat.Border.Side[cbsLeft] := True;
      end;
      if CBP.Sides[xcbpsTop].Changed then begin
        FXLS.CmdFormat.Border.Color.Color := CBP.Sides[xcbpsTop].Color;
        FXLS.CmdFormat.Border.Style := CBP.Sides[xcbpsTop].Style;
        FXLS.CmdFormat.Border.Side[cbsTop] := True;
      end;
      if CBP.Sides[xcbpsRight].Changed then begin
        FXLS.CmdFormat.Border.Color.Color := CBP.Sides[xcbpsRight].Color;
        FXLS.CmdFormat.Border.Style := CBP.Sides[xcbpsRight].Style;
        FXLS.CmdFormat.Border.Side[cbsRight] := True;
      end;
      if CBP.Sides[xcbpsBottom].Changed then begin
        FXLS.CmdFormat.Border.Color.Color := CBP.Sides[xcbpsBottom].Color;
        FXLS.CmdFormat.Border.Style := CBP.Sides[xcbpsBottom].Style;
        FXLS.CmdFormat.Border.Side[cbsBottom] := True;
      end;
      if CBP.Sides[xcbpsInsideVert].Changed then begin
        FXLS.CmdFormat.Border.Color.Color := CBP.Sides[xcbpsInsideVert].Color;
        FXLS.CmdFormat.Border.Style := CBP.Sides[xcbpsInsideVert].Style;
        FXLS.CmdFormat.Border.Side[cbsInsideVert] := True;
      end;
      if CBP.Sides[xcbpsInsideHoriz].Changed then begin
        FXLS.CmdFormat.Border.Color.Color := CBP.Sides[xcbpsInsideHoriz].Color;
        FXLS.CmdFormat.Border.Style := CBP.Sides[xcbpsInsideHoriz].Style;
        FXLS.CmdFormat.Border.Side[cbsInsideHoriz] := True;
      end;
    end;

    if FFillChanged then begin
      if TExcelColorPicker(FControls[xuicFillColor].Control).ExcelColor.ColorType = exctUnassigned then
        FXLS.CmdFormat.Fill.NoColor := True
      else
        FXLS.CmdFormat.Fill.BackgroundColor.Color := TExcelColorPicker(FControls[xuicFillColor].Control).ExcelColor;
    end;

    FXLS.CmdFormat.Apply;
  end;
end;

procedure TXLSBookUI.NumberCurrencyCloseUp(ASender: TObject);
begin
  UpdateNumFmt;
end;

procedure TXLSBookUI.NumberCustomChanged(ASender: TObject);
begin
  FXLS.CmdFormat.Number.Format := TEdit(FControls[xuicFormatNumberCustom].Control).Text;
  UpdateNumberSample;
end;

procedure TXLSBookUI.NumberTypeClick(ASender: TObject);
begin
  UpdateNumFmt;
end;

procedure TXLSBookUI.NumberDecimalsChanged(ASender: TObject);
begin
  UpdateNumFmt;
end;

procedure TXLSBookUI.NumberNegNumbersClick(ASender: TObject);
begin
  case TListBox(ASender).ItemIndex of
    0: begin
      FXLS.CmdFormat.Number.NegativeColor := clDefault;
      FXLS.CmdFormat.Number.NegativeColorShowSign := False;
    end;
    1: begin
      FXLS.CmdFormat.Number.NegativeColor := clRed;
      FXLS.CmdFormat.Number.NegativeColorShowSign := False;
    end;
    2: begin // Not correct. There shall be a space before the number
      FXLS.CmdFormat.Number.NegativeColor := clDefault;
      FXLS.CmdFormat.Number.NegativeColorShowSign := False;
    end;
    3: begin
      FXLS.CmdFormat.Number.NegativeColor := clRed;
      FXLS.CmdFormat.Number.NegativeColorShowSign := True;
    end;
  end;
end;

procedure TXLSBookUI.NumberNegNumbersDrawItem(Control: TWinControl; Index: Integer; Rect: TRect; State: TOwnerDrawState);
begin
  if Odd(Index) then
    TListBox(Control).Canvas.Font.Color := clRed;

  TListBox(Control).Canvas.FillRect(Rect);
  TListBox(Control).Canvas.TextOut(Rect.Left, Rect.Top, (Control as TListBox).Items[Index]);
  if odFocused In State then begin
    TListBox(Control).Canvas.Brush.Color := clWindow;
    TListBox(Control).Canvas.DrawFocusRect(Rect);
  end;
end;

procedure TXLSBookUI.Prepare;
var
  Fill: TXc12Fill;
begin
  TListBox(FControls[xuicFormatNumberCategory].Control).ItemIndex := 0;

  TUpDown(FControls[xuicFormatNumberNumberDecimals].Control).Position := 2;

  FSampleNum := FSheet.MMUCells.GetFloat(FSheet.SelectedAreas.ActiveCol,FSheet.SelectedAreas.ActiveRow);

  TEdit(FControls[xuicFormatFontFontEd].Control).Text := FXLS.CmdFormat.Font.Name;

  UpdateFontStyle;

  TEdit(FControls[xuicFormatFontSizeEd].Control).Text := IntToStr(Round(FXLS.CmdFormat.Font.Size));

  FXLS.CmdFormat.Font.ClearAssigned;

  Fill := FindCommonFill;
  if Fill <> Nil then begin
    TExcelColorPicker(FControls[xuicFillColor].Control).ExcelColor := Fill.FgColor;
    TShape(FControls[xuicFillSample].Control).Brush.Color := Fill.FgColor.ARGB;
  end
  else
    TShape(FControls[xuicFillSample].Control).Brush.Style := bsClear;
end;

procedure TXLSBookUI.ProtectionHiddenClick(ASender: TObject);
begin
  FXLS.CmdFormat.Protection.Locked := TCheckBox(ASender).Checked;
end;

procedure TXLSBookUI.ProtectionLockedClick(ASender: TObject);
begin
  FXLS.CmdFormat.Protection.Hidden := TCheckBox(ASender).Checked;
end;

procedure TXLSBookUI.RegisterControl(AControl: TControl; const AId: TXLSBookUIControlId);
var
  i: integer;
  W,H: integer;
begin
  if not (AControl is FControls[AId].ControlClass) then
    raise XLSRWException.Create('Control is not of the excpected type: ' + FControls[AId].ControlClass.ClassName);

  FControls[AId].Control := AControl;

  case AId of
    xuicFormat: begin
      TForm(AControl).OnActivate := DialogActivate;
      TForm(AControl).OnCloseQuery := DialogCloseQuery;
    end;

    xuicFormatNumberCategory          : TListBox(AControl).OnClick := NumberCategoryClick;
    xuicFormatNumberNumberDecimals    : TEdit(TUpDown(AControl).Associate).OnChange := NumberDecimalsChanged;
    xuicFormatNumberNumberUse1000Sep  : TCheckBox(AControl).OnClick := NumberChanged;
    xuicFormatNumberNumberNegNumbers  : begin
      TListBox(AControl).Style := lbOwnerDrawFixed;
      TListBox(AControl).OnDrawItem := NumberNegNumbersDrawItem;
      TListBox(AControl).OnClick := NumberNegNumbersClick;
    end;
    xuicFormatNumberCurrencyDecimals  : ;
    xuicFormatNumberCurrencySymbol    : begin
      FillCurrencyStrings(TComboBox(AControl).Items);
      TComboBox(AControl).Text :=  FormatSettings.CurrencyString;
      TComboBox(AControl).OnCloseUp := NumberCurrencyCloseUp;
    end;
    xuicFormatNumberCurrencyNegNumbers: ;

    xuicFormatNumberDate              : TListBox(AControl).OnClick := NumberTypeClick;

    xuicFormatNumberTime              : TListBox(AControl).OnClick := NumberTypeClick;

    xuicFormatNumberCustom            : TEdit(AControl).OnChange := NumberCustomChanged;

    xuicFormatAlignOrientation        : begin
      TPaintBox(AControl).OnMouseDown := AlignOrientationMouseDown;
      TPaintBox(AControl).OnMouseMove := AlignOrientationMouseMove;
      TPaintBox(AControl).OnMouseUp := AlignOrientationMouseUp;
      TPaintBox(AControl).OnPaint := AlignOrientationPaint;
    end;
    xuicFormatAlignDegrees            : begin
      TUpDown(AControl).Position := FXLS.CmdFormat.Alignment.Rotation;
      TEdit(TUpDown(AControl).Associate).OnChange := AlignOrientationDegreesChanged;
    end;
    xuicFormatAlignHoriz              : begin
      TComboBox(AControl).ItemIndex := Integer(FXLS.CmdFormat.Alignment.Horizontal);
      TComboBox(AControl).OnCloseUp := AlignChanged;
    end;
    xuicFormatAlignVert               : begin
      TComboBox(AControl).ItemIndex := Integer(FXLS.CmdFormat.Alignment.Vertical);
      TComboBox(AControl).OnCloseUp := AlignChanged;
    end;
    xuicFormatAlignIndent             : begin
      TUpDown(AControl).Position := FXLS.CmdFormat.Alignment.Indent;
      TEdit(TUpDown(AControl).Associate).OnChange := AlignChanged;
    end;
    xuicFormatAlignJustifyDistributed : begin
      TCheckBox(AControl).Checked := FXLS.CmdFormat.Alignment.JustifyDistributed;
      TCheckBox(AControl).OnClick := AlignChanged;
    end;
    xuicFormatAlignWrapText           : begin
      TCheckBox(AControl).Checked := FXLS.CmdFormat.Alignment.WrapText;
      TCheckBox(AControl).OnClick := AlignChanged;
    end;
    xuicFormatAlignShrinkToFit        : begin
      TCheckBox(AControl).Checked := FXLS.CmdFormat.Alignment.ShrinkToFit;
      TCheckBox(AControl).OnClick := AlignChanged;
    end;
    xuicFormatAlignMergeCells         : begin
      TCheckBox(AControl).Checked := FXLS.CmdFormat.Alignment.MergeCells;
      TCheckBox(AControl).OnClick := AlignChanged;
    end;
    xuicFormatAlignTextDirection      : begin
      TComboBox(AControl).ItemIndex := Integer(FXLS.CmdFormat.Alignment.TextDirection);
      TComboBox(AControl).OnCloseUp := AlignChanged;
    end;

    xuicFormatFontFontEd              : begin      // TEdit
      TEdit(AControl).OnChange := FontFontEdChanged;
    end;
    xuicFormatFontFont                : begin
      TListBox(AControl).Items.BeginUpdate;

      GetAvailableFonts(TListBox(AControl).Items);
      TListBox(AControl).Sorted := True;

      TListBox(AControl).Items.EndUpdate;

      TListBox(AControl).OnDrawItem := FontFontDrawItem;
      TListBox(AControl).OnClick := FontFontClick;
    end;
    xuicFormatFontStyle               : begin       // TListBox
      TListBox(AControl).OnClick := FontStyleClick;
    end;
    xuicFormatFontStyleEd             : begin       // TEdit
    end;
    xuicFormatFontSize                : begin       // TListBox
      for i := 0 to High(XSSDefaultFontSizes) do
        TListBox(AControl).Items.Add(IntToStr(XSSDefaultFontSizes[i]));
      TListBox(AControl).OnClick := FontSizeClick;
    end;
    xuicFormatFontSizeEd              : begin       // TEdit
      TEdit(AControl).OnChange := FontSizeEdChanged;
    end;
    xuicFormatFontUnderline           : begin       // TComboBox
      TListBox(AControl).ItemIndex := Integer(FXLS.CmdFormat.Font.Underline);
    end;
    xuicFormatFontColor               : begin
      TXBookColorComboBox(AControl).Color := FXLS.CmdFormat.Font.Color.Color;
      TXBookColorComboBox(AControl).OnSelect := FontColorChanged;
    end;
    xuicFormatFontNormal              : begin       // TButton
      TButton(AControl).OnClick := FontNormalFontClick;
    end;
    xuicFormatFontStriketrough        : begin       // TCheckBox
      TCheckBox(AControl).Checked := xfsStrikeOut in FXLS.CmdFormat.Font.Style;
      TCheckBox(AControl).OnClick := FontStriketroughClick;
    end;
    xuicFormatFontSuperscript         : begin       // TCheckBox
      TCheckBox(AControl).Checked := FXLS.CmdFormat.Font.SubSuperscript = xssSuperscript;
      TCheckBox(AControl).OnClick := FontSuperscriptClick;
    end;
    xuicFormatFontSubscript           : begin       // TCheckBox
      TCheckBox(AControl).Checked := FXLS.CmdFormat.Font.SubSuperscript = xssSubscript;
      TCheckBox(AControl).OnClick := FontSubscriptClick;
    end;
    xuicFormatFontPreview             : TPaintBox(AControl).OnPaint := FontSamplePaint;

    xuicFormatBorderColor             : TXBookColorComboBox(AControl).OnSelect := BorderColorSelect;
    xuicFormatBorderStyle             : TXBookCellBorderStylePicker(AControl).OnSelect := BorderStyleSelect;
    xuicFormatBorderBorders: begin
      TXBookCellBorderPicker(AControl).OnSelect := BorderBorderSelect;

      W := FSheet.SelectedAreas.MaxWidth;
      H := FSheet.SelectedAreas.MaxHeight;
      if (W = 1) and (H = 1) then
        TXBookCellBorderPicker(AControl).Style := xcbpsCell
      else if (W > 1) and (H = 1) then
        TXBookCellBorderPicker(AControl).Style := xcbpsRow
      else if (W = 1) and (H > 1) then
        TXBookCellBorderPicker(AControl).Style := xcbpsColumn
      else
        TXBookCellBorderPicker(AControl).Style := xcbpsArea;
    end;

   xuicFillColor                      : TExcelColorPicker(AControl).OnClick := FillColorClick;
   xuicFillNoColor                    : TButton(AControl).OnClick := FillNoColorClick;

   xuicProtectionLocked               : TCheckBox(AControl).OnClick := ProtectionLockedClick;
   xuicProtectionHidden               : TCheckBox(AControl).OnClick := ProtectionHiddenClick;
  end;
end;

function TXLSBookUI.SearchStrings(const AText: AxUCString; AStrings: TSTrings): integer;
var
  S: AxUCString;
begin
  S := AnsiLowercase(AText);
  for Result := 0 to AStrings.Count - 1 do begin
    if Pos(S,AnsiLowercase(AStrings[Result])) > 0 then
      Exit;
  end;
  Result := -1;
end;

procedure TXLSBookUI.SelectNumberType(const AIndex: integer);
var
  i: integer;
  LB: TListBox;
begin
  for i := UI_NUM_FIRST to UI_NUM_LAST do
    ShowControls(i,False);

  ShowControls(AIndex,True);

  case AIndex of
    UI_NUM_DATE: begin
      LB := TListBox(FControls[xuicFormatNumberDate].Control);
      LB.Clear;
      for i := 14 to 22 do
        LB.Items.Add(ExcelStandardNumFormats[i]);
      LB.ItemIndex := 0;
      LB.OnClick := NumberTypeClick;
    end;
    UI_NUM_TIME: begin
      LB := TListBox(FControls[xuicFormatNumberTime].Control);
      LB.Clear;
      for i := 45 to 47 do
        LB.Items.Add(ExcelStandardNumFormats[i]);
      LB.ItemIndex := 0;
      LB.OnClick := NumberTypeClick;
    end;
    UI_NUM_FRACTION: begin
      LB := TListBox(FControls[xuicFormatNumberTime].Control);
      LB.Clear;

      LB.Items.AddObject('Up to one digit (1/4)',TObject(xnffOneDigit));
      LB.Items.AddObject('Up to two digits (21/25)',TObject(xnffTwoDigits));
      LB.Items.AddObject('Up to three digit (312/943)',TObject(xnffThreeDigits));
      LB.Items.AddObject('As halves (1/2)',TObject(xnffAsHalves));
      LB.Items.AddObject('As quarters (2/4)',TObject(xnffAsQuarters));
      LB.Items.AddObject('As eights (4/8)',TObject(xnffAsEights));
      LB.Items.AddObject('As sixtheenths (8/16)',TObject(xnffAsSixtheenths));
      LB.Items.AddObject('As tenths (3/10)',TObject(xnffAsTenths));
      LB.Items.AddObject('As hundredths (30/100)',TObject(xnffAsHundredths));

      LB.ItemIndex := 0;
      LB.OnClick := NumberTypeClick;
    end;
  end;

  UpdateNumFmt;
end;

procedure TXLSBookUI.ShowControls(const AGroup: integer; const AVisible: boolean);
var
  i: TXLSBookUIControlId;
begin
  for i := Low(TXLSBookUIControlId) to High(TXLSBookUIControlId) do begin
    if (FControls[i].Group = AGroup) then begin
      FControls[i].Control.Visible := AVisible;
      if FControls[i].Control is TUpDown then
        TUpDown(FControls[i].Control).Associate.Visible := AVisible;
    end;
  end;
end;

procedure TXLSBookUI.UpdateFontSample;
const
  SampleStr = 'AaBbCcYyZz';
var
  PB: TPaintBox;
  X,Y: integer;
  W,H: integer;
  TM: TEXTMETRIC;
begin
  PB := TPaintBox(FControls[xuicFormatFontPreview].Control);

  FXLS.CmdFormat.Font.CopyToTFont(PB.Canvas.Font);

  GetTextMetrics(PB.Canvas.Handle,TM);

  PB.Canvas.Pen.Color := clBlack;
  PB.Canvas.Brush.Style := bsSolid;
  PB.Canvas.Brush.Color := clWhite;

  PB.Canvas.Rectangle(0,0,PB.Width,PB.Height);

  PB.Canvas.Brush.Style := bsClear;

  W := PB.Canvas.TextWidth(SampleStr);
  H := PB.Canvas.TextHeight(SampleStr);
  X := (PB.Width div 2) - (W div 2);
  Y := (PB.Height div 2) - (H div 2);

  PB.Canvas.TextOut(X,Y,SampleStr);

  Dec(Y,TM.tmDescent);
  PB.Canvas.MoveTo(1,Y + H);
  PB.Canvas.LineTo(X - 5,Y + H);
  PB.Canvas.MoveTo(X + W + 5,Y + H);
  PB.Canvas.LineTo(PB.Width,Y + H);

  PB.Canvas.Rectangle(0,0,PB.Width,PB.Height);
end;

procedure TXLSBookUI.UpdateFontStyle;
var
  Ed: TEdit;
  LB: TListBox;
begin
  Ed := TEdit(FControls[xuicFormatFontStyleEd].Control);
  LB := TListBox(FControls[xuicFormatFontStyle].Control);
  if FXLS.CmdFormat.Font.Style = [] then
    LB.ItemIndex := 0
  else if (xfsItalic in FXLS.CmdFormat.Font.Style) and (xfsBold in FXLS.CmdFormat.Font.Style) then
    LB.ItemIndex := 3
  else if xfsItalic in FXLS.CmdFormat.Font.Style then
    LB.ItemIndex := 1
  else if xfsBold in FXLS.CmdFormat.Font.Style then
    LB.ItemIndex := 2;
  Ed.Text := LB.Items[LB.ItemIndex];
end;

procedure TXLSBookUI.UpdateNegNumSample;
var
  i: integer;
  Selected: integer;
  LB: TListBox;
  S: AxUCString;
  Dec: AxUCString;
begin
  LB := TListBox(FControls[xuicFormatNumberNumberNegNumbers].Control);
  Selected := LB.ItemIndex;
  LB.Clear;

  if FXLS.CmdFormat.Number.Thousands then
    S := AxUCString('1') + FormatSettings.ThousandSeparator + AxUCString('234')
  else
    S := '1234';

  if FDecimals > 0 then begin
    for i := 0 to FDecimals - 1 do
      Dec := StrDigits[(i mod 10) + 1] + Dec;

    S := S + FormatSettings.DecimalSeparator + Dec;
  end;

  LB.Items.Add('-' + S);
  LB.Items.Add(' ' + S);
  LB.Items.Add('-' + S);
  LB.Items.Add('-' + S);

  if Selected < 0 then
    Selected := 0;
  LB.ItemIndex := Selected;
end;

procedure TXLSBookUI.UpdateNumFmt;
var
  i: integer;
  CB: TComboBox;
  LB: TListBox;
begin
  i := TListBox(FControls[xuicFormatNumberCategory].Control).ItemIndex;

  FXLS.CmdFormat.Number.Clear;
  case i of
    UI_NUM_GENERAL: ;
    UI_NUM_NUMBER: begin
      FDecimals := TUpDown(FControls[xuicFormatNumberNumberDecimals].Control).Position;
      FXLS.CmdFormat.Number.Decimals := FDecimals;
      FXLS.CmdFormat.Number.Thousands := TCheckBox(FControls[xuicFormatNumberNumberUse1000Sep].Control).Checked;

      UpdateNegNumSample;
    end;
    UI_NUM_CURRENCY: begin
      FDecimals := TUpDown(FControls[xuicFormatNumberNumberDecimals].Control).Position;

      CB := TComboBox(FControls[xuicFormatNumberCurrencySymbol].Control);
      if CB.ItemIndex < 0 then
        FXLS.CmdFormat.Number.CurrencySymbolLCID := 0
      else if CB.ItemIndex = 0 then
        FXLS.CmdFormat.Number.CurrencySymbolLCID := -1
      else
        FXLS.CmdFormat.Number.CurrencySymbolLCID := Integer(CB.Items.Objects[CB.ItemIndex]);

      FXLS.CmdFormat.Number.Decimals := FDecimals;
      FXLS.CmdFormat.Number.Thousands := True;

      UpdateNegNumSample;
    end;
    UI_NUM_ACCOUNTING: ;
    UI_NUM_DATE: begin
      FXLS.CmdFormat.Number.Format := TListBox(FControls[xuicFormatNumberDate].Control).Items[TListBox(FControls[xuicFormatNumberDate].Control).ItemIndex];
    end;
    UI_NUM_TIME: begin
      FXLS.CmdFormat.Number.Format := TListBox(FControls[xuicFormatNumberTime].Control).Items[TListBox(FControls[xuicFormatNumberTime].Control).ItemIndex];
    end;
    UI_NUM_PERCENTAGE: begin
      FDecimals := TUpDown(FControls[xuicFormatNumberNumberDecimals].Control).Position;

      FXLS.CmdFormat.Number.Decimals := FDecimals;
      FXLS.CmdFormat.Number.Percentage := True;
    end;
    UI_NUM_FRACTION: begin
      LB := TListBox(FControls[xuicFormatNumberFractionType].Control);
      FXLS.CmdFormat.Number.Fractions := TXLSNumFmtFractions(LB.Items.Objects[LB.ItemIndex]);
    end;
    UI_NUM_SCIENTIFIC: begin
      FXLS.CmdFormat.Number.Scientific := True;
    end;
    UI_NUM_TEXT: begin
      FXLS.CmdFormat.Number.Format := '@';
    end;
    UI_NUM_SPECIAL: ;
    UI_NUM_CUSTOM: begin
      TEdit(FControls[xuicFormatNumberCustom].Control).Text := FXLS.CmdFormat.Number.Format;
    end;
  end;
  UpdateNumberSample;
end;

procedure TXLSBookUI.UpdateNumberSample;
var
  lbl: TLabel;
  Mask: TExcelMask;
begin
  lbl := TLabel(FControls[xuicFormatNumberSample].Control);
  if lbl <> Nil then begin
    if FSampleNum = 0 then
      lbl.Caption := ''
    else begin
      Mask := TExcelMask.Create;
      try
        Mask.Mask := FXLS.CmdFormat.Number.Format;

        if Mask.Mask <> '' then
          lbl.Caption := Mask.FormatNumber(FSampleNum)
        else
          lbl.Caption := FloatToStr(FSampleNum);
      finally
        Mask.Free;
      end;
    end;
  end;
end;

end.
