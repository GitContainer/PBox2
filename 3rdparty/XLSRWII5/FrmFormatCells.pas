unit FrmFormatCells;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics,
  Controls, Forms, Dialogs, StdCtrls, XLSBook2, ComCtrls, XBookUI2,
  ExtCtrls, Math, XLSUtils5, XBookStdComponents, ImgList, Buttons,
  Menus, ExcelColorPicker, System.ImageList;

type
  TfrmFmtCells = class(TForm)
    Button1: TButton;
    Button2: TButton;
    PageControl: TPageControl;
    tsNumber: TTabSheet;
    Label1: TLabel;
    lbNumCategory: TListBox;
    GroupBox1: TGroupBox;
    lblNumSample: TLabel;
    lbNumNegNumbers: TListBox;
    lblNumNegNumbers: TLabel;
    cbNumUse1000Sep: TCheckBox;
    edNumDecimals: TEdit;
    lblNumDecimal: TLabel;
    udNumDecimals: TUpDown;
    cbNumCurrencySymbol: TComboBox;
    lblNumCurrencySymbol: TLabel;
    lbNumType: TListBox;
    lblNumType: TLabel;
    edNumberCustom: TEdit;
    tsAlignment: TTabSheet;
    pbAlignOrientation: TPaintBox;
    udAlignDegrees: TUpDown;
    edAlignDegrees: TEdit;
    Label2: TLabel;
    cbAlignTextHoriz: TComboBox;
    Label3: TLabel;
    Label4: TLabel;
    cbAlignTextVert: TComboBox;
    udAlignTextIndent: TUpDown;
    edAlignTextIndent: TEdit;
    Label5: TLabel;
    cbAlignJustifyDistr: TCheckBox;
    cbAlignWrapText: TCheckBox;
    cbAlignShrinkToFit: TCheckBox;
    cbAlignMergeCells: TCheckBox;
    Label6: TLabel;
    cbAlignTextDirection: TComboBox;
    Label7: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    tsFont: TTabSheet;
    lbFontFont: TListBox;
    Label11: TLabel;
    edFontFont: TEdit;
    lbFontStyle: TListBox;
    edFontStyle: TEdit;
    lbFontSize: TListBox;
    edFontSize: TEdit;
    Label12: TLabel;
    Label13: TLabel;
    Label14: TLabel;
    cbFontUnderline: TComboBox;
    Label15: TLabel;
    btnFontNormal: TButton;
    Label16: TLabel;
    cbFontStriketrough: TCheckBox;
    cbFontSuperscript: TCheckBox;
    cbFontSubscript: TCheckBox;
    Label17: TLabel;
    pbFontPreview: TPaintBox;
    tsBorder: TTabSheet;
    ImageList: TImageList;
    ImageList_27: TImageList;
    Label18: TLabel;
    Label19: TLabel;
    Label20: TLabel;
    Label21: TLabel;
    tsFill: TTabSheet;
    Label22: TLabel;
    tsProtection: TTabSheet;
    cbProtectionLocked: TCheckBox;
    cbProtectionHidden: TCheckBox;
    xccbFontColor: TXBookColorComboBox;
    xbccbBorder: TXBookColorComboBox;
    xcbspBorder: TXBookCellBorderStylePicker;
    ecpFillColors: TExcelColorPicker;
    ecpFillStdColors: TExcelColorPicker;
    Shape1: TShape;
    btnFillNoColor: TButton;
    shpFillSample: TShape;
    Label23: TLabel;
    Shape2: TShape;
    btnPresetNone: TBitBtn;
    btnPresetOutline: TBitBtn;
    xcbpBorder: TXBookCellBorderPicker;
    btnBorderDiagDown: TBitBtn;
    btnBorderRight: TBitBtn;
    btnBorderInsideVert: TBitBtn;
    btnBorderLeft: TBitBtn;
    btnBorderDiagUp: TBitBtn;
    btnBorderBottom: TBitBtn;
    btnBorderInsideHoriz: TBitBtn;
    btnBorderTop: TBitBtn;
    btnPresetInside: TBitBtn;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
  private
    FXSS: TXLSSpreadSheet;

    procedure AddLocalizedTexts;
    procedure RegisterControls;
  public
    function Execute(AXSS: TXLSSpreadSheet; APage: integer = -1): boolean;
  end;

implementation

{$R *.dfm}

{ TfrmFmtCells }

procedure TfrmFmtCells.AddLocalizedTexts;
begin
  // The listbox must have exactly this number of items. The name doesn't matter.
  lbNumCategory.Items.Add('General');
  lbNumCategory.Items.Add('Number');
  lbNumCategory.Items.Add('Currency');
  lbNumCategory.Items.Add('Accounting');
  lbNumCategory.Items.Add('Date');
  lbNumCategory.Items.Add('Time');
  lbNumCategory.Items.Add('Percentage');
  lbNumCategory.Items.Add('Fraction');
  lbNumCategory.Items.Add('Scientific');
  lbNumCategory.Items.Add('Text');
  lbNumCategory.Items.Add('Special');
  lbNumCategory.Items.Add('Custom');

  cbNumCurrencySymbol.Items.Add('None');

  cbAlignTextHoriz.Items.Add('General');
  cbAlignTextHoriz.Items.Add('Left');
  cbAlignTextHoriz.Items.Add('Center');
  cbAlignTextHoriz.Items.Add('Right');
  cbAlignTextHoriz.Items.Add('Fill');
  cbAlignTextHoriz.Items.Add('Justify');
  cbAlignTextHoriz.Items.Add('Center across sections');
  cbAlignTextHoriz.Items.Add('Distributed');

  cbAlignTextVert.Items.Add('Top');
  cbAlignTextVert.Items.Add('Center');
  cbAlignTextVert.Items.Add('Bottom');
  cbAlignTextVert.Items.Add('Justify');
  cbAlignTextVert.Items.Add('Distributed');

  cbAlignTextDirection.Items.Add('Context');
  cbAlignTextDirection.Items.Add('Left-to-right');
  cbAlignTextDirection.Items.Add('Right-to-left');

  lbFontStyle.Items.Add('Regular');
  lbFontStyle.Items.Add('Italic');
  lbFontStyle.Items.Add('Bold');
  lbFontStyle.Items.Add('Bold Italic');

  cbFontUnderline.Items.Add('None');
  cbFontUnderline.Items.Add('Single');
  cbFontUnderline.Items.Add('Double');
  cbFontUnderline.Items.Add('Single Accounting');
  cbFontUnderline.Items.Add('Double Accounting');
end;

function TfrmFmtCells.Execute(AXSS: TXLSSpreadSheet; APage: integer = -1): boolean;
begin
  FXSS := AXSS;

  AddLocalizedTexts;
  RegisterControls;

  if (APage >= 0) and (APage < PageControl.PageCount) then
    PageControl.ActivePageIndex := APage;

  ShowModal;

  Result := ModalResult = mrOk;
  if Result and FXSS.XLS.CmdFormat.Alignment.MergeCells then
    FXSS.XLSSheet.MergeCells;
end;

procedure TfrmFmtCells.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Action := caFree;
end;

procedure TfrmFmtCells.RegisterControls;
begin
  FXSS.UI.RegisterControl(Self,xuicFormat);
  FXSS.UI.RegisterControl(lbNumCategory,xuicFormatNumberCategory);
  FXSS.UI.RegisterControl(lblNumSample,xuicFormatNumberSample);

  FXSS.UI.RegisterControl(udNumDecimals,xuicFormatNumberNumberDecimals);
  FXSS.UI.RegisterControl(lblNumDecimal,xuicFormatNumberNumberDecimals_Txt);
  FXSS.UI.RegisterControl(cbNumUse1000Sep,xuicFormatNumberNumberUse1000Sep);
  FXSS.UI.RegisterControl(lbNumNegNumbers,xuicFormatNumberNumberNegNumbers);
  FXSS.UI.RegisterControl(lblNumNegNumbers,xuicFormatNumberNumberNegNumbers_Txt);

  FXSS.UI.RegisterControl(udNumDecimals,xuicFormatNumberCurrencyDecimals);
  FXSS.UI.RegisterControl(lblNumDecimal,xuicFormatNumberCurrencyDecimals_Txt);
  FXSS.UI.RegisterControl(cbNumCurrencySymbol,xuicFormatNumberCurrencySymbol);
  FXSS.UI.RegisterControl(lblNumCurrencySymbol,xuicFormatNumberCurrencySymbol_Txt);
  FXSS.UI.RegisterControl(lbNumNegNumbers,xuicFormatNumberCurrencyNegNumbers);
  FXSS.UI.RegisterControl(lblNumNegNumbers,xuicFormatNumberCurrencyNegNumbers_Txt);

  FXSS.UI.RegisterControl(udNumDecimals,xuicFormatNumberAccountingDecimals);
  FXSS.UI.RegisterControl(lblNumDecimal,xuicFormatNumberAccountingDecimals_Txt);
  FXSS.UI.RegisterControl(cbNumCurrencySymbol,xuicFormatNumberAccountingSymbol);
  FXSS.UI.RegisterControl(lblNumCurrencySymbol,xuicFormatNumberAccountingSymbol_Txt);

  FXSS.UI.RegisterControl(lbNumType,xuicFormatNumberDate);
  FXSS.UI.RegisterControl(lblNumType,xuicFormatNumberDateType_Txt);

  FXSS.UI.RegisterControl(lbNumType,xuicFormatNumberTime);
  FXSS.UI.RegisterControl(lblNumType,xuicFormatNumberTimeType_Txt);

  FXSS.UI.RegisterControl(udNumDecimals,xuicFormatNumberPercentageDecimals);
  FXSS.UI.RegisterControl(lblNumDecimal,xuicFormatNumberPercentageDecimals_Txt);

  FXSS.UI.RegisterControl(lbNumType,xuicFormatNumberFractionType);
  FXSS.UI.RegisterControl(lblNumType,xuicFormatNumberFractionType_Txt);

  FXSS.UI.RegisterControl(udNumDecimals,xuicFormatNumberScientificDecimals);
  FXSS.UI.RegisterControl(lblNumDecimal,xuicFormatNumberScientificDecimals_Txt);

  FXSS.UI.RegisterControl(lbNumType,xuicFormatNumberSpecialType);
  FXSS.UI.RegisterControl(lblNumType,xuicFormatNumberSpecialType_Txt);

  FXSS.UI.RegisterControl(edNumberCustom,xuicFormatNumberCustom);
  FXSS.UI.RegisterControl(lblNumType,xuicFormatNumberCustomType_Txt);

  FXSS.UI.RegisterControl(pbAlignOrientation,xuicFormatAlignOrientation);
  FXSS.UI.RegisterControl(udAlignDegrees,xuicFormatAlignDegrees);
  FXSS.UI.RegisterControl(cbAlignTextHoriz,xuicFormatAlignHoriz);
  FXSS.UI.RegisterControl(cbAlignTextVert,xuicFormatAlignVert);
  FXSS.UI.RegisterControl(udAlignTextIndent,xuicFormatAlignIndent);
  FXSS.UI.RegisterControl(cbAlignJustifyDistr,xuicFormatAlignJustifyDistributed);
  FXSS.UI.RegisterControl(cbAlignWrapText,xuicFormatAlignWrapText);
  FXSS.UI.RegisterControl(cbAlignShrinkToFit,xuicFormatAlignShrinkToFit);
  FXSS.UI.RegisterControl(cbAlignMergeCells,xuicFormatAlignMergeCells);
  FXSS.UI.RegisterControl(cbAlignTextDirection,xuicFormatAlignTextDirection);


  FXSS.UI.RegisterControl(edFontFont,xuicFormatFontFontEd);
  FXSS.UI.RegisterControl(lbFontFont,xuicFormatFontFont);
  FXSS.UI.RegisterControl(edFontStyle,xuicFormatFontStyleEd);
  FXSS.UI.RegisterControl(lbFontStyle,xuicFormatFontStyle);
  FXSS.UI.RegisterControl(edFontSize,xuicFormatFontSizeEd);
  FXSS.UI.RegisterControl(lbFontSize,xuicFormatFontSize);
  FXSS.UI.RegisterControl(cbFontUnderline,xuicFormatFontUnderline);
  FXSS.UI.RegisterControl(xccbFontColor,xuicFormatFontColor);
  FXSS.UI.RegisterControl(btnFontNormal,xuicFormatFontNormal);
  FXSS.UI.RegisterControl(cbFontStriketrough,xuicFormatFontStriketrough);
  FXSS.UI.RegisterControl(cbFontSuperscript,xuicFormatFontSuperscript);
  FXSS.UI.RegisterControl(cbFontSubscript,xuicFormatFontSubscript);
  FXSS.UI.RegisterControl(pbFontPreview,xuicFormatFontPreview);

  FXSS.UI.RegisterControl(xbccbBorder,xuicFormatBorderColor);
  FXSS.UI.RegisterControl(xcbspBorder,xuicFormatBorderStyle);
  FXSS.UI.RegisterControl(xcbpBorder,xuicFormatBorderBorders);

  FXSS.UI.RegisterControl(ecpFillColors,xuicFillColor);
  FXSS.UI.RegisterControl(btnFillNoColor,xuicFillNoColor);
  FXSS.UI.RegisterControl(shpFillSample,xuicFillSample);

  FXSS.UI.RegisterControl(cbProtectionLocked,xuicProtectionLocked);
  FXSS.UI.RegisterControl(cbProtectionHidden,xuicProtectionHidden);
end;

procedure TfrmFmtCells.FormCreate(Sender: TObject);
begin
//  btnTest.Glyph.TransparentColor := clWhite;
//  btnTest.Glyph.TransparentMode := tmAuto;
//  btnTest.Glyph.Transparent := True;
end;

end.
