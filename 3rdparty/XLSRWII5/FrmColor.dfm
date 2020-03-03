object frmSelectColor: TfrmSelectColor
  Left = 648
  Top = 128
  BorderStyle = bsToolWindow
  Caption = 'Select color'
  ClientHeight = 237
  ClientWidth = 204
  Color = clBtnFace
  Font.Charset = ANSI_CHARSET
  Font.Color = clBlack
  Font.Height = -15
  Font.Name = 'Calibri'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  PixelsPerInch = 96
  TextHeight = 18
  object ecpTheme: TExcelColorPicker
    Left = 12
    Top = 24
    Width = 180
    Height = 115
    ColorMode = ecpmExcel2007Theme
  end
  object ecpStandard: TExcelColorPicker
    Left = 12
    Top = 160
    Width = 180
    Height = 18
    ColorMode = ecpmExcel2007Standard
    LinkedPicker = ecpTheme
  end
  object Label1: TLabel
    Left = 12
    Top = 8
    Width = 73
    Height = 15
    Caption = 'Theme colors'
    Font.Charset = ANSI_CHARSET
    Font.Color = clNavy
    Font.Height = -13
    Font.Name = 'Calibri'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object Label2: TLabel
    Left = 12
    Top = 144
    Width = 84
    Height = 15
    Caption = 'Standard colors'
    Font.Charset = ANSI_CHARSET
    Font.Color = clNavy
    Font.Height = -13
    Font.Name = 'Calibri'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object Button1: TButton
    Left = 12
    Top = 204
    Width = 75
    Height = 25
    Caption = 'OK'
    Default = True
    ModalResult = 1
    TabOrder = 0
  end
  object Button2: TButton
    Left = 112
    Top = 204
    Width = 75
    Height = 25
    Cancel = True
    Caption = 'Cancel'
    Font.Charset = ANSI_CHARSET
    Font.Color = clBlack
    Font.Height = -15
    Font.Name = 'Calibri'
    Font.Style = []
    ModalResult = 2
    ParentFont = False
    TabOrder = 1
  end
end
