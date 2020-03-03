object frmSelectChart: TfrmSelectChart
  Left = 0
  Top = 0
  Caption = 'Select Chart'
  ClientHeight = 87
  ClientWidth = 260
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poMainFormCenter
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 20
    Top = 16
    Width = 53
    Height = 13
    Caption = 'Chart style'
  end
  object cbStyle: TComboBox
    Left = 80
    Top = 12
    Width = 161
    Height = 21
    Style = csDropDownList
    TabOrder = 0
    Items.Strings = (
      'Bar'
      'Line'
      'Area'
      'Bubble'
      'Doughnut'
      'Pie'
      'Radar'
      'Scatter')
  end
  object Button1: TButton
    Left = 16
    Top = 48
    Width = 75
    Height = 25
    Caption = 'OK'
    Default = True
    ModalResult = 1
    TabOrder = 1
  end
  object Button2: TButton
    Left = 164
    Top = 48
    Width = 75
    Height = 25
    Cancel = True
    Caption = 'Button2'
    ModalResult = 2
    TabOrder = 2
  end
end
