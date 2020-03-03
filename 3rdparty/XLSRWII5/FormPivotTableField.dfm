object frmPivotTableField: TfrmPivotTableField
  Left = 0
  Top = 0
  BorderStyle = bsDialog
  Caption = 'Field Settings'
  ClientHeight = 329
  ClientWidth = 315
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
  object lblSourceName: TLabel
    Left = 8
    Top = 8
    Width = 66
    Height = 13
    Caption = 'Source name:'
  end
  object Label1: TLabel
    Left = 8
    Top = 32
    Width = 69
    Height = 13
    Caption = 'Custom name:'
  end
  object Label2: TLabel
    Left = 8
    Top = 72
    Width = 54
    Height = 13
    Caption = 'Subtotals'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object Label3: TLabel
    Left = 12
    Top = 157
    Width = 141
    Height = 13
    Caption = 'Select one or more functions:'
  end
  object btnOk: TButton
    Left = 141
    Top = 291
    Width = 75
    Height = 25
    Caption = 'OK'
    Default = True
    ModalResult = 1
    TabOrder = 3
  end
  object btnCancel: TButton
    Left = 222
    Top = 291
    Width = 75
    Height = 25
    Cancel = True
    Caption = 'Cancel'
    ModalResult = 2
    TabOrder = 4
  end
  object edCustomName: TEdit
    Left = 79
    Top = 29
    Width = 218
    Height = 21
    TabOrder = 0
  end
  object rgSubtotals: TRadioGroup
    Left = 0
    Top = 84
    Width = 105
    Height = 73
    Items.Strings = (
      'Automatic'
      'None'
      'Custom')
    TabOrder = 1
    OnClick = rgSubtotalsClick
  end
  object lbFunctions: TListBox
    Left = 8
    Top = 176
    Width = 208
    Height = 97
    Enabled = False
    ItemHeight = 13
    TabOrder = 2
    OnClick = lbFunctionsClick
  end
end
