object frmPivotTableOptions: TfrmPivotTableOptions
  Left = 0
  Top = 0
  BorderStyle = bsDialog
  Caption = 'Pivot Table Options'
  ClientHeight = 266
  ClientWidth = 338
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
    Left = 12
    Top = 16
    Width = 27
    Height = 13
    Caption = 'Name'
  end
  object btnOk: TButton
    Left = 12
    Top = 228
    Width = 75
    Height = 25
    Caption = 'OK'
    Default = True
    ModalResult = 1
    TabOrder = 3
  end
  object btnCancel: TButton
    Left = 252
    Top = 228
    Width = 75
    Height = 25
    Cancel = True
    Caption = 'Cancel'
    ModalResult = 2
    TabOrder = 4
  end
  object edName: TEdit
    Left = 44
    Top = 12
    Width = 217
    Height = 21
    TabOrder = 0
  end
  object cbTotalsCols: TCheckBox
    Left = 12
    Top = 88
    Width = 201
    Height = 17
    Caption = 'Show grand totals for columns'
    TabOrder = 2
  end
  object cbTotalsRows: TCheckBox
    Left = 12
    Top = 64
    Width = 205
    Height = 17
    Caption = 'Show grand totals for rows'
    TabOrder = 1
  end
end
