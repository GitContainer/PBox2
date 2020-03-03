object frmSelectValues: TfrmSelectValues
  Left = 0
  Top = 0
  BorderStyle = bsSizeToolWin
  Caption = 'SelectValues'
  ClientHeight = 372
  ClientWidth = 189
  Color = clBtnFace
  Constraints.MinHeight = 180
  Constraints.MinWidth = 180
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  DesignSize = (
    189
    372)
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 8
    Top = 4
    Width = 31
    Height = 13
    Caption = 'Values'
  end
  object clbValues: TCheckListBox
    Left = 4
    Top = 20
    Width = 182
    Height = 313
    OnClickCheck = clbValuesClickCheck
    Anchors = [akLeft, akTop, akRight, akBottom]
    ItemHeight = 13
    TabOrder = 0
  end
  object btnOk: TButton
    Left = 4
    Top = 340
    Width = 75
    Height = 25
    Anchors = [akLeft, akBottom]
    Caption = 'OK'
    Default = True
    ModalResult = 1
    TabOrder = 1
  end
  object btnCancel: TButton
    Left = 109
    Top = 340
    Width = 75
    Height = 25
    Anchors = [akRight, akBottom]
    Cancel = True
    Caption = 'Cancel'
    ModalResult = 2
    TabOrder = 2
  end
end
