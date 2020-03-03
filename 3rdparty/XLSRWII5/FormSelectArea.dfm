object frmSelectArea: TfrmSelectArea
  Left = 61
  Top = 86
  Caption = 'Select Area'
  ClientHeight = 79
  ClientWidth = 382
  Color = clBtnFace
  Constraints.MinHeight = 118
  Constraints.MinWidth = 398
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poMainFormCenter
  OnClose = FormClose
  DesignSize = (
    382
    79)
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 12
    Top = 16
    Width = 23
    Height = 13
    Caption = 'Area'
  end
  object edArea: TEdit
    Left = 41
    Top = 13
    Width = 316
    Height = 21
    Anchors = [akLeft, akTop, akRight]
    TabOrder = 0
  end
  object btnOk: TButton
    Left = 200
    Top = 45
    Width = 75
    Height = 25
    Anchors = [akRight, akBottom]
    Caption = 'OK'
    Default = True
    ModalResult = 1
    TabOrder = 1
  end
  object btnCancel: TButton
    Left = 281
    Top = 45
    Width = 75
    Height = 25
    Anchors = [akRight, akBottom]
    Cancel = True
    Caption = 'Cancel'
    ModalResult = 2
    TabOrder = 2
  end
end
