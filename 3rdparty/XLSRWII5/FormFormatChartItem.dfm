object frmFormatChartItem: TfrmFormatChartItem
  Left = 1132
  Top = 880
  Width = 430
  Height = 462
  Caption = 'Format Chart Item'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  PixelsPerInch = 96
  TextHeight = 13
  object Button2: TButton
    Left = 16
    Top = 390
    Width = 75
    Height = 25
    Caption = 'Close'
    Default = True
    ModalResult = 1
    TabOrder = 0
  end
  object rgFill: TRadioGroup
    Left = 16
    Top = 13
    Width = 86
    Height = 132
    Caption = 'Fill'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = []
    ItemIndex = 3
    Items.Strings = (
      'None'
      'Solid'
      'Gradient'
      'Picture'
      'Automatic')
    ParentFont = False
    TabOrder = 1
    OnClick = rgFillClick
  end
  object rgLine: TRadioGroup
    Left = 13
    Top = 277
    Width = 86
    Height = 88
    Caption = 'Line'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = []
    ItemIndex = 2
    Items.Strings = (
      'None'
      'Solid'
      'Automatic')
    ParentFont = False
    TabOrder = 2
    OnClick = rgLineClick
  end
  object pcFill: TPageControl
    Left = 108
    Top = 15
    Width = 281
    Height = 264
    ActivePage = tsSolid
    TabOrder = 3
    object tsSolid: TTabSheet
      Caption = 'tsSolid'
      TabVisible = False
      object lblSolidTransp: TLabel
        Left = 21
        Top = 167
        Width = 66
        Height = 13
        Caption = 'Transparency'
      end
      object pbFillSolid: TPaintBox
        Left = 12
        Top = 44
        Width = 93
        Height = 93
        OnPaint = pbFillSolidPaint
      end
      object btnFillColor: TButton
        Left = 12
        Top = 8
        Width = 75
        Height = 25
        Caption = 'Color...'
        TabOrder = 0
        OnClick = btnFillColorClick
      end
      object tbSolidTransp: TTrackBar
        Left = 12
        Top = 186
        Width = 192
        Height = 20
        Max = 100
        TabOrder = 1
        TickStyle = tsNone
        OnChange = tbSolidTranspChange
      end
    end
    object tsGradient: TTabSheet
      Caption = 'tsGradient'
      ImageIndex = 1
      TabVisible = False
      object pbFillGrad: TPaintBox
        Left = 12
        Top = 52
        Width = 253
        Height = 112
        OnPaint = pbFillGradPaint
      end
      object Label1: TLabel
        Left = 3
        Top = -2
        Width = 70
        Height = 13
        Caption = 'Gradient stops'
      end
      object lblGradTransp: TLabel
        Left = 12
        Top = 177
        Width = 66
        Height = 13
        Caption = 'Transparency'
      end
      object btnGradAdd: TButton
        Left = 71
        Top = 15
        Width = 62
        Height = 25
        Caption = 'Add'
        TabOrder = 0
        OnClick = btnGradAddClick
      end
      object btnGradRemove: TButton
        Left = 139
        Top = 15
        Width = 62
        Height = 25
        Caption = 'Remove'
        TabOrder = 1
        OnClick = btnGradRemoveClick
      end
      object cbGradStops: TComboBox
        Left = 3
        Top = 17
        Width = 62
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 2
      end
      object btnGradColor: TButton
        Left = 208
        Top = 15
        Width = 62
        Height = 25
        Caption = 'Color...'
        TabOrder = 3
        OnClick = btnGradColorClick
      end
      object tbGradTransp: TTrackBar
        Left = 3
        Top = 191
        Width = 192
        Height = 20
        Max = 100
        TabOrder = 4
        TickStyle = tsNone
        OnChange = tbGradTranspChange
      end
    end
    object tsPicture: TTabSheet
      Caption = 'tsPicture'
      ImageIndex = 2
      TabVisible = False
      object Label2: TLabel
        Left = 6
        Top = 3
        Width = 42
        Height = 13
        Caption = 'Filename'
      end
      object lblPictTransp: TLabel
        Left = 13
        Top = 204
        Width = 66
        Height = 13
        Caption = 'Transparency'
      end
      object pbFillPict: TPaintBox
        Left = 6
        Top = 79
        Width = 264
        Height = 119
        OnPaint = pbFillPictPaint
      end
      object edFillFilename: TEdit
        Left = 6
        Top = 22
        Width = 235
        Height = 21
        TabOrder = 0
      end
      object Button4: TButton
        Left = 243
        Top = 20
        Width = 27
        Height = 25
        Caption = '...'
        TabOrder = 1
        OnClick = Button4Click
      end
      object cbFillPictTile: TCheckBox
        Left = 6
        Top = 56
        Width = 97
        Height = 17
        Caption = 'Tile picture'
        TabOrder = 2
        OnClick = cbFillPictTileClick
      end
      object tbPictTransp: TTrackBar
        Left = 3
        Top = 223
        Width = 192
        Height = 20
        Max = 100
        TabOrder = 3
        TickStyle = tsNone
        OnChange = tbPictTranspChange
      end
    end
  end
  object pcLine: TPageControl
    Left = 112
    Top = 279
    Width = 277
    Height = 86
    ActivePage = tsLineSolid
    TabOrder = 4
    object tsLineSolid: TTabSheet
      Caption = 'tsLineSolid'
      TabVisible = False
      object Label3: TLabel
        Left = 111
        Top = 10
        Width = 28
        Height = 13
        Caption = 'Width'
      end
      object Label4: TLabel
        Left = 214
        Top = 10
        Width = 29
        Height = 13
        Caption = 'points'
      end
      object pbLineSolid: TPaintBox
        Left = 3
        Top = 32
        Width = 259
        Height = 41
        OnPaint = pbLineSolidPaint
      end
      object udLineWidth: TUpDown
        Left = 192
        Top = 6
        Width = 16
        Height = 21
        Associate = edLineWidth
        Max = 72
        TabOrder = 0
      end
      object edLineWidth: TEdit
        Left = 145
        Top = 6
        Width = 47
        Height = 21
        TabOrder = 1
        Text = '0'
        OnChange = edLineWidthChange
      end
      object Button3: TButton
        Left = 3
        Top = 3
        Width = 75
        Height = 25
        Caption = 'Color...'
        TabOrder = 2
        OnClick = Button3Click
      end
    end
  end
  object dlgOpenPict: TOpenPictureDialog
    Left = 24
    Top = 168
  end
end
