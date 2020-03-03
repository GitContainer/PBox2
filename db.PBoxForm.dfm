object frmPBox: TfrmPBox
  Left = 0
  Top = 0
  Caption = 'PBox '#24037#20855#31665' v2.0'
  ClientHeight = 723
  ClientWidth = 1127
  Color = clBtnFace
  Constraints.MinHeight = 649
  Constraints.MinWidth = 1020
  DoubleBuffered = True
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  KeyPreview = True
  OldCreateOrder = False
  Position = poScreenCenter
  OnActivate = FormActivate
  OnClose = FormClose
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  OnResize = FormResize
  PixelsPerInch = 96
  TextHeight = 13
  object pnlBottom: TPanel
    Left = 0
    Top = 698
    Width = 1127
    Height = 25
    Align = alBottom
    BevelOuter = bvNone
    Caption = 'pnlBottom'
    Color = 15109422
    Ctl3D = False
    ParentBackground = False
    ParentCtl3D = False
    ShowCaption = False
    TabOrder = 0
    object pnlInfo: TPanel
      Left = 213
      Top = 0
      Width = 235
      Height = 25
      Align = alRight
      BevelOuter = bvNone
      Caption = 'pnlInfo'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWhite
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
      ShowCaption = False
      TabOrder = 0
      object lblInfo: TLabel
        Left = 6
        Top = 4
        Width = 8
        Height = 15
        Font.Charset = GB2312_CHARSET
        Font.Color = clWhite
        Font.Height = -15
        Font.Name = #23435#20307
        Font.Style = []
        ParentFont = False
      end
      object bvlModule01: TBevel
        Left = 233
        Top = 0
        Width = 2
        Height = 25
        Align = alRight
        ExplicitLeft = 118
      end
      object bvlModule02: TBevel
        Left = 0
        Top = 0
        Width = 2
        Height = 25
        Align = alLeft
        ExplicitLeft = 118
      end
    end
    object pnlTime: TPanel
      Left = 903
      Top = 0
      Width = 224
      Height = 25
      Align = alRight
      BevelOuter = bvNone
      Caption = 'pnlDateTime'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWhite
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
      ShowCaption = False
      TabOrder = 1
      object lblTime: TLabel
        Left = 8
        Top = 4
        Width = 8
        Height = 15
        Font.Charset = GB2312_CHARSET
        Font.Color = clWhite
        Font.Height = -15
        Font.Name = #23435#20307
        Font.Style = []
        ParentFont = False
      end
    end
    object pnlIP: TPanel
      Left = 750
      Top = 0
      Width = 153
      Height = 25
      Align = alRight
      BevelOuter = bvNone
      Caption = 'pnlIP'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWhite
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
      ShowCaption = False
      TabOrder = 2
      object lblIP: TLabel
        Left = 4
        Top = 4
        Width = 8
        Height = 15
        Font.Charset = GB2312_CHARSET
        Font.Color = clWhite
        Font.Height = -15
        Font.Name = #23435#20307
        Font.Style = []
        ParentFont = False
      end
      object bvlIP: TBevel
        Left = 151
        Top = 0
        Width = 2
        Height = 25
        Align = alRight
        ExplicitLeft = 118
      end
    end
    object pnlWeb: TPanel
      Left = 448
      Top = 0
      Width = 302
      Height = 25
      Align = alRight
      BevelOuter = bvNone
      Caption = 'pnlWeb'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWhite
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
      ShowCaption = False
      TabOrder = 3
      object lblWeb: TLabel
        Left = 4
        Top = 4
        Width = 8
        Height = 15
        Font.Charset = GB2312_CHARSET
        Font.Color = clWhite
        Font.Height = -15
        Font.Name = #23435#20307
        Font.Style = []
        ParentFont = False
      end
      object bvlWeb: TBevel
        Left = 300
        Top = 0
        Width = 2
        Height = 25
        Align = alRight
        ExplicitLeft = 118
      end
    end
    object pnlLogin: TPanel
      Left = 0
      Top = 0
      Width = 153
      Height = 25
      Align = alLeft
      BevelOuter = bvNone
      Caption = 'pnlIP'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWhite
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
      ShowCaption = False
      TabOrder = 4
      object lblLogin: TLabel
        Left = 4
        Top = 4
        Width = 8
        Height = 15
        Font.Charset = GB2312_CHARSET
        Font.Color = clWhite
        Font.Height = -15
        Font.Name = #23435#20307
        Font.Style = []
        ParentFont = False
      end
    end
  end
  object clbrPModule: TCoolBar
    Left = 0
    Top = 0
    Width = 1127
    Height = 28
    AutoSize = True
    Bands = <
      item
        Control = tlbPModule
        ImageIndex = -1
        MinHeight = 24
        Width = 1121
      end>
    object tlbPModule: TToolBar
      Left = 11
      Top = 0
      Width = 1112
      Height = 24
      ButtonHeight = 24
      ButtonWidth = 43
      Caption = 'tlbPModule'
      Font.Charset = GB2312_CHARSET
      Font.Color = clWindowText
      Font.Height = -14
      Font.Name = #23435#20307
      Font.Style = []
      ParentFont = False
      ShowCaptions = True
      TabOrder = 0
    end
  end
  object rzpgcntrlAll: TPageControl
    Left = 0
    Top = 28
    Width = 1127
    Height = 670
    ActivePage = tsButton
    Align = alClient
    Style = tsButtons
    TabOrder = 2
    object tsButton: TTabSheet
      Caption = 'tsButton'
      object imgButtonBack: TImage
        Left = 0
        Top = 0
        Width = 1119
        Height = 639
        Align = alClient
        Stretch = True
        ExplicitLeft = 272
        ExplicitTop = 128
        ExplicitWidth = 105
        ExplicitHeight = 105
      end
      object pnlModuleDialog: TPanel
        Left = 176
        Top = 96
        Width = 655
        Height = 385
        BevelOuter = bvNone
        BorderStyle = bsSingle
        Caption = 'pnlModuleDialog'
        Color = clWhite
        Ctl3D = False
        ParentBackground = False
        ParentCtl3D = False
        ShowCaption = False
        TabOrder = 0
        object pnlModuleDialogTitle: TPanel
          Left = 0
          Top = 0
          Width = 653
          Height = 38
          Align = alTop
          Caption = 'pnlModuleDialogTitle'
          Color = 9916930
          Font.Charset = GB2312_CHARSET
          Font.Color = clWhite
          Font.Height = -19
          Font.Name = #24494#36719#38597#40657
          Font.Style = [fsBold]
          ParentBackground = False
          ParentFont = False
          TabOrder = 0
          DesignSize = (
            653
            38)
          object imgSubModuleClose: TImage
            Left = 616
            Top = 2
            Width = 31
            Height = 32
            Anchors = [akTop, akRight]
            AutoSize = True
            Transparent = True
            OnClick = imgSubModuleCloseClick
            OnMouseEnter = imgSubModuleCloseMouseEnter
            OnMouseLeave = imgSubModuleCloseMouseLeave
            ExplicitLeft = 584
          end
        end
      end
    end
    object tsList: TTabSheet
      Caption = 'tsList'
      ExplicitLeft = 0
      ExplicitTop = 0
      ExplicitWidth = 0
      ExplicitHeight = 0
      object imgListBack: TImage
        Left = 0
        Top = 0
        Width = 1119
        Height = 642
        Align = alClient
        Stretch = True
        ExplicitLeft = 272
        ExplicitTop = 128
        ExplicitWidth = 105
        ExplicitHeight = 105
      end
    end
    object tsDll: TTabSheet
      Caption = 'tsDll'
      ExplicitLeft = 0
      ExplicitTop = 0
      ExplicitWidth = 0
      ExplicitHeight = 0
      object imgDllFormBack: TImage
        Left = 0
        Top = 0
        Width = 1119
        Height = 642
        Align = alClient
        Stretch = True
        ExplicitLeft = 272
        ExplicitTop = 128
        ExplicitWidth = 105
        ExplicitHeight = 105
      end
    end
  end
  object ilMainMenu: TImageList
    ColorDepth = cd32Bit
    Height = 32
    Width = 32
    Left = 96
    Top = 154
  end
  object mmMainMenu: TMainMenu
    AutoHotkeys = maManual
    AutoMerge = True
    Images = ilMainMenu
    Left = 96
    Top = 262
  end
  object tmrDateTime: TTimer
    OnTimer = tmrDateTimeTimer
    Left = 96
    Top = 316
  end
  object pmTray: TPopupMenu
    AutoHotkeys = maManual
    Left = 96
    Top = 208
    object mniTrayShowForm: TMenuItem
      Caption = #26174#31034
      OnClick = mniTrayShowFormClick
    end
    object mniTrayLine01: TMenuItem
      Caption = '-'
    end
    object mniTrayExit: TMenuItem
      Caption = #36864#20986
      OnClick = mniTrayExitClick
    end
  end
  object ilPModule: TImageList
    ColorDepth = cd32Bit
    Height = 32
    Width = 32
    Left = 96
    Top = 100
  end
end
