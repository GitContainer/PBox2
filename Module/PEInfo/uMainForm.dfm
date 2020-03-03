object frmPEInfo: TfrmPEInfo
  Left = 320
  Top = 62
  Caption = 'PEInfo'
  ClientHeight = 760
  ClientWidth = 1454
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnClose = FormClose
  OnCreate = FormCreate
  DesignSize = (
    1454
    760)
  PixelsPerInch = 96
  TextHeight = 13
  object lbl1: TLabel
    Left = 8
    Top = 9
    Width = 72
    Height = 13
    Caption = #36873#25321'PE'#25991#20214#65306
  end
  object srchbxSelectPEFile: TSearchBox
    Left = 86
    Top = 5
    Width = 1360
    Height = 21
    Anchors = [akLeft, akTop, akRight]
    ReadOnly = True
    TabOrder = 0
    OnDblClick = srchbxSelectPEFileInvokeSearch
    OnInvokeSearch = srchbxSelectPEFileInvokeSearch
  end
  object pnlHex: TPanel
    Left = 2
    Top = 40
    Width = 631
    Height = 704
    Anchors = [akLeft, akTop, akBottom]
    BevelOuter = bvNone
    BorderStyle = bsSingle
    Caption = 'pnlHex'
    Color = clWhite
    Ctl3D = False
    ParentBackground = False
    ParentCtl3D = False
    ShowCaption = False
    TabOrder = 1
  end
  object pgcPEInfo: TPageControl
    Left = 639
    Top = 40
    Width = 798
    Height = 705
    ActivePage = tsNTHeader
    Anchors = [akLeft, akTop, akRight, akBottom]
    TabOrder = 2
    OnChange = pgcPEInfoChange
    object tsAll: TTabSheet
      Caption = 'PE '#31616#20171
      ExplicitLeft = 0
      ExplicitTop = 0
      ExplicitWidth = 0
      ExplicitHeight = 0
      object lbl57: TLabel
        Left = 20
        Top = 12
        Width = 15
        Height = 27
        Font.Charset = GB2312_CHARSET
        Font.Color = clRed
        Font.Height = -27
        Font.Name = #23435#20307
        Font.Style = [fsBold]
        ParentFont = False
      end
    end
    object tsDosHeader: TTabSheet
      Caption = 'DOS '#22836
      ImageIndex = 1
      ExplicitLeft = 0
      ExplicitTop = 0
      ExplicitWidth = 0
      ExplicitHeight = 0
      object lbl39: TLabel
        Left = 16
        Top = 23
        Width = 726
        Height = 19
        Caption = 
          #8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213 +
          #8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213
        Font.Charset = GB2312_CHARSET
        Font.Color = clSilver
        Font.Height = -13
        Font.Name = #24494#36719#38597#40657
        Font.Style = []
        ParentFont = False
      end
      object lbl38: TLabel
        Left = 184
        Top = 430
        Width = 72
        Height = 15
        Caption = 'e_lfanew'
        Font.Charset = GB2312_CHARSET
        Font.Color = clRed
        Font.Height = -15
        Font.Name = #23435#20307
        Font.Style = [fsBold]
        ParentFont = False
      end
      object lbl37: TLabel
        Left = 184
        Top = 406
        Width = 48
        Height = 15
        Caption = 'e_res2'
        Font.Charset = GB2312_CHARSET
        Font.Color = clBlack
        Font.Height = -15
        Font.Name = #23435#20307
        Font.Style = []
        ParentFont = False
      end
      object lbl40: TLabel
        Left = 16
        Top = 46
        Width = 726
        Height = 19
        Caption = 
          #8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213 +
          #8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213
        Font.Charset = GB2312_CHARSET
        Font.Color = clSilver
        Font.Height = -13
        Font.Name = #24494#36719#38597#40657
        Font.Style = []
        ParentFont = False
      end
      object lbl43: TLabel
        Left = 16
        Top = 116
        Width = 726
        Height = 19
        Caption = 
          #8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213 +
          #8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213
        Font.Charset = GB2312_CHARSET
        Font.Color = clSilver
        Font.Height = -13
        Font.Name = #24494#36719#38597#40657
        Font.Style = []
        ParentFont = False
      end
      object lbl42: TLabel
        Left = 16
        Top = 69
        Width = 726
        Height = 19
        Caption = 
          #8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213 +
          #8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213
        Font.Charset = GB2312_CHARSET
        Font.Color = clSilver
        Font.Height = -13
        Font.Name = #24494#36719#38597#40657
        Font.Style = []
        ParentFont = False
      end
      object lbl41: TLabel
        Left = 16
        Top = 93
        Width = 726
        Height = 19
        Caption = 
          #8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213 +
          #8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213
        Font.Charset = GB2312_CHARSET
        Font.Color = clSilver
        Font.Height = -13
        Font.Name = #24494#36719#38597#40657
        Font.Style = []
        ParentFont = False
      end
      object lbl32: TLabel
        Left = 184
        Top = 290
        Width = 64
        Height = 15
        Caption = 'e_lfarlc'
        Font.Charset = GB2312_CHARSET
        Font.Color = clBlack
        Font.Height = -15
        Font.Name = #23435#20307
        Font.Style = []
        ParentFont = False
      end
      object lbl31: TLabel
        Left = 184
        Top = 267
        Width = 32
        Height = 15
        Caption = 'e_cs'
        Font.Charset = GB2312_CHARSET
        Font.Color = clBlack
        Font.Height = -15
        Font.Name = #23435#20307
        Font.Style = []
        ParentFont = False
      end
      object lbl30: TLabel
        Left = 184
        Top = 244
        Width = 32
        Height = 15
        Caption = 'e_ip'
        Font.Charset = GB2312_CHARSET
        Font.Color = clBlack
        Font.Height = -15
        Font.Name = #23435#20307
        Font.Style = []
        ParentFont = False
      end
      object lbl33: TLabel
        Left = 184
        Top = 313
        Width = 48
        Height = 15
        Caption = 'e_ovno'
        Font.Charset = GB2312_CHARSET
        Font.Color = clBlack
        Font.Height = -15
        Font.Name = #23435#20307
        Font.Style = []
        ParentFont = False
      end
      object lbl36: TLabel
        Left = 184
        Top = 383
        Width = 72
        Height = 15
        Caption = 'e_oeminfo'
        Font.Charset = GB2312_CHARSET
        Font.Color = clBlack
        Font.Height = -15
        Font.Name = #23435#20307
        Font.Style = []
        ParentFont = False
      end
      object lbl35: TLabel
        Left = 184
        Top = 360
        Width = 56
        Height = 15
        Caption = 'e_oemid'
        Font.Charset = GB2312_CHARSET
        Font.Color = clBlack
        Font.Height = -15
        Font.Name = #23435#20307
        Font.Style = []
        ParentFont = False
      end
      object lbl34: TLabel
        Left = 184
        Top = 337
        Width = 40
        Height = 15
        Caption = 'e_res'
        Font.Charset = GB2312_CHARSET
        Font.Color = clBlack
        Font.Height = -15
        Font.Name = #23435#20307
        Font.Style = []
        ParentFont = False
      end
      object lbl53: TLabel
        Left = 16
        Top = 233
        Width = 726
        Height = 19
        Caption = 
          #8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213 +
          #8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213
        Font.Charset = GB2312_CHARSET
        Font.Color = clSilver
        Font.Height = -13
        Font.Name = #24494#36719#38597#40657
        Font.Style = []
        ParentFont = False
      end
      object lbl52: TLabel
        Left = 16
        Top = 256
        Width = 726
        Height = 19
        Caption = 
          #8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213 +
          #8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213
        Font.Charset = GB2312_CHARSET
        Font.Color = clSilver
        Font.Height = -13
        Font.Name = #24494#36719#38597#40657
        Font.Style = []
        ParentFont = False
      end
      object lbl51: TLabel
        Left = 16
        Top = 279
        Width = 726
        Height = 19
        Caption = 
          #8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213 +
          #8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213
        Font.Charset = GB2312_CHARSET
        Font.Color = clSilver
        Font.Height = -13
        Font.Name = #24494#36719#38597#40657
        Font.Style = []
        ParentFont = False
      end
      object lbl54: TLabel
        Left = 16
        Top = 209
        Width = 726
        Height = 19
        Caption = 
          #8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213 +
          #8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213
        Font.Charset = GB2312_CHARSET
        Font.Color = clSilver
        Font.Height = -13
        Font.Name = #24494#36719#38597#40657
        Font.Style = []
        ParentFont = False
      end
      object lbl56: TLabel
        Left = 16
        Top = 396
        Width = 726
        Height = 19
        Caption = 
          #8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213 +
          #8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213
        Font.Charset = GB2312_CHARSET
        Font.Color = clSilver
        Font.Height = -13
        Font.Name = #24494#36719#38597#40657
        Font.Style = []
        ParentFont = False
      end
      object lbl55: TLabel
        Left = 16
        Top = 420
        Width = 726
        Height = 19
        Caption = 
          #8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213 +
          #8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213
        Font.Charset = GB2312_CHARSET
        Font.Color = clSilver
        Font.Height = -13
        Font.Name = #24494#36719#38597#40657
        Font.Style = []
        ParentFont = False
      end
      object lbl46: TLabel
        Left = 16
        Top = 186
        Width = 726
        Height = 19
        Caption = 
          #8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213 +
          #8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213
        Font.Charset = GB2312_CHARSET
        Font.Color = clSilver
        Font.Height = -13
        Font.Name = #24494#36719#38597#40657
        Font.Style = []
        ParentFont = False
      end
      object lbl45: TLabel
        Left = 16
        Top = 163
        Width = 726
        Height = 19
        Caption = 
          #8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213 +
          #8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213
        Font.Charset = GB2312_CHARSET
        Font.Color = clSilver
        Font.Height = -13
        Font.Name = #24494#36719#38597#40657
        Font.Style = []
        ParentFont = False
      end
      object lbl44: TLabel
        Left = 16
        Top = 139
        Width = 726
        Height = 19
        Caption = 
          #8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213 +
          #8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213
        Font.Charset = GB2312_CHARSET
        Font.Color = clSilver
        Font.Height = -13
        Font.Name = #24494#36719#38597#40657
        Font.Style = []
        ParentFont = False
      end
      object lbl47: TLabel
        Left = 16
        Top = 373
        Width = 726
        Height = 19
        Caption = 
          #8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213 +
          #8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213
        Font.Charset = GB2312_CHARSET
        Font.Color = clSilver
        Font.Height = -13
        Font.Name = #24494#36719#38597#40657
        Font.Style = []
        ParentFont = False
      end
      object lbl50: TLabel
        Left = 16
        Top = 303
        Width = 726
        Height = 19
        Caption = 
          #8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213 +
          #8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213
        Font.Charset = GB2312_CHARSET
        Font.Color = clSilver
        Font.Height = -13
        Font.Name = #24494#36719#38597#40657
        Font.Style = []
        ParentFont = False
      end
      object lbl49: TLabel
        Left = 16
        Top = 326
        Width = 726
        Height = 19
        Caption = 
          #8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213 +
          #8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213
        Font.Charset = GB2312_CHARSET
        Font.Color = clSilver
        Font.Height = -13
        Font.Name = #24494#36719#38597#40657
        Font.Style = []
        ParentFont = False
      end
      object lbl48: TLabel
        Left = 16
        Top = 349
        Width = 726
        Height = 19
        Caption = 
          #8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213 +
          #8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213
        Font.Charset = GB2312_CHARSET
        Font.Color = clSilver
        Font.Height = -13
        Font.Name = #24494#36719#38597#40657
        Font.Style = []
        ParentFont = False
      end
      object lbl29: TLabel
        Left = 184
        Top = 221
        Width = 48
        Height = 15
        Caption = 'e_csum'
        Font.Charset = GB2312_CHARSET
        Font.Color = clBlack
        Font.Height = -15
        Font.Name = #23435#20307
        Font.Style = []
        ParentFont = False
      end
      object lbl10: TLabel
        Left = 20
        Top = 244
        Width = 32
        Height = 15
        Caption = 'e_ip'
        Font.Charset = GB2312_CHARSET
        Font.Color = clBlue
        Font.Height = -15
        Font.Name = #23435#20307
        Font.Style = []
        ParentFont = False
      end
      object lbl9: TLabel
        Left = 20
        Top = 267
        Width = 32
        Height = 15
        Caption = 'e_cs'
        Font.Charset = GB2312_CHARSET
        Font.Color = clBlue
        Font.Height = -15
        Font.Name = #23435#20307
        Font.Style = []
        ParentFont = False
      end
      object lbl8: TLabel
        Left = 20
        Top = 104
        Width = 72
        Height = 15
        Caption = 'e_cparhdr'
        Font.Charset = GB2312_CHARSET
        Font.Color = clBlue
        Font.Height = -15
        Font.Name = #23435#20307
        Font.Style = []
        ParentFont = False
      end
      object lbl11: TLabel
        Left = 20
        Top = 221
        Width = 48
        Height = 15
        Caption = 'e_csum'
        Font.Charset = GB2312_CHARSET
        Font.Color = clBlue
        Font.Height = -15
        Font.Name = #23435#20307
        Font.Style = []
        ParentFont = False
      end
      object lbl14: TLabel
        Left = 20
        Top = 337
        Width = 40
        Height = 15
        Caption = 'e_res'
        Font.Charset = GB2312_CHARSET
        Font.Color = clBlue
        Font.Height = -15
        Font.Name = #23435#20307
        Font.Style = []
        ParentFont = False
      end
      object lbl13: TLabel
        Left = 20
        Top = 360
        Width = 56
        Height = 15
        Caption = 'e_oemid'
        Font.Charset = GB2312_CHARSET
        Font.Color = clBlue
        Font.Height = -15
        Font.Name = #23435#20307
        Font.Style = []
        ParentFont = False
      end
      object lbl12: TLabel
        Left = 20
        Top = 197
        Width = 32
        Height = 15
        Caption = 'e_sp'
        Font.Charset = GB2312_CHARSET
        Font.Color = clBlue
        Font.Height = -15
        Font.Name = #23435#20307
        Font.Style = []
        ParentFont = False
      end
      object lbl3: TLabel
        Left = 20
        Top = 58
        Width = 32
        Height = 15
        Caption = 'e_cp'
        Font.Charset = GB2312_CHARSET
        Font.Color = clBlue
        Font.Height = -15
        Font.Name = #23435#20307
        Font.Style = []
        ParentFont = False
      end
      object lbl2: TLabel
        Left = 20
        Top = 35
        Width = 48
        Height = 15
        Caption = 'e_cblp'
        Font.Charset = GB2312_CHARSET
        Font.Color = clBlue
        Font.Height = -15
        Font.Name = #23435#20307
        Font.Style = []
        ParentFont = False
      end
      object Label1: TLabel
        Left = 20
        Top = 12
        Width = 56
        Height = 15
        Caption = 'e_magic'
        Font.Charset = GB2312_CHARSET
        Font.Color = clBlue
        Font.Height = -15
        Font.Name = #23435#20307
        Font.Style = []
        ParentFont = False
      end
      object lbl4: TLabel
        Left = 20
        Top = 81
        Width = 48
        Height = 15
        Caption = 'e_crlc'
        Font.Charset = GB2312_CHARSET
        Font.Color = clBlue
        Font.Height = -15
        Font.Name = #23435#20307
        Font.Style = []
        ParentFont = False
      end
      object lbl7: TLabel
        Left = 20
        Top = 128
        Width = 80
        Height = 15
        Caption = 'e_minalloc'
        Font.Charset = GB2312_CHARSET
        Font.Color = clBlue
        Font.Height = -15
        Font.Name = #23435#20307
        Font.Style = []
        ParentFont = False
      end
      object lbl6: TLabel
        Left = 20
        Top = 151
        Width = 80
        Height = 15
        Caption = 'e_maxalloc'
        Font.Charset = GB2312_CHARSET
        Font.Color = clBlue
        Font.Height = -15
        Font.Name = #23435#20307
        Font.Style = []
        ParentFont = False
      end
      object lbl5: TLabel
        Left = 20
        Top = 174
        Width = 32
        Height = 15
        Caption = 'e_ss'
        Font.Charset = GB2312_CHARSET
        Font.Color = clBlue
        Font.Height = -15
        Font.Name = #23435#20307
        Font.Style = []
        ParentFont = False
      end
      object lbl24: TLabel
        Left = 184
        Top = 104
        Width = 72
        Height = 15
        Caption = 'e_cparhdr'
        Font.Charset = GB2312_CHARSET
        Font.Color = clBlack
        Font.Height = -15
        Font.Name = #23435#20307
        Font.Style = []
        ParentFont = False
      end
      object lbl23: TLabel
        Left = 184
        Top = 81
        Width = 48
        Height = 15
        Caption = 'e_crlc'
        Font.Charset = GB2312_CHARSET
        Font.Color = clBlack
        Font.Height = -15
        Font.Name = #23435#20307
        Font.Style = []
        ParentFont = False
      end
      object lbl22: TLabel
        Left = 184
        Top = 58
        Width = 32
        Height = 15
        Caption = 'e_cp'
        Font.Charset = GB2312_CHARSET
        Font.Color = clBlack
        Font.Height = -15
        Font.Name = #23435#20307
        Font.Style = []
        ParentFont = False
      end
      object lbl25: TLabel
        Left = 184
        Top = 128
        Width = 80
        Height = 15
        Caption = 'e_minalloc'
        Font.Charset = GB2312_CHARSET
        Font.Color = clBlack
        Font.Height = -15
        Font.Name = #23435#20307
        Font.Style = []
        ParentFont = False
      end
      object lbl28: TLabel
        Left = 184
        Top = 197
        Width = 32
        Height = 15
        Caption = 'e_sp'
        Font.Charset = GB2312_CHARSET
        Font.Color = clBlack
        Font.Height = -15
        Font.Name = #23435#20307
        Font.Style = []
        ParentFont = False
      end
      object lbl27: TLabel
        Left = 184
        Top = 174
        Width = 32
        Height = 15
        Caption = 'e_ss'
        Font.Charset = GB2312_CHARSET
        Font.Color = clBlack
        Font.Height = -15
        Font.Name = #23435#20307
        Font.Style = []
        ParentFont = False
      end
      object lbl26: TLabel
        Left = 184
        Top = 151
        Width = 80
        Height = 15
        Caption = 'e_maxalloc'
        Font.Charset = GB2312_CHARSET
        Font.Color = clBlack
        Font.Height = -15
        Font.Name = #23435#20307
        Font.Style = []
        ParentFont = False
      end
      object lbl18: TLabel
        Left = 20
        Top = 430
        Width = 64
        Height = 15
        Caption = 'e_lfanew'
        Font.Charset = GB2312_CHARSET
        Font.Color = clBlue
        Font.Height = -15
        Font.Name = #23435#20307
        Font.Style = []
        ParentFont = False
      end
      object lbl16: TLabel
        Left = 20
        Top = 290
        Width = 64
        Height = 15
        Caption = 'e_lfarlc'
        Font.Charset = GB2312_CHARSET
        Font.Color = clBlue
        Font.Height = -15
        Font.Name = #23435#20307
        Font.Style = []
        ParentFont = False
      end
      object lbl15: TLabel
        Left = 20
        Top = 313
        Width = 48
        Height = 15
        Caption = 'e_ovno'
        Font.Charset = GB2312_CHARSET
        Font.Color = clBlue
        Font.Height = -15
        Font.Name = #23435#20307
        Font.Style = []
        ParentFont = False
      end
      object lbl19: TLabel
        Left = 20
        Top = 406
        Width = 48
        Height = 15
        Caption = 'e_res2'
        Font.Charset = GB2312_CHARSET
        Font.Color = clBlue
        Font.Height = -15
        Font.Name = #23435#20307
        Font.Style = []
        ParentFont = False
      end
      object lbl21: TLabel
        Left = 184
        Top = 35
        Width = 48
        Height = 15
        Caption = 'e_cblp'
        Font.Charset = GB2312_CHARSET
        Font.Color = clBlack
        Font.Height = -15
        Font.Name = #23435#20307
        Font.Style = []
        ParentFont = False
      end
      object lbl17: TLabel
        Left = 184
        Top = 12
        Width = 63
        Height = 15
        Caption = 'e_magic'
        Font.Charset = GB2312_CHARSET
        Font.Color = clRed
        Font.Height = -15
        Font.Name = #23435#20307
        Font.Style = [fsBold]
        ParentFont = False
      end
      object lbl20: TLabel
        Left = 20
        Top = 383
        Width = 72
        Height = 15
        Caption = 'e_oeminfo'
        Font.Charset = GB2312_CHARSET
        Font.Color = clBlue
        Font.Height = -15
        Font.Name = #23435#20307
        Font.Style = []
        ParentFont = False
      end
      object btnMagic: TButton
        Left = 260
        Top = 6
        Width = 75
        Height = 25
        Caption = #21487#36873#20540
        TabOrder = 0
        OnClick = btnMagicClick
      end
    end
    object tsDosStub: TTabSheet
      Caption = 'DOS '#25554#26729
      ImageIndex = 8
      ExplicitLeft = 0
      ExplicitTop = 0
      ExplicitWidth = 0
      ExplicitHeight = 0
    end
    object tsNTHeader: TTabSheet
      Caption = 'NT '#22836
      ImageIndex = 2
      ExplicitLeft = 0
      ExplicitTop = 0
      ExplicitWidth = 0
      ExplicitHeight = 0
      DesignSize = (
        790
        677)
      object Label2: TLabel
        Left = 24
        Top = 16
        Width = 81
        Height = 15
        Caption = 'Signature'
        Font.Charset = GB2312_CHARSET
        Font.Color = 33023
        Font.Height = -15
        Font.Name = #23435#20307
        Font.Style = [fsBold]
        ParentFont = False
      end
      object lbl58: TLabel
        Left = 24
        Top = 41
        Width = 90
        Height = 15
        Caption = 'FileHeader'
        Font.Charset = GB2312_CHARSET
        Font.Color = 33023
        Font.Height = -15
        Font.Name = #23435#20307
        Font.Style = [fsBold]
        ParentFont = False
      end
      object lbl59: TLabel
        Left = 20
        Top = 234
        Width = 126
        Height = 15
        Caption = 'OptionalHeader'
        Font.Charset = GB2312_CHARSET
        Font.Color = 33023
        Font.Height = -15
        Font.Name = #23435#20307
        Font.Style = [fsBold]
        ParentFont = False
      end
      object lbl60: TLabel
        Left = 52
        Top = 63
        Width = 56
        Height = 15
        Caption = 'Machine'
        Font.Charset = GB2312_CHARSET
        Font.Color = clBlue
        Font.Height = -15
        Font.Name = #23435#20307
        Font.Style = []
        ParentFont = False
      end
      object lbl61: TLabel
        Left = 52
        Top = 86
        Width = 128
        Height = 15
        Caption = 'NumberOfSections'
        Font.Charset = GB2312_CHARSET
        Font.Color = clBlue
        Font.Height = -15
        Font.Name = #23435#20307
        Font.Style = []
        ParentFont = False
      end
      object lbl62: TLabel
        Left = 52
        Top = 109
        Width = 104
        Height = 15
        Caption = 'TimeDateStamp'
        Font.Charset = GB2312_CHARSET
        Font.Color = clBlue
        Font.Height = -15
        Font.Name = #23435#20307
        Font.Style = []
        ParentFont = False
      end
      object lbl63: TLabel
        Left = 52
        Top = 132
        Width = 160
        Height = 15
        Caption = 'PointerToSymbolTable'
        Font.Charset = GB2312_CHARSET
        Font.Color = clBlue
        Font.Height = -15
        Font.Name = #23435#20307
        Font.Style = []
        ParentFont = False
      end
      object lbl64: TLabel
        Left = 52
        Top = 156
        Width = 120
        Height = 15
        Caption = 'NumberOfSymbols'
        Font.Charset = GB2312_CHARSET
        Font.Color = clBlue
        Font.Height = -15
        Font.Name = #23435#20307
        Font.Style = []
        ParentFont = False
      end
      object lbl65: TLabel
        Left = 52
        Top = 179
        Width = 160
        Height = 15
        Caption = 'SizeOfOptionalHeader'
        Font.Charset = GB2312_CHARSET
        Font.Color = clBlue
        Font.Height = -15
        Font.Name = #23435#20307
        Font.Style = []
        ParentFont = False
      end
      object lbl66: TLabel
        Left = 52
        Top = 202
        Width = 120
        Height = 15
        Caption = 'Characteristics'
        Font.Charset = GB2312_CHARSET
        Font.Color = clBlue
        Font.Height = -15
        Font.Name = #23435#20307
        Font.Style = []
        ParentFont = False
      end
      object lbl67: TLabel
        Left = 304
        Top = 16
        Width = 81
        Height = 15
        Caption = 'Signature'
        Font.Charset = GB2312_CHARSET
        Font.Color = clRed
        Font.Height = -15
        Font.Name = #23435#20307
        Font.Style = [fsBold]
        ParentFont = False
      end
      object lbl68: TLabel
        Left = 304
        Top = 63
        Width = 63
        Height = 15
        Caption = 'Machine'
        Font.Charset = GB2312_CHARSET
        Font.Color = clRed
        Font.Height = -15
        Font.Name = #23435#20307
        Font.Style = [fsBold]
        ParentFont = False
      end
      object lbl69: TLabel
        Left = 304
        Top = 86
        Width = 144
        Height = 15
        Caption = 'NumberOfSections'
        Font.Charset = GB2312_CHARSET
        Font.Color = clRed
        Font.Height = -15
        Font.Name = #23435#20307
        Font.Style = [fsBold]
        ParentFont = False
      end
      object lbl70: TLabel
        Left = 304
        Top = 109
        Width = 104
        Height = 15
        Caption = 'TimeDateStamp'
        Font.Charset = GB2312_CHARSET
        Font.Color = clBlack
        Font.Height = -15
        Font.Name = #23435#20307
        Font.Style = []
        ParentFont = False
      end
      object lbl71: TLabel
        Left = 304
        Top = 132
        Width = 160
        Height = 15
        Caption = 'PointerToSymbolTable'
        Font.Charset = GB2312_CHARSET
        Font.Color = clBlack
        Font.Height = -15
        Font.Name = #23435#20307
        Font.Style = []
        ParentFont = False
      end
      object lbl72: TLabel
        Left = 304
        Top = 156
        Width = 120
        Height = 15
        Caption = 'NumberOfSymbols'
        Font.Charset = GB2312_CHARSET
        Font.Color = clBlack
        Font.Height = -15
        Font.Name = #23435#20307
        Font.Style = []
        ParentFont = False
      end
      object lbl73: TLabel
        Left = 304
        Top = 179
        Width = 160
        Height = 15
        Caption = 'SizeOfOptionalHeader'
        Font.Charset = GB2312_CHARSET
        Font.Color = clBlack
        Font.Height = -15
        Font.Name = #23435#20307
        Font.Style = []
        ParentFont = False
      end
      object lbl74: TLabel
        Left = 304
        Top = 202
        Width = 135
        Height = 15
        Caption = 'Characteristics'
        Font.Charset = GB2312_CHARSET
        Font.Color = clRed
        Font.Height = -15
        Font.Name = #23435#20307
        Font.Style = [fsBold]
        ParentFont = False
      end
      object lbl75: TLabel
        Left = 20
        Top = 28
        Width = 738
        Height = 19
        Caption = 
          #8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213 +
          #8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213
        Font.Charset = GB2312_CHARSET
        Font.Color = clSilver
        Font.Height = -13
        Font.Name = #24494#36719#38597#40657
        Font.Style = []
        ParentFont = False
      end
      object lbl76: TLabel
        Left = 20
        Top = 216
        Width = 738
        Height = 19
        Caption = 
          #8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213 +
          #8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213
        Font.Charset = GB2312_CHARSET
        Font.Color = clSilver
        Font.Height = -13
        Font.Name = #24494#36719#38597#40657
        Font.Style = []
        ParentFont = False
      end
      object scrlbx1: TScrollBox
        Left = 24
        Top = 248
        Width = 731
        Height = 409
        VertScrollBar.Smooth = True
        VertScrollBar.Tracking = True
        Anchors = [akLeft, akTop, akRight, akBottom]
        BorderStyle = bsNone
        Color = clWhite
        Ctl3D = False
        ParentColor = False
        ParentCtl3D = False
        TabOrder = 0
        OnMouseWheel = scrlbx1MouseWheel
        object lbl77: TLabel
          Left = 31
          Top = 11
          Width = 40
          Height = 15
          Caption = 'Magic'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlue
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl78: TLabel
          Left = 31
          Top = 34
          Width = 144
          Height = 15
          Caption = 'MajorLinkerVersion'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlue
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl79: TLabel
          Left = 31
          Top = 57
          Width = 144
          Height = 15
          Caption = 'MinorLinkerVersion'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlue
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl80: TLabel
          Left = 31
          Top = 80
          Width = 80
          Height = 15
          Caption = 'SizeOfCode'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlue
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl81: TLabel
          Left = 31
          Top = 103
          Width = 168
          Height = 15
          Caption = 'SizeOfInitializedData'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlue
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl82: TLabel
          Left = 31
          Top = 126
          Width = 184
          Height = 15
          Caption = 'SizeOfUninitializedData'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlue
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl83: TLabel
          Left = 31
          Top = 150
          Width = 152
          Height = 15
          Caption = 'AddressOfEntryPoint'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlue
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl84: TLabel
          Left = 31
          Top = 173
          Width = 80
          Height = 15
          Caption = 'BaseOfCode'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlue
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl85: TLabel
          Left = 31
          Top = 196
          Width = 80
          Height = 15
          Caption = 'BaseOfData'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlue
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl86: TLabel
          Left = 31
          Top = 219
          Width = 72
          Height = 15
          Caption = 'ImageBase'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlue
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl87: TLabel
          Left = 31
          Top = 242
          Width = 128
          Height = 15
          Caption = 'SectionAlignment'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlue
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl88: TLabel
          Left = 31
          Top = 265
          Width = 104
          Height = 15
          Caption = 'FileAlignment'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlue
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl89: TLabel
          Left = 31
          Top = 289
          Width = 216
          Height = 15
          Caption = 'MajorOperatingSystemVersion'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlue
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl90: TLabel
          Left = 31
          Top = 312
          Width = 216
          Height = 15
          Caption = 'MinorOperatingSystemVersion'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlue
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl91: TLabel
          Left = 31
          Top = 335
          Width = 136
          Height = 15
          Caption = 'MajorImageVersion'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlue
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl92: TLabel
          Left = 31
          Top = 358
          Width = 136
          Height = 15
          Caption = 'MinorImageVersion'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlue
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl93: TLabel
          Left = 31
          Top = 381
          Width = 168
          Height = 15
          Caption = 'MajorSubsystemVersion'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlue
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl94: TLabel
          Left = 31
          Top = 404
          Width = 168
          Height = 15
          Caption = 'MinorSubsystemVersion'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlue
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl95: TLabel
          Left = 31
          Top = 428
          Width = 136
          Height = 15
          Caption = 'Win32VersionValue'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlue
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl96: TLabel
          Left = 31
          Top = 451
          Width = 88
          Height = 15
          Caption = 'SizeOfImage'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlue
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl97: TLabel
          Left = 31
          Top = 474
          Width = 104
          Height = 15
          Caption = 'SizeOfHeaders'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlue
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl98: TLabel
          Left = 31
          Top = 497
          Width = 64
          Height = 15
          Caption = 'CheckSum'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlue
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl99: TLabel
          Left = 31
          Top = 520
          Width = 72
          Height = 15
          Caption = 'Subsystem'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlue
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl100: TLabel
          Left = 31
          Top = 543
          Width = 144
          Height = 15
          Caption = 'DllCharacteristics'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlue
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl101: TLabel
          Left = 31
          Top = 567
          Width = 144
          Height = 15
          Caption = 'SizeOfStackReserve'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlue
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl102: TLabel
          Left = 31
          Top = 590
          Width = 136
          Height = 15
          Caption = 'SizeOfStackCommit'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlue
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl103: TLabel
          Left = 31
          Top = 613
          Width = 136
          Height = 15
          Caption = 'SizeOfHeapReserve'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlue
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl104: TLabel
          Left = 31
          Top = 636
          Width = 128
          Height = 15
          Caption = 'SizeOfHeapCommit'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlue
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl105: TLabel
          Left = 31
          Top = 659
          Width = 88
          Height = 15
          Caption = 'LoaderFlags'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlue
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl106: TLabel
          Left = 31
          Top = 682
          Width = 152
          Height = 15
          Caption = 'NumberOfRvaAndSizes'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlue
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl107: TLabel
          Left = 31
          Top = 706
          Width = 104
          Height = 15
          Caption = 'DataDirectory'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlue
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl108: TLabel
          Left = 92
          Top = 728
          Width = 32
          Height = 15
          Caption = '[00]'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlue
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl109: TLabel
          Left = 92
          Top = 801
          Width = 32
          Height = 15
          Caption = '[03]'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlue
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl110: TLabel
          Left = 92
          Top = 777
          Width = 32
          Height = 15
          Caption = '[02]'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlue
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl111: TLabel
          Left = 92
          Top = 752
          Width = 32
          Height = 15
          Caption = '[01]'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlue
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl112: TLabel
          Left = 92
          Top = 826
          Width = 32
          Height = 15
          Caption = '[04]'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlue
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl113: TLabel
          Left = 92
          Top = 851
          Width = 32
          Height = 15
          Caption = '[05]'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlue
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl114: TLabel
          Left = 92
          Top = 875
          Width = 32
          Height = 15
          Caption = '[06]'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlue
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl115: TLabel
          Left = 92
          Top = 900
          Width = 32
          Height = 15
          Caption = '[07]'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlue
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl116: TLabel
          Left = 92
          Top = 924
          Width = 32
          Height = 15
          Caption = '[08]'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlue
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl117: TLabel
          Left = 92
          Top = 949
          Width = 32
          Height = 15
          Caption = '[08]'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlue
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl118: TLabel
          Left = 92
          Top = 974
          Width = 32
          Height = 15
          Caption = '[10]'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlue
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl119: TLabel
          Left = 92
          Top = 998
          Width = 32
          Height = 15
          Caption = '[11]'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlue
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl120: TLabel
          Left = 92
          Top = 1023
          Width = 32
          Height = 15
          Caption = '[12]'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlue
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl121: TLabel
          Left = 92
          Top = 1047
          Width = 32
          Height = 15
          Caption = '[13]'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlue
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl122: TLabel
          Left = 92
          Top = 1072
          Width = 32
          Height = 15
          Caption = '[14]'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlue
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl123: TLabel
          Left = 92
          Top = 1097
          Width = 32
          Height = 15
          Caption = '[15]'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlue
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl124: TLabel
          Left = 280
          Top = 11
          Width = 45
          Height = 15
          Caption = 'Magic'
          Font.Charset = GB2312_CHARSET
          Font.Color = clRed
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = [fsBold]
          ParentFont = False
        end
        object lbl125: TLabel
          Left = 280
          Top = 34
          Width = 144
          Height = 15
          Caption = 'MajorLinkerVersion'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl126: TLabel
          Left = 280
          Top = 57
          Width = 144
          Height = 15
          Caption = 'MinorLinkerVersion'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl127: TLabel
          Left = 280
          Top = 80
          Width = 80
          Height = 15
          Caption = 'SizeOfCode'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl128: TLabel
          Left = 280
          Top = 103
          Width = 168
          Height = 15
          Caption = 'SizeOfInitializedData'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl129: TLabel
          Left = 280
          Top = 126
          Width = 184
          Height = 15
          Caption = 'SizeOfUninitializedData'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl130: TLabel
          Left = 280
          Top = 150
          Width = 171
          Height = 15
          Caption = 'AddressOfEntryPoint'
          Font.Charset = GB2312_CHARSET
          Font.Color = clRed
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = [fsBold]
          ParentFont = False
        end
        object lbl131: TLabel
          Left = 280
          Top = 173
          Width = 80
          Height = 15
          Caption = 'BaseOfCode'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl132: TLabel
          Left = 280
          Top = 196
          Width = 80
          Height = 15
          Caption = 'BaseOfData'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl133: TLabel
          Left = 280
          Top = 219
          Width = 81
          Height = 15
          Caption = 'ImageBase'
          Font.Charset = GB2312_CHARSET
          Font.Color = clRed
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = [fsBold]
          ParentFont = False
        end
        object lbl134: TLabel
          Left = 280
          Top = 242
          Width = 144
          Height = 15
          Caption = 'SectionAlignment'
          Font.Charset = GB2312_CHARSET
          Font.Color = clRed
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = [fsBold]
          ParentFont = False
        end
        object lbl135: TLabel
          Left = 280
          Top = 265
          Width = 117
          Height = 15
          Caption = 'FileAlignment'
          Font.Charset = GB2312_CHARSET
          Font.Color = clRed
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = [fsBold]
          ParentFont = False
        end
        object lbl136: TLabel
          Left = 280
          Top = 289
          Width = 216
          Height = 15
          Caption = 'MajorOperatingSystemVersion'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl137: TLabel
          Left = 280
          Top = 312
          Width = 216
          Height = 15
          Caption = 'MinorOperatingSystemVersion'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl138: TLabel
          Left = 280
          Top = 335
          Width = 136
          Height = 15
          Caption = 'MajorImageVersion'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl139: TLabel
          Left = 280
          Top = 358
          Width = 136
          Height = 15
          Caption = 'MinorImageVersion'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl140: TLabel
          Left = 280
          Top = 381
          Width = 168
          Height = 15
          Caption = 'MajorSubsystemVersion'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl141: TLabel
          Left = 280
          Top = 404
          Width = 168
          Height = 15
          Caption = 'MinorSubsystemVersion'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl142: TLabel
          Left = 280
          Top = 428
          Width = 136
          Height = 15
          Caption = 'Win32VersionValue'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl143: TLabel
          Left = 280
          Top = 451
          Width = 88
          Height = 15
          Caption = 'SizeOfImage'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl144: TLabel
          Left = 280
          Top = 474
          Width = 104
          Height = 15
          Caption = 'SizeOfHeaders'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl145: TLabel
          Left = 280
          Top = 497
          Width = 64
          Height = 15
          Caption = 'CheckSum'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl146: TLabel
          Left = 280
          Top = 520
          Width = 81
          Height = 15
          Caption = 'Subsystem'
          Font.Charset = GB2312_CHARSET
          Font.Color = clRed
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = [fsBold]
          ParentFont = False
        end
        object lbl147: TLabel
          Left = 280
          Top = 543
          Width = 162
          Height = 15
          Caption = 'DllCharacteristics'
          Font.Charset = GB2312_CHARSET
          Font.Color = clRed
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = [fsBold]
          ParentFont = False
        end
        object lbl148: TLabel
          Left = 280
          Top = 567
          Width = 144
          Height = 15
          Caption = 'SizeOfStackReserve'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl149: TLabel
          Left = 280
          Top = 590
          Width = 136
          Height = 15
          Caption = 'SizeOfStackCommit'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl150: TLabel
          Left = 280
          Top = 613
          Width = 136
          Height = 15
          Caption = 'SizeOfHeapReserve'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl151: TLabel
          Left = 280
          Top = 636
          Width = 128
          Height = 15
          Caption = 'SizeOfHeapCommit'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl152: TLabel
          Left = 280
          Top = 657
          Width = 88
          Height = 15
          Caption = 'LoaderFlags'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl153: TLabel
          Left = 280
          Top = 682
          Width = 152
          Height = 15
          Caption = 'NumberOfRvaAndSizes'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl154: TLabel
          Left = 280
          Top = 728
          Width = 99
          Height = 15
          Cursor = crHandPoint
          Caption = 'LoaderFlags'
          Font.Charset = GB2312_CHARSET
          Font.Color = 8421440
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = [fsBold]
          ParentFont = False
          OnClick = lbl154Click
        end
        object lbl155: TLabel
          Left = 396
          Top = 728
          Width = 171
          Height = 15
          Cursor = crHandPoint
          Caption = 'NumberOfRvaAndSizes'
          Font.Charset = GB2312_CHARSET
          Font.Color = 8421440
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = [fsBold]
          ParentFont = False
          OnClick = lbl154Click
        end
        object lbl156: TLabel
          Left = 280
          Top = 752
          Width = 99
          Height = 15
          Cursor = crHandPoint
          Caption = 'LoaderFlags'
          Font.Charset = GB2312_CHARSET
          Font.Color = 8421440
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = [fsBold]
          ParentFont = False
          OnClick = lbl156Click
        end
        object lbl157: TLabel
          Left = 396
          Top = 752
          Width = 171
          Height = 15
          Cursor = crHandPoint
          Caption = 'NumberOfRvaAndSizes'
          Font.Charset = GB2312_CHARSET
          Font.Color = 8421440
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = [fsBold]
          ParentFont = False
          OnClick = lbl156Click
        end
        object lbl158: TLabel
          Left = 280
          Top = 777
          Width = 99
          Height = 15
          Cursor = crHandPoint
          Caption = 'LoaderFlags'
          Font.Charset = GB2312_CHARSET
          Font.Color = 8421440
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = [fsBold]
          ParentFont = False
          OnClick = lbl158Click
        end
        object lbl160: TLabel
          Left = 280
          Top = 801
          Width = 88
          Height = 15
          Caption = 'LoaderFlags'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl159: TLabel
          Left = 396
          Top = 777
          Width = 171
          Height = 15
          Cursor = crHandPoint
          Caption = 'NumberOfRvaAndSizes'
          Font.Charset = GB2312_CHARSET
          Font.Color = 8421440
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = [fsBold]
          ParentFont = False
          OnClick = lbl158Click
        end
        object lbl161: TLabel
          Left = 396
          Top = 801
          Width = 152
          Height = 15
          Caption = 'NumberOfRvaAndSizes'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl162: TLabel
          Left = 280
          Top = 826
          Width = 88
          Height = 15
          Caption = 'LoaderFlags'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl164: TLabel
          Left = 280
          Top = 850
          Width = 88
          Height = 15
          Caption = 'LoaderFlags'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl166: TLabel
          Left = 280
          Top = 875
          Width = 88
          Height = 15
          Caption = 'LoaderFlags'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl168: TLabel
          Left = 280
          Top = 899
          Width = 88
          Height = 15
          Caption = 'LoaderFlags'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl163: TLabel
          Left = 396
          Top = 826
          Width = 152
          Height = 15
          Caption = 'NumberOfRvaAndSizes'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl165: TLabel
          Left = 396
          Top = 850
          Width = 152
          Height = 15
          Caption = 'NumberOfRvaAndSizes'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl167: TLabel
          Left = 396
          Top = 875
          Width = 152
          Height = 15
          Caption = 'NumberOfRvaAndSizes'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl169: TLabel
          Left = 396
          Top = 899
          Width = 152
          Height = 15
          Caption = 'NumberOfRvaAndSizes'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl170: TLabel
          Left = 280
          Top = 924
          Width = 88
          Height = 15
          Caption = 'LoaderFlags'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl172: TLabel
          Left = 280
          Top = 948
          Width = 88
          Height = 15
          Caption = 'LoaderFlags'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl174: TLabel
          Left = 280
          Top = 973
          Width = 88
          Height = 15
          Caption = 'LoaderFlags'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl176: TLabel
          Left = 280
          Top = 997
          Width = 88
          Height = 15
          Caption = 'LoaderFlags'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl178: TLabel
          Left = 280
          Top = 1022
          Width = 88
          Height = 15
          Caption = 'LoaderFlags'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl180: TLabel
          Left = 280
          Top = 1047
          Width = 88
          Height = 15
          Caption = 'LoaderFlags'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl182: TLabel
          Left = 280
          Top = 1071
          Width = 88
          Height = 15
          Caption = 'LoaderFlags'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl184: TLabel
          Left = 280
          Top = 1095
          Width = 88
          Height = 15
          Caption = 'LoaderFlags'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl171: TLabel
          Left = 396
          Top = 924
          Width = 152
          Height = 15
          Caption = 'NumberOfRvaAndSizes'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl173: TLabel
          Left = 396
          Top = 948
          Width = 152
          Height = 15
          Caption = 'NumberOfRvaAndSizes'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl175: TLabel
          Left = 396
          Top = 973
          Width = 152
          Height = 15
          Caption = 'NumberOfRvaAndSizes'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl177: TLabel
          Left = 396
          Top = 997
          Width = 152
          Height = 15
          Caption = 'NumberOfRvaAndSizes'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl179: TLabel
          Left = 396
          Top = 1022
          Width = 152
          Height = 15
          Caption = 'NumberOfRvaAndSizes'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl181: TLabel
          Left = 396
          Top = 1046
          Width = 152
          Height = 15
          Caption = 'NumberOfRvaAndSizes'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl183: TLabel
          Left = 396
          Top = 1071
          Width = 152
          Height = 15
          Caption = 'NumberOfRvaAndSizes'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl185: TLabel
          Left = 396
          Top = 1095
          Width = 152
          Height = 15
          Caption = 'NumberOfRvaAndSizes'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl186: TLabel
          Left = 56
          Top = 739
          Width = 636
          Height = 19
          Caption = 
            #8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213 +
            #8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213
          Font.Charset = GB2312_CHARSET
          Font.Color = clSilver
          Font.Height = -13
          Font.Name = #24494#36719#38597#40657
          Font.Style = []
          ParentFont = False
        end
        object lbl187: TLabel
          Left = 56
          Top = 764
          Width = 636
          Height = 19
          Caption = 
            #8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213 +
            #8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213
          Font.Charset = GB2312_CHARSET
          Font.Color = clSilver
          Font.Height = -13
          Font.Name = #24494#36719#38597#40657
          Font.Style = []
          ParentFont = False
        end
        object lbl188: TLabel
          Left = 56
          Top = 813
          Width = 636
          Height = 19
          Caption = 
            #8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213 +
            #8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213
          Font.Charset = GB2312_CHARSET
          Font.Color = clSilver
          Font.Height = -13
          Font.Name = #24494#36719#38597#40657
          Font.Style = []
          ParentFont = False
        end
        object lbl189: TLabel
          Left = 56
          Top = 788
          Width = 636
          Height = 19
          Caption = 
            #8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213 +
            #8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213
          Font.Charset = GB2312_CHARSET
          Font.Color = clSilver
          Font.Height = -13
          Font.Name = #24494#36719#38597#40657
          Font.Style = []
          ParentFont = False
        end
        object lbl190: TLabel
          Left = 56
          Top = 911
          Width = 636
          Height = 19
          Caption = 
            #8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213 +
            #8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213
          Font.Charset = GB2312_CHARSET
          Font.Color = clSilver
          Font.Height = -13
          Font.Name = #24494#36719#38597#40657
          Font.Style = []
          ParentFont = False
        end
        object lbl191: TLabel
          Left = 56
          Top = 886
          Width = 636
          Height = 19
          Caption = 
            #8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213 +
            #8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213
          Font.Charset = GB2312_CHARSET
          Font.Color = clSilver
          Font.Height = -13
          Font.Name = #24494#36719#38597#40657
          Font.Style = []
          ParentFont = False
        end
        object lbl192: TLabel
          Left = 56
          Top = 862
          Width = 636
          Height = 19
          Caption = 
            #8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213 +
            #8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213
          Font.Charset = GB2312_CHARSET
          Font.Color = clSilver
          Font.Height = -13
          Font.Name = #24494#36719#38597#40657
          Font.Style = []
          ParentFont = False
        end
        object lbl193: TLabel
          Left = 56
          Top = 837
          Width = 636
          Height = 19
          Caption = 
            #8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213 +
            #8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213
          Font.Charset = GB2312_CHARSET
          Font.Color = clSilver
          Font.Height = -13
          Font.Name = #24494#36719#38597#40657
          Font.Style = []
          ParentFont = False
        end
        object lbl195: TLabel
          Left = 56
          Top = 1082
          Width = 636
          Height = 19
          Caption = 
            #8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213 +
            #8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213
          Font.Charset = GB2312_CHARSET
          Font.Color = clSilver
          Font.Height = -13
          Font.Name = #24494#36719#38597#40657
          Font.Style = []
          ParentFont = False
        end
        object lbl196: TLabel
          Left = 56
          Top = 1058
          Width = 636
          Height = 19
          Caption = 
            #8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213 +
            #8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213
          Font.Charset = GB2312_CHARSET
          Font.Color = clSilver
          Font.Height = -13
          Font.Name = #24494#36719#38597#40657
          Font.Style = []
          ParentFont = False
        end
        object lbl197: TLabel
          Left = 56
          Top = 1033
          Width = 636
          Height = 19
          Caption = 
            #8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213 +
            #8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213
          Font.Charset = GB2312_CHARSET
          Font.Color = clSilver
          Font.Height = -13
          Font.Name = #24494#36719#38597#40657
          Font.Style = []
          ParentFont = False
        end
        object lbl198: TLabel
          Left = 56
          Top = 1009
          Width = 636
          Height = 19
          Caption = 
            #8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213 +
            #8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213
          Font.Charset = GB2312_CHARSET
          Font.Color = clSilver
          Font.Height = -13
          Font.Name = #24494#36719#38597#40657
          Font.Style = []
          ParentFont = False
        end
        object lbl199: TLabel
          Left = 56
          Top = 984
          Width = 636
          Height = 19
          Caption = 
            #8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213 +
            #8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213
          Font.Charset = GB2312_CHARSET
          Font.Color = clSilver
          Font.Height = -13
          Font.Name = #24494#36719#38597#40657
          Font.Style = []
          ParentFont = False
        end
        object lbl200: TLabel
          Left = 56
          Top = 960
          Width = 636
          Height = 19
          Caption = 
            #8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213 +
            #8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213
          Font.Charset = GB2312_CHARSET
          Font.Color = clSilver
          Font.Height = -13
          Font.Name = #24494#36719#38597#40657
          Font.Style = []
          ParentFont = False
        end
        object lbl201: TLabel
          Left = 56
          Top = 935
          Width = 636
          Height = 19
          Caption = 
            #8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213 +
            #8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213#8213
          Font.Charset = GB2312_CHARSET
          Font.Color = clSilver
          Font.Height = -13
          Font.Name = #24494#36719#38597#40657
          Font.Style = []
          ParentFont = False
        end
        object lbl194: TLabel
          Left = 576
          Top = 728
          Width = 93
          Height = 15
          Caption = '// 00 '#23548#20986#34920
          Font.Charset = GB2312_CHARSET
          Font.Color = clGreen
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl202: TLabel
          Left = 576
          Top = 752
          Width = 93
          Height = 15
          Caption = '// 01 '#23548#20837#34920
          Font.Charset = GB2312_CHARSET
          Font.Color = clGreen
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl203: TLabel
          Left = 576
          Top = 776
          Width = 108
          Height = 15
          Caption = '// 02 '#36164#28304#30446#24405
          Font.Charset = GB2312_CHARSET
          Font.Color = clGreen
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl204: TLabel
          Left = 576
          Top = 801
          Width = 108
          Height = 15
          Caption = '// 03 '#24322#24120#30446#24405
          Font.Charset = GB2312_CHARSET
          Font.Color = clGreen
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl205: TLabel
          Left = 576
          Top = 825
          Width = 108
          Height = 15
          Caption = '// 04 '#23433#20840#30446#24405
          Font.Charset = GB2312_CHARSET
          Font.Color = clGreen
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl206: TLabel
          Left = 576
          Top = 850
          Width = 138
          Height = 15
          Caption = '// 05 '#37325#23450#20301#22522#26412#34920
          Font.Charset = GB2312_CHARSET
          Font.Color = clGreen
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl207: TLabel
          Left = 576
          Top = 874
          Width = 108
          Height = 15
          Caption = '// 06 '#35843#35797#30446#24405
          Font.Charset = GB2312_CHARSET
          Font.Color = clGreen
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl208: TLabel
          Left = 576
          Top = 899
          Width = 108
          Height = 15
          Caption = '// 07 '#25551#26415#23383#20018
          Font.Charset = GB2312_CHARSET
          Font.Color = clGreen
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl209: TLabel
          Left = 576
          Top = 923
          Width = 93
          Height = 15
          Caption = '// 08 '#26426#22120#20540
          Font.Charset = GB2312_CHARSET
          Font.Color = clGreen
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl210: TLabel
          Left = 576
          Top = 948
          Width = 102
          Height = 15
          Caption = '// 09 TLS'#30446#24405
          Font.Charset = GB2312_CHARSET
          Font.Color = clGreen
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl211: TLabel
          Left = 576
          Top = 972
          Width = 138
          Height = 15
          Caption = '// 10 '#36733#20837#37197#20540#30446#24405
          Font.Charset = GB2312_CHARSET
          Font.Color = clGreen
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl212: TLabel
          Left = 576
          Top = 997
          Width = 123
          Height = 15
          Caption = '// 11 '#32465#23450#36755#20837#34920
          Font.Charset = GB2312_CHARSET
          Font.Color = clGreen
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl213: TLabel
          Left = 576
          Top = 1021
          Width = 123
          Height = 15
          Caption = '// 12 '#23548#20837#22320#22336#34920
          Font.Charset = GB2312_CHARSET
          Font.Color = clGreen
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl214: TLabel
          Left = 576
          Top = 1046
          Width = 138
          Height = 15
          Caption = '// 13 '#24310#36831#36733#20837#25551#36848
          Font.Charset = GB2312_CHARSET
          Font.Color = clGreen
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl215: TLabel
          Left = 576
          Top = 1070
          Width = 102
          Height = 15
          Caption = '// 14 COM'#20449#24687
          Font.Charset = GB2312_CHARSET
          Font.Color = clGreen
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object lbl216: TLabel
          Left = 576
          Top = 1095
          Width = 78
          Height = 15
          Caption = '// 15 '#20445#30041
          Font.Charset = GB2312_CHARSET
          Font.Color = clGreen
          Font.Height = -15
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object btnOptionHeaderMagic: TButton
          Left = 440
          Top = 8
          Width = 75
          Height = 25
          Caption = #21487#36873#20540
          TabOrder = 0
          OnClick = btnOptionHeaderMagicClick
        end
        object btnSubSystem: TButton
          Left = 453
          Top = 516
          Width = 75
          Height = 25
          Caption = #21487#36873#20540
          TabOrder = 1
          OnClick = btnSubSystemClick
        end
        object btnDll: TButton
          Left = 453
          Top = 541
          Width = 75
          Height = 25
          Caption = #21487#36873#20540
          TabOrder = 2
          OnClick = btnDllClick
        end
      end
      object btnMachine: TButton
        Left = 464
        Top = 60
        Width = 75
        Height = 25
        Caption = #21487#36873#20540
        TabOrder = 1
        OnClick = btnMachineClick
      end
    end
    object tsSectionTable: TTabSheet
      Caption = #33410#34920
      ImageIndex = 3
      ExplicitLeft = 0
      ExplicitTop = 0
      ExplicitWidth = 0
      ExplicitHeight = 0
      DesignSize = (
        790
        677)
      object lvSectionTable: TListView
        Left = 12
        Top = 12
        Width = 590
        Height = 653
        Anchors = [akLeft, akTop, akRight, akBottom]
        Columns = <
          item
            Caption = 'Name'
            Width = 80
          end
          item
            Caption = 'VSize'
            Width = 80
          end
          item
            Caption = 'VOffset'
            Width = 100
          end
          item
            Caption = 'RSize'
            Width = 100
          end
          item
            Caption = 'ROffset'
            Width = 100
          end
          item
            Caption = 'Characteristics'
            Width = 120
          end>
        Font.Charset = GB2312_CHARSET
        Font.Color = clWindowText
        Font.Height = -14
        Font.Name = #23435#20307
        Font.Style = []
        GridLines = True
        ReadOnly = True
        RowSelect = True
        ParentFont = False
        TabOrder = 0
        ViewStyle = vsReport
        OnClick = lvSectionTableClick
      end
      object pnl2: TPanel
        Left = 601
        Top = 12
        Width = 171
        Height = 653
        Anchors = [akTop, akRight, akBottom]
        BevelOuter = bvNone
        BorderStyle = bsSingle
        Caption = 'pnl2'
        Ctl3D = False
        ParentCtl3D = False
        ShowCaption = False
        TabOrder = 1
        object pnl3: TPanel
          Left = 7
          Top = -2
          Width = 161
          Height = 421
          BevelOuter = bvNone
          Caption = 'pnl1'
          Enabled = False
          ShowCaption = False
          TabOrder = 0
          object chk1: TCheckBox
            Tag = 268435456
            Left = 2
            Top = 8
            Width = 160
            Height = 17
            Caption = #20869#23384#20013#21487#20849#20139
            Font.Charset = GB2312_CHARSET
            Font.Color = clWindowText
            Font.Height = -13
            Font.Name = #23435#20307
            Font.Style = []
            ParentFont = False
            TabOrder = 0
          end
          object chk2: TCheckBox
            Tag = 536870912
            Left = 2
            Top = 35
            Width = 160
            Height = 17
            Caption = #20316#20026#20195#30721#25191#34892
            Font.Charset = GB2312_CHARSET
            Font.Color = clWindowText
            Font.Height = -13
            Font.Name = #23435#20307
            Font.Style = []
            ParentFont = False
            TabOrder = 1
          end
          object chk3: TCheckBox
            Tag = 1073741824
            Left = 2
            Top = 62
            Width = 160
            Height = 17
            Caption = #21487#35835#21462
            Font.Charset = GB2312_CHARSET
            Font.Color = clWindowText
            Font.Height = -13
            Font.Name = #23435#20307
            Font.Style = []
            ParentFont = False
            TabOrder = 2
          end
          object chk4: TCheckBox
            Tag = 2147483647
            Left = 2
            Top = 90
            Width = 160
            Height = 17
            Caption = #21487#20889#20837
            Font.Charset = GB2312_CHARSET
            Font.Color = clWindowText
            Font.Height = -13
            Font.Name = #23435#20307
            Font.Style = []
            ParentFont = False
            TabOrder = 3
          end
          object chk5: TCheckBox
            Tag = 16777216
            Left = 2
            Top = 117
            Width = 160
            Height = 17
            Caption = #21253#21547#25193#23637#37325#23450#20301
            Font.Charset = GB2312_CHARSET
            Font.Color = clWindowText
            Font.Height = -13
            Font.Name = #23435#20307
            Font.Style = []
            ParentFont = False
            TabOrder = 4
          end
          object chk6: TCheckBox
            Tag = 33554432
            Left = 2
            Top = 145
            Width = 160
            Height = 17
            Caption = #38656#35201#26102#21487#24223#24323
            Font.Charset = GB2312_CHARSET
            Font.Color = clWindowText
            Font.Height = -13
            Font.Name = #23435#20307
            Font.Style = []
            ParentFont = False
            TabOrder = 5
          end
          object chk7: TCheckBox
            Tag = 67108864
            Left = 2
            Top = 172
            Width = 160
            Height = 17
            Caption = #19981#21487#32531#23384
            Font.Charset = GB2312_CHARSET
            Font.Color = clWindowText
            Font.Height = -13
            Font.Name = #23435#20307
            Font.Style = []
            ParentFont = False
            TabOrder = 6
          end
          object chk8: TCheckBox
            Tag = 134217728
            Left = 2
            Top = 200
            Width = 160
            Height = 17
            Caption = #19981#21487#20998#39029
            Font.Charset = GB2312_CHARSET
            Font.Color = clWindowText
            Font.Height = -13
            Font.Name = #23435#20307
            Font.Style = []
            ParentFont = False
            TabOrder = 7
          end
          object chk9: TCheckBox
            Tag = 4096
            Left = 2
            Top = 227
            Width = 160
            Height = 17
            Caption = #21253#21547' COMDAT '#25968#25454
            Font.Charset = GB2312_CHARSET
            Font.Color = clWindowText
            Font.Height = -13
            Font.Name = #23435#20307
            Font.Style = []
            ParentFont = False
            TabOrder = 8
          end
          object chk10: TCheckBox
            Tag = 512
            Left = 2
            Top = 254
            Width = 160
            Height = 17
            Caption = #21253#21547#27880#37322#25110#20854#20182#20449#24687
            Font.Charset = GB2312_CHARSET
            Font.Color = clWindowText
            Font.Height = -13
            Font.Name = #23435#20307
            Font.Style = []
            ParentFont = False
            TabOrder = 9
          end
          object chk11: TCheckBox
            Tag = 2048
            Left = 2
            Top = 282
            Width = 160
            Height = 17
            Caption = #19981#20250#25104#20026#38236#20687#19968#37096#20998
            Font.Charset = GB2312_CHARSET
            Font.Color = clWindowText
            Font.Height = -13
            Font.Name = #23435#20307
            Font.Style = []
            ParentFont = False
            TabOrder = 10
          end
          object chk12: TCheckBox
            Tag = 32
            Left = 2
            Top = 309
            Width = 160
            Height = 17
            Caption = #21253#21547#21487#25191#34892#20195#30721
            Font.Charset = GB2312_CHARSET
            Font.Color = clWindowText
            Font.Height = -13
            Font.Name = #23435#20307
            Font.Style = []
            ParentFont = False
            TabOrder = 11
          end
          object chk13: TCheckBox
            Tag = 64
            Left = 2
            Top = 337
            Width = 160
            Height = 17
            Caption = #21253#21547#24050#21021#22987#21270#25968#25454
            Font.Charset = GB2312_CHARSET
            Font.Color = clWindowText
            Font.Height = -13
            Font.Name = #23435#20307
            Font.Style = []
            ParentFont = False
            TabOrder = 12
          end
          object chk14: TCheckBox
            Tag = 128
            Left = 2
            Top = 364
            Width = 160
            Height = 17
            Caption = #21253#21547#26410#21021#22987#21270#20449#24687
            Font.Charset = GB2312_CHARSET
            Font.Color = clWindowText
            Font.Height = -13
            Font.Name = #23435#20307
            Font.Style = []
            ParentFont = False
            TabOrder = 13
          end
          object chk15: TCheckBox
            Tag = 8
            Left = 2
            Top = 392
            Width = 160
            Height = 17
            Caption = #19981#20250#34987#22635#34917#21040#19979#19968#36793#30028
            Font.Charset = GB2312_CHARSET
            Font.Color = clWindowText
            Font.Height = -13
            Font.Name = #23435#20307
            Font.Style = []
            ParentFont = False
            TabOrder = 14
          end
        end
      end
    end
    object tsSectionData: TTabSheet
      Caption = #33410#25968#25454
      ImageIndex = 4
      ExplicitLeft = 0
      ExplicitTop = 0
      ExplicitWidth = 0
      ExplicitHeight = 0
      DesignSize = (
        790
        677)
      object lvSectionData: TListView
        Left = 12
        Top = 12
        Width = 763
        Height = 654
        Anchors = [akLeft, akTop, akRight, akBottom]
        Columns = <
          item
            Caption = 'Name'
            Width = 80
          end
          item
            Caption = 'VSize'
            Width = 80
          end
          item
            Caption = 'VOffset'
            Width = 100
          end
          item
            Caption = 'RSize'
            Width = 100
          end
          item
            Caption = 'ROffset'
            Width = 100
          end
          item
            Caption = 'Characteristics'
            Width = 120
          end>
        Font.Charset = GB2312_CHARSET
        Font.Color = clWindowText
        Font.Height = -14
        Font.Name = #23435#20307
        Font.Style = []
        GridLines = True
        ReadOnly = True
        RowSelect = True
        ParentFont = False
        TabOrder = 0
        ViewStyle = vsReport
        OnClick = lvSectionDataClick
      end
    end
    object tsExport: TTabSheet
      Caption = #23548#20986#20989#25968
      ImageIndex = 5
      ExplicitLeft = 0
      ExplicitTop = 0
      ExplicitWidth = 0
      ExplicitHeight = 0
      DesignSize = (
        790
        677)
      object lvFunc: TListView
        Left = 8
        Top = 131
        Width = 767
        Height = 535
        Anchors = [akLeft, akTop, akRight, akBottom]
        Columns = <
          item
            Caption = #24207#21015
          end
          item
            Caption = #22320#22336'(VA)'
            Width = 100
          end
          item
            Caption = #20559#31227'(RA)'
            Width = 100
          end
          item
            Caption = #21517#31216
            Width = 300
          end>
        Font.Charset = GB2312_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = #23435#20307
        Font.Style = []
        GridLines = True
        ReadOnly = True
        RowSelect = True
        ParentFont = False
        TabOrder = 0
        ViewStyle = vsReport
        OnClick = lvFuncClick
      end
      object grpExport: TGroupBox
        Left = 8
        Top = 8
        Width = 767
        Height = 117
        Anchors = [akLeft, akTop, akRight]
        Caption = #36755#20986#20449#24687#65306
        TabOrder = 1
        object Label3: TLabel
          Left = 12
          Top = 19
          Width = 72
          Height = 12
          Caption = #36755#20986#34920#20559#31227#65306
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlue
          Font.Height = -12
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object Label4: TLabel
          Left = 12
          Top = 38
          Width = 72
          Height = 12
          Caption = #29305#24449#20540'    '#65306
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlue
          Font.Height = -12
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object Label5: TLabel
          Left = 12
          Top = 58
          Width = 72
          Height = 12
          Caption = #22522#22336'      '#65306
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlue
          Font.Height = -12
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object Label6: TLabel
          Left = 12
          Top = 77
          Width = 72
          Height = 12
          Caption = #21517#31216#22320#22336'  '#65306
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlue
          Font.Height = -12
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object Label7: TLabel
          Left = 12
          Top = 97
          Width = 72
          Height = 12
          Caption = #21517#31216#23383#31526'  '#65306
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlue
          Font.Height = -12
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object Label8: TLabel
          Left = 120
          Top = 97
          Width = 90
          Height = 12
          Caption = 'Characteristics'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object Label9: TLabel
          Left = 120
          Top = 77
          Width = 120
          Height = 12
          Caption = 'SizeOfOptionalHeader'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object Label10: TLabel
          Left = 120
          Top = 58
          Width = 90
          Height = 12
          Caption = 'NumberOfSymbols'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object Label11: TLabel
          Left = 120
          Top = 38
          Width = 120
          Height = 12
          Caption = 'PointerToSymbolTable'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object Label12: TLabel
          Left = 120
          Top = 19
          Width = 78
          Height = 12
          Caption = 'TimeDateStamp'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object Label13: TLabel
          Left = 296
          Top = 19
          Width = 96
          Height = 12
          Caption = #20989#25968#25968#37327'      '#65306
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlue
          Font.Height = -12
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object Label14: TLabel
          Left = 296
          Top = 38
          Width = 96
          Height = 12
          Caption = #20989#25968#21517#25968#37327'    '#65306
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlue
          Font.Height = -12
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object Label15: TLabel
          Left = 296
          Top = 58
          Width = 96
          Height = 12
          Caption = #20989#25968#22320#22336'      '#65306
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlue
          Font.Height = -12
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object Label16: TLabel
          Left = 296
          Top = 77
          Width = 96
          Height = 12
          Caption = #20989#25968#21517#31216#22320#22336'  '#65306
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlue
          Font.Height = -12
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object Label17: TLabel
          Left = 296
          Top = 97
          Width = 96
          Height = 12
          Caption = #20989#25968#21517#24207#21495#22320#22336#65306
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlue
          Font.Height = -12
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object Label18: TLabel
          Left = 404
          Top = 19
          Width = 78
          Height = 12
          Caption = 'TimeDateStamp'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object Label19: TLabel
          Left = 404
          Top = 38
          Width = 120
          Height = 12
          Caption = 'PointerToSymbolTable'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object Label20: TLabel
          Left = 404
          Top = 58
          Width = 90
          Height = 12
          Caption = 'NumberOfSymbols'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object Label21: TLabel
          Left = 404
          Top = 77
          Width = 120
          Height = 12
          Caption = 'SizeOfOptionalHeader'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
        object Label22: TLabel
          Left = 404
          Top = 97
          Width = 90
          Height = 12
          Caption = 'Characteristics'
          Font.Charset = GB2312_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = #23435#20307
          Font.Style = []
          ParentFont = False
        end
      end
    end
    object tsImport: TTabSheet
      Caption = #23548#20837#20989#25968
      ImageIndex = 6
      ExplicitLeft = 0
      ExplicitTop = 0
      ExplicitWidth = 0
      ExplicitHeight = 0
      DesignSize = (
        790
        677)
      object lvImportDllFileName: TListView
        Left = 8
        Top = 3
        Width = 767
        Height = 166
        Anchors = [akLeft, akTop, akRight]
        Columns = <
          item
            Caption = 'Dll '#25991#20214#21517#31216
            Width = 120
          end
          item
            Caption = 'OriginalFirstThunk'
            Width = 140
          end
          item
            Caption = 'TimeDateStamp'
            Width = 110
          end
          item
            Caption = 'ForwarderChain'
            Width = 120
          end
          item
            Caption = 'Name'
            Width = 100
          end
          item
            Caption = 'FirstThunk'
            Width = 120
          end>
        Font.Charset = GB2312_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = #23435#20307
        Font.Style = []
        GridLines = True
        ReadOnly = True
        RowSelect = True
        ParentFont = False
        TabOrder = 0
        ViewStyle = vsReport
        OnClick = lvImportDllFileNameClick
      end
      object lvImportFuncList: TListView
        Left = 8
        Top = 175
        Width = 767
        Height = 491
        Anchors = [akLeft, akTop, akRight, akBottom]
        Columns = <
          item
            Caption = 'ThunkRVA'
            Width = 120
          end
          item
            Caption = 'ThunkValue'
            Width = 140
          end
          item
            Caption = 'Hint'
            Width = 110
          end
          item
            Caption = #20989#25968#21517#31216
            Width = 220
          end>
        Font.Charset = GB2312_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = #23435#20307
        Font.Style = []
        GridLines = True
        ReadOnly = True
        RowSelect = True
        ParentFont = False
        TabOrder = 1
        ViewStyle = vsReport
        OnClick = lvImportFuncListClick
      end
    end
    object tsResource: TTabSheet
      Caption = #36164#28304
      ImageIndex = 7
      ExplicitLeft = 0
      ExplicitTop = 0
      ExplicitWidth = 0
      ExplicitHeight = 0
      DesignSize = (
        790
        677)
      object tvResource: TTreeView
        Left = 8
        Top = 8
        Width = 779
        Height = 661
        Anchors = [akLeft, akTop, akRight, akBottom]
        Font.Charset = GB2312_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = #23435#20307
        Font.Style = []
        Indent = 19
        ParentFont = False
        TabOrder = 0
      end
    end
  end
  object dlgOpenSelectPEFile: TOpenDialog
    Filter = 'PE(*.DLL;*.EXE)|*.EXE;*.DLL'
    Left = 104
    Top = 136
  end
  object pmDosMagic: TPopupMenu
    AutoHotkeys = maManual
    OwnerDraw = True
    Left = 345
    Top = 139
    object IMAGEDOSSIGNATURE5A4DMZ1: TMenuItem
      Caption = 'IMAGE_DOS_SIGNATURE      = $5A4D; { MZ }'
      OnDrawItem = IMAGEVXDSIGNATURE454CLE1DrawItem
      OnMeasureItem = IMAGEVXDSIGNATURE454CLE1MeasureItem
    end
    object IMAGEOS2SIGNATURE454ENE1: TMenuItem
      Caption = 'IMAGE_OS2_SIGNATURE      = $454E; { NE }'
      OnDrawItem = IMAGEVXDSIGNATURE454CLE1DrawItem
      OnMeasureItem = IMAGEVXDSIGNATURE454CLE1MeasureItem
    end
    object IMAGEOS2SIGNATURELE454CLE1: TMenuItem
      Caption = 'IMAGE_OS2_SIGNATURE_LE   = $454C; { LE }'
      OnDrawItem = IMAGEVXDSIGNATURE454CLE1DrawItem
      OnMeasureItem = IMAGEVXDSIGNATURE454CLE1MeasureItem
    end
    object IMAGEVXDSIGNATURE454CLE1: TMenuItem
      Caption = 'IMAGE_VXD_SIGNATURE      = $454C; { LE }'
      OnDrawItem = IMAGEVXDSIGNATURE454CLE1DrawItem
      OnMeasureItem = IMAGEVXDSIGNATURE454CLE1MeasureItem
    end
  end
  object pmOptionHeaderMagic: TPopupMenu
    AutoHotkeys = maManual
    OwnerDraw = True
    Left = 340
    Top = 272
    object IMAGENTOPTIONALHDR32MAGIC010B1: TMenuItem
      Caption = '  IMAGE_NT_OPTIONAL_HDR32_MAGIC    = $010B;'
    end
    object IMAGENTOPTIONALHDR64MAGIC020B1: TMenuItem
      Caption = '  IMAGE_NT_OPTIONAL_HDR64_MAGIC    = $020B;'
    end
    object IMAGEROMOPTIONALHDRMAGIC01071: TMenuItem
      Caption = '  IMAGE_ROM_OPTIONAL_HDR_MAGIC     = $0107;'
    end
  end
  object pmOptionHeaderSubSystem: TPopupMenu
    AutoHotkeys = maManual
    OwnerDraw = True
    Left = 333
    Top = 352
    object IMAGESUBSYSTEMUNKNOWN0Unknownsubsystem1: TMenuItem
      Caption = 
        'IMAGE_SUBSYSTEM_UNKNOWN                  = $0000   // Unknown su' +
        'bsystem.  '
      OnDrawItem = IMAGESUBSYSTEMWINDOWSBOOTAPPLICATION161DrawItem
      OnMeasureItem = IMAGESUBSYSTEMWINDOWSBOOTAPPLICATION161MeasureItem
    end
    object IMAGESUBSYSTEMNATIVE1Imagedoesntrequireasubsystem1: TMenuItem
      Caption = 
        'IMAGE_SUBSYSTEM_NATIVE                   = $0001   // Image does' +
        'n'#39't require a subsystem.  '
      OnDrawItem = IMAGESUBSYSTEMWINDOWSBOOTAPPLICATION161DrawItem
      OnMeasureItem = IMAGESUBSYSTEMWINDOWSBOOTAPPLICATION161MeasureItem
    end
    object IMAGESUBSYSTEMWINDOWSGUI2ImagerunsintheWindowsGUIsubsystem1: TMenuItem
      Caption = 
        'IMAGE_SUBSYSTEM_WINDOWS_GUI              = $0002   // Image runs' +
        ' in the Windows GUI subsystem.'
      OnDrawItem = IMAGESUBSYSTEMWINDOWSBOOTAPPLICATION161DrawItem
      OnMeasureItem = IMAGESUBSYSTEMWINDOWSBOOTAPPLICATION161MeasureItem
    end
    object IMAGESUBSYSTEMWINDOWSCUI3ImagerunsintheWindowscharactersubsystem1: TMenuItem
      Caption = 
        'IMAGE_SUBSYSTEM_WINDOWS_CUI              = $0003   // Image runs' +
        ' in the Windows character subsystem. '
      OnDrawItem = IMAGESUBSYSTEMWINDOWSBOOTAPPLICATION161DrawItem
      OnMeasureItem = IMAGESUBSYSTEMWINDOWSBOOTAPPLICATION161MeasureItem
    end
    object IMAGESUBSYSTEMOS2CUI5imagerunsintheOS2charactersubsystem1: TMenuItem
      Caption = 
        'IMAGE_SUBSYSTEM_OS2_CUI                  = $0005   // image runs' +
        ' in the OS/2 character subsystem.  '
      OnDrawItem = IMAGESUBSYSTEMWINDOWSBOOTAPPLICATION161DrawItem
      OnMeasureItem = IMAGESUBSYSTEMWINDOWSBOOTAPPLICATION161MeasureItem
    end
    object IMAGESUBSYSTEMPOSIXCUI7imagerunsinthePosixcharactersubsystem1: TMenuItem
      Caption = 
        'IMAGE_SUBSYSTEM_POSIX_CUI                = $0007   // image runs' +
        ' in the Posix character subsystem. '
      OnDrawItem = IMAGESUBSYSTEMWINDOWSBOOTAPPLICATION161DrawItem
      OnMeasureItem = IMAGESUBSYSTEMWINDOWSBOOTAPPLICATION161MeasureItem
    end
    object IMAGESUBSYSTEMNATIVEWINDOWS8imageisanativeWin9xdriver1: TMenuItem
      Caption = 
        'IMAGE_SUBSYSTEM_NATIVE_WINDOWS           = $0008   // image is a' +
        ' native Win9x driver. '
      OnDrawItem = IMAGESUBSYSTEMWINDOWSBOOTAPPLICATION161DrawItem
      OnMeasureItem = IMAGESUBSYSTEMWINDOWSBOOTAPPLICATION161MeasureItem
    end
    object IMAGESUBSYSTEMWINDOWSCEGUI9ImagerunsintheWindowsCEsubsystem1: TMenuItem
      Caption = 
        'IMAGE_SUBSYSTEM_WINDOWS_CE_GUI           = $0009   // Image runs' +
        ' in the Windows CE subsystem.  '
      OnDrawItem = IMAGESUBSYSTEMWINDOWSBOOTAPPLICATION161DrawItem
      OnMeasureItem = IMAGESUBSYSTEMWINDOWSBOOTAPPLICATION161MeasureItem
    end
    object IMAGESUBSYSTEMEFIAPPLICATION101: TMenuItem
      Caption = 'IMAGE_SUBSYSTEM_EFI_APPLICATION          = $000A'
      OnDrawItem = IMAGESUBSYSTEMWINDOWSBOOTAPPLICATION161DrawItem
      OnMeasureItem = IMAGESUBSYSTEMWINDOWSBOOTAPPLICATION161MeasureItem
    end
    object IMAGESUBSYSTEMEFIBOOTSERVICEDRIVER111: TMenuItem
      Caption = 'IMAGE_SUBSYSTEM_EFI_BOOT_SERVICE_DRIVER  = $000B'
      OnDrawItem = IMAGESUBSYSTEMWINDOWSBOOTAPPLICATION161DrawItem
      OnMeasureItem = IMAGESUBSYSTEMWINDOWSBOOTAPPLICATION161MeasureItem
    end
    object IMAGESUBSYSTEMEFIRUNTIMEDRIVER121: TMenuItem
      Caption = 'IMAGE_SUBSYSTEM_EFI_RUNTIME_DRIVER       = $000C'
      OnDrawItem = IMAGESUBSYSTEMWINDOWSBOOTAPPLICATION161DrawItem
      OnMeasureItem = IMAGESUBSYSTEMWINDOWSBOOTAPPLICATION161MeasureItem
    end
    object IMAGESUBSYSTEMEFIROM131: TMenuItem
      Caption = 'IMAGE_SUBSYSTEM_EFI_ROM                  = $000D'
      OnDrawItem = IMAGESUBSYSTEMWINDOWSBOOTAPPLICATION161DrawItem
      OnMeasureItem = IMAGESUBSYSTEMWINDOWSBOOTAPPLICATION161MeasureItem
    end
    object IMAGESUBSYSTEMXBOX141: TMenuItem
      Caption = 'IMAGE_SUBSYSTEM_XBOX                     = $000E'
      OnDrawItem = IMAGESUBSYSTEMWINDOWSBOOTAPPLICATION161DrawItem
      OnMeasureItem = IMAGESUBSYSTEMWINDOWSBOOTAPPLICATION161MeasureItem
    end
    object IMAGESUBSYSTEMWINDOWSBOOTAPPLICATION161: TMenuItem
      Caption = 'IMAGE_SUBSYSTEM_WINDOWS_BOOT_APPLICATION = $0010'
      OnDrawItem = IMAGESUBSYSTEMWINDOWSBOOTAPPLICATION161DrawItem
      OnMeasureItem = IMAGESUBSYSTEMWINDOWSBOOTAPPLICATION161MeasureItem
    end
  end
  object pmDll: TPopupMenu
    AutoHotkeys = maManual
    OwnerDraw = True
    Left = 425
    Top = 271
    object IMAGEDLLCHARACTERISTICSDYNAMICBASE0x0040DLLcanmove1: TMenuItem
      Caption = 
        'IMAGE_DLLCHARACTERISTICS_DYNAMIC_BASE              = $0040     /' +
        '/ DLL can move.  '
      OnDrawItem = mniIMAGEDLLCHARACTERISTICSTERMINALSERVERAWARE0x80001DrawItem
      OnMeasureItem = mniIMAGEDLLCHARACTERISTICSTERMINALSERVERAWARE0x80001MeasureItem
    end
    object IMAGEDLLCHARACTERISTICSFORCEINTEGRITY0x0080CodeIntegrityImage1: TMenuItem
      Caption = 
        'IMAGE_DLLCHARACTERISTICS_FORCE_INTEGRITY           = $0080     /' +
        '/ Code Integrity Image'
      OnDrawItem = mniIMAGEDLLCHARACTERISTICSTERMINALSERVERAWARE0x80001DrawItem
      OnMeasureItem = mniIMAGEDLLCHARACTERISTICSTERMINALSERVERAWARE0x80001MeasureItem
    end
    object IMAGEDLLCHARACTERISTICSNXCOMPAT0x0100ImageisNXcompatible1: TMenuItem
      Caption = 
        'IMAGE_DLLCHARACTERISTICS_NX_COMPAT                 = $0100     /' +
        '/ Image is NX compatible  '
      OnDrawItem = mniIMAGEDLLCHARACTERISTICSTERMINALSERVERAWARE0x80001DrawItem
      OnMeasureItem = mniIMAGEDLLCHARACTERISTICSTERMINALSERVERAWARE0x80001MeasureItem
    end
    object IMAGEDLLCHARACTERISTICSNOISOLATION0x0200Imageunderstandsisolationanddoesntwantit1: TMenuItem
      Caption = 
        'IMAGE_DLLCHARACTERISTICS_NO_ISOLATION              = $0200     /' +
        '/ Image understands isolation and doesn'#39't want it'
      OnDrawItem = mniIMAGEDLLCHARACTERISTICSTERMINALSERVERAWARE0x80001DrawItem
      OnMeasureItem = mniIMAGEDLLCHARACTERISTICSTERMINALSERVERAWARE0x80001MeasureItem
    end
    object IMAGEDLLCHARACTERISTICSNOSEH0x0400ImagedoesnotuseSEHNoSEhandlermayresidei1: TMenuItem
      Caption = 
        'IMAGE_DLLCHARACTERISTICS_NO_SEH                    = $0400     /' +
        '/ Image does not use SEH.  No SE handler may reside in this imag' +
        'e'
      OnDrawItem = mniIMAGEDLLCHARACTERISTICSTERMINALSERVERAWARE0x80001DrawItem
      OnMeasureItem = mniIMAGEDLLCHARACTERISTICSTERMINALSERVERAWARE0x80001MeasureItem
    end
    object IMAGEDLLCHARACTERISTICSNOBIND0x0800Donotbindthisimage1: TMenuItem
      Caption = 
        'IMAGE_DLLCHARACTERISTICS_NO_BIND                   = $0800     /' +
        '/ Do not bind this image.'
      OnDrawItem = mniIMAGEDLLCHARACTERISTICSTERMINALSERVERAWARE0x80001DrawItem
      OnMeasureItem = mniIMAGEDLLCHARACTERISTICSTERMINALSERVERAWARE0x80001MeasureItem
    end
    object N0x1000Reserved1: TMenuItem
      Caption = 
        '                                                   = $1000     /' +
        '/ Reserved. '
      OnDrawItem = mniIMAGEDLLCHARACTERISTICSTERMINALSERVERAWARE0x80001DrawItem
      OnMeasureItem = mniIMAGEDLLCHARACTERISTICSTERMINALSERVERAWARE0x80001MeasureItem
    end
    object IMAGEDLLCHARACTERISTICSWDMDRIVER0x2000DriverusesWDMmodel1: TMenuItem
      Caption = 
        'IMAGE_DLLCHARACTERISTICS_WDM_DRIVER                = $2000     /' +
        '/ Driver uses WDM model  '
      OnDrawItem = mniIMAGEDLLCHARACTERISTICSTERMINALSERVERAWARE0x80001DrawItem
      OnMeasureItem = mniIMAGEDLLCHARACTERISTICSTERMINALSERVERAWARE0x80001MeasureItem
    end
    object N0x4000Reserved1: TMenuItem
      Caption = 
        '                                                   = $4000     /' +
        '/ Reserved. '
      OnDrawItem = mniIMAGEDLLCHARACTERISTICSTERMINALSERVERAWARE0x80001DrawItem
      OnMeasureItem = mniIMAGEDLLCHARACTERISTICSTERMINALSERVERAWARE0x80001MeasureItem
    end
    object mniIMAGEDLLCHARACTERISTICSTERMINALSERVERAWARE0x80001: TMenuItem
      Caption = 'IMAGE_DLLCHARACTERISTICS_TERMINAL_SERVER_AWARE     = $8000  '
      OnDrawItem = mniIMAGEDLLCHARACTERISTICSTERMINALSERVERAWARE0x80001DrawItem
      OnMeasureItem = mniIMAGEDLLCHARACTERISTICSTERMINALSERVERAWARE0x80001MeasureItem
    end
  end
  object pmMachine: TPopupMenu
    AutoHotkeys = maManual
    OwnerDraw = True
    Left = 236
    Top = 80
    object MenuItem1: TMenuItem
      Caption = '  IMAGE_FILE_MACHINE_UNKNOWN       = $0000;'
      OnDrawItem = IMAGEFILEMACHINEAMD648664AMD64K81DrawItem
      OnMeasureItem = IMAGEFILEMACHINEAMD648664AMD64K81MeasureItem
    end
    object MenuItem2: TMenuItem
      Caption = '  IMAGE_FILE_MACHINE_I386          = $014C;  { Intel 386. }'
      OnDrawItem = IMAGEFILEMACHINEAMD648664AMD64K81DrawItem
      OnMeasureItem = IMAGEFILEMACHINEAMD648664AMD64K81MeasureItem
    end
    object MenuItem3: TMenuItem
      Caption = 
        '  IMAGE_FILE_MACHINE_R3000         = $0162;  { MIPS little-endia' +
        'n, 0x160 big-endian }'
      OnDrawItem = IMAGEFILEMACHINEAMD648664AMD64K81DrawItem
      OnMeasureItem = IMAGEFILEMACHINEAMD648664AMD64K81MeasureItem
    end
    object MenuItem4: TMenuItem
      Caption = 
        '  IMAGE_FILE_MACHINE_R4000         = $0166;  { MIPS little-endia' +
        'n }'
      OnDrawItem = IMAGEFILEMACHINEAMD648664AMD64K81DrawItem
      OnMeasureItem = IMAGEFILEMACHINEAMD648664AMD64K81MeasureItem
    end
    object MenuItem5: TMenuItem
      Caption = 
        '  IMAGE_FILE_MACHINE_R10000        = $0168;  { MIPS little-endia' +
        'n }'
      OnDrawItem = IMAGEFILEMACHINEAMD648664AMD64K81DrawItem
      OnMeasureItem = IMAGEFILEMACHINEAMD648664AMD64K81MeasureItem
    end
    object MenuItem6: TMenuItem
      Caption = '  IMAGE_FILE_MACHINE_ALPHA         = $0184;  { Alpha_AXP }'
      OnDrawItem = IMAGEFILEMACHINEAMD648664AMD64K81DrawItem
      OnMeasureItem = IMAGEFILEMACHINEAMD648664AMD64K81MeasureItem
    end
    object MenuItem7: TMenuItem
      Caption = 
        '  IMAGE_FILE_MACHINE_POWERPC       = $01F0;  { IBM PowerPC Littl' +
        'e-Endian }'
      OnDrawItem = IMAGEFILEMACHINEAMD648664AMD64K81DrawItem
      OnMeasureItem = IMAGEFILEMACHINEAMD648664AMD64K81MeasureItem
    end
    object MenuItem8: TMenuItem
      Caption = '  IMAGE_FILE_MACHINE_IA64          = $0200;  { Intel 64 }'
      OnDrawItem = IMAGEFILEMACHINEAMD648664AMD64K81DrawItem
      OnMeasureItem = IMAGEFILEMACHINEAMD648664AMD64K81MeasureItem
    end
    object MenuItem9: TMenuItem
      Caption = '  IMAGE_FILE_MACHINE_ALPHA64       = $0284;  { Alpha_64 }'
      OnDrawItem = IMAGEFILEMACHINEAMD648664AMD64K81DrawItem
      OnMeasureItem = IMAGEFILEMACHINEAMD648664AMD64K81MeasureItem
    end
    object MenuItem10: TMenuItem
      Caption = '  IMAGE_FILE_MACHINE_AMD64         = $8664;  { AMD64 (K8) }'
      OnDrawItem = IMAGEFILEMACHINEAMD648664AMD64K81DrawItem
      OnMeasureItem = IMAGEFILEMACHINEAMD648664AMD64K81MeasureItem
    end
  end
end
