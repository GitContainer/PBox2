object frmDBView: TfrmDBView
  Left = 0
  Top = 0
  Caption = #25968#25454#26597#30475#22120' v2.0'
  ClientHeight = 750
  ClientWidth = 1228
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnClose = FormClose
  OnCreate = FormCreate
  DesignSize = (
    1228
    750)
  PixelsPerInch = 96
  TextHeight = 13
  object lblTip: TLabel
    Left = 940
    Top = 32
    Width = 267
    Height = 26
    Anchors = [akTop, akRight]
    Caption = #27809#26377#21457#29616#33258#22686#38271#23383#27573#65292#21482#33021#26368#22810#26174#31034'1'#19975#26465#35760#24405#13#10
    Font.Charset = GB2312_CHARSET
    Font.Color = clRed
    Font.Height = -13
    Font.Name = #23435#20307
    Font.Style = []
    ParentFont = False
    Visible = False
  end
  object btnDBLink: TButton
    Left = 8
    Top = 10
    Width = 177
    Height = 41
    Caption = #25968#25454#24211#36830#25509
    TabOrder = 0
    OnClick = btnDBLinkClick
  end
  object grpAllFields: TGroupBox
    Left = 8
    Top = 408
    Width = 177
    Height = 334
    Anchors = [akLeft, akTop, akBottom]
    Caption = #26174#31034#23383#27573#65306
    TabOrder = 1
    DesignSize = (
      177
      334)
    object chklstFields: TCheckListBox
      Left = 3
      Top = 22
      Width = 166
      Height = 302
      Anchors = [akLeft, akTop, akRight, akBottom]
      ItemHeight = 13
      TabOrder = 0
    end
  end
  object grpAllTable: TGroupBox
    Left = 8
    Top = 64
    Width = 177
    Height = 329
    Caption = #25152#26377#34920#65306
    TabOrder = 2
    DesignSize = (
      177
      329)
    object lstTable: TListBox
      Left = 3
      Top = 20
      Width = 166
      Height = 298
      Anchors = [akLeft, akTop, akBottom]
      ItemHeight = 13
      TabOrder = 0
      OnClick = lstTableClick
    end
  end
  object btnDBSearch: TButton
    Left = 191
    Top = 10
    Width = 177
    Height = 41
    Caption = #25968#25454#26597#35810
    Enabled = False
    TabOrder = 3
    OnClick = btnDBSearchClick
  end
  object grpDispalyData: TGroupBox
    Left = 191
    Top = 64
    Width = 1029
    Height = 678
    Anchors = [akLeft, akTop, akRight, akBottom]
    Caption = #25968#25454#26174#31034#65306
    TabOrder = 4
    DesignSize = (
      1029
      678)
    object lvData: TListView
      Left = 12
      Top = 24
      Width = 1004
      Height = 644
      Anchors = [akLeft, akTop, akRight, akBottom]
      Columns = <>
      Font.Charset = GB2312_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = #23435#20307
      Font.Style = []
      GridLines = True
      OwnerData = True
      ReadOnly = True
      RowSelect = True
      ParentFont = False
      ParentShowHint = False
      ShowHint = False
      TabOrder = 0
      ViewStyle = vsReport
      OnData = lvDataData
    end
  end
  object btnExportToExcel: TButton
    Left = 374
    Top = 10
    Width = 177
    Height = 41
    Caption = #25968#25454#23548#20986#21040' Excel'
    Enabled = False
    TabOrder = 5
    OnClick = btnExportToExcelClick
  end
  object ADOQuery1: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    Left = 244
    Top = 170
  end
  object ADOConnection1: TADOConnection
    LoginPrompt = False
    Left = 244
    Top = 112
  end
  object qryTemp: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    Left = 244
    Top = 229
  end
  object dlgSaveExcel: TSaveDialog
    Filter = 'EXCEL(*.XLSX)|*.XLSX'
    Left = 244
    Top = 288
  end
end
