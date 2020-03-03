object frmDllExport: TfrmDllExport
  Left = 0
  Top = 0
  Caption = 'Dll '#36755#20986#20989#25968
  ClientHeight = 569
  ClientWidth = 940
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnClose = FormClose
  OnResize = FormResize
  DesignSize = (
    940
    569)
  PixelsPerInch = 96
  TextHeight = 13
  object lbl1: TLabel
    Left = 14
    Top = 8
    Width = 107
    Height = 13
    Caption = #36873#25321#19968#20010' DLL '#25991#20214#65306
  end
  object lbl2: TLabel
    Left = 898
    Top = 8
    Width = 9
    Height = 15
    Anchors = [akTop, akRight]
    Font.Charset = GB2312_CHARSET
    Font.Color = clRed
    Font.Height = -15
    Font.Name = #23435#20307
    Font.Style = [fsBold]
    ParentFont = False
  end
  object srchbxDllFile: TSearchBox
    Left = 14
    Top = 28
    Width = 911
    Height = 21
    Anchors = [akLeft, akTop, akRight]
    ReadOnly = True
    TabOrder = 0
    OnDblClick = srchbxDllFileInvokeSearch
    OnInvokeSearch = srchbxDllFileInvokeSearch
  end
  object lvDllExport: TListView
    Left = 14
    Top = 59
    Width = 911
    Height = 496
    Anchors = [akLeft, akTop, akRight, akBottom]
    Columns = <
      item
        Caption = #24207#21015
        Width = 100
      end
      item
        Caption = #21517#31216
        Width = 500
      end
      item
        Caption = #20869#23384#20559#31227#22320#22336
        Width = 100
      end
      item
        Caption = #25991#20214#20559#31227#22320#22336
        Width = 100
      end>
    Font.Charset = GB2312_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = #23435#20307
    Font.Style = []
    GridLines = True
    RowSelect = True
    ParentFont = False
    TabOrder = 1
    ViewStyle = vsReport
  end
  object dlgOpenDllFile: TOpenDialog
    Filter = 'DLL(*.DLL)|*.DLL'
    Left = 240
    Top = 192
  end
end
