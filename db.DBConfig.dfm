object DBConfig: TDBConfig
  Left = 0
  Top = 0
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = #25968#25454#24211#37197#32622
  ClientHeight = 495
  ClientWidth = 832
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poMainFormCenter
  DesignSize = (
    832
    495)
  PixelsPerInch = 96
  TextHeight = 13
  object btnSave: TButton
    Left = 728
    Top = 451
    Width = 92
    Height = 36
    Anchors = [akRight, akBottom]
    Caption = #20445#23384
    Font.Charset = GB2312_CHARSET
    Font.Color = clWindowText
    Font.Height = -15
    Font.Name = #23435#20307
    Font.Style = []
    ParentFont = False
    TabOrder = 0
    OnClick = btnSaveClick
  end
  object btnCancel: TButton
    Left = 616
    Top = 451
    Width = 92
    Height = 36
    Anchors = [akRight, akBottom]
    Caption = #21462#28040
    Font.Charset = GB2312_CHARSET
    Font.Color = clWindowText
    Font.Height = -15
    Font.Name = #23435#20307
    Font.Style = []
    ParentFont = False
    TabOrder = 1
    OnClick = btnCancelClick
  end
  object pgcAll: TPageControl
    Left = 8
    Top = 8
    Width = 812
    Height = 437
    ActivePage = ts7
    MultiLine = True
    TabHeight = 35
    TabOrder = 2
    TabWidth = 89
    object tsCreateDBLink: TTabSheet
      Caption = #36830#25509
      ExplicitLeft = 0
      ExplicitTop = 0
      ExplicitWidth = 0
      ExplicitHeight = 0
      object btnCreateDBLink: TButton
        Left = 24
        Top = 24
        Width = 129
        Height = 41
        Caption = #24314#31435#25968#25454#24211#36830#25509
        TabOrder = 0
        OnClick = btnCreateDBLinkClick
      end
    end
    object ts2: TTabSheet
      Caption = #21019#24314
      ImageIndex = 1
      ExplicitLeft = 0
      ExplicitTop = 0
      ExplicitWidth = 0
      ExplicitHeight = 0
      object btnCreateDB: TButton
        Left = 24
        Top = 24
        Width = 129
        Height = 41
        Caption = #36873#25321#21019#24314#33050#26412
        TabOrder = 0
        OnClick = btnCreateDBClick
      end
    end
    object ts3: TTabSheet
      Caption = #21319#32423
      ImageIndex = 2
      ExplicitLeft = 0
      ExplicitTop = 0
      ExplicitWidth = 0
      ExplicitHeight = 0
      object lblAutoUpdateDBSQLScriptFileNameDelete: TLabel
        Left = 40
        Top = 147
        Width = 210
        Height = 15
        Caption = #21319#32423#25104#21151#21518#65292#20250#33258#21160#21024#38500#27492#25991#20214
        Font.Charset = GB2312_CHARSET
        Font.Color = clRed
        Font.Height = -15
        Font.Name = #23435#20307
        Font.Style = []
        ParentFont = False
        Visible = False
      end
      object lblAutoUpdateDBSQLScriptFileName: TLabel
        Left = 40
        Top = 123
        Width = 301
        Height = 15
        Caption = #21319#32423#33050#26412#25991#20214#21517'('#24517#39035#25918#32622#22312#20027#31243#24207#30446#24405#19979')'#65306
        Font.Charset = GB2312_CHARSET
        Font.Color = clRed
        Font.Height = -15
        Font.Name = #23435#20307
        Font.Style = []
        ParentFont = False
        Visible = False
      end
      object btnSelectUpdataDB: TButton
        Left = 24
        Top = 24
        Width = 129
        Height = 41
        Caption = #36873#25321#21319#32423#33050#26412
        TabOrder = 0
        OnClick = btnSelectUpdataDBClick
      end
      object chkAutoUpdateDB: TCheckBox
        Left = 24
        Top = 96
        Width = 273
        Height = 17
        Caption = #27599#27425#36719#20214#21551#21160#37117#26816#26597#21644#25191#34892#25968#25454#24211#21319#32423
        Font.Charset = GB2312_CHARSET
        Font.Color = clWindowText
        Font.Height = -15
        Font.Name = #23435#20307
        Font.Style = []
        ParentFont = False
        TabOrder = 1
        OnClick = chkAutoUpdateDBClick
      end
      object edtUpdateDBSqlScriptFileName: TEdit
        Left = 338
        Top = 119
        Width = 265
        Height = 23
        Font.Charset = GB2312_CHARSET
        Font.Color = clWindowText
        Font.Height = -15
        Font.Name = #23435#20307
        Font.Style = []
        ParentFont = False
        TabOrder = 2
        Text = 'update.sql'
        Visible = False
      end
    end
    object ts4: TTabSheet
      Caption = #25910#32553
      ImageIndex = 3
      ExplicitLeft = 0
      ExplicitTop = 0
      ExplicitWidth = 0
      ExplicitHeight = 0
      object btnZoomOut: TButton
        Left = 24
        Top = 24
        Width = 129
        Height = 41
        Caption = #25910#32553#25968#25454#24211
        TabOrder = 0
        OnClick = btnZoomOutClick
      end
    end
    object ts5: TTabSheet
      Caption = #22791#20221
      ImageIndex = 4
      ExplicitLeft = 0
      ExplicitTop = 0
      ExplicitWidth = 0
      ExplicitHeight = 0
      object lblLoginName: TLabel
        Left = 24
        Top = 84
        Width = 153
        Height = 16
        Caption = #26412#26426#30331#24405#29992#25143#21517#31216#65306
        Font.Charset = GB2312_CHARSET
        Font.Color = clBlue
        Font.Height = -16
        Font.Name = #23435#20307
        Font.Style = [fsBold]
        ParentFont = False
      end
      object lblLoginPass: TLabel
        Left = 24
        Top = 116
        Width = 153
        Height = 16
        Caption = #26412#26426#30331#24405#29992#25143#23494#30721#65306
        Font.Charset = GB2312_CHARSET
        Font.Color = clBlue
        Font.Height = -16
        Font.Name = #23435#20307
        Font.Style = [fsBold]
        ParentFont = False
      end
      object lblTip: TLabel
        Left = 24
        Top = 160
        Width = 708
        Height = 16
        Caption = #37325#35201#65306#26412#26426#24517#39035#24320#21551#25991#20214#21644#25171#21360#26426#20849#20139#65292#36828#31243#26426#22120'('#23433#35013#25968#25454#24211#30340#26426#22120')'#24517#39035#20851#38381'360'#31561#23433#20840#36719#20214
        Font.Charset = GB2312_CHARSET
        Font.Color = clRed
        Font.Height = -16
        Font.Name = #23435#20307
        Font.Style = [fsBold]
        ParentFont = False
      end
      object btnBackupDatabase: TButton
        Left = 24
        Top = 24
        Width = 129
        Height = 41
        Caption = #22791#20221#25968#25454#24211
        TabOrder = 0
        OnClick = btnBackupDatabaseClick
      end
      object edtLoginPass1: TEdit
        Left = 174
        Top = 111
        Width = 243
        Height = 29
        Font.Charset = SYMBOL_CHARSET
        Font.Color = clBlack
        Font.Height = -19
        Font.Name = 'Wingdings'
        Font.Style = [fsBold]
        ParentFont = False
        PasswordChar = 'l'
        TabOrder = 1
      end
      object edtLoginName1: TEdit
        Left = 174
        Top = 81
        Width = 243
        Height = 24
        Enabled = False
        Font.Charset = GB2312_CHARSET
        Font.Color = clBlack
        Font.Height = -16
        Font.Name = #23435#20307
        Font.Style = [fsBold]
        ParentFont = False
        TabOrder = 2
      end
    end
    object ts6: TTabSheet
      Caption = #36824#21407
      ImageIndex = 5
      ExplicitLeft = 0
      ExplicitTop = 0
      ExplicitWidth = 0
      ExplicitHeight = 0
      object Label1: TLabel
        Left = 24
        Top = 84
        Width = 153
        Height = 16
        Caption = #26412#26426#30331#24405#29992#25143#21517#31216#65306
        Font.Charset = GB2312_CHARSET
        Font.Color = clBlue
        Font.Height = -16
        Font.Name = #23435#20307
        Font.Style = [fsBold]
        ParentFont = False
      end
      object Label2: TLabel
        Left = 24
        Top = 116
        Width = 153
        Height = 16
        Caption = #26412#26426#30331#24405#29992#25143#23494#30721#65306
        Font.Charset = GB2312_CHARSET
        Font.Color = clBlue
        Font.Height = -16
        Font.Name = #23435#20307
        Font.Style = [fsBold]
        ParentFont = False
      end
      object Label3: TLabel
        Left = 24
        Top = 160
        Width = 708
        Height = 16
        Caption = #37325#35201#65306#26412#26426#24517#39035#24320#21551#25991#20214#21644#25171#21360#26426#20849#20139#65292#36828#31243#26426#22120'('#23433#35013#25968#25454#24211#30340#26426#22120')'#24517#39035#20851#38381'360'#31561#23433#20840#36719#20214
        Font.Charset = GB2312_CHARSET
        Font.Color = clRed
        Font.Height = -16
        Font.Name = #23435#20307
        Font.Style = [fsBold]
        ParentFont = False
      end
      object btnRestoreDatabase: TButton
        Left = 24
        Top = 24
        Width = 129
        Height = 41
        Caption = #36824#21407#25968#25454#24211
        TabOrder = 0
        OnClick = btnRestoreDatabaseClick
      end
      object edtLoginPass2: TEdit
        Left = 174
        Top = 111
        Width = 243
        Height = 29
        Font.Charset = SYMBOL_CHARSET
        Font.Color = clBlack
        Font.Height = -19
        Font.Name = 'Wingdings'
        Font.Style = [fsBold]
        ParentFont = False
        PasswordChar = 'l'
        TabOrder = 1
      end
      object edtLoginName2: TEdit
        Left = 174
        Top = 81
        Width = 243
        Height = 24
        Enabled = False
        Font.Charset = GB2312_CHARSET
        Font.Color = clBlack
        Font.Height = -16
        Font.Name = #23435#20307
        Font.Style = [fsBold]
        ParentFont = False
        TabOrder = 2
      end
    end
    object ts7: TTabSheet
      Caption = #30331#24405
      ImageIndex = 6
      ExplicitLeft = 0
      ExplicitTop = 0
      ExplicitWidth = 0
      ExplicitHeight = 0
      object grpEffectLogin: TGroupBox
        Left = 11
        Top = 11
        Width = 782
        Height = 366
        Caption = #30331#24405#34920':'
        Font.Charset = GB2312_CHARSET
        Font.Color = clWindowText
        Font.Height = -15
        Font.Name = #23435#20307
        Font.Style = []
        ParentFont = False
        TabOrder = 0
        object lblLoginTable: TLabel
          Left = 16
          Top = 32
          Width = 75
          Height = 15
          Caption = #30331#24405#34920#21517#65306
        end
        object lbl1: TLabel
          Left = 16
          Top = 63
          Width = 105
          Height = 15
          Caption = #30331#24405#21517#31216#23383#27573#65306
        end
        object lbl2: TLabel
          Left = 16
          Top = 94
          Width = 105
          Height = 15
          Caption = #30331#24405#23494#30721#23383#27573#65306
        end
        object cbbLoginTable: TComboBox
          Left = 121
          Top = 29
          Width = 272
          Height = 23
          Style = csDropDownList
          DropDownCount = 20
          TabOrder = 0
          OnChange = cbbLoginTableChange
        end
        object cbbLoginName: TComboBox
          Left = 121
          Top = 58
          Width = 272
          Height = 23
          Style = csDropDownList
          DropDownCount = 20
          TabOrder = 1
        end
        object cbbLoginPass: TComboBox
          Left = 121
          Top = 88
          Width = 272
          Height = 23
          Style = csDropDownList
          DropDownCount = 20
          TabOrder = 2
        end
        object chkPassword: TCheckBox
          Left = 16
          Top = 125
          Width = 129
          Height = 17
          Caption = #30331#24405#23494#30721#26159#23494#25991
          TabOrder = 3
          OnClick = chkPasswordClick
        end
        object pnlPassword: TPanel
          Left = 35
          Top = 148
          Width = 730
          Height = 205
          BevelOuter = bvNone
          BorderStyle = bsSingle
          Caption = 'pnlPassword'
          Ctl3D = False
          ParentCtl3D = False
          ShowCaption = False
          TabOrder = 4
          Visible = False
          object grpPassword: TGroupBox
            Left = 9
            Top = 6
            Width = 700
            Height = 187
            Caption = 'Dll '#21152#23494#20989#25968
            TabOrder = 0
            object lbl5: TLabel
              Left = 8
              Top = 26
              Width = 69
              Height = 15
              Caption = 'Dll'#25991#20214#65306
              Font.Charset = GB2312_CHARSET
              Font.Color = clRed
              Font.Height = -15
              Font.Name = #23435#20307
              Font.Style = []
              ParentFont = False
            end
            object lbl6: TLabel
              Left = 8
              Top = 56
              Width = 75
              Height = 15
              Caption = #21152#23494#20989#25968#65306
              Font.Charset = GB2312_CHARSET
              Font.Color = clRed
              Font.Height = -15
              Font.Name = #23435#20307
              Font.Style = []
              ParentFont = False
            end
            object lbl7: TLabel
              Left = 12
              Top = 87
              Width = 568
              Height = 90
              Caption = 
                #27880#65306#13#10'  '#20989#25968#24517#39035#26159#19968#20010#36755#20837#21442#25968#65292#19968#20010#36755#20986#21442#25968#12290#13#10'  '#36755#20837#21442#25968#20026#24453#21152#23494#30340#23383#31526#20018#65307#13#10'  '#36755#20986#21442#25968#20026#21152#23494#21518#30340#23383#31526#20018#65307#13#10'  f' +
                'unction FuncName(const strPassword: PAnsiChar):PAnsiChar;  // De' +
                'lphi'#13#10'  char *   FuncName(const char * strPassword);            ' +
                '    // VC'
              Color = clBlue
              Font.Charset = GB2312_CHARSET
              Font.Color = clBlue
              Font.Height = -15
              Font.Name = #23435#20307
              Font.Style = []
              ParentColor = False
              ParentFont = False
            end
            object cbbDllFunc: TComboBox
              Left = 87
              Top = 52
              Width = 478
              Height = 23
              Style = csDropDownList
              TabOrder = 0
            end
            object srchbxDecFuncFile: TSearchBox
              Left = 87
              Top = 24
              Width = 503
              Height = 21
              TabOrder = 1
              OnDblClick = srchbxDecFuncFileInvokeSearch
              OnInvokeSearch = srchbxDecFuncFileInvokeSearch
            end
          end
        end
      end
    end
    object ts8: TTabSheet
      Caption = #23450#26399#28165#38500#25968#25454
      ImageIndex = 7
      ExplicitLeft = 0
      ExplicitTop = 0
      ExplicitWidth = 0
      ExplicitHeight = 0
    end
    object ts9: TTabSheet
      Caption = #23450#26399#22791#20221#25968#25454
      ImageIndex = 8
      ExplicitLeft = 0
      ExplicitTop = 0
      ExplicitWidth = 0
      ExplicitHeight = 0
    end
  end
  object adoCNN: TADOConnection
    LoginPrompt = False
    Left = 652
    Top = 185
  end
end
