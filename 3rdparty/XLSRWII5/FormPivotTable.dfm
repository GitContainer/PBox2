object frmPivotTable: TfrmPivotTable
  Left = 711
  Top = 121
  BorderStyle = bsDialog
  Caption = 'Pivot Table'
  ClientHeight = 675
  ClientWidth = 275
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  FormStyle = fsStayOnTop
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 8
    Top = 8
    Width = 148
    Height = 13
    Caption = 'Choose fileds to add to report:'
  end
  object Label2: TLabel
    Left = 8
    Top = 301
    Width = 58
    Height = 13
    Caption = 'Report filter'
  end
  object Label3: TLabel
    Left = 148
    Top = 301
    Width = 65
    Height = 13
    Caption = 'Column labels'
  end
  object Label4: TLabel
    Left = 8
    Top = 471
    Width = 51
    Height = 13
    Caption = 'Row labels'
  end
  object Label5: TLabel
    Left = 148
    Top = 471
    Width = 31
    Height = 13
    Caption = 'Values'
  end
  object Label6: TLabel
    Left = 8
    Top = 278
    Width = 161
    Height = 13
    Caption = 'Drag fields between areas below:'
  end
  object clbFields: TCheckListBox
    Tag = 1
    Left = 4
    Top = 24
    Width = 265
    Height = 249
    OnClickCheck = clbFieldsClickCheck
    Style = lbOwnerDrawFixed
    TabOrder = 0
    OnDragDrop = ListBoxDragDrop
    OnDragOver = ListBoxDragOver
    OnDrawItem = ListBoxDrawItem
    OnEndDrag = ListBoxEndDrag
    OnMouseDown = LsitBoxMouseDown
  end
  object lbFilters: TListBox
    Tag = 2
    Left = 4
    Top = 316
    Width = 125
    Height = 150
    Style = lbOwnerDrawFixed
    TabOrder = 1
    OnDragDrop = ListBoxDragDrop
    OnDragOver = ListBoxDragOver
    OnDrawItem = ListBoxDrawItem
    OnEndDrag = ListBoxEndDrag
    OnMouseDown = LsitBoxMouseDown
  end
  object lbColumn: TListBox
    Tag = 3
    Left = 144
    Top = 316
    Width = 123
    Height = 150
    Style = lbOwnerDrawFixed
    TabOrder = 2
    OnDragDrop = ListBoxDragDrop
    OnDragOver = ListBoxDragOver
    OnDrawItem = ListBoxDrawItem
    OnEndDrag = ListBoxEndDrag
    OnMouseDown = LsitBoxMouseDown
  end
  object lbValues: TListBox
    Tag = 5
    Left = 144
    Top = 486
    Width = 123
    Height = 150
    Style = lbOwnerDrawFixed
    TabOrder = 4
    OnDragDrop = ListBoxDragDrop
    OnDragOver = ListBoxDragOver
    OnDrawItem = ListBoxDrawItem
    OnEndDrag = ListBoxEndDrag
    OnMouseDown = LsitBoxMouseDown
  end
  object lbRow: TListBox
    Tag = 4
    Left = 4
    Top = 486
    Width = 125
    Height = 150
    Style = lbOwnerDrawFixed
    TabOrder = 3
    OnDragDrop = ListBoxDragDrop
    OnDragOver = ListBoxDragOver
    OnDrawItem = ListBoxDrawItem
    OnEndDrag = ListBoxEndDrag
    OnMouseDown = LsitBoxMouseDown
  end
  object btnClose: TButton
    Left = 4
    Top = 642
    Width = 75
    Height = 25
    Cancel = True
    Caption = 'Close'
    Default = True
    ModalResult = 2
    TabOrder = 6
    OnClick = btnCloseClick
  end
  object btnUpdate: TButton
    Left = 192
    Top = 642
    Width = 75
    Height = 25
    Caption = 'Update'
    TabOrder = 5
    OnClick = btnUpdateClick
  end
  object btnOptions: TButton
    Left = 96
    Top = 642
    Width = 75
    Height = 25
    Caption = 'Options...'
    TabOrder = 7
    OnClick = btnOptionsClick
  end
  object popLabels: TPopupMenu
    Left = 228
    Top = 88
    object Moveup1: TMenuItem
      Action = acMoveUp
    end
    object MoveDown1: TMenuItem
      Action = acMoveDown
    end
    object MovetoBeginning1: TMenuItem
      Action = acMoveFirst
    end
    object MovetoEnd1: TMenuItem
      Action = acMoveLast
    end
    object N1: TMenuItem
      Caption = '-'
    end
    object MovetoReportFilter1: TMenuItem
      Action = acToReportFilter
    end
    object MovetoColumnLabels1: TMenuItem
      Action = acToColumnLabels
    end
    object acMovetoRowLables1: TMenuItem
      Action = acToRowLables
    end
    object MovetoValues1: TMenuItem
      Action = acToValues
    end
    object N2: TMenuItem
      Caption = '-'
    end
    object Remove1: TMenuItem
      Action = acRemove
    end
    object N3: TMenuItem
      Caption = '-'
    end
    object Fieldsettings1: TMenuItem
      Action = acSettings
    end
  end
  object ActionList: TActionList
    Left = 172
    Top = 40
    object acRemove: TAction
      Caption = 'Remove Field'
      OnExecute = acRemoveExecute
    end
    object acSettings: TAction
      Caption = 'Field Settings...'
      OnExecute = acSettingsExecute
    end
    object acMoveUp: TAction
      Caption = 'Move up'
      OnExecute = acMoveUpExecute
    end
    object acMoveDown: TAction
      Caption = 'Move Down'
      OnExecute = acMoveDownExecute
    end
    object acMoveFirst: TAction
      Caption = 'Move to Beginning'
      OnExecute = acMoveFirstExecute
    end
    object acMoveLast: TAction
      Caption = 'Move to End'
      OnExecute = acMoveLastExecute
    end
    object acToReportFilter: TAction
      Caption = 'Move to Report Filter'
      OnExecute = acToReportFilterExecute
    end
    object acToColumnLabels: TAction
      Caption = 'Move to Column Labels'
      OnExecute = acToColumnLabelsExecute
    end
    object acToRowLables: TAction
      Caption = 'Move to Row Lables'
      OnExecute = acToRowLablesExecute
    end
    object acToValues: TAction
      Caption = 'Move to Values'
      OnExecute = acToValuesExecute
    end
    object acAddReportFilter: TAction
      Caption = 'Add to Report Filter'
      OnExecute = acAddReportFilterExecute
    end
    object acAddColumnLabels: TAction
      Caption = 'Add to Column Labels'
      OnExecute = acAddColumnLabelsExecute
    end
    object acAddRowLables: TAction
      Caption = 'Add to Row Lables'
      OnExecute = acAddRowLablesExecute
    end
    object acAddValues: TAction
      Caption = 'Add to Values'
      OnExecute = acAddValuesExecute
    end
  end
  object popFields: TPopupMenu
    Left = 228
    Top = 40
    object AddtoReportFilter1: TMenuItem
      Action = acAddReportFilter
    end
    object AddtoColumnLabels1: TMenuItem
      Action = acAddColumnLabels
    end
    object AddtoRowLables1: TMenuItem
      Action = acAddRowLables
    end
    object AddtoValues1: TMenuItem
      Action = acAddValues
    end
  end
end
