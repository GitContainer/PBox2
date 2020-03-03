unit FormPivotTable;

interface

{-
********************************************************************************
******* XLSSpreadSheet V2.00                                             *******
*******                                                                  *******
******* Copyright(C) 2006,2013 Lars Arvidsson, Axolot Data               *******
*******                                                                  *******
******* email: components@axolot.com                                     *******
******* URL:   http://www.axolot.com                                     *******
********************************************************************************
** Users of the XLSSpreadSheet component must accept the following            **
** disclaimer of warranty:                                                    **
**                                                                            **
** XLSSpreadSheet is supplied as is. The author disclaims all warranties,     **
** expressed or implied, including, without limitation, the warranties of     **
** merchantability and of fitness for any purpose. The author assumes no      **
** liability for damages, direct or consequential, which may result from the  **
** use of XLSSpreadSheet.                                                     **
********************************************************************************
}

{$B-}
{$H+}
{$R-}
{$I AxCompilers.inc}

uses
  Windows, Messages, SysUtils, Variants, Classes, vcl.Graphics,
  vcl.Controls, vcl.Forms, vcl.Dialogs, XLSPivotTables5, vcl.StdCtrls, IniFiles,
  vcl.CheckLst, vcl.ActnList, vcl.Menus, FormPivotTableField, FormSelectValues,
{$ifdef DELPHI_XE3_OR_LATER}
  System.Diagnostics,
{$endif}
  FormPivotTableOptions, System.Actions;

type
  TfrmPivotTable = class(TForm)
    clbFields: TCheckListBox;
    Label1: TLabel;
    lbFilters: TListBox;
    lbColumn: TListBox;
    lbValues: TListBox;
    lbRow: TListBox;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    btnClose: TButton;
    btnUpdate: TButton;
    Label6: TLabel;
    popLabels: TPopupMenu;
    ActionList: TActionList;
    acRemove: TAction;
    acSettings: TAction;
    Remove1: TMenuItem;
    Fieldsettings1: TMenuItem;
    acMoveUp: TAction;
    acMoveDown: TAction;
    acMoveFirst: TAction;
    acMoveLast: TAction;
    acToReportFilter: TAction;
    acToColumnLabels: TAction;
    acToRowLables: TAction;
    acToValues: TAction;
    N1: TMenuItem;
    Moveup1: TMenuItem;
    MoveDown1: TMenuItem;
    MovetoBeginning1: TMenuItem;
    MovetoEnd1: TMenuItem;
    N2: TMenuItem;
    N3: TMenuItem;
    MovetoReportFilter1: TMenuItem;
    MovetoColumnLabels1: TMenuItem;
    acMovetoRowLables1: TMenuItem;
    MovetoValues1: TMenuItem;
    popFields: TPopupMenu;
    acAddReportFilter: TAction;
    acAddColumnLabels: TAction;
    acAddRowLables: TAction;
    acAddValues: TAction;
    AddtoReportFilter1: TMenuItem;
    AddtoColumnLabels1: TMenuItem;
    AddtoRowLables1: TMenuItem;
    AddtoValues1: TMenuItem;
    btnOptions: TButton;
    procedure ListBoxDragOver(Sender, Source: TObject; X, Y: Integer;
      State: TDragState; var Accept: Boolean);
    procedure LsitBoxMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure ListBoxEndDrag(Sender, Target: TObject; X, Y: Integer);
    procedure ListBoxDragDrop(Sender, Source: TObject; X, Y: Integer);
    procedure clbFieldsClickCheck(Sender: TObject);
    procedure btnUpdateClick(Sender: TObject);
    procedure ListBoxDrawItem(Control: TWinControl; Index: Integer; Rect: TRect;
      State: TOwnerDrawState);
    procedure acMoveUpExecute(Sender: TObject);
    procedure acMoveDownExecute(Sender: TObject);
    procedure acMoveFirstExecute(Sender: TObject);
    procedure acMoveLastExecute(Sender: TObject);
    procedure acToReportFilterExecute(Sender: TObject);
    procedure acToColumnLabelsExecute(Sender: TObject);
    procedure acToRowLablesExecute(Sender: TObject);
    procedure acToValuesExecute(Sender: TObject);
    procedure acRemoveExecute(Sender: TObject);
    procedure acAddReportFilterExecute(Sender: TObject);
    procedure acAddColumnLabelsExecute(Sender: TObject);
    procedure acAddRowLablesExecute(Sender: TObject);
    procedure acAddValuesExecute(Sender: TObject);
    procedure acSettingsExecute(Sender: TObject);
    procedure btnCloseClick(Sender: TObject);
    procedure btnOptionsClick(Sender: TObject);
  private
    FPivTable    : TXLSPivotTable;
    FCurrLB      : TCustomListBox;
    FCurrItem    : integer;
    FUpdateEvent : TNotifyEvent;
    FIniFile     : TCustomIniFile;
    FAutoUpdate  : boolean;
    FAutoUpdateMS: Int64;

    procedure CheckByName(AName: string; AChecked: boolean = True);
    procedure InsertFields(AListBox: TCustomListBox; AFields: TXLSPivotFieldsDest);
    procedure DoTable;
    procedure RemoveByName(AName: string); overload;
    procedure RemoveByName(AListBox: TCustomListBox; AName: string); overload;

    procedure UpdateFields(AListBox: TCustomListBox; AFields: TXLSPivotFieldsDest);
    procedure UpdateAllFields;

    procedure EnableActions;

    function  CurrName: string;
    function  CurrField: TXLSPivotField;
    procedure DoUpdate;
    procedure DoExecute(APivTable: TXLSPivotTable; AUpdateEvent: TNotifyEvent);
  public
    procedure Execute(APivTable: TXLSPivotTable; ALeft, ATop: integer; AUpdateEvent: TNotifyEvent); overload;
    procedure Execute(APivTable: TXLSPivotTable; AIniFile: TCustomIniFile; AUpdateEvent: TNotifyEvent); overload;

    property PivotTable  : TXLSPivotTable read FPivTable;
    // If the time to create the piviot table is less than AutoUpdateMS (milliseconds),
    // auto update is on.
    property AutoUpdateMS: Int64 read FAutoUpdateMS write FAutoUpdateMS;
    property OnUpdate    : TNotifyEvent read FUpdateEvent write FUpdateEvent;
  end;

implementation

{$R *.dfm}

{ TfrmPivotTable }

procedure TfrmPivotTable.LsitBoxMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
  P  : TPoint;
  P2 : TPoint;
  R  : TRect;
  Frm: TfrmSelectValues;
begin
  if Button <> mbLeft then
    Exit;

  P.X := X;
  P.Y := Y;

  FCurrLB := TCustomListBox(Sender);
  FCurrItem := FCurrLB.ItemAtPos(P,True);

  EnableActions;

  if FCurrItem >= 0 then begin
    R := FCurrLB.ItemRect(FCurrItem);
    if X > R.Right - 25 then begin
      R := FCurrLB.ItemRect(FCurrItem);
      P2.X := FCurrLB.Left;
      P2.Y := FCurrLB.Top;
      P := ClientToScreen(P2);
      if Sender is TCheckListBox then begin
        if CurrField.CacheField.SharedItems.Values.ValueList.Count <= 0 then
          FPivTable.TableDef.Cache.CacheValues(FCurrItem);

        Frm := TfrmSelectValues.Create(Self);
        Frm.Left := P.X - Frm.Width;
        Frm.Top := P.Y + Y;
        if Frm.Execute(CurrField.CacheField.SharedItems.Values.ValueList,CurrField.Filter) and FAutoUpdate then
          DoUpdate;
      end
      else
        popLabels.Popup(P.X + R.Left,P.Y + R.Bottom + 2);
    end
    else
      TListBox(Sender).BeginDrag(False);
  end;
end;

procedure TfrmPivotTable.RemoveByName(AName: string);
begin
  RemoveByName(lbFilters,AName);
  RemoveByName(lbColumn,AName);
  RemoveByName(lbRow,AName);
  RemoveByName(lbValues,AName);
end;

procedure TfrmPivotTable.RemoveByName(AListBox: TCustomListBox; AName: string);
var
  i: integer;
begin
  i := AListBox.Items.IndexOf(AName);
  if i >= 0 then begin
    case AListBox.Tag of
      2: FPivTable.ReportFilter.Remove(CurrName);
      3: FPivTable.ColumnLabels.Remove(CurrName);
      4: FPivTable.RowLabels.Remove(CurrName);
      5: FPivTable.DataValues.Remove(CurrName);
    end;

    AListBox.Items.Delete(i);
  end;
end;

procedure TfrmPivotTable.UpdateAllFields;
begin
  FPivTable.ClearReport;

  UpdateFields(lbFilters,FPivTable.ReportFilter);
  UpdateFields(lbColumn,FPivTable.ColumnLabels);
  UpdateFields(lbRow,FPivTable.RowLabels);
  UpdateFields(lbValues,FPivTable.DataValues);
end;

procedure TfrmPivotTable.UpdateFields(AListBox: TCustomListBox; AFields: TXLSPivotFieldsDest);
var
  i: integer;
begin
  for i := 0 to AListBox.Items.Count - 1 do
    AFields.Add(FPivTable.Fields.Find(AListBox.Items[i]));
end;

procedure TfrmPivotTable.ListBoxEndDrag(Sender, Target: TObject; X, Y: Integer);
begin
  FCurrLB := Nil;
  FCurrItem := -1;
end;

procedure TfrmPivotTable.acAddColumnLabelsExecute(Sender: TObject);
var
  N: string;
begin
  N := clbFields.Items[clbFields.ItemIndex];

  RemoveByName(N);
  CheckByName(N);

  lbColumn.Items.AddObject(N,clbFields.Items.Objects[clbFields.ItemIndex]);

  if FAutoUpdate then
    DoUpdate;
end;

procedure TfrmPivotTable.acAddReportFilterExecute(Sender: TObject);
var
  N: string;
begin
  N := clbFields.Items[clbFields.ItemIndex];

  RemoveByName(N);
  CheckByName(N);

  lbFilters.Items.AddObject(N,clbFields.Items.Objects[clbFields.ItemIndex]);

  if FAutoUpdate then
    DoUpdate;
end;

procedure TfrmPivotTable.acAddRowLablesExecute(Sender: TObject);
var
  N: string;
begin
  N := clbFields.Items[clbFields.ItemIndex];

  RemoveByName(N);
  CheckByName(N);

  lbRow.Items.AddObject(N,clbFields.Items.Objects[clbFields.ItemIndex]);

  if FAutoUpdate then
    DoUpdate;
end;

procedure TfrmPivotTable.acAddValuesExecute(Sender: TObject);
var
  N: string;
begin
  N := clbFields.Items[clbFields.ItemIndex];

  RemoveByName(N);
  CheckByName(N);

  lbValues.Items.AddObject(N,clbFields.Items.Objects[clbFields.ItemIndex]);

  if FAutoUpdate then
    DoUpdate;
end;

procedure TfrmPivotTable.acMoveDownExecute(Sender: TObject);
begin
  FCurrLB.Items.Exchange(FCurrItem,FCurrItem + 1);

  if FAutoUpdate then
    DoUpdate;
end;

procedure TfrmPivotTable.acMoveFirstExecute(Sender: TObject);
begin
  FCurrLB.Items.Exchange(FCurrItem,0);

  if FAutoUpdate then
    DoUpdate;
end;

procedure TfrmPivotTable.acMoveLastExecute(Sender: TObject);
begin
  FCurrLB.Items.Exchange(FCurrItem,FCurrLB.Items.Count - 1);

  if FAutoUpdate then
    DoUpdate;
end;

procedure TfrmPivotTable.acMoveUpExecute(Sender: TObject);
begin
  FCurrLB.Items.Exchange(FCurrItem,FCurrItem - 1);

  if FAutoUpdate then
    DoUpdate;
end;

procedure TfrmPivotTable.acRemoveExecute(Sender: TObject);
begin
  CheckByName(CurrName,False);

  RemoveByName(CurrName);

  if FAutoUpdate then
    DoUpdate;
end;

procedure TfrmPivotTable.acSettingsExecute(Sender: TObject);
begin
  if TfrmPivotTableField.Create(Self).Execute(FPivTable.Fields.Find(CurrName)) then begin
    DoTable;

    if FAutoUpdate then
      DoUpdate;
  end;
end;

procedure TfrmPivotTable.acToColumnLabelsExecute(Sender: TObject);
begin
  lbColumn.Items.AddObject(CurrName,CurrField);
  FCurrLB.Items.Delete(FCurrItem);

  if FAutoUpdate then
    DoUpdate;
end;

procedure TfrmPivotTable.acToReportFilterExecute(Sender: TObject);
begin
  lbFilters.Items.AddObject(CurrName,CurrField);
  FCurrLB.Items.Delete(FCurrItem);

  if FAutoUpdate then
    DoUpdate;
end;

procedure TfrmPivotTable.acToRowLablesExecute(Sender: TObject);
begin
  lbRow.Items.AddObject(CurrName,CurrField);
  FCurrLB.Items.Delete(FCurrItem);

  if FAutoUpdate then
    DoUpdate;
end;

procedure TfrmPivotTable.acToValuesExecute(Sender: TObject);
begin
  lbValues.Items.AddObject(CurrName,CurrField);
  FCurrLB.Items.Delete(FCurrItem);

  if FAutoUpdate then
    DoUpdate;
end;

procedure TfrmPivotTable.btnCloseClick(Sender: TObject);
begin
  if FIniFile <> Nil then begin
    FIniFile.WriteInteger('PivDialog','Left',Left);
    FIniFile.WriteInteger('PivDialog','Top',Top);
  end;

  Close;
end;

procedure TfrmPivotTable.btnOptionsClick(Sender: TObject);
begin
  if TfrmPivotTableOptions.Create(Self).Execute(FPivTable) then
    DoUpdate;
end;

procedure TfrmPivotTable.btnUpdateClick(Sender: TObject);
begin
  DoUpdate;
end;

procedure TfrmPivotTable.CheckByName(AName: string; AChecked: boolean = True);
var
  i: integer;
begin
  i := clbFields.Items.IndexOf(AName);

  if i >= 0 then
    clbFields.Checked[i] := AChecked;
end;

procedure TfrmPivotTable.clbFieldsClickCheck(Sender: TObject);
var
  S: string;
begin
  if clbFields.ItemIndex >= 0 then begin
    if clbFields.Checked[clbFields.ItemIndex] then begin
      lbRow.Items.AddObject(clbFields.Items[clbFields.ItemIndex],clbFields.Items.Objects[clbFields.ItemIndex]);
    end
    else if FCurrLB <> Nil then begin
      S := FCurrLB.Items[FCurrItem];

      RemoveByName(S);
    end;

    if FAutoUpdate then
      DoUpdate;
  end;
end;

function TfrmPivotTable.CurrField: TXLSPivotField;
begin
  Result := FPivTable.Fields.Find(CurrName);
end;

function TfrmPivotTable.CurrName: string;
begin
  Result := FCurrLB.Items[FCurrItem];
end;

procedure TfrmPivotTable.DoExecute(APivTable: TXLSPivotTable; AUpdateEvent: TNotifyEvent);
begin
  FAutoUpdateMS := 500;
  FAutoUpdate := True;

  FPivTable := APivTable;
  FUpdateEvent := AUpdateEvent;

  DoTable;

  if FPivTable.Dirty then
    DoUpdate;

  Show;
end;

procedure TfrmPivotTable.DoTable;
var
  i: integer;
begin
  clbFields.Clear;

  for i := 0 to FPivTable.Fields.Count - 1 do
    clbFields.Items.AddObject(FPivTable.Fields[i].DisplayName,FPivTable.Fields[i]);

  InsertFields(lbFilters,FPivTable.ReportFilter);
  InsertFields(lbColumn,FPivTable.ColumnLabels);
  InsertFields(lbRow,FPivTable.RowLabels);
  InsertFields(lbValues,FPivTable.DataValues);

  clbFields.Repaint;
end;

procedure TfrmPivotTable.DoUpdate;
{$ifdef DELPHI_XE3_OR_LATER}
var
  SW: TStopWatch;
{$endif}
begin
{$ifdef DELPHI_XE3_OR_LATER}
  SW := TStopWatch.Create;
  SW.Start;
{$endif}

  UpdateAllFields;

  FPivTable.Make;

  if Assigned(FUpdateEvent) then
    FUpdateEvent(Self);

{$ifdef DELPHI_XE3_OR_LATER}
  SW.Stop;

  FAutoUpdate := SW.ElapsedMilliseconds < FAutoUpdateMS;
{$else}
  FAutoUpdate := False;
{$endif}
end;

procedure TfrmPivotTable.EnableActions;
begin
  acMoveUp.Enabled := (FCurrItem >= 1);
  acMoveDown.Enabled := (FCurrItem >= 0) and (FCurrItem < (FCurrLB.Items.Count - 1));
  acMoveFirst.Enabled := (FCurrItem > 0);
  acMoveLast.Enabled := (FCurrItem >= 0) and (FCurrItem < (FCurrLB.Items.Count - 1));

  acToReportFilter.Enabled := (FcurrLB <> lbFilters);
  acToColumnLabels.Enabled := (FcurrLB <> lbColumn);
  acToRowLables.Enabled := (FcurrLB <> lbRow);
  acToValues.Enabled := (FcurrLB <> lbValues);
end;

procedure TfrmPivotTable.Execute(APivTable: TXLSPivotTable; AIniFile: TCustomIniFile; AUpdateEvent: TNotifyEvent);
var
  L,T   : integer;
begin
  FIniFile := AIniFile;

  L := FIniFile.ReadInteger('PivDialog','Left',-1);
  T := FIniFile.ReadInteger('PivDialog','Top',-1);
  if (L >= 0) and (T >= 0) then begin
    Left := L;
    Top := T;
  end;

  DoExecute(APivTable,AUpdateEvent);
end;

procedure TfrmPivotTable.Execute(APivTable: TXLSPivotTable; ALeft, ATop: integer; AUpdateEvent: TNotifyEvent);
begin
  if ALeft < 0 then
    Left := -ALeft - Width
  else
    Left := ALeft;
  Top := ATop;

  DoExecute(APivTable,AUpdateEvent);
end;

procedure TfrmPivotTable.InsertFields(AListBox: TCustomListBox; AFields: TXLSPivotFieldsDest);
var
  i: integer;
begin
  AListBox.Clear;

  for i := 0 to AFields.Count - 1 do begin
    AListBox.Items.AddObject(AFields[i].DisplayName,AFields[i]);
    CheckByName(AFields[i].DisplayName);
  end;
end;

procedure TfrmPivotTable.ListBoxDrawItem(Control: TWinControl; Index: Integer; Rect: TRect; State: TOwnerDrawState);
var
  H  : integer;
  S  : string;
  LB : TListBox;
  Pts: array[0..3] of TPoint;
  Fld: TXLSPivotField;
begin
  Fld := Nil;

  LB := TListBox(Control);

  if odSelected in State then begin
    LB.Canvas.Pen.Color := $004598A5;
    LB.Canvas.Brush.Color := $0080EEFF;
    LB.Canvas.Rectangle(Rect);
  end
  else begin
    LB.Canvas.Brush.Color := $00B9F2FB;
    LB.Canvas.FillRect(Rect);
  end;

  if (Control is TCheckListBox) then begin
    if TCheckListBox(Control).Checked[Index] then
      LB.Canvas.Font.Style := [fsBold];
    Fld := TXLSPivotField(LB.Items.Objects[Index]);
  end
  else
    LB.Canvas.Font.Style := [];

  LB.Canvas.Font.Color := clBlack;

  if LB.Tag = 5 then
    S := TXLSPivotField(LB.Items.Objects[Index]).FuncName + ' of ' + LB.Items[Index]
  else
    S := LB.Items[Index];

  LB.Canvas.TextOut(Rect.Left + 2,Rect.Top + 1,S);

  if odFocused In State then begin
    LB.Canvas.Brush.Color := LB.Color;
    LB.Canvas.DrawFocusRect(Rect);
  end;

  Pts[0].X := Rect.Right - 14;
  Pts[0].Y := Rect.Top   +  4;

  Pts[1].X := Rect.Right -  6;
  Pts[1].Y := Rect.Top   +  4;

  Pts[2].X := Rect.Right - 10;
  Pts[2].Y := Rect.Bottom - 7;

  Pts[3].X := Rect.Right - 14;
  Pts[3].Y := Rect.Top   +  4;

  LB.Canvas.Pen.Color := clBlack;
  LB.Canvas.Brush.Color := clBlack;
  LB.Canvas.Brush.Style := bsSolid;

  LB.Canvas.Polygon(Pts);

  if (Fld <> Nil) and (Fld.Filter.Count > 0) then begin
    LB.Canvas.Pen.Color := $F0B000;
    LB.Canvas.Brush.Color := $F0B000;
    LB.Canvas.Brush.Style := bsSolid;

    H := (Rect.Bottom - Rect.Top)- 8;
    LB.Canvas.Ellipse(Rect.Right - 20 - H,Rect.Top + 4,Rect.Right - 20,Rect.Bottom - 4);
  end;
end;

procedure TfrmPivotTable.ListBoxDragDrop(Sender, Source: TObject; X, Y: Integer);
var
  i       : integer;
  S       : string;
  O       : TObject;
  P       : TPoint;
  LBSender: TCustomListBox;
  LBSource: TCustomListBox;
begin
  if (FCurrLB <> Nil) and (FCurrItem >= 0) then begin
    LBSender := TCustomListBox(Sender);
    LBSource := TCustomListBox(Source);

    S := LBSource.Items[FCurrItem];
    O := LBSource.Items.Objects[FCurrItem];

    if LBSender is TCheckListBox then begin
      i := TCheckListBox(LBSender).Items.IndexOf(S);
      if i >= 0 then begin
        TCheckListBox(LBSender).Checked[i] := False;
        LBSource.Items.Delete(FCurrItem);
      end;
    end
    else if LBSender = LBSource then begin
      P.X := X;
      P.Y := Y;
      i := FCurrLB.ItemAtPos(P,True);
      if (i >= 0) and (i <> FCurrItem) then
        LBSender.Items.Exchange(i,FCurrItem);
    end
    else if LBSender.Items.IndexOf(S) < 0 then begin
      if LBSource is TCheckListBox then
        TCheckListBox(LBSource).Checked[FCurrItem] := True
      else
        LBSource.Items.Delete(FCurrItem);
      LBSender.Items.AddObject(S,O);
    end;
  end;

  if FAutoUpdate then
    DoUpdate;
end;

procedure TfrmPivotTable.ListBoxDragOver(Sender, Source: TObject; X, Y: Integer; State: TDragState; var Accept: Boolean);
begin
  Accept := (Sender is TCustomListBox) and not (Sender is TCheckListBox);
end;

end.
