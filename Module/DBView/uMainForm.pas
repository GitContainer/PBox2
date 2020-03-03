unit uMainForm;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.StrUtils, System.Variants, System.Classes, System.IniFiles, Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ComCtrls, Vcl.StdCtrls, Vcl.CheckLst, Data.DB, Data.Win.ADODB,
  Data.Win.ADOConEd, XLSReadWriteII5, Xc12Utils5, XLSUtils5, Xc12DataStyleSheet5, db.uCommon;

type
  TfrmDBView = class(TForm)
    btnDBLink: TButton;
    grpAllFields: TGroupBox;
    chklstFields: TCheckListBox;
    grpAllTable: TGroupBox;
    lstTable: TListBox;
    btnDBSearch: TButton;
    grpDispalyData: TGroupBox;
    lvData: TListView;
    btnExportToExcel: TButton;
    ADOQuery1: TADOQuery;
    ADOConnection1: TADOConnection;
    qryTemp: TADOQuery;
    dlgSaveExcel: TSaveDialog;
    lblTip: TLabel;
    procedure btnDBLinkClick(Sender: TObject);
    procedure btnDBSearchClick(Sender: TObject);
    procedure btnExportToExcelClick(Sender: TObject);
    procedure lstTableClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure lvDataData(Sender: TObject; Item: TListItem);
  private
    procedure DispItemData;
    procedure ReadDBFromConfig(ADOCNN: TADOConnection);
    procedure EnumAllTables;
    function GetTableRecordCount(const strTableName: string): Integer;
    { 获取自增长字段名称 }
    function GetAutoAddField(const strTableName: String): String;
  public
    { Public declarations }
  end;

procedure db_ShowDllForm_Plugins(var frm: TFormClass; var ft: TSPFileType; var strParentModuleName, strModuleName, strClassName, strWindowName, strIconFileName: PAnsiChar; const bShow: Boolean = True); stdcall;

implementation

{$R *.dfm}

procedure db_ShowDllForm_Plugins(var frm: TFormClass; var ft: TSPFileType; var strParentModuleName, strModuleName, strClassName, strWindowName, strIconFileName: PAnsiChar; const bShow: Boolean = True); stdcall;
begin
  frm                     := TfrmDBView;
  ft                      := ftDelphiDll;
  strParentModuleName     := '程序员工具';
  strModuleName           := '数据库查看器';
  strIconFileName         := '';
  strClassName            := '';
  strWindowName           := '';
  Application.Handle      := GetMainFormApplication.Handle;
  Application.Icon.Handle := GetMainFormApplication.Icon.Handle;
end;

const
  c_intMaxRecordShow = 10000;

procedure TfrmDBView.btnDBLinkClick(Sender: TObject);
begin
  if not EditConnectionString(ADOConnection1) then
    Exit;

  try
    ADOConnection1.Connected := True;
    EnumAllTables;
  except
  end;
end;

{ 获取自增长字段名称 }
function TfrmDBView.GetAutoAddField(const strTableName: String): String;
begin
  Result := '';
  qryTemp.Close;
  qryTemp.SQL.Text := Format('select colstat, name from syscolumns where id=object_id(%s) and colstat = 1', [QuotedStr(strTableName)]);
  qryTemp.Open;
  if qryTemp.RecordCount > 0 then
    Result := qryTemp.Fields[1].AsString;
end;

function TfrmDBView.GetTableRecordCount(const strTableName: string): Integer;
begin
  with TADOQuery.Create(nil) do
  begin
    Connection := ADOConnection1;
    SQL.Text   := 'select Count(*) from' + strTableName;
    Open;
    Result := Fields[0].AsInteger;
    Free;
  end;
end;

procedure TfrmDBView.lstTableClick(Sender: TObject);
begin
  ADOConnection1.GetFieldNames(lstTable.Items.Strings[lstTable.ItemIndex], chklstFields.Items);
  chklstFields.CheckAll(cbChecked);
  if GetAutoAddField(lstTable.Items.Strings[lstTable.ItemIndex]) = '' then
  begin
    lblTip.Visible := GetTableRecordCount(lstTable.Items.Strings[lstTable.ItemIndex]) >= c_intMaxRecordShow;
  end;
  btnDBSearch.Enabled      := True;
  btnExportToExcel.Enabled := False;
end;

procedure TfrmDBView.lvDataData(Sender: TObject; Item: TListItem);
var
  I: Integer;
begin
  ADOQuery1.RecNo := Item.Index + 1;
  Item.Caption    := ADOQuery1.Fields[1].AsString;
  for I           := 2 to ADOQuery1.Fields.Count - 1 do
  begin
    if ADOQuery1.Fields[I].DataType = ftBlob then
      Item.SubItems.Add('')
    else
      Item.SubItems.Add(ADOQuery1.Fields[I].AsString);
  end;
end;

procedure TfrmDBView.btnDBSearchClick(Sender: TObject);
var
  strTableName: String;
  strID       : String;
  strFields   : String;
  I           : Integer;
begin
  lvData.Items.Clear;
  lvData.Columns.Clear;
  ADOQuery1.Close;

  for I := 0 to chklstFields.Count - 1 do
  begin
    if chklstFields.Checked[I] then
      strFields := strFields + ',' + chklstFields.Items[I];
  end;
  strFields := RightStr(strFields, Length(strFields) - 1);

  strTableName := lstTable.Items.Strings[lstTable.ItemIndex];
  strID        := GetAutoAddField(strTableName);
  if strID = '' then
  begin
    if lblTip.Visible then
      ADOQuery1.SQL.Text := Format('select Top %d %s from %s', [c_intMaxRecordShow, strFields, strTableName])
    else
      ADOQuery1.SQL.Text := 'select ' + strFields + ' from ' + strTableName;
  end
  else
  begin
    ADOQuery1.SQL.Text := 'select ROW_NUMBER() over(order by ' + strID + ') as RowNum, ' + strFields + ' from ' + strTableName;
  end;

  ADOQuery1.Open;
  DispItemData;
  btnExportToExcel.Enabled := ADOQuery1.RecordCount > 0;
end;

procedure TfrmDBView.DispItemData;
const
  c_strFieldChineseName =                                                                                   //
    ' SELECT c.[name] AS 字段名, cast(ep.[value] as varchar(100)) AS [字段说明] FROM sys.tables AS t' +            //
    ' INNER JOIN sys.columns AS c ON t.object_id = c.object_id' +                                           //
    ' LEFT JOIN sys.extended_properties AS ep ON ep.major_id = c.object_id AND ep.minor_id = c.column_id' + //
    ' WHERE ep.class = 1 AND t.name=%s';
var
  I              : Integer;
  strTableName   : String;
  strFieldDisplay: String;
begin
  lvData.Items.Clear;
  lvData.Columns.Clear;
  strTableName := lstTable.Items.Strings[lstTable.ItemIndex];

  qryTemp.Close;
  qryTemp.SQL.Clear;
  qryTemp.SQL.Text := Format(c_strFieldChineseName, [QuotedStr(strTableName)]);
  qryTemp.Open;

  for I := 1 to ADOQuery1.Fields.Count - 1 do
  begin
    if qryTemp.Locate('字段名', ADOQuery1.Fields[I].FieldName, []) then
    begin
      strFieldDisplay := qryTemp.Fields[1].AsString;
      if Trim(strFieldDisplay) = '' then
        strFieldDisplay := ADOQuery1.Fields[I].FieldName;
    end
    else
    begin
      strFieldDisplay := ADOQuery1.Fields[I].FieldName;
    end;

    with lvData.Columns.Add do
    begin
      Caption := strFieldDisplay;
      Width   := 160;
    end;
  end;
  qryTemp.Close;

  ADOQuery1.ControlsDisabled;
  lvData.Items.Count := ADOQuery1.RecordCount;
end;

procedure TfrmDBView.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Action := caFree;
end;

procedure TfrmDBView.FormCreate(Sender: TObject);
begin
  ReadDBFromConfig(ADOConnection1);
end;

procedure TfrmDBView.ReadDBFromConfig(ADOCNN: TADOConnection);
var
  strIniFileName: String;
  strEncLink    : String;
  strSQLLink    : String;
begin
  strIniFileName := string(GetConfigFileName);
  if Trim(strIniFileName) = '' then
    Exit;

  with TIniFile.Create(strIniFileName) do
  begin
    strEncLink := ReadString('DB', 'Name', '');
    if Trim(strEncLink) <> '' then
    begin
      strSQLLink       := DecryptString(strEncLink, c_strAESKey);
      ADOCNN.Connected := False;
      try
        ADOCNN.ConnectionString := strSQLLink;
        ADOCNN.KeepConnection   := True;
        ADOCNN.LoginPrompt      := False;
        ADOCNN.Connected        := True;
        btnDBLink.Enabled       := False;
        EnumAllTables;
      except
        btnDBLink.Enabled := True;
        ADOCNN.Connected  := False;
      end;
    end;
    Free;
  end;
end;

procedure TfrmDBView.EnumAllTables;
var
  lstTables: TStringList;
begin
  lstTable.Clear;
  lstTables := TStringList.Create;
  try
    ADOConnection1.GetTableNames(lstTables);
    lstTable.Items.AddStrings(lstTables);
  finally
    lstTables.Free;
  end;
end;

procedure TfrmDBView.btnExportToExcelClick(Sender: TObject);
var
  XLS    : TXLSReadWriteII5;
  I, J, K: Integer;
begin
  if not dlgSaveExcel.Execute then
    Exit;

  btnDBLink.Enabled        := False;
  btnDBSearch.Enabled      := False;
  btnExportToExcel.Enabled := False;
  Application.ProcessMessages;
  XLS := TXLSReadWriteII5.Create(nil);
  try
    XLS.Filename := dlgSaveExcel.Filename + '.xlsx';
    for I        := 1 to lvData.Columns.Count do
    begin
      for J := 1 to ADOQuery1.RecordCount + 1 do
      begin
        XLS.Sheets[0].Range.Items[I, J, I, J].BorderOutlineStyle := cbsThin;
        XLS.Sheets[0].Range.Items[I, J, I, J].BorderOutlineColor := 0;
      end;
    end;

    for I := 1 to lvData.Columns.Count do
    begin
      Application.ProcessMessages;
      XLS.Sheets[0].AsString[I, 1]                  := lvData.Column[I - 1].Caption;
      XLS.Sheets[0].Columns[I].Width                := 4000;
      XLS.Sheets[0].Cell[I, 1].FontColor            := clWhite;
      XLS.Sheets[0].Cell[I, 1].FontStyle            := XLS.Sheets[0].Cell[I, 1].FontStyle + [xfsBold];
      XLS.Sheets[0].Cell[I, 1].FillPatternForeColor := xcBlue;
      XLS.Sheets[0].Cell[I, 1].HorizAlignment       := chaCenter;
      XLS.Sheets[0].Cell[I, 1].VertAlignment        := cvaCenter;
    end;

    K := 2;
    ADOQuery1.First;
    while not ADOQuery1.Eof do
    begin
      J     := 1;
      for I := 1 to lvData.Columns.Count do
      begin
        if I = 1 then
          XLS.Sheets[0].AsInteger[J, K] := ADOQuery1.Fields[I].AsInteger
        else
          XLS.Sheets[0].AsString[J, K] := ADOQuery1.Fields[I].AsString;

        XLS.Sheets[0].Cell[J, K].HorizAlignment := chaCenter;
        XLS.Sheets[0].Cell[J, K].VertAlignment  := cvaCenter;
        Inc(J);
      end;
      Inc(K);
      btnExportToExcel.Caption := Format('正在导出：%d', [K - 2]);
      ADOQuery1.Next;
    end;

    XLS.Write;
  finally
    XLS.Free;
    btnExportToExcel.Caption := '数据导出到 Excel';
    btnDBLink.Enabled        := True;
    btnDBSearch.Enabled      := True;
    btnExportToExcel.Enabled := True;
  end;
end;

end.
