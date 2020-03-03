unit db.LoginForm;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, System.IniFiles, Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls,
  Data.db, Data.Win.ADODB, db.uCommon, Vcl.ExtCtrls, System.ImageList, Vcl.ImgList;

type
  TfrmLogin = class(TForm)
    lbl1: TLabel;
    lbl2: TLabel;
    edtUserName: TEdit;
    edtUserPass: TEdit;
    btnSave: TButton;
    btnCancel: TButton;
    imgLogo: TImage;
    chkUserName: TCheckBox;
    chkAutoLogin: TCheckBox;
    ilButton: TImageList;
    procedure btnCancelClick(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
    procedure edtUserPassKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure edtUserNameKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure chkAutoLoginClick(Sender: TObject);
    procedure chkUserNameClick(Sender: TObject);
  private
    { Private declarations }
    FbResult: Boolean;
    procedure LoadLoginInfo(var ini: TIniFile);
    procedure SaveLoginInfo(var ini: TIniFile);
  public
    { Public declarations }
  end;

procedure CheckLoginForm;

implementation

{$R *.dfm}

var
  g_ADOCNN       : TADOConnection = nil;
  g_strLoginTable: string         = '';
  g_strLoginName : string         = '';
  g_strLoginPass : String         = '';

procedure CheckLoginForm;
var
  IniFile  : TIniFile;
  strLinkDB: String;
begin
  g_ADOCNN := TADOConnection.Create(nil);
  try
    IniFile         := TIniFile.Create(ChangeFileExt(ParamStr(0), '.ini'));
    strLinkDB       := DecryptString(IniFile.ReadString(c_strIniDBSection, 'Name', ''), c_strAESKey);
    g_strLoginTable := IniFile.ReadString(c_strIniDBSection, 'LoginTable', '');
    g_strLoginName  := IniFile.ReadString(c_strIniDBSection, 'LoginNameField', '');
    g_strLoginPass  := IniFile.ReadString(c_strIniDBSection, 'LoginPassField', '');
    if TryLinkDataBase(strLinkDB, g_ADOCNN) then
    begin
      if (strLinkDB <> '') and (g_strLoginTable <> '') and (g_strLoginName <> '') and (g_strLoginPass <> '') then
      begin
        with TfrmLogin.Create(nil) do
        begin
          FbResult             := False;
          imgLogo.Picture.Icon := Application.Icon;
          Position             := poScreenCenter;
          LoadLoginInfo(IniFile);
          ShowModal;
          if not FbResult then
          begin
            Halt(0);
          end
          else
          begin
            { 登录成功 }
            try
              SaveLoginInfo(IniFile);
              g_strCurrentLoginName := edtUserName.Text;
              UpdateDataBaseScript(IniFile, g_ADOCNN, True);
            except

            end;
          end;
          Free;
        end
      end;
    end;
    IniFile.Free;
  finally
    g_ADOCNN.Free;
  end;
end;

procedure TfrmLogin.LoadLoginInfo(var ini: TIniFile);
begin
  if ini.ReadBool(c_strIniDBSection, 'CheckLoginAuto', False) then
  begin
    edtUserName.Text     := ini.ReadString(c_strIniDBSection, 'LoginUserName', '');
    edtUserPass.Text     := DecryptString(ini.ReadString(c_strIniUISection, 'LoginUserPass', ''), c_strAESKey);
    chkUserName.Checked  := True;
    chkAutoLogin.Checked := True;
  end
  else
  begin
    if ini.ReadBool(c_strIniDBSection, 'CheckLoginUserName', False) then
    begin
      edtUserName.Text     := ini.ReadString(c_strIniDBSection, 'LoginUserName', '');
      edtUserPass.Text     := '';
      chkUserName.Checked  := True;
      chkAutoLogin.Checked := False;
      Winapi.Windows.SetFocus(edtUserPass.Handle);
    end;
  end;
end;

procedure TfrmLogin.SaveLoginInfo(var ini: TIniFile);
begin
  ini.WriteBool(c_strIniDBSection, 'CheckLoginUserName', chkUserName.Checked);
  ini.WriteBool(c_strIniDBSection, 'CheckLoginAuto', chkAutoLogin.Checked);

  if chkAutoLogin.Checked then
  begin
    ini.WriteString(c_strIniDBSection, 'LoginUserName', edtUserName.Text);
    ini.WriteString(c_strIniUISection, 'LoginUserPass', EncryptString(edtUserPass.Text, c_strAESKey));
  end
  else
  begin
    if chkUserName.Checked then
    begin
      ini.WriteString(c_strIniDBSection, 'LoginUserName', edtUserName.Text);
      ini.WriteString(c_strIniUISection, 'LoginUserPass', '');
    end
    else
    begin
      ini.WriteString(c_strIniDBSection, 'LoginUserName', '');
      ini.WriteString(c_strIniUISection, 'LoginUserPass', '');
    end;
  end;
end;

procedure TfrmLogin.btnCancelClick(Sender: TObject);
begin
  FbResult := False;
  Close;
end;

procedure TfrmLogin.btnSaveClick(Sender: TObject);
var
  strSQL: String;
  qry   : TADOQuery;
begin
  if Trim(edtUserName.Text) = '' then
  begin
    MessageBox(Handle, '用户名称不能为空，请输入', '系统提示：', MB_OK OR MB_ICONERROR);
    edtUserName.SetFocus;
    Exit;
  end;

  if Trim(edtUserPass.Text) = '' then
  begin
    MessageBox(Handle, '用户密码不能为空，请输入', '系统提示：', MB_OK OR MB_ICONERROR);
    edtUserPass.SetFocus;
    Exit;
  end;

  if (g_ADOCNN.Connected) and (g_strLoginTable <> '') and (g_strLoginName <> '') and (g_strLoginPass <> '') then
  begin
    strSQL         := Format('select * from %s where %s=%s and %s=%s', [g_strLoginTable, g_strLoginName, QuotedStr(edtUserName.Text), g_strLoginPass, QuotedStr(EncDatabasePassword(edtUserPass.Text))]);
    qry            := TADOQuery.Create(nil);
    qry.Connection := g_ADOCNN;
    qry.SQL.Text   := strSQL;
    qry.Open;
    if qry.RecordCount > 0 then
    begin
      { 登录成功 }
      FbResult := True;
      Close;
    end
    else
    begin
      MessageBox(Handle, '用户不存在或密码错误，请重新输入', '系统提示：', MB_OK OR MB_ICONERROR);
    end;
    qry.Free;
  end;
end;

procedure TfrmLogin.chkAutoLoginClick(Sender: TObject);
begin
  if chkAutoLogin.Checked then
    chkUserName.Checked := True;
end;

procedure TfrmLogin.chkUserNameClick(Sender: TObject);
begin
  if not chkUserName.Checked then
    chkAutoLogin.Checked := False;
end;

procedure TfrmLogin.edtUserNameKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if Key = VK_RETURN then
    edtUserPass.SetFocus;
end;

procedure TfrmLogin.edtUserPassKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if Key = VK_RETURN then
    btnSave.Click;
end;

end.
