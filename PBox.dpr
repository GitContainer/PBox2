program PBox;
{$IF CompilerVersion >= 21.0}
{$WEAKLINKRTTI ON}
{$RTTI EXPLICIT METHODS([]) PROPERTIES([]) FIELDS([])}
{$IFEND}

uses
  Vcl.Forms,
  db.uBaseForm in 'db.uBaseForm.pas',
  db.uCommon in 'db.uCommon.pas',
  db.uHashCode in 'db.uHashCode.pas',
  db.uNetworkManager in 'db.uNetworkManager.pas',
  db.uCreateVCDialogDll in 'db.uCreateVCDialogDll.pas',
  db.uCreateDelphiDll in 'db.uCreateDelphiDll.pas',
  db.uCreateEXE in 'db.uCreateEXE.pas',
  db.PBoxForm in 'db.PBoxForm.pas' {frmPBox},
  db.ConfigForm in 'db.ConfigForm.pas' {frmConfig},
  db.AddEXE in 'db.AddEXE.pas' {frmAddEXE},
  db.DBConfig in 'db.DBConfig.pas' {DBConfig},
  db.LoginForm in 'db.LoginForm.pas' {frmLogin};

{$R *.res}

begin
  OnlyOneRunInstance;
  CheckLoginForm;
  Application.Initialize;
  ReportMemoryLeaksOnShutdown   := True;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TfrmPBox, frmPBox);
  Application.Run;

end.
