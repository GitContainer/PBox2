library HexEdit;
{$IF CompilerVersion >= 21.0}
{$WEAKLINKRTTI ON}
{$RTTI EXPLICIT METHODS([]) PROPERTIES([]) FIELDS([])}
{$IFEND}

uses
  System.SysUtils,
  System.Classes,
  uMainForm in 'uMainForm.pas' {frmHexEdit},
  sample2 in 'sample2.pas' {fmPreview},
  sample3 in 'sample3.pas' {Form2};

{$R *.res}

exports
  db_ShowDllForm_Plugins;

begin

end.
