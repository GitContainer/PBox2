library PEInfo;
{$IF CompilerVersion >= 21.0}
{$WEAKLINKRTTI ON}
{$RTTI EXPLICIT METHODS([]) PROPERTIES([]) FIELDS([])}
{$IFEND}

uses
  System.SysUtils,
  System.Classes,
  uHexEditor in 'uHexEditor.pas',
  uResource in 'uResource.pas',
  uMainForm in 'uMainForm.pas' {frmPEInfo};

{$R *.res}

exports
  db_ShowDllForm_Plugins;

begin

end.
