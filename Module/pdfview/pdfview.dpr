library pdfview;

{$IF CompilerVersion >= 21.0}
{$WEAKLINKRTTI ON}
{$RTTI EXPLICIT METHODS([]) PROPERTIES([]) FIELDS([])}
{$IFEND}

uses
  System.SysUtils,
  System.Classes,
  Main in 'Main.pas' {MainForm},
  PDFium.Frame in 'PDFium.Frame.pas' {PDFiumFrame: TFrame};

{$R *.res}

exports
  db_ShowDllForm_Plugins;


begin
end.
