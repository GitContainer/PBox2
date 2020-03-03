unit FormSelectChart;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics,
  Controls, Forms, Dialogs, StdCtrls;

type
  TfrmSelectChart = class(TForm)
    Label1: TLabel;
    cbStyle: TComboBox;
    Button1: TButton;
    Button2: TButton;
  private
    { Private declarations }
  public
    function Execute: integer;
  end;

var
  frmSelectChart: TfrmSelectChart;

implementation

{$R *.dfm}

{ TfrmSelectChart }

function TfrmSelectChart.Execute: integer;
begin
  cbStyle.ItemIndex := 0;

  ShowModal;

  if ModalResult = mrOk then
    Result := cbStyle.ItemIndex
  else
    Result := -1;
end;

end.
