unit FrmColor;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExcelColorPicker, StdCtrls,
  Xc12Utils5;

type
  TfrmSelectColor = class(TForm)
    Button1: TButton;
    Button2: TButton;
    ecpTheme: TExcelColorPicker;
    ecpStandard: TExcelColorPicker;
    Label1: TLabel;
    Label2: TLabel;
  private
  public
    function Execute(var AColor: TXc12Color): boolean;
  end;

implementation

{$R *.dfm}

{ TfrmSelectColor }

function TfrmSelectColor.Execute(var AColor: TXc12Color): boolean;
begin
  ecpTheme.ExcelColor := AColor;

  ShowModal;

  Result := ModalResult = mrOk;

  if Result then
    AColor := ecpTheme.ExcelColor;
end;

end.

