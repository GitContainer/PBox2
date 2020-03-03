unit FormPivotTableField;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, vcl.Graphics,
  vcl.Controls, vcl.Forms, vcl.Dialogs, vcl.ComCtrls, vcl.StdCtrls, vcl.ExtCtrls,
  XLSPivotTables5;

type
  TfrmPivotTableField = class(TForm)
    btnOk: TButton;
    btnCancel: TButton;
    lblSourceName: TLabel;
    Label1: TLabel;
    edCustomName: TEdit;
    Label2: TLabel;
    rgSubtotals: TRadioGroup;
    lbFunctions: TListBox;
    Label3: TLabel;
    procedure rgSubtotalsClick(Sender: TObject);
    procedure lbFunctionsClick(Sender: TObject);
  private
    FField: TXLSPivotField;
  public
    function Execute(AField: TXLSPivotField): boolean;
  end;

implementation

{$R *.dfm}

{ TfrmPivotTableField }

function TfrmPivotTableField.Execute(AField: TXLSPivotField): boolean;
var
  i: TXLSPivotFunc;
begin
  FField := AField;

  lblSourceName.Caption := 'Source name: ' + FField.Name;
  edCustomName.Text := FField.UserName;

  for i := Low(TXLSPivotFunc) to Pred(xpfDefault) do
    lbFunctions.Items.Add(XLSPivotFuncStr[i]);

  case FField.Func of
    xpfDefault: rgSubtotals.ItemIndex := 0;
    xpfNone   : rgSubtotals.ItemIndex := 1;
    else begin
      rgSubtotals.ItemIndex := 2;
      lbFunctions.ItemIndex := Integer(FField.Func);
    end;
  end;

  ShowModal;

  Result := ModalResult = mrOk;

  if Result then begin
    AField.UserName := edCustomName.Text;

    case rgSubtotals.ItemIndex of
      0: FField.Func := xpfDefault;
      1: FField.Func := xpfNone;
    end;
  end;
end;

procedure TfrmPivotTableField.lbFunctionsClick(Sender: TObject);
begin
  FField.Func := TXLSPivotFunc(lbFunctions.ItemIndex);
end;

procedure TfrmPivotTableField.rgSubtotalsClick(Sender: TObject);
begin
  lbFunctions.Enabled := rgSubtotals.ItemIndex > 1;

  if lbFunctions.Enabled then
    lbFunctions.ItemIndex := Integer(FField.Func);
end;

end.
