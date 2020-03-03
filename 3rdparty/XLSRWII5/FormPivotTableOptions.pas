unit FormPivotTableOptions;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, vcl.Graphics,
  vcl.Controls, vcl.Forms, vcl.Dialogs, vcl.StdCtrls, XLSPivotTables5;

type
  TfrmPivotTableOptions = class(TForm)
    btnOk: TButton;
    btnCancel: TButton;
    Label1: TLabel;
    edName: TEdit;
    cbTotalsCols: TCheckBox;
    cbTotalsRows: TCheckBox;
  private
    { Private declarations }
  public
    function Execute(APivTable: TXLSPivotTable): boolean;
  end;

implementation

{$R *.dfm}

{ TfrmPivotTableOptions }

function TfrmPivotTableOptions.Execute(APivTable: TXLSPivotTable): boolean;
begin
  edName.Text := APivTable.Name;
  cbTotalsRows.Checked := APivTable.RowsGrandTotals;
  cbTotalsCols.Checked := APivTable.ColsGrandTotals;

  ShowModal;

  Result := ModalResult = mrOk;

  if Result then begin
    APivTable.Name := edName.Text;
    APivTable.RowsGrandTotals := cbTotalsRows.Checked;
    APivTable.ColsGrandTotals := cbTotalsCols.Checked;
  end;
end;

end.
