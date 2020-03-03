unit FormSelectValues;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, vcl.Graphics,
  vcl.Controls, vcl.Forms, vcl.Dialogs, vcl.StdCtrls, vcl.CheckLst, XLSSharedItems5;

type
  TfrmSelectValues = class(TForm)
    clbValues: TCheckListBox;
    Label1: TLabel;
    btnOk: TButton;
    btnCancel: TButton;
    procedure clbValuesClickCheck(Sender: TObject);
  private
    procedure AddValues(AValues,AFilter: TXLSUniqueSharedItemsValues);
  public
    function Execute(AValues,AFilter: TXLSUniqueSharedItemsValues): boolean;
  end;

implementation

{$R *.dfm}

{ TfrmSelectValues }

procedure TfrmSelectValues.AddValues(AValues,AFilter: TXLSUniqueSharedItemsValues);
var
  i: integer;
begin
  clbValues.Clear;

  clbValues.Items.Add('(Select all)');

  for i := 0 to AValues.Count - 1 do begin
    clbValues.AddItem(AValues[i].AsText,AValues[i]);

    clbValues.Checked[i + 1] := (AFilter.Count = 0) or AFilter.Find(AValues[i]);
  end;

  clbValues.Checked[0] := AFilter.Count = 0;
end;

procedure TfrmSelectValues.clbValuesClickCheck(Sender: TObject);
var
  i: integer;
  C: boolean;
begin
  if clbValues.ItemIndex = 0 then begin
    C := clbValues.Checked[0];

    for i := 1 to clbValues.Count - 1 do
      clbValues.Checked[i] := C;
  end
  else begin
    clbValues.Checked[0] := False;
    for i := 1 to clbValues.Count - 1 do begin
      if not clbValues.Checked[i] then
        Exit;
    end;
    clbValues.Checked[0] := True;
  end;
end;

function TfrmSelectValues.Execute(AValues,AFilter: TXLSUniqueSharedItemsValues): boolean;
var
  i: integer;
begin
  AddValues(AValues,AFilter);

  ShowModal;

  Result := ModalResult = mrOk;

  if Result then begin
    AFilter.Clear;

    if not clbValues.Checked[0] then begin
      for i := 0 to AValues.Count - 1 do begin
        if clbValues.Checked[i + 1] then
          AFilter.Add(AValues[i].Clone);
      end;
    end;
  end;
end;

end.
