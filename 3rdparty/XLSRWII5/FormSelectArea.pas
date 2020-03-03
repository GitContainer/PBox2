unit FormSelectArea;

{$B-}
{$H+}
{$R-}
{$I AxCompilers.inc}

interface

uses
  Windows, Messages, SysUtils, Variants, Classes,
{$ifdef DELPHI_XE5_OR_LATER}
  UITypes,
{$endif}
  Graphics, Controls, Forms, Dialogs, StdCtrls,
  Xc12Utils5, XLSRelCells5;

type
  TfrmSelectArea = class(TForm)
    Label1: TLabel;
    edArea: TEdit;
    btnOk: TButton;
    btnCancel: TButton;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }
  public
    function Execute(ARef: TXLSRelCells): boolean;
  end;

implementation

{$R *.dfm}

{ TfrmSelectArea }

function TfrmSelectArea.Execute(ARef: TXLSRelCells): boolean;
begin
  edArea.Text := ARef.ShortRef;

  ShowModal;

  Result := ModalResult = mrOk;

  if Result then
    ARef.Ref := edArea.Text;
end;

procedure TfrmSelectArea.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  if IsAreaStr(edArea.Text) then
    Action := caFree
  else begin
    MessageDlg('Area is not valid',mtError,[mbOk],0);
    Action := caNone;
  end;
end;

end.
