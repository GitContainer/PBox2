unit sample3;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls;

type
  TForm2 = class(TForm)
    Label1: TLabel;
    Label2: TLabel;
    ListBox1: TListBox;
    ListBox2: TListBox;
    Button1: TButton;
    Button2: TButton;
    procedure ListBox1Click(Sender: TObject);
  private
    { Private-Deklarationen }
  public
    { Public-Deklarationen }
  end;

var
  Form2: TForm2;

implementation

{$R *.DFM}


procedure TForm2.ListBox1Click(Sender: TObject);
begin
     Button1.Enabled := (ListBox1.ItemIndex > -1 ) and
                        (ListBox2.ItemIndex > -1 ) and
                        (ListBox2.ItemIndex <> ListBox1.ItemIndex );
end;

end.
