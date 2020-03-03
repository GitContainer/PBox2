unit sample2;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ExtCtrls;

type
  TfmPreview = class(TForm)
    Image1: TImage;
  private
    { Private-Deklarationen }
  public
    { Public-Deklarationen }
  end;

var
  fmPreview: TfmPreview;

implementation

{$R *.DFM}

end.
