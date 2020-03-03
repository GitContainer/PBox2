unit BIFF_EscherFileObjs5;

{-
********************************************************************************
******* XLSReadWriteII V6.00                                             *******
*******                                                                  *******
******* Copyright(C) 1999,2017 Lars Arvidsson, Axolot Data               *******
*******                                                                  *******
******* email: components@axolot.com                                     *******
******* URL:   http://www.axolot.com                                     *******
********************************************************************************
** Users of the XLSReadWriteII component must accept the following            **
** disclaimer of warranty:                                                    **
**                                                                            **
** XLSReadWriteII is supplied as is. The author disclaims all warranties,     **
** expressedor implied, including, without limitation, the warranties of      **
** merchantability and of fitness for any purpose. The author assumes no      **
** liability for damages, direct or consequential, which may result from the  **
** use of XLSReadWriteII.                                                     **
********************************************************************************
}

{$B-}
{$H+}
{$R-}
{$I AxCompilers.inc}
{$I XLSRWII.inc}

interface

uses Classes, SysUtils,
     Xc12Utils5,
     XLSUtils5,
     BIFF_RecsII5, BIFF_EscherTypes5, BIFF_Stream5;

type TMSOObjCLIENTANCHOR = class(TXLSStreamObj)
private
     FOptions: longword;
     FCol1: longword;
     FCol1Offset: longword;
     FRow1: longword;
     FRow1Offset: longword;
     FCol2: longword;
     FCol2Offset: longword;
     FRow2: longword;
     FRow2Offset: longword;
public
     procedure Read(Stream: TXLSStream); override;
     procedure Write(Stream: TXLSStream); override;
     procedure Assign(Anchor: TMSOObjCLIENTANCHOR);

     property Options:    longword read FOptions    write FOptions;
     property Col1:       longword read FCol1       write FCol1;
     property Col1Offset: longword read FCol1Offset write FCol1Offset;
     property Row1:       longword read FRow1       write FRow1;
     property Row1Offset: longword read FRow1Offset write FRow1Offset;
     property Col2:       longword read FCol2       write FCol2;
     property Col2Offset: longword read FCol2Offset write FCol2Offset;
     property Row2:       longword read FRow2       write FRow2;
     property Row2Offset: longword read FRow2Offset write FRow2Offset;
     end;

implementation

{ TMSOObjCLIENTANCHOR }

procedure TMSOObjCLIENTANCHOR.Assign(Anchor: TMSOObjCLIENTANCHOR);
begin
  Anchor.Options    := FOptions;
  Anchor.Col1       := FCol1;
  Anchor.Col1Offset := FCol1Offset;
  Anchor.Row1       := FRow1;
  Anchor.Row1Offset := FRow1Offset;
  Anchor.Col2       := FCol2;
  Anchor.Col2Offset := FCol2Offset;
  Anchor.Row2       := FRow2;
  Anchor.Row2Offset := FRow2Offset;
end;

procedure TMSOObjCLIENTANCHOR.Read(Stream: TXLSStream);
var
  Rec: TMSORecCLIENTANCHOR;
begin
  case Stream.FileVersion of
    xvExcel97: begin
      Stream.Read(Rec,SizeOf(TMSORecCLIENTANCHOR));
      FOptions    := Rec.Options;
      FCol1       := Rec.Col1;
      FCol1Offset := Rec.Col1Offset;
      FRow1       := Rec.Row1;
      FRow1Offset := Rec.Row1Offset;
      FCol2       := Rec.Col2;
      FCol2Offset := Rec.Col2Offset;
      FRow2       := Rec.Row2;
      FRow2Offset := Rec.Row2Offset;
    end;
    else
      raise XLSRWException.Create('Unhandled excel version');
  end;
end;

procedure TMSOObjCLIENTANCHOR.Write(Stream: TXLSStream);
var
  Rec: TMSORecCLIENTANCHOR;
begin
  case Stream.FileVersion of
    xvExcel97: begin
      Rec.Options    := FOptions;
      Rec.Col1       := FCol1;
      Rec.Col1Offset := FCol1Offset;
      Rec.Row1       := FRow1;
      Rec.Row1Offset := FRow1Offset;
      Rec.Col2       := FCol2;
      Rec.Col2Offset := FCol2Offset;
      Rec.Row2       := FRow2;
      Rec.Row2Offset := FRow2Offset;
      Stream.WriteMSOHeader($00,$0000,MSO_CLIENTANCHOR,SizeOf(TMSORecCLIENTANCHOR));
      Stream.Write(Rec,SizeOf(TMSORecCLIENTANCHOR));
    end;
    else
      raise XLSRWException.Create('Unhandled excel version');
  end;
end;

end.
