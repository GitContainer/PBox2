unit XSSIERenderTextIE;

{-
********************************************************************************
******* XLSSpreadSheet V3.00                                             *******
*******                                                                  *******
******* Copyright(C) 2006,2017 Lars Arvidsson, Axolot Data               *******
*******                                                                  *******
******* email: components@axolot.com                                     *******
******* URL:   http://www.axolot.com                                     *******
********************************************************************************
** Users of the XLSSpreadSheet component must accept the following            **
** disclaimer of warranty:                                                    **
**                                                                            **
** XLSSpreadSheet is supplied as is. The author disclaims all warranties,     **
** expressed or implied, including, without limitation, the warranties of     **
** merchantability and of fitness for any purpose. The author assumes no      **
** liability for damages, direct or consequential, which may result from the  **
** use of XLSSpreadSheet.                                                     **
********************************************************************************
}

{$B-}
{$H+}
{$R-}
{$I AxCompilers.inc}

interface

uses Classes, SysUtils,
     XBookPaintGDI2,
     XSSIEUtils, XSSIELogPhyPosition, XSSIEGDIText, XSSIEDocProps,
     XSSIELogParas, XSSIELogEditor, XSSIEPhyRow, XSSIECaretRow;

type TAXWTextRender = class(TObject)
protected
     FGDI       : TAXWGDI;
     FDoc       : TAXWLogDocEditor;
     FArea      : TAXWClientArea;
     FDOP       : TAXWDOP;

     FTextPrint : TAXWTextPrint;
     FCaretRow  : TAXWCaretRowEditor;

     FVertAlign : TAXWVertAlign;

     procedure SetupText;
     function  ParasHeight: integer;
public
     constructor Create(AGDI: TAXWGDI; ADoc: TAXWLogDocEditor; ACaretRow: TAXWCaretRowEditor; AArea: TAXWClientArea);
     destructor Destroy; override;

     function  CalcOriginY: integer;

     procedure Render;
     end;


implementation

{ TAXWTextRender }

constructor TAXWTextRender.Create(AGDI: TAXWGDI; ADoc: TAXWLogDocEditor; ACaretRow: TAXWCaretRowEditor; AArea: TAXWClientArea);
begin
  FGDI := AGDI;
  FDoc := ADoc;
  FCaretRow := ACaretRow;
  FArea := AArea;

  FDOP := TAXWDOP.Create;
  FTextPrint := TAXWTextPrint.Create(FGDI,FArea,FDoc.Selections,FDOP);
end;

destructor TAXWTextRender.Destroy;
begin
  FTextPrint.Free;
  FDOP.Free;

  inherited;
end;

function TAXWTextRender.CalcOriginY: integer;
var
  H: integer;
  Row: TAXWPhyRow;
begin
  Row := FDoc.Paras[0].Rows[0];
  case FVertAlign of
    avaBottom     : begin
      Result := Round(FArea.Y1 + FArea.PaperHeight - FDoc.Paras[0].Rows.Height + Row.Ascent);
      if (Result - Row.Ascent) < FArea.Y1 then
        Result := Round(FArea.Y1 + Row.Ascent);
    end;
    avaDistributed,
    avaCenter     : begin
      if FDoc.Paras[0].Rows.Count > 1 then
        Result := Round(FArea.Y1 + (FArea.PaperHeight / 2) - (FDoc.Paras[0].Rows.Height / 2) + Row.Ascent)
      else begin
        H := ParasHeight;
        if H > FArea.PaperHeight then
          Result := Round(FArea.Y1 + (FArea.PaperHeight / 2) - (Row.Height / 2))
        else
          Result := Round(FArea.Y1 + (FArea.PaperHeight / 2) + (Row.Ascent / 2) - (Row.Descent / 2));
      end;
    end;
    avaTop        : Result := Round(FArea.Y1 + Row.Ascent);
    else            Result := FArea.Y1;
  end;
end;

function TAXWTextRender.ParasHeight: integer;
var
  i,j: integer;
  H: double;
begin
  H := 0;

  for i := 0 to FDoc.Paras.Count - 1 do begin
    for j := 0 to FDoc.Paras[i].Rows.Count - 1 do
      H := H + FDoc.Paras[i].Rows.Height;
  end;
  Result := Round(H);
end;

procedure TAXWTextRender.Render;
var
  Y    : double;
  iPara: integer;
  iRow : integer;
  Para : TAXWLogPara;
  Row  : TAXWPhyRow;
  CHP  : TAXWCHP;
  PAP  : TAXWPAP;
begin
  SetupText;

  Y := CalcOriginY;

  for iPara := 0 to FDoc.Paras.Count - 1 do begin
    Para := FDoc.Paras[iPara];
    PAP := TAXWPAP.Create(FDoc.MasterPAP);
    PAP.Assign(Para.PAPX);
    try
      for iRow := 0 to Para.Rows.Count - 1 do begin
        Row := Para.Rows[iRow];

        CHP := TAXWCHP.Create(FDoc.MasterCHP);
        try
          FDoc.Selections.Hit(Para,Row);
          FTextPrint.PrintRow(Para,CHP,PAP,Row,Y);

          if FCaretRow.CaretPos.Row = Row then begin
            FCaretRow.ScrCaretY := Y - Row.Ascent;
            FCaretRow.ScrCaretX := FArea.CX1 + Round(FCaretRow[FCaretRow.CaretPos.RelativePos].X);
            FCaretRow.ScrCaretH := FCaretRow.CaretPos.Row.Height;
          end;
          Y := Y + Row.Height;
        finally
          CHP.Free;
        end;
      end;
    finally
      PAP.Free;
    end;
  end;
end;

procedure TAXWTextRender.SetupText;
begin
  FGDI.TransparentMode := True;
  FGDI.SetTextAlign(xhtaDefault,xvtaBaseline);
end;

end.
