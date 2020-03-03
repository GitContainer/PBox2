unit BIFF_DrawingObjAnchor5;

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
     BIFF_Escher5;

type
//* Base class for drawing objects and control objects, that have their
//* coordinates given in rows and columns.
    TDrwAnchor = class(TCollectionItem)
private
     function GetSpId: longword;
     function  GetCol1: integer;
     function  GetCol1Offset: double;
     function  GetCol2: integer;
     function  GetCol2Offset: double;
     function  GetRow1: integer;
     function  GetRow1Offset: double;
     function  GetRow2: integer;
     function  GetRow2Offset: double;
     procedure SetCol1(const Value: integer);
     procedure SetCol1Offset(Value: double);
     procedure SetCol2(const Value: integer);
     procedure SetCol2Offset(Value: double);
     procedure SetRow1(const Value: integer);
     procedure SetRow1Offset(Value: double);
     procedure SetRow2(const Value: integer);
     procedure SetRow2Offset(Value: double);
     function  GetFlipHorizontal: boolean;
     function  GetFlipVertical: boolean;
     procedure SetFlipHorizontal(const Value: boolean);
     procedure SetFlipVertical(const Value: boolean);
protected
     FName: string;
     FFileAdd: boolean;
     FShape: TShapeClientAnchor;

     function GetDisplayName: string; override;
     procedure SetShape(const Value: TShapeClientAnchor); virtual;
public
     property Shape: TShapeClientAnchor read FShape;
published
     //* Left column for the object.
     property Col1:       integer read GetCol1 write SetCol1;
     //* Offset in Col1 for the object. Units are in percent, where the
     //* value can be between 0 and 1.
     property Col1Offset: double  read GetCol1Offset write SetCol1Offset;
     //* Right column for the object.
     property Col2:       integer read GetCol2 write SetCol2;
     //* Offset in Col2 for the object. Units are in percent, where the
     //* value can be between 0 and 1.
     property Col2Offset: double  read GetCol2Offset write SetCol2Offset;
     //* Top row for the object.
     property Row1:       integer read GetRow1 write SetRow1;
     //* Offset in Row1 for the object. Units are in percent, where the
     //* value can be between 0 and 1.
     property Row1Offset: double  read GetRow1Offset write SetRow1Offset;
     //* Top bottom for the object.
     property Row2:       integer read GetRow2 write SetRow2;
     //* Offset in Row2 for the object. Units are in percent, where the
     //* value can be between 0 and 1.
     property Row2Offset: double  read GetRow2Offset write SetRow2Offset;
     //* Set to True if the object shall be flipped horizontal, i.e. mirrored.
     property FlipHorizontal: boolean read GetFlipHorizontal write SetFlipHorizontal;
     //* Set to True if the object shall be flipped vertical, i.e. upside
     //* down.
     property FlipVertical: boolean read GetFlipVertical write SetFlipVertical;
     //* ~exclude
     property SpId: longword read GetSpId;
     end;

implementation

{ TDrwAnchor }

function TDrwAnchor.GetCol1: integer;
begin
  Result := FShape.Anchor.Col1;
end;

function TDrwAnchor.GetCol1Offset: double;
begin
  Result := FShape.Anchor.Col1Offset / 1024;
end;

function TDrwAnchor.GetCol2: integer;
begin
  Result := FShape.Anchor.Col2;
end;

function TDrwAnchor.GetCol2Offset: double;
begin
  Result := FShape.Anchor.Col2Offset / 1024;
end;

function TDrwAnchor.GetRow1: integer;
begin
  Result := FShape.Anchor.Row1;
end;

function TDrwAnchor.GetRow1Offset: double;
begin
  Result := FShape.Anchor.Row1Offset / 256;
end;

function TDrwAnchor.GetRow2: integer;
begin
  Result := FShape.Anchor.Row2;
end;

function TDrwAnchor.GetRow2Offset: double;
begin
  Result := FShape.Anchor.Row2Offset / 256;
end;

function TDrwAnchor.GetSpId: longword;
begin
  Result := FShape.SpId;
end;

procedure TDrwAnchor.SetCol1(const Value: integer);
begin
  FShape.Anchor.Col1 := Value;
end;

procedure TDrwAnchor.SetCol1Offset(Value: double);
begin
  if Value > 1.0 then
    Value := 1.0
  else if Value < 0.0 then
    Value := 0.0;
  FShape.Anchor.Col1Offset := Round(Value * 1024);
end;

procedure TDrwAnchor.SetCol2(const Value: integer);
begin
  FShape.Anchor.Col2 := Value;
end;

procedure TDrwAnchor.SetCol2Offset(Value: double);
begin
  if Value > 1.0 then
    Value := 1.0
  else if Value < 0.0 then
    Value := 0.0;
  FShape.Anchor.Col2Offset := Round(Value * 1024);
end;

procedure TDrwAnchor.SetShape(const Value: TShapeClientAnchor);
begin
  FShape := Value;
end;

procedure TDrwAnchor.SetRow1(const Value: integer);
begin
  FShape.Anchor.Row1 := Value;
end;

procedure TDrwAnchor.SetRow1Offset(Value: double);
begin
  if Value > 1.0 then
    Value := 1.0
  else if Value < 0.0 then
    Value := 0.0;
  FShape.Anchor.Row1Offset := Round(Value * 256);
end;

procedure TDrwAnchor.SetRow2(const Value: integer);
begin
  FShape.Anchor.Row2 := Value;
end;

procedure TDrwAnchor.SetRow2Offset(Value: double);
begin
  if Value > 1.0 then
    Value := 1.0
  else if Value < 0.0 then
    Value := 0.0;
  FShape.Anchor.Row2Offset := Round(Value * 256);
end;

function TDrwAnchor.GetDisplayName: string;
begin
  inherited GetDisplayName;
  Result := FName;
end;

function TDrwAnchor.GetFlipHorizontal: boolean;
begin
  Result := FShape.FlipHorizontal;
end;

function TDrwAnchor.GetFlipVertical: boolean;
begin
  Result := FShape.FlipVertical;
end;

procedure TDrwAnchor.SetFlipHorizontal(const Value: boolean);
begin
  FShape.FlipHorizontal := Value;
end;

procedure TDrwAnchor.SetFlipVertical(const Value: boolean);
begin
  FShape.FlipVertical := Value;
end;

end.
