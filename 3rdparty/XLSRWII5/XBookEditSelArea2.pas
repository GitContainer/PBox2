unit XBookEditSelArea2;

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

uses { Delphi  } Classes, SysUtils,
                 Xc12Utils5,
     { XLSRWII } XLSUtils5, XLSSheetData5,
     { XLSBook } XBookUtils2, XBookSkin2, XBookPaint2, XBookColumns2, XBookRows2;

type TSelEditMode = (semMove,semSize);

type TXLSBookEditSelArea = class(TObject)
protected
     FSkin: TXLSBookSkin;
     FLastCol,FLastRow: integer;
     FSelectedArea: TXLSSelectedArea;
     FSheetArea: TXLSCellArea;
     FColumns: TXLSBookColumns;
     FRows: TXLSBookRows;
     FStarted: boolean;

     procedure ClipToSheet(var C1, R1, C2, R2: integer);
public
     constructor Create(Skin: TXLSBookSkin; Columns: TXLSBookColumns; Rows: TXLSBookRows);
     procedure PaintRect; virtual; abstract;
     procedure BeginEdit(SelectedArea: TXLSSelectedArea; Col,Row: integer); virtual; abstract;
     procedure UpdateEdit(Col,Row: integer; ClipToSheet: boolean); virtual; abstract;
     procedure EndEdit(Col,Row: integer); virtual; abstract;
     function  Mode: TSelEditMode; virtual; abstract;
     procedure ScrollUpdate(Col,Row: integer); virtual; abstract;
     procedure SheetAreaChanged;

     property Started: boolean read FStarted;
     end;

type TXLSBookMoveSelArea = class(TXLSBookEditSelArea)
protected
     FMouseDownCol,FMouseDownRow: integer;
     FDeltaCol,FDeltaRow: integer;

     procedure MakeMoveRect(C1, R1, C2, R2: integer; var Lines: TStraightLineArray);
     procedure ClipMoveRect;
public
     constructor Create(Skin: TXLSBookSkin; Columns: TXLSBookColumns; Rows: TXLSBookRows; MouseCol,MouseRow: integer);
     procedure PaintRect; override;
     procedure BeginEdit(SelectedArea: TXLSSelectedArea; Col,Row: integer); override;
     procedure UpdateEdit(Col,Row: integer; ClipToSheet: boolean); override;
     procedure EndEdit(Col,Row: integer); override;
     function  Mode: TSelEditMode; override;
     procedure ScrollUpdate(Col,Row: integer); override;

     property DeltaCol: integer read FDeltaCol;
     property DeltaRow: integer read FDeltaRow;
     end;

type TXLSBookSizeSelArea = class(TXLSBookEditSelArea)
protected
     FSizeArea: TXLSCellArea;
     FDeleteArea: TXLSCellArea;
     FIsDelete,FLastWasDelete: boolean;

     procedure SetSizeArea(Col,Row: integer);
     procedure SetDeleteArea(Col,Row: integer);
     procedure MakeSizeRect(C1, R1, C2, R2: integer; var Lines: TStraightLineArray);
     procedure DoDeleteArea;
public
     procedure PaintRect; override;
     procedure BeginEdit(SelectedArea: TXLSSelectedArea; Col,Row: integer); override;
     procedure UpdateEdit(Col,Row: integer; ClipToSheet: boolean); override;
     procedure EndEdit(Col,Row: integer); override;
     function  Mode: TSelEditMode; override;
     procedure ScrollUpdate(Col,Row: integer); override;

     property IsDelete: boolean read FIsDelete write FIsDelete;
     property SizeArea: TXLSCellArea read FSizeArea;
     property DeleteArea: TXLSCellArea read FDeleteArea;
     end;

implementation

{ TXLSBookEditSelArea }

procedure TXLSBookEditSelArea.ClipToSheet(var C1, R1, C2, R2: integer);
begin
       if R1 < FSheetArea.Row1 then R1 := FSheetArea.Row1
  else if R1 > FSheetArea.Row2 then R1 := FSheetArea.Row2;
       if R2 < FSheetArea.Row1 then R2 := FSheetArea.Row1
  else if R2 > FSheetArea.Row2 then R2 := FSheetArea.Row2;

       if C1 < FSheetArea.Col1 then C1 := FSheetArea.Col1
  else if C1 > FSheetArea.Col2 then C1 := FSheetArea.Col2;
       if C2 < FSheetArea.Col1 then C2 := FSheetArea.Col1
  else if C2 > FSheetArea.Col2 then C2 := FSheetArea.Col2;
end;

constructor TXLSBookEditSelArea.Create(Skin: TXLSBookSkin; Columns: TXLSBookColumns; Rows: TXLSBookRows);
begin
  FSkin := Skin;
  FColumns := Columns;
  FRows := Rows;
  SheetAreaChanged;
end;

procedure TXLSBookEditSelArea.SheetAreaChanged;
begin
  FSheetArea.Col1 := FColumns.LeftCol;
  FSheetArea.Col2 := FColumns.RightCol;
  FSheetArea.Row1 := FRows.TopRow;
  FSheetArea.Row2 := FRows.BottomRow;
end;

{ TXLSBookMoveSelArea }

procedure TXLSBookMoveSelArea.BeginEdit(SelectedArea: TXLSSelectedArea; Col, Row: integer);
begin
  FSelectedArea := SelectedArea;
  FStarted := True;
  if FMouseDownRow < FSelectedArea.Row1 then
    Inc(FMouseDownRow)
  else if FMouseDownRow > FSelectedArea.Row2 then
    Dec(FMouseDownRow);
  if FMouseDownCol < FSelectedArea.Col1 then
    Inc(FMouseDownCol)
  else if FMouseDownCol > FSelectedArea.Col2 then
    Dec(FMouseDownCol);
  FDeltaCol := Col - FMouseDownCol;
  FDeltaRow := Row - FMouseDownRow;
  PaintRect;
  FLastCol := Col;
  FLastRow := Row;
end;

procedure TXLSBookMoveSelArea.ClipMoveRect;
begin
  with FSelectedArea do begin
    if (Col1 + FDeltaCol) < 0 then
      FDeltaCol := -Col1
    else if (Col2 + FDeltaCol) > XLS_MAXCOL then
      FDeltaCol := XLS_MAXCOL - Col2;
    if (Row1 + FDeltaRow) < 0 then
      FDeltaRow := -Row1
    else if (Row2 + FDeltaRow) > XLS_MAXROW then
      FDeltaRow := XLS_MAXROW - Row2;
  end;
end;

constructor TXLSBookMoveSelArea.Create(Skin: TXLSBookSkin; Columns: TXLSBookColumns; Rows: TXLSBookRows; MouseCol,MouseRow: integer);
begin
  inherited Create(Skin,Columns,Rows);
  FMouseDownCol := MouseCol;
  FMouseDownRow := MouseRow;
end;

procedure TXLSBookMoveSelArea.EndEdit(Col, Row: integer);
begin
  PaintRect;
end;

procedure TXLSBookMoveSelArea.MakeMoveRect(C1, R1, C2, R2: integer; var Lines: TStraightLineArray);
var
  Col,Row: integer;
  x1,y1,x2,y2: integer;
  bC1,bR1,bC2,bR2: boolean;

procedure AddLine(A1,A2,B: integer; Horiz: boolean);
var
  i: integer;
begin
  i := Length(Lines);
  SetLength(Lines,i + 1);
  Lines[i].A1 := A1;
  Lines[i].A2 := A2;
  Lines[i].B := B;
  Lines[i].Horiz := Horiz;
end;

begin
  bC1 := C1 >= FSheetArea.Col1;
  bR1 := R1 >= FSheetArea.Row1;
  bC2 := C2 <= FSheetArea.Col2;
  bR2 := R2 <= FSheetArea.Row2;

  ClipToSheet(C1,R1,C2,R2);

  Dec(C1,FSheetArea.Col1);
  Dec(C2,FSheetArea.Col1);
  Dec(R1,FSheetArea.Row1);
  Dec(R2,FSheetArea.Row1);

  if bC1 then bC1 := FColumns.VisibleSize[C1] > 0;
  if bR1 then bR1 := FRows.VisibleSize[R1] > 0;
  if bC2 then bC2 := FColumns.VisibleSize[C2] > 0;
  if bR2 then bR2 := FRows.VisibleSize[R2] > 0;

  x1 := -1;
  x2 := -1;
  Col := C1;
  while (Col <= C2) and (FColumns.VisibleSize[Col] <= 0) do
    Inc(Col);
  if Col <= C2 then begin
    x1 := FColumns.Pos[Col];
    Col := C2;
    while (Col >= C1) and (FColumns.VisibleSize[Col] <= 0) do
      Dec(Col);
    x2 := FColumns.Pos[Col] + FColumns.VisibleSize[Col];
  end;
  y1 := -1;
  y2 := -1;
  Row := R1;
  while (Row <= R2) and (FRows.VisibleSize[Row] <= 0) do
    Inc(Row);
  if Row <= R2 then begin
    y1 := FRows.Pos[Row];
    Row := R2;
    while (Row >= R1) and (FRows.VisibleSize[Row] <= 0) do
      Dec(Row);
    y2 := FRows.Pos[Row] + FRows.VisibleSize[Row];
  end;

  if bC1 then AddLine(y1,y2,x1,False);
  if bR1 then AddLine(x1,x2,y1,True);
  if bC2 then AddLine(y1,y2,x2,False);
  if bR2 then AddLine(x1,x2,y2,True);
end;

function TXLSBookMoveSelArea.Mode: TSelEditMode;
begin
  Result := semMove;
end;

procedure TXLSBookMoveSelArea.PaintRect;
var
  Lines: TStraightLineArray;
begin
  with FSelectedArea do begin
    MakeMoveRect(Col1 + FDeltaCol,Row1 + FDeltaRow,Col2 + FDeltaCol,Row2 + FDeltaRow,Lines);
    FSkin.PaintMoveRect(Lines);
  end;
end;

procedure TXLSBookMoveSelArea.ScrollUpdate(Col, Row: integer);
begin
  Inc(FDeltaCol,Col);
  Inc(FDeltaRow,Row);
  ClipMoveRect;
  with FSelectedArea do begin
    if (Row <> 0) and ((Row1 + FDeltaRow) >= FSheetArea.Row2) then
      FDeltaRow := FSheetArea.Row2 - Row1;
    if (Col <> 0) and ((Col1 + FDeltaCol) >= FSheetArea.Col2) then
      FDeltaCol := FSheetArea.Col2 - Col1;
  end;
end;


procedure TXLSBookMoveSelArea.UpdateEdit(Col, Row: integer; ClipToSheet: boolean);
var
  Lines: TStraightLineArray;
begin
  if (Col <> FLastCol) or (Row <> FLastRow) then begin
    PaintRect;
    SetLength(Lines,0);
    FDeltaCol := Col - FMouseDownCol;
    FDeltaRow := Row - FMouseDownRow;
    ClipMoveRect;
    with FSelectedArea do begin
      if (Row1 + FDeltaRow) >= FSheetArea.Row2 then
        FDeltaRow := FSheetArea.Row2 - Row1 - 1;
      if (Col1 + FDeltaCol) >= FSheetArea.Col2 then
        FDeltaCol := FSheetArea.Col2 - Col1 - 1;
      MakeMoveRect(Col1 + FDeltaCol,Row1 + FDeltaRow,Col2 + FDeltaCol,Row2 + FDeltaRow,Lines);
    end;
    FSkin.PaintMoveRect(Lines);
  end;
  FLastCol := Col;
  FLastRow := Row;
end;

{ TXLSBookSizeSelArea }

procedure TXLSBookSizeSelArea.BeginEdit(SelectedArea: TXLSSelectedArea; Col, Row: integer);
var
  Lines: TStraightLineArray;
begin
  FSelectedArea := SelectedArea;
  FStarted := True;
  FSizeArea := FSelectedArea.AsRecArea;
  if not FIsDelete then
    SetSizeArea(Col,Row);
  MakeSizeRect(FSizeArea.Col1,FSizeArea.Row1,FSizeArea.Col2,FSizeArea.Row2,Lines);
  FSkin.PaintMoveRect(Lines);
  if FIsDelete then begin
    FDeleteArea := FSizeArea;
    FDeleteArea.Col2 := Col;
    FDeleteArea.Row2 := Row;
  end;
  FLastCol := Col;
  FLastRow := Row;
end;

procedure TXLSBookSizeSelArea.DoDeleteArea;
var
  x1,y1,x2,y2: integer;
  C1,R1,C2,R2: integer;
begin
  with FDeleteArea do begin
    C1 := Col1;
    R1 := Row1;
    C2 := FSizeArea.Col2;
    R2 := FSizeArea.Row2;
    ClipToSheet(C1,R1,C2,R2);
    Dec(C1,FSheetArea.Col1);
    Dec(R1,FSheetArea.Row1);
    Dec(C2,FSheetArea.Col1);
    Dec(R2,FSheetArea.Row1);
    x1 := FColumns.Pos[C1] + 1;
    y1 := FRows.Pos[R1] + 1;
    x2 := FColumns.Pos[C2] + FColumns.VisibleSize[C2] - 1;
    y2 := FRows.Pos[R2] + FRows.VisibleSize[R2] - 1;
    FSkin.PaintDeleteRect(x1,y1,x2,y2);
  end;
end;

procedure TXLSBookSizeSelArea.EndEdit(Col, Row: integer);
begin
  PaintRect;
end;

procedure TXLSBookSizeSelArea.MakeSizeRect(C1, R1, C2, R2: integer; var Lines: TStraightLineArray);
var
  x1,y1,x2,y2: integer;
  oC1,oR1,oC2,oR2: integer;

procedure AddLine(A1,A2,B: integer; Horiz: boolean);
var
  i: integer;
begin
  SetLength(Lines,Length(Lines) + 1);
  i := High(Lines);
  Lines[i].A1 := A1;
  Lines[i].A2 := A2;
  Lines[i].B := B;
  Lines[i].Horiz := Horiz;
end;

begin
  oC1 := C1; oR1 := R1; oC2 := C2; oR2 := R2;
  ClipToSheet(C1,R1,C2,R2);
  Dec(C1,FSheetArea.Col1);
  Dec(R1,FSheetArea.Row1);
  Dec(C2,FSheetArea.Col1);
  Dec(R2,FSheetArea.Row1);
  x1 := FColumns.Pos[C1];
  y1 := FRows.Pos[R1];
  x2 := FColumns.Pos[C2] + FColumns.VisibleSize[C2];
  y2 := FRows.Pos[R2] + FRows.VisibleSize[R2];

  if oC1 >= FSheetArea.Col1 then
    AddLine(y1,y2,x1,False);
  if oR1 >= FSheetArea.Row1 then
    AddLine(x1,x2,y1,True);
  if oC2 <= FSheetArea.Col2 then
    AddLine(y1,y2,x2,False);
  if oR2 <= FSheetArea.Row2 then
    AddLine(x1,x2,y2,True);
end;

function TXLSBookSizeSelArea.Mode: TSelEditMode;
begin
  Result := semSize;
end;

procedure TXLSBookSizeSelArea.PaintRect;
var
  Lines: TStraightLineArray;
begin
  SetLength(Lines,4);
  MakeSizeRect(FSizeArea.Col1,FSizeArea.Row1,FSizeArea.Col2,FSizeArea.Row2,Lines);
  FSkin.PaintMoveRect(Lines);
end;

procedure TXLSBookSizeSelArea.ScrollUpdate(Col, Row: integer);
begin
  case Row of
    -1: Dec(FSizeArea.Row1);
     1: Inc(FSizeArea.Row2);
  end;
  case Col of
    -1: Dec(FSizeArea.Col1);
     1: Inc(FSizeArea.Col2);
  end;
  if (Row <> 0) and (FSizeArea.Row2 >= FSheetArea.Row2) then
    FSizeArea.Row2 := FSheetArea.Row2;
  if (Col <> 0) and (FSizeArea.Col2 >= FSheetArea.Col2) then
    FSizeArea.Col2 := FSheetArea.Col2;

  with FSelectedArea do begin
    if FSizeArea.Row1 < Row1 then
      FSizeArea.Row2 := Row2
    else if FSizeArea.Row2 < Row2 then
      FSizeArea.Row1 := Row1;
    if FSizeArea.Col1 < Col1 then
      FSizeArea.Col2 := Col2
    else if FSizeArea.Col2 > Col2 then
      FSizeArea.Col1 := Col1;
  end;
end;

procedure TXLSBookSizeSelArea.SetDeleteArea(Col, Row: integer);
begin
  FSizeArea.Col2 := Col;
  FSizeArea.Row2 := Row;
end;

procedure TXLSBookSizeSelArea.SetSizeArea(Col, Row: integer);
begin
  with FSelectedArea do begin
         if Col > Col2 then FSizeArea.Col2 := Col
    else if Col < Col1 then FSizeArea.Col1 := Col;
         if Row > Row2 then FSizeArea.Row2 := Row
    else if Row < Row1 then FSizeArea.Row1 := Row;
  end;
end;

procedure TXLSBookSizeSelArea.UpdateEdit(Col, Row: integer; ClipToSheet: boolean);
var
  Lines: TStraightLineArray;
begin
  if (Col <> FLastCol) or (Row <> FLastRow) then begin
    with FSelectedArea do begin
      MakeSizeRect(FSizeArea.Col1,FSizeArea.Row1,FSizeArea.Col2,FSizeArea.Row2,Lines);
      FSkin.PaintMoveRect(Lines);

      FSizeArea := FSelectedArea.AsRecArea;

      if FLastWasDelete then
        DoDeleteArea;

      if not FIsDelete then begin
        SetSizeArea(Col,Row);
        if ClipToSheet then begin
          if (Row <> 0) and (FSizeArea.Row2 >= FSheetArea.Row2) then
            FSizeArea.Row2 := FSheetArea.Row2 - 1;
          if (Col <> 0) and (FSizeArea.Col2 >= FSheetArea.Col2) then
            FSizeArea.Col2 := FSheetArea.Col2 - 1;
        end;
      end;

      SetLength(Lines,0);

      MakeSizeRect(FSizeArea.Col1,FSizeArea.Row1,FSizeArea.Col2,FSizeArea.Row2,Lines);
      FSkin.PaintMoveRect(Lines);


      if FIsDelete then begin
        FDeleteArea := FSizeArea;
        FDeleteArea.Col1 := Col;
        FDeleteArea.Row1 := Row;
        DoDeleteArea;
      end;
      FLastWasDelete := FIsDelete;

    end;
  end;
  FLastCol := Col;
  FLastRow := Row;
end;

end.
