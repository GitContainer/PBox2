unit XSSIEPhyRow;

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

uses { Delphi } Classes, SysUtils, Contnrs, Math,
     { AxWord } XSSIEUtils, XSSIEDocProps, XSSIECharRun;

type TAXWPhyRowFlag = (aprfDirty,aprfComplex,aprfFirst,aprfLast,aprfDebug);
     TAXWPhyRowFlags = set of TAXWPhyRowFlag;

type TAXWPhyRow = class(TObject)
protected
     FHeight : double;
     FWidth  : double;
     FXOffset: double;
     FY      : double;
     FAscent : double;
     FDescent: double;
     FCR1    : TAXWCharRun;
     FCR2    : TAXWCharRun;
     FCRPos1 : integer;
     FCRPos2 : integer;
     FLogPos1: integer;
     FFlags  : TAXWPhyRowFlags;
     FBreakCharCount: integer;
public
     constructor Create;

     procedure Assign(Row: TAXWPhyRow);
     function  Empty: boolean;

     function  RowHeight: double; {$ifdef D2006PLUS} inline; {$endif}
     function  Len: integer;
     function  SameCR: boolean;
     function  FirstPos: integer;
     function  LastPos: integer;

     property Height: double read FHeight write FHeight;
     property Ascent: double read FAscent write FAscent;
     property Descent: double read FDescent write FDescent;
     property Width: double read FWidth write FWidth;
     property XOffset: double read FXOffset write FXOffset;
     property Y: double read FY write FY;
     property BreakCharCount: integer read FBreakCharCount write FBreakCharCount;
     property CR1: TAXWCharRun read FCR1 write FCR1;
     property CR2: TAXWCharRun read FCR2 write FCR2;
     property CRPos1: integer read FCRPos1 write FCRPos1;
     // The last pos in the last CR (may be more than one).
     // _NOT_ the last pos of the entire CR text.
     property CRPos2: integer read FCRPos2 write FCRPos2;
     property LogPos1: integer read FLogPos1 write FLogPos1;
     property Flags: TAXWPhyRowFlags read FFlags write FFlags;
     end;

type TAXWPhyRowComplex = class(TAXWPhyRow)
private
     function  GetBreakWidth(Index: integer): TBreakCharData;
     function  GetBreakWidthsCount: integer;
protected
     FBreakWidths: array of TBreakCharData;
     FIteratePos: integer;
public
     constructor Create(Widths: array of TBreakCharData);

     procedure BeginIterate;
     function  IterateNext(CharCount: integer): double;

     property BreakWidth[Index: integer]: TBreakCharData read GetBreakWidth;
     property BreakWidthsCount: integer read GetBreakWidthsCount;
     end;

type TAXWPhyRows = class(TObjectList)
protected
     FWidth: double;
     FHeight: double;

     function  GetItems(Index: integer): TAXWPhyRow;
public
     constructor Create;

     procedure Clear; override;

     function  Add: TAXWPhyRow;
     function  FindRow(Pos: integer): integer;
     procedure Replace(Index: integer; Row: TAXWPhyRow);
     procedure ClearFirstLast;
     procedure SetFirstLast;
     procedure UpdateArea;
     function  Last: TAXWPhyRow;

     property Width: double read FWidth;
     property Height: double read FHeight;
     property Items[Index: integer]: TAXWPhyRow read GetItems; default;
     end;

implementation

{ TAXWPhyRow }

procedure TAXWPhyRow.Assign(Row: TAXWPhyRow);
begin
  FHeight := Row.FHeight;
  FWidth := Row.FWidth;
  FDescent := Row.FDescent;
  FCR1 := Row.FCR1;
  FCR2 := Row.FCR2;
  FCRPos1 := Row.FCRPos1;
  FCRPos2 := Row.FCRPos2;
  FLogPos1 := Row.FLogPos1;
  FFlags := Row.FFlags;
end;

constructor TAXWPhyRow.Create;
begin
  FFlags := [aprfDirty,aprfDebug];
end;

function TAXWPhyRow.Empty: boolean;
begin
  Result := FCR1 = Nil;
end;

function TAXWPhyRow.FirstPos: integer;
begin
  Result := FCR1.StartPos + FCRPos1 - 1;
end;

function TAXWPhyRow.LastPos: integer;
begin
  Result := FCR2.StartPos + FCRPos2 - 1;
end;

function TAXWPhyRow.Len: integer;
var
  i: integer;
begin
  if FCR1 = FCR2 then
    Result := FCRPos2 - FCRPos1 + 1
  else begin
    Result := FCR1.Len - FCRPos1 + 1;
    for i := FCR1.Index + 1 to FCR2.Index - 1 do
      Inc(Result,FCR1.Parent[i].Len);
    Inc(Result,FCRPos2);
  end;
end;

function TAXWPhyRow.RowHeight: double;
begin
  Result := FHeight + FDescent;
end;

function TAXWPhyRow.SameCR: boolean;
begin
  Result := FCR1 = FCR2;
end;

{ TAXWPhyRows }

function TAXWPhyRows.Add: TAXWPhyRow;
begin
  Result := TAXWPhyRow.Create;
  inherited Add(Result);
end;

procedure TAXWPhyRows.Clear;
begin
  inherited Clear;

  FWidth := 0;
  FHeight := 0;
end;

procedure TAXWPhyRows.ClearFirstLast;
begin
  if Count > 0 then begin
    Items[0].FFlags := Items[0].FFlags - [aprfFirst];
    Items[Count - 1].FFlags := Items[Count - 1].FFlags - [aprfLast];
  end;
end;

constructor TAXWPhyRows.Create;
begin
  inherited;
end;

function TAXWPhyRows.FindRow(Pos: integer): integer;
var
  First: Integer;
  Last: Integer;
  LP: integer;
begin
  if Count < 0 then
    Result := -1
  else begin
    First  := 0;
    Last   := Count - 2;

    while First <= Last do begin
      Result := (First + Last) div 2;
      if (Pos >= Items[Result].LogPos1) and (Pos < Items[Result + 1].LogPos1) then
        Exit
      else if Items[Result].LogPos1 > Pos then
        Last := Result - 1
      else
        First := Result + 1;
    end;
    LP := Items[Count - 1].LogPos1;
    if (Pos >= LP) and (Pos <= LP + Items[Count - 1].Len) then
      Result := Count - 1
    else
      Result := -1;
  end;
end;

function TAXWPhyRows.GetItems(Index: integer): TAXWPhyRow;
begin
  Result := TAXWPhyRow(inherited Items[Index]);
end;

function TAXWPhyRows.Last: TAXWPhyRow;
begin
  Result := TAXWPhyRow(inherited Last);
end;

procedure TAXWPhyRows.Replace(Index: integer; Row: TAXWPhyRow);
begin
  inherited Items[Index] := Row;
end;

procedure TAXWPhyRows.SetFirstLast;
begin
  if Count > 0 then begin
    Items[0].FFlags := Items[0].FFlags + [aprfFirst];
    Items[Count - 1].FFlags := Items[Count - 1].FFlags + [aprfLast];
  end;
end;

procedure TAXWPhyRows.UpdateArea;
begin
  FWidth := Max(FWidth,TAXWPhyRow(Last).FWidth);
  FHeight := FHeight + TAXWPhyRow(Last).FHeight;
end;

{ TAXWPhyRowComplex }

procedure TAXWPhyRowComplex.BeginIterate;
begin
  FIteratePos := 0;
end;

constructor TAXWPhyRowComplex.Create(Widths: array of TBreakCharData);
var
  i: integer;
  BreakCharsWidth: double;
begin
  inherited Create;
  SetLength(FBreakWidths,Length(Widths));
  System.Move(Widths[0],FBreakWidths[0],SizeOf(TBreakCharData) * Length(FBreakWidths));

  BreakCharsWidth := 0;
  for i := 0 to High(FBreakWidths) do
    BreakCharsWidth := BreakCharsWidth + FBreakWidths[i].Width;
  for i := 0 to High(FBreakWidths) do
    FBreakWidths[i].Percent := FBreakWidths[i].Width / BreakCharsWidth;
end;

function TAXWPhyRowComplex.GetBreakWidth(Index: integer): TBreakCharData;
begin
  Result := FBreakWidths[Index];
end;

function TAXWPhyRowComplex.GetBreakWidthsCount: integer;
begin
  Result := Length(FBreakWidths);
end;

function TAXWPhyRowComplex.IterateNext(CharCount: integer): double;
begin
  if FIteratePos > High(FBreakWidths) then
    raise Exception.Create('Row iterate out of range');
  Result := FBreakWidths[FIteratePos].Percent * CharCount;
  Inc(FIteratePos,CharCount);
end;

end.
