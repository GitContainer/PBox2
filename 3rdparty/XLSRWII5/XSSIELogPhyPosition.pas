unit XSSIELogPhyPosition;

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

uses Classes, SysUtils, Contnrs, Math,
     XLSUtils5,
     XSSIEUtils, XSSIEDocProps, XSSIECharRun, XSSIELogParas, XSSIEPhyRow;

type TAXWLogPhyPosition = class(TAXWLogPosition)
protected
     FParaChanged: boolean;
     FRowIndex: integer;
     FRow: TAXWPhyRow;

     function  FindRow(P: integer): integer;
     function  BOLPos(R: integer): integer; overload; {$ifdef D2006PLUS} inline; {$endif}
     function  EOLPos(R: integer): integer; overload; {$ifdef D2006PLUS} inline; {$endif}
     function  EOLPlusPos(R: integer): integer; overload; {$ifdef D2006PLUS} inline; {$endif}
     procedure SetPos(const Value: integer); override;
     procedure SetRowFromPos;
public
     constructor Create;

     procedure Clear;

     procedure SyncRow;
     function  Valid: boolean;
     procedure MakeValid;
     procedure Reset(Para: TAXWLogPara); virtual;
     function  EqualRow(LPP: TAXWLogPhyPosition): boolean;
     function  GreaterRow(LPP: TAXWLogPhyPosition): boolean;
     function  LessRow(LPP: TAXWLogPhyPosition): boolean;
     procedure Assign(LPP: TAXWLogPhyPosition); overload;
     procedure Assign(APara: TAXWLogPara; ARowIndex: integer); overload;
     procedure Update; override;
     // Relative pos is the position from the beginning of the row on the
     // current row. First pos is 1.
     function  RelativePos: integer;
     function  BOLPos: integer; overload; // {$ifdef D2006PLUS} inline; {$endif}
     function  EOLPos: integer; overload; {$ifdef D2006PLUS} inline; {$endif}
     // Position after EOL (EOL + 1).
     function  EOLPlusPos: integer; overload; {$ifdef D2006PLUS} inline; {$endif}

     function  RowIsAtBeginningOfPara: boolean;
     function  RowIsAtEndOfPara: boolean;
     function  RowIsAtEndOfDoc: boolean;
     function  RowIsAtBeginningOfDoc: boolean;

     function  HaveNextRow: boolean;
     function  HavePrevRow: boolean;

     procedure MovePosNext; override;
     procedure MovePosPrev; override;
     function  MoveRowNext: boolean; virtual;
     function  MoveRowPrev: boolean; virtual;
     procedure MoveParaNext; virtual;
     procedure MoveParaPrev; virtual;
     procedure MoveToBOL; virtual;
     procedure MoveToEOLPlus; virtual;
     procedure MoveToEndOfPara; override;
     procedure MoveToBeginningOfPara; override;
     procedure MoveToEndOfDoc; override;
     procedure MoveToBeginningOfDoc; override;
     procedure MoveToPos(Pos: TAXWLogPhyPosition); overload; virtual;
     procedure MoveToPos(Pos: TAXWLogPosition); overload; virtual;

     property Row: TAXWPhyRow read FRow;
     property RowIndex: integer read FRowIndex;
     property ParaChanged: boolean read FParaChanged;
     end;

type TAXWEventLogPhyPosition = class(TAXWLogPhyPosition)
private
     procedure SetPosChangeEvent(const Value: TNotifyEvent);
protected
     FBeforePosChangeEvent: TNotifyEvent;
     FPosChangeEvent: TNotifyEvent;
     FOldPara: TAXWLogPara;
     FOldRowIndex: integer;
     FOldPos: integer;
public
     constructor Create;

     procedure Reset(Para: TAXWLogPara); override;
     procedure UpdatePosRelative(RelativePos: integer);

     procedure BeforeMove;
     procedure AfterMove;
     procedure AfterMoveKeepPos;

     procedure MovePosNext; override;
     procedure MovePosPrev; override;
     function  MoveRowNext: boolean; override;
     function  MoveRowPrev: boolean; override;
     procedure MoveParaNext; override;
     procedure MoveParaPrev; override;
     procedure MoveToBOL; override;
     procedure MoveToEOLPlus; override;
     procedure MoveToEndOfPara; override;
     procedure MoveToBeginningOfPara; override;
     procedure MoveToEndOfDoc; override;
     procedure MoveToBeginningOfDoc; override;
     procedure MoveToPos(Pos: TAXWLogPhyPosition); override;
     procedure MoveToPos(Pos: TAXWLogPosition); override;

     property OldPos: integer read FOldPos;

     property OnPosChange: TNotifyEvent read FPosChangeEvent write SetPosChangeEvent;
     end;

implementation

{ TAXWLogPhyPosition }

procedure TAXWLogPhyPosition.Assign(LPP: TAXWLogPhyPosition);
begin
  inherited Assign(LPP);
  FRowIndex := LPP.FRowIndex;
  FRow := LPP.FRow;
end;

function TAXWLogPhyPosition.BOLPos(R: integer): integer;
begin
  Result := FPara.Rows[R].LogPos1;
end;

procedure TAXWLogPhyPosition.Assign(APara: TAXWLogPara; ARowIndex: integer);
begin
  FPara := APara;
  FRowIndex := ARowIndex;
end;

function TAXWLogPhyPosition.BOLPos: integer;
begin
  if FRowIndex < 0 then
    Result := 0
  else
    Result := BOLPos(FRowIndex);
end;

procedure TAXWLogPhyPosition.Clear;
begin
  inherited Clear;
  FRowIndex := 0;
end;

constructor TAXWLogPhyPosition.Create;
begin
  inherited Create;

  Clear;
end;

function TAXWLogPhyPosition.EOLPos(R: integer): integer;
begin
  if R < (FPara.Rows.Count - 1) then
    Result := FPara.Rows[R + 1].LogPos1 - 1
  else
    Result := FPara.Rows[R].LogPos1 + FPara.Rows[R].Len - 1;
end;

function TAXWLogPhyPosition.EOLPlusPos: integer;
begin
  Result := EOLPlusPos(FRowIndex);
end;

function TAXWLogPhyPosition.EOLPlusPos(R: integer): integer;
begin
  Result := EOLPos(R);
  if R < (FPara.Rows.Count - 1) then
    Inc(Result);
end;

function TAXWLogPhyPosition.EOLPos: integer;
begin
  if FRowIndex < 0 then
    Result := 0
  else
    Result := EOLPos(FRowIndex);
end;

function TAXWLogPhyPosition.EqualRow(LPP: TAXWLogPhyPosition): boolean;
begin
  Result := (LPP.FPara = FPara) and (LPP.FRowIndex = FRowIndex);
end;

function TAXWLogPhyPosition.FindRow(P: integer): integer;
var
  First: Integer;
  Last: Integer;
  P1,P2: integer;
begin
  First  := 0;
  Last   := FPara.Rows.Count - 1 ;

  while First <= Last do begin
    Result := (First + Last) div 2;
    P1 := BOLPos(Result);
    P2 := EOLPos(Result);
    if (P >= P1) and (P <= P2) then
      Exit
    else if EOLPos(Result) > P then
      Last := Result - 1
    else
      First := Result + 1;
  end;
  if P = EndOfParaPos then
    Result := FPara.Rows.Count - 1
  else
    Result := -1;
end;

function TAXWLogPhyPosition.GreaterRow(LPP: TAXWLogPhyPosition): boolean;
begin
  Result := FPara.Index > LPP.Para.Index;
  if not Result and (FPara.Index = LPP.Para.Index) then
    Result := FRowIndex > LPP.FRowIndex;
end;

function TAXWLogPhyPosition.RowIsAtBeginningOfDoc: boolean;
begin
  Result := (FPara = FPara.First) and (FRowIndex = 0);
end;

function TAXWLogPhyPosition.RowIsAtBeginningOfPara: boolean;
begin
  Result := FRowIndex = 0;
end;

function TAXWLogPhyPosition.RowIsAtEndOfDoc: boolean;
begin
  Result := (FPara = FPara.Last) and (FRowIndex = (FPara.Rows.Count - 1));
end;

function TAXWLogPhyPosition.RowIsAtEndOfPara: boolean;
begin
  Result := FRowIndex = (FPara.Rows.Count - 1);
end;

function TAXWLogPhyPosition.LessRow(LPP: TAXWLogPhyPosition): boolean;
begin
  Result := FPara.Index < LPP.Para.Index;
  if not Result and (FPara.Index = LPP.Para.Index) then
    Result := FRowIndex < LPP.FRowIndex;
end;

procedure TAXWLogPhyPosition.MakeValid;
begin
  // TODO Check that para is valid?
  if FCharPos < 1 then
    FCharPos := 1
  else if FCharPos > FPara.EndOfParaPos then
    FCharPos := FPara.EndOfParaPos;
end;

procedure TAXWLogPhyPosition.MoveParaNext;
begin
  FParaChanged := False;
  if FPara.Next <> Nil then begin
    FPara := FPara.Next;
    if FPara.Rows.Count > 0 then
      FRowIndex := 0
    else
      FRowIndex := -1;
    FCharPos := 1;
    FParaChanged := True;
  end;
  SyncRow;
end;

procedure TAXWLogPhyPosition.MoveParaPrev;
begin
  FParaChanged := False;
  if FRowIndex = 0 then begin
    if FPara.Prev <> Nil then begin
      FPara := FPara.Prev;
      FParaChanged := True;
    end;
  end
  else
    FRowIndex := 0;
  FCharPos := 1;
  SyncRow;
end;

procedure TAXWLogPhyPosition.MovePosNext;
begin
  inherited MovePosNext;
  SetRowFromPos;
  SyncRow;
end;

procedure TAXWLogPhyPosition.MovePosPrev;
begin
  inherited MovePosPrev;
  SetRowFromPos;
  SyncRow;
end;

function TAXWLogPhyPosition.MoveRowNext: boolean;
begin
  Result := True;
  FParaChanged := False;
  if FRowIndex < (FPara.Rows.Count - 1) then begin
    Inc(FRowIndex);
    FCharPos := BOLPos;
  end
  else if FPara.Next <> Nil then begin
    FPara := FPara.Next;
    FRowIndex := 0;
    FCharPos := BOLPos;
    FParaChanged := True;
  end
  else
    Result := False;
  SyncRow;
end;

function TAXWLogPhyPosition.MoveRowPrev: boolean;
begin
  Result := True;
  FParaChanged := False;
  if FRowIndex > 0 then begin
    Dec(FRowIndex);
    SyncRow;
    FCharPos := BOLPos;
  end
  else if FPara.Prev <> Nil then begin
    FPara := FPara.Prev;
    FRowIndex := FPara.Rows.Count - 1;
    FCharPos := BOLPos;
    FParaChanged := True;
  end
  else
    Result := False;
  SyncRow;
end;

procedure TAXWLogPhyPosition.MoveToBeginningOfDoc;
begin
  inherited MoveToBeginningOfDoc;
  FRowIndex := 0;
  SyncRow;
end;

procedure TAXWLogPhyPosition.MoveToBeginningOfPara;
begin
  inherited MoveToBeginningOfPara;
  FRowIndex := 0;
  SyncRow;
end;

procedure TAXWLogPhyPosition.MoveToBOL;
begin
  FParaChanged := False;
  FCharPos := BOLPos;
  SyncRow;
end;

procedure TAXWLogPhyPosition.MoveToEndOfDoc;
begin
  inherited MoveToEndOfDoc;
  FRowIndex := FPara.Rows.Count - 1;
  SyncRow;
end;

procedure TAXWLogPhyPosition.MoveToEndOfPara;
begin
  inherited MoveToEndOfPara;
  FRowIndex := FPara.Rows.Count - 1;
  SyncRow;
end;

procedure TAXWLogPhyPosition.MoveToEOLPlus;
begin
  FParaChanged := False;
  if FRowIndex = (FPara.Rows.Count - 1) then begin
    FCharPos := EndOfParaPos;
    SetRowFromPos;
  end
  else begin
    FCharPos := EOLPos;
    SetRowFromPos;
    Inc(FCharPos);
  end;
  SyncRow;
end;

procedure TAXWLogPhyPosition.MoveToPos(Pos: TAXWLogPosition);
begin
  Assign(Pos);
  SetRowFromPos;
  SyncRow;
end;

procedure TAXWLogPhyPosition.MoveToPos(Pos: TAXWLogPhyPosition);
begin
  Assign(Pos);
  SetRowFromPos;
  SyncRow;
end;

function TAXWLogPhyPosition.HaveNextRow: boolean;
begin
  if (FRowIndex < (FPara.Rows.Count - 1)) or (FPara.Next <> Nil) then
    Result := True
  else
    Result := False;
end;

function TAXWLogPhyPosition.HavePrevRow: boolean;
begin
  if (FRowIndex > 0) or (FPara.Prev <> Nil) then
    Result := True
  else
    Result := False;
end;

function TAXWLogPhyPosition.RelativePos: integer;
var
  RS: integer;
begin
  RS := BOLPos;
  if RS = 0 then
    Result := 0
  else if (FCharPos >= RS) and (FCharPos <= EOLPos) then
    Result := FCharPos - RS + 1
  else if FCharPos = EOLPlusPos then
    Result := FCharPos - RS + 1
  else if FCharPos = EndOfParaPos then
    Result := FCharPos - RS + 1
  else
    raise XLSRWException.Create('Invalid row position');
end;

procedure TAXWLogPhyPosition.Reset(Para: TAXWLogPara);
begin
  // TODO Maybe more checks...
  // Don't call any events.
  if FPara = Nil then begin
    FPara := Para;
    FRowIndex := 0;
    FCharPos := 1;
  end;
end;

procedure TAXWLogPhyPosition.SetPos(const Value: integer);
begin
  FCharPos := Value;
  SetRowFromPos;
end;

procedure TAXWLogPhyPosition.Update;
begin
  inherited;
  SetRowFromPos;
end;

procedure TAXWLogPhyPosition.SetRowFromPos;
begin
  FRowIndex := FindRow(FCharPos);
  if FRowIndex < 0 then begin
    raise XLSRWException.CreateFmt('Can not find row. Pos=%d',[FCharPos]);
  end;
end;

procedure TAXWLogPhyPosition.SyncRow;
begin
  if (FPara <> Nil) and (FRowIndex <= FPara.Rows.Count - 1) then
    FRow := FPara.Rows[FRowIndex]
  else
    FRow := Nil;
end;

function TAXWLogPhyPosition.Valid: boolean;
begin
  Result := FPara <> Nil;
end;

{ TAXWLogPhyPositionEvent }

procedure TAXWEventLogPhyPosition.AfterMove;
begin
  if FRowIndex >= 0 then
    FRow := FPara.Rows[FRowIndex]
  else
    FRow := Nil;

  FPara.SetDocPropsAtPos(FCharPos);
//    if FOldPos < 1 then
//      FOldPos := FCharPos;
  if (FOldPos <> FCharPos) or (FOldPara <> FPara) or (FOldRowIndex <> FRowIndex) then
    FPosChangeEvent(Self);
end;

procedure TAXWEventLogPhyPosition.AfterMoveKeepPos;
begin
  AfterMove;
end;

procedure TAXWEventLogPhyPosition.BeforeMove;
begin
  FOldPara := FPara;
  FOldRowIndex := FRowIndex;
//  if FOldPos >= 1 then
    FOldPos := FCharPos;
end;

constructor TAXWEventLogPhyPosition.Create;
begin
  inherited Create;

  FOldPos := -1;
end;

procedure TAXWEventLogPhyPosition.MoveParaNext;
begin
  BeforeMove;
  inherited MoveParaNext;
  AfterMove;
end;

procedure TAXWEventLogPhyPosition.MoveParaPrev;
begin
  BeforeMove;
  inherited MoveParaPrev;
  AfterMove;
end;

procedure TAXWEventLogPhyPosition.MovePosNext;
begin
  BeforeMove;
  inherited MovePosNext;
  AfterMove;
end;

procedure TAXWEventLogPhyPosition.MovePosPrev;
begin
  BeforeMove;
  inherited MovePosPrev;
  AfterMove;
end;

function TAXWEventLogPhyPosition.MoveRowNext: boolean;
begin
  BeforeMove;
  Result := inherited MoveRowNext;
  AfterMoveKeepPos;
end;

function TAXWEventLogPhyPosition.MoveRowPrev: boolean;
begin
  BeforeMove;
  Result := inherited MoveRowPrev;
  AfterMove;
end;

procedure TAXWEventLogPhyPosition.MoveToBeginningOfDoc;
begin
  BeforeMove;
  inherited MoveToBeginningOfDoc;
  AfterMove;
end;

procedure TAXWEventLogPhyPosition.MoveToBeginningOfPara;
begin
  BeforeMove;
  inherited MoveToBeginningOfPara;
  AfterMove;
end;

procedure TAXWEventLogPhyPosition.MoveToBOL;
begin
  BeforeMove;
  inherited MoveToBOL;
  AfterMove;
end;

procedure TAXWEventLogPhyPosition.MoveToEndOfDoc;
begin
  BeforeMove;
  inherited MoveToEndOfDoc;
  AfterMove;
end;

procedure TAXWEventLogPhyPosition.MoveToEndOfPara;
begin
  BeforeMove;
  inherited MoveToEndOfPara;
  AfterMove;
end;

procedure TAXWEventLogPhyPosition.MoveToEOLPlus;
begin
  BeforeMove;
  inherited MoveToEOLPlus;
  AfterMove;
end;

procedure TAXWEventLogPhyPosition.MoveToPos(Pos: TAXWLogPosition);
begin
  BeforeMove;
  inherited MoveToPos(Pos);
  AfterMove;
end;

procedure TAXWEventLogPhyPosition.MoveToPos(Pos: TAXWLogPhyPosition);
begin
  BeforeMove;
  inherited MoveToPos(Pos);
  AfterMove;
end;

procedure TAXWEventLogPhyPosition.Reset(Para: TAXWLogPara);
begin
  inherited Reset(Para);
end;

procedure TAXWEventLogPhyPosition.SetPosChangeEvent(const Value: TNotifyEvent);
begin
  FPosChangeEvent := Value;
end;

procedure TAXWEventLogPhyPosition.UpdatePosRelative(RelativePos: integer);
begin
  FCharPos := BOLPos + RelativePos - 1;
end;

end.
