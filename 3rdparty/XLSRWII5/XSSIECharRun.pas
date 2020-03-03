unit XSSIECharRun;

{-
********************************************************************************
******* XLSSpreadSheet V2.00                                             *******
*******                                                                  *******
******* Copyright(C) 2006,2013 Lars Arvidsson, Axolot Data               *******
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
     { AxWord } XSSIEDefs, XSSIEUtils, XSSIEDocProps;

// acrfNoWhitespace is only valid if acrfPartOfNext is set.
type TAXWCharRunFlag = (acrfPartOfNext,acrfPartOfPrev,acrfNoWhitespace);
     TAXWCharRunFlags = set of TAXWCharRunFlag;

type TAXWCharRuns = class;

     TAXWCharRun = class(TAXWIndexObject)
private
     function  CPGetChar(Position: integer): AxUCChar;
     procedure SetText(Value: AxUCString);
     function  GetEndPos: integer;
     function  GetParent: TAXWCharRuns;
protected
     FStartPos  : integer;
     FFlags     : TAXWCharRunFlags;
     FCHPX      : TAXWCHPX;
     FText      : AxUCString;

     FFieldStart: integer;
     FFieldEnd  : integer;
public
     constructor Create(AParent: TAXWCharRuns; List: TAXWCHPX);
     destructor Destroy; override;

     function  Clone: TAXWCharRun; virtual;

     procedure Assign(CR: TAXWCharRun);
     procedure DeleteChars(const AIndex,ACount: integer);
     function  Len: integer; override;
     procedure Append(S: AxUCString);
     function  FirstChar: AxUCChar; {$ifdef D2006PLUS} inline; {$endif}
     function  LastChar: AxUCChar; {$ifdef D2006PLUS} inline; {$endif}
     function  Next: TAXWCharRun; {$ifdef D2006PLUS} inline; {$endif}
     function  Prev: TAXWCharRun; {$ifdef D2006PLUS} inline; {$endif}
     function  First: TAXWCharRun; {$ifdef D2006PLUS} inline; {$endif}
     function  Last: TAXWCharRun; {$ifdef D2006PLUS} inline; {$endif}
     function  AddCHPX: TAXWCHPX; {$ifdef D2006PLUS} inline; {$endif}
     function  ScanCharNext(var P: integer): boolean;
     procedure DeleteCHPX;
     procedure DeleteCHPXIfEmpty;
     procedure CheckPartOfNext;

     property Parent: TAXWCharRuns read GetParent;
     property Text: AxUCString read FText write SetText;
     property StartPos: integer read FStartPos write FStartPos;
     property EndPos: integer read GetEndPos;
     property Flags: TAXWCharRunFlags read FFlags write FFlags;
     property CHPX: TAXWCHPX read FCHPX;
     property CPChar[Position: integer]: AxUCChar read CPGetChar;
     property FieldStart: integer read FFieldStart write FFieldStart;
     property FieldEnd: integer read FFieldEnd write FFieldEnd;
     end;

     TAXWCharRuns = class(TAXWIndexObjectList)
private
     function GetItems(Index: integer): TAXWCharRun;
protected
     FCHP: TAXWCHP;

     FDeleteCREvent: TNotifyEvent;
public
     constructor Create(AOwner: TAXWIndexObject; ACHP: TAXWCHP);

     function  Add(List: TAXWCHPX): TAXWCharRun;
     function  Insert(const AIndex: integer; AList: TAXWCHPX): TAXWCharRun;

     function  Next(CR: TAXWCharRun): TAXWCharRun; {$ifdef D2006PLUS} inline; {$endif}
     function  Prev(CR: TAXWCharRun): TAXWCharRun; {$ifdef D2006PLUS} inline; {$endif}
     function  First: TAXWCharRun; {$ifdef D2006PLUS} inline; {$endif}
     function  Last: TAXWCharRun; {$ifdef D2006PLUS} inline; {$endif}
     function  Split(CR: TAXWCharRun; P: integer): TAXWCharRun;
     procedure Append(CRS: TAXWCharRuns);
     procedure Combine; overload;
     procedure Combine(CR1,CR2: TAXWCharRun); overload;
     function  IsEqualCHPX(CR1,CR2: TAXWCharRun): boolean;
     function  IsSplitWordPrev(CR: TAXWCharRun): boolean;
     function  IsSplitWordNext(CR: TAXWCharRun): boolean;
     function  Find(CP: integer): TAXWCharRun;
     function  FindBeginningOfWord(var P: integer; CR: TAXWCharRun): TAXWCharRun;
     function  FindEndOfWord(var P: integer; CR: TAXWCharRun): TAXWCharRun;
     function  PrevChar(var CR: TAXWCharRun; var P: integer; var C: AxUCChar): boolean;
     function  NextChar(var CR: TAXWCharRun; var P: integer; var C: AxUCChar): boolean;
     function  EndOfParaPos: integer; {$ifdef D2006PLUS} inline; {$endif}

     property OnDeleteCR: TNotifyEvent read FDeleteCREvent write FDeleteCREvent;
     property Items[Index: integer]: TAXWCharRun read GetItems; default;
     end;

implementation

{ TAXWCharRun }

procedure TAXWCharRun.Append(S: AxUCString);
begin
  FText := FText + S;
  Dirty := True;
end;

procedure TAXWCharRun.Assign(CR: TAXWCharRun);
begin
  FText := CR.FText;
  FStartPos := CR.StartPos;
  if CR.FCHPX <> Nil then begin
    FCHPX := TAXWCHPX.Create(TAXWCharRuns(FOwner).FCHP);
    FCHPX.Assign(CR.FCHPX);
  end
  else
    FCHPX := Nil;
  if Index > 0 then
    Prev.CheckPartOfNext;
end;

function TAXWCharRun.AddCHPX: TAXWCHPX;
begin
  if FCHPX = Nil then
    FCHPX := TAXWCHPX.Create(TAXWCharRuns(FOwner).FCHP);
  Result := FCHPX;
end;

procedure TAXWCharRun.CheckPartOfNext;
var
  CR: TAXWCharRun;
begin
  CR := Next;
  if CR = Nil then begin
    FFlags := FFlags - [acrfPartOfNext];
    FFlags := FFlags - [acrfNoWhitespace];
    Exit;
  end;
  if (FText <> '') and not (CharIsWhitespaceOrHardLB(FText[Length(FText)]) or ((CR.Text <> '') and CharIsWhitespaceOrHardLB(CR.Text[1]))) then begin
    FFlags := FFlags + [acrfPartOfNext];
    CR.Flags := CR.Flags + [acrfPartOfPrev];
    if WhitespacePos(FText) <= 0 then
      FFlags := FFlags + [acrfNoWhitespace]
    else
      FFlags := FFlags - [acrfNoWhitespace];
    CR.CheckPartOfNext;
  end
  else begin
    FFlags := FFlags - [acrfPartOfNext];
    FFlags := FFlags - [acrfNoWhitespace];
    CR.Flags := CR.Flags - [acrfPartOfPrev];
  end;
end;

function TAXWCharRun.Clone: TAXWCharRun;
begin
  Result := TAXWCharRun.Create(TAXWCharRuns(FOwner),FCHPX);
end;

constructor TAXWCharRun.Create(AParent: TAXWCharRuns; List: TAXWCHPX);
begin
  FOwner := AParent;
  if (List <> Nil) and (List.Count > 0) then begin
    FCHPX := TAXWCHPX.Create(TAXWCharRuns(FOwner).FCHP);
    FCHPX.Assign(List);
  end;
  FStartPos := 1;
  Dirty := True;
end;

procedure TAXWCharRun.DeleteChars(const AIndex, ACount: integer);
begin
  System.Delete(FText,AIndex,ACount);
  Dirty := True;
end;

procedure TAXWCharRun.DeleteCHPX;
begin
  if FCHPX <> Nil then
    FreeAndNil(FCHPX);
end;

procedure TAXWCharRun.DeleteCHPXIfEmpty;
begin
  if (FCHPX <> Nil) and (FCHPX.Count <= 0) then
    FreeAndNil(FCHPX);
end;

destructor TAXWCharRun.Destroy;
begin
  if Assigned(TAXWCharRuns(FOwner).FDeleteCREvent) then
    TAXWCharRuns(FOwner).FDeleteCREvent(Self);

  Dirty := True;
  DeleteCHPX;
  inherited;
end;

function TAXWCharRun.First: TAXWCharRun;
begin
  if FOwner.Count > 0 then
    Result := TAXWCharRun(FOwner.Items[0])
  else
    Result := Nil;
end;

function TAXWCharRun.FirstChar: AxUCChar;
begin
  if FText = '' then
    Result := AXW_CHAR_EOL
  else
    Result := FText[1];
end;

function TAXWCharRun.CPGetChar(Position: integer): AxUCChar;
var
  i: integer;
begin
  i := Position - FStartPos + 1;
  if (i < 1) or (i > Length(FText)) then
    Result := AXW_CHAR_EOL
  else
    Result := FText[i];
end;

function TAXWCharRun.GetEndPos: integer;
begin
  Result := FStartPos + Length(FText) - 1;
end;

function TAXWCharRun.GetParent: TAXWCharRuns;
begin
  Result := TAXWCharRuns(FOwner);
end;

function TAXWCharRun.Last: TAXWCharRun;
begin
  if FOwner.Count > 0 then
    Result := TAXWCharRun(FOwner.Items[FOwner.Count - 1])
  else
    Result := Nil;
end;

function TAXWCharRun.LastChar: AxUCChar;
begin
  if FText = '' then
    Result := AXW_CHAR_EOL
  else
    Result := FText[Length(FText)];
end;

function TAXWCharRun.Len: integer;
begin
  Result := Length(FText);
end;

function TAXWCharRun.Next: TAXWCharRun;
begin
  if Index < (FOwner.Count - 1) then
    Result := TAXWCharRun(FOwner[Index + 1])
  else
    Result := Nil;
end;

function TAXWCharRun.Prev: TAXWCharRun;
begin
  if Index > 0 then
    Result := TAXWCharRun(FOwner[Index - 1])
  else
    Result := Nil;
end;

function TAXWCharRun.ScanCharNext(var P: integer): boolean;
var
  L: integer;
  i: integer;
begin
  i := P - FStartPos + 1;
  L := Length(FText);
  if (i < 1) or (i > L) or (FText = '') then
    Result := False
  else begin
    while (i <= L) and not CharIsWhitespace(FText[i]) do
      Inc(i);
    while (i <= L) and CharIsWhitespace(FText[i]) do
      Inc(i);
    P := FStartPos + i - 1;
    Result := i <= L;
  end;
end;

procedure TAXWCharRun.SetText(Value: AxUCString);
begin
  FText := Value;
  if Index > 0 then
    Prev.CheckPartOfNext;
  CheckPartOfNext;
  Dirty := True;
end;

{ TAXWCharRuns }

function TAXWCharRuns.Add(List: TAXWCHPX): TAXWCharRun;
begin
  Result := TAXWCharRun.Create(Self,List);
  inherited Add(Result);
end;

procedure TAXWCharRuns.Append(CRS: TAXWCharRuns);
var
  i: integer;
  CR: TAXWCharRun;
begin
  for i := 0 to CRS.Count - 1 do begin
    CR := TAXWCharRun.Create(Self,Nil);
    CR.Assign(CRS[i]);
    inherited Add(CR);
  end;
end;

procedure TAXWCharRuns.Combine(CR1, CR2: TAXWCharRun);
var
  i1,i2: integer;
begin

  if (CR1 = Nil) and (CR2 = Nil) then
    Exit;
  if CR1 = Nil then
    CR1 := CR2
  else if CR2 = Nil then
    CR2 := CR1;

  i1 := Max(CR1.Index - 1,0);
  i2 := Min(CR2.Index + 1,Count - 1);
  while i1 < i2 do begin
    CR1 := Items[i1];
    CR2 := Items[i1 + 1];
    if IsEqualCHPX(CR1,CR2) then begin
      Items[i1].FText := Items[i1].FText + Items[i1 + 1].FText;
      Delete(i1 + 1);
      Items[i1].CheckPartOfNext;
      Dec(i2);
    end
    else
      Inc(i1);
  end;

end;

procedure TAXWCharRuns.Combine;
var
  i1,i2: integer;
  CR1,CR2: TAXWCharRun;
begin
  i1 := 0;
  i2 := Count - 1;
  while i1 < i2 do begin
    CR1 := Items[i1];
    CR2 := Items[i1 + 1];
    if IsEqualCHPX(CR1,CR2) then begin
      Items[i1].FText := CR1.FText + CR2.FText;
      Delete(i1 + 1);
      CR1.CheckPartOfNext;
      Dec(i2);
    end
    else
      Inc(i1);
  end;
end;

constructor TAXWCharRuns.Create(AOwner: TAXWIndexObject; ACHP: TAXWCHP);
begin
  inherited Create(AOwner);

  FCHP := ACHP;
end;

function TAXWCharRuns.EndOfParaPos: integer;
begin
  if Count < 1 then
    Result := 0
  else
    Result := Items[Count - 1].EndPos + 1;
end;

function TAXWCharRuns.Find(CP: integer): TAXWCharRun;
var
  First,Last,Pivot: Integer;
begin
  if CP = EndOfParaPos then begin
    Result := Items[Count - 1];
    Exit;
  end;

  First  := 0;
  Last   := Count - 1 ;

  while First <= Last do begin
    Pivot := (First + Last) div 2;
    Result := Items[Pivot];
    if (CP >= Result.StartPos) and (CP < (Result.StartPos + Result.Len)) then
      Exit;
    if Result.StartPos > CP then
      Last := Pivot - 1
    else
      First := Pivot + 1;
  end;
  Result := Nil;
end;

function TAXWCharRuns.FindBeginningOfWord(var P: integer; CR: TAXWCharRun): TAXWCharRun;
var
  i: integer;
  C: AxUCChar;
begin
  i := P - CR.StartPos + 1;
  if CharIsWhitespace(CR.Text[i]) then begin
    Result := Nil;
    Exit;
  end;
  Result := CR;

  while PrevChar(Result,P,C) do begin
    if CharIsWhitespace(C) then begin
      NextChar(Result,P,C);
      Exit;
    end;
  end;
  P := 1;
  Result := TAXWCharRun(First);
end;

function TAXWCharRuns.FindEndOfWord(var P: integer; CR: TAXWCharRun): TAXWCharRun;
var
  i: integer;
  C: AxUCChar;
begin
  i := P - CR.StartPos + 1;
  if CharIsWhitespace(CR.Text[i]) then begin
    Result := Nil;
    Exit;
  end;
  Result := CR;

  while NextChar(Result,P,C) do begin
    if CharIsWhitespace(C) then begin
//      PrevChar(Result,P,C);
      Exit;
    end;
  end;
  P := EndOfParaPos;
  Result := TAXWCharRun(Last);
end;

function TAXWCharRuns.First: TAXWCharRun;
begin
  if Count > 0 then
    Result := Items[0]
  else
    Result := Nil;
end;

function TAXWCharRuns.GetItems(Index: integer): TAXWCharRun;
begin
  Result := TAXWCharRun(inherited Items[Index]);
end;

function TAXWCharRuns.Insert(const AIndex: integer; AList: TAXWCHPX): TAXWCharRun;
begin
  Result := TAXWCharRun.Create(Self,AList);
  inherited Insert(AIndex,Result);
end;

function TAXWCharRuns.IsEqualCHPX(CR1, CR2: TAXWCharRun): boolean;
begin
  Result := (CR1.CHPX = Nil) and (CR2.CHPX = Nil);
  if Result then
    Exit;
  Result := ((CR1.CHPX <> Nil) and (CR2.CHPX <> Nil));
  if Result then
    Result := CR1.CHPX.Equal(CR2.CHPX);
end;

function TAXWCharRuns.IsSplitWordNext(CR: TAXWCharRun): boolean;
begin
  Result := CR.Index < (Count - 1);
  if Result then
    Result := not CharIsWhitespace(CR.LastChar) and not CharIsWhitespace(Items[CR.Index + 1].FirstChar);
end;

function TAXWCharRuns.IsSplitWordPrev(CR: TAXWCharRun): boolean;
begin
  Result := CR.Index > 0;
  if Result then
    Result := not CharIsWhitespace(CR.FirstChar) and not CharIsWhitespace(Items[CR.Index - 1].LastChar);
end;

function TAXWCharRuns.Last: TAXWCharRun;
begin
  if Count > 0 then
    Result := Items[Count - 1]
  else
    Result := Nil;
end;

function TAXWCharRuns.Next(CR: TAXWCharRun): TAXWCharRun;
var
  i: integer;
begin
 i := CR.Index + 1;
 if i < Count then
   Result := Items[i]
 else
   Result := Nil;
end;

function TAXWCharRuns.NextChar(var CR: TAXWCharRun; var P: integer; var C: AxUCChar): boolean;
var
  i: integer;
begin
  Inc(P);
  Result := P <= EndOfParaPos;
  if Result then begin
    repeat
      i := P - CR.StartPos + 1;
      if (i >= 1) and (i <= CR.Len) then begin
        C := CR.FText[i];
        Break;
      end
      else
        CR := Next(CR);
    until CR = Nil;
    Result := CR <> Nil;
  end;
end;

function TAXWCharRuns.Prev(CR: TAXWCharRun): TAXWCharRun;
var
  i: integer;
begin
 i := CR.Index - 1;
 if i >= 0 then
   Result := Items[i]
 else
   Result := Nil;
end;

function TAXWCharRuns.PrevChar(var CR: TAXWCharRun; var P: integer; var C: AxUCChar): boolean;
var
  i: integer;
begin
  Dec(P);
  Result := P >= 1;
  if Result then begin
    repeat
      i := P - CR.StartPos + 1;
      if (i >= 1) and (i <= CR.Len) then begin
        C := CR.FText[i];
        Break;
      end
      else
        CR := Prev(CR);
    until CR = Nil;
    Result := CR <> Nil;
  end;
end;

function TAXWCharRuns.Split(CR: TAXWCharRun; P: integer): TAXWCharRun;
var
  i: integer;
begin
  i := P - CR.StartPos + 1;
  if (i < 1) or (i > CR.Len) then begin
{$ifdef AXW_DEBUG}
    raise XLSRWException.Create('Invalid pos in Split');
{$endif}
    Result := Nil;
  end
  else begin
//    Result := TAXWCharRun.Create(Self,CR.CHPX);
    Result := CR.Clone;
    Result.FText := Copy(CR.FText,i,MAXINT);
    CR.FText := Copy(CR.FText,1,i - 1);
    inherited Insert(CR.Index + 1,Result);
    CR.CheckPartOfNext;
  end;
end;

end.
