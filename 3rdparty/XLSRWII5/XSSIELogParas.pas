unit XSSIELogParas;

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

uses Classes, SysUtils, Contnrs, Math, Masks,
     XLSUtils5,
     XSSIEDefs, XSSIEUtils, XSSIEDocProps, XSSIECharRun, XSSIEPhyRow;

type TAXWSelectionHit = (axshNone,axshFirstOnRow,axshLastOnRow,axshBothOnRow,axshFullRow);

type TAXWLogParas = class;
     TAXWLogPosition = class;

     TAXWLogPara = class(TAXWIndexObject)
private
     function  GetItems(Index: integer): TAXWCharRun;
     function  GetText: AxUCString;
     function  GetParent: TAXWLogParas;
protected
     FCharRuns : TAXWCharRuns;
     FCHP      : TAXWCHP;
     FPAPX     : TAXWPAPX;

     FTabs     : TAXWTabStops;
     FNumText  : AxUCString;

     FRows     : TAXWPhyRows;

     function  FindRun(CP: integer): TAXWCharRun;
     procedure Sync(StartCR: TAXWCharRun);
     // Merges format with List
     procedure ApplyFormat(List: TAXWCHPX; CR1,CR2: TAXWCharRun); overload;
     procedure ApplyFormat(List: TAXWCHPX; CP1,CP2: integer; CR1,CR2: TAXWCharRun); overload;
     // Replaces format with List
     procedure SetFormat(List: TAXWCHPX; CR1,CR2: TAXWCharRun); overload;
     procedure SetFormat(List: TAXWCHPX; CP1,CP2: integer; CR1,CR2: TAXWCharRun); overload;
public
     constructor Create(Parent: TAXWLogParas; CHP: TAXWCHP);
     destructor Destroy; override;

     procedure Clear;
     procedure ClearDirty;
     function  Count: integer;

     function  Next: TAXWLogPara; {$ifdef D2006PLUS} inline; {$endif}
     function  Prev: TAXWLogPara; {$ifdef D2006PLUS} inline; {$endif}
     function  First: TAXWLogPara; {$ifdef D2006PLUS} inline; {$endif}
     function  Last: TAXWLogPara; {$ifdef D2006PLUS} inline; {$endif}
     function  Empty: boolean; {$ifdef D2006PLUS} inline; {$endif}
     function  Len: integer; override;
     function  CRFirst: TAXWCharRun; {$ifdef D2006PLUS} inline; {$endif}
     function  CRLast: TAXWCharRun; {$ifdef D2006PLUS} inline; {$endif}
     function  EndOfParaPos: integer; {$ifdef D2006PLUS} inline; {$endif}
     function  FindPrevWord(var CP: integer): boolean;
     function  FindNextWord(var CP: integer): boolean;
     function  WordAtPos(CP: integer; var CP1,CP2: integer): boolean; overload;
     function  WordAtPos(CP: integer; var P1,P2: integer; var CR1,CR2: TAXWCharRun): boolean; overload;
     function  CHPXAtPos(CP: integer): TAXWCHPX;
     procedure Append(Para: TAXWLogPara);
     function  AppendRaw(const S: AxUCString): TAXWCharRun;
     function  FindCharRun(CP: integer): TAXWCharRun; {$ifdef D2006PLUS} inline; {$endif}
     function  SplitCharRun(CR: TAXWCharRun; CP1,CP2: integer): TAXWCharRun; overload;
     function  SplitCharRun(CR: TAXWCharRun; CP: integer): TAXWCharRun; overload;

     procedure InsertText(S: AxUCString; P: integer; InCRLeft: boolean = False);
     procedure DeleteText(CP1,CP2: integer);
     procedure DeleteTextToEnd(CP: integer);
     procedure AppendText(List: TAXWCHPX; const S: AxUCString);
     // TAXWLogParas.Sync muste be called after AppendTextUncond. If not, CharRuns start index will be wrong.
     procedure AppendTextUncond(List: TAXWCHPX; const S: AxUCString);
     function  AsPlainText(CP1,CP2: integer): AxUCString;

     function  FormatWord(List: TAXWCHPX; CP: integer): boolean;
     procedure FormatAll(List: TAXWCHPX);
     procedure FormatRange(List: TAXWCHPX; CP1,CP2: integer);
     procedure FormatAllUncond(List: TAXWCHPX);
     procedure FormatRangeUncond(List: TAXWCHPX; CP1,CP2: integer);

     procedure AddTabs;
     function  MapTab(var APixelPos: double; AIndex: integer): TAXWTabStop;

     // Assumes that the text fits in one char run. Will not find text that is splitted over several char runs.
     function  FindText(const AText: AxUCString): integer;

     function  GetChar(P: integer): AxUCChar;
     procedure SetDocPropsAtPos(CP: integer);

     property Parent: TAXWLogParas read GetParent;
     property Text: AxUCString read GetText;
     property PAPX: TAXWPAPX read FPAPX;
     property CHP: TAXWCHP read FCHP;
     property Rows: TAXWPhyRows read FRows;
     property Runs: TAXWCharRuns read FCharRuns;
     property Tabs: TAXWTabStops read FTabs;
     property NumText: AxUCString read FNumText write FNumText;
     property Items[Index: integer]: TAXWCharRun read GetItems; default;
     end;

     TAXWLogParas = class(TAXWIndexObjectList)
private
     function GetItems(Index: integer): TAXWLogPara;
protected
     FPPIX         : integer;
     FPPIY         : integer;
     FDOP          : TAXWDOP;
     FMasterPAP    : TAXWPAP;
     FMasterCHP    : TAXWCHP;
     FCaretPAPX    : TAXWPAPX;
     FCaretCHPX    : TAXWCHPX;
public
     constructor Create(ADOP: TAXWDOP; AMasterPAP: TAXWPAP; AMasterCHP: TAXWCHP);
     destructor Destroy; override;

     function  CountCR: integer;

     procedure ClearRows;

     function  Empty: boolean;
     function  First: TAXWLogPara;
     function  Last: TAXWLogPara;
     function  New: TAXWLogPara;
     function  Add: TAXWLogPara; overload;
     procedure Add(Para: TAXWLogPara); overload;
     function  Insert(Index: integer): TAXWLogPara;
     procedure MergeNext(Index: integer);

     procedure AddRawText(S: AxUCString);
     procedure Sync;

     procedure ApplyFormat(CHPX: TAXWCHPX; Pos1,Pos2: TAXWLogPosition);
     procedure ApplyFormatUncond(CHPX: TAXWCHPX; Pos1,Pos2: TAXWLogPosition);
     procedure ToggleFormat(Id: TAXWChpId; Pos1,Pos2: TAXWLogPosition);

     procedure ApplyParaFormat(PAPX: TAXWPAPX; Para1,Para2: TAXWLogPara);

     property PPIX: integer read FPPIX write FPPIX;
     property PPIY: integer read FPPIY write FPPIY;
     property DOP: TAXWDOP read FDOP;
     property MasterPAP: TAXWPAP read FMasterPAP;
     property MasterCHP: TAXWCHP read FMasterCHP;
     property CaretPAPX: TAXWPAPX read FCaretPAPX;
     property CaretCHPX: TAXWCHPX read FCaretCHPX;
     property Items[Index: integer]: TAXWLogPara read GetItems; default;
     end;

     TAXWLogPosition = class(TObject)
protected
     FPara: TAXWLogPara;
     FCharPos: integer;
     FParaChanged: boolean;

     procedure SetPos(const Value: integer); virtual;
     procedure IncPos; {$ifdef D2006PLUS} inline; {$endif}
     procedure DecPos; {$ifdef D2006PLUS} inline; {$endif}
public
     constructor Create;
     procedure Clear;
     function  Equal(LPP: TAXWLogPosition): boolean;
     function  Greater(LPP: TAXWLogPosition): boolean;
     function  GreaterEqual(LPP: TAXWLogPosition): boolean;
     function  Less(LPP: TAXWLogPosition): boolean;
     function  LessEqual(LPP: TAXWLogPosition): boolean;
     function  CPIsAtEOP: boolean; {$ifdef D2006PLUS} inline; {$endif}
     function  CPIsAtBOP: boolean; {$ifdef D2006PLUS} inline; {$endif}
     procedure Assign(LPP: TAXWLogPosition); overload;
     procedure Assign(APara: TAXWLogPara); overload;
     procedure SetWithinLimits;
     procedure Update; virtual;
     procedure SetDocPropsAtPos;

     function  EndOfParaPos: integer; virtual;
     // Used for selected text. The (caret) positions in Selections are one
     // char more than the selected text. CharPosDec will decrese char pos
     // by one, assuming char pos is greater than zero.
     function  CharPosDec: integer;

     procedure MovePosNext; virtual;
     procedure MovePosPrev; virtual;
     procedure MoveToEndOfPara; virtual;
     procedure MoveToBeginningOfPara; virtual;
     procedure MoveToEndOfDoc; virtual;
     procedure MoveToBeginningOfDoc; virtual;

     property Para: TAXWLogPara read FPara write FPara;
     property CharPos: integer read FCharPos write SetPos;
     end;

type TAXWLogSelections = class;

     TAXWLogSelection = class(TObject)
protected
     FParent: TAXWLogSelections;
     FFirst,FLast: TAXWLogPosition;
public
     constructor Create(Parent: TAXWLogSelections);
     destructor Destroy; override;

     function  Hit(Para: TAXWLogPara; Row: TAXWPhyRow): TAXWSelectionHit;
     function  AsPlainText: AxUCString;

     property First: TAXWLogPosition read FFirst;
     property Last: TAXWLogPosition read FLast;
     end;

     TAXWLogSelections = class(TObjectList)
private
     function GetItems(Index: integer): TAXWLogSelection;
protected
     // Only valid after Hit is called.
     FHitResult: TAXWSelectionHit;
     // Char pos from beginning of the row.
     FHitP1Offs,FHitP2Offs: integer;
public
     constructor Create;
     procedure Reset; virtual;
     procedure Add;
     function  Hit(Para: TAXWLogPara; Row: TAXWPhyRow): TAXWSelectionHit;

     property HitResult: TAXWSelectionHit read FHitResult;
     property HitP1Offs: integer read FHitP1Offs;
     property HitP2Offs: integer read FHitP2Offs;
     property Items[Index: integer]: TAXWLogSelection read GetItems; default;
     end;

type TAXWFindTextData = class(TObject)
private
     function GetFound: boolean;
protected
     FCharPos: integer;
     FCharRun: TAXWCharRun;
     FPara: TAXWLogPara;
     FOptions: TAXWFindTextOption;
     FText: AxUCString;
     FAtEndOfDoc: boolean;
public
     procedure BeginFind(Para: TAXWLogPara);
     function  FindNext: boolean;

     property Options: TAXWFindTextOption read FOptions write FOptions;
     property Text: AxUCString read FText write FText;
     property Found: boolean read GetFound;
     end;

implementation

{ TXSSLogPara }

function TAXWLogPara.AppendRaw(const S: AxUCString): TAXWCharRun;
begin
  Result := FCharRuns.Add(Nil);
//  Result.CheckCHPX;
  Result.Text := S;
end;

procedure TAXWLogPara.AppendText(List: TAXWCHPX; const S: AxUCString);
var
  CR: TAXWCharRun;
  ListsEqual: boolean;
begin
  if Count > 0 then begin
    CR := Items[Count - 1];
    ListsEqual := (List <> Nil) and (CR.CHPX <> Nil) and CR.CHPX.Equal(List);
    if not ListsEqual then
      ListsEqual := not (((List <> Nil) and (List.Count > 0)) or ((CR.CHPX <> Nil) and (CR.CHPX.Count > 0)));
  end
  else
    ListsEqual := False;
  if (Count > 0) and ListsEqual then
    Items[Count - 1].Append(S)
  else begin
    CR := FCharRuns.Add(List);
    CR.Text := S;
  end;
  if Count > 1 then
    Items[Count - 1].StartPos := Items[Count - 2].EndPos + 1;
end;

procedure TAXWLogPara.AppendTextUncond(List: TAXWCHPX; const S: AxUCString);
var
  CR: TAXWCharRun;
begin
  CR := FCharRuns.Add(List);
  CR.Text := S;
end;

procedure TAXWLogPara.ApplyFormat(List: TAXWCHPX; CR1, CR2: TAXWCharRun);
var
  i: integer;
  CR: TAXWCharRun;
begin
  if (List = Nil) or (List.Count = 0) then
    Exit;
  for i := CR1.Index to CR2.Index do begin
    CR := FCharRuns[i];
    CR.AddCHPX;
    CR.CHPX.Merge(List);
    FCHP.CompactCHPX(CR.CHPX);
    if FCHP.IsEmptyCHPX(CR.CHPX) then
      CR.DeleteCHPX;

    CR.Dirty := True;
  end;
  FCharRuns.Combine(CR1,CR2);
end;

procedure TAXWLogPara.ApplyFormat(List: TAXWCHPX; CP1, CP2: integer; CR1, CR2: TAXWCharRun);
var
  i: integer;
  CR1b, CR2b: TAXWCharRun;

procedure ApplySingle(CRun: TAXWCharRun; Pos1,Pos2: integer);
var
  CR: TAXWCharRun;
begin
  CR := SplitCharRun(CRun,Pos1,Pos2);
  CR.AddCHPX;
  CR.CHPX.Merge(List);
  FCHP.CompactCHPX(CR.CHPX);
  if FCHP.IsEmptyCHPX(CR.CHPX) then
    CR.DeleteCHPX;
  CR.Dirty := True;
end;

begin
  { TODO Check that there is no empty char runs left after split (possibly with only spaces)}
  if CR1 = CR2 then begin
    ApplySingle(CR1,CP1,CP2);
  end
  else begin
    if CP1 > 1 then begin
      ApplySingle(CR1,CP1,CR1.EndPos);
      CR1b := FCharRuns.Next(CR1);
    end
    else
      CR1b := CR1;
    if CP2 < CR2.EndPos then begin
      ApplySingle(CR2,1,CP2);
      CR2b := FCharRuns.Prev(CR2);
    end
    else
      CR2b := CR2;
    for i := CR1b.Index to CR2b.Index do
      ApplySingle(FCharRuns[i],1,FCharRuns[i].EndPos);
  end;
  FCharRuns.Combine(CR1,CR2);

  Sync(CR1);
end;

procedure TAXWLogPara.AddTabs;
begin
  if FTabs = Nil then
    FTabs := TAXWTabStops.Create;
end;

procedure TAXWLogPara.Append(Para: TAXWLogPara);
var
  CR1,CR2: TAXWCharRun;
begin
  { TODO Sync CHPX with this para? }

  CR1 := CRLast;
  CR2 := Para.CRFirst;
  FCharRuns.Append(Para.FCharRuns);
  FCharRuns.Combine(CR1,CR2);
  Sync(CR1);
end;

function TAXWLogPara.CHPXAtPos(CP: integer): TAXWCHPX;
var
  CR: TAXWCharRun;
begin
  Result := Nil;
  CR := FindRun(CP);
  if CR <> Nil then
    Result := CR.CHPX;
end;

procedure TAXWLogPara.Clear;
begin
  FCharRuns.Clear;
  if FTabs <> Nil then
    FreeAndNil(FTabs);
end;

procedure TAXWLogPara.ClearDirty;
var
  i: integer;
begin
  FCharRuns.Dirty := False;
  for i := 0 to FCharRuns.Count - 1 do begin
    FCharRuns[i].Dirty := False;
  end;
end;

function TAXWLogPara.AsPlainText(CP1,CP2: integer): AxUCString;
var
  i: integer;
  CR1,CR2: TAXWCharRun;
begin
  CR1 := FindRun(CP1);
  CR2 := FindRun(CP2);
  if (CR1 <> Nil) and (CR2 <> NiL) then begin
    if CR1 = CR2 then
      Result := Copy(CR1.Text,CP1,CP2 - CP1 + 1)
    else begin
      Result := Copy(CR1.Text,CP1,MAXINT);
      for i := CR1.Index + 1 to CR2.Index - 1 do
        Result := Result + Items[i].Text;
      Result := Result + Copy(CR2.Text,1,CP2);
    end;
  end;
end;

function TAXWLogPara.Count: integer;
begin
  Result := FCharRuns.Count;
end;

constructor TAXWLogPara.Create(Parent: TAXWLogParas; CHP: TAXWCHP);
begin
  inherited Create(Parent);
  FCHP := CHP;
  FPAPX := TAXWPAPX.Create(Parent.MasterPAP);
  FCharRuns := TAXWCharRuns.Create(Self,FCHP);
  FRows := TAXWPhyRows.Create;
end;

procedure TAXWLogPara.DeleteText(CP1,CP2: integer);
var
  i,Cnt: integer;
  DelStart,DelEnd: integer;
  CR,CR1,CR2: TAXWCharRun;
  DelAtStart: boolean;
begin
  CR1 := FindRun(CP1);
  if CP2 <> CP1 then
    CR2 := FindRun(CP2)
  else
    CR2 := CR1;

  if (CR1 = Nil) or (CR2 = Nil) then
    raise XLSRWException.Create('Delete index out or range');

  DelAtStart := CP1 = CR1.StartPos;
  i := CP1 - CR1.StartPos + 1;
  if CP1 = CP2 then
    Cnt := CP2 - CP1 + 1
  else
    Cnt := CP2 - CP1;
  if CR1 = CR2 then begin
    CR1.DeleteChars(i,Cnt);
    if CR1.Text = '' then begin
      i := CR1.Index;
      CR1 := CR1.Prev;
      DelAtStart := False;
      FCharRuns.Delete(i);
    end;
  end
  else begin
    DelStart := (CR1.Index + 1);
    DelEnd := (CR2.Index - 1);
    CR1.DeleteChars(i,MAXINT);
    if CR1.Text = '' then
      Dec(DelStart);
    CR2.DeleteChars(1,CP2 - CR2.StartPos);
    if CR2.Text = '' then
      Inc(DelEnd);

    for i := DelStart to DelEnd do
      FCharRuns.Delete(DelStart);

    if DelStart = 0 then
      CR1 := CRFirst;
  end;

  if CR1 <> Nil then
    FCharRuns.Combine(CR1,CR2);

  if DelAtStart then
    CR := CR1.Prev
  else
    CR := CR1;
  if CR <> Nil then
    CR.CheckPartOfNext;

  Sync(CR1);
end;

procedure TAXWLogPara.DeleteTextToEnd(CP: integer);
begin
  DeleteText(CP,Len + 1);
end;

destructor TAXWLogPara.Destroy;
begin
  FreeAndNil(FCharRuns);
  FreeAndNil(FPAPX);
  FreeAndNil(FRows);
  if FTabs <> Nil then
    FreeAndNil(FTabs);
  inherited;
end;

function TAXWLogPara.Empty: boolean;
begin
  Result := Count <= 0;
end;

function TAXWLogPara.EndOfParaPos: integer;
begin
  Result := FCharRuns.EndOfParaPos;
end;

function TAXWLogPara.FindCharRun(CP: integer): TAXWCharRun;
begin
  Result := FCharRuns.Find(CP);
end;

function TAXWLogPara.FindNextWord(var CP: integer): boolean;
var
  CR: TAXWCharRun;

function ScanForLast(CharType: TAXWCharType): boolean;
var
  C: AxUCChar;
begin
  Result := True;
  while Result and FCharRuns.NextChar(CR,CP,C) do begin
    case CharType of
      actWhitespace   : Result := CharIsWhitespace(C);
      actNotWhitespace: Result := not CharIsWhitespace(C);
    end;
  end;
end;

begin
  CR := FindRun(CP);
  if (CR = Nil) or (CP >= EndOfParaPos) then begin
    Result := False;
    Exit;
  end;

  ScanForLast(actNotWhitespace);
  if CR = Nil then begin
    if CP = EndOfParaPos then
      Result := True
    else
      Result := True;
  end
  else begin
    ScanForLast(actWhitespace);
    Result := True;
  end;
end;

function TAXWLogPara.FindPrevWord(var CP: integer): boolean;
var
  CR: TAXWCharRun;

function ScanForFirst(CharType: TAXWCharType): boolean;
var
  C: AxUCChar;
begin
  Result := True;
  while Result and FCharRuns.PrevChar(CR,CP,C) do begin
    case CharType of
      actWhitespace   : Result := CharIsWhitespace(C);
      actNotWhitespace: Result := not CharIsWhitespace(C);
    end;
  end;
end;

begin
  CR := FindRun(CP);
  if (CR = Nil) or (CP <= 1) then begin
    Result := False;
    Exit;
  end;

  ScanForFirst(actWhitespace);
  ScanForFirst(actNotWhitespace);
  Inc(CP);
  Result := True;
end;

function TAXWLogPara.FindRun(CP: integer): TAXWCharRun;
var
  i: integer;
begin
  { TODO Binsearch }
  Result := Nil;
  for i := 0 to Count - 1 do begin
    if (CP >= Items[i].StartPos) and (CP <= Items[i].EndPos) then begin
      Result := Items[i];
      Exit;
    end;
  end;
  i := Count - 1;
  // At EOL
  if (i >= 0) and (CP = (Items[i].EndPos + 1)) then
    Result := Items[i];
end;

function TAXWLogPara.FindText(const AText: AxUCString): integer;
var
  i: integer;
  p: integer;
begin
  for i := 0 to FCharRuns.Count - 1 do begin
    p := Pos(AText,FCharRuns[i].Text);
    if p > 0 then begin
      Result := FCharRuns[i].StartPos + p - 1;
      Exit;
    end;
  end;
  Result := -1;
end;

function TAXWLogPara.First: TAXWLogPara;
begin
  if Parent.Count > 0 then
    Result := Parent.Items[0]
  else
    Result := Nil;
end;

function TAXWLogPara.CRFirst: TAXWCharRun;
begin
  if Count > 0 then
    Result := Items[0]
  else
    Result := Nil;
end;

procedure TAXWLogPara.FormatRange(List: TAXWCHPX; CP1, CP2: integer);
var
  CR1,CR2: TAXWCharRun;
begin
  CR1 := FindRun(CP1);
  CR2 := FindRun(CP2);
  if (CR1 <> Nil) and (CR2 <> Nil) then
    ApplyFormat(List,CP1,CP2,CR1,CR2);
end;

procedure TAXWLogPara.FormatAll(List: TAXWCHPX);
var
  CR1,CR2: TAXWCharRun;
begin
  CR1 := FCharRuns.First;
  CR2 := FCharRuns.Last;
  if (CR1 <> Nil) and (CR2 <> Nil) then
    ApplyFormat(List,CR1,CR2);
end;

procedure TAXWLogPara.FormatAllUncond(List: TAXWCHPX);
var
  CR1,CR2: TAXWCharRun;
begin
  CR1 := FCharRuns.First;
  CR2 := FCharRuns.Last;
  if (CR1 <> Nil) and (CR2 <> Nil) then
    SetFormat(List,CR1,CR2);
end;

procedure TAXWLogPara.FormatRangeUncond(List: TAXWCHPX; CP1, CP2: integer);
var
  CR1,CR2: TAXWCharRun;
begin
  CR1 := FindRun(CP1);
  CR2 := FindRun(CP2);
  if (CR1 <> Nil) and (CR2 <> Nil) then
    SetFormat(List,CP1,CP2,CR1,CR2);
end;

function TAXWLogPara.FormatWord(List: TAXWCHPX; CP: integer): boolean;
var
  C: AxUCChar;
  P1,P2: integer;
  CR1,CR2: TAXWCharRun;
begin
  Result := WordAtPos(CP,P1,P2,CR1,CR2);
  if Result then begin
    if (CP >= CR1.StartPos) and (CP <= CR1.EndPos) then begin
      // Start position must be after the first char.
      C := CR1.CPChar[CP - 1];
      Result := (C <> AXW_CHAR_EOL) and not CharIsWhitespace(C);
    end;
    if Result then begin
      if (P2 < CR2.EndPos) and CharIsWhitespace(CR2.Text[P2 - CR2.StartPos + 2]) then
        Inc(P2);
      ApplyFormat(List,P1,P2,CR1,CR2);
    end;
  end;
end;

function TAXWLogPara.GetChar(P: integer): AxUCChar;
var
  i: integer;
  CR: TAXWCharRun;
begin
  CR := FindRun(P);
  i := P - CR.StartPos + 1;
  if (i >= 1) and (i <= CR.Len) then
    Result := CR.Text[i]
  else
    Result := AXW_CHAR_EOL;
end;

function TAXWLogPara.GetItems(Index: integer): TAXWCharRun;
begin
  Result := FCharRuns.Items[Index];
end;

function TAXWLogPara.GetParent: TAXWLogParas;
begin
  Result := TAXWLogParas(FOwner)
end;

function TAXWLogPara.GetText: AxUCString;
var
  i: integer;
begin
  Result := '';
  for i := 0 to Count - 1 do
    Result := Result + Items[i].Text;
end;

procedure TAXWLogPara.InsertText(S: AxUCString; P: integer; InCRLeft: boolean = False);
var
  CR,CR2: TAXWCharRun;
  S2: AxUCString;
  CHPX: TAXWCHPX;
begin
  CR := FindRun(P);
  if CR = Nil then begin
    CHPX := TAXWCHPX.Create(FCHP);
    try
      FCHP.CopyToCHPX(CHPX);

      CR := FCharRuns.Add(CHPX);
      CR.Text := S;
    finally
      CHPX.Free;
    end;
  end
  else begin
    if InCRLeft and (P = CR.StartPos) and (FCharRuns.Prev(CR) <> Nil) then begin
      CR2 := FCharRuns.Prev(CR);
      CR2.Text := CR2.Text + S;
    end
    else begin
      S2 := CR.Text;
      System.Insert(S,S2,P - CR.StartPos + 1);
      CR.Text := S2;
    end;
  end;

  Sync(CR);

  FCharRuns.Dirty := True;
end;

function TAXWLogPara.CRLast: TAXWCharRun;
begin
  if Count > 0 then
    Result := Items[Count - 1]
  else
    Result := Nil;
end;

function TAXWLogPara.Last: TAXWLogPara;
begin
  if Parent.Count > 0 then
    Result := Parent.Items[Parent.Count - 1]
  else
    Result := Nil;
end;

function TAXWLogPara.Len: integer;
begin
  if Count > 0 then
    Result := CRLast.EndPos
  else
    Result := 0;
end;

function TAXWLogPara.MapTab(var APixelPos: double; AIndex: integer): TAXWTabStop;
var
  p: double;
begin
  Result := Parent.DOP.DefTab;
  if FTabs <> Nil then begin
    if AIndex < FTabs.Count then begin
      p := FTabs[AIndex].Position;
      Result := FTabs[AIndex];
    end
    else
      p := 0;
  end
  else
    p := Parent.DOP.TabWidth * (AIndex + 1);
  p := p * (Parent.PPIX / 72);
  if p > APixelPos then
    APixelPos := p
  else
    Result := Parent.DOP.DefTab;
end;

function TAXWLogPara.Next: TAXWLogPara;
begin
  if Index < (Parent.Count - 1) then
    Result := Parent[Index + 1]
  else
    Result := Nil;
end;

function TAXWLogPara.Prev: TAXWLogPara;
begin
  if Index > 0 then
    Result := Parent[Index - 1]
  else
    Result := Nil;
end;

procedure TAXWLogPara.SetDocPropsAtPos(CP: integer);
var
  CHPX: TAXWCHPX;
begin
  CHPX := CHPXAtPos(CP);
  Parent.FCaretCHPX := CHPX;
  Parent.FCaretPAPX := FPAPX
end;

procedure TAXWLogPara.SetFormat(List: TAXWCHPX; CP1, CP2: integer; CR1, CR2: TAXWCharRun);
var
  i: integer;
  CR1b, CR2b: TAXWCharRun;

procedure ApplySingle(CRun: TAXWCharRun; Pos1,Pos2: integer);
var
  CR: TAXWCharRun;
begin
  CR := SplitCharRun(CRun,Pos1,Pos2);
  CR.AddCHPX;
  CR.CHPX.Clear;
  CR.CHPX.Assign(List);
end;

begin
  { TODO Check that there is no empty char runs left after split (possibly with only spaces)}
  if CR1 = CR2 then begin
    ApplySingle(CR1,CP1,CP2);
  end
  else begin
    if CP1 > 1 then begin
      ApplySingle(CR1,CP1,CR1.EndPos);
      CR1b := FCharRuns.Next(CR1);
    end
    else
      CR1b := CR1;
    if CP2 < CR2.EndPos then begin
      ApplySingle(CR2,1,CP2);
      CR2b := FCharRuns.Prev(CR2);
    end
    else
      CR2b := CR2;
    for i := CR1b.Index to CR2b.Index do
      ApplySingle(FCharRuns[i],1,FCharRuns[i].EndPos);
  end;
  FCharRuns.Combine(CR1,CR2);

  Sync(CR1);
end;

procedure TAXWLogPara.SetFormat(List: TAXWCHPX; CR1, CR2: TAXWCharRun);
var
  i: integer;
  CR: TAXWCharRun;
begin
  if (List = Nil) or (List.Count = 0) then
    Exit;
  for i := CR1.Index to CR2.Index do begin
    CR := FCharRuns[i];
    CR.AddCHPX;
    CR.CHPX.Clear;
    CR.CHPX.Assign(List);
  end;
  FCharRuns.Combine(CR1,CR2);
end;

function TAXWLogPara.SplitCharRun(CR: TAXWCharRun; CP: integer): TAXWCharRun;
begin
  if (CP > CR.StartPos) and (CP < CR.EndPos) then begin
    Result := FCharRuns.Split(CR,CP);
    Sync(CR);
  end
  else
    Result := CR;
end;

function TAXWLogPara.SplitCharRun(CR: TAXWCharRun; CP1, CP2: integer): TAXWCharRun;
begin
  if (CP1 > CR.StartPos) and (CP2 < CR.EndPos) then begin
    Result := FCharRuns.Split(CR,CP1);
    Sync(CR);
    FCharRuns.Split(Result,CP2 + 1);
  end
  else if CP1 > CR.StartPos then begin
    Result := FCharRuns.Split(CR,CP1);
  end
  else if CP2 < CR.EndPos then begin
    Result := CR;
    FCharRuns.Split(Result,CP2 + 1);
  end
  else
    Result := CR;
end;

procedure TAXWLogPara.Sync(StartCR: TAXWCharRun);
var
  i,StartIndex: integer;
  CP: integer;
begin
  if Count < 2 then begin
    if Count > 0 then
      Items[0].StartPos := 1;
  end
  else begin
    StartIndex := StartCR.Index;
    if StartIndex < 1 then begin
      CP := 1;
      for i := 0 to Count - 1 do begin
        Items[i].StartPos := CP;
        Inc(CP,Items[i].Len);
      end;
    end
    else begin
      for i := StartIndex to Count - 1 do
        Items[i].StartPos := Items[i - 1].EndPos + 1;
    end;
  end;
end;

function TAXWLogPara.WordAtPos(CP: integer; var CP1, CP2: integer): boolean;
var
  CR1,CR2: TAXWCharRun;
begin
  Result := WordAtPos(CP,CP1,CP2,CR1,CR2);
end;

function TAXWLogPara.WordAtPos(CP: integer; var P1, P2: integer; var CR1, CR2: TAXWCharRun): boolean;
var
  CR: TAXWCharRun;
begin
  CR := FindRun(CP);
  if CR = Nil then begin
    Result := False;
    Exit;
  end;

  P1 := CP;
  P2 := CP;
  CR1 := FCharRuns.FindBeginningOfWord(P1,CR);
  CR2 := FCharRuns.FindEndOfWord(P2,CR);
  Result := (CR1 <> Nil) and (CR2 <> Nil);
end;

{ TXSSLogParas }

function TAXWLogParas.New: TAXWLogPara;
begin
  Result := TAXWLogPara.Create(Self,FMasterCHP);
end;

procedure TAXWLogParas.Sync;
var
  i: integer;
begin
  for i := 0 to Count - 1 do begin
    if Items[i].Count > 0 then
      Items[i].Sync(Items[i][0]);
  end;
end;

procedure TAXWLogParas.ToggleFormat(Id: TAXWChpId; Pos1,Pos2: TAXWLogPosition);
var
  ValBool: boolean;
  ValUnderline: TAXWChpUnderline;
  CHPX: TAXWCHPX;
begin
  ValBool := False;
  ValUnderline := axcuSingle;
  CHPX := Pos1.Para.CHPXAtPos(Pos1.CharPos);
  if Id = axciUnderline then begin
    if CHPX <> Nil then begin
      if CHPX.Underline > axcuNone then
        ValUnderline := axcuNone
      else
        ValUnderline := axcuSingle;
    end
    else
      ValUnderline := axcuSingle;
  end
  else begin
    if CHPX <> Nil then
      ValBool := not CHPX.AsBoolean[Integer(Id)]
    else
      ValBool := True;
  end;

  CHPX := TAXWCHPX.Create(FMasterCHP);
  try
    if Id = axciUnderline then
      CHPX.Underline := ValUnderline
    else
      CHPX.AsBoolean[Integer(Id)] := ValBool;
    ApplyFormat(CHPX,Pos1,Pos2);
  finally
    CHPX.Free;
  end;
end;

procedure TAXWLogParas.ClearRows;
var
  i: integer;
begin
  for i := 0 to Count - 1 do
    Items[i].Rows.Clear;
end;

function TAXWLogParas.CountCR: integer;
var
  i: integer;
 begin
  Result := 0;
  for i := 0 to Count - 1 do
    Inc(Result,Items[i].Count);
end;

constructor TAXWLogParas.Create(ADOP: TAXWDOP; AMasterPAP: TAXWPAP; AMasterCHP: TAXWCHP);
begin
  inherited Create(Nil);
  FDOP := ADOP;
  FMasterPAP := AMasterPAP;
  FMasterCHP := AMasterCHP;
end;

destructor TAXWLogParas.Destroy;
begin
  inherited;
end;

function TAXWLogParas.Empty: boolean;
begin
  Result := Count <= 0;
end;

function TAXWLogParas.First: TAXWLogPara;
begin
  if Count > 0 then
    Result := Items[0]
  else
    Result := Nil;
end;

function TAXWLogParas.GetItems(Index: integer): TAXWLogPara;
begin
  Result := TAXWLogPara(inherited Items[Index]);
end;

function TAXWLogParas.Insert(Index: integer): TAXWLogPara;
begin
  Result := TAXWLogPara.Create(Self,FMasterCHP);
  inherited Insert(Index,Result);
end;

function TAXWLogParas.Last: TAXWLogPara;
begin
  if Count > 0 then
    Result := Items[Count - 1]
  else
    Result := Nil;
end;

procedure TAXWLogParas.MergeNext(Index: integer);
begin
  if (Count < 2) or (Index >= Count) then
    Exit;
  Items[Index].Append(Items[Index + 1]);
  Delete(Index + 1);
end;

function TAXWLogParas.Add: TAXWLogPara;
begin
  Result := TAXWLogPara.Create(Self,FMasterCHP);
  Add(Result);
end;

procedure TAXWLogParas.Add(Para: TAXWLogPara);
begin
  inherited Add(Para);
end;

procedure TAXWLogParas.AddRawText(S: AxUCString);
begin
  Add.FCharRuns.Add(Nil).Text := S;
end;

procedure TAXWLogParas.ApplyFormat(CHPX: TAXWCHPX; Pos1,Pos2: TAXWLogPosition);
var
  i,iStart,iEnd: integer;
  CR: TAXWCharRun;
begin
  if Pos1.Para = Pos2.Para then begin
    if (Pos1.CharPos = 1) and (Pos2.CharPos = Pos1.Para.EndOfParaPos) then
      Pos1.Para.FormatAll(CHPX)
    else if Pos1.CPIsAtEOP and Pos2.CPIsAtEOP then begin
      CR := Pos1.Para.CRLast;
      if CR <> Nil then
        CHPX.Merge(CR.CHPX);
      Pos1.Para.AppendText(CHPX,'');
    end
    else
      Pos1.Para.FormatRange(CHPX,Pos1.CharPos,Pos2.CharPosDec)
  end
  else begin
    if Pos1.CharPos = 1 then
      iStart := Pos1.Para.Index
    else begin
      Pos1.Para.FormatRange(CHPX,Pos1.CharPos,Pos1.Para.EndOfParaPos);
      iStart := Pos1.Para.Index + 1;
    end;
    if Pos2.CharPos = Pos2.Para.EndOfParaPos then
      iEnd := Pos2.Para.Index
    else begin
      Pos2.Para.FormatRange(CHPX,1,Pos2.CharPosDec);
      iEnd := Pos2.Para.Index - 1;
    end;
    for i := iStart to iEnd do
      Items[i].FormatAll(CHPX);
  end;
end;

procedure TAXWLogParas.ApplyFormatUncond(CHPX: TAXWCHPX; Pos1,Pos2: TAXWLogPosition);
var
  i,iStart,iEnd: integer;
begin
  if Pos1.Para = Pos2.Para then begin
    if (Pos1.CharPos = 1) and (Pos2.CharPos = Pos1.Para.EndOfParaPos) then
      Pos1.Para.FormatAllUncond(CHPX)
    else
      Pos1.Para.FormatRangeUncond(CHPX,Pos1.CharPos,Pos2.CharPosDec)
  end
  else begin
    if Pos1.CharPos = 1 then
      iStart := Pos1.Para.Index
    else begin
      Pos1.Para.FormatRangeUncond(CHPX,Pos1.CharPos,Pos1.Para.EndOfParaPos);
      iStart := Pos1.Para.Index + 1;
    end;
    if Pos2.CharPos = Pos2.Para.EndOfParaPos then
      iEnd := Pos2.Para.Index
    else begin
      Pos2.Para.FormatRangeUncond(CHPX,1,Pos2.CharPosDec);
      iEnd := Pos2.Para.Index - 1;
    end;
    for i := iStart to iEnd do
      Items[i].FormatAllUncond(CHPX);
  end;
end;

procedure TAXWLogParas.ApplyParaFormat(PAPX: TAXWPAPX; Para1, Para2: TAXWLogPara);
var
  i: integer;
begin
  for i := Para1.Index to Para2.Index do
    Items[i].PAPX.Merge(PAPX);
end;

{ TAXWCustomLogPosition }

procedure TAXWLogPosition.Assign(LPP: TAXWLogPosition);
begin
  FPara := LPP.FPara;
  FCharPos := LPP.FCharPos;
end;

procedure TAXWLogPosition.Assign(APara: TAXWLogPara);
begin
  FPara := APara;
  FCharPos := 1;
end;

function TAXWLogPosition.CharPosDec: integer;
begin
  if FCharPos > 1 then
    Result := FCharPos - 1
  else
    Result := FCharPos;
end;

procedure TAXWLogPosition.Clear;
begin
  FPara := Nil;
  FCharPos := 1;
end;

function TAXWLogPosition.CPIsAtBOP: boolean;
begin
  Result := FCharPos = 1;
end;

function TAXWLogPosition.CPIsAtEOP: boolean;
begin
  Result := FCharPos = FPara.EndOfParaPos;
end;

constructor TAXWLogPosition.Create;
begin
  FCharPos := 1;
end;

procedure TAXWLogPosition.DecPos;
begin
  Dec(FCharPos);
end;

function TAXWLogPosition.EndOfParaPos: integer;
begin
  Result := FPara.EndOfParaPos;
end;

function TAXWLogPosition.Equal(LPP: TAXWLogPosition): boolean;
begin
  Result := (LPP.FPara = FPara) and (LPP.FCharPos = FCharPos);
end;

function TAXWLogPosition.Greater(LPP: TAXWLogPosition): boolean;
begin
  Result := FPara.Index > LPP.FPara.Index;
  if not Result then
    Result := (FPara.Index = LPP.FPara.Index) and (FCharPos > LPP.FCharPos);
end;

function TAXWLogPosition.GreaterEqual(LPP: TAXWLogPosition): boolean;
begin
  Result := FPara.Index > LPP.FPara.Index;
  if not Result then
    Result := (FPara.Index = LPP.FPara.Index) and (FCharPos >= LPP.FCharPos);
end;

procedure TAXWLogPosition.IncPos;
begin
  Inc(FCharPos);
end;

function TAXWLogPosition.Less(LPP: TAXWLogPosition): boolean;
begin
  Result := FPara.Index < LPP.FPara.Index;
  if not Result then
    Result := (FPara.Index = LPP.FPara.Index) and (FCharPos < LPP.FCharPos);
end;

function TAXWLogPosition.LessEqual(LPP: TAXWLogPosition): boolean;
begin
  Result := FPara.Index < LPP.FPara.Index;
  if not Result then
    Result := (FPara.Index = LPP.FPara.Index) and (FCharPos <= LPP.FCharPos);
end;

procedure TAXWLogPosition.MovePosNext;
begin
  FParaChanged := False;
  if FCharPos < EndOfParaPos then
    Inc(FCharPos)
  else if FPara.Next <> Nil then begin
    FPara := FPara.Next;
    FCharPos := 1;
    FParaChanged := True;
  end;
end;

procedure TAXWLogPosition.MovePosPrev;
begin
  FParaChanged := False;
  if FCharPos > 1 then begin
    Dec(FCharPos);
  end
  else if (FPara <> Nil) and (FPara.Prev <> Nil) then begin
    FPara := FPara.Prev;
    FCharPos := EndOfParaPos;
    FParaChanged := True;
  end;
end;

procedure TAXWLogPosition.MoveToBeginningOfDoc;
begin
  FParaChanged := FPara <> FPara.First;
  FPara := FPara.First;
  FCharPos := 1;
end;

procedure TAXWLogPosition.MoveToBeginningOfPara;
begin
  FParaChanged := False;
  FCharPos := 1;
end;

procedure TAXWLogPosition.MoveToEndOfDoc;
begin
  FParaChanged := FPara <> FPara.Last;
  FPara := FPara.Last;
  FCharPos := FPara.Len + 1;
end;

procedure TAXWLogPosition.MoveToEndOfPara;
begin
  FParaChanged := False;
  FCharPos := EndOfParaPos;
end;

procedure TAXWLogPosition.SetDocPropsAtPos;
begin
  FPara.SetDocPropsAtPos(FCharPos);
end;

procedure TAXWLogPosition.SetPos(const Value: integer);
begin
  FCharPos := Value;
end;

procedure TAXWLogPosition.SetWithinLimits;
begin
  if FPara.Index < 0 then
    FPara := FPara.First;
  if FCharPos < 1 then
    FCharPos := 1;
  if FPara.Index > FPara.Last.Index then begin
    FPara := FPara.Last;
    FCharPos := FPara.EndOfParaPos;
  end
  else if FCharPos > FPara.EndOfParaPos then
    FCharPos := FPara.EndOfParaPos;
end;

procedure TAXWLogPosition.Update;
begin

end;

{ TAXWFindTextData }

procedure TAXWFindTextData.BeginFind(Para: TAXWLogPara);
begin
  FPara := Para;
  FAtEndOfDoc := Para.Count <= 0;
  if not FAtEndOfDoc then
    FCharRun := Para[0];
  FCharPos := 1;
end;

function TAXWFindTextData.FindNext: boolean;
var
  i: integer;
begin
  Result := False;
  if FAtEndOfDoc or (FText = '') then
    Exit;

  for i := FCharRun.Index to FPara.Count - 1 do begin

  end;
end;

function TAXWFindTextData.GetFound: boolean;
begin
  Result := False;
end;

{ TAXWLogSelection }

function TAXWLogSelection.AsPlainText: AxUCString;
//var
//  i: integer;
begin
  if FFirst.Para = FLast.Para then
    Result := FFirst.Para.AsPlainText(FFirst.CharPos,FLast.CharPos)
  else begin
    raise XLSRWException.Create('TODO');
//    Result := '';
//    for i := FFirst.Para.Index to FLast.Para.Index do begin
//
//    end;
  end;
end;

constructor TAXWLogSelection.Create(Parent: TAXWLogSelections);
begin
  FParent := Parent;
  FFirst := TAXWLogPosition.Create;
  FLast := TAXWLogPosition.Create;
end;

destructor TAXWLogSelection.Destroy;
begin
  FFirst.Free;
  FLast.Free;
  inherited;
end;

function TAXWLogSelection.Hit(Para: TAXWLogPara; Row: TAXWPhyRow): TAXWSelectionHit;
var
  PFirst,PLast,P1,P2: int64;
begin
  if (Para.Index < FFirst.Para.Index) or (Para.Index > FLast.Para.Index) then
    Result := axshNone
  else if (Para.Index > FFirst.Para.Index) and (Para.Index < FLast.Para.Index) then
    Result := axshFullRow
  else begin
    Int64Rec(PFirst).Hi := FFirst.Para.Index;
    Int64Rec(PFirst).Lo := FFirst.CharPos;
    Int64Rec(PLast).Hi := FLast.Para.Index;
    Int64Rec(PLast).Lo := FLast.CharPos;
    Int64Rec(P1).Hi := Para.Index;
    Int64Rec(P1).Lo := Row.FirstPos;
    Int64Rec(P2).Hi := Para.Index;
    Int64Rec(P2).Lo := Row.LastPos;

    if (PFirst <= P1) and (PLast >= P2) then
      Result := axshFullRow
    else if (PFirst >= P1) and (PFirst <= P2) then begin
      FParent.FHitP1Offs := PFirst - P1 + 1;
      if (Plast >= P1) and (PLast <= P2) then begin
        FParent.FHitP2Offs := PLast - P1 + 1;
        Result := axshBothOnRow;
      end
      else
        Result := axshFirstOnRow;
    end
    else if (PLast >= P1) and (PLast <= P2) then begin
      FParent.FHitP2Offs := PLast - P1 + 1;
      if (PFirst >= P1) and (PFirst <= P2) then begin
        FParent.FHitP1Offs := PFirst - P1 + 1;
        Result := axshBothOnRow;
      end
      else begin
        Result := axshLastOnRow;
      end;
    end
    else
      Result := axshNone;
  end;
  FParent.FHitResult := Result;
end;

{ TAXWLogSelectionsList }

procedure TAXWLogSelections.Add;
begin
  inherited Add(TAXWLogSelection.Create(Self));
end;

constructor TAXWLogSelections.Create;
begin
  inherited Create;
end;

function TAXWLogSelections.GetItems(Index: integer): TAXWLogSelection;
begin
  Result := TAXWLogSelection(inherited Items[Index]);
end;

function TAXWLogSelections.Hit(Para: TAXWLogPara; Row: TAXWPhyRow): TAXWSelectionHit;
var
  i: integer;
begin
  // TODO More than one hit per row.
  for I := 0 to Count - 1 do begin
    Result := Items[i].Hit(Para,Row);
    if Result <> axshNone then
      Exit;
  end;
  Result := axshNone;
end;

procedure TAXWLogSelections.Reset;
begin
  FHitResult := axshNone;
end;

end.
