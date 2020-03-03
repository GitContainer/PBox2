unit XSSIELogEditor;

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

uses { Delphi } Classes, SysUtils, Contnrs, Math, vcl.Clipbrd,
     { Windows} Windows,
                XLSUtils5,
     { AxWord } XSSIEDefs, XSSIEUtils, XSSIEDocProps, XSSIECharRun, XSSIELogParas,
                XSSIELogPhyPosition, XSSIEKeys;

type TAXWSelectionStatus = (assNone,assUpdating);

type TAXWPhyPosition = (apsNone,apsTopRow,apsBottomRow,apsPageUp,apsPageDown);

type TAXWLogDocument = class(TObject)
private
     function  GetDirty: boolean;
protected
     FLogParas  : TAXWLogParas;
     FDOP       : TAXWDOP;
     FMasterPAP : TAXWPAP;
     FMasterCHP : TAXWCHP;
     FTempCHPX  : TAXWCHPX;
     FTempPAPX  : TAXWPAPX;

     procedure Paginate; virtual; abstract;
     procedure PaginateCaretPara; virtual; abstract;
     procedure ClearDirty;
public
     constructor Create;
     destructor Destroy; override;

     procedure Clear; virtual;

     procedure AppendTextSimple(const AText: AxUCString);

     procedure AppendText(const AText: AxUCString); overload;
     procedure AppendText(const AText: AxUCString; CHPX: TAXWCHPX); overload;
//     procedure AppendText(const AText: AxUCString; AChpId: TAXWChpId; AValue: integer); overload;

     function  AppendPara: TAXWLogPara; overload;
     procedure AppendPara(S: AxUCString); overload;
     procedure AppendPara(S: AxUCString; PAPX: TAXWPAPX); overload;

     procedure AppendTextFmt(S: AxUCString);

     function  PlainText: AxUCString;

     property Paras: TAXWLogParas read FLogParas;

     property DOP: TAXWDOP read FDOP;
     property MasterCHP: TAXWCHP read FMasterCHP;
     property MasterPAP: TAXWPAP read FMasterPAP;
     property TempCHPX: TAXWCHPX read FTempCHPX;
     property TempPAPX: TAXWPAPX read FTempPAPX;
     property Dirty: boolean read GetDirty;
     end;

type TAXWUndoRedoOp = (uroInsertText,uroInsertPara,uroMove,uroDeleteText);

type TAXWUndoItem = class(TObject)
protected
     FValue: integer;
     FBuffer: AxUCString;
     FOperation: TAXWUndoRedoOp;
public
     property Buffer: AxUCString read FBuffer write FBuffer;
     property Operation: TAXWUndoRedoOp read FOperation write FOperation;
     property Value: integer read FValue write FValue;
     end;

type TAXWUndo = class(TObjectStack)
protected
public
     destructor Destroy; override;

     procedure AddInsert(Text: AxUCString);
     procedure AddDelete(Text: AxUCString);
     procedure AddMove(Distance: integer);

     function Peek: TAXWUndoItem;
     end;

// Never move First/Last position directly as Caret/Anchor may go out of sync.
type TAXWSelections = class(TAXWLogSelections)
private
     function GetFirstPos: TAXWLogPosition; {$ifdef D2006PLUS} inline; {$endif}
     function GetLastPos: TAXWLogPosition; {$ifdef D2006PLUS} inline; {$endif}
protected
     FAnchor: TAXWLogPosition;
     FCaretPos: TAXWLogPosition;
     FTempPos1: TAXWLogPosition;
     FTempPos2: TAXWLogPosition;

     FStatus: TAXWSelectionStatus;

     FSavedParaIndex: integer;
     FSavedCharPos: integer;
     FSavedCharPosIsEOP: boolean;

     FIteratePara: TAXWLogPara;
     FIterateCR,FIterateCREnd: TAXWCharRun;

     procedure Normalize;

     procedure BeginIterateCR;
     function  NextCR: boolean;

     property IterateCR: TAXWCharRun read FIterateCR write FIterateCR;
public
     constructor Create(Caret: TAXWLogPosition);
     destructor Destroy; override;

     procedure Reset; override;

     function  HasSelection: boolean;
     function  IsEntireDoc: boolean;
     function  IsZeroLen: boolean;

     procedure PushCaretPos;
     procedure PushFirstPos;
     procedure PopCaretPos(Paras: TAXWLogParas);

     procedure DropAnchor;
     procedure Update;
     procedure SetToCaret;
     procedure CopyToTemp;

     property Status: TAXWSelectionStatus read FStatus;
     property Anchor:   TAXWLogPosition read FAnchor;
     property FirstPos: TAXWLogPosition read GetFirstPos;
     property LastPos:  TAXWLogPosition read GetLastPos;
     property TempPos1: TAXWLogPosition read FTempPos1;
     property TempPos2: TAXWLogPosition read FTempPos2;
     end;

type TAXWLogDocEditor = class(TAXWLogDocument)
protected
     FCaretPos: TAXWEventLogPhyPosition;
     FSelections: TAXWSelections;

     FUndo: TAXWUndo;

     FRepaint: boolean;

     FCaretPosEvent: TNotifyEvent;

     procedure DoMoveCaret(Cmd: TAXWCommand); overload;
     procedure ClearSelection;
     procedure ClearCHPFormatting;
     function  GetPhyPosition(const APosition: TAXWPhyPosition): TAXWLogPhyPosition; virtual;
public
     constructor Create;
     destructor Destroy; override;

     procedure ClearData;

     procedure SetEmpty;
     function  IsEmpty: boolean;

     procedure BeginUpdate;
     procedure EndUpdate;
     procedure AddPara;
     procedure InsertText(S: AxUCString; InCRLeft: boolean = False);
     procedure InsertPara(S: AxUCString);

     procedure CmdControl(Cmd: TAXWCommand);
     procedure CmdMove(Cmd: TAXWCommand);
     procedure CmdClipboard(Cmd: TAXWCommand);
     procedure CmdSelect(Cmd: TAXWCommand);
     procedure CmdDelete(Cmd: TAXWCommand);
     procedure CmdFormat(Cmd: TAXWCommand);

     procedure Undo;

     procedure ReadTextFile(Filename: AxUCString);
     procedure ReadTextStream(Stream: TStream);

     procedure UpdateSelection(Direction: TAXWCommand);
     procedure FormatRange(List: TAXWCHPX; LPP1,LPP2: TAXWLogPosition);

     property CaretPos: TAXWEventLogPhyPosition read FCaretPos;
     property Selections: TAXWSelections read FSelections;

     property OnCaretPos: TNotifyEvent read FCaretPosEvent write FCaretPosEvent;
     end;

implementation

{ TAXWLogDocument }

// TODO S can only contain valid characters. Do a check for this.
procedure TAXWLogDocument.AppendPara(S: AxUCString);
var
  Para: TAXWLogPara;
  CHPX: TAXWCHPX;
begin
  Para := FLogParas.New;

  CHPX := TAXWCHPX.Create(FMasterCHP);
  try
    FLogParas.MasterCHP.CopyToCHPX(CHPX);

    Para.AppendText(CHPX,ValidateText(S));
  finally
    CHPX.Free;
  end;
  FLogParas.Add(Para);
end;

procedure TAXWLogDocument.AppendPara(S: AxUCString; PAPX: TAXWPAPX);
var
  Para: TAXWLogPara;
  CHPX: TAXWCHPX;
begin
  Para := FLogParas.New;

  CHPX := TAXWCHPX.Create(FMasterCHP);
  try
    FLogParas.MasterCHP.CopyToCHPX(CHPX);
    Para.AppendText(CHPX,ValidateText(S));
    if PAPX <> Nil then
      Para.PAPX.Merge(PAPX)
    else
      Para.PAPX.Merge(FTempPAPX);
  finally
    CHPX.Free;
  end;
  FLogParas.Add(Para);
end;

procedure TAXWLogDocument.AppendText(const AText: AxUCString);
var
  S,S2: AxUCString;
  Para: TAXWLogPara;
  CHPX: TAXWCHPX;
  NewPara: boolean;
begin
  NewPara := FLogParas.Count <= 0;
  if NewPara then
    Para := FLogParas.New
  else
    Para := FLogParas.Last;

  CHPX := TAXWCHPX.Create(FMasterCHP);
  try
    S := AText;
    repeat
      S2 := SplitAtChar(AxUCChar(VK_RETURN),S);
      FLogParas.MasterCHP.CopyToCHPX(CHPX);
      Para.AppendText(CHPX,S2);
      if NewPara then
        FLogParas.Add(Para);
      if S <> '' then begin
        Para := FLogParas.New;
        NewPara := True;
      end;
    until (S = '');
  finally
    CHPX.Free;
  end;
end;

procedure TAXWLogDocument.AppendTextSimple(const AText: AxUCString);
var
  Para: TAXWLogPara;
  NewPara: boolean;
begin
  NewPara := FLogParas.Count <= 0;
  if NewPara then
    Para := FLogParas.New
  else
    Para := FLogParas.Last;

//  Para.PAPX.AddInteger(axpiAlignment,Integer(axptaJustify));

  Para.CRLast.Text := Para.CRLast.Text + AText;
  if NewPara then
    FLogParas.Add(Para);
end;

function TAXWLogDocument.AppendPara: TAXWLogPara;
begin
  Result := FLogParas.New;
  FLogParas.Add(Result);
end;

procedure TAXWLogDocument.AppendText(const AText: AxUCString; CHPX: TAXWCHPX);
var
  Para: TAXWLogPara;
  NewPara: boolean;
begin
  NewPara := FLogParas.Count <= 0;
  if NewPara then
    Para := FLogParas.New
  else
    Para := FLogParas.Last;

  Para.AppendText(CHPX,AText);

  if NewPara then
    FLogParas.Add(Para);
end;

//procedure TAXWLogDocument.AppendText(const AText: AxUCString; AChpId: TAXWChpId; AValue: integer);
//var
//  CHPX: TAXWCHPX;
//begin
//  CHPX := TAXWCHPX.Create;
//  try
//    CHPX.AddInteger(AChpId,AValue);
//    AppendText(AText,CHPX);
//  finally
//    CHPX.Free;
//  end;
//end;

procedure TAXWLogDocument.AppendTextFmt(S: AxUCString);
type
  TFmtAttribute = (axfaNewPara,axfaBold,axfaItalic,axfaUnderline,axfaFontSize,axfaFontName,axfaFontColor);
var
  NewPara: boolean;
  Txt,Token,Value: AxUCString;
  Attributes: TStringList;

function GetToken(var S,Token,Value: AxUCString): AxUCString;
var
  L,p1,p2: integer;
begin
  Token := '';
  p1 := CPos(AxUCChar('\'),S);
  p2 := CPos(AxUCChar(';'),S);

  if (p1 > 0) and (p2 > 0) and (p2 > p1) then begin
    if S[p1 + 1] = '\' then begin
      Result := Copy(S,1,p1 + 1);
      S := Copy(S,p1 + 2,MAXINT);
      Exit;
    end;

    Token := Copy(S,p1 + 1,p2 - p1 - 1);
    Result := Copy(S,1,p1 - 1);
    S := Copy(S,p2 + 1,MAXINT);
    if (Length(Token) = 1) and CharInSet(Token[1],['n','B','I','U']) then
      Exit
    else begin
      p1 := CPos(AxUCChar(':'),Token);
      if p1 <= 0 then
        Exit;
      Value := Copy(Token,p1 + 1,MAXINT);
      Token := Copy(Token,1,p1 - 1);
      L := Length(Value);
      if (Length(Value) > 1) and (Value[1] = '"') and (Value[L] = '"') then
        Value := Copy(Value,2,L - 2);
    end;
  end
  else if p1 > 0 then begin
    Result := Copy(S,1,p1 + 1);
    S := Copy(S,p1 + 2,MAXINT);
  end
  else begin
    Result := S;
    S := '';
  end;
end;

begin
  // TODO Use a global list
  Attributes := TStringList.Create;
  Attributes.AddObject('n',TObject(axfaNewPara));
  Attributes.AddObject('B',TObject(axfaBold));
  Attributes.AddObject('I',TObject(axfaItalic));
  Attributes.AddObject('U',TObject(axfaUnderline));
  Attributes.AddObject('Sz',TObject(axfaFontSize));
  Attributes.AddObject('F',TObject(axfaFontName));
  Attributes.AddObject('FC',TObject(axfaFontColor));
  Attributes.Sort;

  NewPara := False;
  try
    repeat
      Txt := GetToken(S,Token,Value);
      if Txt <> '' then begin
        if NewPara then begin
          AppendPara(Txt);
          NewPara := False;
        end
        else
          AppendText(Txt);
      end;
//      if Attributes.Find(Token,i) then begin
//        case TFmtAttribute(Attributes.Objects[i]) of
//          axfaNewPara:   NewPara := True;
//          axfaBold:      TextFormat.Bold := not TextFormat.Bold;
//          axfaItalic:    TextFormat.Italic := not TextFormat.Italic;
//          axfaUnderline: TextFormat.UL := not TextFormat.UL;
//          axfaFontSize:  TextFormat.Size := StrToIntDef(Value,TextFormat.Size);
//          axfaFontName:  TextFormat.FontName := Value;
////          axfaFontColor: TextFormat.Color := StrToColor(Value);
//        end;
//      end;
    until (S = '');
  finally
    Attributes.Free;
  end;
end;

procedure TAXWLogDocument.Clear;
begin
  FLogParas.Clear;
end;

procedure TAXWLogDocument.ClearDirty;
var
  i: integer;
begin
  FLogParas.Dirty := False;
  for i := 0 to FLogParas.Count - 1 do begin
    if FLogParas[i].Dirty then
      FLogParas[i].ClearDirty;
  end;
end;

constructor TAXWLogDocument.Create;
begin
  FDOP := TAXWDOP.Create;

  FMasterPAP := TAXWPAP.Create(Nil);
  FMasterCHP := TAXWCHP.Create(Nil);

  FTempPAPX := TAXWPAPX.Create(FMasterPAP);
  FTempCHPX := TAXWCHPX.Create(FMasterCHP);

  FLogParas := TAXWLogParas.Create(FDOP,FMasterPAP,FMasterCHP);
end;

destructor TAXWLogDocument.Destroy;
begin
  FreeAndNil(FLogParas);
  FreeAndNil(FMasterCHP);
  FreeAndNil(FMasterPAP);
  FDOP.Free;
  FreeAndNil(FTempCHPX);
  FreeAndNil(FTempPAPX);

  inherited;
end;

function TAXWLogDocument.GetDirty: boolean;
begin
  Result := FLogParas.Dirty;
end;

function TAXWLogDocument.PlainText: AxUCString;
var
  i: integer;
begin
  Result := '';
  for i := 0 to FLogParas.Count - 1 do
    Result := Result + FlogParas[i].Text;
end;

{ TAXWLogDocEditor }

procedure TAXWLogDocEditor.AddPara;
var
  CHPX: TAXWCHPX;
begin
  if FCaretPos.Para <> Nil then
    FLogParas.Insert(FCaretPos.Para.Index)
  else
    FLogParas.Add;
  CHPX := TAXWCHPX.Create(FMasterCHP);
  try
    FLogParas.MasterCHP.CopyToCHPX(CHPX);
  finally
    CHPX.Free;
  end;
end;

procedure TAXWLogDocEditor.BeginUpdate;
begin
//  FLogParas.BeginUpdate;
end;

procedure TAXWLogDocEditor.ClearCHPFormatting;
var
  S: AxUCString;
begin
  S := PlainText;
  ClearData;
  FCaretPos.Para := Nil;
  BeginUpdate;
  AppendPara(S);
  EndUpdate;
  Paginate;
  FRepaint := True;
end;

procedure TAXWLogDocEditor.ClearData;
begin
  inherited;
  FLogParas.Clear;
  FCaretPos.Clear;
  FSelections.Reset;
end;

procedure TAXWLogDocEditor.ClearSelection;
begin
  if FSelections.HasSelection then begin
    FSelections.CopyToTemp;
    FSelections.Reset;
    FRepaint := True;
  end;
end;

constructor TAXWLogDocEditor.Create;
begin
  inherited Create;
  FCaretPos := TAXWEventLogPhyPosition.Create;
  FSelections := TAXWSelections.Create(FCaretPos);
  FUndo := TAXWUndo.Create;
end;

destructor TAXWLogDocEditor.Destroy;
begin
  FreeAndNil(FUndo);
  FreeAndNil(FCaretPos);
  FreeAndNil(FSelections);

  inherited;
end;

procedure TAXWLogDocEditor.DoMoveCaret(Cmd: TAXWCommand);
var
  P: integer;
  LPP: TAXWLogPhyPosition;
  LPP2: TAXWLogPhyPosition;
begin
  if FLogParas.Empty then
    Exit;
  LPP2 := Nil;
  LPP := TAXWLogPhyPosition.Create;
  try
    LPP.Assign(FCaretPos);
    FCaretPos.Reset(FLogParas.First);
    case Cmd of
      axcNone: ;
      axcCurrent: begin
        FCaretPos.MakeValid; // TODO Check that caret has a valid pos. The area may have been deleted.
      end;
      axcMoveCharLeft: begin
        FCaretPos.MovePosPrev;
      end;
      axcMoveCharRight: begin
        FCaretPos.MovePosNext;
      end;
      axcMoveWordLeft: begin
        FCaretPos.BeforeMove;
        P := FCaretPos.CharPos;
        if FCaretPos.Para.FindPrevWord(P) then
          FCaretPos.CharPos := P
        else if FCaretPos.Para.Index > 0 then begin
          FCaretPos.Para := FLogParas[FCaretPos.Para.Index - 1];
          FCaretPos.CharPos := FCaretPos.EndOfParaPos;
        end;
        FCaretPos.AfterMove;
      end;
      axcMoveWordRight: begin
        FCaretPos.BeforeMove;
        P := FCaretPos.CharPos;
        if FCaretPos.Para.FindNextWord(P) then
          FCaretPos.CharPos := P
        else if FCaretPos.Para.Index < (FLogParas.Count - 1) then begin
          FCaretPos.Para := FLogParas[FCaretPos.Para.Index + 1];
          FCaretPos.CharPos := FCaretPos.BOLPos;
        end;
        FCaretPos.AfterMove;
      end;
      axcMoveParaUp: begin
        FCaretPos.MoveParaPrev;
      end;
      axcMoveParaDown: begin
        FCaretPos.MoveParaNext;
      end;
      axcMoveCellLeft: ;
      axcMoveCellRight: ;
      axcMoveLineUp: begin
        FCaretPos.MoveRowPrev;
      end;
      axcMoveLineDown: begin
        FCaretPos.MoveRowNext;
      end;
      axcMoveEndOfLine: begin
        FCaretPos.MoveToEOLPlus;
      end;
      axcMoveBeginningOfLine: begin
        FCaretPos.MoveToBOL;
      end;
      axcMoveTopOfWindow: begin
        LPP2 := GetPhyPosition(apsTopRow);
        if LPP2 <> Nil then
          FCaretPos.MoveToPos(LPP2);
      end;
      axcMoveEndOfWindow: begin
        LPP2 := GetPhyPosition(apsBottomRow);
        if LPP2 <> Nil then
          FCaretPos.MoveToPos(LPP2);
      end;
      axcMovePgUp: begin
        LPP2 := GetPhyPosition(apsPageUp);
        if LPP2 <> Nil then
          FCaretPos.MoveToPos(LPP2);
      end;
      axcMovePgDown: begin
        LPP2 := GetPhyPosition(apsPageDown);
        if LPP2 <> Nil then
          FCaretPos.MoveToPos(LPP2);
      end;
      axcMoveTopOfNextPg: ;
      axcMoveTopOfPrevPg: ;
      axcMoveEndOfDoc: begin
        FCaretPos.MoveToEndOfDoc;
      end;
      axcMoveBeginningOfDoc: begin
        FCaretPos.MoveToBeginningOfDoc;
      end;
      axcMoveBeginningOfPara: begin
        FCaretPos.MoveToBeginningOfPara;
      end;
      axcMoveEndOfPara: begin
        FCaretPos.MoveToEndOfPara;
      end;
      axcMoveLastLocation: ;
    end;
    if not LPP.Equal(FCaretPos) and Assigned(FCaretPosEvent) then begin
      FCaretPosEvent(Self);
  //    FUndo.AddMove();
    end;
  finally
    LPP.Free;
    if LPP2 <> Nil then
      LPP2.Free;
  end;
end;

procedure TAXWLogDocEditor.EndUpdate;
begin
  inherited;
  if not FCaretPos.Valid and (FLogParas.Count > 0) then
    FCaretPos.Reset(FLogParas[0]);
  FLogParas.Sync;
//  FLogParas.EndUpdate;
  CmdMove(axcCurrent);
end;

procedure TAXWLogDocEditor.FormatRange(List: TAXWCHPX; LPP1,LPP2: TAXWLogPosition);
var
  i: integer;
  LP1,LP2: TAXWLogPara;
begin
  if LPP1.Para = LPP2.Para then
    LPP1.Para.FormatRange(List,LPP1.CharPos,LPP2.CharPos - 1)
  else begin
    LPP1.Para.FormatRange(List,LPP1.CharPos,LPP1.Para.EndOfParaPos);
    LP1 := FLogParas[LPP1.Para.Index + 1];
    LPP2.Para.FormatRange(List,1,LPP2.CharPos - 1);
    LP2 := FLogParas[LPP2.Para.Index - 1];
    for i := LP1.Index to LP2.Index do
      FLogParas[i].FormatRange(List,1,FLogParas[i].EndOfParaPos);
  end;
end;

function TAXWLogDocEditor.GetPhyPosition(const APosition: TAXWPhyPosition): TAXWLogPhyPosition;
begin
  Result := Nil;
end;

procedure TAXWLogDocEditor.CmdClipboard(Cmd: TAXWCommand);
var
  S: AxUCString;
begin
  case Cmd of
    axcClipboardCopy       : begin
      if FSelections.HasSelection then
        Clipboard.AsText := FSelections[0].AsPlainText;
    end;
    axcClipboardCut        : begin
      if FSelections.HasSelection then begin
        Clipboard.AsText := FSelections[0].AsPlainText;
        CmdDelete(axcDelSelection);
      end;
    end;
    axcClipboardPaste,
    axcClipboardPastePlain : begin
      S := Clipboard.AsText;
      if S <> '' then begin
        InsertText(S);

        FCaretPos.BeforeMove;
        FCaretPos.CharPos := FCaretPos.CharPos + Length(S) - 1;
        FCaretPos.AfterMove;

        FRepaint := True;
      end;
    end;
  end;
end;

procedure TAXWLogDocEditor.CmdControl(Cmd: TAXWCommand);
begin
  case Cmd of
    axcControlNewPara: begin
      if FCaretPos.Para <> Nil then begin
        FLogParas.Insert(FCaretPos.Para.Index + 1);
        Paginate;
        CmdMove(axcMoveParaDown);
      end;
    end;
    axcControlLineBreak: begin
      InsertText(AXW_CHAR_SPECIAL_HARDLINEBREAK);
    end;
  end;
end;

{

}
procedure TAXWLogDocEditor.CmdDelete(Cmd: TAXWCommand);
var
  i: integer;
  StartIndex,EndIndex: integer;
  MergeParas: boolean;
begin
  case Cmd of
    axcDelCharLeft: begin
      if not IsEmpty then begin
        FSelections.DropAnchor;
        FSelections.Anchor.MovePosPrev;
        FSelections.Update;
      end;
    end;
    axcDelCharRight: begin
      if not FSelections.HasSelection then begin
        FSelections.DropAnchor;
        FSelections.Anchor.MovePosNext;
        FSelections.Update;
      end;
    end;
    axcDelWordLeft: ;  // TODO
    axcDelWordRight: ; // TODO
    axcDelSelection: begin
      if not FSelections.HasSelection or FSelections.IsZeroLen then
        Exit;
      if FSelections.IsEntireDoc then begin
        SetEmpty;
        CmdMove(axcCurrent);
        Exit;
      end;
      MergeParas := False;
      FSelections.PushFirstPos;
      if FSelections.FirstPos.Para.Index <> FSelections.LastPos.Para.Index then begin
        MergeParas := (FSelections.FirstPos.CharPos > 1) {and (FSelections.LastPos.Para.Next <> Nil)};
        StartIndex := FSelections.FirstPos.Para.Index + 1;
        EndIndex := FSelections.LastPos.Para.Index - 1;
        if FSelections.FirstPos.CharPos = 1 then
          Dec(StartIndex)
        else
          FSelections.FirstPos.Para.DeleteTextToEnd(FSelections.FirstPos.CharPos);
        if FSelections.LastPos.CPIsAtEOP then
          Inc(EndIndex)
        else if FSelections.LastPos.CharPos > 1 then
          FSelections.LastPos.Para.DeleteText(1,FSelections.LastPos.CharPos);
        for i := EndIndex downto StartIndex do
          FLogParas.Delete(i);
      end
      else begin
        if (FSelections.FirstPos.CharPos = 1) and FSelections.LastPos.CPIsAtEOP then
          FLogParas.Delete(FSelections.FirstPos.Para.Index)
        else
          FSelections.FirstPos.Para.DeleteText(FSelections.FirstPos.CharPos,FSelections.LastPos.CharPos - 1);
      end;

      if MergeParas then
        FLogParas.MergeNext(FSelections.FirstPos.Para.Index);

      FSelections.Reset;

      Paginate;

      FRepaint := True;

      FSelections.PopCaretPos(FLogParas);

      DoMoveCaret(axcMoveCharLeft);
      if (FCaretPos.CharPos >= 1) and (FCaretPos.CharPos < FCaretPos.OldPos) then
        DoMoveCaret(axcMoveCharRight);

      CmdMove(axcCurrent);
      Exit;
    end
    else
      Exit;
  end;
  CmdDelete(axcDelSelection);
end;

procedure TAXWLogDocEditor.CmdFormat(Cmd: TAXWCommand);
var
  V: double;
  CHPX: TAXWCHPX;

procedure GetSelectionCHPX;
begin
  FTempCHPX.Clear;
end;

procedure ApplyFormatToSelection;
begin
  if FSelections.FirstPos.Para.Index = FSelections.LastPos.Para.Index then begin
    FSelections.FirstPos.Para.FormatRange(FTempCHPX,FSelections.FirstPos.CharPos,FSelections.LastPos.CharPos);
  end
  else begin

  end;
end;

procedure ApplyBoolean(Id: TAXWChpId);
var
  CP1,CP2: integer;
begin
//  FLogParas.ToggleFormat(Id,FSelections.FirstPos.Para,FSelections.LastPos.Para,FSelections.FirstPos.CharPos,FSelections.LastPos.CharPos - 1);
  if FSelections.IsZeroLen and FSelections.FirstPos.Para.WordAtPos(FSelections.FirstPos.CharPos,CP1,CP2) then begin
    FSelections.CopyToTemp;
    FSelections.TempPos1.CharPos := CP1;
    FSelections.TempPos2.CharPos := CP2 + 1;
  end
  else
    FSelections.CopyToTemp;
  FLogParas.ToggleFormat(Id,FSelections.TempPos1,FSelections.TempPos2);
  FCaretPos.SetDocPropsAtPos;
end;

procedure ApplyInteger(Id: TAXWChpId; Value: integer);
var
  List: TAXWCHPX;
begin
  List := TAXWCHPX.Create(FMasterCHP);
  try
    List.AddInteger(Integer(Id),Value);
    FLogParas.ApplyFormat(List,FSelections.FirstPos,FSelections.LastPos);
    FCaretPos.SetDocPropsAtPos;
  finally
    List.Free;
  end;
end;

procedure ApplyFloat(Id: TAXWChpId; Value: double);
var
  List: TAXWCHPX;
  CP1,CP2: integer;
begin
  if FSelections.IsZeroLen and FSelections.FirstPos.Para.WordAtPos(FSelections.FirstPos.CharPos,CP1,CP2) then begin
    FSelections.CopyToTemp;
    FSelections.TempPos1.CharPos := CP1;
    FSelections.TempPos2.CharPos := CP2 + 1;
  end
  else
    FSelections.CopyToTemp;

  List := TAXWCHPX.Create(FMasterCHP);
  try
    List.AddFloat(Integer(Id),Value);
    FLogParas.ApplyFormat(List,FSelections.TempPos1,FSelections.TempPos2);
    FCaretPos.SetDocPropsAtPos;
  finally
    List.Free;
  end;
end;

procedure ApplyParaInteger(Id: TAXWPapId; Value: integer);
var
  List: TAXWPAPX;
begin
  List := TAXWPAPX.Create(FMasterPAP);
  try
    List.AddInteger(Integer(Id),Value);
    FLogParas.ApplyParaFormat(List,FSelections.FirstPos.Para,FSelections.LastPos.Para);
    FCaretPos.SetDocPropsAtPos;
  finally
    List.Free;
  end;
end;

begin
  if not FSelections.HasSelection then
    FSelections.SetToCaret;

  case CMD of
    axcFormatCHP          : begin
//      FTempCHPX.Clear;
//      FLogParas.ApplyFormatUncond(FTempCHPX,FSelections.FirstPos,FSelections.LastPos);
      FLogParas.ApplyFormat(FTempCHPX,FSelections.FirstPos,FSelections.LastPos);
    end;
    axcFormatBold         : ApplyBoolean(axciBold);
    axcFormatItalic       : ApplyBoolean(axciItalic);
    axcFormatUnderline    : ApplyBoolean(axciUnderline);

    axcFormatDecFont      : begin
      FCaretPos.SetDocPropsAtPos;
      CHPX := FCaretPos.Para.CHPXAtPos(FCaretPos.CharPos);
      if CHPX <> Nil then
        V := CHPX.Size
      else
        V := FMasterCHP.Size;
      if V > 3 then
        ApplyFloat(axciSize,V - 1);
    end;
    axcFormatIncFont      : begin
      CHPX := FCaretPos.Para.CHPXAtPos(FCaretPos.CharPos);
      if CHPX <> Nil then
        V := CHPX.Size
      else
        V := FMasterCHP.Size;
      if V < 72 then
        ApplyFloat(axciSize,V + 1);
    end;

    axcClearAllCharFmt    : ClearCHPFormatting;

    axcFormatParaLeft     : ApplyParaInteger(axpiAlignment,Integer(axptaLeft));
    axcFormatParaCenter   : ApplyParaInteger(axpiAlignment,Integer(axptaCenter));
    axcFormatParaRight    : ApplyParaInteger(axpiAlignment,Integer(axptaRight));
    axcFormatParaJustify  : ApplyParaInteger(axpiAlignment,Integer(axptaJustify));
  end;
end;

procedure TAXWLogDocEditor.InsertText(S: AxUCString; InCRLeft: boolean = False);
begin
  if FSelections.HasSelection then
    CmdDelete(axcDelSelection);
  if FLogParas.Count <= 0 then begin
    InsertPara(S);
//    CmdMove(axcMoveCharRight);
  end
  else
    FCaretPos.Para.InsertText(S,FCaretPos.CharPos,InCRLeft);
  PaginateCaretPara;
  CmdMove(axcMoveCharRight);
end;

function TAXWLogDocEditor.IsEmpty: boolean;
begin
  Result := FLogParas.Count <= 0;
end;

procedure TAXWLogDocEditor.CmdMove(Cmd: TAXWCommand);
begin
  ClearSelection;
  DoMoveCaret(Cmd);
end;

procedure TAXWLogDocEditor.CmdSelect(Cmd: TAXWCommand);
begin
  if not FSelections.HasSelection then
    FSelections.DropAnchor
  else
    FSelections.CopyToTemp;
  case Cmd of
    axcSelectCharLeft        : DoMoveCaret(axcMoveCharLeft);
    axcSelectCharRight       : DoMoveCaret(axcMoveCharRight);
    axcSelectWordLeft        : DoMoveCaret(axcMoveWordLeft);
    axcSelectWordRight       : DoMoveCaret(axcMoveWordRight);
    axcSelectParaUp          : DoMoveCaret(axcMoveParaUp);
    axcSelectParaDown        : DoMoveCaret(axcMoveParaDown);
    axcSelectLineUp          : DoMoveCaret(axcMoveLineUp);
    axcSelectLineDown        : DoMoveCaret(axcMoveLineDown);
    axcSelectEndOfLine       : DoMoveCaret(axcMoveEndOfLine);
    axcSelectBeginningOfLine : DoMoveCaret(axcMoveBeginningOfLine);
    axcSelectPgUp            : DoMoveCaret(axcMovePgUp);
    axcSelectPgDown          : DoMoveCaret(axcMovePgDown);
    axcSelectTopOfNextPg     : DoMoveCaret(axcMoveTopOfNextPg);
    axcSelectTopOfPrevPg     : DoMoveCaret(axcMoveTopOfPrevPg);
    axcSelectEndOfDoc        : DoMoveCaret(axcMoveEndOfDoc);
    axcSelectBeginningOfDoc  : DoMoveCaret(axcMoveBeginningOfDoc);

    axcSelectWord            : ;
    axcSelectPara : begin
      FSelections.Anchor.MoveToBeginningOfPara;
      DoMoveCaret(axcMoveEndOfPara);
    end;
    axcSelectAll : begin
      FSelections.Anchor.MoveToBeginningOfDoc;
      FCaretPos.MoveToEndOfDoc;
    end
    else
      Exit;
  end;
  FSelections.Update;

  FRepaint := True;
end;

procedure TAXWLogDocEditor.ReadTextfile(Filename: AxUCString);
var
  Stream: TFileStream;
begin
  Stream := TFileStream.Create(Filename,fmOpenRead,fmShareDenyNone);
  try
    ReadTextStream(Stream);
  finally
    Stream.Free;
  end;
end;

procedure TAXWLogDocEditor.ReadTextStream(Stream: TStream);
var
  p1,p2,Sz,L: integer;
//  Timestamp: longword;
  C: AxUCChar;
  S,S2: AxUCString;
  Buf: PByteArray;
begin
  Clear;

  GetMem(Buf,Stream.Size);
  try
    AppendPara('');
    p1 := 1;
    L := Stream.Read(Buf^,Stream.Size);
//    Timestamp := GetTickCount + 20;
    SetLength(S,L);
    p2 := 1;
    while p2 <= L do begin
      S[p2] := AxUCChar(Buf[p2 - 1]);
      C := S[p2];
      if CharInSet(C,[#10,#13]) then begin
        if (C = #13) and (p2 < L) and (AxUCChar(Buf[p2]) = #10) then
          Sz := 2
        else if (C = #10) and (p2 > 1) and (AxUCChar(Buf[p2 - 2]) = #13) then
          Sz := 2
        else
          Sz := 1;

        S2 := Copy(S,p1,p2 - p1);
        if S2 = '' then
          AppendPara('')
        else
          AppendTextSimple(S2);

        // Signal when reading lengthy docs. Display the read content.
{
        if (L > 100000) and (GetTickCount > Timestamp) then begin
          Timestamp := MAXINT;

          FLogParas.EndUpdate;
          FLogParas.SyncSimple(asosAll,Nil,0);
          FLogParas.BeginUpdate;
        end;
}
        p1 := p2 + Sz;
        Inc(p2,Sz);
      end
      else
        Inc(p2);
    end;
  finally
    FreeMem(Buf);
  end;
end;

procedure TAXWLogDocEditor.InsertPara(S: AxUCString);
var
  Para: TAXWLogPara;
  CHPX: TAXWCHPX;
begin
  if FCaretPos.Para <> Nil then
    Para := FLogParas.Insert(FCaretPos.Para.Index)
  else
    Para := FLogParas.Add;
  CHPX := TAXWCHPX.Create(FMasterCHP);
  try
    FLogParas.MasterCHP.CopyToCHPX(CHPX);
    Para.AppendText(CHPX,S);
  finally
    CHPX.Free;
  end;
end;

procedure TAXWLogDocEditor.SetEmpty;
begin
  ClearData;
  BeginUpdate;
//  AppendPara('');
  EndUpdate;
  FCaretPos.Para := Nil;
end;

procedure TAXWLogDocEditor.Undo;
begin
  if FUndo.Count <= 0 then
    Exit;
  case FUndo.Peek.Operation of
    uroInsertText: ;
    uroInsertPara: ;
    uroMove      : ;
    uroDeleteText: InsertText(FUndo.Peek.Buffer,True);
  end;
  FUndo.Pop;
end;

procedure TAXWLogDocEditor.UpdateSelection(Direction: TAXWCommand);
begin
  case FSelections.Status of
    assNone: begin
      FSelections.DropAnchor;
      DoMoveCaret(Direction);
      FSelections.Update;
    end;
    assUpdating: begin
      DoMoveCaret(Direction);
      FSelections.Update;
    end;
  end;
end;

{ TAXWUndo }

procedure TAXWUndo.AddDelete(Text: AxUCString);
var
  U: TAXWUndoItem;
begin
  U := TAXWUndoItem.Create;
  U.Buffer := Text;
  U.Operation := uroDeleteText;
  Push(U);
end;

procedure TAXWUndo.AddInsert(Text: AxUCString);
var
  U: TAXWUndoItem;
begin
  U := TAXWUndoItem.Create;
  U.Value := Length(Text);
  U.Operation := uroInsertText;
  Push(U);
end;

procedure TAXWUndo.AddMove(Distance: integer);
var
  U: TAXWUndoItem;
begin
  U := TAXWUndoItem.Create;
  U.Value := Distance;
  U.Operation := uroMove;
  Push(U);
end;

destructor TAXWUndo.Destroy;
var
  i: integer;
begin
  for i := 0 to List.Count - 1 do
    TAXWUndoItem(List[i]).Free;
  inherited;
end;

function TAXWUndo.Peek: TAXWUndoItem;
begin
  Result := TAXWUndoItem(inherited Peek);
end;

{ TAXWSelections }

procedure TAXWSelections.DropAnchor;
begin
  FAnchor.Assign(FCaretPos);
  FStatus := assUpdating;
end;

function TAXWSelections.GetFirstPos: TAXWLogPosition;
begin
  Result := Items[0].First;
end;

function TAXWSelections.GetLastPos: TAXWLogPosition;
begin
  Result := Items[0].Last;
end;

procedure TAXWSelections.BeginIterateCR;
begin
  if HasSelection then begin
    FIteratePara := GetFirstPos.Para;
    FIterateCR := GetFirstPos.Para.FindCharRun(GetFirstPos.CharPos);
    FIterateCREnd := GetLastPos.Para.FindCharRun(GetLastPos.CharPos);
  end
  else begin
    FIterateCR := Nil;
    FIteratePara := Nil;
  end;
end;

procedure TAXWSelections.CopyToTemp;
begin
  FTempPos1.Assign(GetFirstPos);
  FTempPos2.Assign(GetLastPos);
end;

constructor TAXWSelections.Create(Caret: TAXWLogPosition);
begin
  inherited Create;
  FCaretPos := Caret;
  FAnchor := TAXWLogPosition.Create;
//  FFirstPos := TAXWLogPosition.Create;
//  FLastPos := TAXWLogPosition.Create;
  FTempPos1 := TAXWLogPosition.Create;
  FTempPos2 := TAXWLogPosition.Create;
end;

destructor TAXWSelections.Destroy;
begin
  inherited;  // inherited calls Clear, and Clear clears FAnchor...
  FreeAndNil(FAnchor);
//  FreeAndNil(FFirstPos);
//  FreeAndNil(FLastPos);
  FreeAndNil(FTempPos1);
  FreeAndNil(FTempPos2);
end;

function TAXWSelections.HasSelection: boolean;
begin
  Result := (Status <> assNone) and GetLastPos.GreaterEqual(GetFirstPos);
end;

function TAXWSelections.IsEntireDoc: boolean;
begin
  Normalize;
  Result := (GetFirstPos.Para = GetFirstPos.Para.First) and (GetLastPos.Para = GetLastPos.Para.Last);
  if Result then
    Result := (GetFirstPos.CharPos = 1) and GetLastPos.CPIsAtEOP;
end;

function TAXWSelections.IsZeroLen: boolean;
begin
  Result := HasSelection and GetFirstPos.Equal(GetLastPos);
end;

function TAXWSelections.NextCR: boolean;
begin
  Result := HasSelection;
  if Result and (FIteratePara <> Nil) and (FIterateCR <> Nil) then begin
    FIterateCR := FIterateCR.Next;
    if FIterateCR = Nil then begin
      FIteratePara := FIteratePara.Next;
      if FIteratePara <> Nil then
        FIterateCR := FIteratePara.CRFirst;
    end;
  end;
  Result := (FIteratePara <> Nil) and (FIterateCR <> Nil);
  if Result then begin
    Result := (FIteratePara.Index <= GetLastPos.Para.Index) and (FIterateCR.Index <= FIterateCREnd.Index);
    if not Result then begin
      FIterateCR := Nil;
      FIteratePara := Nil;
    end;
  end;
end;

procedure TAXWSelections.Normalize;
begin
  if Count <= 0 then
    Add;
  if FCaretPos.Para.Index < FAnchor.Para.Index then begin
    GetFirstPos.Assign(FCaretPos);
    GetLastPos.Assign(FAnchor);
  end
  else if FCaretPos.Para.Index > FAnchor.Para.Index then begin
    GetFirstPos.Assign(FAnchor);
    GetLastPos.Assign(FCaretPos);
  end
  else begin
    if FCaretPos.CharPos > FAnchor.CharPos then begin
      GetFirstPos.Assign(FAnchor);
      GetLastPos.Assign(FCaretPos);
    end
    else begin
      GetFirstPos.Assign(FCaretPos);
      GetLastPos.Assign(FAnchor);
    end;
  end;
end;

procedure TAXWSelections.PopCaretPos(Paras: TAXWLogParas);
var
  Para: TAXWLogPara;
  CP: integer;
begin
  if FSavedParaIndex >= Paras.Count then begin
    Para := Paras[Paras.Count - 1];
    FSavedCharPos := Para.EndOfParaPos;
  end
  else
    Para := Paras[FSavedParaIndex];

  if FSavedCharPos >= Para.EndOfParaPos then
    CP := Para.EndOfParaPos
  else
    CP := FSavedCharPos;

  FCaretPos.Assign(Para);
  FCaretPos.CharPos := CP;
  FCaretPos.Update;
end;

procedure TAXWSelections.PushCaretPos;
begin
  FSavedParaIndex := FCaretPos.Para.Index;
  FSavedCharPos := FCaretPos.CharPos;
  FSavedCharPosIsEOP := FSavedCharPos = FCaretPos.Para.EndOfParaPos;
end;

procedure TAXWSelections.PushFirstPos;
begin
  FSavedParaIndex := GetFirstPos.Para.Index;
  FSavedCharPos := GetFirstPos.CharPos;
  FSavedCharPosIsEOP := FSavedCharPos = GetFirstPos.Para.EndOfParaPos;
end;

procedure TAXWSelections.Reset;
begin
  inherited Reset;
  Clear;
  FAnchor.Clear;
  FStatus := assNone;
end;

procedure TAXWSelections.SetToCaret;
begin
  if Count <= 0 then
    Add;
  GetFirstPos.Assign(FCaretPos);
  GetLastPos.Assign(FCaretPos);
  FStatus := assUpdating;
end;

procedure TAXWSelections.Update;
begin
  Normalize;
end;

end.
