unit XLSFmlaDebugger5;

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

uses Classes, Sysutils, Contnrs, Math,
     Xc12Utils5, Xc12Manager5,
     XLSUtils5, XLSCellMMU5, XLSFormulaTypes5, XLSEvaluate5, XLSEvaluateFmla5,
     XLSDecodeFmla5, XLSEncodeFmla5, XLSFormula5, XLSFmlaDebugData5;

type TFmlaDebugger = class(TObject)
protected
     FManager   : TXc12Manager;
     FSheetIndex: integer;
     FCol       : integer;
     FRow       : integer;
     FPtgs      : PXLSPtgs;
     FPtgsSz    : integer;
     FStack     : TStringList;
     FItems     : TFmlaDebugItems;
     FCurrItem  : TFmlaDebugItem;
     FFormula   : AxUCString;
     FFormulas  : TStringList;
     FLastEval  : AxUCString;
     FMaxSteps  : integer;
     FCurrSteps : integer;

     function  GetFormulaValue(const AFV: TXLSFormulaValue): AxUCString;
     procedure CollectFormulas;
     procedure DumpStack(AStack: TValueStack);
     procedure DumpFormulas;
     function  Decode(APtgs,AStopPtgs: PXLSPtgs; APtgsSz: integer): AxUCString;
     function  Encode(const AFormula: AxUCString; out APtgs: PXLSPtgs): integer;
     function  EvaluateSteps(const ASteps: integer): boolean;
public
     constructor Create(const AManager: TXc12Manager; const ASheetIndex,ACol,ARow: integer);
     destructor Destroy; override;

     function  AtFirst: boolean;
     function  AtLast: boolean;
     procedure MoveNext;
     procedure MovePrev;
     procedure MoveSteps(const ASteps: integer);

     function  StepsToPos(const APos: integer): integer;
     function  LastResult: AxUCString;

     property Formula: AxUCString read FFormula;
     property Formulas: TStringList read FFormulas;
     property LastEval: AxUCString read FLastEval;
     property Stack: TStringList read FStack;
     property CurrItem: TFmlaDebugItem read FCurrItem;
     end;

implementation

{ TFmlaDebugger }

function TFmlaDebugger.AtFirst: boolean;
begin
  Result := FCurrSteps < 1;
end;

function TFmlaDebugger.AtLast: boolean;
begin
  Result := FCurrSteps >= FMaxSteps;
end;

procedure TFmlaDebugger.CollectFormulas;
var
  Evaluator: TXLSFormulaEvaluator;
begin
  Evaluator := TXLSFormulaEvaluator.Create(FManager);
  try
    try
      FItems.Steps := MAXINT;
      FItems.CollectFormulas := True;
      Evaluator.DebugEvaluate(FItems,FSheetIndex,FCol,FRow,FPtgs,FPtgsSz,Nil);
      FItems.CollectFormulas := False;
    except

    end;
    DumpFormulas;
  finally
    Evaluator.Free;
  end;
end;

constructor TFmlaDebugger.Create(const AManager: TXc12Manager; const ASheetIndex,ACol,ARow: integer);
var
  S: AxUCString;
  Sz: integer;
  Ptgs: PXLSPtgs;
  Cell: TXLSCellItem;
  Cells: TXLSCellMMU;
  TempLS: AxUCChar;
begin
  FManager := AManager;
  FSheetIndex := ASheetIndex;
  FCol := ACol;
  FRow := ARow;

  // TODO Correct value for this = number of Ptgs in formula.
  FMaxSteps := 100;

  FItems := TFmlaDebugItems.Create;
  FStack := TStringList.Create;
  FFormulas := TStringList.Create;

  Cells := FManager.Worksheets[FSheetIndex].Cells;
  Cells.CalcDimensions;
  if Cells.FindCell(FCol,FRow,Cell) then begin
    if not Cells.IsFormulaCompiled(@Cell) then begin
      S := Cells.GetFormula(@Cell);
      Sz := Encode(S,Ptgs);
      if Sz > 0 then begin
        Cells.SetFormulaPtgs(FCol,FRow,@Cell,Ptgs,Sz);
        FreeMem(Ptgs);
      end;
    end;
    FPtgsSz := FManager.Worksheets[FSheetIndex].Cells.GetFormulaPtgs(@Cell,FPtgs);
    TempLS := FormatSettings.ListSeparator;
    FormatSettings.ListSeparator := ',';
    try
      FFormula := Decode(FPtgs,Nil,FPtgsSz);
    finally
      FormatSettings.ListSeparator := TempLS;
    end;
    Encode(FFormula,FPtgs);

    CollectFormulas;
  end
  else
    FPtgs := Nil;
end;

function TFmlaDebugger.Decode(APtgs,AStopPtgs: PXLSPtgs; APtgsSz: integer): AxUCString;
var
  Decoder: TXLSFormulaDecoder;
begin
  Decoder := TXLSFormulaDecoder.Create(FManager);
  try
    Result := Decoder.Decode(FManager.Worksheets[FSheetIndex].Cells,FSheetIndex,FPtgs,AStopPtgs,FPtgsSz,False);
  finally
    Decoder.Free;
  end;
end;

destructor TFmlaDebugger.Destroy;
begin
  FItems.Free;
  FStack.Free;
  FFormulas.Free;
  if FPtgs <> Nil then
    FreeMem(FPtgs);
  inherited;
end;

procedure TFmlaDebugger.DumpFormulas;
var
  i: integer;
  S1,S2: AxUCString;
begin
  for i := 0 to FItems.FmlaItems.Count - 1 do begin
    S1 := Decode(FPtgs,FItems.FmlaItems[i].Ptgs,FPtgsSz);
    S2 := GetFormulaValue(FItems.FmlaItems[i].Result);
    FFormulas.Add(S2 + #9 + S1);
  end;
end;

procedure TFmlaDebugger.DumpStack(AStack: TValueStack);
var
  i: integer;
  S: AxUCString;
begin
  FStack.Clear;
  for i := 0 to AStack.StackSize do begin
    S := GetFormulaValue(AStack.Peek(i));
    FStack.Add(S);
  end;
end;

function TFmlaDebugger.Encode(const AFormula: AxUCString; out APtgs: PXLSPtgs): integer;
var
  FH: TXLSFormulaHandler;
begin
  FItems.Clear;
  FH := TXLSFormulaHandler.Create(FManager);
  try
    Result := FH.EncodeFormula(AFormula,APtgs,FSheetIndex,FCol,FRow,FItems);
  finally
    FH.Free;
  end;
end;

function TFmlaDebugger.EvaluateSteps(const ASteps: integer): boolean;
var
  Evaluator: TXLSFormulaEvaluator;
begin
  Result := FPtgs <> Nil;
  if not Result then
    Exit;

  FItems.Steps := ASteps;

  Evaluator := TXLSFormulaEvaluator.Create(FManager);
  try
    Evaluator.DebugEvaluate(FItems,FSheetIndex,FCol,FRow,FPtgs,FPtgsSz,Nil);

    FCurrItem := FItems.Find(FItems.CurrPtgsOffs);

    DumpStack(Evaluator.Stack);

    FLastEval := Decode(FPtgs,FItems.CurrPtgs,FPtgsSz)
  finally
    Evaluator.Free;
  end;
end;

function TFmlaDebugger.StepsToPos(const APos: integer): integer;
var
  i: integer;
  MinLen,L: integer;
begin
  Result := -1;
  MinLen := MAXINT;
  for i := 0 to FItems.Count - 1 do begin
    if (APos >= FItems[i].P1) and (APos <= FItems[i].P2) then begin
      L := FItems[i].P2 - FItems[i].P1;
      if L < MinLen then begin
        MinLen := L;
        Result := FItems[i].Index;
      end;
    end;
  end;
  if Result >= 0 then
    Inc(Result);
end;

function TFmlaDebugger.GetFormulaValue(const AFV: TXLSFormulaValue): AxUCString;
var
  C,R: integer;
  S,S2: AxUCString;
begin
  Result := '';
  case AFV.RefType of
    xfrtNone    : ;
    xfrtRef     : Result := ColRowToRefStr(AFV.Col1,AFV.Row1) + ' = ';
    xfrtArea    : begin
      Result := AreaToRefStr(AFV.Col1,AFV.Row1,AFV.Col2,AFV.Row2) + ' = ';
      if (AFV.Col2 = XLS_MAXCOL) or (AFV.Row2 = XLS_MAXROW) then
        Result := '{...}'
      else begin
        S := '{';
        for R := AFV.Row1 to AFV.Row2 do begin
          for C := AFV.Col1 to AFV.Col2 do begin
            if not AFV.Cells.AsString(C,R,S2) then
              S2 := '';
            S := S + S2 + ';';
          end;
          S[Length(S)] := '\';
        end;
        S[Length(S)] := '}';
        Result := S;
      end;
    end;
    xfrtXArea   : ;
    xfrtAreaList: ;
    xfrtArray   : ;
    xfrtArrayArg: ;
  end;

  case AFV.ValType of
//    xfvtUnknown: Result := Result + '[???]';
{$ifdef DELPHI_6_OR_LATER}
    xfvtFloat  : Result := Result + FloatToStr(RoundTo(AFV.vFloat,-4));
{$endif}    
    xfvtBoolean: if AFV.vBool then Result := Result + G_StrTRUE else Result := Result + G_StrFALSE;
    xfvtError  : Result := Result + Xc12CellErrorNames[AFV.vErr];
    xfvtString : Result := Result + AFV.vStr;
  end;
end;

function TFmlaDebugger.LastResult: AxUCString;
begin
  if FStack.Count > 0 then
    Result := FStack[0]
  else
    Result := '';
end;

procedure TFmlaDebugger.MoveNext;
begin
  if FCurrSteps < FMaxSteps then begin
    Inc(FCurrSteps);
    EvaluateSteps(FCurrSteps);
  end;
end;

procedure TFmlaDebugger.MovePrev;
begin
  if FCurrSteps > 1 then begin
    Dec(FCurrSteps);
    EvaluateSteps(FCurrSteps);
  end;
end;

procedure TFmlaDebugger.MoveSteps(const ASteps: integer);
begin
  if (ASteps > 0) and (ASteps <= FMaxSteps) then begin
    FCurrSteps := ASteps;
    EvaluateSteps(FCurrSteps);
  end;
end;

end.
