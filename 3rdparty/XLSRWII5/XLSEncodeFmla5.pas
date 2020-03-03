unit XLSEncodeFmla5;

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
{$ifdef AXW}
  {$I AXW.Inc}
{$else}
  {$I XLSRWII.inc}
{$endif}

interface

// http://en.wikipedia.org/wiki/Shunting-yard_algorithm

// SUM(A+10-90,(Sheet1!B4,G7:H8),800,900)+Grymt(899)
// 3+4*2/(1-5)^2^3

uses Classes, SysUtils, Contnrs, Math, IniFiles,
     Xc12Utils5,
{$ifdef XLSRWII}
     Xc12Manager5, Xc12DataTable5,
     XLSFmlaDebugData5,
{$endif}
{$ifdef AXW}
     AXWXLSRWII,
{$endif}
     XLSUtils5, XLSTokenizer5, XLSFormulaTypes5;

type PAttributeData = ^TAttributeData;
     TAttributeData = packed record
     Attribute: word;
     Count: word;
     end;

type TXLSFormulaEncoder = class(TObject)
protected
     FManager     : TXc12Manager;

     FTokens      : TXLSTokenList;
     FPrevToken   : TXLSTokenType;
     // Attributes (spaces) are stored in a negativ offset as an TAttributeData.
     // PAttributeData(Integer(FBuffer) - SizeOf(TAttributeData))
     FBuffer      : Pointer;
     FBufSize     : integer;
     FBufPtr      : integer;
     FStack       : PXLSArrayPtgs;
     FStackPtr    : integer;
     // Stack for paranthesis (Left or function). TOS indicates if we are
     // in a paranthesis or function. The meaning of the Comma operator
     // is different if we are in a function (function arg) or a parathesis
     // (union operator).
     FPStack      : PXLSArrayPtgs;
     FPStackPtr   : integer;
     FPostfix     : PXLSArrayPtgs;
     FPostfixSz   : integer;
     FAttrData    : TAttributeData;
     FAttrCount   : integer;

     // Common ptgs.
     FPtgsRPar    : TXLSPtgs;
     FPtgsComma   : TXLSPtgs;
{$ifdef XLSRWII}
     FCurrTables  : TXc12Tables;
{$endif}
     // Used for identify names with local scope. If the formula is in a name,
     // FSheetId is -1.
     FSheetId     : integer;
     // Used by columns without table name is tables.
     FCol,FRow    : integer;

     // Used when entering formulas in editor.
     FStopOnError : boolean;
     // If FStopOnError is true, FDone is set to True in order to end encoding
     // and save the (first) parts of the ptgs that are valid.
     FDone        : boolean;

{$ifdef XLSRWII}
     FDebugItems  : TFmlaDebugItems;
{$endif}

     procedure Push(APtgs: PXLSPtgs); {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     function  Pop: PXLSPtgs; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     function  Peek: integer;
     function  PPeek: integer;
     function  Last(AN: integer): integer; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     function  NextToken(const AIndex: integer; const ACount: integer = 1): TXLSTokenType;
     procedure SetString(AToken: PXLSToken; APtgs: PXLSPtgs);
     function  GetString(APtgs: PXLSPtgs): AxUCString;
     function  Precedence(APtgs: PXLSPtgs): integer;
     function  AddCellRef(var AIndex: integer): PXLSPtgs;
     procedure AddPostfix(APtgs: PXLSPtgs);
     procedure AddOperator(APtgs: PXLSPtgs);
     function  PrecheckTable(var AIndex: integer): boolean;
{$ifdef XLSRWII}
     procedure AddTable(var AIndex: integer);
     function  SaveTable(AName: AxUCString): PXLSPtgsTable;
{$endif}

     function  AllocPtgs(ASize: integer; APtgs: integer): PXLSPtgs;
     function  ParseToken(AToken: PXLSToken; var AIndex: integer): boolean;
     function  InfixToPostfix: boolean;
     function  CopyBuffer(var APtgsSize: integer): PXLSPtgs;
public
     constructor Create(AManager: TXc12Manager);
     destructor Destroy; override;

     function  Encode(ATokens: TXLSTokenList; out APtgs: PXLSPtgs; const ASheetId,ACol,ARow: integer{$ifdef XLSRWII}; ADebug: TFmlaDebugItems = Nil{$endif}): integer;

{$ifdef XLSRWII}
     function  Dump: AxUCString;
     function  DumpStack: AxUCString;
{$endif}

     property StopOnError: boolean read FStopOnError write FStopOnError;
{$ifdef XLSRWII}
     property CurrTables: TXc12Tables read FCurrTables write FCurrTables;
{$endif}
     end;

function XLSEncodeFormula(AManager: TXc12Manager; const AFormula: AxUCString; out APtgs: PXLSPtgs; out APtgsSz: integer; const AStrictMode: boolean{$ifdef XLSRWII}; ADebug: TFmlaDebugItems = Nil{$endif}): boolean;
function IsSimpleName(AManager: TXc12Manager; APtgs: PXLSPtgs; APtgsSz: integer; AArea: PXLS3dCellArea): TXc12SimpleNameType;

implementation

function IsSimpleName(AManager: TXc12Manager; APtgs: PXLSPtgs; APtgsSz: integer; AArea: PXLS3dCellArea): TXc12SimpleNameType;
{$ifdef XLS_BIFF}
var
  i1,i2: integer;
  P: PXLSPtgs;
{$endif}
begin
  Result := xsntNone;
//  Result := xsntNone;
  if (APtgsSz = SizeOf(TXLSPtgsArea1d)) and (APtgs.Id = xptgArea1d) then begin
    AArea.SheetIndex := PXLSPtgsArea1d(APtgs).Sheet;
    AArea.Col1 := PXLSPtgsArea1d(APtgs).Col1;
    AArea.Col2 := PXLSPtgsArea1d(APtgs).Col2;
    AArea.Row1 := PXLSPtgsArea1d(APtgs).Row1;
    AArea.Row2 := PXLSPtgsArea1d(APtgs).Row2;
    Result := xsntArea;
  end
  else if (APtgsSz = SizeOf(TXLSPtgsRef1d)) and (APtgs.Id = xptgRef1d) then begin
    AArea.SheetIndex := PXLSPtgsRef1d(APtgs).Sheet;
    AArea.Col1 := PXLSPtgsRef1d(APtgs).Col;
    AArea.Col2 := PXLSPtgsRef1d(APtgs).Col;
    AArea.Row1 := PXLSPtgsRef1d(APtgs).Row;
    AArea.Row2 := PXLSPtgsRef1d(APtgs).Row;
    Result := xsntRef;
  end
{$ifdef XLS_BIFF}
  else if APtgs.Id in [xptg_EXCEL_97,xptg_ARRAYCONSTS_97] then begin
    P := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgs));
    if P.Id = xptgRef3d97 then begin
      if AManager._ExtNames97.Get3dSheets(PXLSPtgsRef3d97(P).ExtSheetIndex,i1,i2) then begin
        AArea.SheetIndex := i1;
        AArea.Col1 := PXLSPtgsRef3d97(P).Col;
        AArea.Col2 := PXLSPtgsRef3d97(P).Col;
        AArea.Row1 := PXLSPtgsRef3d97(P).Row;
        AArea.Row2 := PXLSPtgsRef3d97(P).Row;
        if i1 >= XLS_MAXSHEETS then
          Result := xsntError
        else
          Result := xsntRef;
      end
      else
        Result := xsntNone;
    end
    else if P.Id = xptgArea3d97 then begin
      if AManager._ExtNames97.Get3dSheets(PXLSPtgsArea3d97(P).ExtSheetIndex,i1,i2) then begin
        AArea.SheetIndex := i1;
        AArea.Col1 := PXLSPtgsArea3d97(P).Col1;
        AArea.Col2 := PXLSPtgsArea3d97(P).Col2;
        AArea.Row1 := PXLSPtgsArea3d97(P).Row1;
        AArea.Row2 := PXLSPtgsArea3d97(P).Row2;
        if i1 >= XLS_MAXSHEETS then
          Result := xsntError
        else
          Result := xsntArea;
      end
      else
        Result := xsntNone;
    end;
  end
{$endif}
  else
    Result := xsntNone;
end;

function XLSEncodeFormula(AManager: TXc12Manager; const AFormula: AxUCString; out APtgs: PXLSPtgs; out APtgsSz: integer; const AStrictMode: boolean{$ifdef XLSRWII}; ADebug: TFmlaDebugItems = Nil{$endif}): boolean;
var
  ErrCnt: integer;
  Tokenizer: TXLSTokenizer;
  List: TXLSTokenList;
  Encoder: TXLSFormulaEncoder;
begin
  ErrCnt := AManager.Errors.ErrorCount;
  Tokenizer := TXLSTokenizer.Create(AManager.Errors);
  try
    Tokenizer.StrictMode := AStrictMode;
    Tokenizer.R1C1Mode := False;
    Tokenizer.Formula := Uppercase(AFormula);

    List := Tokenizer.Parse;
    Result := List <> Nil;
    if Result then begin
      Encoder := TXLSFormulaEncoder.Create(AManager);
      try
        APtgsSz := Encoder.Encode(List,APtgs,0,0,0{$ifdef XLSRWII},ADebug{$endif});

        Result := (APtgsSz > 0) and (ErrCnt = AManager.Errors.ErrorCount);
      finally
        Encoder.Free;
      end;
      List.Free;
    end;
  finally
    Tokenizer.Free;
  end;
end;

{ TXLSFormulaEncoder }

procedure TXLSFormulaEncoder.AddPostfix(APtgs: PXLSPtgs);
begin
  if (APtgs.Id = xptgFuncArg) {and not (Last(0) in [xptgTableCol,xptgTableSpecial]) } then begin
    if PPeek in [xptgNone,xptgLPar] then begin
      if (Last(0) in XPtgsRefAreaType + [xptgOpISect,xptgOpUnion,xptgOpRange,xptgLPar]) and (Last(1) in XPtgsRefAreaType + [xptgOpISect,xptgOpUnion,xptgOpRange,xptgLPar]) then
        APtgs.Id := xptgOpUnion
      else begin
        FManager.Errors.Error('',XLSERR_FMLA_BADUSEOFOP);
        Exit;
      end;
    end
    else
      Exit;
  end;
  FPostfix[FPostfixSz] := APtgs;
  Inc(FPostfixSz);
end;

function CompareTableArgs(Item1, Item2: Pointer): Integer;
var
  V1,V2: integer;
begin
  if PXLSPtgs(Item1).Id = xptgTableSpecial then
    V1 := PXLSPtgsTableSpecial(Item1).SpecialId
  else
    V1 := 0;
  if PXLSPtgs(Item2).Id = xptgTableSpecial then
    V2 := PXLSPtgsTableSpecial(Item2).SpecialId
  else
    V2 := 0;

  Result := V1 - V2;
end;

// Possible special combinations
// -----------------------------
// All: [none]
// Data: Headers | Totals
// Headers: Data
// Totals: Data
// This row: [none]

{$ifdef XLSRWII}
procedure TXLSFormulaEncoder.AddTable(var AIndex: integer);
var
  S: AxUCString;
  sTbl: AxUCString;
  Id: integer;
  R1,R2: integer;
  Err: TXc12CellError;
  IdCol: integer;
  Tbl: TXc12Table;
  HasTblToken: boolean;
  ptgsTable: PXLSPtgsTable;
  ptgsTableCol: PXLSPtgsTableCol;
  ptgsTableSpecial: PXLSPtgsTableSpecial;
  Special: TXLSTableSpecialSpecifier;
begin
  ptgsTable := Nil;
//  Special := xtssNone;
  if FPrevToken = xttTable then begin
    ptgsTable := SaveTable(TokenAsString(FTokens[AIndex - 1]));
    if ptgsTable = Nil then
      Exit;
    ptgsTable.ArgCount := 1;
    Id := ptgsTable.TableId;
    HasTblToken := True;
  end
  else begin
    Id := FCurrTables.Find(FCol,FRow);
    if Id < 0 then begin
      FManager.Errors.Error('at ' + ColRowToRefStr(FCol,FRow),XLSERR_FMLA_UNKNOWNTABLE);
      Exit;
    end;
    HasTblToken := False;
  end;
  Tbl := FCurrTables[Id];

  while NextToken(AIndex,0) in [xttTableCol,xttTableSpecial] do begin
    if FTokens[AIndex].TokenType = xttTableCol then begin
      S := TokenAsString(FTokens[AIndex]);
      IdCol := FCurrTables.FindCol(S);
      if IdCol < 0 then begin
        FManager.Errors.Error(S,XLSERR_FMLA_UNKNOWNTABLECOL);
        Exit;
      end;

      PXLSPtgs(ptgsTableCol) := AllocPtgs(SizeOf(TXLSPtgsTableCol),xptgTableCol);
      ptgsTableCol.TableId := Id;
      ptgsTableCol.ColId := IdCol;

      ptgsTableCol.AreaPtg := xptgArea;
      ptgsTableCol.Col1 := Tbl.Ref.Col1 + Tbl.TableColumns[IdCol].Id - 1;
      ptgsTableCol.Row1 := Tbl.Ref.Row1 + Tbl.HeaderRowCount;
      ptgsTableCol.Col2 := Tbl.Ref.Col1 + Tbl.TableColumns[IdCol].Id - 1;
      ptgsTableCol.Row2 := Tbl.Ref.Row2 - Tbl.TotalsRowCount;

      AddPostfix(PXLSPtgs(ptgsTableCol));
    end
    else if FTokens[AIndex].TokenType = xttTableSpecial then begin
      R1 := 0;
      R2 := 0;
      S := AnsiUppercase(TokenAsString(FTokens[AIndex]));
      if S = 'ALL' then
        Special := xtssAll
      else if S = 'DATA' then
        Special := xtssData
      else if S = 'HEADERS' then
        Special := xtssHeaders
      else if S = 'TOTALS' then
        Special := xtssTotals
      else if S = 'THIS ROW' then
        Special := xtssThisRow
      else begin
        FManager.Errors.Error(S,XLSERR_FMLA_UNKNOWNTBLSPEC);
        Exit;
      end;

      Err := errUnknown;
      case Special of
        xtssAll    : begin
          R1 := Tbl.Ref.Row1;
          R2 := Tbl.Ref.Row2;
        end;
        xtssData   : begin
          R1 := Tbl.Ref.Row1 + Tbl.HeaderRowCount;
          R2 := Tbl.Ref.Row2 - Tbl.TotalsRowCount;
        end;
        xtssHeaders: begin
          if Tbl.HeaderRowCount > 0 then begin
            R1 := Tbl.Ref.Row1;
            R2 := Tbl.Ref.Row1 + Tbl.HeaderRowCount - 1;
          end
          else
            Err := errRef;
        end;
        xtssTotals : begin
          if Tbl.TotalsRowCount > 0 then begin
            R1 := Tbl.Ref.Row2 - Tbl.TotalsRowCount + 1;
            R2 := Tbl.Ref.Row2;
          end
          else
            Err := errRef;
        end;
        xtssThisRow: begin
          // Nothing to do. The row is where the formula is. Set by the evaluator.
        end;
      end;

      PXLSPtgs(ptgsTableSpecial) := AllocPtgs(SizeOf(TXLSPtgsTableSpecial),xptgTableSpecial);
      ptgsTableSpecial.SpecialId := Byte(Special);
      ptgsTableSpecial.Error := Byte(Err);

      ptgsTableSpecial.AreaId := xptgArea;
      ptgsTableSpecial.Col1 := Tbl.Ref.Col1;
      ptgsTableSpecial.Col2 := Tbl.Ref.Col2;

      ptgsTableSpecial.Row1 := R1;
      ptgsTableSpecial.Row2 := R2;

      AddPostfix(PXLSPtgs(ptgsTableSpecial));
    end
    else begin
      FManager.Errors.Error(S,XLSERR_FMLA_BADTABLE);
      Exit;
    end;

    Inc(AIndex);

    case NextToken(AIndex,0) of
      xttColon: begin
        AddOperator(AllocPtgs(SizeOf(TXLSPtgs),xptgOpRange));
        Inc(AIndex);
      end;
      xttComma: begin
        if HasTblToken then
          Inc(ptgsTable.ArgCount);
//        AddOperator(@FPtgsComma);
        Inc(AIndex);
      end;
      xttSpace: begin
        AddOperator(AllocPtgs(SizeOf(TXLSPtgs),xptgOpISect));
        Inc(AIndex);
      end;
    end;
  end;
  if HasTblToken and (FTokens[AIndex].TokenType <> xttRSqBracket) then
    FManager.Errors.Error(sTbl,XLSERR_FMLA_TABLESQBRACKET);
  if not HasTblToken then
    Dec(AIndex);
end;
{$endif}

function TXLSFormulaEncoder.AllocPtgs(ASize: integer; APtgs: integer): PXLSPtgs;
var
  P: PAttributeData;
begin
  if (FBufPtr + ASize + SizeOf(TAttributeData)) > FBufSize then
    // FBuffer cqan not be reallocated as FPostfix points into the buffer.
    raise XLSRWException.Create('Out of memory in Encode Formula');

  P := PAttributeData(NativeInt(FBuffer) + FBufPtr - SizeOf(TAttributeData));
  if FAttrData.Attribute <> 0 then begin
    P.Attribute := FAttrData.Attribute;
    P.Count := FAttrData.Count;
    FAttrData.Attribute := 0;
    Inc(FAttrCount);
  end
  else
    P.Attribute := 0;

  Result := PXLSPtgs(NativeInt(FBuffer) + FBufPtr);
  Result.Id := APtgs;
  Inc(FBufPtr,ASize + SizeOf(TAttributeData));
end;

function TXLSFormulaEncoder.CopyBuffer(var APtgsSize: integer): PXLSPtgs;
var
  i: integer;
  PtgsRes,Ptgs: PXLSPtgs;
  AllocSz: integer;

procedure CopyPtgs(ASize: integer);
var
  P: PAttributeData;
begin
  while (ASize + APtgsSize) > AllocSz do begin
    Inc(AllocSz,$FF);
    ReAllocMem(Result,AllocSz);

    PtgsRes := PXLSPtgs(NativeInt(Result) + APtgsSize);
  end;

  P := PAttributeData(NativeInt(Ptgs) - SizeOf(TAttributeData));
  if P.Attribute <> 0 then begin
    PXLSPtgsWS(PtgsRes).Id := P.Attribute;
    PXLSPtgsWS(PtgsRes).Count := P.Count;
    PtgsRes := PXLSPtgs(NativeInt(PtgsRes) + SizeOf(TXLSPtgsWS));
  end;
  System.Move(Ptgs^,PtgsRes^,ASize);

{$ifdef XLSRWII}
  if FDebugItems <> Nil then
    FDebugItems.UpdateOffs(Ptgs,NativeInt(PtgsRes) - NativeInt(Result),i);
{$endif}

  PtgsRes := PXLSPtgs(NativeInt(PtgsRes) + ASize);
  Inc(APtgsSize,ASize);
end;

begin
  APtgsSize := 0;
  AllocSz := $FF;
  GetMem(Result,AllocSz);
  PtgsRes := Result;

  for i := 0 to FPostfixSz - 1 do begin
    Ptgs := FPostfix[i];
    case Ptgs.Id of
      xptgLPar      : CopyPtgs(SizeOf(TXlsPtgs));
      xptgArray     : CopyPtgs(SizeOf(TXlsPtgsArray));

      xptgErr       : CopyPtgs(SizeOf(TXlsPtgsErr));
      xptgBool      : CopyPtgs(SizeOf(TXlsPtgsBool));
      xptgStr       : CopyPtgs(FixedSzPtgsStr + PXLSPtgsStr(Ptgs).Len * 2);

      xptgOpAdd,
      xptgOpSub,
      xptgOpUPlus,
      xptgOpUMinus,
      xptgOpMult,
      xptgOpDiv,
      xptgOpPower,
      xptgOpConcat,
      xptgOpLT,
      xptgOpLE,
      xptgOpEQ,
      xptgOpGE,
      xptgOpGT,
      xptgOpNE,
      xptgOpUnion,
      xptgOpRange,
      xptgOpPercent : CopyPtgs(SizeOf(TXlsPtgs));
      xptgOpIsect   : CopyPtgs(SizeOf(TXlsPtgsISect));

      xptgRef       : CopyPtgs(SizeOf(TXlsPtgsRef));
      xptgArea      : CopyPtgs(SizeOf(TXlsPtgsArea));
      xptgAreaErr   : CopyPtgs(SizeOf(TXlsPtgsArea));
      xptgRefErr    : CopyPtgs(SizeOf(TXlsPtgsRef));
      xptgRef1d     : CopyPtgs(SizeOf(TXlsPtgsRef1d));
      xptgRef1dErr  : CopyPtgs(SizeOf(TXlsPtgsRef1dError));
      xptgArea1d    : CopyPtgs(SizeOf(TXlsPtgsArea1d));
      xptgRef3d     : CopyPtgs(SizeOf(TXlsPtgsRef3d));
      xptgArea3d    : CopyPtgs(SizeOf(TXlsPtgsArea3d));
      xptgXRef1d    : CopyPtgs(SizeOf(TXlsPtgsXRef1d));
      xptgXRef3d    : CopyPtgs(SizeOf(TXlsPtgsXRef3d));
      xptgXArea1d   : CopyPtgs(SizeOf(TXlsPtgsXArea1d));
      xptgXArea3d   : CopyPtgs(SizeOf(TXlsPtgsXArea3d));

      xptgInt       : CopyPtgs(SizeOf(TXlsPtgsInt));
      xptgNum       : CopyPtgs(SizeOf(TXlsPtgsNum));
      xptgName      : CopyPtgs(SizeOf(TXlsPtgsName));
      xptgFunc      : CopyPtgs(SizeOf(TXlsPtgsFunc));
      xptgFuncVar   : CopyPtgs(SizeOf(TXlsPtgsFuncVar));
      xptgUserFunc  : CopyPtgs(FixedSzPtgsUserFunc + PXLSPtgsUserFunc(Ptgs).Len * 2);
      xptgFuncArg   : CopyPtgs(SizeOf(TXLSPtgs));
      xptgMissingArg: CopyPtgs(SizeOf(TXLSPtgs));

      xptgTable       : CopyPtgs(SizeOf(TXLSPtgsTable));
      xptgTableCol    : CopyPtgs(SizeOf(TXLSPtgsTableCol));
      xptgTableSpecial: CopyPtgs(SizeOf(TXLSPtgsTableSpecial));

      xptgWS      : CopyPtgs(SizeOf(TXlsPtgsWS));

      xptgArrayFmlaChild: CopyPtgs(SizeOf(TXLSPtgsArrayChildFmla));
      xptgDataTableFmlaChild: CopyPtgs(SizeOf(TXLSPtgsDataTableChildFmla));
      else          raise XLSRWException.CreateFmt('Unknown ptgs "%.2X"',[Ptgs.Id]);
    end;
  end;
  ReAllocMem(Result,APtgsSize);
end;

constructor TXLSFormulaEncoder.Create(AManager: TXc12Manager);
begin
  FManager := AManager;
  FBufSize := $FFFF;
  GetMem(FBuffer,FBufSize);
  FPtgsRPar.Id := xptgRPar;
  FPtgsComma.Id := xptgFuncArg;

  FStopOnError := True;
end;

destructor TXLSFormulaEncoder.Destroy;
begin
  FreeMem(FBuffer);
  if FStack <> Nil then
    FreeMem(FStack);
  if FPStack <> Nil then
    FreeMem(FPStack);
  if FPostfix <> Nil then
    FreeMem(FPostfix);
  inherited;
end;

// [T1.xlsx]Sheet1:Sheet3!B2:C3
//
//[workbook name] T1.xlsx
//[sheet name] Sheet1
//:
//[sheet name] Sheet3
//!
//[col] B
//[row] 2
//:
//[col] C
//[row] 3

// C:C,R:R,CR,CR:CR
// X-S, X-S:S, S, S:S
// tsError = Sheet1!#REF!
function TXLSFormulaEncoder.AddCellRef(var AIndex: integer): PXLSPtgs;
type TTokenSequence = (tsNone,tsCC,tsRR,tsCR,tsCRCR,tsError,tsXS,tsXSS,tsS,tsSS);

var
  StartIndex: integer;
  Token: PXLSToken;
  RefTokens: TTokenSequence;
  SheetTokens: TTokenSequence;
  SNamePos: integer;
  RefPos: integer;

function GetSheetIndex(const AName: AxUCString; out ASheetIndex: integer): boolean;
begin
{$ifdef XLSRWII}
  ASheetIndex := FManager.Worksheets.Find(AName);
{$endif}
{$ifdef AXW}
  ASheetIndex := FManager.FindBookmark(AName);
{$endif}
  Result := ASheetIndex >= 0;
  if not Result then
    FManager.Errors.Error(AName,XLSERR_FMLA_UNKNOWNSHEET);
end;

procedure SaveRef;
var
  i1,i2: integer;
{$ifdef XLSRWII}
  x: integer;
{$endif}

function EncodeCol(AToken: PXLSToken): integer;
begin
  if AToken.TokenType = xttAbsCol then
    Result := AToken.vInteger or COL_ABSFLAG
  else
    Result := AToken.vInteger;
end;

function EncodeRow(AToken: PXLSToken): integer;
begin
  if AToken.TokenType = xttAbsRow then
    Result := AToken.vInteger or Integer(ROW_ABSFLAG)
  else
    Result := AToken.vInteger;
end;

begin
  SNamePos := -1;
  case SheetTokens of
{$ifdef XLSRWII}
    tsXS : begin
      if TryStrToInt(TokenAsString(FTokens[StartIndex]),x) then
        Dec(x)
      else begin
        x := FManager.XLinks.FindBook(TokenAsString(FTokens[StartIndex]));
        if x < 0 then begin
          FManager.Errors.Error(TokenAsString(FTokens[StartIndex]),XLSERR_FMLA_BADWORKBOOKNAME);
          Exit;
        end;
      end;
      i1 := FManager.XLinks[x].ExternalBook.FindSheetName(TokenAsString(FTokens[StartIndex + 1]));
      if i1 < 0 then
        FManager.Errors.Error(TokenAsString(FTokens[StartIndex + 1]),XLSERR_FMLA_UNKNOWNSHEET)
      else begin
        if RefTokens = tsCR then begin
          Result := AllocPtgs(SizeOf(TXLSPtgsXRef1d),xptgXRef1d);
          PXLSPtgsXRef1d(Result).XBook := x;
          PXLSPtgsXRef1d(Result).Sheet := i1;
        end
        else begin
          Result := AllocPtgs(SizeOf(TXLSPtgsXArea1d),xptgXArea1d);
          PXLSPtgsXArea1d(Result).XBook := x;
          PXLSPtgsXArea1d(Result).Sheet := i1;
        end;
      end;
      SNamePos := FTokens[StartIndex].Pos.P1;
    end;
    tsXSS: begin
      x := FManager.XLinks.FindBook(TokenAsString(FTokens[StartIndex]));
      if x < 0 then
        FManager.Errors.Error(TokenAsString(FTokens[StartIndex]),XLSERR_FMLA_BADWORKBOOKNAME)
      else begin
        i1 := FManager.XLinks[x].ExternalBook.FindSheetName(TokenAsString(FTokens[StartIndex + 1]));
        if i1 < 0 then
          FManager.Errors.Error(TokenAsString(FTokens[StartIndex + 1]),XLSERR_FMLA_UNKNOWNSHEET);
        i2 := FManager.XLinks[x].ExternalBook.FindSheetName(TokenAsString(FTokens[StartIndex + 3]));
        if i2 < 0 then
          FManager.Errors.Error(TokenAsString(FTokens[StartIndex + 3]),XLSERR_FMLA_UNKNOWNSHEET);
        if (i1 >= 0) and (i2 >= 0) then begin
          if RefTokens = tsCR then begin
            Result := AllocPtgs(SizeOf(TXLSPtgsXRef3d),xptgXRef3d);
            PXLSPtgsXRef3d(Result).XBook := x;
            PXLSPtgsXRef3d(Result).Sheet1 := i1;
            PXLSPtgsXRef3d(Result).Sheet2 := i2;
          end
          else begin
            Result := AllocPtgs(SizeOf(TXLSPtgsXArea3d),xptgXArea3d);
            PXLSPtgsXArea3d(Result).XBook := x;
            PXLSPtgsXArea3d(Result).Sheet1 := i1;
            PXLSPtgsXArea3d(Result).Sheet2 := i2;
          end;
        end;
      end;
      SNamePos := FTokens[StartIndex].Pos.P1;
    end;
{$endif}
    tsS : begin
      if not GetSheetIndex(TokenAsString(FTokens[StartIndex]),i1) then
        Exit;
      if RefTokens = tsError then begin
        Result := AllocPtgs(SizeOf(TXLSPtgsRef1dError),xptgRef1dErr);
        PXLSPtgsRef1dError(Result).Error := Byte(errRef);
      end
      else if RefTokens = tsCR then begin
        Result := AllocPtgs(SizeOf(TXLSPtgsRef1d),xptgRef1d);
        PXLSPtgsRef1d(Result).Sheet := i1;
      end
      else begin
        Result := AllocPtgs(SizeOf(TXLSPtgsArea1d),xptgArea1d);
        PXLSPtgsArea1d(Result).Sheet := i1;
      end;
      SNamePos := FTokens[StartIndex].Pos.P1;
    end;
    tsSS: begin
      if not GetSheetIndex(TokenAsString(FTokens[StartIndex]),i1) then
        Exit;
      if not GetSheetIndex(TokenAsString(FTokens[StartIndex + 2]),i2) then
        Exit;
      if RefTokens = tsCR then begin
        Result := AllocPtgs(SizeOf(TXLSPtgsRef3d),xptgRef3d);
        PXLSPtgsRef3d(Result).Sheet1 := i1;
        PXLSPtgsRef3d(Result).Sheet2 := i2;
      end
      else begin
        Result := AllocPtgs(SizeOf(TXLSPtgsArea3d),xptgArea3d);
        PXLSPtgsArea3d(Result).Sheet1 := i1;
        PXLSPtgsArea3d(Result).Sheet2 := i2;
      end;
      SNamePos := FTokens[StartIndex].Pos.P1;
    end;
  end;
  case RefTokens of
    tsCC  : begin
      if Result = Nil then
        Result := AllocPtgs(SizeOf(TXLSPtgsArea),xptgArea);
      PXLSPtgsArea(Result).Col1 := EncodeCol(FTokens[AIndex - 2]);
      PXLSPtgsArea(Result).Row1 := 0;
      PXLSPtgsArea(Result).Col2 := EncodeCol(FTokens[AIndex]);
      PXLSPtgsArea(Result).Row2 := XLS_MAXROW;
      RefPos := FTokens[AIndex - 2].Pos.P1;
      AddPostfix(Result);
    end;
    tsRR  : begin
      if Result = Nil then
        Result := AllocPtgs(SizeOf(TXLSPtgsArea),xptgArea);
      PXLSPtgsArea(Result).Col1 := 0;
      PXLSPtgsArea(Result).Row1 := EncodeRow(FTokens[AIndex - 1]);
      PXLSPtgsArea(Result).Col2 := XLS_MAXCOL;
      PXLSPtgsArea(Result).Row2 := EncodeRow(FTokens[AIndex + 1]);
      RefPos := FTokens[AIndex - 1].Pos.P1;
      AddPostfix(Result);
    end;
    tsCR  : begin
      if Result = Nil then
        Result := AllocPtgs(SizeOf(TXLSPtgsRef),xptgRef);
      PXLSPtgsRef(Result).Col := EncodeCol(FTokens[AIndex - 1]);
      PXLSPtgsRef(Result).Row := EncodeRow(FTokens[AIndex]);
      RefPos := FTokens[AIndex - 1].Pos.P1;
      AddPostfix(Result);
    end;
    tsCRCR: begin
      if Result = Nil then
        Result := AllocPtgs(SizeOf(TXLSPtgsArea),xptgArea);
      PXLSPtgsArea(Result).Col1 := EncodeCol(FTokens[AIndex - 4]);
      PXLSPtgsArea(Result).Row1 := EncodeRow(FTokens[AIndex - 3]);
      PXLSPtgsArea(Result).Col2 := EncodeCol(FTokens[AIndex - 1]);
      PXLSPtgsArea(Result).Row2 := EncodeRow(FTokens[AIndex]);
      RefPos := FTokens[AIndex - 4].Pos.P1;
      AddPostfix(Result);
    end;
    tsError: begin
      if Result = Nil then
        Result := AllocPtgs(SizeOf(TXLSPtgsRef1dError),xptgRef1dErr);
      PXLSPtgsRef1dError(Result).Error := Byte(FTokens[AIndex - 1].TokenType) - Byte(xttConstErrNull) + 1;
      RefPos := FTokens[AIndex - 1].Pos.P1;
      AddPostfix(Result);
    end;
  end;
{$ifdef XLSRWII}
  if FDebugItems <> Nil then begin
    if SNamePos < 0 then
      FDebugItems.Add(Result,RefPos,FTokens[AIndex].Pos.P2)
    else
      FDebugItems.Add(Result,SNamePos,FTokens[AIndex].Pos.P2);
  end;
{$endif}
end;

procedure SaveName;
var
  i: integer;
  S: AxUCString;
  Id: integer;
  Ptgs: PXLSPtgs;
begin
  case SheetTokens of
    tsXS: begin
//      S := TokenAsString(FTokens[StartIndex]);
//      if S <> '0' then
//        raise XLSRWException.CreateFmt('Unknown workbook "%s" name',[S]);
      S := TokenAsString(FTokens[AIndex]);
{$ifdef XLSRWII}
      Id := FManager.Workbook.DefinedNames.FindId(S,-1);
{$endif}
{$ifdef AXW}
      Id := FManager.FindBookmark(S);
{$endif}
    end;
    tsS : begin
      if not GetSheetIndex(TokenAsString(FTokens[StartIndex]),i) then
        Exit;
      S := TokenAsString(FTokens[AIndex]);
{$ifdef XLSRWII}
      Id := FManager.Workbook.DefinedNames.FindIdInSheet(S,i);
{$endif}
{$ifdef AXW}
      Id := FManager.FindBookmark(S);
{$endif}
    end;
    else raise XLSRWException.Create('Unknown name reference');
  end;
  Ptgs := AllocPtgs(SizeOf(TXLSPtgsName),xptgName);
  if Id < 0 then
    PXLSPtgsName(Ptgs).NameId := XLS_NAME_UNKNOWN
  else
    PXLSPtgsName(Ptgs).NameId := Id;
  AddPostfix(Ptgs);
end;

begin
  StartIndex := AIndex;
  RefTokens := tsNone;
  SheetTokens := tsNone;
  if FTokens[AIndex].TokenType = xttWorkbookName then begin
    Inc(AIndex);
    if NextToken(AIndex) = xttColon then begin
      SheetTokens := tsXSS;
      Inc(AIndex,4);
    end
// Do not comment out below.
// Handles error xref error formulas like: [2]Sheet1!#REF!
    else if NextToken(AIndex) = xttExclamation then begin
      SheetTokens := tsXS;
      Inc(AIndex,2);
    end
// --------------------------
    else begin
     SheetTokens := tsXS;
      if NextToken(AIndex + 1) in [xttConstErrNull..xttConstErrGettingData] then
        Inc(AIndex)
      else
       Inc(AIndex,2);
    end;
  end
  else if FTokens[AIndex].TokenType = xttSheetName then begin
    if NextToken(AIndex) = xttColon then begin
      SheetTokens := tsSS;
      Inc(AIndex,4);
    end
    else begin
      SheetTokens := tsS;
//      if NextToken(AIndex + 1) in [xttConstErrNull..xttConstErrGettingData] then
//        Inc(AIndex)
//      else
        Inc(AIndex,2);
    end;
  end;

  Result := Nil;
  Token := FTokens[AIndex];
  case Token.TokenType of
    xttConstErrNull..xttConstErrGettingData: begin
      Inc(AIndex);
      RefTokens := tsError;
      SaveRef;
      Exit;
    end;
    xttRelCol,
    xttAbsCol      : begin
      case NextToken(AIndex) of
        xttRelRow,
        xttAbsRow: begin
          if NextToken(AIndex,2) = xttColon then begin
            Inc(AIndex,2);
            if NextToken(AIndex) in [xttRelCol,xttAbsCol] then begin
              Inc(AIndex);
              if NextToken(AIndex) in [xttRelRow,xttAbsRow] then begin
                Inc(AIndex);
                RefTokens := tsCRCR;
                SaveRef;
                Exit;
              end;
            end
            // A1:(A2)
            else begin
              Dec(AIndex);
              RefTokens := tsCR;
              SaveRef;
              Exit;
            end;
          end
          else begin
            Inc(AIndex);
            RefTokens := tsCR;
            SaveRef;
            Exit;
          end;
        end;
        xttColon: begin
          Inc(AIndex);
          case NextToken(AIndex) of
            xttRelCol,
            xttAbsCol: begin
              Inc(AIndex);
              RefTokens := tsCC;
              SaveRef;
              Exit;
            end;
          end;
        end;
      end;
    end;
    xttRelRow,
    xttAbsRow       : begin
      if (NextToken(AIndex) = xttColon) and (NextToken(AIndex + 1) in [xttRelRow,xttAbsRow]) then begin
        Inc(AIndex);
        RefTokens := tsRR;
        SaveRef;
        Exit;
      end;
    end;

    xttR1C1RelCol,
    xttR1C1AbsCol,
    xttR1C1RelRow,
    xttR1C1AbsRow   : ;
    xttName         : SaveName;
    else raise XLSRWException.CreateFmt('Unknown sheet reference type: "%.2X"',[Integer(Token.TokenType)]);
  end;
end;

procedure TXLSFormulaEncoder.AddOperator(APtgs: PXLSPtgs);
var
  Prec: integer;
begin
//  F_TEST := Dump;
//  F_TESTSTACK := DumpStack;

  if APtgs.Id = xptgRPar then begin
    while (FStackPtr >= 0) and not (FStack[FStackPtr].Id in XPtgsLParType) do
      AddPostfix(Pop);
    if FStackPtr >= 0 then
      AddPostfix(Pop);
  end
  else begin
    if (FStackPtr < 0) or (APtgs.Id in [xptgOpUPlus,xptgOpUMinus]) then
      Push(APtgs)
    else begin
      if APtgs.Id in XPtgsLParType then
        Prec := 100
      else
        Prec := Precedence(APtgs);
      if Precedence(FStack[FStackPtr]) >= Prec then begin
        while (FStackPtr >= 0) and (Precedence(FStack[FStackPtr]) >= Prec) do begin
          AddPostfix(Pop);
        end;
        Push(APtgs);
      end
      else
        Push(APtgs);
    end;
  end;
//  F_TEST := Dump;
//  F_TESTSTACK := DumpStack;
end;

{$ifdef XLSRWII}
function TXLSFormulaEncoder.Dump: AxUCString;
var
  i: integer;
  Ref: PXLSPtgsRef;
  Area: PXLSPtgsArea;
  Ref1d: PXLSPtgsRef1d;
  Ref1dErr: PXLSPtgsRef1dError;
  Area1d: PXLSPtgsArea1d;
  Ref3d: PXLSPtgsRef3d;
  Area3d: PXLSPtgsArea3d;
  XRef1d: PXLSPtgsXRef1d;
  XRef3d: PXLSPtgsXRef3d;
  XArea1d: PXLSPtgsXArea1d;
  XArea3d: PXLSPtgsXArea3d;
begin
  Result := '';
  for i := 0 to FPostfixSz - 1 do begin
    case FPostfix[i].Id of
      xptgLPar    : Result := Result + '(';

      xptgErr     : Result := Result + Xc12CellErrorNames[TXc12CellError(PXLSPtgsErr(FPostfix[i]).Value)];
      xptgBool    : begin
        if PXLSPtgsBool(FPostfix[i]).Value <> 0 then
          Result := Result + 'TRUE'
        else
          Result := Result + 'FALSE';
      end;
      xptgStr     : Result := Result + GetString(FPostfix[i]);

      xptgOpAdd   : Result := Result + '+';
      xptgOpSub   : Result := Result + '-';
      xptgOpUPlus : Result := Result + 'u+';
      xptgOpUMinus: Result := Result + 'u-';
      xptgOpMult  : Result := Result + '*';
      xptgOpDiv   : Result := Result + '/';
      xptgOpPower : Result := Result + '^';
      xptgOpConcat: Result := Result + '&';
      xptgOpLT    : Result := Result + '<';
      xptgOpLE    : Result := Result + '<=';
      xptgOpEQ    : Result := Result + '=';
      xptgOpGE    : Result := Result + '>=';
      xptgOpGT    : Result := Result + '>';
      xptgOpNE    : Result := Result + '<>';

      xptgOpIsect : begin
        if PXLSPtgsWS(FPostfix[i]).Count > 1 then
          Result := Result + Format('{I + ws * %d}',[PXLSPtgsWS(FPostfix[i]).Count])
        else
          Result := Result + '{I}';
      end;
      xptgOpUnion : Result := Result + '{U}';
      xptgOpRange : Result := Result + ':';

      xptgRef     : begin
        Ref := PXLSPtgsRef(FPostfix[i]);
        Result := Result + ColRowToRefStrEnc(Ref.Col,Ref.Row);
      end;
      xptgArea    : begin
        Area := PXLSPtgsArea(FPostfix[i]);
        Result := Result + AreaToRefStrEnc(Area.Col1,Area.Row1,Area.Col2,Area.Row2);
      end;
      xptgRef1d    : begin
        Ref1d := PXLSPtgsRef1d(FPostfix[i]);
        Result := Result + FManager.Worksheets[Ref1d.Sheet].Name + '!' + ColRowToRefStrEnc(Ref1d.Col,Ref1d.Row);
      end;
      xptgRef1dErr  : begin
        Ref1dErr := PXLSPtgsRef1dError(FPostfix[i]);
        Result := Result + Xc12CellErrorNames[TXc12CellError(Ref1dErr.Error)] + '!' + ColRowToRefStrEnc(Ref1dErr.Col,Ref1dErr.Row);
      end;
      xptgArea1d    : begin
        Area1d := PXLSPtgsArea1d(FPostfix[i]);
        Result := Result + FManager.Worksheets[Area1d.Sheet].Name + '!' + AreaToRefStrEnc(Area1d.Col1,Area1d.Row1,Area1d.Col2,Area1d.Row2);
      end;
      xptgRef3d     : begin
        Ref3d := PXLSPtgsRef3d(FPostfix[i]);
        Result := Result + FManager.Worksheets[Ref3d.Sheet1].Name + ':' + FManager.Worksheets[Ref3d.Sheet2].Name + '!' + ColRowToRefStrEnc(Ref3d.Col,Ref3d.Row);
      end;
      xptgArea3d    : begin
        Area3d := PXLSPtgsArea3d(FPostfix[i]);
        Result := Result + FManager.Worksheets[Area3d.Sheet1].Name + ':' + FManager.Worksheets[Area3d.Sheet2].Name + '!' + AreaToRefStrEnc(Area3d.Col1,Area3d.Row1,Area3d.Col2,Area3d.Row2);
      end;
      xptgXRef1d   : begin
        XRef1d := PXLSPtgsXRef1d(FPostfix[i]);
        Result := Result + '[' + FManager.XLinks[XRef1d.XBook].ExternalBook.Filename + ']' +
                  FManager.XLinks[XRef1d.XBook].ExternalBook.SheetNames[XRef1d.Sheet] + '!' +
                  ColRowToRefStrEnc(XRef1d.Col and not COL_ABSFLAG,XRef1d.Row and not ROW_ABSFLAG);
      end;
      xptgXRef3d   : begin
        XRef3d := PXLSPtgsXRef3d(FPostfix[i]);
        Result := Result + '[' + FManager.XLinks[XRef3d.XBook].ExternalBook.FileName + ']' +
                  FManager.XLinks[XRef3d.XBook].ExternalBook.SheetNames[XRef3d.Sheet1] + ':' +
                  FManager.XLinks[XRef3d.XBook].ExternalBook.SheetNames[XRef3d.Sheet2] + '!' +
                  ColRowToRefStrEnc(XRef3d.Col,XRef3d.Row);
      end;
      xptgXArea1d   : begin
        XArea1d := PXLSPtgsXArea1d(FPostfix[i]);
        Result := Result + '[' + FManager.XLinks[XArea1d.XBook].ExternalBook.Filename + ']' +
                  FManager.XLinks[XArea1d.XBook].ExternalBook.SheetNames[XArea1d.Sheet] + '!' +
                  AreaToRefStrEnc(XArea1d.Col1,XArea1d.Row1,XArea1d.Col2,XArea1d.Row2);
      end;
      xptgXArea3d   : begin
        XArea3d := PXLSPtgsXArea3d(FPostfix[i]);
        Result := Result + '[' + FManager.XLinks[XArea3d.XBook].ExternalBook.FileName + ']' +
                  FManager.XLinks[XArea3d.XBook].ExternalBook.SheetNames[XArea3d.Sheet1] + ':' +
                  FManager.XLinks[XArea3d.XBook].ExternalBook.SheetNames[XArea3d.Sheet2] + '!' +
                  AreaToRefStrEnc(XArea3d.Col1,XArea3d.Row1,XArea3d.Col2,XArea3d.Row2);
      end;

      xptgInt     : Result := Result + IntToStr(PXLSPtgsInt(FPostfix[i]).Value);
      xptgNum     : Result := Result + FloatToStr(PXLSPtgsNum(FPostfix[i]).Value);
      xptgName    : begin
        if PXLSPtgsName(FPostfix[i]).NameId = XLS_NAME_UNKNOWN then
          Result := Result + Xc12CellErrorNames[errName]
        else
          Result := Result + FManager.Workbook.DefinedNames[PXLSPtgsName(FPostfix[i]).NameId].Name;
      end;
      xptgFunc    : Result := Result + G_XLSExcelFuncNames.FindName(PXLSPtgsFunc(FPostfix[i]).FuncId);
      xptgFuncVar : Result := Result + G_XLSExcelFuncNames.FindName(PXLSPtgsFuncVar(FPostfix[i]).FuncId) + '(' + IntToStr(PXLSPtgsFuncVar(FPostfix[i]).ArgCount) +')';
      xptgUserFunc: Result := Result + GetString(FPostfix[i]) + '(' + IntToStr(PXLSPtgsUserFunc(FPostfix[i]).ArgCount) +')';
      xptgFuncArg : Result := Result + ',';
      xptgArray   : Result := Result + Format('{a[%d:%d]}',[PXLSPtgsArray(FPostfix[i]).Cols,PXLSPtgsArray(FPostfix[i]).Rows]);
      xptgArrayFmlaChild: Result := Result + '{array child}';
      xptgMissingArg: Result := Result + '{,}';

      xptgWS      : begin
        if PXLSPtgsWS(FPostfix[i]).Count > 1 then
          Result := Result + Format('{ws * %d}',[PXLSPtgsWS(FPostfix[i]).Count])
        else
          Result := Result + '{ws}';
      end;
      else          Result := Result + '{?}';
    end;
    Result := Result + ' ';
  end;
end;

function TXLSFormulaEncoder.DumpStack: AxUCString;
var
  i: integer;
begin
  Result := '';
  for i := 0 to FStackPtr do begin
    case FStack[i].Id of
      xptgLPar    : Result := Result + '(';

      xptgOpPower : Result := Result + '^';
      xptgOpAdd   : Result := Result + '+';
      xptgOpSub   : Result := Result + '-';
      xptgOpMult  : Result := Result + '*';
      xptgOpDiv   : Result := Result + '/';

      xptgOpUPlus : Result := Result + '+';
      xptgOpUMinus: Result := Result + '-';

      xptgFuncArg   : Result := Result + ',';

      xptgUserFunc: Result := Result + GetString(FStack[i]);
    end;
  end;
end;
{$endif}

function TXLSFormulaEncoder.Encode(ATokens: TXLSTokenList; out APtgs: PXLSPtgs; const ASheetId,ACol,ARow: integer{$ifdef XLSRWII}; ADebug: TFmlaDebugItems = Nil{$endif}): integer;
begin
  FTokens := ATokens;
  FSheetId := ASheetId;
  FCol := ACol;
  FRow := ARow;
{$ifdef XLSRWII}
  FDebugItems := ADebug;
{$endif}
  ReallocMem(FStack,FTokens.Count * SizeOf(PXLSPtgs));
  ReallocMem(FPStack,FTokens.Count * SizeOf(PXLSPtgs));
  ReallocMem(FPostfix,(FTokens.Count + 1) * SizeOf(PXLSPtgs));

  FStackPtr := -1;
  FPStackPtr := -1;
  FPostfixSz := 0;

  FAttrCount := 0;
  FAttrData.Attribute := 0;

  FDone := False;

  FBufPtr := SizeOf(TAttributeData);
  if InfixToPostfix then begin
    APtgs := CopyBuffer(Result);
//    Result := 0;
  end
  else
    Result := 0;
end;

function TXLSFormulaEncoder.GetString(APtgs: PXLSPtgs): AxUCString;
begin
  case APtgs.Id of
    xptgStr : begin
      SetLength(Result,PXLSPtgsStr(APtgs).Len);
      Move(PXLSPtgsStr(APtgs).Str[0],Pointer(Result)^,PXLSPtgsStr(APtgs).Len * 2);
    end;
    xptgUserFunc: begin
      SetLength(Result,PXLSPtgsUserFunc(APtgs).Len);
      Move(PXLSPtgsUserFunc(APtgs).Name[0],Pointer(Result)^,PXLSPtgsUserFunc(APtgs).Len * 2);
    end;
    else raise XLSRWException.CreateFmt('Unknown ptg: %d',[APtgs.Id]);
  end;
end;

function TXLSFormulaEncoder.InfixToPostfix: boolean;
var
  i: integer;
begin
  i := 0;
  while i < FTokens.Count do begin
    Result := ParseToken(FTokens[i],i);
    if not Result or FDone then
      Exit;
  end;
  Result := True;

  while FStackPtr >= 0 do
    AddPostfix(Pop);
end;

function TXLSFormulaEncoder.Last(AN: integer): integer;
begin
  if FPostfixSz > AN then
    Result := FPostfix[FPostfixSz - AN - 1].Id
  else
    Result := xptgNone;
end;

function TXLSFormulaEncoder.NextToken(const AIndex: integer; const ACount: integer): TXLSTokenType;
begin
  if AIndex < (FTokens.Count - ACount) then
    Result := FTokens[AIndex + ACount].TokenType
  else
    Result := xttNone;
end;

function TXLSFormulaEncoder.ParseToken(AToken: PXLSToken; var AIndex: integer): boolean;
var
  S: AxUCString;
  Id: integer;
  Ptgs: PXLSPtgs;

procedure DoError(AError: TXc12CellError);
begin
  Ptgs := AllocPtgs(SizeOf(TXLSPtgsErr),xptgErr);
  PXLSPtgsErr(Ptgs).Value := Byte(AError);
{$ifdef XLSRWII}
  if FDebugItems <> Nil then
    FDebugItems.Add(Ptgs,AToken.Pos.P1,AToken.Pos.P2);
{$endif}
  AddPostfix(Ptgs);
end;

procedure DoOperator(AOperator: byte);
begin
  Ptgs := AllocPtgs(SizeOf(TXLSPtgs),AOperator);
{$ifdef XLSRWII}
  if FDebugItems <> Nil then
    FDebugItems.Add(Ptgs,AToken.Pos.P1,AToken.Pos.P2);
{$endif}
  AddOperator(Ptgs);
end;

begin
  Result := True;
  case AToken.TokenType of
    xttWhitespace: begin
      Ptgs := AllocPtgs(SizeOf(TXLSPtgsWS),xptgWS);
      PXLSPtgsWS(Ptgs).Count := Min(AToken.vInteger,$FF);
      AddPostfix(Ptgs);
    end;
    xttComma      : begin
      if PPeek in [xptgArray] + XPtgsFuncType then begin
        if PPeek in XPtgsVarFuncType then
          PXLSPtgsFuncVar(FPStack[FPStackPtr]).ArgCount := PXLSPtgsFuncVar(FPStack[FPStackPtr]).ArgCount + 1
        else if PPeek = xptgArray then
          PXLSPtgsArray(FPStack[FPStackPtr]).Cols := PXLSPtgsArray(FPStack[FPStackPtr]).Cols + 1;

        AddOperator(@FPtgsComma);

        if FPrevToken = xttComma then begin
          Ptgs := AllocPtgs(SizeOf(TXLSPtgs),xptgMissingArg);
          AddPostfix(Ptgs);
        end;
      end
      else
        AddOperator(@FPtgsComma);
    end;
    xttColon      : DoOperator(xptgOpRange);
    xttSemicolon  : begin
      if PPeek in [xptgArray] + XPtgsFuncType then begin
        if PPeek = xptgArray then begin
          PXLSPtgsArray(FPStack[FPStackPtr]).Cols := PXLSPtgsArray(FPStack[FPStackPtr]).Cols + 1;
          PXLSPtgsArray(FPStack[FPStackPtr]).Rows := PXLSPtgsArray(FPStack[FPStackPtr]).Rows + 1;
        end;

        AddOperator(@FPtgsComma);

        if FPrevToken = xttSemicolon then begin
          Ptgs := AllocPtgs(SizeOf(TXLSPtgs),xptgMissingArg);
          AddPostfix(Ptgs);
        end;
      end
      else begin
        Ptgs := AllocPtgs(SizeOf(TXLSPtgs),xptgFuncArg);
        AddOperator(Ptgs);
      end;
    end;
    xttExclamation: ;
    xttLPar       : begin
      Ptgs := AllocPtgs(SizeOf(TXLSPtgs),xptgLPar);
      AddOperator(Ptgs);
    end;
    xttRPar       : begin
      if FPrevToken = xttComma then begin
        Ptgs := AllocPtgs(SizeOf(TXLSPtgs),xptgMissingArg);
        AddPostfix(Ptgs);
      end;
{$ifdef XLSRWII}
      if (FDebugItems <> Nil) and (PPeek in [xptgArray] + XPtgsFuncType) then
        FDebugItems.UpdatePos(FPStack[FPStackPtr],AToken.Pos.P2);
{$endif}
      AddOperator(@FPtgsRPar);
    end;
    xttLSqBracket : ;
    xttRSqBracket : ;
    xttAphostrophe: ;
    xttSpace      : begin
      FAttrData.Attribute := xptgWS;
      FAttrData.Count := AToken.vInteger;
    end;

    // Order must match order in Xc12Utils5
    xttConstErrNull       : DoError(errNull);
    xttConstErrDiv0       : DoError(errDiv0);
    xttConstErrValue      : DoError(errValue);
    xttConstErrRef        : DoError(errRef);
    xttConstErrName       : DoError(errName);
    xttConstErrNum        : DoError(errNum);
    xttConstErrNA         : DoError(errNA);
    xttConstErrGettingData: DoError(errGettingData);

    xttConstLogicalTrue,
    xttConstLogicalFalse  : begin
      Ptgs := AllocPtgs(SizeOf(TXLSPtgsBool),xptgBool);
      PXLSPtgsBool(Ptgs).Value := Integer(Boolean(AToken.TokenType = xttConstLogicalTrue));
{$ifdef XLSRWII}
      if FDebugItems <> Nil then
        FDebugItems.Add(Ptgs,AToken.Pos.P1,AToken.Pos.P2);
{$endif}
      AddPostfix(Ptgs);
    end;

    xttConstString        : begin
      Ptgs := AllocPtgs(FixedSzPtgsStr + AToken.vPStr[0] * 2,xptgStr);
      SetString(AToken,Ptgs);
{$ifdef XLSRWII}
      if FDebugItems <> Nil then
        FDebugItems.Add(Ptgs,AToken.Pos.P1,AToken.Pos.P2);
{$endif}
      AddPostfix(Ptgs);
    end;

    xttConstArrayBegin    : begin
      Ptgs := AllocPtgs(SizeOf(TXLSPtgsArray),xptgArray);
      if FTokens.NextTokenType(AIndex) = xttConstArrayEnd then begin
        PXLSPtgsArray(Ptgs).Cols := 0;
        PXLSPtgsArray(Ptgs).Rows := 0;
      end
      else begin
        PXLSPtgsArray(Ptgs).Cols := 1;
        PXLSPtgsArray(Ptgs).Rows := 1;
      end;
{$ifdef XLSRWII}
      if FDebugItems <> Nil then
        FDebugItems.TempPosStack.Add(AToken.Pos.P1 - 1);
{$endif}
      AddOperator(Ptgs);
    end;
    xttConstArrayEnd: begin
      if PPeek = xptgArray then begin
        if Frac(PXLSPtgsArray(FPStack[FPStackPtr]).Cols / PXLSPtgsArray(FPStack[FPStackPtr]).Rows) <> 0 then
          FManager.Errors.Error('',XLSERR_FMLA_BADARRAYCOLS)
        else
          PXLSPtgsArray(FPStack[FPStackPtr]).Cols := PXLSPtgsArray(FPStack[FPStackPtr]).Cols div PXLSPtgsArray(FPStack[FPStackPtr]).Rows;
      end;
      AddOperator(@FPtgsRPar);
{$ifdef XLSRWII}
      if (FDebugItems <> Nil) and (FDebugItems.TempPosStack.Count > 0) then
        FDebugItems.Add(FPostfix[FPostfixSz - 1],FDebugItems.TempPosStack.Pop,AToken.Pos.P2 - 1);
{$endif}
    end;

    xttConstNumber        : begin
      if (Frac(AToken.vFloat) = 0) and (AToken.vFloat <= MAXINT) and (AToken.vFloat >= XLS_MININT) then begin
        Ptgs := AllocPtgs(SizeOf(TXLSPtgsInt),xptgInt);
        PXLSPtgsInt(Ptgs).Value := Round(AToken.vFloat);
      end
      else begin
        Ptgs := AllocPtgs(SizeOf(TXLSPtgsNum),xptgNum);
        PXLSPtgsNum(Ptgs).Value := AToken.vFloat;
      end;
{$ifdef XLSRWII}
      if FDebugItems <> Nil then
        FDebugItems.Add(Ptgs,AToken.Pos.P1,AToken.Pos.P2);
{$endif}
      AddPostfix(Ptgs);
    end;

    xttOpHat         : DoOperator(xptgOpPower);
    xttOpMult        : DoOperator(xptgOpMult);
    xttOpDiv         : DoOperator(xptgOpDiv);
    xttOpPlus        : DoOperator(xptgOpAdd);
    xttOpMinus       : DoOperator(xptgOpSub);
    xttOpUPlus       : DoOperator(xptgOpUPlus);
    xttOpUMinus      : DoOperator(xptgOpUMinus);
    xttOpAmp         : DoOperator(xptgOpConcat);
    xttOpEqual       : DoOperator(xptgOpEQ);
    xttOpNotEqual    : DoOperator(xptgOpNE);
    xttOpLess        : DoOperator(xptgOpLT);
    xttOpLessEqual   : DoOperator(xptgOpLE);
    xttOpGreater     : DoOperator(xptgOpGT);
    xttOpGreaterEqual: DoOperator(xptgOpGE);
    xttOpPercent     : DoOperator(xptgOpPercent);
    xttOpRange       : DoOperator(xptgOpRange);
    xttOpISect       : begin
      Ptgs := AllocPtgs(SizeOf(TXLSPtgsISect),xptgOpISect);
      PXLSPtgsISect(Ptgs).SpacesCount := AToken.vInteger;
      PXLSPtgsISect(Ptgs).SpacesCount := AToken.vInteger;
      AddOperator(Ptgs);
    end;

    xttRelCol,
    xttAbsCol,
    xttRelRow,
    xttAbsRow,

    xttR1C1RelCol,
    xttR1C1AbsCol,
    xttR1C1RelRow,
    xttR1C1AbsRow    : Ptgs := AddCellRef(AIndex);

    xttName          : begin
{$ifdef XLSRWII}
      Id := FManager.Workbook.DefinedNames.FindId(TokenAsString(AToken),FSheetId);
{$endif}
{$ifdef AXW}
      Id := FManager.FindBookmark(TokenAsString(AToken));
{$endif}
      if Id < 0 then begin
        Ptgs := AllocPtgs(SizeOf(TXLSPtgsErr),xptgErr);
        PXLSPtgsErr(Ptgs).Value := Byte(errName);
//        FManager.Errors.Error(TokenAsString(AToken),XLSERR_FMLA_UNKNOWNNAME)
      end
      else begin
        Ptgs := AllocPtgs(SizeOf(TXLSPtgsName),xptgName);
        PXLSPtgsName(Ptgs).NameId := Id;
        AddPostfix(Ptgs);
      end;
    end;
    xttFuncName      : begin
      S := TokenAsString(AToken);
      Id := G_XLSExcelFuncNames.FindId(S);
      if Id >= 0 then begin
        if G_XLSExcelFuncNames.FixedArgs(Id) then begin
          Ptgs := AllocPtgs(SizeOf(TXLSPtgsFunc),xptgFunc);
          PXLSPtgsFunc(Ptgs).FuncId := Id;
        end
        else begin
          Ptgs := AllocPtgs(SizeOf(TXLSPtgsFuncVar),xptgFuncVar);
          PXLSPtgsFuncVar(Ptgs).FuncId := Id;
          if FTokens.NextTokenType(AIndex) = xttRPar then
            PXLSPtgsFuncVar(Ptgs).ArgCount := 0
          else
            PXLSPtgsFuncVar(Ptgs).ArgCount := 1;
        end;
      end
      else begin
        Ptgs := AllocPtgs(FixedSzPtgsUserFunc + AToken.vPStr[0] * 2,xptgUserFunc);
        if (AIndex < (FTokens.Count - 1)) and (FTokens[AIndex + 1].TokenType = xttRPar) then
          PXLSPtgsUserFunc(Ptgs).ArgCount := 0
        else
          PXLSPtgsUserFunc(Ptgs).ArgCount := 1;
        SetString(AToken,Ptgs);
      end;
{$ifdef XLSRWII}
      if FDebugItems <> Nil then
        FDebugItems.Add(Ptgs,AToken.Pos.P1,AToken.Pos.P2);
{$endif}
      AddOperator(Ptgs);
    end;

    xttSheetName,
    xttWorkbookName  : Ptgs := AddCellRef(AIndex);
{$ifdef XLSRWII}
    xttTable         : begin
      if not (NextToken(AIndex,1) in [xttTableCol,xttTableSpecial]) then
        SaveTable(TokenAsString(FTokens[AIndex]));
    end;
    xttTableCol,
    xttTableSpecial  : begin
      Result := PrecheckTable(AIndex);
      if Result then
        AddTable(AIndex);
    end;
{$endif}
    else begin
      if FStopOnError then begin
        FDone := True;
        Exit;
      end
      else
        raise XLSRWException.CreateFmt('Invalid xtt token %d',[Integer(AToken.TokenType)]);
    end;
  end;
  FPrevToken := AToken.TokenType;
  Inc(AIndex);
end;

function TXLSFormulaEncoder.Peek: integer;
begin
  if FStackPtr >= 0 then
    Result := FStack[FStackPtr].Id
  else
    Result := xptgNone;
end;

function TXLSFormulaEncoder.Pop: PXLSPtgs;
begin
  Result := FStack[FStackPtr];
  Dec(FStackPtr);
  if Result.Id in XPtgsLParType then
    Dec(FPStackPtr);
end;

function TXLSFormulaEncoder.PPeek: integer;
begin
  if FPStackPtr >= 0 then
    Result := FPStack[FPStackPtr].Id
  else
    Result := xptgNone;
end;

// http://en.wikipedia.org/wiki/Operators_in_C_and_C%2B%2B#Operator_precedence
function TXLSFormulaEncoder.Precedence(APtgs: PXLSPtgs): integer;
begin
  case APtgs.Id of
    xptgArray     : Result :=  1;
    xptgLPar      : Result :=  2;
    xptgUserFunc  : Result :=  3;
    xptgFunc      : Result :=  3;
    xptgFuncVar   : Result :=  4;
    xptgTable     : Result :=  5;
    xptgFuncArg   : Result :=  6;

    xptgOpRange   : Result := 38;
    xptgOpUnion   : Result := 38;
    xptgOPIsect   : Result := 38;

    xptgOpUPlus   : Result := 37;
    xptgOpUMinus  : Result := 37;

    xptgOpPercent : Result := 36;

    xptgOpPower   : Result := 35;

    xptgOpMult    : Result := 34;
    xptgOpDiv     : Result := 34;

    xptgOpAdd     : Result := 33;
    xptgOpSub     : Result := 33;

    xptgOpConcat  : Result := 32;

    xptgOpLT      : Result := 31;
    xptgOpLE      : Result := 31;
    xptgOpGT      : Result := 31;
    xptgOpGE      : Result := 31;

    xptgOpEQ      : Result := 30;
    xptgOpNE      : Result := 30;
    else          raise XLSRWException.CreateFmt('Unknown operator %.2X',[APtgs.Id]);
  end;
end;

function TXLSFormulaEncoder.PrecheckTable(var AIndex: integer): boolean;
type
  TTableArgData = record
  Token: TXLSTokenType;
  Special: integer;
  Index: integer;
  end;

var
  i: integer;
  S: AxUCString;
  Cnt: integer;
  Args: array of TTableArgData;

procedure QSortArgs(iLo, iHi: Integer) ;
var
  Lo,Hi: Integer;
  Pivot: integer;
  T: TTableArgData;
begin
  Lo := iLo;
  Hi := iHi;
  Pivot := Args[(Lo + Hi) div 2].Special;
  repeat
    while Args[Lo].Special < Pivot do Inc(Lo);
    while Args[Hi].Special > Pivot do Dec(Hi);
      if Lo <= Hi then begin
        T := Args[Lo];
        Args[Lo] := Args[Hi];
        Args[Hi] := T;
        Inc(Lo) ;
        Dec(Hi) ;
      end;
   until Lo > Hi;
   if Hi > iLo then QSortArgs(iLo,Hi);
   if Lo < iHi then QSortArgs(Lo,iHi);
 end;

begin
  Result := True;
  i := 0;
  Cnt := 0;
  SetLength(Args,32);
  while NextToken(AIndex,i) in [xttTableCol,xttTableSpecial] do begin
    Args[Cnt].Token := NextToken(AIndex,i);
    if Args[Cnt].Token = xttTableSpecial then begin
      S := AnsiUppercase(TokenAsString(FTokens[AIndex + i]));
      if S = 'ALL' then
        Args[Cnt].Special := Integer(xtssAll)
      else if S = 'DATA' then
        Args[Cnt].Special := Integer(xtssData)
      else if S = 'HEADERS' then
        Args[Cnt].Special := Integer(xtssHeaders)
      else if S = 'TOTALS' then
        Args[Cnt].Special := Integer(xtssTotals)
      else if S = 'THIS ROW' then
        Args[Cnt].Special := Integer(xtssThisRow)
      else begin
        FManager.Errors.Error(S,XLSERR_FMLA_UNKNOWNTBLSPEC);
        Result := False;
        Exit;
      end;
    end
    else
      Args[Cnt].Special := $FFFF;
    Args[Cnt].Index := AIndex + i;
    Inc(Cnt);
    if Cnt > Length(Args) then
      SetLength(Args,Length(Args) + 32);

    if (FPrevToken = xttTable) and (NextToken(AIndex,i + 1) = xttRSqBracket) then
      Break;

    if NextToken(AIndex,i + 1) in [xttColon,xttSpace,xttComma] then begin
      if NextToken(AIndex,i) = xttTableCol then
        Result := NextToken(AIndex,i + 1) in [xttColon,xttSpace,xttComma]
      else
        Result := NextToken(AIndex,i + 1) in [xttComma];
      Inc(i);
    end
    else
      Break;

    if not Result then begin
      FManager.Errors.Error('',XLSERR_FMLA_BADUSEOFTABLEOP);
      Exit;
    end;

    Inc(i);
  end;
  SetLength(Args,Cnt);

  if Cnt <= 1 then
    Exit;

  QSortArgs(0,High(Args));

  // TODO
  // Specials must be first, before columns. Excel sorts this by itself.
  Result := Args[0].Token = FTokens[AIndex].TokenType;
  if not Result then begin
    FManager.Errors.Error('',XLSERR_FMLA_TBLSPECNOTFIRST);
    Exit;
  end;


// Possible special combinations
// -----------------------------
// All: [none]
// Data: Headers | Totals
// Headers: Data
// Totals: Data
// This row: [none]
  if Args[0].Token = xttTableSpecial then begin
    case TXLSTableSpecialSpecifier(Args[0].Special) of
      xtssNone   : Result := False;
      xtssAll    : Result := not (TXLSTableSpecialSpecifier(Args[1].Special) in [xtssAll,xtssData,xtssHeaders,xtssTotals,xtssThisRow]);
      xtssData   : Result := not (TXLSTableSpecialSpecifier(Args[1].Special) in [xtssAll,xtssData,xtssThisRow]);
      xtssHeaders: Result := not (TXLSTableSpecialSpecifier(Args[1].Special) in [xtssAll,xtssHeaders,xtssTotals,xtssThisRow]);
      xtssTotals : Result := not (TXLSTableSpecialSpecifier(Args[1].Special) in [xtssAll,xtssHeaders,xtssTotals,xtssThisRow]);
      xtssThisRow: Result := not (TXLSTableSpecialSpecifier(Args[1].Special) in [xtssAll,xtssData,xtssHeaders,xtssTotals,xtssThisRow]);
    end;
    if Result and (Cnt > 2) then
      Result := Args[2].Token <> xttTableSpecial;

    if not Result then
      FManager.Errors.Error(S,XLSERR_FMLA_UNKNOWNTBLSPEC);
  end;
end;

procedure TXLSFormulaEncoder.Push(APtgs: PXLSPtgs);
begin
  Inc(FStackPtr);
  FStack[FStackPtr] := APtgs;
  if APtgs.Id in XPtgsLParType then begin
    Inc(FPStackPtr);
    FPStack[FPStackPtr] := APtgs;
  end;
end;

{$ifdef XLSRWII}
function TXLSFormulaEncoder.SaveTable(AName: AxUCString): PXLSPtgsTable;
var
  Id: integer;
  Sheet: integer;
begin
  Result := Nil;
  Sheet := $FFFF;
  Id := -1;
  if FCurrTables <> Nil then
    Id := FCurrTables.Find(AName);
  if Id < 0 then
    Id := FManager.FindTable(AName,FCurrTables,Sheet);
  if Id < 0 then
    FManager.Errors.Error(AName,XLSERR_FMLA_UNKNOWNTABLE)
  else begin
    PXLSPtgs(Result) := AllocPtgs(SizeOf(TXLSPtgsTable),xptgTable);
    Result.TableId := Id;
    Result.Sheet := Sheet;
    Result.ArgCount := 0;
    Result.AreaPtg := xptgArea;
    Result.Col1 := FCurrTables[Id].Ref.Col1;
    Result.Row1 := FCurrTables[Id].Ref.Row1 + FCurrTables[Id].HeaderRowCount;
    Result.Col2 := FCurrTables[Id].Ref.Col2;
    Result.Row2 := FCurrTables[Id].Ref.Row2 - FCurrTables[Id].TotalsRowCount;
    AddOperator(PXLSPtgs(Result));
  end;
end;
{$endif}

procedure TXLSFormulaEncoder.SetString(AToken: PXLSToken; APtgs: PXLSPtgs);
begin
  case APtgs.Id of
    xptgName,
    xptgStr: begin
      PXLSPtgsStr(APtgs).Len := AToken.vPStr[0];
      Move(AToken.vPStr[1],PXLSPtgsStr(APtgs).Str[0],AToken.vPStr[0] * 2);
    end;
    xptgUserFunc: begin
      PXLSPtgsUserFunc(APtgs).Len := AToken.vPStr[0];
      Move(AToken.vPStr[1],PXLSPtgsUserFunc(APtgs).Name[0],AToken.vPStr[0] * 2);
    end;
    else raise XLSRWException.CreateFmt('Unknown ptgs %d',[Aptgs.Id]);
  end;
end;

end.
