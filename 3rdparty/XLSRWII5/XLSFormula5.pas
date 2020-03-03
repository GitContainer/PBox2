unit XLSFormula5;

interface

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

uses Classes, SysUtils,
{$ifdef XLS_BIFF}
     BIFF_Utils5, BIFF_DecodeFormula5,
{$endif}
     Xc12Utils5, Xc12Manager5, Xc12DataTable5,
     XLSUtils5, XLSFormulaTypes5, XLSTokenizer5, XLSEncodeFmla5, XLSDecodeFmla5,
     XLSTSort, XLSEvaluate5, XLSEvaluateFmla5, XLSCalcChain5, XLSMMU5, XLSCellMMU5,
     XLSFmlaDebugData5;

type TXLSIteratePtgs = class(TObject)
protected
     FPtgs    : PXLSPtgs;
     FCurrPtgs: PXLSPtgs;
     FPtgsSz  : integer;
public
     constructor Create(APtgs: PXLSPtgs; const APtgsSz: integer);

     function  Next: boolean;
     function  CurrPtgs: PXLSPtgs;
     end;

type TXLSFormulaHandler = class(TObject)
private
     procedure SetTables(const Value: TXc12Tables);
     function  GetUserFuncEvent: TUserFunctionEvent;
     procedure SetUserFuncEvent(const Value: TUserFunctionEvent);
protected
     FManager   : TXc12Manager;
     FTokenizer : TXLSTokenizer;
     FDecoder   : TXLSFormulaDecoder;
     FCCBuilder : TXLSCalcChainBuilder;
     FEvaluator : TXLSFormulaEvaluator;
     FCalcChain : TXLSCalcChain;
     FFormulasUncalced: boolean;

     // Used by Encoder
     FCurrTables: TXc12Tables;

     function  CheckAdjustSheetIndex(const ACurrIndex,ASheetIndex,ACount: integer; out ANewIndex: integer): boolean; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     function  GetAreaErrorId(AId: integer): integer;
public
     constructor Create(AManager: TXc12Manager);
     destructor Destroy; override;

     procedure Calculate;
     procedure CalculateAndVerify;

     function  EncodeFormula(const AFormula: AxUCString; out APtgs: PXLSPtgs; const ASheetId: integer): integer; overload;
     function  EncodeFormula(const AFormula: AxUCString; out APtgs: PXLSPtgs; const ASheetId,ACol,ARow: integer; AStrictMode: boolean = True): integer; overload;
     function  EncodeFormula(const AFormula: AxUCString; out APtgs: PXLSPtgs; const ASheetId,ACol,ARow: integer; const ADebugItems: TFmlaDebugItems): integer; overload;

     function  DecodeFormula(ACells: TXLSCellMMU; const ASheetIndex: integer; APtgs: PXLSPtgs; const APtgsSize: integer; const ASaveToFile: boolean): AxUCString;
     function  EvaluateFormula(ASheetIndex,ACol,ARow: integer; APtgs: PXLSPtgs; APtgsSize: integer): TXLSFormulaValue;
     function  EvaluateFormulaStr(ACells: TXLSCellMMU; ASheetIndex,ACol,ARow: integer; APtgs: PXLSPtgs; APtgsSize: integer): AxUCString;
     procedure AdjustCell(const APtgs: PXLSPtgs; const APtgsSize, ADeltaCol, ADeltaRow: integer; ALockStartRow: boolean = False);
     procedure AdjustCellRowsChanged(const APtgs: PXLSPtgs; const APtgsSize,ASheetIndex,ARow,ADeltaRow: integer);
     procedure AdjustCellRowsChanged1d(const APtgs: PXLSPtgs; const APtgsSize,ASheetIndex,ARow,ADeltaRow: integer);
     procedure AdjustCellColsChanged(const APtgs: PXLSPtgs; const APtgsSize,ASheetIndex,ACol,ADeltaCol: integer);
     procedure AdjustCellColsChanged1d(const APtgs: PXLSPtgs; const APtgsSize,ASheetIndex,ACol,ADeltaCol: integer);
     procedure AdjustSheetIndex(const APtgs: PXLSPtgs; const APtgsSize, ASheetIndex, ACount: integer);

     function  CollectNames(ANames: TIntegerList; const APtgs: PXLSPtgs; const APtgsSize: integer): boolean;

     property CalcChain: TXLSCalcChain read FCalcChain;
     property CurrTables: TXc12Tables write SetTables;
     property FormulasUncalced: boolean read FFormulasUncalced write FFormulasUncalced;


     property OnUserFunction: TUserFunctionEvent read GetUserFuncEvent write SetUserFuncEvent;
     end;

function NameEncodeFormula(const AManager: TXc12Manager; const AFormula: AxUCString; out APtgs: PXLSPtgs; const ASheetId: integer): integer;

implementation

// Encoded non-cells formulas (names etc).
function NameEncodeFormula(const AManager: TXc12Manager; const AFormula: AxUCString; out APtgs: PXLSPtgs; const ASheetId: integer): integer;
var
  Tokens: TXLSTokenList;
  Encoder: TXLSFormulaEncoder;
  Tokenizer : TXLSTokenizer;
begin
  Tokenizer := TXLSTokenizer.Create(AManager.Errors);
  try
    Tokenizer.Formula := AFormula;
    Tokens := Tokenizer.Parse;

    if Tokens <> Nil then begin
//      Tokens.Dump(AManager.Errors.ErrorList);

      Encoder := TXLSFormulaEncoder.Create(AManager);
      try
        Encoder.CurrTables := Nil;
        Result := Encoder.Encode(Tokens,APtgs,ASheetId,-1,-1);
      finally
        Encoder.Free;
      end;
      Tokens.Free;
    end
    else
      Result := 0;
  finally
    Tokenizer.Free;
  end;
end;

{ TXLSIteratePtgs }

constructor TXLSIteratePtgs.Create(APtgs: PXLSPtgs; const APtgsSz: integer);
begin
  FPtgs := APtgs;
  FPtgsSz := APtgsSz;
end;

function TXLSIteratePtgs.CurrPtgs: PXLSPtgs;
begin
  Result := FCurrPtgs;
end;

function TXLSIteratePtgs.Next: boolean;
var
  L: integer;
  P: Pointer;
begin
  Result := FPtgsSz > 0;
  if Result then begin
    FCurrPtgs := FPtgs;

    case GetBasePtgs(FPtgs.Id) of
      xptgNone    : raise XLSRWException.Create('Illegal ptgs');

      xptgRef,
      xptgRefErr    : L := SizeOf(TXlsPtgsRef);
      xptgArea,
      xptgAreaErr   : L := SizeOf(TXlsPtgsArea);
      xptgStr       : L := FixedSzPtgsStr + PXLSPtgsStr(FPtgs).Len * 2;
      xptgUserFunc  : L := FixedSzPtgsUserFunc + PXLSPtgsUserFunc(FPtgs).Len * 2;

      xptg_EXCEL_97 : begin
        FPtgs := PXLSPtgs(NativeInt(FPtgs) + SizeOf(TXLSPtgs));
        Dec(FPtgsSz,SizeOf(TXLSPtgs));
        Exit;
      end;

      xptg_ARRAYCONSTS_97 : begin
        FPtgs := PXLSPtgs(NativeInt(FPtgs) + SizeOf(TXLSPtgs));

        L := PWord(FPtgs)^;
        Dec(FPtgsSz,L + SizeOf(TXLSPtgs) + SizeOf(word));
        FPtgs := PXLSPtgs(NativeInt(FPtgs) + SizeOf(word));

        Exit;
      end;

      xptgRef97      : L := SizeOf(TXlsPtgsRef97);
      xptgArea97     : L := SizeOf(TXlsPtgsArea97);

      xptgStr97      : begin
        L := 3;
        P := Pointer(NativeInt(FPtgs) + 1);
        if PByteArray(P)[1] = 0 then
          Inc(L,Byte(P^))
        else
          Inc(L,Byte(P^) * 2);
      end;

      xptgAttr97    : begin
        L := 1;
        P := Pointer(NativeInt(FPtgs) + 1);
        if (Byte(P^) and $04) = $04 then begin
          Inc(L);
          P := Pointer(NativeInt(P) + 1);
          Inc(L,(Word(P^) + 2) * SizeOf(word) - 3);
        end;
        Inc(L,3);
      end;

      xptgMemErr97,
      xptgMemArea97,
      xptgMemFunc97  : begin
        P := Pointer(NativeInt(FPtgs) + 1);
        L := 3 + PWord(P)^;
      end

      else            L := G_XLSPtgsSize[FPtgs.Id];
    end;
    if L < 0 then
      raise XLSRWException.CreateFmt('Unknown ptgs %.2X',[FPtgs.Id]);

    FPtgs := PXLSPtgs(NativeInt(FPtgs) + L);
    Dec(FPtgsSz,L);
  end;
end;

{ TXLSFormulaHandler }

procedure TXLSFormulaHandler.AdjustCell(const APtgs: PXLSPtgs; const APtgsSize, ADeltaCol, ADeltaRow: integer; ALockStartRow: boolean = False);
var
  C1,R1,C2,R2: integer;
  Mask: integer;
  Ptgs: PXLSPtgs;
  Iter: TXLSIteratePtgs;
begin
  Iter := TXLSIteratePtgs.Create(APtgs,APtgsSize);
  try
    while Iter.Next do begin
      Ptgs := Iter.CurrPtgs;
      case Ptgs.Id of
        xptgRef,
        xptgRef1d: begin
          C1 := 0;
          R1 := 0;
          if (PXlsPtgsRef(Ptgs).Col and COL_ABSFLAG) = 0 then begin
            C1 := PXlsPtgsRef(Ptgs).Col + ADeltaCol;
            PXlsPtgsRef(Ptgs).Col := C1;
          end;
          if ((PXlsPtgsRef(Ptgs).Row and ROW_ABSFLAG) = 0) and not ALockStartRow then begin
            R1 := Integer(PXlsPtgsRef(Ptgs).Row) + ADeltaRow;
            PXlsPtgsRef(Ptgs).Row := Longword(R1);
          end;

          if not InsideExtent(C1,R1) then
            Ptgs.Id := xptgRefErr;
        end;
        xptgArea,
        xptgArea1d: begin
          C1 := 0;
          R1 := 0;
          C2 := 0;
          R2 := 0;
          if (PXlsPtgsArea(Ptgs).Col1 and COL_ABSFLAG) = 0 then begin
            C1 := PXlsPtgsArea(Ptgs).Col1 + ADeltaCol;
            PXlsPtgsArea(Ptgs).Col1 := C1;
          end;
          if ((PXlsPtgsArea(Ptgs).Row1 and ROW_ABSFLAG) = 0) and not ALockStartRow then begin
            R1 := Integer(PXlsPtgsArea(Ptgs).Row1) + ADeltaRow;
            PXlsPtgsArea(Ptgs).Row1 := Longword(R1);
          end;
          if (PXlsPtgsArea(Ptgs).Col2 and COL_ABSFLAG) = 0 then begin
            C2 := PXlsPtgsArea(Ptgs).Col2 + ADeltaCol;
            PXlsPtgsArea(Ptgs).Col2 := C2;
          end;
          if (PXlsPtgsArea(Ptgs).Row2 and ROW_ABSFLAG) = 0 then begin
            R2 := Integer(PXlsPtgsArea(Ptgs).Row2) + ADeltaRow;
            PXlsPtgsArea(Ptgs).Row2 := Longword(R2);
          end;

          if not InsideExtent(C1,R1,C2,R2) then
            Ptgs.Id := xptgAreaErr;
        end;

        xptgRef97,
        xptgRefV97,
        xptgRefA97     : begin
          C1 := 0;
          R1 := 0;
          if (PXlsPtgsRef97(Ptgs).Col and $4000) <> 0 then begin
            Mask := PXlsPtgsRef97(Ptgs).Col and $C000;
            C1 := (PXlsPtgsRef97(Ptgs).Col and $3FFF) + ADeltaCol;
            PXlsPtgsRef97(Ptgs).Col := C1 or Mask;
          end;
          if ((PXlsPtgsRef97(Ptgs).Col and $8000) <> 0) and not ALockStartRow then begin
            R1 := Integer(PXlsPtgsRef97(Ptgs).Row) + ADeltaRow;
            PXlsPtgsRef97(Ptgs).Row := Longword(R1);
          end;

          if not InsideExtent(C1,R1) then
            Ptgs.Id := xptgRefErr97;
        end;
        xptgArea97,
        xptgAreaV97,
        xptgAreaA97    : begin
          C1 := 0;
          R1 := 0;
          C2 := 0;
          R2 := 0;
          if (PXlsPtgsArea97(Ptgs).Col1 and $4000) <> 0 then begin
            Mask := PXlsPtgsArea97(Ptgs).Col1 and $C000;
            C1 := (PXlsPtgsArea97(Ptgs).Col1 and $3FFF) + ADeltaCol;
            PXlsPtgsArea97(Ptgs).Col1 := C1 or Mask;
          end;
          if ((PXlsPtgsArea97(Ptgs).Col1 and $8000) <> 0) and not ALockStartRow then begin
            R1 := Integer(PXlsPtgsArea97(Ptgs).Row1) + ADeltaRow;
            PXlsPtgsArea97(Ptgs).Row1 := Longword(R1);
          end;
          if (PXlsPtgsArea97(Ptgs).Col2 and $4000) <> 0 then begin
            Mask := PXlsPtgsArea97(Ptgs).Col1 and $C000;
            C2 := (PXlsPtgsArea97(Ptgs).Col2 and $3FFF) + ADeltaCol;
            PXlsPtgsArea97(Ptgs).Col2 := C2 or Mask;
          end;
          if (PXlsPtgsArea97(Ptgs).Col2 and $8000) <> 0 then begin
            R2 := Integer(PXlsPtgsArea97(Ptgs).Row2) + ADeltaRow;
            PXlsPtgsArea97(Ptgs).Row2 := Longword(R2);
          end;

          if not InsideExtent(C1,R1,C2,R2) then
            Ptgs.Id := xptgAreaErr97;
        end;
      end;
    end;
  finally
    Iter.Free;
  end;
end;

procedure TXLSFormulaHandler.AdjustCellColsChanged(const APtgs: PXLSPtgs; const APtgsSize, ASheetIndex, ACol, ADeltaCol: integer);
var
  C1,C2: integer;
  Col2 : integer;
  Mask: word;
  Ptgs: PXLSPtgs;
  Iter: TXLSIteratePtgs;

function CalcAdjust(ACol2,D: integer): integer;
begin
  if ACol2 >= Col2 then
    Result := ADeltaCol
  else
    Result := -(ACol2 - ACol + D);
end;

begin
  Col2 := ACol - ADeltaCol - 1;

  Iter := TXLSIteratePtgs.Create(APtgs,APtgsSize);
  try
    while Iter.Next do begin
      Ptgs := Iter.CurrPtgs;

      case Ptgs.Id of
        xptgRef1d : if PXlsPtgsRef1d(Ptgs).Sheet <> ASheetIndex then Continue;
        xptgArea1d: if PXlsPtgsRef1d(Ptgs).Sheet <> ASheetIndex then Continue;
      end;

      case Ptgs.Id of
        xptgRef,
        xptgRef1d: begin
          Mask := PXlsPtgsRef(Ptgs).Col and COL_ABSFLAG;
          C1 := PXlsPtgsRef(Ptgs).Col and NOT_COL_ABSFLAG;
          if C1 >= ACol then begin
            Inc(C1,ADeltaCol);
            PXlsPtgsRef(Ptgs).Col := Longword(C1) or Mask;
          end;

          if not InsideExtent(C1,0) then begin
            case Ptgs.Id of
              xptgRef  : Ptgs.Id := xptgRefErr;
              xptgRef1d: Ptgs.Id := xptgRef1dErr;
            end;
          end;
        end;
        xptgArea,
        xptgArea1d: begin
          C1 := PXlsPtgsArea(Ptgs).Col1 and NOT_COL_ABSFLAG;
          C2 := PXlsPtgsArea(Ptgs).Col2 and NOT_COL_ABSFLAG;

          Mask := PXlsPtgsArea(Ptgs).Col1 and COL_ABSFLAG;
          if (C1 >= ACol) and (C2 <= Col2) then
            Ptgs.Id := GetAreaErrorId(Ptgs.Id)
          else begin
            if C1 > ACol then begin
              Inc(C1,CalcAdjust(C1,0));
              PXlsPtgsArea(Ptgs).Col1 := Longword(C1) or Mask;
            end;
            Mask := PXlsPtgsArea(Ptgs).Col2 and COL_ABSFLAG;
            if C2 >= ACol then begin
              Inc(C2,CalcAdjust(C2,1));
              PXlsPtgsArea(Ptgs).Col2 := Longword(C2) or Mask;
            end;

            if not InsideExtent(C1,0,C2,0) then
              Ptgs.Id := GetAreaErrorId(Ptgs.Id);
          end;
        end;

        xptgRef97,
        xptgRefA97,
        xptgRefV97      : begin
          Mask := PXlsPtgsRef97(Ptgs).Col and $C000;
          C1 := PXlsPtgsRef97(Ptgs).Col and $3FFF;
          if C1 >= ACol then begin
            Inc(C1,ADeltaCol);
            PXlsPtgsRef97(Ptgs).Col := C1 or Mask;
          end;

          if not InsideExtent(C1,0) then
            Ptgs.Id := xptgRefErr97;
        end;

        xptgRef3d97,
        xptgRef3dA97,
        xptgRef3dV97      : begin
          Mask := PXlsPtgsRef3d97(Ptgs).Col and $C000;
          C1 := PXlsPtgsRef3d97(Ptgs).Col and $3FFF;
          if C1 >= ACol then begin
            Inc(C1,ADeltaCol);
            PXlsPtgsRef3d97(Ptgs).Col := C1 or Mask;
          end;

          if not InsideExtent(C1,0) then
            Ptgs.Id := xptgRefErr97;
        end;

        xptgArea97,
        xptgAreaA97,
        xptgAreaV97   : begin
          C1 := PXlsPtgsArea97(Ptgs).Col1 and $3FFF;
          C2 := PXlsPtgsArea97(Ptgs).Col2 and $3FFF;
          if (C1 >= ACol) and (C2 <= Col2) then
            Ptgs.Id := xptgAreaErr97
          else begin
            Mask := PXlsPtgsArea97(Ptgs).Col1 and $C000;
            if C1 > ACol then begin
              Inc(C1,CalcAdjust(C1,0));
              PXlsPtgsArea97(Ptgs).Col1 := C1 or Mask;
            end;
            Mask := PXlsPtgsArea97(Ptgs).Col2 and $C000;
            if C2 >= ACol then begin
              Inc(C2,CalcAdjust(C2,1));
              PXlsPtgsArea97(Ptgs).Col2 := C2 or Mask;
            end;

            if not InsideExtent(C1,0,C2,0) or (C1 > C2) then
              Ptgs.Id := xptgAreaErr97;
          end;
        end;

        xptgArea3d97,
        xptgArea3dA97,
        xptgArea3dV97   : begin
          C1 := PXlsPtgsArea3d97(Ptgs).Col1 and $3FFF;
          C2 := PXlsPtgsArea3d97(Ptgs).Col2 and $3FFF;
          if (C1 >= ACol) and (C2 <= Col2) then
            Ptgs.Id := xptgAreaErr97
          else begin
            Mask := PXlsPtgsArea3d97(Ptgs).Col1 and $C000;
            if C1 >= ACol then begin
              Inc(C1,ADeltaCol);
              PXlsPtgsArea3d97(Ptgs).Col1 := C1 or Mask;
            end;
            Mask := PXlsPtgsArea3d97(Ptgs).Col2 and $C000;
            if C2 >= ACol then begin
              Inc(C2,ADeltaCol);
              PXlsPtgsArea3d97(Ptgs).Col2 := C2 or Mask;
            end;

            if not InsideExtent(C1,0,C2,0) or (C1 > C2) then
              Ptgs.Id := xptgAreaErr97;
          end;
        end;

      end;
    end;
  finally
    Iter.Free;
  end;
end;

procedure TXLSFormulaHandler.AdjustCellColsChanged1d(const APtgs: PXLSPtgs; const APtgsSize, ASheetIndex, ACol, ADeltaCol: integer);
var
  C1,C2: integer;
  Col2: integer;
  Mask: word;
  Ptgs: PXLSPtgs;
  Iter: TXLSIteratePtgs;

function CalcAdjust(ACol2,D: integer): integer;
begin
  if ACol2 >= Col2 then
    Result := ADeltaCol
  else
    Result := -(ACol2 - ACol + D);
end;

begin
  Col2 := ACol - ADeltaCol - 1;
  Iter := TXLSIteratePtgs.Create(APtgs,APtgsSize);
  try
    while Iter.Next do begin
      Ptgs := Iter.CurrPtgs;

      case Ptgs.Id of
        xptgRef1d : if PXlsPtgsRef1d(Ptgs).Sheet <> ASheetIndex then Continue;
        xptgArea1d: if PXlsPtgsRef1d(Ptgs).Sheet <> ASheetIndex then Continue;
      end;

      case Ptgs.Id of
        xptgRef1d: begin
          Mask := PXlsPtgsRef(Ptgs).Col and COL_ABSFLAG;
          C1 := PXlsPtgsRef(Ptgs).Col and NOT_COL_ABSFLAG;
          if C1 >= ACol then begin
            Inc(C1,ADeltaCol);
            PXlsPtgsRef(Ptgs).Col := Longword(C1) or Mask;
          end;

          if not InsideExtent(C1,0) then
            Ptgs.Id := xptgRefErr;
        end;
        xptgArea1d: begin
          C1 := PXlsPtgsArea(Ptgs).Col1 and NOT_COL_ABSFLAG;
          C2 := PXlsPtgsArea(Ptgs).Col2 and NOT_COL_ABSFLAG;
          if (C1 >= ACol) and (C2 <= Col2) then
            Ptgs.Id := GetAreaErrorId(Ptgs.Id)
          else begin
            Mask := PXlsPtgsArea(Ptgs).Col1 and COL_ABSFLAG;
            if C1 > ACol then begin
              Inc(C1,CalcAdjust(C1,0));
              PXlsPtgsArea(Ptgs).Col1 := Longword(C1) or Mask;
            end;
            Mask := PXlsPtgsArea(Ptgs).Col2 and COL_ABSFLAG;
            if C2 >= ACol then begin
              Inc(C2,CalcAdjust(C2,1));
              PXlsPtgsArea(Ptgs).Col2 := Longword(C2) or Mask;
            end;

            if not InsideExtent(C1,0,C2,0) or (C1 > C2) then
              Ptgs.Id := GetAreaErrorId(Ptgs.Id);
          end;
        end;
      end;
    end;
  finally
    Iter.Free;
  end;
end;

procedure TXLSFormulaHandler.AdjustCellRowsChanged(const APtgs: PXLSPtgs; const APtgsSize, ASheetIndex, ARow, ADeltaRow: integer);
var
  R1,R2: integer;
  Row2 : integer;
  Ptgs : PXLSPtgs;
  Iter : TXLSIteratePtgs;
  Mask:  longword;

function CalcAdjust(ARow2,D: integer): integer;
begin
  if ARow2 >= Row2 then
    Result := ADeltaRow
  else
    Result := -(ARow2 - ARow + D);
end;

begin
  Row2 := ARow - ADeltaRow - 1;

  Iter := TXLSIteratePtgs.Create(APtgs,APtgsSize);
  try
    while Iter.Next do begin
      Ptgs := Iter.CurrPtgs;

      case Ptgs.Id of
        xptgRef1d : if PXlsPtgsRef1d(Ptgs).Sheet <> ASheetIndex then Continue;
        xptgArea1d: if PXlsPtgsArea1d(Ptgs).Sheet <> ASheetIndex then Continue;
      end;

      case Ptgs.Id of
        xptgRef,
        xptgRef1d: begin
          Mask := PXlsPtgsRef(Ptgs).Row and ROW_ABSFLAG;
          R1 := PXlsPtgsRef(Ptgs).Row and NOT_ROW_ABSFLAG;
          if R1 >= ARow then begin
            Inc(R1,ADeltaRow);
            PXlsPtgsRef(Ptgs).Row := Longword(R1) or Mask;
          end;

          if not InsideExtent(0,R1) then
            Ptgs.Id := xptgRefErr;
        end;
        xptgArea,
        xptgArea1d: begin
          R1 := PXlsPtgsArea(Ptgs).Row1 and NOT_ROW_ABSFLAG;
          R2 := PXlsPtgsArea(Ptgs).Row2 and NOT_ROW_ABSFLAG;

          if (R1 >= ARow) and (R2 <= Row2) then
            Ptgs.Id := GetAreaErrorId(Ptgs.Id)
          else begin
            Mask := PXlsPtgsArea(Ptgs).Row1 and ROW_ABSFLAG;
            if R1 > ARow then begin
              Inc(R1,CalcAdjust(R1,0));
              PXlsPtgsArea(Ptgs).Row1 := Longword(R1) or Mask;
            end;

            Mask := PXlsPtgsArea(Ptgs).Row2 and ROW_ABSFLAG;
            if R2 >= ARow then begin
              Inc(R2,CalcAdjust(R2,1));
              PXlsPtgsArea(Ptgs).Row2 := Longword(R2) or Mask;
            end;

            if not InsideExtent(0,R1,0,R2) or (R1 > R2) then
              Ptgs.Id := GetAreaErrorId(Ptgs.Id);
          end;
        end;

        xptgRef97,
        xptgRefA97,
        xptgRefV97     : begin
          R1 := PXlsPtgsRef97(Ptgs).Row;
          if R1 >= ARow then begin
            Inc(R1,ADeltaRow);
            PXlsPtgsRef97(Ptgs).Row := Longword(R1);
          end;

          if not InsideExtent(0,R1) then
            Ptgs.Id := xptgRefErr97;
        end;

        xptgRef3d97,
        xptgRef3dA97,
        xptgRef3dV97     : begin
          R1 := PXlsPtgsRef3d97(Ptgs).Row;
          if R1 >= ARow then begin
            Inc(R1,ADeltaRow);
            PXlsPtgsRef3d97(Ptgs).Row := Longword(R1);
          end;

          if not InsideExtent(0,R1) then
            Ptgs.Id := xptgRefErr97;
        end;

        xptgArea97,
        xptgAreaA97,
        xptgAreaV97     : begin
          R1 := PXlsPtgsArea97(Ptgs).Row1;
          R2 := PXlsPtgsArea97(Ptgs).Row2;
          if (R1 >= ARow) and (R2 <= Row2) then
            Ptgs.Id := xptgAreaErr97
          else begin
            if R1 >= ARow then begin
              Inc(R1,CalcAdjust(R1,0));
              PXlsPtgsArea97(Ptgs).Row1 := Longword(R1);
            end;
            if R2 >= ARow then begin
              Inc(R2,CalcAdjust(R2,1));
              PXlsPtgsArea97(Ptgs).Row2 := Longword(R2);
            end;
            if not InsideExtent(0,R1,0,R2) or (R1 > R2) then
              Ptgs.Id := xptgAreaErr97;
          end;
        end;

        xptgArea3d97,
        xptgArea3dA97,
        xptgArea3dV97     : begin
          R1 := PXlsPtgsArea3d97(Ptgs).Row1;
          R2 := PXlsPtgsArea3d97(Ptgs).Row2;
          if (R1 >= ARow) and (R2 <= Row2) then
            Ptgs.Id := xptgAreaErr97
          else begin
            if R1 >= ARow then begin
              Inc(R1,CalcAdjust(R1,0));
              PXlsPtgsArea3d97(Ptgs).Row1 := Longword(R1);
            end;
            if R2 >= ARow then begin
              Inc(R2,CalcAdjust(R2,1));
              PXlsPtgsArea3d97(Ptgs).Row2 := Longword(R2);
            end;

            if not InsideExtent(0,R1,0,R2) or (R1 > R2) then
              Ptgs.Id := xptgAreaErr97;
          end;
        end;

      end;
    end;
  finally
    Iter.Free;
  end;
end;

procedure TXLSFormulaHandler.AdjustCellRowsChanged1d(const APtgs: PXLSPtgs; const APtgsSize, ASheetIndex, ARow, ADeltaRow: integer);
var
  R1,R2: integer;
  Row2 : integer;
  Ptgs: PXLSPtgs;
  Iter: TXLSIteratePtgs;
  Mask: word;

function CalcAdjust(ARow2,D: integer): integer;
begin
  if ARow2 >= Row2 then
    Result := ADeltaRow
  else
    Result := -(ARow2 - ARow + D);
end;

begin
  Row2 := ARow - ADeltaRow - 1;

  Iter := TXLSIteratePtgs.Create(APtgs,APtgsSize);
  try
    while Iter.Next do begin
      Ptgs := Iter.CurrPtgs;

      case Ptgs.Id of
        xptgRef1d : if PXlsPtgsRef1d(Ptgs).Sheet <> ASheetIndex then Continue;
        xptgArea1d: if PXlsPtgsRef1d(Ptgs).Sheet <> ASheetIndex then Continue;
      end;

      case Ptgs.Id of
        xptgRef1d: begin
          Mask := PXlsPtgsRef(Ptgs).Row and ROW_ABSFLAG;
          R1 := PXlsPtgsRef(Ptgs).Row and NOT_ROW_ABSFLAG;
          if R1 >= ARow then begin
            Inc(R1,ADeltaRow);
            PXlsPtgsRef(Ptgs).Row := Longword(R1) or Mask;
          end;

          if not InsideExtent(0,R1) then
            Ptgs.Id := xptgRefErr;
        end;
        xptgArea1d: begin
          R1 := PXlsPtgsArea(Ptgs).Row1 and NOT_ROW_ABSFLAG;
          R2 := PXlsPtgsArea(Ptgs).Row2 and NOT_ROW_ABSFLAG;

          if (R1 >= ARow) and (R2 <= Row2) then
            Ptgs.Id := GetAreaErrorId(Ptgs.Id)
          else begin
            Mask := PXlsPtgsArea(Ptgs).Row1 and ROW_ABSFLAG;
            if R1 > ARow then begin
              Inc(R1,CalcAdjust(R1,0));
              PXlsPtgsArea(Ptgs).Row1 := Longword(R1) or Mask;
            end;
            Mask := PXlsPtgsArea(Ptgs).Row2 and ROW_ABSFLAG;
            if R2 >= ARow then begin
              Inc(R2,CalcAdjust(R2,1));
              PXlsPtgsArea(Ptgs).Row2 := Longword(R2) or Mask;
            end;

            if not InsideExtent(0,R1,0,R2) or (R1 > R2) then
              Ptgs.Id := GetAreaErrorId(Ptgs.Id);
          end;
        end;
      end;
    end;
  finally
    Iter.Free;
  end;
end;

procedure TXLSFormulaHandler.AdjustSheetIndex(const APtgs: PXLSPtgs; const APtgsSize, ASheetIndex, ACount: integer);
var
  i,j: integer;
  Ok1,Ok2: boolean;
  Ptgs: PXLSPtgs;
  Iter: TXLSIteratePtgs;
begin
  Iter := TXLSIteratePtgs.Create(APtgs,APtgsSize);
  try
    while Iter.Next do begin
      Ptgs := Iter.CurrPtgs;
      case Ptgs.Id of
        xptgRef1d     : begin
          if CheckAdjustSheetIndex(PXLSPtgsRef1d(Ptgs).Sheet,ASheetIndex,ACount,i) then
            PXLSPtgsRef1d(Ptgs).Sheet := i
          else begin
            PXLSPtgsRef1d(Ptgs).Id := xptgRef1dErr;
            PXLSPtgsRef1dError(Ptgs).Error := Byte(errRef);
          end;
        end;
        xptgArea1d     : begin
          if CheckAdjustSheetIndex(PXLSPtgsArea1d(Ptgs).Sheet,ASheetIndex,ACount,i) then
            PXLSPtgsArea1d(Ptgs).Sheet := i
          else begin
            PXLSPtgsArea1d(Ptgs).Id := xptgRef1dErr;
            PXLSPtgsArea1dError(Ptgs).Error := Byte(errRef);
          end;
        end;
        xptgRef3d     : begin
          Ok1 := CheckAdjustSheetIndex(PXLSPtgsRef3d(Ptgs).Sheet1,ASheetIndex,ACount,i);
          Ok2 := CheckAdjustSheetIndex(PXLSPtgsRef3d(Ptgs).Sheet2,ASheetIndex,ACount,j);

          if Ok1 then
            PXLSPtgsRef3d(Ptgs).Sheet1 := i
          else begin
            PXLSPtgsRef3d(Ptgs).Id := xptgRef3dErr;
            PXLSPtgsRef3dError(Ptgs).Error1 := Byte(errRef);
          end;
          if Ok2 then
            PXLSPtgsRef3d(Ptgs).Sheet2 := j
          else begin
            PXLSPtgsRef3d(Ptgs).Id := xptgRef3dErr;
            PXLSPtgsRef3dError(Ptgs).Error2 := Byte(errRef);
          end;

          if Ok1 xor Ok2 then begin
            PXLSPtgsRef3d(Ptgs).Id := xptgRef3d;
            if Ok1 then
              PXLSPtgsRef3d(Ptgs).Sheet2 := ASheetIndex - 1
            else
              PXLSPtgsRef3d(Ptgs).Sheet1 := ASheetIndex + -ACount - 1;
          end;
        end;
        xptgArea3d     : begin
          Ok1 := CheckAdjustSheetIndex(PXLSPtgsArea3d(Ptgs).Sheet1,ASheetIndex,ACount,i);
          Ok2 := CheckAdjustSheetIndex(PXLSPtgsArea3d(Ptgs).Sheet2,ASheetIndex,ACount,j);

          if Ok1 then
            PXLSPtgsArea3d(Ptgs).Sheet1 := i
          else begin
            PXLSPtgsArea3d(Ptgs).Id := xptgArea3dErr;
            PXLSPtgsArea3dError(Ptgs).Error1 := Byte(errRef);
          end;
          if Ok2 then
            PXLSPtgsArea3d(Ptgs).Sheet2 := j
          else begin
            PXLSPtgsArea3d(Ptgs).Id := xptgArea3dErr;
            PXLSPtgsArea3dError(Ptgs).Error2 := Byte(errRef);
          end;

          if Ok1 xor Ok2 then begin
            PXLSPtgsArea3d(Ptgs).Id := xptgArea3d;
            if Ok1 then
              PXLSPtgsArea3d(Ptgs).Sheet2 := ASheetIndex - 1
            else
              PXLSPtgsArea3d(Ptgs).Sheet1 := ASheetIndex + -ACount - 1;
          end;
        end;
        xptgAreaErr3d97: begin
{$ifdef XLS_BIFF}
          FManager._ExtNames97.DeleteSheet(PXLSPtgsArea3d97(Ptgs).ExtSheetIndex);
{$endif}
        end;
      end;
    end;
  finally
    Iter.Free;
  end;
end;

procedure TXLSFormulaHandler.Calculate;
var
  i: integer;
  Item: PXLS3dCompactRef;
  Ptgs: PXLSPtgs;
  PtgsSz: integer;
  TargetArea: PXLSFormulaArea;
  Cells: TXLSCellMMU;
  C: TXLSCellItem;
  Res: TXLSFormulaValue;
  ProgressMod: integer;
begin
  FCCBuilder.Build(FCalcChain);

  FManager.BeginProgress(xptCalculateFormulas,FCalcChain.Count);

  ProgressMod := FCalcChain.Count div 25;
  if ProgressMod <= 0 then
    ProgressMod := 1;
  
  for i := 0 to FCalcChain.Count - 1 do begin
    if (i mod ProgressMod) = 0 then
      FManager.WorkProgress(i);

    Item := FCalcChain[i];
    Cells := FManager.Worksheets[Item.SheetId].Cells;
    C := Cells.FindCell(Item.Col,Item.Row);
    PtgsSz := Cells.GetFormulaPtgs(@C,Ptgs);
    TargetArea := Cells.GetFormulaTargetArea(@C);
    Res := FEvaluator.Evaluate(Item.SheetId,Item.Col,Item.Row,Ptgs,PtgsSz,TargetArea);
    if Res.RefType = xfrtArea then begin
    end
    else if not (Res.RefType in [xfrtNone,xfrtRef]) then
      raise XLSRWException.Create('Invalid formula result');
    case Res.ValType of
      xfvtUnknown: ;
      xfvtFloat   : FManager.Worksheets[Item.SheetId].Cells.AddFormulaVal(@C,Item.Col,Item.Row,Res.vFloat);
      xfvtBoolean : FManager.Worksheets[Item.SheetId].Cells.AddFormulaVal(@C,Item.Col,Item.Row,Res.vBool);
      xfvtError   : FManager.Worksheets[Item.SheetId].Cells.AddFormulaVal(@C,Item.Col,Item.Row,Res.vErr);
      xfvtString  : FManager.Worksheets[Item.SheetId].Cells.AddFormulaVal(@C,Item.Col,Item.Row,Res.vStr);
      xfvtAreaList: raise XLSRWException.Create('Invalid formula result');
    end;
  end;
  FManager.EndProgress;
end;

procedure TXLSFormulaHandler.CalculateAndVerify;
var
  i: integer;
  S: AxUCString;
  Ok: boolean;
  Item: PXLS3dCompactRef;
  Ptgs: PXLSPtgs;
  PtgsSz: integer;
  TargetArea: PXLSFormulaArea;
  Cells: TXLSCellMMU;
  Res: TXLSFormulaValue;
  C: TXLSCellItem;
  ProgressMod: integer;
begin
  FCCBuilder.Build(FCalcChain);

  FManager.BeginProgress(xptCalculateFormulas,FCalcChain.Count);
  ProgressMod := FCalcChain.Count div 25;
  if ProgressMod <= 0 then
    ProgressMod := 1;

  for i := 0 to FCalcChain.Count - 1 do begin
    if (i mod ProgressMod) = 0 then
      FManager.WorkProgress(i);

    Item := FCalcChain[i];
    Cells := FManager.Worksheets[Item.SheetId].Cells;
    C := Cells.FindCell(Item.Col,Item.Row);

    if C.Data = Nil then
      raise XLSRWException.Create('Cell is missing');

    // XLSRWException if C not is formula.
    Cells.FormulaType(@C);

    PtgsSz := Cells.GetFormulaPtgs(@C,Ptgs);
    TargetArea := Cells.GetFormulaTargetArea(@C);

//    if i >= 34336 then
//      Ok := False;

    Res := FEvaluator.Evaluate(Item.SheetId,Item.Col,Item.Row,Ptgs,PtgsSz,TargetArea);

//    if not (Res.RefType in [xfrtNone,xfrtRef]) then
//      raise XLSRWException.CreateFmt('Invalid formula result in cell %s',[ColRowToRefStr(Item.Col,Item.Row)]);
    Ok := True;
    case Res.ValType of
      // Usually caused by empty cells, If(TRUE;A1;100), where A1 is empty
      xfvtUnknown : Ok := True;
      xfvtFloat   : begin
        Ok := (Cells.CellType(@C) = xctFloatFormula) and (Abs(Cells.GetFormulaValFloat(@C) - Res.vFloat) < 0.0000001);
        FManager.Worksheets[Item.SheetId].Cells.AddFormulaVal(@C,Item.Col,Item.Row,Res.vFloat);
      end;
      xfvtBoolean : begin
       Ok := (Cells.CellType(@C) = xctBooleanFormula) and (Cells.GetFormulaValBoolean(@C) = Res.vBool);
       FManager.Worksheets[Item.SheetId].Cells.AddFormulaVal(@C,Item.Col,Item.Row,Res.vBool);
      end;
      xfvtError   : begin
        Ok := (Cells.CellType(@C) = xctErrorFormula) and (Cells.GetFormulaValError(@C) = Res.vErr);
        FManager.Worksheets[Item.SheetId].Cells.AddFormulaVal(@C,Item.Col,Item.Row,Res.vErr);
      end;
      xfvtString  : begin
        Ok := (Cells.CellType(@C) = xctStringFormula) and (Cells.GetFormulaValString(@C) = Res.vStr);
        FManager.Worksheets[Item.SheetId].Cells.AddFormulaVal(@C,Item.Col,Item.Row,Res.vStr);
      end;
      xfvtAreaList: raise XLSRWException.Create('Invalid formula result');
    end;
    if not Ok and not FEvaluator.Volatile then begin
      // C may change due to different result after calc.
      C := Cells.FindCell(Item.Col,Item.Row);
      if Cells.IsFormulaCompiled(@C) then begin
        PtgsSz := Cells.GetFormulaPtgs(@C,Ptgs);

        S := '#' + IntToStr(i) + ' ' + DecodeFormula(FManager.Worksheets[Item.SheetId].Cells,Item.SheetId,Ptgs,PtgsSz,False);
        FManager.Errors.Hint(FManager.Worksheets[Item.SheetId].Name + '!' + ColRowToRefStr(Item.Col,Item.Row) + ', ' + S ,XLSHINT_BAD_FORMULA_RESULT);
      end
      else
        FManager.Errors.Hint(FManager.Worksheets[Item.SheetId].Name + '!' + ColRowToRefStr(Item.Col,Item.Row) + ', [Error not compiled]' ,XLSHINT_BAD_FORMULA_RESULT);
//      FManager.Errors.Hint(FManager.Worksheets[Item.SheetId].Name + '!' + ColRowToRefStr(Item.Col,Item.Row),XLSHINT_BAD_FORMULA_RESULT);
    end;
//    Cells.CheckIntegrity(Nil);
  end;
  FManager.EndProgress;

  FManager.Errors.Message('Calculated ' + IntToStr(FCalcChain.Count) + ' formulas');
end;

function TXLSFormulaHandler.CheckAdjustSheetIndex(const ACurrIndex, ASheetIndex, ACount: integer; out ANewIndex: integer): boolean;
begin
  Result := True;
  ANewIndex := ACurrIndex;
  if ACount > 0 then begin
    if ACurrIndex >= ASheetIndex then
      Inc(ANewIndex,ACount);
  end
  else begin
    if ACurrIndex >= ASheetIndex then begin
      if ACurrIndex < (ASheetIndex + -ACount) then
        Result := False
      else
        Dec(ANewIndex,-ACount);
    end;
  end;
end;

function TXLSFormulaHandler.CollectNames(ANames: TIntegerList; const APtgs: PXLSPtgs; const APtgsSize: integer): boolean;
var
  Iter: TXLSIteratePtgs;
  Ptgs: PXLSPtgs;
begin
  Iter := TXLSIteratePtgs.Create(APtgs,APtgsSize);
  try
    while Iter.Next do begin
      Ptgs := Iter.CurrPtgs;
      if Ptgs.Id = xptgName then
        ANames.Add(PXLSPtgsName(Ptgs).NameId);
    end;
  finally
    Iter.Free;
  end;
  Result := ANames.Count > 0;
end;

constructor TXLSFormulaHandler.Create(AManager: TXc12Manager);
begin
  FManager   := AManager;
  FTokenizer := TXLSTokenizer.Create(FManager.Errors);
  FDecoder   := TXLSFormulaDecoder.Create(FManager);
  FCCBuilder := TXLSCalcChainBuilder.Create(FManager);
  FEvaluator := TXLSFormulaEvaluator.Create(FManager);
  FCalcChain := TXLSCalcChain.Create;
end;

function TXLSFormulaHandler.DecodeFormula(ACells: TXLSCellMMU; const ASheetIndex: integer; APtgs: PXLSPtgs; const APtgsSize: integer; const ASaveToFile: boolean): AxUCString;
var
  Sz: integer;
  P: PXLSPtgs;
  Ref: TXLSCellItem;
begin
  P := APtgs;
  Sz := APtgsSize;
  if (P.Id = xptgArrayFmlaChild97) then begin
    Ref := ACells.FindCell(PXLSPtgsArrayChildFmla(P).ParentCol,PXLSPtgsArrayChildFmla(P).ParentRow);
    Sz := ACells.GetFormulaPtgs(@Ref,P);
  end;

{$ifdef XLS_BIFF}
  if P.Id in [xptg_EXCEL_97,xptg_ARRAYCONSTS_97] then
    Result := DecodeFmla(P,Sz,ASheetIndex,0,0,Nil,FormatSettings.ListSeparator)
  else
{$endif}
    Result := FDecoder.Decode(ACells,ASheetIndex,P,Nil,Sz,ASaveToFile);
end;

destructor TXLSFormulaHandler.Destroy;
begin
  FTokenizer.Free;
  FDecoder.Free;
  FCCBuilder.Free;
  FEvaluator.Free;
  FCalcChain.Free;
end;

function TXLSFormulaHandler.EncodeFormula(const AFormula: AxUCString; out APtgs: PXLSPtgs; const ASheetId: integer): integer;
var
  Tokens: TXLSTokenList;
  Encoder: TXLSFormulaEncoder;
begin
  FTokenizer.Formula := AFormula;
  Tokens := FTokenizer.Parse;

  if Tokens <> Nil then begin
    Encoder := TXLSFormulaEncoder.Create(FManager);
    try
      Encoder.CurrTables := FCurrTables;
      Result := Encoder.Encode(Tokens,APtgs,ASheetId,-1,-1);
    finally
      Encoder.Free;
    end;
//    FManager.Errors.ErrorList.Text := FManager.Errors.ErrorList.Text + #13 + FEncoder.Dump;
    Tokens.Free;
  end
  else
    Result := 0;
end;

function TXLSFormulaHandler.EncodeFormula(const AFormula: AxUCString; out APtgs: PXLSPtgs; const ASheetId,ACol, ARow: integer; AStrictMode: boolean = True): integer;
var
  Tokens: TXLSTokenList;
  Encoder: TXLSFormulaEncoder;
begin
  FTokenizer.Formula := AFormula;
  FTokenizer.StrictMode := AStrictMode;
  Tokens := FTokenizer.Parse;

  if Tokens <> Nil then begin
    Encoder := TXLSFormulaEncoder.Create(FManager);
    try
      Encoder.CurrTables := FCurrTables;
      Result := Encoder.Encode(Tokens,APtgs,ASheetId,ACol,ARow);
    finally
      Encoder.Free;
    end;
    Tokens.Free;
  end
  else
    Result := 0;
end;

function TXLSFormulaHandler.EvaluateFormula(ASheetIndex, ACol, ARow: integer; APtgs: PXLSPtgs; APtgsSize: integer): TXLSFormulaValue;
begin
  Result := FEvaluator.Evaluate(ASheetIndex,ACol,ARow,APtgs,APtgsSize,Nil);
end;

function TXLSFormulaHandler.EvaluateFormulaStr(ACells: TXLSCellMMU; ASheetIndex,ACol,ARow: integer; APtgs: PXLSPtgs; APtgsSize: integer): AxUCString;
var
  Cell: TXLSCellItem;
  TargetArea: PXLSFormulaArea;
begin
  Cell := ACells.FindCell(ACol,ARow);
  TargetArea := ACells.GetFormulaTargetArea(@Cell);
  Result := FEvaluator.EvaluateStr(ASheetIndex,ACol,ARow,APtgs,APtgsSize,TargetArea);
end;

function TXLSFormulaHandler.GetAreaErrorId(AId: integer): integer;
begin
  case AId of
    xptgArea  : Result := xptgAreaErr;
    xptgArea1d: Result := xptgArea1dErr;
    else        Result := xptgAreaErr;
  end;
end;

function TXLSFormulaHandler.GetUserFuncEvent: TUserFunctionEvent;
begin
  Result := FEvaluator.OnUserFunction;
end;

procedure TXLSFormulaHandler.SetTables(const Value: TXc12Tables);
begin
  FCurrTables := Value;
end;

procedure TXLSFormulaHandler.SetUserFuncEvent(const Value: TUserFunctionEvent);
begin
  FEvaluator.OnUserFunction := Value;
end;

function TXLSFormulaHandler.EncodeFormula(const AFormula: AxUCString; out APtgs: PXLSPtgs; const ASheetId, ACol, ARow: integer; const ADebugItems: TFmlaDebugItems): integer;
var
  Tokens: TXLSTokenList;
  Encoder: TXLSFormulaEncoder;
begin
  FTokenizer.Formula := AFormula;
  Tokens := FTokenizer.Parse;

  if Tokens <> Nil then begin
    Encoder := TXLSFormulaEncoder.Create(FManager);
    try
      Encoder.CurrTables := FCurrTables;
      Result := Encoder.Encode(Tokens,APtgs,ASheetId,ACol,ARow,ADebugItems);
    finally
      Encoder.Free;
    end;
    Tokens.Free;
  end
  else
    Result := 0;
end;

end.
