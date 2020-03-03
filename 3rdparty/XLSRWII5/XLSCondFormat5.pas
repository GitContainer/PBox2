unit XLSCondFormat5;

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

uses Classes, SysUtils, Contnrs, Math,
     Xc12Utils5, Xc12DataStylesheet5, Xc12DataWorksheet5, Xc12Manager5,
     Xc12DataAutofilter5,
     XLSUtils5, XLSClassFactory5, XLSCellAreas5, XLSMoveCopy5, XLSFormula5,
     XLSFormulaTypes5, XLSEvaluate5, XLSCellMMU5;

type TXSSCondFmtResultType = (xcfrtNone,xcfrtDXF,xcfrtColor,xcfrtDataBar,xcfrtDataBarNoText,xcfrtIcon,xcfrtIconNoText);

type TXLSCFRule = class(TObject)
private
     function  GetAboveAverage: boolean;
     function  GetBottom: boolean;
     function  GetColorScale: TXc12ColorScale;
     function  GetDataBar: TXc12DataBar;
     function  GetDXF: TXc12DXF;
     function  GetEqualAverage: boolean;
     function  GetFormulas(Index: integer): AxUCString;
     function  GetIconSet: TXc12IconSet;
     function  GetOperator: TXc12ConditionalFormattingOperator;
     function  GetPercent: boolean;
     function  GetPriority: integer;
     function  GetRank: integer;
     function  GetStdDev: integer;
     function  GetStopIfTrue: boolean;
     function  GetText: AxUCString;
     function  GetTimePeriod: TXc12TimePeriod;
     function  GetType: TXc12CfType;
     procedure SerFormulas(Index: integer; const Value: AxUCString);
     procedure SetAboveAverage(const Value: boolean);
     procedure SetBottom(const Value: boolean);
     procedure SetEqualAverage(const Value: boolean);
     procedure SetOperator(const Value: TXc12ConditionalFormattingOperator);
     procedure SetPercent(const Value: boolean);
     procedure SetPriority(const Value: integer);
     procedure SetRank(const Value: integer);
     procedure SetStdDev(const Value: integer);
     procedure SetStopIfTrue(const Value: boolean);
     procedure SetText(const Value: AxUCString);
     procedure SetTimePeriod(const Value: TXc12TimePeriod);
     procedure SetType(const Value: TXc12CfType);
protected
     FManager: TXc12Manager;
     FData   : TXc12CFRule;
public
     constructor Create(AManager: TXc12Manager; AData: TXc12CFRule);

     property Formulas[Index: integer]: AxUCString read GetFormulas write SerFormulas;

     property RuleType: TXc12CfType read GetType write SetType;
     property Style: TXc12DXF read GetDXF;
     property Priority: integer read GetPriority write SetPriority;
     property StopIfTrue: boolean read GetStopIfTrue write SetStopIfTrue;
     property AboveAverage: boolean read GetAboveAverage write SetAboveAverage;
     property Percent: boolean read GetPercent write SetPercent;
     property Bottom: boolean read GetBottom write SetBottom;
     property Operator_: TXc12ConditionalFormattingOperator read GetOperator write SetOperator;
     property Text: AxUCString read GetText write SetText;
     property TimePeriod: TXc12TimePeriod read GetTimePeriod write SetTimePeriod;
     property Rank: integer read GetRank write SetRank;
     property StdDev: integer read GetStdDev write SetStdDev;
     property EqualAverage: boolean read GetEqualAverage write SetEqualAverage;

     property ColorScale: TXc12ColorScale read GetColorScale;
     property DataBar: TXc12DataBar read GetDataBar;
     property IconSet: TXc12IconSet read GetIconSet;
     end;

type TXLSConditionalFormat = class(TXc12ConditionalFormatting)
private
     function  GetAreas: TCellAreas;
     function  GetPivot: boolean;
     procedure SetPivot(const Value: boolean);
protected
     FManager     : TXc12Manager;
     FFormulas    : TXLSFormulaHandler;
     FXc12Sheet   : TXc12DataWorksheet;

     FCurrArea    : TCellArea;

     FResColor    : longword;
     FResDXF      : TXc12DXF;
     FResPercent  : double;
     FResIconIndex: integer;
     FResIconStyle: TXc12IconSetType;

     procedure PrepareAboveAverage(ARule: TXc12CfRule);
     procedure PrepareMinMax(ARule: TXc12CfRule);
     procedure PrepareTop10(ARule: TXc12CfRule);

     function  GetColorScaleCfvoVal(ARule: TXc12CfRule; ACfvo: TXc12Cfvo): double;
     function  GetIconSetCfvoVal(ARule: TXc12CfRule; ACfvo: TXc12Cfvo): double;
     function  CompileCfvoFormula(ACfvo: TXc12Cfvo): double;
     procedure CompileFormulas(ARule: TXc12CfRule);
     function  FormulaValAsFloat(ARule: TXc12CfRule; const AFormula: integer): double;
     function  DoExpression(ARule: TXc12CfRule; const ACol,ARow: integer): TXSSCondFmtResultType;
     function  DoCellIs(ARule: TXc12CfRule; const ACol,ARow: integer): TXSSCondFmtResultType;
     function  DoAboveAverage(ARule: TXc12CfRule; const ACol,ARow: integer): TXSSCondFmtResultType;
     function  DoColorScale(ARule: TXc12CfRule; const ACol,ARow: integer): TXSSCondFmtResultType;
     function  DoDuplicates(ARule: TXc12CfRule; const ACol,ARow: integer; const ADoUnique: boolean): TXSSCondFmtResultType;
     function  DoDataBar(ARule: TXc12CfRule; const ACol,ARow: integer): TXSSCondFmtResultType;
     function  DoIconSet(ARule: TXc12CfRule; const ACol,ARow: integer): TXSSCondFmtResultType;
     function  DoTop10(ARule: TXc12CfRule; const ACol,ARow: integer): TXSSCondFmtResultType;
public
     constructor Create(AManager : TXc12Manager; AFormulas: TXLSFormulaHandler);
     destructor Destroy; override;

     procedure SetStyle(ARule: TXc12CfRule; AFontStyle: TXc12FontStyles; AFill: TXc12Color);

     procedure BeginExecute(const ASqRefIndex: integer);
     function  Execute(const ACol,ARow: integer): TXSSCondFmtResultType;

     function  Intersect(Col1,Row1,Col2,Row2: integer): boolean; override;
     procedure Copy(Col1,Row1,Col2,Row2,DeltaCol,DeltaRow: integer); override;
     procedure Delete(Col1,Row1,Col2,Row2: integer); override;
     procedure Include(Col1,Row1,Col2,Row2: integer); override;
     procedure Move(DeltaCol,DeltaRow: integer); override;
     procedure Move(Col1,Row1,Col2,Row2,DeltaCol,DeltaRow: integer); override;

     property Pivot: boolean read GetPivot write SetPivot;
     property Areas: TCellAreas read GetAreas;

     property ResultDXF: TXc12DXF read FResDXF;
     property ResultColor: longword read FResColor;
     property ResPercent: double read FResPercent;
     property ResIconIndex: integer read FResIconIndex;
     property ResIconStyle: TXc12IconSetType read FResIconStyle;
     end;

type TXLSConditionalFormats = class(TXc12ConditionalFormattings)
private
     function GetItems(Index: integer): TXLSConditionalFormat;
protected
     function  CreateMember: TXc12ConditionalFormatting; override;
public
     constructor Create(AClassFactory: TXLSClassFactory);

     function AddCF: TXLSConditionalFormat;

     function  Find(ACol,ARow: integer): TXLSConditionalFormat;
     procedure FillHitList(const Col1,Row1,Col2,Row2: integer; AList: TBaseCellAreas);

     property Items[Index: integer]: TXLSConditionalFormat read GetItems;
     end;

implementation

{ TXLSConditionalFormats }

function TXLSConditionalFormats.AddCF: TXLSConditionalFormat;
begin
  Result := TXLSConditionalFormat(inherited AddCF);
end;

constructor TXLSConditionalFormats.Create(AClassFactory: TXLSClassFactory);
begin
  inherited Create(AClassFactory);
end;

function TXLSConditionalFormats.CreateMember: TXc12ConditionalFormatting;
begin
  Result := TXc12ConditionalFormatting(FClassFactory.CreateAClass(xcftConditionalFormat));
  TXLSConditionalFormat(Result).FXc12Sheet := TXc12DataWorksheet(FXc12Sheet);
end;

procedure TXLSConditionalFormats.FillHitList(const Col1, Row1, Col2, Row2: integer; AList: TBaseCellAreas);
var
  i: integer;
begin
  for i := 0 to Count - 1 do
    Items[i].Areas.FillAndClipHitList(Col1,Row1,Col2,Row2,AList,Items[i]);
end;

function TXLSConditionalFormats.Find(ACol, ARow: integer): TXLSConditionalFormat;
var
  i,j: integer;
begin
  for i := 0 to Count - 1 do begin
    j := Items[i].Areas.CellInAreas(ACol,ARow);
    if j >= 0 then begin
      Result := Items[i];
//      Result.FCurrArea := Items[i].Areas[j];
      Exit;
    end;
  end;
  Result := Nil;
end;

function TXLSConditionalFormats.GetItems(Index: integer): TXLSConditionalFormat;
begin
  Result := TXLSConditionalFormat(inherited Items[Index]);
end;

{ TXLSConditionalFormat }

procedure TXLSConditionalFormat.BeginExecute(const ASqRefIndex: integer);
var
  i: integer;
begin
  FCurrArea := FSqRef[ASqRefIndex];
  for i := 0 to FCfRules.Count - 1 do begin
    CompileFormulas(FCfRules[i]);
    case FCfRules[i].Type_ of
      x12ctColorScale       : PrepareMinMax(FCfRules[i]);
      x12ctDataBar          : PrepareMinMax(FCfRules[i]);
      x12ctIconSet          : PrepareMinMax(FCfRules[i]);
      x12ctTop10            : PrepareTop10(FCfRules[i]);
      x12ctUniqueValues     : ;
      x12ctDuplicateValues  : ;

      x12ctAboveAverage     : PrepareAboveAverage(FCfRules[i]);
    end;
  end;
end;

function TXLSConditionalFormat.CompileCfvoFormula(ACfvo: TXc12Cfvo): double;
var
  FV: TXLSFormulaValue;
  Ptgs: PXLSPtgs;
begin
  if ACfvo.Ptgs = Nil then begin
    Ptgs := ACfvo.Ptgs;
    ACfvo.PtgsSz := FFormulas.EncodeFormula(ACfvo.Val,Ptgs,FXc12Sheet.Index);
    ACfvo.Ptgs := Ptgs;
  end;
  FV := FFormulas.EvaluateFormula(FXc12Sheet.Index,0,0,ACfvo.Ptgs,ACfvo.PtgsSz);
  if FV.ValType = xfvtFloat then
    Result := FV.vFloat
  else
    Result := 0;
end;

procedure TXLSConditionalFormat.CompileFormulas(ARule: TXc12CfRule);
var
  i: integer;
  Ptgs: PXLSPtgs;
begin
  for i := 0 to ARule.FormulaMaxCount - 1 do begin
    if (ARule.Formulas[i] <> '') and (ARule.Ptgs[i] = Nil) then begin
      Ptgs := ARule.Ptgs[i];
      ARule.PtgsSz[i] := FFormulas.EncodeFormula(ARule.Formulas[i],Ptgs,FXc12Sheet.Index);
      ARule.Ptgs[i] := Ptgs;
    end;
  end;
end;

procedure TXLSConditionalFormat.Copy(Col1, Row1, Col2, Row2, DeltaCol, DeltaRow: integer);
begin
  FSQRef.Copy(Col1, Row1, Col2, Row2,DeltaCol,DeltaRow);
end;

constructor TXLSConditionalFormat.Create(AManager : TXc12Manager; AFormulas: TXLSFormulaHandler);
begin
  inherited Create;

  FManager := AManager;
  FFormulas := AFormulas;

//  FRules := TXLSCFRules.Create(FData.CfRules);
end;

procedure TXLSConditionalFormat.Delete(Col1, Row1, Col2, Row2: integer);
begin
  FSQRef.Delete(Col1, Row1, Col2, Row2);
end;

destructor TXLSConditionalFormat.Destroy;
begin
  inherited;
end;

// Stdev is not implemented.
function TXLSConditionalFormat.DoAboveAverage(ARule: TXc12CfRule; const ACol,ARow: integer): TXSSCondFmtResultType;
var
  V: double;
  Ok: boolean;
begin
  V := FXc12Sheet.Cells.GetFloat(ACol,ARow);
  if ARule.AboveAverage then begin
    if ARule.EqualAverage then
      Ok := V >= ARule.ValAverage
    else
      Ok := V > ARule.ValAverage
  end
  else begin
    if ARule.EqualAverage then
      Ok := V <= ARule.ValAverage
    else
      Ok := V < ARule.ValAverage;
  end;

  if Ok then
    Result := xcfrtDXF
  else
    Result := xcfrtNone;
end;

function TXLSConditionalFormat.DoCellIs(ARule: TXc12CfRule; const ACol,ARow: integer): TXSSCondFmtResultType;
var
  Val: double;
  Cmp1,Cmp2: double;
  Ok: boolean;
begin
  Val := FXc12Sheet.Cells.GetFloat(ACol,ARow);
  Cmp1 := FormulaValAsFloat(ARule,0);
  Cmp2 := FormulaValAsFloat(ARule,1);
  case ARule.Operator_ of
    x12cfoLessThan          : Ok := Val < Cmp1;
    x12cfoLessThanOrEqual   : Ok := Val <= Cmp1;
    x12cfoEqual             : Ok := Val = Cmp1;
    x12cfoNotEqual          : Ok := Val <> Cmp1;
    x12cfoGreaterThanOrEqual: Ok := Val >= Cmp1;
    x12cfoGreaterThan       : Ok := Val > Cmp1;
    x12cfoBetween           : Ok := (Val >= Cmp1) and (Val <= Cmp2);
    x12cfoNotBetween        : Ok := not ((Val >= Cmp1) and (Val <= Cmp2));
    else                      Ok := False;
  end;
  if Ok then
    Result := xcfrtDXF
  else
    Result := xcfrtNone;
end;

function TXLSConditionalFormat.DoColorScale(ARule: TXc12CfRule; const ACol, ARow: integer): TXSSCondFmtResultType;
var
  Val: double;
  V1,V2,V3: double;
  ValRange: double;
begin
  Result := xcfrtColor;

  Val := FXc12Sheet.Cells.GetFloat(ACol,ARow);

  case ARule.ColorScale.Colors.Count of
    2: begin
      V1 := GetColorScaleCfvoVal(ARule,ARule.ColorScale.Cfvos[0]);
      V2 := GetColorScaleCfvoVal(ARule,ARule.ColorScale.Cfvos[1]);

      if Val < V1 then
        FResColor := ARule.ColorScale.Colors[0].ARGB
      else if Val > V2 then
        FResColor := ARule.ColorScale.Colors[1].ARGB
      else begin
        if (V2 - V1) = 0 then
          ValRange := 0
        else
          ValRange := (Val - V1) / (V2 - V1);

        FResColor := MakeGradientColor(ARule.ColorScale.Colors[0].ARGB,ARule.ColorScale.Colors[1].ARGB,ValRange);
      end;
    end;
    3: begin
      V1 := GetColorScaleCfvoVal(ARule,ARule.ColorScale.Cfvos[0]);
      V2 := GetColorScaleCfvoVal(ARule,ARule.ColorScale.Cfvos[1]);
      V3 := GetColorScaleCfvoVal(ARule,ARule.ColorScale.Cfvos[2]);

      if Val < V1 then
        FResColor := ARule.ColorScale.Colors[0].ARGB
      else if Val > V3 then
        FResColor := ARule.ColorScale.Colors[2].ARGB
      else begin
        if Val <= V2 then begin
          if (V2 - V1) = 0 then
            ValRange := 0
          else
            ValRange := (Val - V1) / (V2 - V1);
          FResColor := MakeGradientColor(ARule.ColorScale.Colors[0].ARGB,ARule.ColorScale.Colors[1].ARGB,ValRange );
        end
        else begin
          if (V2 - V1) = 0 then
            ValRange := 0
          else
            ValRange := (Val - V2) / (V3 - V2);
          FResColor := MakeGradientColor(ARule.ColorScale.Colors[1].ARGB,ARule.ColorScale.Colors[2].ARGB,ValRange);
        end;
      end;
    end;
  end;
end;

function TXLSConditionalFormat.DoDataBar(ARule: TXc12CfRule; const ACol, ARow: integer): TXSSCondFmtResultType;
var
  Val: double;
  V1,V2: double;
  Cell: TXLSCellItem;
begin
  if not FXc12Sheet.Cells.FindCell(ACol,ARow,Cell) then begin
    Result := xcfrtNone;
    Exit;
  end;

  if ARule.DataBar.ShowValue then
    Result := xcfrtDataBar
  else
    Result := xcfrtDataBarNoText;

  V1 := GetIconSetCfvoVal(ARule,ARule.DataBar.Cfvo1);
  V2 := GetIconSetCfvoVal(ARule,ARule.DataBar.Cfvo2);

  Val := FXc12Sheet.Cells.GetFloat(@Cell);

  if Val <= V1 then
    FResPercent := ARule.ValMin / ARule.ValMax
  else if Val >= V2 then
    FResPercent := 1
  else
    FResPercent := Val / ARule.ValMax;
  FResColor := ARule.DataBar.Color.ARGB;
end;

function TXLSConditionalFormat.DoDuplicates(ARule: TXc12CfRule; const ACol, ARow: integer; const ADoUnique: boolean): TXSSCondFmtResultType;
var
  C,R: integer;
  Cell1,Cell2: TXLSCellItem;
  CT1,CT2: TXLSCellType;
  Ok: boolean;
begin
  if FXc12Sheet.Cells.FindCell(ACol,ARow,Cell1) and (FXc12Sheet.Cells.CellType(@Cell1) <> xctBlank)then begin
    CT1 := FXc12Sheet.Cells.CellType(@Cell1);
    case CT1 of
      xctFloatFormula  : CT1 := xctFloat;
      xctStringFormula : CT1 := xctString;
      xctBooleanFormula: CT1 := xctBoolean;
      xctErrorFormula  : CT1 := xctError;
    end;
    for R := FCurrArea.Row1 to FCurrArea.Row2 do begin
      for C := FCurrArea.Col1 to FCurrArea.Col2 do begin
        if not ((C = ACol) and (R = ARow)) and FXc12Sheet.Cells.FindCell(C,R,Cell2) and (FXc12Sheet.Cells.CellType(@Cell2) <> xctBlank) then begin
          CT2 := FXc12Sheet.Cells.CellType(@Cell2);
          case CT2 of
            xctFloatFormula  : CT2 := xctFloat;
            xctStringFormula : CT2 := xctString;
            xctBooleanFormula: CT2 := xctBoolean;
            xctErrorFormula  : CT2 := xctError;
          end;
          Ok := False;
          if CT1 = CT2 then begin
            case CT1 of
              xctBoolean: Ok := FXc12Sheet.Cells.GetBoolean(@Cell1) = FXc12Sheet.Cells.GetBoolean(@Cell2);
              xctError  : Ok := FXc12Sheet.Cells.GetError(@Cell1) = FXc12Sheet.Cells.GetError(@Cell2);
              xctString : Ok := SameText(FXc12Sheet.Cells.GetString(@Cell1),FXc12Sheet.Cells.GetString(@Cell2));
              xctFloat  : Ok := FXc12Sheet.Cells.GetFloat(@Cell1) = FXc12Sheet.Cells.GetFloat(@Cell2);
            end;
            if Ok then begin
              if ADoUnique then
                Result := xcfrtNone
              else
                Result := xcfrtDXF;
              Exit;
            end;
          end;
        end;
      end;
    end;
  end;
  if ADoUnique then
    Result := xcfrtDXF
  else
    Result := xcfrtNone;
end;

function TXLSConditionalFormat.DoExpression(ARule: TXc12CfRule; const ACol,ARow: integer): TXSSCondFmtResultType;
var
  Ptgs: PXLSPtgs;
  FV: TXLSFormulaValue;
begin
  GetMem(Ptgs,ARule.PtgsSz[0]);
  try
    System.Move(ARule.Ptgs[0]^,Ptgs^,ARule.PtgsSz[0]);
    FFormulas.AdjustCell(Ptgs,ARule.PtgsSz[0],ACol - FCurrArea.Col1,ARow - FCurrArea.Row1);
    FV := FFormulas.EvaluateFormula(FXc12Sheet.Index,ACol,ARow,Ptgs,ARule.PtgsSz[0]);
    if (FV.ValType = xfvtBoolean) and FV.vBool then
      Result := xcfrtDXF
    else
      Result := xcfrtNone;
  finally
    FreeMem(Ptgs);
  end;
end;

function TXLSConditionalFormat.DoIconSet(ARule: TXc12CfRule; const ACol, ARow: integer): TXSSCondFmtResultType;
var
  i: integer;
  Val,Val1,Val2: double;
  V1,V2: double;
  Cell: TXLSCellItem;
begin
  if ARule.IconSet.ShowValue then
    Result := xcfrtIcon
  else
    Result := xcfrtIconNoText;

  FResIconStyle := ARule.IconSet.IconSet;

  if not FXc12Sheet.Cells.FindCell(ACol,ARow,Cell) or (ARule.IconSet.Cfvos.Count < 3) then begin
    Result := xcfrtNone;
    Exit;
  end;

  V1 := 0;

  Val := FXc12Sheet.Cells.GetFloat(@Cell);
  if ARule.IconSet.Cfvos[1].Type_ in [x12ctPercent,x12ctPercentile] then
    Val1 := Val - ARule.ValMin
  else
    Val1 := Val;
  if ARule.IconSet.Cfvos[ARule.IconSet.Cfvos.Count - 1].Type_ in [x12ctPercent,x12ctPercentile] then
    Val2 := Val - ARule.ValMin
  else
    Val2 := Val;

  if Val1 < GetIconSetCfvoVal(ARule,ARule.IconSet.Cfvos[1]) then
    FResIconIndex := ARule.IconSet.Cfvos.Count - 1
  else if Val2 >= GetIconSetCfvoVal(ARule,ARule.IconSet.Cfvos[ARule.IconSet.Cfvos.Count - 1]) then
    FResIconIndex := 0
  else begin
    for i := 1 to ARule.IconSet.Cfvos.Count - 1 do begin
      V2 := GetIconSetCfvoVal(ARule,ARule.IconSet.Cfvos[i]);

      if ARule.IconSet.Cfvos[i].Type_ in [x12ctPercent,x12ctPercentile] then
        Val2 := Val - ARule.ValMin
      else
        Val2 := Val;

      if (Val1 >= V1) and (Val2 < V2) then begin
        FResIconIndex := ARule.IconSet.Cfvos.Count - i;
        Break;
      end;

      V1 := V2;
    end;
  end;
  if ARule.IconSet.Reverse then
    FResIconIndex := ARule.IconSet.Cfvos.Count - 1 - FResIconIndex;
end;

function TXLSConditionalFormat.DoTop10(ARule: TXc12CfRule; const ACol, ARow: integer): TXSSCondFmtResultType;
var
  i: integer;
  Val: double;
begin
  Result := xcfrtNone;
  if not ARule.Percent and (ARule.Rank >= ARule.ValuesCount) then
    Exit;

  Val := FXc12Sheet.Cells.GetFloat(ACol,ARow);
  if ARule.Bottom then begin
    if ARule.Percent then
      i := Floor(ARule.ValuesCount * (ARule.Rank / 100))
    else
      i := ARule.Rank;
    if Val < ARule.Values[i] then
      Result := xcfrtDXF;
  end
  else begin
    if ARule.Percent then
      i := Ceil(ARule.ValuesCount * (1 - (ARule.Rank / 100)))
    else
      i := ARule.ValuesCount - ARule.Rank;
    if Val >= ARule.Values[i] then
      Result := xcfrtDXF;
  end;
end;

function TXLSConditionalFormat.Execute(const ACol,ARow: integer): TXSSCondFmtResultType;
var
  i: integer;
begin
  Result := xcfrtNone;
  for i := 0 to FCfRules.Count - 1 do begin
    case FCfRules[i].Type_ of
      x12ctExpression       : Result := DoExpression(FCfRules[i],ACol,ARow);
      x12ctCellIs           : Result := DoCellIs(FCfRules[i],ACol,ARow);
      x12ctColorScale       : Result := DoColorScale(FCfRules[i],ACol,ARow);
      x12ctDataBar          : Result := DoDataBar(FCfRules[i],ACol,ARow);
      x12ctIconSet          : Result := DoIconSet(FCfRules[i],ACol,ARow);
      x12ctTop10            : Result := DoTop10(FCfRules[i],ACol,ARow);
      x12ctUniqueValues     : Result := DoDuplicates(FCfRules[i],ACol,ARow,True);
      x12ctDuplicateValues  : Result := DoDuplicates(FCfRules[i],ACol,ARow,False);

      x12ctContainsText,
      x12ctNotContainsText,
      x12ctBeginsWith,
      x12ctEndsWith,
      x12ctContainsBlanks,
      x12ctNotContainsBlanks,
      x12ctContainsErrors,
      x12ctNotContainsErrors,
      x12ctTimePeriod       : Result := DoExpression(FCfRules[i],ACol,ARow);

      x12ctAboveAverage     : Result := DoAboveAverage(FCfRules[i],ACol,ARow);
    end;
    if Result <> xcfrtNone then begin
      if Result = xcfrtDXF then
        FResDXF := FCfRules[i].DXF;
      Exit;
    end;
  end;
end;

function TXLSConditionalFormat.FormulaValAsFloat(ARule: TXc12CfRule; const AFormula: integer): double;
var
  FV: TXLSFormulaValue;
begin
  if ARule.Ptgs[AFormula] <> Nil then begin
    FV := FFormulas.EvaluateFormula(FXc12Sheet.Index,0,0,ARule.Ptgs[AFormula],ARule.PtgsSz[AFormula]);
    if FV.ValType = xfvtFloat then
      Result := FV.vFloat
    else
      Result := 0;
  end
  else
    Result := 0;
end;

function TXLSConditionalFormat.GetAreas: TCellAreas;
begin
  Result := FSQRef;
end;

function TXLSConditionalFormat.GetColorScaleCfvoVal(ARule: TXc12CfRule; ACfvo: TXc12Cfvo): double;
var
  n: integer;
  R,C: integer;
begin
  if not (ACfvo.Type_ in [x12ctMax,x12ctMin]) then
    Result := CompileCfvoFormula(ACfvo)
  else
    Result := 0;

  case ACfvo.Type_ of
    x12ctNum       : ;
    x12ctPercentile,
    x12ctPercent   : begin
      Result := Result / 100;
      n := Round(FCurrArea.CellCount * Result);

      if n < 0 then
        Result := ARule.ValMin
      else if n >= FCurrArea.CellCount then
        Result := ARule.ValMax
      else begin
        R := FCurrArea.Row1 + (n div FCurrArea.Width);
        C := FCurrArea.Col1 + (n mod FCurrArea.Width);
        Result := FXc12Sheet.Cells.GetFloat(C,R);
      end;
    end;
    x12ctMax       : Result := ARule.ValMax;
    x12ctMin       : Result := ARule.ValMin;
    x12ctFormula   : ;
  end;
end;

function TXLSConditionalFormat.GetIconSetCfvoVal(ARule: TXc12CfRule; ACfvo: TXc12Cfvo): double;
begin
  Result := CompileCfvoFormula(ACfvo);

  case ACfvo.Type_ of
    x12ctNum       : ;
    x12ctMin       : Result := ARule.ValMin;
    x12ctMax       : Result := ARule.ValMax;
    x12ctPercent   : Result := (ARule.ValMax - ARule.ValMin) * (Result / 100);
    x12ctFormula   : ;
    x12ctPercentile: Result := (ARule.ValMax - ARule.ValMin) * (Result / 100);
  end;
end;

function TXLSConditionalFormat.GetPivot: boolean;
begin
  Result := FPivot;
end;

procedure TXLSConditionalFormat.Include(Col1, Row1, Col2, Row2: integer);
begin
  FSQRef.Include(Col1, Row1, Col2, Row2);
end;

function TXLSConditionalFormat.Intersect(Col1, Row1, Col2, Row2: integer): boolean;
begin
  Result := FSQRef.AreaInAreas(Col1, Row1, Col2, Row2);
end;

procedure TXLSConditionalFormat.Move(Col1, Row1, Col2, Row2, DeltaCol, DeltaRow: integer);
begin
  FSQRef.Move(Col1, Row1, Col2, Row2, DeltaCol, DeltaRow);
end;

procedure TXLSConditionalFormat.PrepareAboveAverage(ARule: TXc12CfRule);
var
  V: double;
  C,R: integer;
begin
  V := 0;
  for R := FCurrArea.Row1 to FCurrArea.Row2 do begin
    for C := FCurrArea.Col1 to FCurrArea.Col2 do
      V := V + FXc12Sheet.Cells.GetFloat(C,R);
  end;
  ARule.ValAverage := V / ((FCurrArea.Row2 - FCurrArea.Row1 + 1) * (FCurrArea.Col2 - FCurrArea.Col1 + 1));
end;

procedure TXLSConditionalFormat.PrepareMinMax(ARule: TXc12CfRule);
var
  V: double;
  VMin,VMax: double;
  C,R: integer;
begin
  VMin := MAXDOUBLE;
  VMax := MINDOUBLE;
  for R := FCurrArea.Row1 to FCurrArea.Row2 do begin
    for C := FCurrArea.Col1 to FCurrArea.Col2 do begin
      V := FXc12Sheet.Cells.GetFloat(C,R);
      VMin := Min(VMin,V);
      VMax := Max(VMax,V);
    end;
  end;
  ARule.ValMin := VMin;
  ARule.ValMax := VMax;
end;

procedure TXLSConditionalFormat.PrepareTop10(ARule: TXc12CfRule);
var
  i: integer;
  C,R: integer;
begin
  i := 0;
  ARule.ValuesCount := (FCurrArea.Row2 - FCurrArea.Row1 + 1) * (FCurrArea.Col2 - FCurrArea.Col1 + 1);
  for R := FCurrArea.Row1 to FCurrArea.Row2 do begin
    for C := FCurrArea.Col1 to FCurrArea.Col2 do begin
      ARule.Values[i] := FXc12Sheet.Cells.GetFloat(C,R);
      Inc(i);
    end;
  end;
  ARule.SortValues;
end;

procedure TXLSConditionalFormat.Move(DeltaCol, DeltaRow: integer);
begin
  FSQRef.Move(DeltaCol, DeltaRow);
end;

procedure TXLSConditionalFormat.SetPivot(const Value: boolean);
begin
  FPivot := Value;
end;

procedure TXLSConditionalFormat.SetStyle(ARule: TXc12CfRule; AFontStyle: TXc12FontStyles; AFill: TXc12Color);
var
  DXF: TXc12DXF;
  Font: TXc12Font;
  Fill: TXc12Fill;
begin
//  if FManager.StyleSheet.DXFs.Count <= 0 then
//    FManager.StyleSheet.DXFs.Add;

  DXF := FManager.StyleSheet.DXFs.Add;

  if AFontStyle <> [] then begin
    Font := DXF.AddFont;
    Font.Style := AFontStyle;
  end;

  if AFill.ColorType <> exctUnassigned then begin
    Fill := DXF.AddFill;
    Fill.BgColor := AFill;
  end;

  ARule.DXF := DXF;
end;

{ TXLSCFRule }

constructor TXLSCFRule.Create(AManager: TXc12Manager; AData: TXc12CFRule);
begin
  FManager := AManager;
  FData := AData;
end;

function TXLSCFRule.GetAboveAverage: boolean;
begin
  Result := FData.AboveAverage;
end;

function TXLSCFRule.GetBottom: boolean;
begin
  Result := FData.Bottom;
end;

function TXLSCFRule.GetColorScale: TXc12ColorScale;
begin
  Result := FData.ColorScale;
end;

function TXLSCFRule.GetDataBar: TXc12DataBar;
begin
  Result := FData.DataBar;
end;

function TXLSCFRule.GetDXF: TXc12DXF;
begin
  Result := FData.DXF;
end;

function TXLSCFRule.GetEqualAverage: boolean;
begin
  Result := FData.EqualAverage;
end;

function TXLSCFRule.GetFormulas(Index: integer): AxUCString;
begin
  Result := FData.Formulas[Index];
end;

function TXLSCFRule.GetIconSet: TXc12IconSet;
begin
  Result := FData.IconSet;
end;

function TXLSCFRule.GetOperator: TXc12ConditionalFormattingOperator;
begin
  Result := FData.Operator_;
end;

function TXLSCFRule.GetPercent: boolean;
begin
  Result := FData.Percent;
end;

function TXLSCFRule.GetPriority: integer;
begin
  Result := FData.Priority;
end;

function TXLSCFRule.GetRank: integer;
begin
  Result := FData.Rank;
end;

function TXLSCFRule.GetStdDev: integer;
begin
  Result := FData.StdDev;
end;

function TXLSCFRule.GetStopIfTrue: boolean;
begin
  Result := FData.StopIfTrue;
end;

function TXLSCFRule.GetText: AxUCString;
begin
  Result := FData.Text;
end;

function TXLSCFRule.GetTimePeriod: TXc12TimePeriod;
begin
  Result := FData.TimePeriod;
end;

function TXLSCFRule.GetType: TXc12CfType;
begin
  Result := FData.Type_;
end;

procedure TXLSCFRule.SerFormulas(Index: integer; const Value: AxUCString);
begin
  FData.Formulas[Index] := Value;
end;

procedure TXLSCFRule.SetAboveAverage(const Value: boolean);
begin
  FData.AboveAverage := Value;
end;

procedure TXLSCFRule.SetBottom(const Value: boolean);
begin
  FData.Bottom := Value;
end;

procedure TXLSCFRule.SetEqualAverage(const Value: boolean);
begin
  FData.EqualAverage := Value;
end;

procedure TXLSCFRule.SetOperator(const Value: TXc12ConditionalFormattingOperator);
begin
  FData.Operator_ := Value;
end;

procedure TXLSCFRule.SetPercent(const Value: boolean);
begin
  FData.Percent := Value;
end;

procedure TXLSCFRule.SetPriority(const Value: integer);
begin
  FData.Priority := Value;
end;

procedure TXLSCFRule.SetRank(const Value: integer);
begin
  FData.Rank := Value;
end;

procedure TXLSCFRule.SetStdDev(const Value: integer);
begin
  FData.StdDev:= Value;
end;

procedure TXLSCFRule.SetStopIfTrue(const Value: boolean);
begin
  FData.StopIfTrue := Value;
end;

procedure TXLSCFRule.SetText(const Value: AxUCString);
begin
  FData.Text := Value;
end;

procedure TXLSCFRule.SetTimePeriod(const Value: TXc12TimePeriod);
begin
  FData.TimePeriod := Value;
end;

procedure TXLSCFRule.SetType(const Value: TXc12CfType);
begin
  FData.Type_ := Value;
end;

end.
