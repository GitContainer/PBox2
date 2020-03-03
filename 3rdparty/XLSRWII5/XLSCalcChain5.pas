unit XLSCalcChain5;

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

uses Classes, SysUtils, Contnrs, Math, IniFiles,
{$ifdef XLS_BIFF}
     BIFF_Utils5,
{$endif}
     Xc12Utils5, Xc12Manager5, Xc12DataWorkbook5,
     XLSUtils5, XLSFormulaTypes5, XLSCellMMU5, XLSTSort;

type PCCFormulaData = ^TCCFormulaData;
     TCCFormulaData = record
     Ref: TXLS3dCompactRef;
     Ptgs: PXLSPtgs;
     PtgsSize: integer;
     end;

type TCCFormulaDataArray = array[0..10] of TCCFormulaData;
     PCCFormulaDataArray = ^TCCFormulaDataArray;

type TCircPtgsItem = class(TObject)
private
     function GetItems(Index: integer): PXLSPtgs;
protected
     FPtgs: array of PXLSPtgs;
public
     procedure Add(APtgs: PXLSPtgs); overload;
     procedure Add(AItem: TCircPtgsItem); overload;
     function Count: integer;

     property Items[Index: integer]: PXLSPtgs read GetItems; default;
     end;

type TCircPtgsStack = class(TObjectList)
private
     function  GetItems(Index: integer): TCircPtgsItem;
     procedure SetItems(Index: integer; const Value: TCircPtgsItem);
protected
     FFoundIgnoreCircularEvent: TNotifyEvent;
public
     constructor Create;

     procedure PushIgnore;
     procedure OpIgnore;
     procedure FuncIgnore(const AArgCount: integer);
     procedure ArrayIgnore(const ACols,ARows: integer);

     procedure Push(APtgs: PXLSPtgs);
     procedure Op;
     procedure Func(const AFuncId,AArgCount: integer);

     property OnFoundIgnoreCircular: TNotifyEvent read FFoundIgnoreCircularEvent write FFoundIgnoreCircularEvent;
     property Items[Index: integer]: TCircPtgsItem read GetItems write SetItems; default;
     end;

type TXLSCalcChain = class(TObject)
private
     function GetItems(Index: integer): PXLS3dCompactRef;
protected
     FItems: PDepRefArray;
     FCount: integer;
     FAllocSize: integer;
public
     destructor Destroy; override;

     procedure Clear;
     function  Count: integer;

     procedure Add(ARef: PXLS3dCompactRef);
     procedure Append(ACalcChain: TXLSCalcChain);
     procedure Done;

     property Itens[Index: integer]: PXLS3dCompactRef read GetItems; default;
     end;

type TXLSCalcChainBuilder = class(TObject)
protected
     FManager    : TXc12Manager;
     FAllocSize  : integer;
     FFormulaCnt : integer;
     FFormulas   : PCCFormulaDataArray;
     FPtgsList   : array of PXLSPtgs;
     FPtgsCount  : integer;
     FDepList    : TDepRefsItems;
     FCalcChain  : TXLSCalcChain;
     FHasIgnoreCircularFunc: boolean;

     function  FindFormula(ASheetId: integer; ACol,ARow: integer): integer;
     function  FindFormulaClosest(ASheetId: integer; ACol,ARow: integer): integer;

     procedure AddDependencies(ADep: TDepRefsItem; const ASheetId, ACol, ARow: integer); overload;
     procedure AddDependencies(ADep: TDepRefsItem; const ASheetId, ACol1, ARow1, ACol2, ARow2: integer); overload;

     procedure AddDependencies(ADep: TDepRefsItem; const ASheetId: integer; APtgsRef: PXLSPtgsRef); overload;
     procedure AddDependencies(ADep: TDepRefsItem; const ASheetId: integer; APtgsArea: PXLSPtgsArea); overload;
     procedure AddDependencies(ADep: TDepRefsItem; APtgsRef3d: PXLSPtgsRef3d); overload;
     procedure AddDependencies(ADep: TDepRefsItem; APtgsArea3d: PXLSPtgsArea3d); overload;
{$ifdef XLS_BIFF}
     procedure AddDependencies(ADep: TDepRefsItem; const ASheetId: integer; APtgsRef: PXLSPtgsRef97); overload;
     procedure AddDependencies(ADep: TDepRefsItem; const ASheetId: integer; APtgsArea: PXLSPtgsArea97); overload;
     procedure AddDependencies(ADep: TDepRefsItem; APtgsRef3d: PXLSPtgsRef3d97); overload;
     procedure AddDependencies(ADep: TDepRefsItem; APtgsArea3d: PXLSPtgsArea3d97); overload;
{$endif}
     procedure IgnoreCircular(ASender: TObject);
     function  CheckFormulaContent(AData: PCCFormulaData): TDepRefsItem;
     procedure CollectFormulaPtgs(APtgs: PXLSPtgs; const APtgsSz: integer);
     procedure RemoveRefsWhereCircularIsOk;
     procedure CheckFormulaRefs(ADep: TDepRefsItem; ADepRef: PXLS3dCompactRef);
     procedure CollectFormulas(ASheetId: integer; ACells: TXLSCellMMU);
public
     constructor Create(AManager: TXc12Manager);
     destructor Destroy; override;

     function DepListAsText: AxUCString;

     procedure Clear;

     function Build(ACalcChain: TXLSCalcChain): boolean;
     end;

implementation

{ TXLSCalcChainBuilder }

procedure TXLSCalcChainBuilder.AddDependencies(ADep: TDepRefsItem; const ASheetId: integer; APtgsRef: PXLSPtgsRef);
var
  i: integer;
begin
  i := FindFormula(ASheetId,APtgsRef.Col and not COL_ABSFLAG,APtgsRef.Row and not ROW_ABSFLAG);
  if i >= 0 then
    ADep.AddRef(@FFormulas[i].Ref);
end;

procedure TXLSCalcChainBuilder.AddDependencies(ADep: TDepRefsItem; const ASheetId: integer; APtgsArea: PXLSPtgsArea);
var
  i,i1,i2: integer;
begin
  i1 := FindFormulaClosest(ASheetId,0,APtgsArea.Row1 and not ROW_ABSFLAG);
  i2 := FindFormulaClosest(ASheetId,XLS_MAXCOL,APtgsArea.Row2 and not ROW_ABSFLAG);

  for i := i1 to i2 do begin
    if FFormulas[i].Ref.Row > (APtgsArea.Row2 and not COL_ABSFLAG) then
      Exit;
    if (FFormulas[i].Ref.Col >= (APtgsArea.Col1 and not COL_ABSFLAG)) and (FFormulas[i].Ref.Col <= (APtgsArea.Col2 and not COL_ABSFLAG)) and
       (FFormulas[i].Ref.Row >= (APtgsArea.Row1 and not ROW_ABSFLAG)) and (FFormulas[i].Ref.Row <= (APtgsArea.Row2 and not ROW_ABSFLAG)) then begin
      ADep.AddRef(@FFormulas[i].Ref);
    end;
  end;
end;

procedure TXLSCalcChainBuilder.AddDependencies(ADep: TDepRefsItem; APtgsArea3d: PXlsPtgsArea3d);
var
  i: integer;
begin
  for i := APtgsArea3d.Sheet1 to APtgsArea3d.Sheet2 do
    AddDependencies(ADep,i,PXlsPtgsArea(APtgsArea3d));
end;

procedure TXLSCalcChainBuilder.AddDependencies(ADep: TDepRefsItem; APtgsRef3d: PXlsPtgsRef3d);
var
  i: integer;
begin
  for i := APtgsRef3d.Sheet1 to APtgsRef3d.Sheet2 do
    AddDependencies(ADep,i,PXlsPtgsRef(APtgsRef3d));
end;

{$ifdef XLS_BIFF}
procedure TXLSCalcChainBuilder.AddDependencies(ADep: TDepRefsItem; const ASheetId: integer; APtgsRef: PXLSPtgsRef97);
var
  i: integer;
begin
  i := FindFormula(ASheetId,APtgsRef.Col,APtgsRef.Row);
  if i >= 0 then
    ADep.AddRef(@FFormulas[i].Ref);
end;

procedure TXLSCalcChainBuilder.AddDependencies(ADep: TDepRefsItem; const ASheetId: integer; APtgsArea: PXLSPtgsArea97);
var
  i,i1,i2: integer;
begin
  i1 := FindFormulaClosest(ASheetId,0,APtgsArea.Row1);
  i2 := FindFormulaClosest(ASheetId,XLS_MAXCOL,APtgsArea.Row2);
  for i := i1 to i2 do begin
    if FFormulas[i].Ref.Row > APtgsArea.Row2 then
      Exit;
    if (FFormulas[i].Ref.Col >= APtgsArea.Col1) and (FFormulas[i].Ref.Col <= APtgsArea.Col2) and
       (FFormulas[i].Ref.Row >= APtgsArea.Row1) and (FFormulas[i].Ref.Row <= APtgsArea.Row2) then
      ADep.AddRef(@FFormulas[i].Ref);
  end;
end;
{$endif}

procedure TXLSCalcChainBuilder.AddDependencies(ADep: TDepRefsItem; const ASheetId, ACol, ARow: integer);
var
  i: integer;
begin
  i := FindFormula(ASheetId,ACol,ARow);
  if i >= 0 then
    ADep.AddRef(@FFormulas[i].Ref);
end;

procedure TXLSCalcChainBuilder.AddDependencies(ADep: TDepRefsItem; const ASheetId, ACol1, ARow1, ACol2, ARow2: integer);
var
  i,i1,i2: integer;
begin
  i1 := FindFormulaClosest(ASheetId,0,ARow1);
  i2 := FindFormulaClosest(ASheetId,XLS_MAXCOL,ARow2);
  for i := i1 to i2 do begin
    if FFormulas[i].Ref.Row > Longword(ARow2) then
      Exit;
    if (FFormulas[i].Ref.Col >= ACol1) and (FFormulas[i].Ref.Col <= ACol2) and
       (FFormulas[i].Ref.Row >= Longword(ARow1)) and (FFormulas[i].Ref.Row <= Longword(ARow2)) then
      ADep.AddRef(@FFormulas[i].Ref);
  end;
end;

{$ifdef XLS_BIFF}
procedure TXLSCalcChainBuilder.AddDependencies(ADep: TDepRefsItem; APtgsRef3d: PXLSPtgsRef3d97);
var
  i: integer;
  i1,i2: integer;
begin
  FManager._ExtNames97.Get3dSheets(APtgsRef3d.ExtSheetIndex,i1,i2);
  if i1 = i2 then
    AddDependencies(ADep,i1,APtgsRef3d.Col and $3FFF,APtgsRef3d.Row)
  else begin
    for i := i1 to i2 do
      AddDependencies(ADep,i,APtgsRef3d.Col and $3FFF,APtgsRef3d.Row)
  end;
end;

procedure TXLSCalcChainBuilder.AddDependencies(ADep: TDepRefsItem; APtgsArea3d: PXLSPtgsArea3d97);
var
  i: integer;
  i1,i2: integer;
begin
  FManager._ExtNames97.Get3dSheets(APtgsArea3d.ExtSheetIndex,i1,i2);
  if i1 = i2 then
    AddDependencies(ADep,i1,APtgsArea3d.Col1 and $3FFF,APtgsArea3d.Row1,APtgsArea3d.Col2 and $3FFF,APtgsArea3d.Row2)
  else begin
    for i := i1 to i2 do
      AddDependencies(ADep,i,APtgsArea3d.Col1 and $3FFF,APtgsArea3d.Row1,APtgsArea3d.Col2 and $3FFF,APtgsArea3d.Row2)
  end;
end;
{$endif}

function TXLSCalcChainBuilder.Build(ACalcChain: TXLSCalcChain): boolean;
var
  i: integer;
  Sorter: TTopoSort;
  Dep: TDepRefsItem;
begin
  FCalcChain := ACalcChain;
  FCalcChain.Clear;
  FHasIgnoreCircularFunc := False;

  Clear;

  for i := 0 to FManager.Worksheets.Count - 1 do
    CollectFormulas(i,FManager.Worksheets[i].Cells);

  for i := 0 to FFormulaCnt - 1 do begin
    Dep := CheckFormulaContent(@FFormulas[i]);
    if Dep.Count = 1 then begin
      ACalcChain.Add(Dep.Items[0]);
      Dep.Free;
    end
    else
      FDepList.Add(Dep);
  end;

  // Will not find circular references where the formula referers to itself
  // such as SUM(A1:A5) and the SUM function is in cell A3.
  Sorter := TTopoSort.Create;
  try
    Result := Sorter.Sort(FDepList);

    if Result then begin
      for i := Sorter.Output.Count - 1 downto 0 do begin
        ACalcChain.Add(PXLS3dCompactRef(Sorter.Output[i]));

//        FManager.Errors.Message(ColRowToRefStr(PDepRef(Sorter.Output[i]).Col,PDepRef(Sorter.Output[i]).Row));
      end;
    end
    else begin
      for i := 0 to Sorter.Circular.Count - 1 do
        FManager.Errors.Warning(Sorter.Circular[i].AsText,XLSWARN_CIRCULAR_FORMULA);
    end;
    ACalcChain.Done;

//    FManager.Errors.Message(DepListAsText);
  finally
    Sorter.Free;
  end;

end;

procedure TXLSCalcChainBuilder.Clear;
begin
  if FFormulas <> Nil then begin
    FreeMem(FFormulas);
    FFormulas := Nil;
  end;
  FDepList.Clear;
  FFormulaCnt := 0;
  FAllocSize := 0;
end;

constructor TXLSCalcChainBuilder.Create(AManager: TXc12Manager);
begin
  FManager := AManager;
  FDepList := TDepRefsItems.Create;;
  SetLength(FPtgsList,$FF);
end;

function TXLSCalcChainBuilder.DepListAsText: AxUCString;
begin
  Result := FDepList.AsText;
end;

destructor TXLSCalcChainBuilder.Destroy;
begin
  Clear;
  FDepList.Free;
  inherited;
end;

function TXLSCalcChainBuilder.FindFormulaClosest(ASheetId, ACol,ARow: integer): integer;
var
  First: Integer;
  Last: Integer;
  V: integer;
  P: PXLS3dCompactRef;
begin
  Result := 0;
  First := 0;
  Last := FFormulaCnt - 1;

  while First <= Last do begin
    Result := (First + Last) div 2;
    P := @FFormulas[Result].Ref;
    V := P.SheetId - ASheetId;
    if V = 0 then begin
      V := Integer(P.Row) - ARow;
      if V = 0 then
        V := P.Col - ACol;
    end;
    if V = 0 then
      Exit
    else If V > 0 then
      Last := Result - 1
    else
      First := Result + 1;
  end;
  if Result < 0 then
    Result := 0
  else if Result >= FFormulaCnt then
    Result := FFormulaCnt - 1;
end;

procedure TXLSCalcChainBuilder.IgnoreCircular(ASender: TObject);
var
  i: integer;
  Item: TCircPtgsItem;
begin
  Item := TCircPtgsItem(ASender);
  if Item <> Nil then begin
    for i := 0 to FPtgsCount - 1 do begin
      if FPtgsList[i] = Item[0] then begin
        FPtgsList[i] := Nil;
        Break;
      end;
    end;
  end;
end;

procedure TXLSCalcChainBuilder.RemoveRefsWhereCircularIsOk;
var
  i: integer;
  Ptgs: PXLSPtgs;
  Stack: TCircPtgsStack;
begin
  Stack := TCircPtgsStack.Create;
  try
    Stack.OnFoundIgnoreCircular := IgnoreCircular;
    for i := 0 to FPtgsCount - 1 do begin
      Ptgs := FPtgsList[i];
      if Ptgs = Nil then
        Continue;
      case GetBasePtgs(Ptgs.Id) of
        xptgNone    : raise XLSRWException.Create('Illegal ptgs');
        xptgOpAdd,
        xptgOpSub,
        xptgOpMult,
        xptgOpDiv,
        xptgOpPower,
        xptgOpConcat,
        xptgOpLT,
        xptgOpLE,
        xptgOpEQ,
        xptgOpGE,
        xptgOpGT,
        xptgOpNE      : Stack.OpIgnore;

        xptgWS,
        xptgOpPercent,
        xptgLPar,
        xptgOpUPlus,
        xptgOpUMinus  : ;

        xptgOpIsect,
        xptgOpUnion,
        xptgOpRange   : Stack.Op;

        xptgArray     : Stack.ArrayIgnore(PXLSPtgsArray(Ptgs).Cols,PXLSPtgsArray(Ptgs).Rows);

        xptgStr,
        xptgErr,
        xptgBool,
        xptgInt,
        xptgNum       : Stack.PushIgnore;

        xptgFunc      : Stack.Func(PXLSPtgsFunc(Ptgs).FuncId,G_XLSExcelFuncNames.ArgCount(PXLSPtgsFunc(Ptgs).FuncId));
        xptgFuncVar   : Stack.Func(PXLSPtgsFuncVar(Ptgs).FuncId,PXLSPtgsFuncVar(Ptgs).ArgCount);
        xptgName      : Stack.PushIgnore; // raise XLSRWException.Create('Name not excpected here');
        xptgRef       : Stack.Push(FPtgsList[i]);
        xptgArea      : Stack.Push(FPtgsList[i]);
        xptgRefErr    : Stack.Push(FPtgsList[i]);
        xptgAreaErr   : Stack.Push(FPtgsList[i]);

        xptgRef1d     : Stack.Push(FPtgsList[i]);
        xptgRef1dErr  : Stack.PushIgnore;
        xptgRef3d     : Stack.Push(FPtgsList[i]);
        xptgXRef1d,
        xptgXRef3d    : Stack.PushIgnore;
        xptgArea1d    : Stack.Push(FPtgsList[i]);
        xptgArea3d    : Stack.Push(FPtgsList[i]);
        xptgXArea1d,
        xptgXArea3d   : Stack.PushIgnore;

        xptgArrayFmlaChild: Stack.PushIgnore;

        xptgDataTableFmla: Stack.PushIgnore;

        xptgDataTableFmlaChild: Stack.PushIgnore;

        xptgTable     : Exit;
        xptgTableSpecial: Exit;
        xptgMissingArg: Stack.PushIgnore;
        xptgUserFunc  : Stack.FuncIgnore(PXLSPtgsUserFunc(Ptgs).ArgCount);

        xptgStr97,
        xptgInt97,
        xptgAttr97,
        xptgName97,
        xptgNameX97   : Stack.PushIgnore;

        xptgRef97     : Stack.Push(FPtgsList[i]);
        xptgRef3d97   : Stack.Push(FPtgsList[i]);
        xptgArea97    : Stack.Push(FPtgsList[i]);
        xptgArea3d97  : Stack.Push(FPtgsList[i]);

        xptgRefErr97  : Stack.PushIgnore;

        xptgFunc97    : Stack.Func(PXLSPtgsFunc(Ptgs).FuncId,G_XLSExcelFuncNames.ArgCount(PXLSPtgsFunc(Ptgs).FuncId));

        xptgFuncVar97 : Stack.Func(PXLSPtgsFuncVar(Ptgs).FuncId and $7F,PXLSPtgsFuncVar(Ptgs).ArgCount and $7FFF);

        else raise XLSRWException.CreateFmt('Unknown Ptg "%.2X"',[Ptgs.Id]);
      end;
    end;
  finally
    Stack.Free;
  end;
end;

function TXLSCalcChainBuilder.FindFormula(ASheetId, ACol, ARow: integer): integer;
var
  First: Integer;
  Last: Integer;
  V: integer;
  P: PXLS3dCompactRef;
begin
  First := 0;
  Last := FFormulaCnt - 1;

  while First <= Last do begin
    Result := (First + Last) div 2;
    P := @FFormulas[Result].Ref;
    V := P.SheetId - ASheetId;
    if V = 0 then begin
      V := Integer(P.Row) - ARow;
      if V = 0 then
        V := P.Col - ACol;
    end;
    if V = 0 then
      Exit
    else If V > 0 then
      Last := Result - 1
    else
      First := Result + 1;
  end;
  Result := -1;
end;

function TXLSCalcChainBuilder.CheckFormulaContent(AData: PCCFormulaData): TDepRefsItem;
begin
  Result := TDepRefsItem.Create(@AData.Ref);
  FPtgsCount := 0;
  CollectFormulaPtgs(AData.Ptgs,AData.PtgsSize);
  if FHasIgnoreCircularFunc then
    RemoveRefsWhereCircularIsOk;
  CheckFormulaRefs(Result,@AData.Ref);
  Result.EndAddRefs;
end;

procedure TXLSCalcChainBuilder.CheckFormulaRefs(ADep: TDepRefsItem; ADepRef: PXLS3dCompactRef);
var
  i: integer;
  Ptgs: PXLSPtgs;
begin
  for i := 0 to FPtgsCount - 1 do begin
    Ptgs := FPtgsList[i];
    if Ptgs = Nil then
      Continue;
    case GetBasePtgs(Ptgs.Id) of
      xptgRef       : AddDependencies(ADep,ADepRef.SheetId,PXLSPtgsRef(Ptgs));
      xptgArea      : AddDependencies(ADep,ADepRef.SheetId,PXLSPtgsArea(Ptgs));
      xptgRef1d     : AddDependencies(ADep,PXlsPtgsRef1d(Ptgs).Sheet,PXLSPtgsRef(Ptgs));
      xptgArea1d    : AddDependencies(ADep,PXlsPtgsArea1d(Ptgs).Sheet,PXLSPtgsArea(Ptgs));
      xptgRef3d     : AddDependencies(ADep,PXLSPtgsRef3d(Ptgs));
      xptgArea3d    : AddDependencies(ADep,PXLSPtgsArea3d(Ptgs));
{$ifdef XLS_BIFF}
      xptgRef97     : AddDependencies(ADep,ADepRef.SheetId,PXLSPtgsRef97(Ptgs));
      xptgArea97    : AddDependencies(ADep,ADepRef.SheetId,PXLSPtgsArea97(Ptgs));
      xptgRef3d97   : AddDependencies(ADep,PXLSPtgsRef3d97(Ptgs));
      xptgArea3d97  : AddDependencies(ADep,PXLSPtgsArea3d97(Ptgs));
{$endif}
    end;
  end;
end;

procedure TXLSCalcChainBuilder.CollectFormulaPtgs(APtgs: PXLSPtgs; const APtgsSz: integer);
var
  L: integer;
  P: Pointer;
  Sz: integer;
  Ptgs: PXLSPtgs;
  PtgsSz: integer;
begin
  PtgsSz := APtgsSz;
  while PtgsSz > 0 do begin
    case GetBasePtgs(APtgs.Id) of
      xptgName      : begin
        L := SizeOf(TXLSPtgsName);
        if (PXLSPtgsName(APtgs).NameId <> XLS_NAME_UNKNOWN) and (FManager.Workbook.DefinedNames[PXLSPtgsName(APtgs).NameId].SimpleName > xsntNone) then begin
          Ptgs := FManager.Workbook.DefinedNames[PXLSPtgsName(APtgs).NameId].Ptgs;
          Sz := FManager.Workbook.DefinedNames[PXLSPtgsName(APtgs).NameId].PtgsSz;
          CollectFormulaPtgs(Ptgs,Sz);
          APtgs := PXLSPtgs(NativeInt(APtgs) + L);
          Dec(PtgsSz,L);
          Continue;
        end;
      end;
      xptgStr       : L := FixedSzPtgsStr + PXLSPtgsStr(APtgs).Len * 2;
      xptgUserFunc  : L := FixedSzPtgsUserFunc + PXLSPtgsUserFunc(APtgs).Len * 2;
      xptgFuncVar   : begin
        if PXLSPtgsFuncVar(APtgs).FuncId in [XLSFUNCID_LOOKUP,XLSFUNCID_HLOOKUP,XLSFUNCID_VLOOKUP] then
          FHasIgnoreCircularFunc := True;
        L := G_XLSPtgsSize[APtgs.Id];
      end;

      xptg_EXCEL_97 : begin
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgs));
        Dec(PtgsSz,SizeOf(TXLSPtgs));
        Continue;
      end;

      xptg_ARRAYCONSTS_97 : begin
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgs));

        L := PWord(APtgs)^;
        Dec(PtgsSz,L + SizeOf(TXLSPtgs) + SizeOf(word));
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(word));

        Continue;
      end;

      xptgName97    : begin
        L := SizeOf(TXLSPtgsName97);
        if FManager.Workbook.DefinedNames[PXLSPtgsName97(APtgs).NameId - 1].SimpleName > xsntNone then begin
          Ptgs := FManager.Workbook.DefinedNames[PXLSPtgsName97(APtgs).NameId - 1].Ptgs;
          Sz := FManager.Workbook.DefinedNames[PXLSPtgsName97(APtgs).NameId - 1].PtgsSz;
          CollectFormulaPtgs(Ptgs,Sz);
          APtgs := PXLSPtgs(NativeInt(APtgs) + L);
          Dec(PtgsSz,L);
          Continue;
        end;
      end;

      xptgStr97      : begin
        L := 3;
        P := Pointer(NativeInt(APtgs) + 1);
        if PByteArray(P)[1] = 0 then
          Inc(L,Byte(P^))
        else
          Inc(L,Byte(P^) * 2);
      end;

      xptgAttr97    : begin
        L := 1;
        P := Pointer(NativeInt(APtgs) + 1);
        if (Byte(P^) and $04) = $04 then begin
          Inc(L);
          P := Pointer(NativeInt(P) + 1);
          Inc(L,(Word(P^) + 2) * SizeOf(word) - 3);
        end;
        Inc(L,3);
      end;

      xptgFuncVar97,
      xptgFuncVarV97,
      xptgFuncVarA97 : begin
        if (PXLSPtgsFuncVar(APtgs).FuncId and $7FFF) in [XLSFUNCID_LOOKUP,XLSFUNCID_HLOOKUP,XLSFUNCID_VLOOKUP] then
          FHasIgnoreCircularFunc := True;
        L := G_XLSPtgsSize[APtgs.Id];
      end;

      xptgMemErr97,
      xptgMemArea97,
      xptgMemFunc97  : begin
        P := Pointer(NativeInt(APtgs) + 1);
        L := 3 + PWord(P)^;
      end

      else             L := G_XLSPtgsSize[APtgs.Id];
    end;

    if L < 0 then
      raise XLSRWException.CreateFmt('Unknown ptgs %.2X',[APtgs.Id]);

    if FPtgsCount > High(FPtgsList) then
      SetLength(FPtgsList,Length(FPtgsList) + $FF);
    FPtgsList[FPtgsCount] := APtgs;
    Inc(FPtgsCount);
    APtgs := PXLSPtgs(NativeInt(APtgs) + L);
    Dec(PtgsSz,L);
  end;
end;

procedure TXLSCalcChainBuilder.CollectFormulas(ASheetId: integer; ACells: TXLSCellMMU);
var
  Ptgs: PXLSPtgs;
  PtgsSize: integer;
  S: AxUCString;
begin
  ACells.BeginIterate;
  while ACells.IterateNext do begin
    if (ACells.IterCellType in XLSCellTypeFormulas) and not (ACells.IterFormulaType in [xcftArrayChild,xcftArrayChild97,xcftDataTableChild]) then begin
      try
        PtgsSize := ACells.IterGetFormulaPtgs(Ptgs);
        if FFormulaCnt >= FAllocSize then begin
          Inc(FAllocSize,$FFF);
          ReAllocMem(FFormulas,SizeOf(TCCFormulaData) * FAllocSize);
        end;
        FFormulas[FFormulaCnt].Ref.SheetId := ASheetId;
        FFormulas[FFormulaCnt].Ref.Col := ACells.IterCellCol;
        FFormulas[FFormulaCnt].Ref.Row := ACells.IterCellRow;
        FFormulas[FFormulaCnt].Ptgs := Ptgs;
        FFormulas[FFormulaCnt].PtgsSize := PtgsSize;
        Inc(FFormulaCnt);
      except
        on e: XLSRWException do begin
          S := FManager.Worksheets[ASheetId].Name + '!' + ColRowToRefStr(ACells.IterCellCol,ACells.IterCellRow);
          raise XLSRWException.Create(e.message + ' [' + S + ']');
        end;
      end;
    end;
  end;
end;

{ TXLSCalcChain }

procedure TXLSCalcChain.Add(ARef: PXLS3dCompactRef);
begin
  if FCount >= FAllocSize then begin
    Inc(FAllocSize,$FFF);
    ReAllocMem(FItems,FAllocSize * SizeOf(TXLS3dCompactRef));
  end;
  FItems[FCount] := ARef^;
  Inc(FCount);
end;

procedure TXLSCalcChain.Append(ACalcChain: TXLSCalcChain);
var
  C: integer;
begin
  C := FCount;
  Inc(FCount,ACalcChain.FCount);
  ReAllocMem(FItems,FCount * SizeOf(TXLS3dCompactRef));
  Move(ACalcChain.FItems^,FItems[C],ACalcChain.FCount * SizeOf(TXLS3dCompactRef));
end;

procedure TXLSCalcChain.Clear;
begin
  if FItems <> Nil then begin
    FreeMem(FItems);
    FItems := Nil;
  end;
  FCount := 0;
  FAllocSize := 0;
end;

function TXLSCalcChain.Count: integer;
begin
  Result := FCount;
end;

destructor TXLSCalcChain.Destroy;
begin
  Clear;
  inherited;
end;

procedure TXLSCalcChain.Done;
begin
  ReAllocMem(FItems,FCount * SizeOf(TXLS3dCompactRef));
end;

function TXLSCalcChain.GetItems(Index: integer): PXLS3dCompactRef;
begin
  Result := @FItems[Index];
end;

{ TCircPtgsItem }

procedure TCircPtgsItem.Add(APtgs: PXLSPtgs);
begin
  SetLength(FPtgs,Length(FPtgs) + 1);
  FPtgs[High(FPtgs)] := APtgs;
end;

procedure TCircPtgsItem.Add(AItem: TCircPtgsItem);
var
  i: integer;
begin
  for i := 0 to AItem.Count - 1 do
    Add(AItem[i]);
end;

function TCircPtgsItem.Count: integer;
begin
  Result := Length(FPtgs);
end;

function TCircPtgsItem.GetItems(Index: integer): PXLSPtgs;
begin
  Result := FPtgs[Index];
end;

{ TCircPtgsStack }

procedure TCircPtgsStack.ArrayIgnore(const ACols, ARows: integer);
begin
  Count := Count - (ACols * ARows - 1);
end;

constructor TCircPtgsStack.Create;
begin
  inherited Create;
end;

procedure TCircPtgsStack.Func(const AFuncId,AArgCount: integer);
begin
  case AFuncId of
    XLSFUNCID_LOOKUP : FFoundIgnoreCircularEvent(Items[Count - 1 - 2]);
    XLSFUNCID_HLOOKUP: FFoundIgnoreCircularEvent(Items[Count - 1 - 2]);
    XLSFUNCID_VLOOKUP: FFoundIgnoreCircularEvent(Items[Count - 1 - 2]);
  end;
  FuncIgnore(AArgCount);
end;

procedure TCircPtgsStack.FuncIgnore(const AArgCount: integer);
begin
  Count := Count - (AArgCount - 1);
end;

function TCircPtgsStack.GetItems(Index: integer): TCircPtgsItem;
begin
  Result := TCircPtgsItem(inherited Items[Index]);
end;

procedure TCircPtgsStack.Op;
var
  I1,I2: TCircPtgsItem;
begin
  I1 := TCircPtgsItem(Items[Count - 1]);
  I2 := TCircPtgsItem(Items[Count - 2]);
  I2.Add(I1);
  Delete(Count - 1);
end;

procedure TCircPtgsStack.OpIgnore;
begin
  Delete(Count - 1);
end;

procedure TCircPtgsStack.Push(APtgs: PXLSPtgs);
var
  Item: TCircPtgsItem;
begin
  Item := TCircPtgsItem.Create;
  Item.Add(APtgs);
  Add(Item);
end;

procedure TCircPtgsStack.PushIgnore;
begin
  Add(Nil);
end;

procedure TCircPtgsStack.SetItems(Index: integer; const Value: TCircPtgsItem);
begin
  inherited Items[Index] := Value;
end;

end.
