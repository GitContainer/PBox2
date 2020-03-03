unit XLSFmlaDebugData5;

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

uses Classes, SysUtils, Contnrs,
     XLSUtils5, XLSFormulaTypes5, XLSEvaluate5;

type TFmlaDebugFmlaItem = class(TObject)
protected
     FSteps     : integer;
     FPtgs      : PXLSPtgs;
     FPtgsOffset: integer;
     FResult    : TXLSFormulaValue;
public
     property Steps: integer read FSteps;
     property PtgsOffset: integer read FPtgsOffset;
     property Result: TXLSFormulaValue read FResult;
     property Ptgs  : PXLSPtgs read FPtgs  ;
     end;

type TFmlaDebugFmlaItems = class(TObjectList)
private
     function GetItems(Index: integer): TFmlaDebugFmlaItem;
protected
public
     constructor Create;

     procedure Add(const ASteps, APtgsOffset: integer; const APtgs: PXLSPtgs; const AResult: TXLSFormulaValue);

     property Items[Index: integer]: TFmlaDebugFmlaItem read GetItems; default;
     end;

type TFmlaDebugItem = class(TObject)
protected
     FP1,FP2  : integer;
     FPtgs    : PXLSPtgs;
     FPtgsOffs: integer;
     FIndex   : integer;
public
     property P1: integer read FP1 write FP1;
     property P2: integer read FP2 write FP2;
     property _Ptgs: PXLSPtgs read FPtgs write FPtgs;
     property PtgsOffs: integer read FPtgsOffs write FPtgsOffs;
     property Index: integer read FIndex write FIndex;
     end;

type TFmlaDebugItems = class(TObjectList)
private
     function  GetItems(Index: integer): TFmlaDebugItem;
protected
     FSteps          : integer;
     FFmlaItems      : TFmlaDebugFmlaItems;
     FCurrPtgs       : PXLSPtgs;
     FCurrPtgsOffs   : integer;
     FCollectFormulas: boolean;
     FTempPosStack   : TIntegerStack;
public
     constructor Create;
     destructor Destroy; override;

     function  Find(const APtgs: PXLSPtgs): TFmlaDebugItem; overload;
     function  Find(const APtgsOffs: integer): TFmlaDebugItem; overload;

     procedure Add(const APtgs: PXLSPtgs; const AP1,AP2: integer);
     procedure UpdatePos(const APtgs: PXLSPtgs; const APos: integer);
     procedure UpdateOffs(const APtgs: PXLSPtgs; const APtgsOffs, AIndex: integer);

     property Items[Index: integer]: TFmlaDebugItem read GetItems; default;
     property FmlaItems: TFmlaDebugFmlaItems read FFmlaItems;
     property Steps: integer read FSteps write FSteps;
     property CurrPtgs: PXLSPtgs read FCurrPtgs write FCurrPtgs;
     property CurrPtgsOffs: integer read FCurrPtgsOffs write FCurrPtgsOffs;
     property CollectFormulas: boolean read FCollectFormulas write FCollectFormulas;
     property TempPosStack: TIntegerStack read FTempPosStack;
     end;

implementation

{ TFmlaDebugItems }

procedure TFmlaDebugItems.Add(const APtgs: PXLSPtgs; const AP1, AP2: integer);
var
  Item: TFmlaDebugItem;
begin
  Item := TFmlaDebugItem.Create;
  Item._Ptgs := APtgs;
  Item.PtgsOffs := -1;
  Item.P1 := AP1;
  Item.P2 := AP2;
  inherited Add(Item);
end;

constructor TFmlaDebugItems.Create;
begin
  inherited Create;
  FFmlaItems := TFmlaDebugFmlaItems.Create;
  FTempPosStack := TIntegerStack.Create;
end;

destructor TFmlaDebugItems.Destroy;
begin
  FFmlaItems.Free;
  FTempPosStack.Free;
  inherited;
end;

function TFmlaDebugItems.Find(const APtgs: PXLSPtgs): TFmlaDebugItem;
var
  i: integer;
begin
  for i := 0 to Count - 1 do begin
    if Items[i]._Ptgs = APtgs then begin
      Result := Items[i];
      Exit;
    end;
  end;
  Result := Nil;
end;

function TFmlaDebugItems.Find(const APtgsOffs: integer): TFmlaDebugItem;
var
  i: integer;
begin
  for i := 0 to Count - 1 do begin
    if Items[i].PtgsOffs = APtgsOffs then begin
      Result := Items[i];
      Exit;
    end;
  end;
  Result := Nil;
end;

function TFmlaDebugItems.GetItems(Index: integer): TFmlaDebugItem;
begin
  Result := TFmlaDebugItem(inherited Items[Index]);
end;

procedure TFmlaDebugItems.UpdateOffs(const APtgs: PXLSPtgs; const APtgsOffs, AIndex: integer);
var
  Item: TFmlaDebugItem;
begin
  Item := Find(APtgs);
  if Item <> Nil then begin
    Item.PtgsOffs := APtgsOffs;
    Item.Index := AIndex;
//    Item._Ptgs := Nil;
  end;
end;

procedure TFmlaDebugItems.UpdatePos(const APtgs: PXLSPtgs; const APos: integer);
var
  Item: TFmlaDebugItem;
begin
  Item := Find(APtgs);
  if Item <> Nil then
    Item.P2 := APos;
end;

{ TFmlaDebugFmlaItems }

procedure TFmlaDebugFmlaItems.Add(const ASteps, APtgsOffset: integer; const APtgs: PXLSPtgs; const AResult: TXLSFormulaValue);
var
  Item: TFmlaDebugFmlaItem;
begin
  Item := TFmlaDebugFmlaItem.Create;
  Item.FSteps := ASteps;
  Item.FPtgsOffset := APtgsOffset;
  Item.FPtgs := APtgs;
  Item.FResult := AResult;
  inherited Add(Item);
end;

constructor TFmlaDebugFmlaItems.Create;
begin
  inherited Create;
end;

function TFmlaDebugFmlaItems.GetItems(Index: integer): TFmlaDebugFmlaItem;
begin
  Result := TFmlaDebugFmlaItem(inherited Items[Index]);
end;

end.
