unit XLSEvaluate5;

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

uses Classes, SysUtils, Contnrs, Math, {$ifdef DELPHI_6_OR_LATER} Masks, {$endif}
     Xc12Utils5,
     XLSUtils5, XLSCellMMU5;

type TXLSFormulaValueType = (xfvtUnknown,xfvtFloat,xfvtBoolean,xfvtError,xfvtString,xfvtAreaList,xfvtTableSpecial);

type TXLSFormulaRefType = (xfrtNone,xfrtRef,xfrtArea,xfrtXArea,xfrtAreaList,xfrtArray,xfrtArrayArg);

type PXLSFormulaValue = ^TXLSFormulaValue;
     TXLSFormulaValue = record
     Empty    : boolean;
     RefType  : TXLSFormulaRefType;
     Cells    : TXLSCellMMU;
     Col1,Col2: integer;
     Row1,Row2: integer;
     vStr     : AxUCString;
     case ValType: TXLSFormulaValueType of
       xfvtUnknown     : ();
       xfvtFloat       : (vFloat : double);
       xfvtBoolean     : (vBool  : boolean);
       xfvtError       : (vErr   : TXc12CellError);
       xfvtTableSpecial: (vSpec  : integer);
       // Can be TXLSArrayItem (RefType = xfrtArray) and TCellAreas (RefType = xfrtAreaList).
       xfvtAreaList    : (vSource: TObject);
     end;

type PXLSVarValue = ^TXLSVarValue;
     TXLSVarValue = record
     Empty: boolean;
     vStr : AxUCString;
     case ValType: TXLSFormulaValueType of
       xfvtFloat   : (vFloat: double);
       xfvtBoolean : (vBool : boolean);
       xfvtError   : (vErr  : TXc12CellError);
     end;

type TXLSArrayItem = class(TObject)
private
     function  GetItems(ACol,ARow: integer): PXLSVarValue;
     function  GetHeight: integer;
     function  GetWidth: integer;
     function  GetAsFloat(ACol, ARow: integer): double;
     function  GetAsBoolean(ACol, ARow: integer): boolean;
     function  GetAsError(ACol, ARow: integer): TXc12CellError;
     function  GetAsString(ACol, ARow: integer): AxUCString;
     procedure SetAsFloat(ACol,ARow: integer; AValue: double);
     procedure SetAsBoolean(ACol, ARow: integer; const Value: boolean);
     procedure SetAsError(ACol, ARow: integer; const Value: TXc12CellError);
     procedure SetAsString(ACol, ARow: integer; const Value: AxUCString);
     function  GetHit(ACol, ARow: integer): boolean;
protected
     FValues: array of array of TXLSVarValue;
     // Ignore second dimension (width or height) on vectors. Only when reading values.
     FVectorMode: boolean;
     FTag: integer;

     procedure CheckVectorMode(var ACol,ARow: integer);
public
     constructor Create(AWidth,AHeight: integer);
     constructor CreateFV(AValue: TXLSFormulaValue);

     procedure Clear;
     procedure Assign(AValue: TXLSArrayItem);

     procedure MakeEqualSize(AValue: TXLSArrayItem);
     procedure FillError;

     function  Add(const ACol,ARow: integer; const AValue: TXLSFormulaValue): boolean; overload;
     procedure Add(const ACol,ARow: integer; const AValue: double); overload;
     procedure Add(const ACol,ARow: integer; const AValue: AxUCString); overload;
     procedure Add(const ACol,ARow: integer; const AValue: boolean); overload;
     procedure Add(const ACol,ARow: integer; const AValue: TXc12CellError); overload;

     function  GetAsFormulaValue(ACol,ARow: integer): TXLSFormulaValue;

     function  IsVector: boolean;

     property Width: integer read GetWidth;
     property Height: integer read GetHeight;

     property VectorMode: boolean read FVectorMode write FVectorMode;
     property Hit[ACol,ARow: integer]: boolean read GetHit;
     property AsBoolean[ACol,ARow: integer]: boolean read GetAsBoolean write SetAsBoolean;
     property AsFloat[ACol,ARow: integer]: double read GetAsFloat write SetAsFloat;
     property AsString[ACol,ARow: integer]: AxUCString read GetAsString write SetAsString;
     property AsError[ACol,ARow: integer]: TXc12CellError read GetAsError write SetAsError;

     property Tag: integer read FTag write FTag;

     property Items[ACol,ARow: integer]: PXLSVarValue read GetItems; default;
     end;

type TXLSArrayItemIterate = class(TXLSArrayItem)
protected
     FCol: integer;
     FRow: integer;
     FLoopOnVectors: boolean;
public
     procedure BeginIterate(ALoopOnVectors: boolean = True);
     function  IterateNext: boolean;

     function  IterFloat: double;
     function  IterError: TXc12CellError;
     end;

type TXLSCellsSource = class(TObject)
private
     function  GetAsFloat(ACol, ARow: integer): double;
     procedure SetAsFloat(ACol, ARow: integer; const Value: double);
     function GetHeight: integer;
     function GetWidth: integer;
protected
     FCol1,FCol2: integer;
     FRow1,FRow2: integer;
     FIterCol,FIterRow: integer;
     FCells : TXLSCellMMU;
     FArray : TXLSArrayItem;
     FEmpty : boolean;
     FError : TXc12CellError;
     FOwnsArray: boolean;
public
     constructor Create(AWidth,AHeight: integer); overload;
     constructor Create(AArray : TXLSArrayItem); overload;
     constructor Create(ACells: TXLSCellMMU; ACol1,ARow1,ACol2,ARow2: integer); overload;
     destructor Destroy; override;

     procedure BeginIterate;
     function  IterateNext: boolean;
     function  IterAsFloat: double;

     property AsFloat[ACol,ARow: integer]: double read GetAsFloat write SetAsFloat;

     property LastWasEmpty: boolean read FEmpty;
     property LastError: TXc12CellError read FError;

     property Col1: integer read FCol1;
     property Row1: integer read FRow1;
     property Col2: integer read FCol2;
     property Row2: integer read FRow2;
     property Width: integer read GetWidth;
     property Height: integer read GetHeight;
     property IterCol: integer read FIterCol;
     property IterRow: integer read FIterRow;
     end;

type TXLSDbCondOperator = (xdcoNone,xdcoLT,xdcoLE,xdcoEQ,xdcoGE,xdcoGT,xdcoNE,xdcoError);

type TXLSDbField = class;

     TXLSDbCondition = class(TObject)
private
     function GetValue: PXLSVarValue;
protected
     FField  : TXLSDbField;
     FOp     : TXLSDbCondOperator;
     FValue  : TXLSVarValue;
public
     destructor Destroy; override;

     function  DoCondition(AValue: PXLSVarValue): boolean; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}

     property Field: TXLSDbField read FField write FField;
     property Op: TXLSDbCondOperator read FOp write FOp;
     property Value: PXLSVarValue read GetValue;
     end;

     TXLSDbConditionAnd = class(TObjectList)
private
     function GetItems(Index: integer): TXLSDbCondition;
protected
public
     constructor Create;

     function Add(AField: TXLSDbField; const ACondStr: AxUCString): TXLSDbCondition;

     property Items[Index: integer]: TXLSDbCondition read GetItems; default;
     end;

     TXLSDbConditionOr = class(TObjectList)
private
     function GetItems(Index: integer): TXLSDbConditionAnd;
protected
public
     constructor Create;

     function  Add: TXLSDbConditionAnd;

     property Items[Index: integer]: TXLSDbConditionAnd read GetItems; default;
     end;

     TXLSDbField = class(TObject)
protected
     FName : AxUCString;
     FIndex: integer;
     // True if the field has a condition.
     FUsed : boolean;
public
     destructor Destroy; override;

     property Name: AxUCString read FName write FName;
     property Index: integer read FIndex;
     property Used: boolean read FUsed write FUsed;
     end;

type TXLSFmlaDatabase = class(TObjectList)
private
     function GetItems(Index: integer): TXLSDbField;
protected
     FFvDb        : TXLSFormulaValue;
     FFvCriteria  : TXLSFormulaValue;
     FConditions  : TXLSDbConditionOr;
     FFieldXLate  : array of integer;
     FSrcField    : integer;

     FResCount    : integer;
     FResCountA   : integer;
     FResSum      : double;

     // Copies after recursive call.
     FResCount2   : integer;
     FResCountA2  : integer;
     FResSum2     : double;

     procedure SaveRes;
     procedure GetCellValue(Cells: TXLSCellMMU; ACol, ARow: integer; AFormulaVal: PXLSFormulaValue);
public
     constructor Create(AFvDb,AFvCriteria: TXLSFormulaValue);
     destructor Destroy; override;

     function  Add(AIndex: integer): TXLSDbField;
     function  Find(AName: AxUCString): TXLSDbField;

     function  BuildDb: TXc12CellError;
     function  ApplyCondition(AValues: TXLSArrayItem): boolean;
     procedure SearchDb(AFuncId: integer; ASrcField: AxUCString; out AResult: TXLSFormulaValue);

     property Conditions: TXLSDbConditionOr read FConditions;
     property Items[Index: integer]: TXLSDbField read GetItems; default;
     end;

procedure FVSetError(AFV: PXLSFormulaValue; AError: TXc12CellError); {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}

function ConditionStrToVarValue(ACondStr: AxUCString; out AOperator: TXLSDbCondOperator; AValue: PXLSVarValue): boolean;
function CompareVarValue(AValue1,AValue2: PXLSVarValue; AOperator: TXLSDbCondOperator): boolean;
function CompareFVValue(AValue1,AValue2: PXLSFormulaValue; AOperator: TXLSDbCondOperator): boolean;

function FVToVarValue(AFV: PXLSFormulaValue; AVarVal: PXLSVarValue): boolean;

implementation

function  FVToVarValue(AFV: PXLSFormulaValue; AVarVal: PXLSVarValue): boolean;
begin
  Result := (AFV.RefType in [xfrtNone,xfrtRef]) and not (AFV.ValType in [xfvtUnknown,xfvtAreaList]);
  if not Result then begin
    AVarVal.ValType := xfvtError;
    AVarVal.vErr := errValue;
    Exit;
  end;

  AVarVal.ValType := AFV.ValType;
  case AVarVal.ValType of
    xfvtFloat   : AVarVal.vFloat := AFV.vFloat;
    xfvtBoolean : AVarVal.vBool := AFV.vBool;
    xfvtError   : AVarVal.vErr := AFV.vErr;
    xfvtString  : AVarVal.vStr := AFV.vStr;
  end;
end;

function CompareFVValue(AValue1,AValue2: PXLSFormulaValue; AOperator: TXLSDbCondOperator): boolean;
var
  V1,V2: TXLSVarValue;
begin
  Result := FVToVarValue(AValue1,@V1) and FVToVarValue(AValue2,@V2);
  if Result then
    Result := CompareVarValue(@V1,@V2,AOperator);
end;

function CompareVarValue(AValue1,AValue2: PXLSVarValue; AOperator: TXLSDbCondOperator): boolean;
begin
  Result := AValue1.ValType = AValue2.ValType;
  if not Result then
    Exit;

  case AValue2.ValType of
    xfvtFloat  : begin
      case AOperator of
        xdcoLT   : Result := AValue1.vFloat < AValue2.vFloat;
        xdcoLE   : Result := AValue1.vFloat <= AValue2.vFloat;
        xdcoEQ   : Result := AValue1.vFloat = AValue2.vFloat;
        xdcoGE   : Result := AValue1.vFloat >= AValue2.vFloat;
        xdcoGT   : Result := AValue1.vFloat > AValue2.vFloat;
        xdcoNE   : Result := AValue1.vFloat <> AValue2.vFloat;
      end;
    end;
    xfvtBoolean: begin
      case AOperator of
        xdcoLT   : Result := AValue1.vBool < AValue2.vBool;
        xdcoLE   : Result := AValue1.vBool <= AValue2.vBool;
        xdcoEQ   : Result := AValue1.vBool = AValue2.vBool;
        xdcoGE   : Result := AValue1.vBool >= AValue2.vBool;
        xdcoGT   : Result := AValue1.vBool > AValue2.vBool;
        xdcoNE   : Result := AValue1.vBool <> AValue2.vBool;
      end;
    end;
    xfvtError  : begin
      case AOperator of
        xdcoLT   : Result := AValue1.vErr < AValue2.vErr;
        xdcoLE   : Result := AValue1.vErr <= AValue2.vErr;
        xdcoEQ   : Result := AValue1.vErr = AValue2.vErr;
        xdcoGE   : Result := AValue1.vErr >= AValue2.vErr;
        xdcoGT   : Result := AValue1.vErr > AValue2.vErr;
        xdcoNE   : Result := AValue1.vErr <> AValue2.vErr;
      end;
    end;
    xfvtString : begin
      case AOperator of
        xdcoLT   : Result := CompareText(AValue1.vStr,AValue2.vStr) < 0;
        xdcoLE   : Result := CompareText(AValue1.vStr,AValue2.vStr) <= 0;
{$ifdef DELPHI_6_OR_LATER}
        xdcoEQ   : Result := MatchesMask(AValue1.vStr,AValue2.vStr);
{$endif}        
//        xdcoEQ   : Result := CompareText(AValue1.vStr,AValue2.vStr) = 0;
        xdcoGE   : Result := CompareText(AValue1.vStr,AValue2.vStr) >= 0;
        xdcoGT   : Result := CompareText(AValue1.vStr,AValue2.vStr) > 0;
        xdcoNE   : Result := CompareText(AValue1.vStr,AValue2.vStr)<> 0;
      end;
    end;
  end;
end;

procedure FVSetError(AFV: PXLSFormulaValue; AError: TXc12CellError);
begin
  AFV.ValType := xfvtError;
  AFV.vErr := AError;
end;

procedure StrToVarValue(AStrVal: AxUCString; AVarVal: PXLSVarValue);
var
  i     : TXc12CellError;
  vFloat: double;
  vDT   : TDateTime;
begin
  if TryStrToFloat(AStrVal,vFloat) then begin
    AVarVal.vFloat := vFloat;
    AVarVal.ValType := xfvtFloat;
  end
  else if TryStrToDateTime(AStrVal,vDT) then begin
    AVarVal.vFloat := vDT;
    AVarVal.ValType := xfvtFloat;
  end
  else if AStrVal = G_StrTRUE then begin
    AVarVal.vBool := True;
    AVarVal.ValType := xfvtBoolean;
  end
  else if AStrVal = G_StrFALSE then begin
    AVarVal.vBool := False;
    AVarVal.ValType := xfvtBoolean;
  end
  else if (AStrVal <> '') and (AStrVal[1] = '#') then begin
    for i := Succ(Low(TXc12CellError)) to High(TXc12CellError) do begin
      if Uppercase(AStrVal) = Xc12CellErrorNames[i] then begin
        AVarVal.vErr := i;
        AVarVal.ValType := xfvtError;
        Break;
      end;
    end;
  end
  else begin
    AVarVal.vStr := AStrVal;
    AVarVal.ValType := xfvtString;
  end;
end;

function ConditionStrToVarValue(ACondStr: AxUCString; out AOperator: TXLSDbCondOperator; AValue: PXLSVarValue): boolean;
var
  S: AxUCString;
begin
  Result := False;
  if Length(ACondStr) < 2 then begin
    StrToVarValue(ACondStr,AValue);
    Result := True;
    Exit;
  end;

  case ACondStr[1] of
    '<': begin
      if ACondStr[2] = '=' then
        AOperator := xdcoLE
      else if ACondStr[2] = '>' then
        AOperator := xdcoNE
      else
        AOperator := xdcoLT;
    end;
    '=': AOperator := xdcoEQ;
    '>': begin
      if ACondStr[2] = '=' then
        AOperator := xdcoGE
      else
        AOperator := xdcoGT;
    end;
    else
      Exit;
  end;
  if AOperator in [xdcoLE,xdcoGE,xdcoNE] then
    S := Copy(ACondStr,3,MAXINT)
  else
    S := Copy(ACondStr,2,MAXINT);
  StrToVarValue(S,AValue);
  Result := True;
end;

{ TXLSArrayItem }

function TXLSArrayItem.Add(const ACol,ARow: integer; const AValue: TXLSFormulaValue): boolean;
begin
  Result := True;
  case AValue.ValType of
    xfvtFloat  : begin
      FValues[ACol,ARow].ValType := xfvtFloat;
      FValues[ACol,ARow].vFloat := AValue.vFloat;
    end;
    xfvtBoolean: begin
      FValues[ACol,ARow].ValType := xfvtBoolean;
      FValues[ACol,ARow].vBool := AValue.vBool;
    end;
    xfvtError  : begin
      FValues[ACol,ARow].ValType := xfvtError;
      FValues[ACol,ARow].vErr := AValue.vErr;
    end;
    xfvtString : begin
      FValues[ACol,ARow].ValType := xfvtString;
      FValues[ACol,ARow].vStr := AValue.vStr;
    end
    else begin
      Result := False;
    end;
  end;
  FValues[ACol,ARow].Empty := False;
end;

procedure TXLSArrayItem.Add(const ACol, ARow: integer; const AValue: double);
begin
  FValues[ACol,ARow].ValType := xfvtFloat;
  FValues[ACol,ARow].vFloat := AValue;
  FValues[ACol,ARow].Empty := False;
end;

procedure TXLSArrayItem.Add(const ACol, ARow: integer; const AValue: AxUCString);
begin
  FValues[ACol,ARow].ValType := xfvtString;
  FValues[ACol,ARow].vStr := AValue;
  FValues[ACol,ARow].Empty := False;
end;

procedure TXLSArrayItem.Assign(AValue: TXLSArrayItem);
var
  C,R: integer;
  H,W: integer;
begin
  H := Min(Height,AValue.Height);
  W := Min(Width,AValue.Width);

  for R := 0 to H - 1 do begin
    for C := 0 to W - 1 do
      FValues[C,R] := AValue.FValues[C,R];
  end;
end;

procedure TXLSArrayItem.CheckVectorMode(var ACol, ARow: integer);
begin
  if FVectorMode and IsVector then begin
    if ACol >= Width then
      ACol := 0;
    if ARow >= Height then
      ARow := 0;
  end;
end;

procedure TXLSArrayItem.Clear;
var
  i,j: integer;
begin
  for i := 0 to High(FValues) do begin
    for j := 0 to High(FValues[0]) do
      FValues[i,j].Empty := True;
  end;
end;

constructor TXLSArrayItem.Create(AWidth,AHeight: integer);
begin
  SetLength(FValues,AWidth,AHeight);
end;

constructor TXLSArrayItem.CreateFV(AValue: TXLSFormulaValue);
begin
  case AValue.RefType of
    xfrtNone,
    xfrtRef
                : Create(1,1);
    xfrtArea,
    xfrtXArea   : Create(AValue.Col2 - AValue.Col1 + 1,AValue.Row2 - AValue.Row1 + 1);
    xfrtArray   : Create(TXLSArrayItem(AValue.vSource).Width,TXLSArrayItem(AValue.vSource).Height);
    else          raise XLSRWException.Create('Invalid value');
  end;
end;

procedure TXLSArrayItem.FillError;
var
  C,R: integer;
begin
  for R := 0 to Height - 1 do begin
    for C := 0 to Width - 1 do
      AsError[C,R] := errNA;
  end;
end;

function TXLSArrayItem.GetAsBoolean(ACol, ARow: integer): boolean;
var
  VV: PXLSVarValue;
begin
  CheckVectorMode(ACol,ARow);
  VV := @FValues[ACol,ARow];
  if VV.ValType = xfvtBoolean then
    Result := VV.vBool
  else
    Result := False;
end;

function TXLSArrayItem.GetAsError(ACol, ARow: integer): TXc12CellError;
var
  VV: PXLSVarValue;
begin
  CheckVectorMode(ACol,ARow);
  VV := @FValues[ACol,ARow];
  if VV.ValType = xfvtError then
    Result := VV.vErr
  else
    Result := errUnknown;
end;

function TXLSArrayItem.GetAsFloat(ACol, ARow: integer): double;
var
  VV: PXLSVarValue;
begin
  CheckVectorMode(ACol,ARow);
  VV := @FValues[ACol,ARow];
  if VV.ValType = xfvtFloat then
    Result := VV.vFloat
  else
    Result := 0;
end;

function TXLSArrayItem.GetAsFormulaValue(ACol,ARow: integer): TXLSFormulaValue;
var
  VV: PXLSVarValue;
begin
  CheckVectorMode(ACol,ARow);
  VV := @FValues[ACol,ARow];
  Result.RefType := xfrtNone;
  Result.ValType := VV.ValType;
  case VV.ValType of
    xfvtString  : Result.vStr := VV.vStr;
    xfvtFloat   : Result.vFloat := VV.vFloat;
    xfvtBoolean : Result.vBool := VV.vBool;
    xfvtError   : Result.vErr := VV.vErr;
  end;
end;

function TXLSArrayItem.GetAsString(ACol, ARow: integer): AxUCString;
var
  VV: PXLSVarValue;
begin
  CheckVectorMode(ACol,ARow);
  VV := @FValues[ACol,ARow];
  if VV.ValType = xfvtString then
    Result := VV.vStr
  else
    Result := '';
end;

function TXLSArrayItem.GetHeight: integer;
begin
  if Length(FValues) > 0 then
    Result := Length(FValues[0])
  else
    Result := 0;
end;

function TXLSArrayItem.GetHit(ACol, ARow: integer): boolean;
begin
  CheckVectorMode(ACol,ARow);
  Result := (ACol < Width) and (ARow < Height);
end;

function TXLSArrayItem.GetItems(ACol,ARow: integer): PXLSVarValue;
begin
  CheckVectorMode(ACol,ARow);
  if (ACol < 0) or (ACol > High(FValues)) then
    raise XLSRWException.Create('Index out of range');

  Result := @FValues[ACol,ARow];
end;

function TXLSArrayItem.GetWidth: integer;
begin
  Result := Length(FValues);
end;

function TXLSArrayItem.IsVector: boolean;
begin
  Result := (Width = 1) or (Height = 1);
end;

procedure TXLSArrayItem.MakeEqualSize(AValue: TXLSArrayItem);
var
  W,H: integer;
begin
  W := Max(Width,AValue.Width);
  H := Max(Height,AValue.Height);
  SetLength(FValues,W,H);
  SetLength(AValue.FValues,W,H);
end;

procedure TXLSArrayItem.SetAsBoolean(ACol, ARow: integer; const Value: boolean);
begin
  FValues[ACol,ARow].Empty := False;
  FValues[ACol,ARow].ValType := xfvtBoolean;
  FValues[ACol,ARow].vBool := Value;
end;

procedure TXLSArrayItem.SetAsError(ACol, ARow: integer; const Value: TXc12CellError);
begin
  FValues[ACol,ARow].Empty := False;
  FValues[ACol,ARow].ValType := xfvtError;
  FValues[ACol,ARow].vErr := Value;
end;

procedure TXLSArrayItem.SetAsFloat(ACol, ARow: integer; AValue: double);
begin
  FValues[ACol,ARow].Empty := False;
  FValues[ACol,ARow].ValType := xfvtFloat;
  FValues[ACol,ARow].vFloat := AValue;
end;

procedure TXLSArrayItem.SetAsString(ACol, ARow: integer; const Value: AxUCString);
begin
  FValues[ACol,ARow].Empty := False;
  FValues[ACol,ARow].ValType := xfvtString;
  FValues[ACol,ARow].vStr := Value;
end;

procedure TXLSArrayItem.Add(const ACol, ARow: integer; const AValue: boolean);
begin
  FValues[ACol,ARow].Empty := False;
  FValues[ACol,ARow].ValType := xfvtBoolean;
  FValues[ACol,ARow].vBool := AValue;
end;

procedure TXLSArrayItem.Add(const ACol, ARow: integer; const AValue: TXc12CellError);
begin
  FValues[ACol,ARow].Empty := False;
  FValues[ACol,ARow].ValType := xfvtError;
  FValues[ACol,ARow].vErr := AValue;
end;

{ TXLSDbCondition }

destructor TXLSDbCondition.Destroy;
begin
  inherited;
end;

function TXLSDbCondition.DoCondition(AValue: PXLSVarValue): boolean;
begin
  Result := CompareVarValue(AValue,@FVAlue,FOp);
end;

function TXLSDbCondition.GetValue: PXLSVarValue;
begin
  Result := @FValue;
end;

{ TXLSDbConditions }

function TXLSDbConditionAnd.Add(AField: TXLSDbField; const ACondStr: AxUCString): TXLSDbCondition;
begin
  Result := TXLSDbCondition.Create;
  Result.Op := xdcoError;
  inherited Add(Result);

  if not ConditionStrToVarValue(ACondStr,Result.FOp,@Result.FValue) then
    Exit;
  Result.FField := AField;
end;

constructor TXLSDbConditionAnd.Create;
begin
  inherited Create;
end;

function TXLSDbConditionAnd.GetItems(Index: integer): TXLSDbCondition;
begin
  Result := TXLSDbCondition(inherited Items[Index]);
end;

{ TXLSDbField }

destructor TXLSDbField.Destroy;
begin
  inherited;
end;

{ TXLSDbFields }

function TXLSFmlaDatabase.Add(AIndex: integer): TXLSDbField;
begin
  if Items[AIndex] <> Nil then
    raise XLSRWException.Create('Db field is not empty');
  Result := TXLSDbField.Create;
  Result.FIndex := AIndex;
  inherited Items[AIndex] := Result;
end;

function TXLSFmlaDatabase.ApplyCondition(AValues: TXLSArrayItem): boolean;
var
  C,R: integer;
  Condition: TXLSDbCondition;
begin
  Result := False;
  for R := 0 to FConditions.Count - 1 do begin
    Result := True;
    for C := 0 to FConditions[R].Count - 1 do begin
      Condition := FConditions[R].Items[C];
      Result := Condition.DoCondition(AValues[Condition.Field.Index,0]);
      if not Result then
        Break;
    end;
    if Result then
      Exit;
  end;
end;

function TXLSFmlaDatabase.BuildDb: TXc12CellError;
var
  i         : integer;
  S         : AxUCString;
  n         : integer;
  C,R       : integer;
  Col,Row   : integer;
  CrFields  : array of TXLSDbField;
begin
  Result := errUnknown;
  for Col := FFvDb.Col1 to FFvDb.Col2 do begin
    if FFvDb.Cells.AsString(Col,FFvDb.Row1,S) then
      Add(Col - FFvDb.Col1).Name := S
    else
      Add(Col - FFvDb.Col1).Name := '';
  end;

  C := 0;
  SetLength(CrFields,FFvCriteria.Col2 - FFvCriteria.Col1 + 1);
  for Col := FFvCriteria.Col1 to FFvCriteria.Col2 do begin
    if FFvCriteria.Cells.AsString(Col,FFvCriteria.Row1,S) then begin
      CrFields[C] := Find(S);
    end;
    Inc(C);
  end;

  n := 0;
  R := 0;
  for Row := FFvCriteria.Row1 + 1 to FFvCriteria.Row2 do begin
    C := 0;
    for Col := FFvCriteria.Col1 to FFvCriteria.Col2 do begin
      if FFvCriteria.Cells.AsString(Col,Row,S) and (CrFields[C] <> Nil) then begin
        FConditions[R].Add(CrFields[C],S);
        if (CrFields[C] <> Nil) and not CrFields[C].Used then begin
          CrFields[C].Used := True;
          Inc(n);
        end;
      end;
      Inc(C);
    end;
    Inc(R);
  end;

  SetLength(FFieldXLate,n);
  n := 0;
  for i := 0 to Count - 1 do begin
    if Items[i].Used then begin
      FFieldXLate[n] := i;
      Inc(n);
    end;
  end;
end;

constructor TXLSFmlaDatabase.Create(AFvDb,AFvCriteria: TXLSFormulaValue);
var
  C,R: integer;
begin
  inherited Create;
  FFvDb := AFvDb;
  FFvCriteria := AFvCriteria;

  FConditions := TXLSDbConditionOr.Create;

  for C := FFvDb.Col1 to FFvDb.Col2 do
    inherited Add(Nil);

  for R := FFvCriteria.Row1 + 1 to FFvCriteria.Row2 do
    FConditions.Add;
end;

destructor TXLSFmlaDatabase.Destroy;
begin
  FConditions.Free;
  inherited;
end;

function TXLSFmlaDatabase.Find(AName: AxUCString): TXLSDbField;
var
  i: integer;
begin
  AName := Uppercase(AName);
  for i := 0 to Count - 1 do begin
    if Uppercase(Items[i].Name) = AName then begin
      Result := Items[i];
      Exit;
    end;
  end;
  Result := Nil;
end;

procedure TXLSFmlaDatabase.GetCellValue(Cells: TXLSCellMMU; ACol, ARow: integer; AFormulaVal: PXLSFormulaValue);
var
  Cell: TXLSCellItem;
begin
  AFormulaVal.Empty := not Cells.FindCell(ACol,ARow,Cell);
  if AFormulaVal.Empty then begin
//    AFormulaVal.ValType := xfvtUnknown
    // Is this correct? Is default always float/zero?
    AFormulaVal.ValType := xfvtFloat;
    AFormulaVal.vFloat := 0;
  end
  else begin
    AFormulaVal.RefType := xfrtNone;
    case Cells.CellType(@Cell) of
      xctNone,
      xctBlank         : AFormulaVal.ValType := xfvtUnknown;
      xctError         : begin
        AFormulaVal.ValType := xfvtError;
        AFormulaVal.vErr := Cells.GetError(@Cell);
      end;
      xctString        : begin
        AFormulaVal.ValType := xfvtString;
        AFormulaVal.vStr := Cells.GetString(@Cell);
      end;
      xctFloat         : begin
        AFormulaVal.ValType := xfvtFloat;
        AFormulaVal.vFloat := Cells.GetFloat(@Cell);
      end;
//      xctCurrency      : begin
//        AFormulaVal.ValType := xfvtFloat;
//        AFormulaVal.vFloat := Cells.GetFloat(@Cell);
//      end;
      xctBoolean       : begin
        AFormulaVal.ValType := xfvtBoolean;
        AFormulaVal.vBool := Cells.GetBoolean(@Cell);
      end;
      xctFloatFormula  : begin
        AFormulaVal.ValType := xfvtFloat;
        AFormulaVal.vFloat := Cells.GetFormulaValFloat(@Cell);
      end;
      xctStringFormula : begin
        AFormulaVal.ValType := xfvtString;
        AFormulaVal.vStr := Cells.GetFormulaValString(@Cell);
      end;
      xctBooleanFormula: begin
        AFormulaVal.ValType := xfvtBoolean;
        AFormulaVal.vBool := Cells.GetFormulaValBoolean(@Cell);
      end;
      xctErrorFormula  : begin
        AFormulaVal.ValType := xfvtError;
        AFormulaVal.vErr := Cells.GetFormulaValError(@Cell);
      end;
    end;
  end;
end;

function TXLSFmlaDatabase.GetItems(Index: integer): TXLSDbField;
begin
  Result := TXLSDbField(inherited Items[Index]);
end;

procedure TXLSFmlaDatabase.SaveRes;
begin
  FResCount2 := FResCount;
  FResCountA2 := FResCountA;
  FResSum2 := FResSum;
end;

procedure TXLSFmlaDatabase.SearchDb(AFuncId: integer; ASrcField: AxUCString; out AResult: TXLSFormulaValue);
var
  i         : integer;
  Row       : integer;
  Added     : integer;
  FV        : TXLSFormulaValue;
  Field     : TXLSDbField;
  fRes      : double;
  fTotVariance: double;
  fMin,fMax : double;
  fvRes     : TXLSFormulaValue;
  RowValues : TXLSArrayItem;
begin
  AResult.RefType := xfrtNone;
  AResult.ValType := xfvtError;

  if TryStrToInt(ASrcField,i) then begin
    if (i < 0) or (i > Count) then begin
      AResult.vErr := errValue;
      Exit;
    end;
    Dec(i);
  end
  else begin
    Field := Find(ASrcField);
    if Field = Nil then begin
      AResult.vErr := errValue;
      Exit;
    end;
    i := Field.Index;
  end;
  FSrcField := FFvDb.Col1 + i;

  FResCount := 0;
  FResCountA := 0;
  FResSum := 0;
  fRes := 0;
  fMax := MINDOUBLE;
  fMin := MAXDOUBLE;
  fTotVariance := 0;

  case AFuncId of
    045,       // DSTDEV
    047,       // DVAR
    195,       // DSTDEVP
    196: begin // DVARP
      SearchDb(000,ASrcField,FV);
      if FV.ValType <> xfvtFloat then
        Exit;
      if FResCount > 0 then
        SaveRes
      else begin
        FVSetError(@AResult,errDiv0);
        Exit;
      end;
    end;
    189: fRes := 1;
  end;

  RowValues := TXLSArrayItem.Create(FFvDb.Col2 - FFvDb.Col1 + 1,1);
  try
    for Row := FFvDb.Row1 + 1 to FFvDb.Row2 do begin
      Added := 0;
      RowValues.Clear;
      for i := 0 to High(FFieldXLate) do begin
        GetCellValue(FFvDb.Cells,FFvDb.Col1 + FFieldXLate[i],Row,@FV);
        if not Fv.Empty and (FV.ValType <> xfvtUnknown) then begin
          RowValues.Add(FFieldXLate[i],0,FV);
          Inc(Added);
        end;
      end;
      if Added <= 0 then
        Break;
      if ApplyCondition(RowValues) then begin
        GetCellValue(FFvDb.Cells,FSrcField,Row,@FV);
        if not Fv.Empty and (FV.ValType <> xfvtUnknown) then begin
          Inc(FResCountA);
          if FV.ValType = xfvtFloat then begin
            Inc(FResCount);
            FResSum := FResSum + FV.vFloat;
          end;
          case AFuncId of
            040: begin  // DCOUNT
            end;
            041: begin  // DSUM
            end;
            042: begin  // DAVERAGE
            end;
            043: begin  // DMIN
              if (FV.ValType = xfvtFloat) and (FV.vFloat < fMin) then
                fMin := FV.vFloat;
            end;
            044: begin  // DMAX
              if (FV.ValType = xfvtFloat) and (FV.vFloat > fMax) then
                fMax := FV.vFloat;
            end;
            045: begin  // DSTDEV
              if FV.ValType = xfvtFloat then
                fRes := fRes + Sqr(FV.vFloat - (FResSum2 / FResCount2));
            end;
            047: begin // DVAR
              if FV.ValType = xfvtFloat then
                fTotVariance := fTotVariance + Sqr((FResSum2 / FResCount2) - FV.vFloat);
            end;
            189: begin // DPRODUCT
              if FV.ValType = xfvtFloat then
                fRes := fRes * FV.vFloat;
            end;
            195: begin // DSTDEVP
              if FV.ValType = xfvtFloat then
                fTotVariance := fTotVariance + Sqr((FResSum2 / FResCount2) - FV.vFloat);
            end;
            196: begin // DVARP
              if FV.ValType = xfvtFloat then
                fTotVariance := fTotVariance + Sqr((FResSum2 / FResCount2) - FV.vFloat);
            end;
            199: begin // DCOUNTA
            end;
            235: begin // DGET
              fvRes := FV;
            end;
          end;
        end;
      end;
    end;
  finally
    RowValues.Free;
  end;

  AResult.ValType := xfvtFloat;
  case AFuncId of
    040: begin  // DCOUNT
      AResult.vFloat := FResCount;
    end;
    041: begin  // DSUM
      AResult.vFloat := FResSum;
    end;
    042: begin  // DAVERAGE
      if FResCount > 0 then
        AResult.vFloat := FResSum / FResCount
      else
        AResult.vFloat := 0;
    end;
    043: begin  // DMIN
      if fMin <> MAXDOUBLE then
        AResult.vFloat := fMin
      else
        AResult.vFloat := 0;
    end;
    044: begin  // DMAX
      if fMax <> MINDOUBLE then
        AResult.vFloat := fMax
      else
        AResult.vFloat := 0;
    end;
    045: begin  // DSTDEV
      if FResCount > 1 then
        AResult.vFloat := Sqrt((1 / (FResCount - 1)) * fRes)
      else
        FVSetError(@AResult,errDiv0);
    end;
    047: begin // DVAR
      if FResCount >= 0 then
        AResult.vFloat := fTotVariance / (FResCount - 1)
      else
        FVSetError(@AResult,errDiv0);
    end;
    189: begin // DPRODUCT
      AResult.vFloat := fRes;
    end;
    195: begin // DSTDEVP
      if FResCount > 0 then
        AResult.vFloat := Sqrt(fTotVariance / FResCount)
      else
        FVSetError(@AResult,errDiv0);
    end;
    196: begin // DVARP
      if FResCount > 0 then
        AResult.vFloat := fTotVariance / FResCount
      else
        FVSetError(@AResult,errDiv0);
    end;
    199: begin // DCOUNTA
      AResult.vFloat := FResCountA;
    end;
    235: begin // DGET
      if FResCountA = 1 then
        AResult := fvRes
      else
        FVSetError(@AResult,errNum);
    end;
  end;
end;

{ TXLSDbConditionOr }

function TXLSDbConditionOr.Add: TXLSDbConditionAnd;
begin
  Result := TXLSDbConditionAnd.Create;
  inherited Add(Result);
end;

constructor TXLSDbConditionOr.Create;
begin
  inherited Create;
end;

function TXLSDbConditionOr.GetItems(Index: integer): TXLSDbConditionAnd;
begin
  Result := TXLSDbConditionAnd(inherited Items[Index]);
end;


{ TXLSCellsSource }

constructor TXLSCellsSource.Create(AWidth,AHeight: integer);
begin
  FCol1 := 0;
  FRow1 := 0;
  FCol2 := AWidth - 1;
  FRow2 := AHeight - 1;
  FArray := TXLSArrayItem.Create(AWidth,AHeight);
  FOwnsArray := True;
end;

constructor TXLSCellsSource.Create(AArray: TXLSArrayItem);
begin
  FArray := AArray;
  FCol1 := 0;
  FRow1 := 0;
  FCol2 := FArray.Width - 1;
  FRow2 := FArray.Height - 1;
end;

procedure TXLSCellsSource.BeginIterate;
begin
  FIterCol := FCol1 - 1;
  FIterRow := FRow1;
end;

constructor TXLSCellsSource.Create(ACells: TXLSCellMMU; ACol1, ARow1, ACol2, ARow2: integer);
begin
  FCells := ACells;
  FCol1 := ACol1;
  FRow1 := ARow1;
  FCol2 := ACol2;
  FRow2 := ARow2;
end;

destructor TXLSCellsSource.Destroy;
begin
  if FOwnsArray then
    FArray.Free;
  inherited;
end;

function TXLSCellsSource.GetAsFloat(ACol, ARow: integer): double;
var
  VV: PXLSVarValue;
begin
  if FCells <> Nil then
    FError := FCells.AsFloat(ACol,Arow,Result,FEmpty)
  else begin
    Result := 0;
    VV := FArray[ACol,ARow];
    case VV.ValType of
      xfvtFloat  : Result := VV.vFloat;
      xfvtBoolean: FEmpty := True;
      xfvtError  : FError := VV.vErr;
      xfvtString : FEmpty := True;
    end;
    FEmpty := VV.Empty;
  end;
end;

function TXLSCellsSource.GetHeight: integer;
begin
  Result := FRow2 - FRow1 + 1;
end;

function TXLSCellsSource.GetWidth: integer;
begin
  Result := FCol2 - FCol1 + 1;
end;

function TXLSCellsSource.IterAsFloat: double;
begin
  Result := GetAsFloat(FIterCol,FIterRow);
end;

function TXLSCellsSource.IterateNext: boolean;
begin
  Inc(FIterCol);
  if FIterCol > FCol2 then begin
    FIterCol := FCol1;
    Inc(FIterRow);
    Result := FIterRow <= FRow2;
  end
  else
    Result := True;
end;

procedure TXLSCellsSource.SetAsFloat(ACol, ARow: integer; const Value: double);
begin
  if FCells <> Nil then
    FCells.UpdateFloat(ACol,ARow,Value)
  else
    FArray.SetAsFloat(ACol,ARow,Value);
end;

{ TXLSArrayItemIterate }

procedure TXLSArrayItemIterate.BeginIterate(ALoopOnVectors: boolean = True);
begin
  FLoopOnVectors := ALoopOnVectors and IsVector;
  FCol := -1;
  FRow := 0;
end;

function TXLSArrayItemIterate.IterateNext: boolean;
begin
  Result := True;
  Inc(FCol);
  if FCol >= Width then begin
    Inc(FRow);
    if FRow >= Height then begin
      if FLoopOnVectors then begin
        FCol := -1;
        FRow := 0;
      end
      else
        Result := False;
    end
  end;
end;

function TXLSArrayItemIterate.IterError: TXc12CellError;
begin
  Result := GetAsError(FCol,FRow);
end;

function TXLSArrayItemIterate.IterFloat: double;
begin
  Result := GetAsFloat(FCol,FRow);
end;

end.
