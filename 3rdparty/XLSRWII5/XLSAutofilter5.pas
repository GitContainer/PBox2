unit XLSAutofilter5;

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
     Xc12Utils5, Xc12DataAutofilter5, Xc12DataWorkbook5, Xc12DataWorksheet5, Xc12Manager5,
     XLSUtils5, XLSClassFactory5, XLSNames5;

type TXLSAutofilterColumn = class(TXc12FilterColumn)
protected
     function  TestCondition(AVal1, AVal2: double; AOp: TXc12FilterOperator): boolean; overload;
     function  TestCondition(AVal1, AVal2: AxUCString; AOp: TXc12FilterOperator): boolean; overload;
     function  TestCondition(AVal1, AVal2: boolean; AOp: TXc12FilterOperator): boolean; overload;
public
     procedure Add(AValue: double; AOperator: TXc12FilterOperator); overload;
     procedure Add(AValue1,AValue2: double; AOperator1,AOperator2: TXc12FilterOperator); overload;
     procedure Add(AValue1,AValue2: AxUCString; AOperator1,AOperator2: TXc12FilterOperator); overload;
     procedure Add(AValue: AxUCString; AOperator: TXc12FilterOperator); overload;

     function  Apply(AValue: double): boolean; overload;
     function  Apply(AValue: AxUCString): boolean; overload;
     function  Apply(AValue: boolean): boolean; overload;
     end;

type TXLSAutofilter = class(TXc12AutoFilter)
protected
     FNames    : TXLSNames;
     FXc12Sheet: TXc12DataWorksheet;

     function GetCol1: integer;
     function GetCol2: integer;
     function GetRow1: integer;
     function GetRow2: integer;
public
     constructor Create(AClassFactory: TXLSClassFactory; AClassOwner: TObject; ANames: TXLSNames);

     procedure Clear;

     function Assigned: boolean;

     function IsFiltered: boolean;

     procedure Add(const AC1,AR1,AC2,AR2: integer);

     procedure Move(Col1, Row1, Col2, Row2, DeltaCol, DeltaRow: integer);

     property Col1: integer read GetCol1;
     property Row1: integer read GetRow1;
     property Col2: integer read GetCol2;
     property Row2: integer read GetRow2;
     end;

implementation

{ TXLSAutofilter }

procedure TXLSAutofilter.Add(const AC1,AR1,AC2,AR2: integer);
var
  N: TXLSName;
begin
  Clear;
  FRef := SetCellArea(AC1,AR1,AC2,AR2);

  N := TXLSName(FNames.FindBuiltIn(bnFilterDatabase,FXc12Sheet.Index));
  if N = Nil then
    N := TXLSName(FNames.Add(bnFilterDatabase,FXc12Sheet.Index));
  N.Definition := '''' + FXc12Sheet.Name + '''!' + AreaToRefStr(AC1,AR1,AC2,AR2,True,True,True,True);
  N.Hidden := True;
end;

function TXLSAutofilter.Assigned: boolean;
begin
  Result := FNames.FindBuiltIn(bnFilterDatabase,FXc12Sheet.Index) <> Nil;
end;

procedure TXLSAutofilter.Clear;
begin
  inherited Clear;

  FNames.Delete(bnFilterDatabase,FXc12Sheet.Index);
end;

constructor TXLSAutofilter.Create(AClassFactory: TXLSClassFactory; AClassOwner: TObject; ANames: TXLSNames);
begin
  inherited Create(AClassFactory);

  FXc12Sheet := TXc12DataWorksheet(AClassOwner);
  FNames := ANames;
end;

function TXLSAutofilter.GetCol1: integer;
begin
  Result := FRef.Col1;
end;

function TXLSAutofilter.GetCol2: integer;
begin
  Result := FRef.Col2;
end;

function TXLSAutofilter.GetRow1: integer;
begin
  Result := FRef.Row1;
end;

function TXLSAutofilter.GetRow2: integer;
begin
  Result := FRef.Row2;
end;

function TXLSAutofilter.IsFiltered: boolean;
begin
  Result := Active;
end;

procedure TXLSAutofilter.Move(Col1, Row1, Col2, Row2, DeltaCol, DeltaRow: integer);
begin

end;

{ TXLSAutofilterColumn }

procedure TXLSAutofilterColumn.Add(AValue: double; AOperator: TXc12FilterOperator);
begin
  FCustomFilters.Filter1.Val := FloatToStr(AValue);
  FCustomFilters.Filter1.Operator_ := AOperator;
end;

procedure TXLSAutofilterColumn.Add(AValue: AxUCString; AOperator: TXc12FilterOperator);
begin
  FCustomFilters.Filter1.Val := AValue;
  FCustomFilters.Filter1.Operator_ := AOperator;
end;

function TXLSAutofilterColumn.Apply(AValue: AxUCString): boolean;
var
  i: integer;
  R: boolean;
begin
  Result := (FCustomFilters.Filter1.Val <> '') or (FFilters.Filter.Count > 0);
  if not Result then
    Exit;

  if FFilters.Filter.Count > 0 then begin
    for i := 0 to FFilters.Filter.Count - 1 do begin
      Result := SameText(FFilters.Filter[i],AValue);
      if Result then
        Exit;
    end;
  end
  else begin
    Result := TestCondition(AValue,FCustomFilters.Filter1.Val,FCustomFilters.Filter1.Operator_);

    if Result and (FCustomFilters.Filter2.Val <> '') then begin
      R := TestCondition(AValue,FCustomFilters.Filter2.Val,FCustomFilters.Filter2.Operator_);

      if FCustomFilters.And_ then
        Result := Result and R;
    end;
  end;
end;

function TXLSAutofilterColumn.Apply(AValue: boolean): boolean;
var
  i: integer;
  V: boolean;
  R: boolean;
begin
  Result := ((FCustomFilters.Filter1.Val <> '') and TryStrToBool(FCustomFilters.Filter1.Val,V)) or (FFilters.Filter.Count > 0);
  if not Result then
    Exit;

  if FFilters.Filter.Count > 0 then begin
    for i := 0 to FFilters.Filter.Count - 1 do begin
      Result := TryStrToBool(FCustomFilters.Filter1.Val,V) and (V = AValue);
      if Result then
        Exit;
    end;
  end
  else begin
    Result := TestCondition(AValue,V,FCustomFilters.Filter1.Operator_);

    if Result and (FCustomFilters.Filter2.Val <> '') and TryStrToBool(FCustomFilters.Filter2.Val,V) then begin
      R := TestCondition(AValue,V,FCustomFilters.Filter2.Operator_);

      if FCustomFilters.And_ then
        Result := Result and R;
    end;
  end;
end;

function TXLSAutofilterColumn.TestCondition(AVal1, AVal2: double; AOp: TXc12FilterOperator): boolean;
begin
  Result := False;

  case AOp of
    x12foEqual             : Result := AVal1 = AVal2;
    x12foLessThan          : Result := AVal1 < AVal2;
    x12foLessThanOrEqual   : Result := AVal1 <= AVal2;
    x12foNotEqual          : Result := AVal1 <> AVal2;
    x12foGreaterThanOrEqual: Result := AVal1 >= AVal2;
    x12foGreaterThan       : Result := AVal1 > AVal2;
  end;
end;

function TXLSAutofilterColumn.TestCondition(AVal1, AVal2: AxUCString; AOp: TXc12FilterOperator): boolean;
var
  V: integer;
begin
  Result := False;

  V := CompareStr(AVal1,AVal2);
  case AOp of
    x12foEqual             : Result := V = 0;
    x12foLessThan          : Result := V < 0;
    x12foLessThanOrEqual   : Result := V >= 0;
    x12foNotEqual          : Result := V <> 0;
    x12foGreaterThanOrEqual: Result := V >= 0;
    x12foGreaterThan       : Result := V > 0;
  end;
end;

function TXLSAutofilterColumn.TestCondition(AVal1, AVal2: boolean; AOp: TXc12FilterOperator): boolean;
begin
  Result := False;

  case AOp of
    x12foEqual             : Result := AVal1 = AVal2;
    x12foLessThan          : Result := AVal1 < AVal2;
    x12foLessThanOrEqual   : Result := AVal1 <= AVal2;
    x12foNotEqual          : Result := AVal1 <> AVal2;
    x12foGreaterThanOrEqual: Result := AVal1 >= AVal2;
    x12foGreaterThan       : Result := AVal1 > AVal2;
  end;
end;

procedure TXLSAutofilterColumn.Add(AValue1, AValue2: AxUCString; AOperator1, AOperator2: TXc12FilterOperator);
begin
  FCustomFilters.Filter1.Val := AValue1;
  FCustomFilters.Filter1.Operator_ := AOperator1;

  FCustomFilters.Filter2.Val := AValue2;
  FCustomFilters.Filter2.Operator_ := AOperator2;
end;

function TXLSAutofilterColumn.Apply(AValue: double): boolean;
var
  i: integer;
  V: double;
  R: boolean;
begin
  Result := ((FCustomFilters.Filter1.Val <> '') and TryStrToFloat(FCustomFilters.Filter1.Val,V)) or (FFilters.Filter.Count > 0);
  if not Result then
    Exit;

  if FFilters.Filter.Count > 0 then begin
    for i := 0 to FFilters.Filter.Count - 1 do begin
      Result := TryStrToFloat(FFilters.Filter[i],V) and SameValue(V,AValue);
      if Result then
        Exit;
    end;
  end
  else begin
    Result := TestCondition(AValue,V,FCustomFilters.Filter1.Operator_);

    if Result and (FCustomFilters.Filter2.Val <> '') and TryStrToFloat(FCustomFilters.Filter2.Val,V) then begin
      R := TestCondition(AValue,V,FCustomFilters.Filter2.Operator_);

      if FCustomFilters.And_ then
        Result := Result and R;
    end;
  end;
end;

procedure TXLSAutofilterColumn.Add(AValue1, AValue2: double; AOperator1, AOperator2: TXc12FilterOperator);
begin
  FCustomFilters.Filter1.Val := FloatToStr(AValue1);
  FCustomFilters.Filter1.Operator_ := AOperator1;

  FCustomFilters.Filter2.Val := FloatToStr(AValue2);
  FCustomFilters.Filter2.Operator_ := AOperator2;
end;

end.
