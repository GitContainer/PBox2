unit XLSMatrix5;

interface

uses Classes, SysUtils, Math;

type TXLSMatrix = class(TObject)
private
     function  GetItems(ACol, ARow: integer): double;
     procedure SetItems(ACol, ARow: integer; const Value: double);
protected
     FCols,FRows: integer;
     FMatrix: array of array of double;

     procedure Swap(C1,R1,C2,R2: integer);
public
     constructor Create(AMatrix: TXLSMatrix); overload;
     constructor Create(ACols,ARows: integer); overload;

     function  Determ(out ARes: double): boolean;
     function  Minor(ACol,ARow: integer): TXLSMatrix;
     function  Inverse: TXLSMatrix;
     procedure Mult(AValue: double); overload;
     function  Mult(AValue: TXLSMatrix): TXLSMatrix; overload;
     procedure Zero;
     procedure Diag;

     property Items[ACol,ARow: integer]: double read GetItems write SetItems; default;
     property Cols: integer read FCols;
     property Rows: integer read FRows;
     end;

implementation


{ TXLSMatrix }

constructor TXLSMatrix.Create(ACols, ARows: integer);
begin
  FCols := ACols;
  FRows := ARows;
  SetLength(FMatrix,ACols,ARows);
end;

constructor TXLSMatrix.Create(AMatrix: TXLSMatrix);
var
  C,R: integer;
begin
  Create(AMatrix.FCols,AMatrix.FRows);
  for C := 0 to FCols - 1 do begin
    for R := 0 to FRows - 1 do
      Items[C,R] := AMatrix[C,R];
  end;
end;

function TXLSMatrix.Determ(out ARes: double): boolean;
var
  C: integer;
  M: TXLSMatrix;
  V: double;
begin
  Result := FRows = FCols;
  if not Result then
    Exit;

  if FRows = 1 then
    ARes := GetItems(0,0)
  else if FRows = 2 then
    ARes := GetItems(0,0) * GetItems(1,1) - GetItems(0,1) * GetItems(1,0)
  else begin
    ARes := 0;
    for C := 0 to FCols - 1 do begin
      M := Minor(C,0);
      try
        Result := M.Determ(V);
        if not Result then
          Exit;
        ARes := ARes + (((C + 1) mod 2 + (C + 1) mod 2 - 1) * Items[C,0] * V);
      finally
        M.Free;
      end;
    end;
  end;
end;

procedure TXLSMatrix.Diag;
var
  i: integer;
begin
  Zero;
  for i := 0 to Min(FCols,FRows) - 1 do
    Items[i,i] := 1;
end;

function TXLSMatrix.GetItems(ACol, ARow: integer): double;
begin
  Result := FMatrix[ACol,ARow];
end;

function TXLSMatrix.Inverse: TXLSMatrix;
var
  D: double;
  AI: TXLSMatrix;
  C,R: integer;
  S: integer;
  F: double;
begin
  if not Determ(D) or (D = 0) or (FCols <> FRows) then begin
    Result := Nil;
    Exit;
  end;
  Result := TXLSMatrix.Create(FCols,FRows);
  if FRows = 1 then
    Result[0,0] := 1 / Items[0,0]
  else if FRows = 2 then begin
    Result[0,0] := Items[1,1];
    Result[1,0] := -Items[1,0];
    Result[0,1] := -Items[0,1];
    Result[1,1] :=  Items[0,0];
    Result.Mult(1 / d);
  end
  else begin
    Result.Diag;
    AI := TXLSMatrix.Create(Self);
    try
      for C := 0 to FCols - 1 do begin
        R := C;
        while (R < FRows) and (AI[C,R] = 0) do
          Inc(R);
        if R <> C then begin
          for S := 0 to FCols - 1 do begin
            AI.Swap(S,C,S,R);
            Result.Swap(S,C,S,R);
          end;
        end;
        for R := 0 to FRows - 1 do begin
          if R <> C then begin
            if AI[C,R] <> 0 then begin
              F := -AI[C,R] / AI[C,C];
              for S := 0 to FCols - 1 do begin
                AI[S,R] := AI[S,R] + F * AI[S,C];
                Result[S,R] := Result[S,R] + F * Result[S,C];
              end;
            end;
          end
          else begin
            F := AI[C,C];
            for S := 0 to FCols - 1 do begin
              AI[S,R] := AI[S,R] / F;
              Result[S,R] := Result[S,R] / F;
            end;
          end;
        end;
      end;
    finally
      AI.Free;
    end;
  end;
end;

function TXLSMatrix.Minor(ACol, ARow: integer): TXLSMatrix;
var
  C,R: integer;
  i,j: integer;
begin
  if (ARow >= 0) and (ARow < FRows) and (ACol >= 0) and (ACol < FCols) then begin
    Result := TXLSMatrix.Create(FCols - 1,FRows - 1);
    i := 0;
    for R := 0 to FRows - 1 do begin
      if R = ARow then
        Continue;
      j := 0;
      for C := 0 to FCols - 1 do begin
        if C = ACol then
          Continue;
        Result[j,i] := Items[C,R];
        Inc(j);
      end;
      Inc(i);
    end;
  end
  else
    Result := Nil;
end;

function TXLSMatrix.Mult(AValue: TXLSMatrix): TXLSMatrix;
var
  C,R: integer;
  CRes: integer;
begin
  if AValue.Cols = FRows then begin
    Result := TXLSMatrix.Create(FCols,AValue.Rows);
    Result.Zero;
    for R := 0 to AValue.Rows - 1 do begin
      for CRes := 0 to FCols - 1 do begin
        for C := 0 to AValue.Cols - 1 do
          Result[CRes,R] := Result[CRes,R] + AValue[C,R] * Items[CRes,C];
      end;
    end;
  end
  else
    Result := Nil;
end;

procedure TXLSMatrix.Mult(AValue: double);
var
  C,R: integer;
begin
  for C := 0 to FCols - 1 do begin
    for R := 0 to FRows - 1 do
      Items[C,R] := Items[C,R] * AValue;
  end;
end;

procedure TXLSMatrix.SetItems(ACol, ARow: integer; const Value: double);
begin
  FMatrix[ACol,ARow] := Value;
end;

procedure TXLSMatrix.Swap(C1, R1, C2, R2: integer);
var
  T: double;
begin
  T := Items[C1,R1];
  Items[C1,R1] := Items[C2,R2];
  Items[C2,R2] := T;
end;

procedure TXLSMatrix.Zero;
var
  C,R: integer;
begin
  for C := 0 to FCols - 1 do begin
    for R := 0 to FRows - 1 do
      Items[C,R] := 0;
  end;
end;

end.
