unit XLSSharedItems5;

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
     Xc12Utils5,
     XLSUtils5;

type TXLSSharedItemsValueType = (xsivtBoolean,xsivtDate,xsivtError,xsivtBlank,xsivtNumeric,xsivtString);

     TXLSSharedItemsValue = class(TObject)
protected
     FUnused: boolean;

     function  HashOf: longword; virtual; abstract;

     function  GetAsBoolean: boolean; virtual;
     function  GetAsDate: TDateTime; virtual;
     function  GetAsError: TXc12CellError; virtual;
     function  GetAsNumeric: double; virtual;
     function  GetAsString: AxUCString; virtual;
     function  GetAsIdString: AxUCString; virtual; abstract;
     function  GetAsText: AxUCString; virtual;
     procedure SetAsBoolean(const Value: boolean); virtual;
     procedure SetAsDate(const Value: TDateTime); virtual;
     procedure SetAsError(const Value: TXc12CellError); virtual;
     procedure SetAsNumeric(const Value: double); virtual;
     procedure SetAsString(const Value: AxUCString); virtual;
public
     function  Type_: TXLSSharedItemsValueType; virtual; abstract;

     function  Clone: TXLSSharedItemsValue; virtual; abstract;

     function  Equal(AValue: boolean): boolean; overload;
     function  Equal(AValue: TDateTime {$ifndef DELPHI_XE_OR_LATER}; Dummy: integer {$endif}): boolean; overload;
     function  Equal: boolean; overload;
     function  Equal(AValue: TXc12CellError): boolean; overload;
     function  Equal(AValue: double): boolean; overload;
     function  Equal(AValue: AxUCString): boolean; overload;
     function  Equal(AValue: TXLSSharedItemsValue): boolean; overload;

     property Unused: boolean read FUnused write FUnused;

     property AsBoolean: boolean read GetAsBoolean write SetAsBoolean;
     property AsDate: TDateTime read GetAsDate write SetAsDate;
     property AsError: TXc12CellError read GetAsError write SetAsError;
     property AsNumeric: double read GetAsNumeric write SetAsNumeric;
     property AsString: AxUCString read GetAsString write SetAsString;
     property AsIdString: AxUCString read GetAsIdString;

     property AsText: AxUCString read GetAsText;
     end;

     TXLSSharedItemsValueBoolean = class(TXLSSharedItemsValue)
protected
     FValue: boolean;

     function  HashOf: longword; override;

     function  GetAsBoolean: boolean; override;
     function  GetAsIdString: AxUCString; override;
     function  GetAsText: AxUCString; override;
     procedure SetAsBoolean(const Value: boolean); override;
public
     function  Type_: TXLSSharedItemsValueType; override;

     function  Clone: TXLSSharedItemsValue; override;
     end;

     TXLSSharedItemsValueDate = class(TXLSSharedItemsValue)
protected
     FValue: TDateTime;

     function  HashOf: longword; override;

     function  GetAsDate: TDateTime; override;
     function  GetAsIdString: AxUCString; override;
     function  GetAsText: AxUCString; override;
     procedure SetAsDate(const Value: TDateTime); override;
public
     function  Type_: TXLSSharedItemsValueType; override;

     function  Clone: TXLSSharedItemsValue; override;
     end;

     TXLSSharedItemsValueError = class(TXLSSharedItemsValue)
protected
     FValue: TXc12CellError;

     function  HashOf: longword; override;

     function  GetAsIdString: AxUCString; override;
     function  GetAsText: AxUCString; override;
     function  GetAsError: TXc12CellError; override;
     procedure SetAsError(const Value: TXc12CellError); override;
public
     function  Type_: TXLSSharedItemsValueType; override;

     function  Clone: TXLSSharedItemsValue; override;
     end;

     TXLSSharedItemsValueBlank = class(TXLSSharedItemsValue)
protected
     function  HashOf: longword; override;
     function  GetAsIdString: AxUCString; override;
     function  GetAsText: AxUCString; override;
public
     function  Type_: TXLSSharedItemsValueType; override;

     function  Clone: TXLSSharedItemsValue; override;
     end;

     TXLSSharedItemsValueNumeric = class(TXLSSharedItemsValue)
protected
     FValue: double;

     function  HashOf: longword; override;

     function  GetAsNumeric: double; override;
     function  GetAsIdString: AxUCString; override;
     function  GetAsText: AxUCString; override;
     procedure SetAsNumeric(const Value: double); override;
public
     function  Type_: TXLSSharedItemsValueType; override;

     function  Clone: TXLSSharedItemsValue; override;
     end;

     TXLSSharedItemsValueString = class(TXLSSharedItemsValue)
protected
     FValue: AxUCString;

     function  HashOf: longword; override;

     function  GetAsString: AxUCString; override;
     function  GetAsIdString: AxUCString; override;
     function  GetAsText: AxUCString; override;
     procedure SetAsString(const Value: AxUCString); override;
public
     function  Type_: TXLSSharedItemsValueType; override;

     function  Clone: TXLSSharedItemsValue; override;
     end;

     TXLSSharedItemsValues = class(TObjectList)
private
     function GetItems(Index: integer): TXLSSharedItemsValue;
protected
     procedure DoSort(ALeft,ARight: integer);
public
     constructor Create(AOwnsItems: boolean = True);

     procedure Sort;

     function  Find(AValue: boolean): boolean; overload; virtual;
     function  Find(AValue: double): boolean; overload; virtual;
     function  Find(AValue: TDateTime{$ifndef DELPHI_XE_OR_LATER}; Dummy: integer {$endif}): boolean; overload; virtual;
     function  Find(AValue: TXc12CellError): boolean; overload; virtual;
     function  Find: boolean; overload; virtual;
     function  Find(AValue: AxUCString): boolean; overload; virtual;
     function  Find(AValue: TXLSSharedItemsValue): boolean; overload; virtual;

     function  FindValue(AValue: TXLSSharedItemsValue): integer;

     function Add(AValue: TXLSSharedItemsValue): boolean; overload; virtual;
     function Add(AValue: boolean): boolean; overload; virtual;
     function Add(AValue: TDateTime{$ifndef DELPHI_XE_OR_LATER}; Dummy: integer {$endif}): boolean; overload; virtual;
     function Add(AValue: TXc12CellError): boolean; overload; virtual;
     function Add: boolean; overload; virtual;
     function Add(AValue: double): boolean; overload; virtual;
     function Add(AValue: AxUCString): boolean; overload; virtual;

     function  Compare(AIndex1,AIndex2: integer): double;

     property Items[Index: integer]: TXLSSharedItemsValue read GetItems; default;
     end;

     TXLSUniqueSharedItemsValues = class(TXLSSharedItemsValues)
protected
     FBuckets     : array of TXLSSharedItemsValues;

//     function  HashOf(AValue: TDateTime): longword; overload;
     function  HashOf(AValue: TXc12CellError): longword; overload;
     function  HashOf(AValue: double): longword; overload;
     function  HashOf(AValue: AxUCString): longword; overload;

     procedure ClearBuckets;
public
     constructor Create(AOwnsItems: boolean = True; Size: Cardinal = 256);

     procedure Resize(ASize: integer);

     procedure Clear; override;

     function  Find(AValue: boolean): boolean; overload; override;
     function  Find(AValue: TDateTime{$ifndef DELPHI_XE_OR_LATER}; Dummy: integer {$endif}): boolean; overload; override;
     function  Find(AValue: TXc12CellError): boolean; overload; override;
     function  Find: boolean; overload; override;
     function  Find(AValue: double): boolean; overload; override;
     function  Find(AValue: AxUCString): boolean; overload; override;
     function  Find(AValue: TXLSSharedItemsValue): boolean; overload; override;

     function Add(AValue: TXLSSharedItemsValue): boolean; overload; override;
     function Add(AValue: boolean): boolean; overload; override;
     function Add(AValue: TDateTime{$ifndef DELPHI_XE_OR_LATER}; Dummy: integer {$endif}): boolean; overload; override;
     function Add(AValue: TXc12CellError): boolean; overload; override;
     function Add: boolean; overload; override;
     function Add(AValue: double): boolean; overload; override;
     function Add(AValue: AxUCString): boolean; overload; override;
     end;

implementation

{ TXLSSharedItemsValue }

// function TXLSSharedItemsValue.Equal(AValue: TDateTime): boolean;
// begin
//   Result := (Type_ = xsivtDate) and (AValue = AsDate);
// end;

function TXLSSharedItemsValue.Equal(AValue: boolean): boolean;
begin
  Result := (Type_ = xsivtBoolean) and (AValue = AsBoolean);
end;

function TXLSSharedItemsValue.Equal(AValue: TXc12CellError): boolean;
begin
  Result := (Type_ = xsivtError) and (AValue = AsError);
end;

function TXLSSharedItemsValue.Equal(AValue: AxUCString): boolean;
begin
  Result := (Type_ = xsivtString) and (AValue = AsString);
end;

function TXLSSharedItemsValue.Equal: boolean;
begin
  Result := Type_ = xsivtBlank;
end;

function TXLSSharedItemsValue.Equal(AValue: double): boolean;
begin
  Result := (Type_ = xsivtNumeric) and (AValue = AsNumeric);
end;

function TXLSSharedItemsValue.GetAsBoolean: boolean;
begin
  raise XLSRWException.Create('Invalid value type');
end;

function TXLSSharedItemsValue.GetAsDate: TDateTime;
begin
  raise XLSRWException.Create('Invalid value type');
end;

function TXLSSharedItemsValue.GetAsError: TXc12CellError;
begin
  raise XLSRWException.Create('Invalid value type');
end;

function TXLSSharedItemsValue.GetAsNumeric: double;
begin
  raise XLSRWException.Create('Invalid value type');
end;

function TXLSSharedItemsValue.GetAsString: AxUCString;
begin
  raise XLSRWException.Create('Invalid value type');
end;

function TXLSSharedItemsValue.GetAsText: AxUCString;
begin
  Result := '';
end;

procedure TXLSSharedItemsValue.SetAsBoolean(const Value: boolean);
begin
  raise XLSRWException.Create('Invalid value type');
end;

procedure TXLSSharedItemsValue.SetAsDate(const Value: TDateTime);
begin
  raise XLSRWException.Create('Invalid value type');
end;

procedure TXLSSharedItemsValue.SetAsError(const Value: TXc12CellError);
begin
  raise XLSRWException.Create('Invalid value type');
end;

procedure TXLSSharedItemsValue.SetAsNumeric(const Value: double);
begin
  raise XLSRWException.Create('Invalid value type');
end;

procedure TXLSSharedItemsValue.SetAsString(const Value: AxUCString);
begin
  raise XLSRWException.Create('Invalid value type');
end;

function TXLSSharedItemsValue.Equal(AValue: TXLSSharedItemsValue): boolean;
begin
  Result := Type_ = AValue.Type_;

  if Result then begin
    case Type_ of
      xsivtBoolean: Result := AsBoolean = AValue.AsBoolean;
      xsivtDate   : Result := AsDate = AValue.AsDate;
      xsivtError  : Result := AsError = AValue.AsError;
      xsivtBlank  : Result := True;
      xsivtNumeric: Result := AsNumeric = AValue.AsNumeric;
      xsivtString : Result := AsString = AValue.AsString;
    end;

  end;
end;

function TXLSSharedItemsValue.Equal(AValue: TDateTime{$ifndef DELPHI_XE_OR_LATER}; Dummy: integer {$endif}): boolean;
begin
  Result := (Type_ = xsivtDate) and (AValue = AsDate);
end;

{ TXLSSharedItemsValueBoolean }

function TXLSSharedItemsValueBoolean.GetAsIdString: AxUCString;
begin
  if FValue then
    Result := 'TRUE'
  else
    Result := 'FALSE'
end;

function TXLSSharedItemsValueBoolean.GetAsText: AxUCString;
begin
  if FValue then
    Result := 'TRUE'
  else
    Result := 'FALSE'
end;

function TXLSSharedItemsValueBoolean.HashOf: longword;
begin
  if FValue then
    Result := 1
  else
    Result := 0;
end;

function TXLSSharedItemsValueBoolean.Clone: TXLSSharedItemsValue;
begin
  Result :=  TXLSSharedItemsValueBoolean.Create;
  TXLSSharedItemsValueBoolean(Result).FValue := FValue;
end;

function TXLSSharedItemsValueBoolean.GetAsBoolean: boolean;
begin
  Result := FValue;
end;

procedure TXLSSharedItemsValueBoolean.SetAsBoolean(const Value: boolean);
begin
  FValue := Value;
end;

function TXLSSharedItemsValueBoolean.Type_: TXLSSharedItemsValueType;
begin
  Result := xsivtBoolean;
end;

{ TXLSSharedItemsValueDate }

function TXLSSharedItemsValueDate.GetAsIdString: AxUCString;
begin
  // Don't use DateTimeToStr
  Result := FloatToStr(FValue);
end;

function TXLSSharedItemsValueDate.GetAsText: AxUCString;
begin
  Result := DateToStr(FValue);
end;

function TXLSSharedItemsValueDate.HashOf: longword;
var
  i: Integer;
  P: PByteArray;
begin
  P := @FValue;

  Result := 0;
  for i := 0 to SizeOf(FValue) - 1 do
    Result := ((Result shl 2) or (Result shr (SizeOf(Result) * 8 - 2))) xor Ord(P[i]);
end;

function TXLSSharedItemsValueDate.Clone: TXLSSharedItemsValue;
begin
  Result :=  TXLSSharedItemsValueDate.Create;
  TXLSSharedItemsValueDate(Result).FValue := FValue;
end;

function TXLSSharedItemsValueDate.GetAsDate: TDateTime;
begin
  Result := FValue;
end;

procedure TXLSSharedItemsValueDate.SetAsDate(const Value: TDateTime);
begin
  FValue := Value;
end;

function TXLSSharedItemsValueDate.Type_: TXLSSharedItemsValueType;
begin
  Result := xsivtDate;
end;

{ TXLSSharedItemsValueError }

function TXLSSharedItemsValueError.GetAsIdString: AxUCString;
begin
  Result := Xc12CellErrorNames[FValue];
end;

function TXLSSharedItemsValueError.GetAsText: AxUCString;
begin
  Result := Xc12CellErrorNames[FValue];
end;

function TXLSSharedItemsValueError.HashOf: longword;
var
  i: Integer;
  P: PByteArray;
begin
  P := @FValue;

  Result := 0;
  for i := 0 to SizeOf(FValue) - 1 do
    Result := ((Result shl 2) or (Result shr (SizeOf(Result) * 8 - 2))) xor Ord(P[i]);
end;

function TXLSSharedItemsValueError.Clone: TXLSSharedItemsValue;
begin
  Result :=  TXLSSharedItemsValueError.Create;
  TXLSSharedItemsValueError(Result).FValue := FValue;
end;

function TXLSSharedItemsValueError.GetAsError: TXc12CellError;
begin
  Result := FValue;
end;

procedure TXLSSharedItemsValueError.SetAsError(const Value: TXc12CellError);
begin
  FValue := Value;
end;

function TXLSSharedItemsValueError.Type_: TXLSSharedItemsValueType;
begin
  Result := xsivtError;
end;

{ TXLSSharedItemsValueNumeric }

function TXLSSharedItemsValueNumeric.Clone: TXLSSharedItemsValue;
begin
  Result :=  TXLSSharedItemsValueNumeric.Create;
  TXLSSharedItemsValueNumeric(Result).FValue := FValue;
end;

function TXLSSharedItemsValueNumeric.GetAsIdString: AxUCString;
begin
  Result := FloatToStr(FValue);
end;

function TXLSSharedItemsValueNumeric.HashOf: longword;
var
  i: Integer;
  P: PByteArray;
begin
  P := @FValue;

  Result := 0;
  for i := 0 to SizeOf(FValue) - 1 do
    Result := ((Result shl 2) or (Result shr (SizeOf(Result) * 8 - 2))) xor Ord(P[i]);
end;

function TXLSSharedItemsValueNumeric.GetAsNumeric: double;
begin
  Result := FValue;
end;

function TXLSSharedItemsValueNumeric.GetAsText: AxUCString;
begin
  if Frac(FValue) = 0 then
    Result := FloatToStr(FValue)
  else
    Result := Format('%.2f',[FValue]);
end;

procedure TXLSSharedItemsValueNumeric.SetAsNumeric(const Value: double);
begin
  FValue := Value;
end;

function TXLSSharedItemsValueNumeric.Type_: TXLSSharedItemsValueType;
begin
  Result := xsivtNumeric;
end;

{ TXLSSharedItemsValueString }

function TXLSSharedItemsValueString.Clone: TXLSSharedItemsValue;
begin
  Result :=  TXLSSharedItemsValueString.Create;
  TXLSSharedItemsValueString(Result).FValue := FValue;
end;

function TXLSSharedItemsValueString.GetAsIdString: AxUCString;
begin
  Result := FValue;
end;

function TXLSSharedItemsValueString.HashOf: longword;
var
  i: Integer;
begin
  Result := 0;
  for i := 1 to Length(FValue) do
    Result := ((Result shl 2) or (Result shr (SizeOf(Result) * 8 - 2))) xor Ord(FValue[i]);
end;

function TXLSSharedItemsValueString.GetAsString: AxUCString;
begin
  Result := FValue;
end;

function TXLSSharedItemsValueString.GetAsText: AxUCString;
begin
  Result := FValue;
end;

procedure TXLSSharedItemsValueString.SetAsString(const Value: AxUCString);
begin
  FValue := Value;
end;

function TXLSSharedItemsValueString.Type_: TXLSSharedItemsValueType;
begin
  Result := xsivtString;
end;

{ TXLSSharedItemsValues }

function TXLSSharedItemsValues.Add: boolean;
var
  V: TXLSSharedItemsValueBlank;
begin
  V := TXLSSharedItemsValueBlank.Create;

  inherited Add(V);

  Result := True;
end;

function TXLSSharedItemsValues.Add(AValue: boolean): boolean;
var
  V: TXLSSharedItemsValueBoolean;
begin
  V := TXLSSharedItemsValueBoolean.Create;
  V.FValue := AValue;

  inherited Add(V);

  Result := True;
end;

function TXLSSharedItemsValues.Add(AValue: TDateTime{$ifndef DELPHI_XE_OR_LATER}; Dummy: integer {$endif}): boolean;
var
  V: TXLSSharedItemsValueDate;
begin
  V := TXLSSharedItemsValueDate.Create;
  V.FValue := AValue;

  inherited Add(V);

  Result := True;
end;

function TXLSSharedItemsValues.Add(AValue: TXc12CellError): boolean;
var
  V: TXLSSharedItemsValueError;
begin
  V := TXLSSharedItemsValueError.Create;
  V.FValue := AValue;

  inherited Add(V);

  Result := True;
end;

function TXLSSharedItemsValues.Add(AValue: double): boolean;
var
  V: TXLSSharedItemsValueNumeric;
begin
  V := TXLSSharedItemsValueNumeric.Create;
  V.FValue := AValue;

  inherited Add(V);

  Result := True;
end;

function TXLSSharedItemsValues.Add(AValue: AxUCString): boolean;
var
  V: TXLSSharedItemsValueString;
begin
  V := TXLSSharedItemsValueString.Create;
  V.FValue := AValue;

  inherited Add(V);

  Result := True;
end;

function TXLSSharedItemsValues.Add(AValue: TXLSSharedItemsValue): boolean;
begin
  inherited Add(AValue);

  Result := True;
end;

function TXLSSharedItemsValues.Compare(AIndex1, AIndex2: integer): double;
begin
  if Items[AIndex1].Type_ = Items[AIndex2].Type_ then begin
    case Items[AIndex1].Type_ of
      xsivtBoolean: Result := Integer(Items[AIndex1].AsBoolean) - Integer(Items[AIndex2].AsBoolean);
      xsivtDate   : Result := Items[AIndex1].AsDate - Items[AIndex2].AsDate;
      xsivtError  : Result := Integer(Items[AIndex1].AsError) - Integer(Items[AIndex2].AsError);
      xsivtBlank  : Result := 0;
      xsivtNumeric: Result := Items[AIndex1].AsDate - Items[AIndex2].AsDate;
      xsivtString : Result := CompareText(Items[AIndex1].AsString,Items[AIndex2].AsString);
      else          Result := 0;
    end;
  end
  else begin
    Result := Integer(Items[AIndex1].Type_) - Integer(Items[AIndex2].Type_);
  end;
end;

constructor TXLSSharedItemsValues.Create(AOwnsItems: boolean = True);
begin
  inherited Create(AOwnsItems);
end;

procedure TXLSSharedItemsValues.DoSort(ALeft, ARight: integer);
var
  I, J, P: Integer;
begin
  if ARight < ALeft then
    Exit;

  repeat
    I := ALeft;
    J := ARight;
    P := (ALeft + ARight) shr 1;
    repeat
      while Compare(I, P) < 0 do Inc(I);
      while Compare(J, P) > 0 do Dec(J);
      if I <= J then
      begin
        if I <> J then
         Exchange(I, J);
        if P = I then
          P := J
        else if P = J then
          P := I;
        Inc(I);
        Dec(J);
      end;
    until I > J;
    if ALeft < J then DoSort(ALeft, J);
    ALeft := I;
  until I >= ARight;
end;

function TXLSSharedItemsValues.Find(AValue: TXc12CellError): boolean;
var
  i: integer;
begin
  for i := 0 to Count - 1 do begin
    Result := Items[i].Equal(AValue);
    if Result then
      Exit;
  end;
  Result := False;
end;

function TXLSSharedItemsValues.Find(AValue: TDateTime{$ifndef DELPHI_XE_OR_LATER}; Dummy: integer {$endif}): boolean;
var
  i: integer;
begin
  for i := 0 to Count - 1 do begin
    Result := Items[i].Equal(AValue);
    if Result then
      Exit;
  end;
  Result := False;
end;

function TXLSSharedItemsValues.Find: boolean;
var
  i: integer;
begin
  for i := 0 to Count - 1 do begin
    Result := Items[i].Equal;
    if Result then
      Exit;
  end;
  Result := False;
end;

function TXLSSharedItemsValues.Find(AValue: AxUCString): boolean;
var
  i: integer;
begin
  Result := True;

  for i := 0 to Count - 1 do begin
    if Items[i].Equal(AValue) then
      Exit;
  end;
  Result := False;
end;

function TXLSSharedItemsValues.Find(AValue: double): boolean;
var
  i: integer;
begin
  for i := 0 to Count - 1 do begin
    Result := Items[i].Equal(AValue);
    if Result then
      Exit;
  end;
  Result := False;
end;

function TXLSSharedItemsValues.Find(AValue: boolean): boolean;
var
  i: integer;
begin
  for i := 0 to Count - 1 do begin
    Result := Items[i].Equal(AValue);
    if Result then
      Exit;
  end;
  Result := False;
end;

function TXLSSharedItemsValues.GetItems(Index: integer): TXLSSharedItemsValue;
begin
  Result := TXLSSharedItemsValue(inherited Items[Index]);
end;

procedure TXLSSharedItemsValues.Sort;
begin
  DoSort(0,Count - 1);
end;

function TXLSSharedItemsValues.Find(AValue: TXLSSharedItemsValue): boolean;
var
  i: integer;
begin
  Result := True;

  for i := 0 to Count - 1 do begin
    if Items[i].Equal(AValue) then
      Exit;
  end;
  Result := False;
end;

function TXLSSharedItemsValues.FindValue(AValue: TXLSSharedItemsValue): integer;
begin
  for Result := 0 to Count - 1 do begin
    if Items[Result].Equal(AValue) then
      Exit;
  end;
  Result := -1;
end;

{ TXLSSharedItemsValueBlank }

function TXLSSharedItemsValueBlank.Clone: TXLSSharedItemsValue;
begin
  Result :=  TXLSSharedItemsValueBlank.Create;
end;

function TXLSSharedItemsValueBlank.GetAsIdString: AxUCString;
begin
  Result := 'BLANK';
end;

function TXLSSharedItemsValueBlank.GetAsText: AxUCString;
begin
  Result := '(blank)';
end;

function TXLSSharedItemsValueBlank.HashOf: longword;
begin
  Result := 0;
end;

function TXLSSharedItemsValueBlank.Type_: TXLSSharedItemsValueType;
begin
  Result := xsivtBlank;
end;

{ TXLSHashedSharedItemsValues }

function TXLSUniqueSharedItemsValues.Add(AValue: TDateTime{$ifndef DELPHI_XE_OR_LATER}; Dummy: integer {$endif}): boolean;
var
  Hash: longword;
begin
  Result := False;

  Hash := HashOf(AValue) mod Longword(Length(FBuckets));
  if (FBuckets[Hash] <> Nil) and FBuckets[Hash].Find(AValue{$ifndef DELPHI_XE_OR_LATER}, 0 {$endif}) then
    Exit
  else if FBuckets[Hash] = Nil then
    FBuckets[Hash] := TXLSSharedItemsValues.Create;

  FBuckets[Hash].Add(AValue{$ifndef DELPHI_XE_OR_LATER}, 0 {$endif});

  inherited Add(AValue);

  Result := True;
end;

function TXLSUniqueSharedItemsValues.Add(AValue: boolean): boolean;
var
  Hash: longword;
begin
  Result := False;

  if AValue then
    Hash := 1
  else
    Hash := 0;
  if (FBuckets[Hash] <> Nil) and FBuckets[Hash].Find(AValue) then
    Exit
  else if FBuckets[Hash] = Nil then
    FBuckets[Hash] := TXLSSharedItemsValues.Create;

  FBuckets[Hash].Add(AValue);

  inherited Add(AValue);

  Result := True;
end;

function TXLSUniqueSharedItemsValues.Add(AValue: TXc12CellError): boolean;
var
  Hash: longword;
begin
  Result := False;

  Hash := HashOf(AValue) mod Longword(Length(FBuckets));
  if (FBuckets[Hash] <> Nil) and FBuckets[Hash].Find(AValue) then
    Exit
  else if FBuckets[Hash] = Nil then
    FBuckets[Hash] := TXLSSharedItemsValues.Create;

  FBuckets[Hash].Add(AValue);

  inherited Add(AValue);

  Result := True;
end;

function TXLSUniqueSharedItemsValues.Add(AValue: AxUCString): boolean;
var
  Hash: longword;
begin
  Result := False;

  Hash := HashOf(AValue) mod Longword(Length(FBuckets));
  if (FBuckets[Hash] <> Nil) and FBuckets[Hash].Find(AValue) then
    Exit
  else if FBuckets[Hash] = Nil then
    FBuckets[Hash] := TXLSSharedItemsValues.Create;

  FBuckets[Hash].Add(AValue);

  inherited Add(AValue);

  Result := True;
end;

function TXLSUniqueSharedItemsValues.Add(AValue: TXLSSharedItemsValue): boolean;
begin
  case AValue.Type_ of
    xsivtBoolean: Result := Add(AValue.AsBoolean);
    xsivtDate   : Result := Add(AValue.AsDate);
    xsivtError  : Result := Add(AValue.AsError);
    xsivtBlank  : Result := Add;
    xsivtNumeric: Result := Add(AValue.AsNumeric);
    xsivtString : Result := Add(AValue.AsString);
    else          Result := False;
  end;

  AValue.Free;
end;

procedure TXLSUniqueSharedItemsValues.Clear;
begin
  inherited Clear;

  ClearBuckets;
end;

procedure TXLSUniqueSharedItemsValues.ClearBuckets;
var
  i: integer;
begin
  for i := 0 to High(FBuckets) do begin
    if FBuckets[i] <> Nil then begin
      FBuckets[i].Free;
      FBuckets[i] := Nil;
    end;
  end;
end;

constructor TXLSUniqueSharedItemsValues.Create(AOwnsItems: boolean; Size: Cardinal);
begin
  inherited Create(AOwnsItems);

  SetLength(FBuckets,Size);
end;

function TXLSUniqueSharedItemsValues.Find(AValue: TXLSSharedItemsValue): boolean;
begin
  case AValue.Type_ of
    xsivtBoolean: Result := Find(AValue.AsBoolean);
    xsivtDate   : Result := Find(AValue.AsDate);
    xsivtError  : Result := Find(AValue.AsError);
    xsivtBlank  : Result := True;
    xsivtNumeric: Result := Find(AValue.AsNumeric);
    xsivtString : Result := Find(AValue.AsString);
    else          Result := False;
  end;
end;

function TXLSUniqueSharedItemsValues.Add(AValue: double): boolean;
var
  Hash: longword;
begin
  Result := False;

  Hash := HashOf(AValue) mod Longword(Length(FBuckets));
  if (FBuckets[Hash] <> Nil) and FBuckets[Hash].Find(AValue) then
    Exit
  else if FBuckets[Hash] = Nil then
    FBuckets[Hash] := TXLSSharedItemsValues.Create;

  FBuckets[Hash].Add(AValue);

  inherited Add(AValue);

  Result := True;
end;

function TXLSUniqueSharedItemsValues.Add: boolean;
var
  Hash: longword;
begin
  Result := False;

  Hash := 0;
  if (FBuckets[Hash] <> Nil) and FBuckets[Hash].Find then
    Exit
  else if FBuckets[Hash] = Nil then
    FBuckets[Hash] := TXLSSharedItemsValues.Create;

  FBuckets[Hash].Add;

  inherited Add;

  Result := True;
end;

function TXLSUniqueSharedItemsValues.Find(AValue: TXc12CellError): boolean;
var
  Hash: longword;
begin
  Hash := HashOf(AValue) mod Longword(Length(FBuckets));
  Result := (FBuckets[Hash] <> Nil) and FBuckets[Hash].Find(AValue);
end;

function TXLSUniqueSharedItemsValues.Find(AValue: TDateTime{$ifndef DELPHI_XE_OR_LATER}; Dummy: integer {$endif}): boolean;
var
  Hash: longword;
begin
  Hash := HashOf(AValue) mod Longword(Length(FBuckets));
  Result := (FBuckets[Hash] <> Nil) and FBuckets[Hash].Find(AValue);
end;

function TXLSUniqueSharedItemsValues.Find(AValue: boolean): boolean;
var
  Hash: longword;
begin
  if AValue then
    Hash := 1
  else
    Hash := 0;
  Result := (FBuckets[Hash] <> Nil) and FBuckets[Hash].Find(AValue);
end;

function TXLSUniqueSharedItemsValues.Find: boolean;
var
  Hash: longword;
begin
  Hash := 0;
  Result := (FBuckets[Hash] <> Nil) and FBuckets[Hash].Find;
end;

function TXLSUniqueSharedItemsValues.HashOf(AValue: TXc12CellError): longword;
var
  i: Integer;
  P: PByteArray;
begin
  P := @AValue;

  Result := 0;
  for i := 0 to SizeOf(AValue) - 1 do
    Result := ((Result shl 2) or (Result shr (SizeOf(Result) * 8 - 2))) xor Ord(P[i]);
end;

// function TXLSUniqueSharedItemsValues.HashOf(AValue: TDateTime): longword;
// begin
//   Result := HashOf(Double(AValue));
// end;

function TXLSUniqueSharedItemsValues.HashOf(AValue: AxUCString): longword;
var
  i: Integer;
begin
  Result := 0;
  for i := 1 to Length(AValue) do
    Result := ((Result shl 2) or (Result shr (SizeOf(Result) * 8 - 2))) xor Ord(AValue[i]);
end;

procedure TXLSUniqueSharedItemsValues.Resize(ASize: integer);
begin
  Clear;

  SetLength(FBuckets,ASize);
end;

function TXLSUniqueSharedItemsValues.HashOf(AValue: double): longword;
var
  i: Integer;
  P: PByteArray;
begin
  P := @AValue;

  Result := 0;
  for i := 0 to SizeOf(AValue) - 1 do
    Result := ((Result shl 2) or (Result shr (SizeOf(Result) * 8 - 2))) xor Ord(P[i]);
end;

function TXLSUniqueSharedItemsValues.Find(AValue: AxUCString): boolean;
var
  Hash: longword;
  Vals: TXLSSharedItemsValues;
begin
  Hash := HashOf(AValue) mod Longword(Length(FBuckets));

  Vals := FBuckets[Hash];
  Result := (FBuckets[Hash] <> Nil) and Vals.Find(AValue);
end;

function TXLSUniqueSharedItemsValues.Find(AValue: double): boolean;
var
  Hash: longword;
begin
  Hash := HashOf(AValue) mod Longword(Length(FBuckets));
  Result := (FBuckets[Hash] <> Nil) and FBuckets[Hash].Find(AValue);
end;

end.
