unit xpgPLists;

{$B-}
{$H+}
{$R-}
{$I AxCompilers.inc}

interface

uses Classes, SysUtils, Contnrs,
     xpgPUtils, xpgPXML;

type TErrorNotifyEvent = procedure(Sender: TObject) of object;

type TBaseXpgList = class(TObject)
private
     function GetAsString(Index: integer): string; virtual; abstract;
     procedure SetAsString(Index: integer; const Value: string); virtual; abstract;
protected
     FDelimiter: AxUCChar;
     FErrorEvent: TErrorNotifyEvent;

     procedure SetDelimiter(const Value: AxUCChar); virtual;
     function  GetDelimitedText: AxUCString; virtual; abstract;
     procedure SetDelimitedText(const Value: AxUCString); virtual; abstract;

     function  Count: integer; virtual; abstract;
public
     constructor Create;
     procedure Clear; virtual; abstract;

     procedure Write(AWriter: TXpgWriteXML; AName: string);
     procedure WriteElements(AWriter: TXpgWriteXML; AName: string);

     property Delimiter: AxUCChar read FDelimiter write SetDelimiter;
     property DelimitedText: AxUCString read GetDelimitedText write SetDelimitedText;
     property OnError: TErrorNotifyEvent read FErrorEvent write FErrorEvent;
     property AsString[Index: integer]: string read GetAsString write SetAsString;
     end;

type TItemXpgList = class(TBaseXpgList)
protected
     FItems: TList;
public
     constructor Create;
     destructor Destroy; override;
     procedure Clear; override;

     function  Count: integer; override;
     end;

type TStringXpgList = class(TBaseXpgList)
private
     function  GetAsString(Index: integer): string; override;
     procedure SetAsString(Index: integer; const Value: string); override;
     function  GetItems(Index: integer): AxUCString;
     procedure SetItems(Index: integer; const Value: AxUCString);
protected
     FItems: TStringList;

     procedure SetDelimiter(const Value: AxUCChar); override;
     function  GetDelimitedText: AxUCString; override;
     procedure SetDelimitedText(const Value: AxUCString); override;
public
     constructor Create;
     destructor Destroy; override;
     procedure Clear; override;
     function  IndexOf(AValue: AxUCString): integer;

     function  Count: integer; override;
     procedure Add(AValue: AxUCString);
     procedure AddNameValue(AName,AValue: AxUCString);
     function  Text: AxUCString;

     property Items[Index: integer]: AxUCString read GetItems write SetItems; default;
     end;

type TIntegerXpgList = class(TItemXpgList)
private
     function  GetAsString(Index: integer): string; override;
     procedure SetAsString(Index: integer; const Value: string); override;
     function  GetItems(Index: integer): integer;
     procedure SetItems(Index: integer; const Value: integer);
protected
     function  GetDelimitedText: AxUCString; override;
     procedure SetDelimitedText(const Value: AxUCString); override;
public
     procedure Add(AValue: integer); overload;
     procedure Add(AValue: AxUCString); overload;

     property Items[Index: integer]: integer read GetItems write SetItems; default;
     end;

type TExtendedXpgList = class(TItemXpgList)
private
     function  GetAsString(Index: integer): string; override;
     procedure SetAsString(Index: integer; const Value: string); override;
     function  GetItems(Index: integer): extended;
protected
     function  GetDelimitedText: AxUCString; override;
     procedure SetDelimitedText(const Value: AxUCString); override;
public
     procedure Add(AValue: extended); overload;
     procedure Add(AValue: AxUCString); overload;

     property Items[Index: integer]: extended read GetItems;
     end;

type TLongwordXpgList = class(TItemXpgList)
private
     function  GetAsString(Index: integer): string; override;
     procedure SetAsString(Index: integer; const Value: string); override;
     function  GetItems(Index: integer): longword;
protected
     function  GetDelimitedText: AxUCString; override;
     procedure SetDelimitedText(const Value: AxUCString); override;
public
     procedure Add(AValue: longword); overload;
     procedure Add(AValue: AxUCString); overload;

     property Items[Index: integer]: longword read GetItems; default;
     end;

type TTDateTimeXpgList = class(TItemXpgList)
private
     function  GetAsString(Index: integer): string; override;
     procedure SetAsString(Index: integer; const Value: string); override;
     function  GetItems(Index: integer): TDateTime;
protected
     function  GetDelimitedText: AxUCString; override;
     procedure SetDelimitedText(const Value: AxUCString); override;
public
     procedure Add(AValue: TDateTime); overload;
     procedure Add(AValue: AxUCString); overload;

     property Items[Index: integer]: TDateTime read GetItems;
     end;

type PXpgEnumData = ^TXpgEnumData;
     TXpgEnumData = record
     Enum : integer;
     Index: integer;
     end;

type TEnumXpgList = class(TItemXpgList)
private
     function  GetAsString(Index: integer): string; override;
     procedure SetAsString(Index: integer; const Value: string); override;
     function  GetItems(Index: integer): integer;
protected
     FEnums: TStrings;
     FPrefix: AxUCString;

     function  GetDelimitedText: AxUCString; override;
     procedure SetDelimitedText(const Value: AxUCString); override;
     procedure DoAdd(AValue,AIndex: integer);
public
     constructor Create(AEnums: TStrings; APrefix: AxUCString);
     procedure Add(AValue: integer); overload;
     procedure Add(AValue: AxUCString); overload;

     property Items[Index: integer]: integer read GetItems;
     end;

type TFloatXpgList = class(TItemXpgList)
private
     function  GetAsString(Index: integer): string; override;
     procedure SetAsString(Index: integer; const Value: string); override;
     function  GetItems(Index: integer): extended;
protected
     function  GetDelimitedText: AxUCString; override;
     procedure SetDelimitedText(const Value: AxUCString); override;
public
     procedure Add(AValue: extended); overload;
     procedure Add(AValue: AxUCString); overload;

     property Items[Index: integer]: extended read GetItems;
     end;

type TBooleanXpgList = class(TItemXpgList)
private
     function  GetAsString(Index: integer): string; override;
     procedure SetAsString(Index: integer; const Value: string); override;
     function  GetItems(Index: integer): boolean;
protected
     function  GetDelimitedText: AxUCString; override;
     procedure SetDelimitedText(const Value: AxUCString); override;
public
     procedure Add(AValue: boolean); overload;
     procedure Add(AValue: AxUCString); overload;

     property Items[Index: integer]: boolean read GetItems;
     end;

implementation

{ TXpgIntegerList }

procedure  TIntegerXpgList.Add(AValue: integer);
var
  P: PInteger;
begin
  GetMem(P,SizeOf(Integer));
  P^ := AValue;
  FItems.Add(P);
end;

procedure TIntegerXpgList.Add(AValue: AxUCString);
begin
  Add(StrToInt(AValue));
end;

function TIntegerXpgList.GetAsString(Index: integer): string;
begin
  Result := IntToStr(PInteger(FItems.Items[Index])^);
end;

function TIntegerXpgList.GetDelimitedText: AxUCString;
var
  i: integer;
begin
  Result := '';
  for i := 0 to FItems.Count - 1 do
    Result := Result + IntToStr(GetItems(i)) + ' ';
  if Length(Result) > 0 then
    SetLength(Result,Length(Result) - 1);
end;

function TIntegerXpgList.GetItems(Index: integer): integer;
begin
  Result := PInteger(FItems.Items[Index])^;
end;

procedure TIntegerXpgList.SetAsString(Index: integer; const Value: string);
begin
  PInteger(FItems.Items[Index])^ := StrToInt(Value);
end;

procedure TIntegerXpgList.SetDelimitedText(const Value: AxUCString);
var
  V: integer;
  S1,S2: AxUCString;
begin
  Clear;
  S1 := Value;
  while S1 <> '' do begin
    S2 := SplitAtChar(FDelimiter,S1);
    if TryStrToInt(S2,V) then
      Add(V)
    else if Assigned(FErrorEvent) then
      FErrorEvent(Self);
  end;
end;

procedure TIntegerXpgList.SetItems(Index: integer; const Value: integer);
begin
  PInteger(FItems.Items[Index])^ := Value;
end;

{ TXpgFloatList }

procedure TFloatXpgList.Add(AValue: extended);
var
  P: PExtended;
begin
  GetMem(P,SizeOf(extended));
  P^ := AValue;
  FItems.Add(P);
end;

procedure TFloatXpgList.Add(AValue: AxUCString);
begin
  Add(XmlStrToFloat(AValue));
end;

function TFloatXpgList.GetAsString(Index: integer): string;
begin
  Result := XmlFloatToStr(PExtended(FItems.Items[Index])^);
end;

function TFloatXpgList.GetDelimitedText: AxUCString;
var
  i: integer;
begin
  Result := '';
  for i := 0 to FItems.Count - 1 do
    Result := Result + XmlFloatToStr(GetItems(i)) + ' ';
  if Length(Result) > 0 then
    SetLength(Result,Length(Result) - 1);
end;

function TFloatXpgList.GetItems(Index: integer): extended;
begin
  Result := PExtended(FItems.Items[Index])^;
end;

procedure TFloatXpgList.SetAsString(Index: integer; const Value: string);
begin
  PExtended(FItems.Items[Index])^ := XmlStrToFloat(Value);
end;

procedure TFloatXpgList.SetDelimitedText(const Value: AxUCString);
var
  V: extended;
  S1,S2: AxUCString;
begin
  Clear;
  S1 := Value;
  while S1 <> '' do begin
    S2 := SplitAtChar(FDelimiter,S1);
    if XmlTryStrToFloat(S2,V) then
      Add(V)
    else if Assigned(FErrorEvent) then
      FErrorEvent(Self);
  end;
end;

{ TXpgBooleanList }

procedure TBooleanXpgList.Add(AValue: boolean);
var
  P: PBoolean;
begin
  GetMem(P,SizeOf(boolean));
  P^ := AValue;
  FItems.Add(P);
end;

procedure TBooleanXpgList.Add(AValue: AxUCString);
begin
  Add(XmlStrToBool(AValue));
end;

function TBooleanXpgList.GetAsString(Index: integer): string;
begin
  if PBoolean(FItems.Items[Index])^ then
    Result := 'true'
  else
    Result := 'false';
end;

function TBooleanXpgList.GetDelimitedText: AxUCString;
var
  i: integer;
begin
  Result := '';
  for i := 0 to FItems.Count - 1 do
    Result := Result + BoolToStr(GetItems(i),True) + ' ';
  if Length(Result) > 0 then
    SetLength(Result,Length(Result) - 1);
end;

function TBooleanXpgList.GetItems(Index: integer): boolean;
begin
  Result := PBoolean(FItems.Items[Index])^;
end;

procedure TBooleanXpgList.SetAsString(Index: integer; const Value: string);
begin
  PBoolean(FItems.Items[Index])^ := XmlStrToBool(Value)
end;

procedure TBooleanXpgList.SetDelimitedText(const Value: AxUCString);
var
  V: boolean;
  S1,S2: AxUCString;
begin
  Clear;
  S1 := Value;
  while S1 <> '' do begin
    S2 := SplitAtChar(FDelimiter,S1);
    if XmlTryStrToBool(S2,V) then
      Add(V)
    else if Assigned(FErrorEvent) then
      FErrorEvent(Self);
  end;
end;

{ TXpgItemList }

procedure TItemXpgList.Clear;
begin
  FItems.Clear;
end;

function TItemXpgList.Count: integer;
begin
  Result := FItems.Count;
end;

constructor TItemXpgList.Create;
begin
  inherited Create;
  FItems := TList.Create;
end;

destructor TItemXpgList.Destroy;
var
  i: integer;
begin
  for i := 0 to FItems.Count - 1 do
    FreeMem(FItems[i]);
  FItems.Free;
  inherited;
end;

{ TXpgBaseList }

constructor TBaseXpgList.Create;
begin
  FDelimiter := ' ';
end;

procedure TBaseXpgList.SetDelimiter(const Value: AxUCChar);
begin
  FDelimiter := Value;
end;

procedure TBaseXpgList.Write(AWriter: TXpgWriteXML; AName: string);
var
  S: string;
begin
  S := GetDelimitedText;
  if S <> '' then
    AWriter.SimpleTextTag(AName,S);
end;

procedure TBaseXpgList.WriteElements(AWriter: TXpgWriteXML; AName: string);
var
  i: integer;
begin
  for i := 0 to Count - 1 do
    AWriter.SimpleTextTag(AName,AsString[i]);
end;

{ TXpgStringList }

procedure TStringXpgList.Add(AValue: AxUCString);
begin
  FItems.Add(AValue);
end;

procedure TStringXpgList.AddNameValue(AName, AValue: AxUCString);
begin
  Add(AName + '="' + AValue + '" ');
end;

procedure TStringXpgList.Clear;
begin
  FItems.Clear;
end;

function TStringXpgList.Count: integer;
begin
  Result := FItems.Count;
end;

constructor TStringXpgList.Create;
begin
  inherited Create;
  FItems := TStringList.Create;
{$ifndef DELPHI_5}
  FItems.Delimiter := Char(FDelimiter);
{$endif}
{$ifdef DELPHI_2006_OR_LATER}
  FItems.StrictDelimiter := True;
{$endif}
end;

destructor TStringXpgList.Destroy;
begin
  FItems.Free;
  inherited;
end;

function TStringXpgList.GetAsString(Index: integer): string;
begin
  Result := Items[Index];
end;

function TStringXpgList.GetDelimitedText: AxUCString;
begin
{$ifdef DELPHI_5}
  Result := StringsToDelimited(FItems,FDelimiter);
{$else}
  Result := FItems.DelimitedText;
{$endif}
end;

function TStringXpgList.GetItems(Index: integer): AxUCString;
begin
  Result := FItems[Index];
end;

function TStringXpgList.IndexOf(AValue: AxUCString): integer;
begin
  Result := FItems.IndexOf(AValue);
end;

procedure TStringXpgList.SetAsString(Index: integer; const Value: string);
begin
  FItems[Index] := Value;
end;

procedure TStringXpgList.SetDelimitedText(const Value: AxUCString);
begin
  Clear;

{$ifdef DELPHI_5}
  DelimitedToStrings(Value,FItems,FDelimiter);
{$else}
  FItems.DelimitedText := Value;
{$endif}
end;

procedure TStringXpgList.SetDelimiter(const Value: AxUCChar);
begin
  FDelimiter := Value;
{$ifndef DELPHI_5}
  FItems.Delimiter := Char(Value);
{$endif}
end;

procedure TStringXpgList.SetItems(Index: integer; const Value: AxUCString);
begin
  FItems[Index] := Value;
end;

function TStringXpgList.Text: AxUCString;
begin
  Result := FItems.Text;
end;

{ TXpgEnumList }

procedure TEnumXpgList.Add(AValue: integer);
begin
  raise Exception.Create('Not implemented');
end;

procedure TEnumXpgList.Add(AValue: AxUCString);
var
  i: integer;
begin
  i := FEnums.IndexOf(AValue);
  if i >= 0 then
    DoAdd(Integer(FEnums.Objects[i]),i)
  else if Assigned(FErrorEvent) then
    FErrorEvent(Self);
end;

constructor TEnumXpgList.Create(AEnums: TStrings; APrefix: AxUCString);
begin
  inherited Create;

  FEnums := AEnums;
  FPrefix := APrefix;
end;

procedure TEnumXpgList.DoAdd(AValue, AIndex: integer);
var
  P: PXpgEnumData;
begin
  GetMem(P,SizeOf(TXpgEnumData));
  P.Enum := AValue;
  P.Index := AIndex;
  FItems.Add(P);
end;

function TEnumXpgList.GetAsString(Index: integer): string;
begin
  Result := FEnums[PXpgEnumData(FItems.Items[Index])^.Index];
end;

function TEnumXpgList.GetDelimitedText: AxUCString;
var
  i,n: integer;
  P: PXpgEnumData;
begin
  Result := '';
  n := Length(FPrefix) + 1;
  for i := 0 to FItems.Count - 1 do begin
    P := PXpgEnumData(FItems[i]);
    Result := Result + Copy(FEnums[P.Index],n,MAXINT) + ' ';
  end;
  if Length(Result) > 0 then
    SetLength(Result,Length(Result) - 1);
end;

function TEnumXpgList.GetItems(Index: integer): integer;
begin
  Result := PXpgEnumData(FItems.Items[Index])^.Enum;
end;

procedure TEnumXpgList.SetAsString(Index: integer; const Value: string);
var
  i: integer;
begin
  i := FEnums.IndexOf(Value);
  if i >= 0 then begin
    PXpgEnumData(FItems.Items[Index])^.Enum := Integer(FEnums.Objects[i]);
    PXpgEnumData(FItems.Items[Index])^.Index := i;
  end
  else if Assigned(FErrorEvent) then
    FErrorEvent(Self);
end;

procedure TEnumXpgList.SetDelimitedText(const Value: AxUCString);
var
  i: integer;
  S1,S2: AxUCString;
begin
  Clear;
  S1 := Value;
  while S1 <> '' do begin
    S2 := SplitAtChar(FDelimiter,S1);
    i := FEnums.IndexOf(FPrefix + S2);
    if i >= 0 then
      DoAdd(Integer(FEnums.Objects[i]),i)
    else if Assigned(FErrorEvent) then
      FErrorEvent(Self);
  end;
end;

{ TExtendedXpgList }

procedure TExtendedXpgList.Add(AValue: extended);
var
  P: PExtended;
begin
  GetMem(P,SizeOf(extended));
  P^ := AValue;
  FItems.Add(P);
end;

procedure TExtendedXpgList.Add(AValue: AxUCString);
begin
  Add(XmlStrToFloat(AValue));
end;

function TExtendedXpgList.GetAsString(Index: integer): string;
begin
  Result := XmlFloatToStr(PExtended(FItems.Items[Index])^);
end;

function TExtendedXpgList.GetDelimitedText: AxUCString;
var
  i: integer;
begin
  Result := '';
  for i := 0 to FItems.Count - 1 do
    Result := Result + FloatToStr(GetItems(i)) + ' ';
  if Length(Result) > 0 then
    SetLength(Result,Length(Result) - 1);
end;

function TExtendedXpgList.GetItems(Index: integer): extended;
begin
  Result := PExtended(FItems.Items[Index])^;
end;

procedure TExtendedXpgList.SetAsString(Index: integer; const Value: string);
begin
  PExtended(FItems.Items[Index])^ := XmlStrToFloat(Value);
end;

procedure TExtendedXpgList.SetDelimitedText(const Value: AxUCString);
var
  V: extended;
  S1,S2: AxUCString;
begin
  Clear;
  S1 := Value;
  while S1 <> '' do begin
    S2 := SplitAtChar(FDelimiter,S1);
    if XmlTryStrToFloat(S2,V) then
      Add(V)
    else if Assigned(FErrorEvent) then
      FErrorEvent(Self);
  end;
end;

{ TLongwordXpgList }

procedure TLongwordXpgList.Add(AValue: longword);
var
  P: PLongword;
begin
  GetMem(P,SizeOf(longword));
  P^ := AValue;
  FItems.Add(P);
end;

procedure TLongwordXpgList.Add(AValue: AxUCString);
begin
  Add(StrToInt(AValue));
end;

function TLongwordXpgList.GetAsString(Index: integer): string;
begin
  Result := IntToStr(PLongword(FItems.Items[Index])^);
end;

function TLongwordXpgList.GetDelimitedText: AxUCString;
var
  i: integer;
begin
  Result := '';
  for i := 0 to FItems.Count - 1 do
    Result := Result + IntToStr(GetItems(i)) + ' ';
  if Length(Result) > 0 then
    SetLength(Result,Length(Result) - 1);
end;

function TLongwordXpgList.GetItems(Index: integer): longword;
begin
  Result := PLongword(FItems.Items[Index])^;
end;

procedure TLongwordXpgList.SetAsString(Index: integer; const Value: string);
begin
  PLongword(FItems.Items[Index])^ := StrToInt(Value);
end;

procedure TLongwordXpgList.SetDelimitedText(const Value: AxUCString);
var
  V: longword;
  S1,S2: AxUCString;
begin
  Clear;
  S1 := Value;
  while S1 <> '' do begin
    S2 := SplitAtChar(FDelimiter,S1);
    if XmlTryStrToInt(S2,V) then
      Add(V)
    else if Assigned(FErrorEvent) then
      FErrorEvent(Self);
  end;
end;

{ TTDateTimeXpgList }

procedure TTDateTimeXpgList.Add(AValue: TDateTime);
var
  P: PDateTime;
begin
  GetMem(P,SizeOf(TDateTime));
  P^ := AValue;
  FItems.Add(P);
end;

procedure TTDateTimeXpgList.Add(AValue: AxUCString);
begin
  Add(XmlStrToDateTime(AValue));
end;

function TTDateTimeXpgList.GetAsString(Index: integer): string;
begin
  Result := XmlDateTimeToStr(PDateTime(FItems.Items[Index])^);
end;

function TTDateTimeXpgList.GetDelimitedText: AxUCString;
var
  i: integer;
begin
  Result := '';
  for i := 0 to FItems.Count - 1 do
    Result := Result + FloatToStr(GetItems(i)) + ' ';
  if Length(Result) > 0 then
    SetLength(Result,Length(Result) - 1);
end;

function TTDateTimeXpgList.GetItems(Index: integer): TDateTime;
begin
  Result := PDateTime(FItems.Items[Index])^;
end;

procedure TTDateTimeXpgList.SetAsString(Index: integer; const Value: string);
begin
  PDateTime(FItems.Items[Index])^ := XmlStrToDateTime(Value);
end;

procedure TTDateTimeXpgList.SetDelimitedText(const Value: AxUCString);
var
  V: TDateTime;
  S1,S2: AxUCString;
begin
  Clear;
  S1 := Value;
  while S1 <> '' do begin
    S2 := SplitAtChar(FDelimiter,S1);
    if XmlTryStrToDateTime(S2,V) then
      Add(V)
    else if Assigned(FErrorEvent) then
      FErrorEvent(Self);
  end;
end;

end.