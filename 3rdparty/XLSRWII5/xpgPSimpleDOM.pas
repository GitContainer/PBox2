unit xpgPSimpleDOM;

interface

uses Classes, SysUtils, Contnrs,
     xpgPUtils, xpgPXML;

type TXpgDOMAttribute = class(TObject)
protected
     FName: AxUCString;
     FValue: AxUCString;

     function  GetAsInteger: integer;
     function  GetAsBase64: AnsiString;
public
     property Name     : AxUCString read FName write FName;
     property Value    : AxUCString read FValue write FValue;
     property AsInteger: integer read GetAsInteger;
     property AsBase64 : AnsiString read GetAsBase64;
     end;

type TXpgDOMAttributes = class(TObject)
private
     function GetItems(Index: integer): TXpgDOMAttribute;
protected
     FItems: TObjectList;
public
     constructor Create;
     destructor Destroy; override;

     function  AsText: AxUCString;

     procedure Clear;
     procedure Assign(ASource: TXpgDOMAttributes);
     function  Add(const AName,AValue: AxUCString): TXpgDOMAttribute; overload;
     function  Add(const AName: AxUCString; AValue: integer): TXpgDOMAttribute; overload;
     function  Add(const AName: AxUCString; AValue: double): TXpgDOMAttribute; overload;
     function  Add(const AName: AxUCString; AValue: boolean): TXpgDOMAttribute; overload;
     function  Update(const AName,AValue: AxUCString): TXpgDOMAttribute; overload;
     function  Update(const AName: AxUCString; AValue: integer): TXpgDOMAttribute; overload;
     function  Update(const AName: AxUCString; AValue: boolean): TXpgDOMAttribute; overload;
     function  Find(const AName: AxUCString): TXpgDOMAttribute;
     function  FindValue(const AName: AxUCString): AxUCString;
     function  FindValueDef(const AName, ADefault: AxUCString): AxUCString;
     function  FindValueInt(const AName: AxUCString): integer;
     function  FindValueIntDef(const AName: AxUCString; const ADefault: integer): integer;
     function  FindValueHexDef(const AName: AxUCString; const ADefault: integer): integer;
     function  FindValueBool(const AName: AxUCString): boolean;
     function  FindValueFloatDef(const AName: AxUCString; const ADefault: integer): double;
     // Used with class and id attributes.
     function  EqualValue(const AName,AMatchPrefix,AValue: AxUCString): boolean;
     function  Clone: TXpgDOMAttributes;

     function  Count: integer;

     property Items[Index: integer]: TXpgDOMAttribute read GetItems; default;
     end;

     type TXpgDOMNode = class;

     TXpgDOMNodeList = class(TObjectList)
private
     function GetItems(Index: integer): TXpgDOMNode;
protected
public
     constructor Create;

     property Items[Index: integer]: TXpgDOMNode read GetItems; default;
     end;

     TXpgDOMNode = class(TObject)
private
     function  GetItems(Index: integer): TXpgDOMNode;
     function  GetAsBoolean: boolean;
     function  GetAsFloat: extended;
     function  GetAsInteger: integer;
     function  GetAsString: AxUCString;
     function  GetAsDateTime: TDateTime;
     procedure SetAsBoolean(const Value: boolean);
     procedure SetAsFloat(const Value: extended);
     procedure SetAsInteger(const Value: integer);
     procedure SetAsString(const Value: AxUCString);
     procedure SetAsDateTime(const Value: TDateTime);
protected
     FParent    : TXpgDOMNode;
     FQName     : AxUCString;
     FContent   : AxUCString;
     FAttributes: TXpgDOMAttributes;
     FChilds    : TObjectList;

     function  CreateNode(AParent: TXpgDOMNode): TXpgDOMNode; virtual;

     procedure DoFindAll(const APathName: AxUCString; AResult: TXpgDOMNodeList);
public
     constructor Create(AParent: TXpgDOMNode);
     destructor Destroy; override;

     procedure Clear;

     procedure Assign(ANode: TXpgDOMNode);

     function  AddChild(const AQName: AxUCString): TXpgDOMNode; overload;
     function  AddChild(const AQName, AContent: AxUCString): TXpgDOMNode; overload;
     function  AddChild(const AQName: AxUCString; AContent: double): TXpgDOMNode; overload;
     function  AddChild(const AQName: AxUCString; AContent: boolean): TXpgDOMNode; overload;

     function  CheckChild(const AQName: AxUCString): TXpgDOMNode;
     procedure AddAttribute(const AQName, AValue: AxUCString); overload;
     procedure AddAttribute(const AQName: AxUCString; AValue: integer); overload;
     procedure AddAttribute(const AQName: AxUCString; AValue: boolean); overload;
     function  CheckAttribute(const AQName, ADefaultValue: AxUCString): TXpgDOMAttribute;
     procedure Delete(const AIndex: integer);
     function  Detach(const AIndex: integer): TXpgDOMNode;
     function  DeleteChildFirst(const AName: AxUCString): boolean; overload;
     function  DeleteChildFirst(const ANames: array of AxUCString): boolean; overload;
     function  DeleteChildAll(const AName: AxUCString): integer;
     function  Count: integer;
     function  Find(const APathName: AxUCString; ARequired: boolean = False): TXpgDOMNode; overload;
     function  Find(const ANames: array of AxUCString; ARequired: boolean; out ANode: TXpgDOMNode): integer; overload;
     function  FindAll(const APathName: AxUCString): TXpgDOMNodeList;
     function  FindChildWithAttrValue(const AAttribute,AValue: string): TXpgDOMNode;
     function  GetContent(const AName: AxUCString): AxUCString;
     function  Clone: TXpgDOMNode;
     procedure Insert(ANode: TXpgDOMNode);

     function  Index: integer;
     function  PrevSibling: TXpgDOMNode;
     function  NextSibling: TXpgDOMNode;

     property Parent: TXpgDOMNode read FParent;
     property QName: AxUCString read FQName write FQName;
     property Content: AxUCString read FContent write FContent;
     property Attributes: TXpgDOMAttributes read FAttributes;

     property AsString: AxUCString read GetAsString write SetAsString;
     property AsInteger: integer read GetAsInteger write SetAsInteger;
     property AsFloat: extended read GetAsFloat write SetAsFloat;
     property AsBoolean: boolean read GetAsBoolean write SetAsBoolean;
     property AsDateTime: TDateTime read GetAsDateTime write SetAsDateTime;

     property Items[Index: integer]: TXpgDOMNode read GetItems; default;
     end;

type TXpgSimpleDOM = class(TXpgReadXML)
protected
     FStack   : TObjectStack;
     FRoot    : TXpgDOMNode;

     function  CreateRoot: TXpgDOMNode; virtual;
public
     constructor Create;
     destructor Destroy; override;

     procedure Clear; override;
     procedure BeginTag; override;
     procedure EndTag; override;

     function  Empty: boolean;
     function  LoadFromFile(const AFilename: string): integer;
     function  LoadFromStream(AStream: TStream): integer; override;
     function  LoadFromStreamHtml(AStream: TStream): integer; override;
     procedure LoadFromString(AStr: AxUCString);
     function  LoadFromBuffer(ABuf: PByteArray; AMaxSize: integer): integer;
     procedure SaveToFile(const AFilename: string);
     procedure SaveToStream(AStream: TStream);
     function  SaveToString: AxUCString;

     property Root: TXpgDOMNode read FRoot;
     end;

implementation

{ TXpgDOMNode }

procedure TXpgDOMNode.AddAttribute(const AQName, AValue: AxUCString);
begin
  FAttributes.Add(AQName,AValue);
end;

function TXpgDOMNode.AddChild(const AQName: AxUCString): TXpgDOMNode;
begin
  Result := CreateNode(Self);
  Result.QName := AQName;
  FChilds.Add(Result);
end;

procedure TXpgDOMNode.AddAttribute(const AQName: AxUCString; AValue: integer);
begin
  FAttributes.Add(AQName,XmlIntToStr(AValue));
end;

procedure TXpgDOMNode.AddAttribute(const AQName: AxUCString; AValue: boolean);
begin
  FAttributes.Add(AQName,XmlBoolToStr(AValue));
end;

function TXpgDOMNode.AddChild(const AQName: AxUCString; AContent: boolean): TXpgDOMNode;
begin
  Result := CreateNode(Self);
  Result.QName := AQName;
  if AContent then
    Result.Content := 'true'
  else
    Result.Content := 'false';
  FChilds.Add(Result);
end;

procedure TXpgDOMNode.Assign(ANode: TXpgDOMNode);
var
  i: integer;
begin
  FQName := ANode.FQName;
  FContent := ANode.FContent;

  FAttributes.Assign(ANode.FAttributes);

  for i := 0 to ANode.Count - 1 do
    AddChild(ANode.FQName).Assign(ANode[i]);
end;

function TXpgDOMNode.AddChild(const AQName, AContent: AxUCString): TXpgDOMNode;
begin
  Result := CreateNode(Self);
  Result.QName := AQName;
  Result.Content := AContent;
  FChilds.Add(Result);
end;

function TXpgDOMNode.CheckAttribute(const AQName, ADefaultValue: AxUCString): TXpgDOMAttribute;
begin
  Result := FAttributes.Find(AQName);
  if Result = Nil then
    Result := FAttributes.Add(AQName,ADefaultValue);
end;

function TXpgDOMNode.CheckChild(const AQName: AxUCString): TXpgDOMNode;
begin
  Result := Find(AQName,False);
  if Result = Nil then
    AddChild(AQName);
end;

procedure TXpgDOMNode.Clear;
begin
  FChilds.Clear;
end;

function TXpgDOMNode.Clone: TXpgDOMNode;

procedure DoClone(ASource, ADest: TXpgDOMNode);
var
  i: integer;
begin
  ADest.FParent := ASource.FParent;
  ADest.FContent := ASource.Content;
  ADest.FAttributes.Assign(ASource.Attributes);

  for i := 0 to ASource.FChilds.Count - 1 do
    DoClone(ASource[i],ADest.AddChild(ASource[i].QName));
end;

begin
  Result := CreateNode(Nil);
  Result.QName := QName;
  DoClone(Self,Result);
end;

function TXpgDOMNode.Count: integer;
begin
  Result := FChilds.Count;
end;

constructor TXpgDOMNode.Create(AParent: TXpgDOMNode);
begin
  FParent := AParent;
  FAttributes := TXpgDOMAttributes.Create;
  FChilds := TObjectList.Create;
end;

function TXpgDOMNode.CreateNode(AParent: TXpgDOMNode): TXpgDOMNode;
begin
  Result := TXpgDOMNode.Create(AParent);
end;

procedure TXpgDOMNode.Delete(const AIndex: integer);
begin
  FChilds.Delete(AIndex);
end;

function TXpgDOMNode.DeleteChildAll(const AName: AxUCString): integer;
var
  i,n: integer;
begin
  Result := 0;
  n := FChilds.Count;
  i := 0;
  while i < n do begin
    if Items[i].QName = AName then begin
      FChilds.Delete(i);
      Dec(n);
      Inc(Result);
    end
    else
      Inc(i);
  end;
end;

function TXpgDOMNode.DeleteChildFirst(const ANames: array of AxUCString): boolean;
var
  i: integer;
begin
  Result := False;
  for i := Low(ANames) to High(ANames) do begin
    if DeleteChildFirst(ANames[i]) then
      Result := True;
  end;
end;

function TXpgDOMNode.DeleteChildFirst(const AName: AxUCString): boolean;
var
  i: integer;
begin
  for i := 0 to FChilds.Count - 1 do begin
    if Items[i].QName = AName then begin
      FChilds.Delete(i);
      Result := True;
      Exit;
    end;
  end;
  Result := False;
end;

destructor TXpgDOMNode.Destroy;
begin
  FAttributes.Free;
  FChilds.Free;
  inherited;
end;

function TXpgDOMNode.Detach(const AIndex: integer): TXpgDOMNode;
begin
  Result := TXpgDOMNode(FChilds[AIndex]);
  FChilds.OwnsObjects := False;
  FChilds[AIndex] := Nil;
  FChilds.OwnsObjects := True;
  FChilds.Delete(AIndex);
end;

procedure TXpgDOMNode.DoFindAll(const APathName: AxUCString; AResult: TXpgDOMNodeList);
var
  i: integer;
  p: integer;
  Path,Name: AxUCString;
begin
  p := CPos('/',APathName);
  if p > 0 then begin
    Name := Copy(APathName,1,p - 1);
    Path := Copy(APathName,p + 1,MAXINT);
  end
  else
    Name := APathName;

  for i := 0 to FChilds.Count - 1 do begin
    if Items[i].QName = Name then begin
      if p > 0 then
        Items[i].DoFindAll(Path,AResult)
      else
        AResult.Add(Items[i]);
    end;
  end;
end;

function TXpgDOMNode.Find(const ANames: array of AxUCString; ARequired: boolean; out ANode: TXpgDOMNode): integer;
var
  i: integer;
begin
  for Result := Low(ANames) to High(ANames) do begin
    for i := 0 to FChilds.Count - 1 do begin
      if Items[i].QName = ANames[Result] then begin
        ANode := Items[i];
        Exit;
      end;
    end;
  end;
  if ARequired then
    raise Exception.Create('Can not find expected node');
  Result := -1;
  ANode := Nil;
end;

function TXpgDOMNode.FindAll(const APathName: AxUCString): TXpgDOMNodeList;
begin
  Result := TXpgDOMNodeList.Create;
  DoFindAll(APathName,Result);
end;

function TXpgDOMNode.FindChildWithAttrValue(const AAttribute, AValue: string): TXpgDOMNode;
var
  i: integer;
  S: string;
begin
  for i := 0 to Count - 1 do begin
    S := Items[i].Attributes.FindValue(AAttribute);
    if S = AValue then begin
      Result := Items[i];
      Exit;
    end;
  end;
  Result := Nil;
end;

function TXpgDOMNode.Find(const APathName: AxUCString; ARequired: boolean): TXpgDOMNode;
var
  i: integer;
  p: integer;
  Path,Name: AxUCString;
begin
  p := CPos('/',APathName);
  if p > 0 then begin
    Name := Copy(APathName,1,p - 1);
    Path := Copy(APathName,p + 1,MAXINT);
  end
  else
    Name := APathName;

  for i := 0 to FChilds.Count - 1 do begin
    if Items[i].QName = Name then begin
      if p > 0 then
        Result := Items[i].Find(Path,ARequired)
      else
        Result := Items[i];
      Exit;
    end;
  end;
  if ARequired then
    raise Exception.CreateFmt('XML error: Can not find expected node "%s"',[APathName]);
  Result := Nil;
end;

function TXpgDOMNode.GetAsBoolean: boolean;
begin
  if not XmlTryStrToBool(FContent,Result) then
    Result := False;
end;

function TXpgDOMNode.GetAsDateTime: TDateTime;
begin
  if not XmlTryStrToDateTime(FContent,Result) then
    Result := 0;
end;

function TXpgDOMNode.GetAsFloat: extended;
begin
  if not XmlTryStrToFloat(FContent,Result) then
    Result := 0;
end;

function TXpgDOMNode.GetAsInteger: integer;
begin
  if not XmlTryStrToInt(FContent,Result) then
    Result := 0;
end;

function TXpgDOMNode.GetAsString: AxUCString;
begin
  Result := FContent;
end;

function TXpgDOMNode.GetContent(const AName: AxUCString): AxUCString;
var
  N: TXpgDOMNode;
begin
  N := Find(Aname,False);
  if N <> Nil then
    Result := N.Content
  else
    Result := '';
end;

function TXpgDOMNode.GetItems(Index: integer): TXpgDOMNode;
begin
  Result := TXpgDOMNode(FChilds[Index]);
end;

function TXpgDOMNode.Index: integer;
begin
  if FParent <> Nil then begin
    for Result := 0 to FParent.FChilds.Count - 1 do begin
      if FParent.FChilds[Result] = Self then
        Exit;
    end;
  end;
  Result := -1;
end;

procedure TXpgDOMNode.Insert(ANode: TXpgDOMNode);
begin
  ANode.FParent := Self;
  FChilds.Add(ANode);
end;

function TXpgDOMNode.NextSibling: TXpgDOMNode;
var
  i: integer;
begin
  i := Index;
  if (i >= 0) and (i < (FParent.FChilds.Count - 1)) then
    Result := FParent[i + 1]
  else
    Result := Nil;
end;

function TXpgDOMNode.PrevSibling: TXpgDOMNode;
var
  i: integer;
begin
  i := Index;
  if i > 0 then
    Result := FParent[i - 1]
  else
    Result := Nil;
end;

procedure TXpgDOMNode.SetAsBoolean(const Value: boolean);
begin
  FContent := XmlBoolToStr(Value);
end;

procedure TXpgDOMNode.SetAsDateTime(const Value: TDateTime);
begin
  FContent := XmlDateTimeToStr(Value);
end;

procedure TXpgDOMNode.SetAsFloat(const Value: extended);
begin
  FContent := XmlFloatToStr(Value);
end;

procedure TXpgDOMNode.SetAsInteger(const Value: integer);
begin
  FContent := XmlIntToStr(Value);
end;

procedure TXpgDOMNode.SetAsString(const Value: AxUCString);
begin
  FContent := Value;
end;

function TXpgDOMNode.AddChild(const AQName: AxUCString; AContent: double): TXpgDOMNode;
begin
  Result := CreateNode(Self);
  Result.QName := AQName;
  Result.Content := XmlFloatToStr(AContent);
  FChilds.Add(Result);
end;

{ TXpgSimpleDOM }

procedure TXpgSimpleDOM.BeginTag;
var
  i: integer;
  S: AxUCString;
  Node: TXpgDOMNode;
begin
  Node := TXpgDOMNode(FStack.Peek);
  Node := Node.AddChild(QName);
  if HasText then begin
    if FPreserveSpaces then
      Node.Content := Text
    else
      Node.Content := Trim(Text);
  end;
  for i := 0 to Attributes.Count - 1 do begin
    S := Attributes.Values[i];
    HtmlCheckText(S);
    Node.AddAttribute(Attributes[i],S);
  end;

  FStack.Push(Node);
end;

procedure TXpgSimpleDOM.Clear;
begin
  inherited Clear;
  while FStack.Count > 0 do
    FStack.Pop;
  FRoot.Clear;
end;

constructor TXpgSimpleDOM.Create;
begin
  inherited Create;

  FStack := TObjectStack.Create;
  FRoot := CreateRoot;
  FRoot.QName := '__ROOT__';
end;

function TXpgSimpleDOM.CreateRoot: TXpgDOMNode;
begin
  Result := TXpgDOMNode.Create(Nil);
end;

destructor TXpgSimpleDOM.Destroy;
begin
  Clear;
  FStack.Free;
  FRoot.Free;
  inherited;
end;

function TXpgSimpleDOM.Empty: boolean;
begin
  Result := FRoot.Count <= 0;
end;

procedure TXpgSimpleDOM.EndTag;
begin
  FStack.Pop;
end;

function TXpgSimpleDOM.LoadFromBuffer(ABuf: PByteArray; AMaxSize: integer): integer;
var
  i: integer;
  S: AxUCString;
begin
  Result := 0;

  while (Result < AMaxSize) and (ABuf[Result] in [10,13,32..126]) do
    Inc(Result);

  if Result <= AMaxSize then begin
    SetLength(S,Result);

    for i := 0 to Result - 1 do
      S[i + 1] := AXUCChar(ABuf[i]);

    LoadFromString(S);
  end;
end;

function TXpgSimpleDOM.LoadFromFile(const AFilename: string): integer;
var
  Stream: TFileStream;
begin
  Stream := TFileStream.Create(AFilename,fmOpenRead);
  try
    Result := LoadFromStream(Stream);
  finally
    Stream.Free;
  end;
end;

function TXpgSimpleDOM.LoadFromStream(AStream: TStream): integer;
begin
  Clear;
  FStack.Push(FRoot);
  try
    Result := inherited LoadFromStream(AStream);
  finally
    FStack.Pop;
  end;
end;

function TXpgSimpleDOM.LoadFromStreamHtml(AStream: TStream): integer;
begin
  Clear;
  FStack.Push(FRoot);
  try
    Result := inherited LoadFromStreamHtml(AStream);
  finally
    if FStack.Count > 0 then
      FStack.Pop;
  end;
end;

procedure TXpgSimpleDOM.LoadFromString(AStr: AxUCString);
var
  Stream: TStringStream;
begin
  Clear;
  Stream := TStringStream.Create('');
  try
    Stream.WriteString(AStr);
    Stream.Seek(0,0);
    LoadFromStream(Stream);
  finally
    Stream.Free;
  end;
end;

procedure TXpgSimpleDOM.SaveToFile(const AFilename: string);
var
  Stream: TFileStream;
begin
  Stream := TFileStream.Create(AFilename,fmCreate);
  try
    SaveToStream(Stream);
  finally
    Stream.Free;
  end;
end;

procedure TXpgSimpleDOM.SaveToStream(AStream: TStream);
var
  Writer: TXpgWriteXML;

procedure WriteNode(ANode: TXpgDOMNode);
var
  i: integer;
begin
  for i := 0 to ANode.Attributes.Count - 1 do
    Writer.AddAttribute(ANode.Attributes[i].Name,ANode.Attributes[i].Value);

  if ANode.Count > 0 then begin
    Writer.BeginTag(ANode.QName);
  end
  else begin
    Writer.Text := ANode.Content;
    Writer.SimpleTag(ANode.QName);
  end;

  for i := 0 to ANode.Count - 1 do
    WriteNode(ANode[i]);

  if ANode.Count > 0 then
    Writer.EndTag;
end;

begin
  Writer := TXpgWriteXML.Create;
  try
    Writer.SaveToStream(AStream);
    if FRoot.Count > 0 then
      WriteNode(FRoot[0]);
  finally
    Writer.Free;
  end;
end;

function TXpgSimpleDOM.SaveToString: AxUCString;
var
  Stream: TStringStream;
begin
  Stream := TStringStream.Create('');
  try
    SaveToStream(Stream);
    Result := Stream.DataString;
  finally
    Stream.Free;
  end;
end;

{ TXpgDOMAttributes }

function TXpgDOMAttributes.Add(const AName, AValue: AxUCString): TXpgDOMAttribute;
begin
  Result := TXpgDOMAttribute.Create;
  Result.Name := AName;
  Result.Value := AValue;
  FItems.Add(Result);
end;

function TXpgDOMAttributes.Add(const AName: AxUCString; AValue: integer): TXpgDOMAttribute;
begin
  Result := TXpgDOMAttribute.Create;
  Result.Name := AName;
  Result.Value := XmlIntToStr(AValue);
  FItems.Add(Result);
end;

function TXpgDOMAttributes.Add(const AName: AxUCString; AValue: boolean): TXpgDOMAttribute;
begin
  Result := TXpgDOMAttribute.Create;
  Result.Name := AName;
  if AValue then
    Result.Value := '1'
  else
    Result.Value := '0';
  FItems.Add(Result);
end;

function TXpgDOMAttributes.Add(const AName: AxUCString; AValue: double): TXpgDOMAttribute;
begin
  Result := TXpgDOMAttribute.Create;
  Result.Name := AName;
  Result.Value := XmlFloatToStr(AValue);
  FItems.Add(Result);
end;

procedure TXpgDOMAttributes.Assign(ASource: TXpgDOMAttributes);
var
  i: integer;
begin
  Clear;
  for i := 0 to ASource.FItems.Count - 1 do
    Add(ASource.Items[i].Name,ASource.Items[i].Value);
end;

function TXpgDOMAttributes.AsText: AxUCString;
var
  i: integer;
begin
  Result := '';
  for i := 0 to Count - 1 do
    Result := Result + Format('%s="%s" ',[Items[i].Name,Items[i].Value]);
end;

procedure TXpgDOMAttributes.Clear;
begin
  FItems.Clear;
end;

function TXpgDOMAttributes.Clone: TXpgDOMAttributes;
begin
  Result := TXpgDOMAttributes.Create;
  Result.Assign(Self);
end;

function TXpgDOMAttributes.Count: integer;
begin
  Result := FItems.Count;
end;

constructor TXpgDOMAttributes.Create;
begin
  FItems := TObjectList.Create;
end;

destructor TXpgDOMAttributes.Destroy;
begin
  FItems.Free;
  inherited;
end;

function TXpgDOMAttributes.EqualValue(const AName, AMatchPrefix, AValue: AxUCString): boolean;
var
  i: integer;
  S: AxUCString;
  List: TStringList;
begin
  Result := False;
  S := FindValue(AName);
  if S <> '' then begin
    List := TStringList.Create;
    try
      List.Delimiter := ' ';
      List.DelimitedText := S;
      for i := 0 to List.Count - 1 do begin
        if (AMatchPrefix + List[i]) = AValue then begin
          Result := True;
          Exit;
        end;
      end;
    finally
      List.Free;
    end;
  end;
end;

function TXpgDOMAttributes.Find(const AName: AxUCString): TXpgDOMAttribute;
var
  i: integer;
begin
  for i := 0 to Count - 1 do begin
    if Items[i].Name = AName then begin
      Result := Items[i];
      Exit;
    end;
  end;
  Result := Nil;
end;

function TXpgDOMAttributes.FindValue(const AName: AxUCString): AxUCString;
var
  A: TXpgDOMAttribute;
begin
  A := Find(AName);
  if A <> Nil then
    Result := A.Value
  else
    Result := '';
end;

function TXpgDOMAttributes.FindValueBool(const AName: AxUCString): boolean;
var
  A: TXpgDOMAttribute;
begin
  A := Find(AName);
  if A <> Nil then
    Result := (Lowercase(A.Value) = 'true') or (A.Value = '1')
  else
    Result := False;
end;

function TXpgDOMAttributes.FindValueDef(const AName, ADefault: AxUCString): AxUCString;
var
  A: TXpgDOMAttribute;
begin
  A := Find(AName);
  if A <> Nil then
    Result := A.Value
  else
    Result := ADefault;
end;

function TXpgDOMAttributes.FindValueFloatDef(const AName: AxUCString; const ADefault: integer): double;
var
  A: TXpgDOMAttribute;
begin
  A := Find(AName);
  if A <> Nil then
    Result := XmlStrToFloat(A.Value)
  else
    Result := ADefault;
end;

function TXpgDOMAttributes.FindValueHexDef(const AName: AxUCString; const ADefault: integer): integer;
var
  A: TXpgDOMAttribute;
begin
  A := Find(AName);
  if A <> Nil then
    Result := XmlStrHexToIntDef('#' + A.Value,ADefault)
  else
    Result := ADefault;
end;

function TXpgDOMAttributes.FindValueInt(const AName: AxUCString): integer;
var
  A: TXpgDOMAttribute;
begin
  A := Find(AName);
  if A <> Nil then
    Result := XmlStrToIntDef(A.Value,0)
  else
    Result := 0;
end;

function TXpgDOMAttributes.FindValueIntDef(const AName: AxUCString; const ADefault: integer): integer;
var
  A: TXpgDOMAttribute;
begin
  A := Find(AName);
  if A <> Nil then
    Result := XmlStrToIntDef(A.Value,ADefault)
  else
    Result := ADefault;
end;

function TXpgDOMAttributes.GetItems(Index: integer): TXpgDOMAttribute;
begin
  Result := TXpgDOMAttribute(FItems.Items[Index]);
end;

function TXpgDOMAttributes.Update(const AName: AxUCString; AValue: boolean): TXpgDOMAttribute;
begin
  Result := Find(AName);
  if Result <> Nil then begin
    if AValue then
      Result.Value := '1'
    else
      Result.Value := '0';
  end
  else
    Result := Add(AName,AValue);
end;

function TXpgDOMAttributes.Update(const AName: AxUCString; AValue: integer): TXpgDOMAttribute;
begin
  Result := Find(AName);
  if Result <> Nil then
    Result.Value := XmlIntToStr(AValue)
  else
    Result := Add(AName,AValue);
end;

function TXpgDOMAttributes.Update(const AName, AValue: AxUCString): TXpgDOMAttribute;
begin
  Result := Find(AName);
  if Result <> Nil then
    Result.Value := AValue
  else
    Result := Add(AName,AValue);
end;

{ TXpgDOMNodeList }

constructor TXpgDOMNodeList.Create;
begin
  inherited Create(False);
end;

function TXpgDOMNodeList.GetItems(Index: integer): TXpgDOMNode;
begin
  Result := TXpgDOMNode(inherited Items[Index]);
end;

{ TXpgDOMAttribute }

function TXpgDOMAttribute.GetAsBase64: AnsiString;
var
  i: integer;
  S: AnsiString;
begin
  SetLength(S,Length(FValue));

  for i := 1 to Length(FValue) do
    S[i] := AnsiChar(FValue[i]);

  Result := XPGDecodeBase64(S);
end;

function TXpgDOMAttribute.GetAsInteger: integer;
begin
  Result := StrToIntDef(FValue,0);
end;

end.
