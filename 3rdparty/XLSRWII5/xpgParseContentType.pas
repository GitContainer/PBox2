unit xpgParseContentType;

// Copyright (c) 2010,2011 Axolot Data
// Web : http://www.axolot.com/xpg
// Mail: xpg@axolot.com
//
// X   X  PPP    GGG
//  X X   P  P  G
//   X    PPP   G  GG
//  X X   P     G   G
// X   X  P      GGG
//
// File generated with Axolot XPG, Xml Parser Generator.
// Version 0.00.90.
// File created on 2011-12-03 17:17:56

{$I AxCompilers.inc}

{$MINENUMSIZE 2}
{$BOOLEVAL OFF}
{$RANGECHECKS OFF}

interface

uses Classes, SysUtils, Contnrs, IniFiles, xpgPUtils, xpgPLists, xpgPXMLUtils,
XLSUtils5,
xpgPXML;

type TXPGDocBase = class(TObject)
protected
     FErrors: TXpgPErrors;
public
     property Errors: TXpgPErrors read FErrors;
     end;

     TXPGBase = class(TObject)
protected
     FOwner: TXPGDocBase;
     FElementCount: integer;
     FAttributeCount: integer;
     FAssigneds: TXpgAssigneds;
     function  CheckAssigned: integer; virtual;
     function  Assigned: boolean;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; virtual;
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); virtual;
     procedure AfterTag; virtual;

     class procedure AddEnums;
     class function  StrToEnum(AValue: AxUCString): integer;
     class function  StrToEnumDef(AValue: AxUCString; ADefault: integer): integer;
     class function  TryStrToEnum(AValue: AxUCString; AText: AxUCString; AEnumNames: array of AxUCString; APtrInt: PInteger): boolean;
public
     function  Available: boolean;

     property ElementCount: integer read FElementCount write FElementCount;
     property AttributeCount: integer read FAttributeCount write FAttributeCount;
     property Assigneds: TXpgAssigneds read FAssigneds write FAssigneds;
     end;

     TXPGBaseObjectList = class(TObjectList)
protected
     FOwner: TXPGDocBase;
     FAssigned: boolean;
     function  GetItems(Index: integer): TXPGBase;
public
     constructor Create(AOwner: TXPGDocBase);
     property Items[Index: integer]: TXPGBase read GetItems;
     end;

     TXPGReader = class(TXpgReadXML)
protected
     FCurrent: TXPGBase;
     FStack: TObjectStack;
     FErrors: TXpgPErrors;
public
     constructor Create(AErrors: TXpgPErrors; ARoot: TXPGBase);
     destructor Destroy; override;

     procedure BeginTag; override;
     procedure EndTag; override;
     end;

     TCT_Default = class(TXPGBase)
protected
     FExtension: AxUCString;
     FContentType: AxUCString;

     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Extension: AxUCString read FExtension write FExtension;
     property ContentType: AxUCString read FContentType write FContentType;
     end;

     TCT_DefaultXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_Default;
public
     function  Add: TCT_Default;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_Default read GetItems; default;
     end;

     TCT_Override = class(TXPGBase)
protected
     FContentType: AxUCString;
     FPartName: AxUCString;

     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property ContentType: AxUCString read FContentType write FContentType;
     property PartName: AxUCString read FPartName write FPartName;
     end;

     TCT_OverrideXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_Override;
public
     function  Add: TCT_Override;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_Override read GetItems; default;
     end;

     TCT_Types = class(TXPGBase)
protected
     FDefaultXpgList: TCT_DefaultXpgList;
     FOverrideXpgList: TCT_OverrideXpgList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property DefaultXpgList: TCT_DefaultXpgList read FDefaultXpgList;
     property OverrideXpgList: TCT_OverrideXpgList read FOverrideXpgList;
     end;

     T__ROOT__ = class(TXPGBase)
protected
     FRootAttributes: TStringXpgList;
     FDefault: TCT_Default;
     FOverride: TCT_Override;
     FTypes: TCT_Types;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property RootAttributes: TStringXpgList read FRootAttributes;
     property Default: TCT_Default read FDefault;
     property Override: TCT_Override read FOverride;
     property Types: TCT_Types read FTypes;
     end;

     TXPGDocContentType = class(TXPGDocBase)
protected
     FRoot: T__ROOT__;
     FReader: TXPGReader;
     FWriter: TXpgWriteXML;

     function  GetDefault: TCT_Default;
     function  GetOverride: TCT_Override;
     function  GetTypes: TCT_Types;
public
     constructor Create;
     destructor Destroy; override;

     procedure LoadFromFile(AFilename: AxUCString);
     procedure LoadFromStream(AStream: TStream);
     procedure SaveToFile(AFilename: AxUCString);

     procedure SaveToStream(AStream: TStream);

     property Root: T__ROOT__ read FRoot;
     property Default: TCT_Default read GetDefault;
     property Override: TCT_Override read GetOverride;
     property Types: TCT_Types read GetTypes;
     end;


implementation

{$ifdef DELPHI_5}
var FEnums: TStringList;
{$else}
var FEnums: THashedStringList;
{$endif}

{ TXPGBase }

function  TXPGBase.CheckAssigned: integer;
begin
  Result := 1;
end;

function  TXPGBase.Assigned: boolean;
begin
  Result := FAssigneds <> [];
end;

function  TXPGBase.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
end;

procedure TXPGBase.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
end;

procedure TXPGBase.AfterTag;
begin
end;

class procedure TXPGBase.AddEnums;
begin
end;

class function  TXPGBase.StrToEnum(AValue: AxUCString): integer;
var
  i: integer;
begin
  i := FEnums.IndexOf(AValue);
  if i >= 0 then
    Result := Integer(FEnums.Objects[i])
  else
    Result := 0;
end;

class function  TXPGBase.StrToEnumDef(AValue: AxUCString; ADefault: integer): integer;
var
  i: integer;
begin
  i := FEnums.IndexOf(AValue);
  if i >= 0 then 
    Result := Integer(FEnums.Objects[i])
  else 
    Result := ADefault;
end;

class function  TXPGBase.TryStrToEnum(AValue: AxUCString; AText: AxUCString; AEnumNames: array of AxUCString; APtrInt: PInteger): boolean;
var
  i: integer;
begin
  i := FEnums.IndexOf(AValue);
  if i >= 0 then 
  begin
    i := Integer(FEnums.Objects[i]);
    Result := (i <= High(AEnumNames)) and (AText = AEnumNames[i]);
    if Result then 
      APtrInt^ := i;
  end
  else 
    Result := False;
end;

function  TXPGBase.Available: boolean;
begin
  Result := xaRead in FAssigneds;
end;

{ TXPGBaseObjectList }

function  TXPGBaseObjectList.GetItems(Index: integer): TXPGBase;
begin
  Result := TXPGBase(inherited Items[Index]);
end;

constructor TXPGBaseObjectList.Create(AOwner: TXPGDocBase);
begin
  inherited Create;
  FOwner := AOwner;
end;

{ TXPGReader }

constructor TXPGReader.Create(AErrors: TXpgPErrors; ARoot: TXPGBase);
begin
  inherited Create;
  FErrors := AErrors;
  FErrors.NoDuplicates := True;
  FCurrent := ARoot;
  FStack := TObjectStack.Create;
end;

destructor TXPGReader.Destroy;
begin
  FStack.Free;
  inherited Destroy;
end;

procedure TXPGReader.BeginTag;
begin
  FStack.Push(FCurrent);
  FCurrent := FCurrent.HandleElement(Self);
  FCurrent.AssignAttributes(Attributes);
end;

procedure TXPGReader.EndTag;
begin
  FCurrent := TXPGBase(FStack.Pop);
  FCurrent.AfterTag;
end;

{ TCT_Default }

function  TCT_Default.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FExtension <> '' then 
    Inc(AttrsAssigned);
  if FContentType <> '' then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_Default.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_Default.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('Extension',FExtension);
  AWriter.AddAttribute('ContentType',FContentType);
end;

procedure TCT_Default.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
  S: AxUCString;
begin
  for i := 0 to AAttributes.Count - 1 do 
  begin
    S := AAttributes[i];
    case CalcHash_A(S) of
      $000003BD: FExtension := AAttributes.Values[i];
      $0000047D: FContentType := AAttributes.Values[i];
      else 
        FOwner.Errors.Error(xemUnknownAttribute,S);
    end;
  end
end;

constructor TCT_Default.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 2;
end;

destructor TCT_Default.Destroy;
begin
end;

procedure TCT_Default.Clear;
begin
  FAssigneds := [];
  FExtension := '';
  FContentType := '';
end;

{ TCT_DefaultXpgList }

function  TCT_DefaultXpgList.GetItems(Index: integer): TCT_Default;
begin
  Result := TCT_Default(inherited Items[Index]);
end;

function  TCT_DefaultXpgList.Add: TCT_Default;
begin
  Result := TCT_Default.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_DefaultXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_DefaultXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

{ TCT_Override }

function  TCT_Override.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FContentType <> '' then 
    Inc(AttrsAssigned);
  if FPartName <> '' then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_Override.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_Override.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('ContentType',FContentType);
  AWriter.AddAttribute('PartName',FPartName);
end;

procedure TCT_Override.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
  S: AxUCString;
begin
  for i := 0 to AAttributes.Count - 1 do 
  begin
    S := AAttributes[i];
    case CalcHash_A(S) of
      $0000047D: FContentType := AAttributes.Values[i];
      $00000318: FPartName := AAttributes.Values[i];
      else 
        FOwner.Errors.Error(xemUnknownAttribute,S);
    end;
  end
end;

constructor TCT_Override.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 2;
end;

destructor TCT_Override.Destroy;
begin
end;

procedure TCT_Override.Clear;
begin
  FAssigneds := [];
  FContentType := '';
  FPartName := '';
end;

{ TCT_OverrideXpgList }

function  TCT_OverrideXpgList.GetItems(Index: integer): TCT_Override;
begin
  Result := TCT_Override(inherited Items[Index]);
end;

function  TCT_OverrideXpgList.Add: TCT_Override;
begin
  Result := TCT_Override.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_OverrideXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_OverrideXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

{ TCT_Types }

function  TCT_Types.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FDefaultXpgList.CheckAssigned);
  Inc(ElemsAssigned,FOverrideXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_Types.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  case CalcHash_A(QName) of
    $000002C5: Result := FDefaultXpgList.Add;
    $00000340: Result := FOverrideXpgList.Add;
    else 
      FOwner.Errors.Error(xemUnknownElement,QName);
  end;
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_Types.Write(AWriter: TXpgWriteXML);
begin
  FDefaultXpgList.Write(AWriter,'Default');
  FOverrideXpgList.Write(AWriter,'Override');
end;

constructor TCT_Types.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 2;
  FAttributeCount := 0;
  FDefaultXpgList := TCT_DefaultXpgList.Create(FOwner);
  FOverrideXpgList := TCT_OverrideXpgList.Create(FOwner);
end;

destructor TCT_Types.Destroy;
begin
  FDefaultXpgList.Free;
  FOverrideXpgList.Free;
end;

procedure TCT_Types.Clear;
begin
  FAssigneds := [];
  FDefaultXpgList.Clear;
  FOverrideXpgList.Clear;
end;

{ T__ROOT__ }

function  T__ROOT__.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FDefault.CheckAssigned);
  Inc(ElemsAssigned,FOverride.CheckAssigned);
  Inc(ElemsAssigned,FTypes.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  T__ROOT__.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  i: integer;
  QName: AxUCString;
begin
  for i := 0 to AReader.Attributes.Count - 1 do 
    FRootAttributes.Add(AReader.Attributes.AsXmlText2(i));
  Result := Self;
  QName := AReader.QName;
  case CalcHash_A(QName) of
    $000002C5: Result := FDefault;
    $00000340: Result := FOverride;
    $00000215: Result := FTypes;
    else 
      FOwner.Errors.Error(xemUnknownElement,QName);
  end;
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure T__ROOT__.Write(AWriter: TXpgWriteXML);
begin
  AWriter.Attributes := FRootAttributes.Text;  // TODO
  if FDefault.Assigned then
  begin
    FDefault.WriteAttributes(AWriter);
    AWriter.SimpleTag('Default');
  end;
  if FOverride.Assigned then 
  begin
    FOverride.WriteAttributes(AWriter);
    AWriter.SimpleTag('Override');
  end;
  if FTypes.Assigned then 
    if xaElements in FTypes.FAssigneds then 
    begin
      AWriter.BeginTag('Types');
      FTypes.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('Types');
end;

constructor T__ROOT__.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FRootAttributes := TStringXpgList.Create;
  FElementCount := 3;
  FAttributeCount := 0;
  FDefault := TCT_Default.Create(FOwner);
  FOverride := TCT_Override.Create(FOwner);
  FTypes := TCT_Types.Create(FOwner);
end;

destructor T__ROOT__.Destroy;
begin
  FRootAttributes.Free;
  FDefault.Free;
  FOverride.Free;
  FTypes.Free;
end;

procedure T__ROOT__.Clear;
begin
  FRootAttributes.Clear;
  FAssigneds := [];
  FDefault.Clear;
  FOverride.Clear;
  FTypes.Clear;
end;

{ TTXPGDocContentType }

function  TXPGDocContentType.GetDefault: TCT_Default;
begin
  Result := FRoot.Default;
end;

function  TXPGDocContentType.GetOverride: TCT_Override;
begin
  Result := FRoot.Override;
end;

function  TXPGDocContentType.GetTypes: TCT_Types;
begin
  Result := FRoot.Types;
end;

constructor TXPGDocContentType.Create;
begin
  FRoot := T__ROOT__.Create(Self);
  FErrors := TXpgPErrors.Create;
  FReader := TXPGReader.Create(FErrors,FRoot);
  FWriter := TXpgWriteXML.Create;
end;

destructor TXPGDocContentType.Destroy;
begin
  FRoot.Free;
  FReader.Free;
  FWriter.Free;
  FErrors.Free;
  inherited Destroy;
end;

procedure TXPGDocContentType.LoadFromFile(AFilename: AxUCString);
var
  Stream: TFileStream;
begin
  Stream := TFileStream.Create(AFilename,fmOpenRead);
  try
    FReader.LoadFromStream(Stream);
  finally
    Stream.Free;
  end;
end;

procedure TXPGDocContentType.LoadFromStream(AStream: TStream);
begin
  FReader.LoadFromStream(AStream);
end;

procedure TXPGDocContentType.SaveToFile(AFilename: AxUCString);
begin
  FWriter.SaveToFile(AFilename);
  FRoot.CheckAssigned;
  FRoot.Write(FWriter);
end;

procedure TXPGDocContentType.SaveToStream(AStream: TStream);
begin
  FWriter.SaveToStream(AStream);
  FRoot.CheckAssigned;
  FRoot.Write(FWriter);
end;

initialization
{$ifdef DELPHI_5}
  FEnums := TStringList.Create;
{$else}
  FEnums := THashedStringList.Create;
{$endif}
  TXPGBase.AddEnums;

finalization
  FEnums.Free;

end.
