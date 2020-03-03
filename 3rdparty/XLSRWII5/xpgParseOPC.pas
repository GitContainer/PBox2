unit xpgParseOPC;

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
// File created on 2011-12-03 17:18:30

{$I AxCompilers.inc}

{$MINENUMSIZE 2}
{$BOOLEVAL OFF}
{$RANGECHECKS OFF}

interface

uses Classes, SysUtils, Contnrs, IniFiles, xpgPUtils, xpgPLists, xpgPXMLUtils,
XLSUtils5,
xpgPXML;

type TST_TargetMode =  (sttmExternal,sttmInternal);
const StrTST_TargetMode: array[0..1] of AxUCString = ('External','Internal');

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

     TCT_Relationship = class(TXPGBase)
protected
     FTargetMode: TST_TargetMode;
     FTarget: AxUCString;
     FType: AxUCString;
     FId: AxUCString;
     FContent: AxUCString;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property TargetMode: TST_TargetMode read FTargetMode write FTargetMode;
     property Target: AxUCString read FTarget write FTarget;
     property Type_: AxUCString read FType write FType;
     property Id: AxUCString read FId write FId;
     property Content: AxUCString read FContent write FContent;
     end;

     TCT_RelationshipXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_Relationship;
public
     function  Add: TCT_Relationship;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_Relationship read GetItems; default;
     end;

     TCT_Relationships = class(TXPGBase)
protected
     FRelationshipXpgList: TCT_RelationshipXpgList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property RelationshipXpgList: TCT_RelationshipXpgList read FRelationshipXpgList;
     end;

     T__ROOT__ = class(TXPGBase)
protected
     FRootAttributes: TStringXpgList;
     FRelationships: TCT_Relationships;
     FRelationship: TCT_Relationship;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property RootAttributes: TStringXpgList read FRootAttributes;
     property Relationships: TCT_Relationships read FRelationships;
     property Relationship: TCT_Relationship read FRelationship;
     end;

     TXPGDocOPC = class(TXPGDocBase)
protected
     FRoot: T__ROOT__;
     FReader: TXPGReader;
     FWriter: TXpgWriteXML;

     function  GetRelationships: TCT_Relationships;
     function  GetRelationship: TCT_Relationship;
public
     constructor Create;
     destructor Destroy; override;

     procedure LoadFromFile(AFilename: AxUCString);
     procedure LoadFromStream(AStream: TStream);
     procedure SaveToFile(AFilename: AxUCString);

     procedure SaveToStream(AStream: TStream);

     property Root: T__ROOT__ read FRoot;
     property Relationships: TCT_Relationships read GetRelationships;
     property Relationship: TCT_Relationship read GetRelationship;
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
  FEnums.AddObject('sttmExternal',TObject(0));
  FEnums.AddObject('sttmInternal',TObject(1));
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

{ TCT_Relationship }

function  TCT_Relationship.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if Integer(FTargetMode) <> XPG_UNKNOWN_ENUM then 
    Inc(AttrsAssigned);
  if FTarget <> '' then 
    Inc(AttrsAssigned);
  if FType <> '' then 
    Inc(AttrsAssigned);
  if FId <> '' then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
  if FContent <> '' then 
  begin
    FAssigneds := FAssigneds + [xaContent];
    Inc(Result);
  end;
end;

function  TCT_Relationship.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_Relationship.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_Relationship.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if Integer(FTargetMode) <> XPG_UNKNOWN_ENUM then 
    AWriter.AddAttribute('TargetMode',StrTST_TargetMode[Integer(FTargetMode)]);
  AWriter.AddAttribute('Target',FTarget);
  AWriter.AddAttribute('Type',FType);
  AWriter.AddAttribute('Id',FId);
end;

procedure TCT_Relationship.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
  S: AxUCString;
begin
  for i := 0 to AAttributes.Count - 1 do 
  begin
    S := AAttributes[i];
    case CalcHash_A(S) of
      $000003EC: FTargetMode := TST_TargetMode(StrToEnum('sttm' + AAttributes.Values[i]));
      $00000267: FTarget := AAttributes.Values[i];
      $000001A2: FType := AAttributes.Values[i];
      $000000AD: FId := AAttributes.Values[i];
      else 
        FOwner.Errors.Error(xemUnknownAttribute,S);
    end;
  end
end;

constructor TCT_Relationship.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 4;
  FTargetMode := TST_TargetMode(XPG_UNKNOWN_ENUM);
end;

destructor TCT_Relationship.Destroy;
begin
end;

procedure TCT_Relationship.Clear;
begin
  FAssigneds := [];
  FTargetMode := TST_TargetMode(XPG_UNKNOWN_ENUM);
  FTarget := '';
  FType := '';
  FId := '';
end;

{ TCT_RelationshipXpgList }

function  TCT_RelationshipXpgList.GetItems(Index: integer): TCT_Relationship;
begin
  Result := TCT_Relationship(inherited Items[Index]);
end;

function  TCT_RelationshipXpgList.Add: TCT_Relationship;
begin
  Result := TCT_Relationship.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_RelationshipXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_RelationshipXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaContent in Items[i].FAssigneds then 
      AWriter.Text := GetItems(i).Content;
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

{ TCT_Relationships }

function  TCT_Relationships.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FRelationshipXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_Relationships.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'Relationship' then 
  begin
    Result := FRelationshipXpgList.Add;
    if AReader.HasText then 
      TCT_Relationship(Result).Content := AReader.Text;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_Relationships.Write(AWriter: TXpgWriteXML);
begin
  FRelationshipXpgList.Write(AWriter,'Relationship');
end;

constructor TCT_Relationships.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 0;
  FRelationshipXpgList := TCT_RelationshipXpgList.Create(FOwner);
end;

destructor TCT_Relationships.Destroy;
begin
  FRelationshipXpgList.Free;
end;

procedure TCT_Relationships.Clear;
begin
  FAssigneds := [];
  FRelationshipXpgList.Clear;
end;

{ T__ROOT__ }

function  T__ROOT__.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FRelationships.CheckAssigned);
  Inc(ElemsAssigned,FRelationship.CheckAssigned);
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
    $00000565: Result := FRelationships;
    $000004F2: begin
      Result := FRelationship;
      if AReader.HasText then 
        TCT_Relationship(Result).Content := AReader.Text;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,QName);
  end;
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure T__ROOT__.Write(AWriter: TXpgWriteXML);
begin
  AWriter.Attributes := FRootAttributes.Text;
  if FRelationships.Assigned then 
    if xaElements in FRelationships.FAssigneds then 
    begin
      AWriter.BeginTag('Relationships');
      FRelationships.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('Relationships');
  if FRelationship.Assigned then 
  begin
    AWriter.Text := FRelationship.Content;
    FRelationship.WriteAttributes(AWriter);
    AWriter.SimpleTag('Relationship');
  end;
end;

constructor T__ROOT__.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FRootAttributes := TStringXpgList.Create;
  FElementCount := 2;
  FAttributeCount := 0;
  FRelationships := TCT_Relationships.Create(FOwner);
  FRelationship := TCT_Relationship.Create(FOwner);
end;

destructor T__ROOT__.Destroy;
begin
  FRootAttributes.Free;
  FRelationships.Free;
  FRelationship.Free;
end;

procedure T__ROOT__.Clear;
begin
  FRootAttributes.Clear;
  FAssigneds := [];
  FRelationships.Clear;
  FRelationship.Clear;
end;

{ TXPGDocOPC }

function  TXPGDocOPC.GetRelationships: TCT_Relationships;
begin
  Result := FRoot.Relationships;
end;

function  TXPGDocOPC.GetRelationship: TCT_Relationship;
begin
  Result := FRoot.Relationship;
end;

constructor TXPGDocOPC.Create;
begin
  FRoot := T__ROOT__.Create(Self);
  FErrors := TXpgPErrors.Create;
  FReader := TXPGReader.Create(FErrors,FRoot);
  FWriter := TXpgWriteXML.Create;
end;

destructor TXPGDocOPC.Destroy;
begin
  FRoot.Free;
  FReader.Free;
  FWriter.Free;
  FErrors.Free;
  inherited Destroy;
end;

procedure TXPGDocOPC.LoadFromFile(AFilename: AxUCString);
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

procedure TXPGDocOPC.LoadFromStream(AStream: TStream);
begin
  FReader.LoadFromStream(AStream);
end;

procedure TXPGDocOPC.SaveToFile(AFilename: AxUCString);
begin
  FWriter.SaveToFile(AFilename);
  FRoot.CheckAssigned;
  FRoot.Write(FWriter);
end;

procedure TXPGDocOPC.SaveToStream(AStream: TStream);
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
