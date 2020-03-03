unit xpgParseDocPropsApp;

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
// File created on 2011-12-03 17:16:44

{$I AxCompilers.inc}

{$MINENUMSIZE 2}
{$BOOLEVAL OFF}
{$RANGECHECKS OFF}

interface

uses Classes, SysUtils, Contnrs, IniFiles, xpgPUtils, xpgPLists, xpgPXMLUtils,
XLSUtils5,
xpgPXML;

type TST_ArrayBaseType =  (stabtVariant,stabtI1,stabtI2,stabtI4,stabtInt,stabtUi1,stabtUi2,stabtUi4,stabtUint,stabtR4,stabtR8,stabtDecimal,stabtBstr,stabtDate,stabtBool,stabtCy,stabtError);
const StrTST_ArrayBaseType: array[0..16] of AxUCString = ('variant','i1','i2','i4','int','ui1','ui2','ui4','uint','r4','r8','decimal','bstr','date','bool','cy','error');
type TST_VectorBaseType =  (stvbtVariant,stvbtI1,stvbtI2,stvbtI4,stvbtI8,stvbtUi1,stvbtUi2,stvbtUi4,stvbtUi8,stvbtR4,stvbtR8,stvbtLpstr,stvbtLpwstr,stvbtBstr,stvbtDate,stvbtFiletime,stvbtBool,stvbtCy,stvbtError,stvbtClsid,stvbtCf);
const StrTST_VectorBaseType: array[0..20] of AxUCString = ('variant','i1','i2','i4','i8','ui1','ui2','ui4','ui8','r4','r8','lpstr','lpwstr','bstr','date','filetime','bool','cy','error','clsid','cf');

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

     TCT_Vector = class;
     TCT_Array = class;

     TCT_Empty = class(TXPGBase)
protected
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     end;

     TCT_Null = class(TXPGBase)
protected
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     end;

     TCT_Vstream = class(TXPGBase)
protected
     FVersion: AxUCString;
     FContent: integer;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Version: AxUCString read FVersion write FVersion;
     property Content: integer read FContent write FContent;
     end;

     TCT_Cf = class(TXPGBase)
protected
     FFormat: AxUCString;
     FContent: integer;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Format: AxUCString read FFormat write FFormat;
     property Content: integer read FContent write FContent;
     end;

     TCT_CfXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_Cf;
public
     function  Add: TCT_Cf;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_Cf read GetItems; default;
     end;

     TCT_Variant = class(TXPGBase)
protected
     FVt_Variant: TCT_Variant;
     FVt_Vector: TCT_Vector;
     FVt_Array: TCT_Array;
     FVt_Blob: longword;
     FVt_Oblob: longword;
     FVt_Empty: TCT_Empty;
     FVt_Null: TCT_Null;
     FVt_I1: longword;
     FVt_I2: longword;
     FVt_I4: longword;
     FVt_I8: longword;
     FVt_Int: longword;
     FVt_Ui1: longword;
     FVt_Ui2: longword;
     FVt_Ui4: longword;
     FVt_Ui8: longword;
     FVt_Uint: longword;
     FVt_R4: extended;
     FVt_R8: extended;
     FVt_Decimal: extended;
     FVt_Lpstr: AxUCString;
     FVt_Lpwstr: AxUCString;
     FVt_Bstr: AxUCString;
     FVt_Date: TDateTime;
     FVt_Filetime: TDateTime;
     FVt_Bool: boolean;
     FVt_Cy: AxUCString;
     FVt_Error: AxUCString;
     FVt_Stream: longword;
     FVt_Ostream: longword;
     FVt_Storage: longword;
     FVt_Ostorage: longword;
     FVt_Vstream: TCT_Vstream;
     FVt_Clsid: AxUCString;
     FVt_Cf: TCT_Cf;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     function  CreateTCT_Variant: TCT_Variant;
     function  CreateTCT_Vector: TCT_Vector;
     function  CreateTCT_Array: TCT_Array;

     property Vt_Variant: TCT_Variant read FVt_Variant;
     property Vt_Vector: TCT_Vector read FVt_Vector write FVt_Vector;
     property Vt_Array: TCT_Array read FVt_Array write FVt_Array;
     property Vt_Blob: longword read FVt_Blob write FVt_Blob;
     property Vt_Oblob: longword read FVt_Oblob write FVt_Oblob;
     property Vt_Empty: TCT_Empty read FVt_Empty;
     property Vt_Null: TCT_Null read FVt_Null;
     property Vt_I1: longword read FVt_I1 write FVt_I1;
     property Vt_I2: longword read FVt_I2 write FVt_I2;
     property Vt_I4: longword read FVt_I4 write FVt_I4;
     property Vt_I8: longword read FVt_I8 write FVt_I8;
     property Vt_Int: longword read FVt_Int write FVt_Int;
     property Vt_Ui1: longword read FVt_Ui1 write FVt_Ui1;
     property Vt_Ui2: longword read FVt_Ui2 write FVt_Ui2;
     property Vt_Ui4: longword read FVt_Ui4 write FVt_Ui4;
     property Vt_Ui8: longword read FVt_Ui8 write FVt_Ui8;
     property Vt_Uint: longword read FVt_Uint write FVt_Uint;
     property Vt_R4: extended read FVt_R4 write FVt_R4;
     property Vt_R8: extended read FVt_R8 write FVt_R8;
     property Vt_Decimal: extended read FVt_Decimal write FVt_Decimal;
     property Vt_Lpstr: AxUCString read FVt_Lpstr write FVt_Lpstr;
     property Vt_Lpwstr: AxUCString read FVt_Lpwstr write FVt_Lpwstr;
     property Vt_Bstr: AxUCString read FVt_Bstr write FVt_Bstr;
     property Vt_Date: TDateTime read FVt_Date write FVt_Date;
     property Vt_Filetime: TDateTime read FVt_Filetime write FVt_Filetime;
     property Vt_Bool: boolean read FVt_Bool write FVt_Bool;
     property Vt_Cy: AxUCString read FVt_Cy write FVt_Cy;
     property Vt_Error: AxUCString read FVt_Error write FVt_Error;
     property Vt_Stream: longword read FVt_Stream write FVt_Stream;
     property Vt_Ostream: longword read FVt_Ostream write FVt_Ostream;
     property Vt_Storage: longword read FVt_Storage write FVt_Storage;
     property Vt_Ostorage: longword read FVt_Ostorage write FVt_Ostorage;
     property Vt_Vstream: TCT_Vstream read FVt_Vstream;
     property Vt_Clsid: AxUCString read FVt_Clsid write FVt_Clsid;
     property Vt_Cf: TCT_Cf read FVt_Cf;
     end;

     TCT_VariantXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_Variant;
public
     function  Add: TCT_Variant;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_Variant read GetItems; default;
     end;

     TCT_Array = class(TXPGBase)
protected
     FLBounds: integer;
     FUBounds: integer;
     FBaseType: TST_ArrayBaseType;
     FVt_VariantXpgList: TCT_VariantXpgList;
     FVt_I1XpgList: TLongwordXpgList;
     FVt_I2XpgList: TLongwordXpgList;
     FVt_I4XpgList: TLongwordXpgList;
     FVt_IntXpgList: TLongwordXpgList;
     FVt_Ui1XpgList: TLongwordXpgList;
     FVt_Ui2XpgList: TLongwordXpgList;
     FVt_Ui4XpgList: TLongwordXpgList;
     FVt_UintXpgList: TLongwordXpgList;
     FVt_R4XpgList: TExtendedXpgList;
     FVt_R8XpgList: TExtendedXpgList;
     FVt_DecimalXpgList: TExtendedXpgList;
     FVt_BstrXpgList: TStringXpgList;
     FVt_DateXpgList: TTDateTimeXpgList;
     FVt_BoolXpgList: TBooleanXpgList;
     FVt_ErrorXpgList: TStringXpgList;
     FVt_CyXpgList: TStringXpgList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     function  CreateTCT_VariantXpgList: TCT_VariantXpgList;

     property LBounds: integer read FLBounds write FLBounds;
     property UBounds: integer read FUBounds write FUBounds;
     property BaseType: TST_ArrayBaseType read FBaseType write FBaseType;
     property Vt_VariantXpgList: TCT_VariantXpgList read FVt_VariantXpgList;
     property Vt_I1XpgList: TLongwordXpgList read FVt_I1XpgList;
     property Vt_I2XpgList: TLongwordXpgList read FVt_I2XpgList;
     property Vt_I4XpgList: TLongwordXpgList read FVt_I4XpgList;
     property Vt_IntXpgList: TLongwordXpgList read FVt_IntXpgList;
     property Vt_Ui1XpgList: TLongwordXpgList read FVt_Ui1XpgList;
     property Vt_Ui2XpgList: TLongwordXpgList read FVt_Ui2XpgList;
     property Vt_Ui4XpgList: TLongwordXpgList read FVt_Ui4XpgList;
     property Vt_UintXpgList: TLongwordXpgList read FVt_UintXpgList;
     property Vt_R4XpgList: TExtendedXpgList read FVt_R4XpgList;
     property Vt_R8XpgList: TExtendedXpgList read FVt_R8XpgList;
     property Vt_DecimalXpgList: TExtendedXpgList read FVt_DecimalXpgList;
     property Vt_BstrXpgList: TStringXpgList read FVt_BstrXpgList;
     property Vt_DateXpgList: TTDateTimeXpgList read FVt_DateXpgList;
     property Vt_BoolXpgList: TBooleanXpgList read FVt_BoolXpgList;
     property Vt_ErrorXpgList: TStringXpgList read FVt_ErrorXpgList;
     property Vt_CyXpgList: TStringXpgList read FVt_CyXpgList;
     end;

     TCT_Vector = class(TXPGBase)
protected
     FBaseType: TST_VectorBaseType;
     FSize: integer;
     FVt_VariantXpgList: TCT_VariantXpgList;
     FVt_I1XpgList: TLongwordXpgList;
     FVt_I2XpgList: TLongwordXpgList;
     FVt_I4XpgList: TLongwordXpgList;
     FVt_I8XpgList: TLongwordXpgList;
     FVt_Ui1XpgList: TLongwordXpgList;
     FVt_Ui2XpgList: TLongwordXpgList;
     FVt_Ui4XpgList: TLongwordXpgList;
     FVt_Ui8XpgList: TLongwordXpgList;
     FVt_R4XpgList: TExtendedXpgList;
     FVt_R8XpgList: TExtendedXpgList;
     FVt_LpstrXpgList: TStringXpgList;
     FVt_LpwstrXpgList: TStringXpgList;
     FVt_BstrXpgList: TStringXpgList;
     FVt_DateXpgList: TTDateTimeXpgList;
     FVt_FiletimeXpgList: TTDateTimeXpgList;
     FVt_BoolXpgList: TBooleanXpgList;
     FVt_CyXpgList: TStringXpgList;
     FVt_ErrorXpgList: TStringXpgList;
     FVt_ClsidXpgList: TStringXpgList;
     FVt_CfXpgList: TCT_CfXpgList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     function  CreateTCT_VariantXpgList: TCT_VariantXpgList;

     property BaseType: TST_VectorBaseType read FBaseType write FBaseType;
     property Size: integer read FSize write FSize;
     property Vt_VariantXpgList: TCT_VariantXpgList read FVt_VariantXpgList;
     property Vt_I1XpgList: TLongwordXpgList read FVt_I1XpgList;
     property Vt_I2XpgList: TLongwordXpgList read FVt_I2XpgList;
     property Vt_I4XpgList: TLongwordXpgList read FVt_I4XpgList;
     property Vt_I8XpgList: TLongwordXpgList read FVt_I8XpgList;
     property Vt_Ui1XpgList: TLongwordXpgList read FVt_Ui1XpgList;
     property Vt_Ui2XpgList: TLongwordXpgList read FVt_Ui2XpgList;
     property Vt_Ui4XpgList: TLongwordXpgList read FVt_Ui4XpgList;
     property Vt_Ui8XpgList: TLongwordXpgList read FVt_Ui8XpgList;
     property Vt_R4XpgList: TExtendedXpgList read FVt_R4XpgList;
     property Vt_R8XpgList: TExtendedXpgList read FVt_R8XpgList;
     property Vt_LpstrXpgList: TStringXpgList read FVt_LpstrXpgList;
     property Vt_LpwstrXpgList: TStringXpgList read FVt_LpwstrXpgList;
     property Vt_BstrXpgList: TStringXpgList read FVt_BstrXpgList;
     property Vt_DateXpgList: TTDateTimeXpgList read FVt_DateXpgList;
     property Vt_FiletimeXpgList: TTDateTimeXpgList read FVt_FiletimeXpgList;
     property Vt_BoolXpgList: TBooleanXpgList read FVt_BoolXpgList;
     property Vt_CyXpgList: TStringXpgList read FVt_CyXpgList;
     property Vt_ErrorXpgList: TStringXpgList read FVt_ErrorXpgList;
     property Vt_ClsidXpgList: TStringXpgList read FVt_ClsidXpgList;
     property Vt_CfXpgList: TCT_CfXpgList read FVt_CfXpgList;
     end;

     TCT_VectorVariant = class(TXPGBase)
protected
     FVt_Vector: TCT_Vector;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Vt_Vector: TCT_Vector read FVt_Vector;
     end;

     TCT_VectorLpstr = class(TXPGBase)
protected
     FVt_Vector: TCT_Vector;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Vt_Vector: TCT_Vector read FVt_Vector;
     end;

     TCT_DigSigBlob = class(TXPGBase)
protected
     FVt_Blob: longword;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Vt_Blob: longword read FVt_Blob write FVt_Blob;
     end;

     TCT_Properties = class(TXPGBase)
protected
     FTemplate: AxUCString;
     FManager: AxUCString;
     FCompany: AxUCString;
     FPages: longword;
     FWords: longword;
     FCharacters: longword;
     FPresentationFormat: AxUCString;
     FLines: longword;
     FParagraphs: longword;
     FSlides: longword;
     FNotes: longword;
     FTotalTime: longword;
     FHiddenSlides: longword;
     FMMClips: longword;
     FScaleCrop: boolean;
     FHeadingPairs: TCT_VectorVariant;
     FTitlesOfParts: TCT_VectorLpstr;
     FLinksUpToDate: boolean;
     FCharactersWithSpaces: longword;
     FSharedDoc: boolean;
     FHyperlinkBase: AxUCString;
     FHLinks: TCT_VectorVariant;
     FHyperlinksChanged: boolean;
     FDigSig: TCT_DigSigBlob;
     FApplication: AxUCString;
     FAppVersion: AxUCString;
     FDocSecurity: longword;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Template: AxUCString read FTemplate write FTemplate;
     property Manager: AxUCString read FManager write FManager;
     property Company: AxUCString read FCompany write FCompany;
     property Pages: longword read FPages write FPages;
     property Words: longword read FWords write FWords;
     property Characters: longword read FCharacters write FCharacters;
     property PresentationFormat: AxUCString read FPresentationFormat write FPresentationFormat;
     property Lines: longword read FLines write FLines;
     property Paragraphs: longword read FParagraphs write FParagraphs;
     property Slides: longword read FSlides write FSlides;
     property Notes: longword read FNotes write FNotes;
     property TotalTime: longword read FTotalTime write FTotalTime;
     property HiddenSlides: longword read FHiddenSlides write FHiddenSlides;
     property MMClips: longword read FMMClips write FMMClips;
     property ScaleCrop: boolean read FScaleCrop write FScaleCrop;
     property HeadingPairs: TCT_VectorVariant read FHeadingPairs;
     property TitlesOfParts: TCT_VectorLpstr read FTitlesOfParts;
     property LinksUpToDate: boolean read FLinksUpToDate write FLinksUpToDate;
     property CharactersWithSpaces: longword read FCharactersWithSpaces write FCharactersWithSpaces;
     property SharedDoc: boolean read FSharedDoc write FSharedDoc;
     property HyperlinkBase: AxUCString read FHyperlinkBase write FHyperlinkBase;
     property HLinks: TCT_VectorVariant read FHLinks;
     property HyperlinksChanged: boolean read FHyperlinksChanged write FHyperlinksChanged;
     property DigSig: TCT_DigSigBlob read FDigSig;
     property Application: AxUCString read FApplication write FApplication;
     property AppVersion: AxUCString read FAppVersion write FAppVersion;
     property DocSecurity: longword read FDocSecurity write FDocSecurity;
     end;

     T__ROOT__ = class(TXPGBase)
protected
     FRootAttributes: TStringXpgList;
     FProperties: TCT_Properties;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property RootAttributes: TStringXpgList read FRootAttributes;
     property Properties: TCT_Properties read FProperties;
     end;

     TXPGDocDocPropsApp = class(TXPGDocBase)
protected
     FRoot: T__ROOT__;
     FReader: TXPGReader;
     FWriter: TXpgWriteXML;
     function  GetProperties: TCT_Properties;
public
     constructor Create;
     destructor Destroy; override;

     procedure LoadFromFile(AFilename: AxUCString);
     procedure LoadFromStream(AStream: TStream);
     procedure SaveToFile(AFilename: AxUCString);

     procedure SaveToStream(AStream: TStream);

     property Root: T__ROOT__ read FRoot;
     property Properties: TCT_Properties read GetProperties;
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
  FEnums.AddObject('stabtVariant',TObject(0));
  FEnums.AddObject('stabtI1',TObject(1));
  FEnums.AddObject('stabtI2',TObject(2));
  FEnums.AddObject('stabtI4',TObject(3));
  FEnums.AddObject('stabtInt',TObject(4));
  FEnums.AddObject('stabtUi1',TObject(5));
  FEnums.AddObject('stabtUi2',TObject(6));
  FEnums.AddObject('stabtUi4',TObject(7));
  FEnums.AddObject('stabtUint',TObject(8));
  FEnums.AddObject('stabtR4',TObject(9));
  FEnums.AddObject('stabtR8',TObject(10));
  FEnums.AddObject('stabtDecimal',TObject(11));
  FEnums.AddObject('stabtBstr',TObject(12));
  FEnums.AddObject('stabtDate',TObject(13));
  FEnums.AddObject('stabtBool',TObject(14));
  FEnums.AddObject('stabtCy',TObject(15));
  FEnums.AddObject('stabtError',TObject(16));
  FEnums.AddObject('stvbtVariant',TObject(0));
  FEnums.AddObject('stvbtI1',TObject(1));
  FEnums.AddObject('stvbtI2',TObject(2));
  FEnums.AddObject('stvbtI4',TObject(3));
  FEnums.AddObject('stvbtI8',TObject(4));
  FEnums.AddObject('stvbtUi1',TObject(5));
  FEnums.AddObject('stvbtUi2',TObject(6));
  FEnums.AddObject('stvbtUi4',TObject(7));
  FEnums.AddObject('stvbtUi8',TObject(8));
  FEnums.AddObject('stvbtR4',TObject(9));
  FEnums.AddObject('stvbtR8',TObject(10));
  FEnums.AddObject('stvbtLpstr',TObject(11));
  FEnums.AddObject('stvbtLpwstr',TObject(12));
  FEnums.AddObject('stvbtBstr',TObject(13));
  FEnums.AddObject('stvbtDate',TObject(14));
  FEnums.AddObject('stvbtFiletime',TObject(15));
  FEnums.AddObject('stvbtBool',TObject(16));
  FEnums.AddObject('stvbtCy',TObject(17));
  FEnums.AddObject('stvbtError',TObject(18));
  FEnums.AddObject('stvbtClsid',TObject(19));
  FEnums.AddObject('stvbtCf',TObject(20));
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

{ TCT_Empty }

function  TCT_Empty.CheckAssigned: integer;
begin
  FAssigneds := [];
  Result := 0;
end;

procedure TCT_Empty.Write(AWriter: TXpgWriteXML);
begin
end;

constructor TCT_Empty.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 0;
end;

destructor TCT_Empty.Destroy;
begin
end;

procedure TCT_Empty.Clear;
begin
  FAssigneds := [];
end;

{ TCT_Null }

function  TCT_Null.CheckAssigned: integer;
begin
  FAssigneds := [];
  Result := 0;
end;

procedure TCT_Null.Write(AWriter: TXpgWriteXML);
begin
end;

constructor TCT_Null.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 0;
end;

destructor TCT_Null.Destroy;
begin
end;

procedure TCT_Null.Clear;
begin
  FAssigneds := [];
end;

{ TCT_Vstream }

function  TCT_Vstream.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FVersion <> '' then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
  if FContent <> 0 then 
  begin
    FAssigneds := FAssigneds + [xaContent];
    Inc(Result);
  end;
end;

function  TCT_Vstream.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_Vstream.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_Vstream.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FVersion <> '' then 
    AWriter.AddAttribute('version',FVersion);
end;

procedure TCT_Vstream.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'version' then 
    FVersion := AAttributes.Values[0]
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_Vstream.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
end;

destructor TCT_Vstream.Destroy;
begin
end;

procedure TCT_Vstream.Clear;
begin
  FAssigneds := [];
  FVersion := '';
end;

{ TCT_Cf }

function  TCT_Cf.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FFormat <> '' then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
  if FContent <> 0 then 
  begin
    FAssigneds := FAssigneds + [xaContent];
    Inc(Result);
  end;
end;

function  TCT_Cf.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_Cf.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_Cf.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FFormat <> '' then 
    AWriter.AddAttribute('format',FFormat);
end;

procedure TCT_Cf.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'format' then 
    FFormat := AAttributes.Values[0]
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_Cf.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
end;

destructor TCT_Cf.Destroy;
begin
end;

procedure TCT_Cf.Clear;
begin
  FAssigneds := [];
  FFormat := '';
end;

{ TCT_CfXpgList }

function  TCT_CfXpgList.GetItems(Index: integer): TCT_Cf;
begin
  Result := TCT_Cf(inherited Items[Index]);
end;

function  TCT_CfXpgList.Add: TCT_Cf;
begin
  Result := TCT_Cf.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_CfXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_CfXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaContent in Items[i].FAssigneds then 
      AWriter.Text := XmlIntToStr(GetItems(i).Content);
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

{ TCT_Variant }

function  TCT_Variant.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FVt_Variant <> Nil then 
    Inc(ElemsAssigned,FVt_Variant.CheckAssigned);
  if FVt_Vector <> Nil then 
    Inc(ElemsAssigned,FVt_Vector.CheckAssigned);
  if FVt_Array <> Nil then 
    Inc(ElemsAssigned,FVt_Array.CheckAssigned);
  if FVt_Blob <> 0 then 
    Inc(ElemsAssigned);
  if FVt_Oblob <> 0 then 
    Inc(ElemsAssigned);
  Inc(ElemsAssigned,FVt_Empty.CheckAssigned);
  Inc(ElemsAssigned,FVt_Null.CheckAssigned);
  if FVt_I1 <> 0 then 
    Inc(ElemsAssigned);
  if FVt_I2 <> 0 then 
    Inc(ElemsAssigned);
  if FVt_I4 <> 0 then 
    Inc(ElemsAssigned);
  if FVt_I8 <> 0 then 
    Inc(ElemsAssigned);
  if FVt_Int <> 0 then 
    Inc(ElemsAssigned);
  if FVt_Ui1 <> 0 then 
    Inc(ElemsAssigned);
  if FVt_Ui2 <> 0 then 
    Inc(ElemsAssigned);
  if FVt_Ui4 <> 0 then 
    Inc(ElemsAssigned);
  if FVt_Ui8 <> 0 then 
    Inc(ElemsAssigned);
  if FVt_Uint <> 0 then 
    Inc(ElemsAssigned);
  if FVt_R4 <> 0 then 
    Inc(ElemsAssigned);
  if FVt_R8 <> 0 then 
    Inc(ElemsAssigned);
  if FVt_Decimal <> 0 then 
    Inc(ElemsAssigned);
  if FVt_Lpstr <> '' then 
    Inc(ElemsAssigned);
  if FVt_Lpwstr <> '' then 
    Inc(ElemsAssigned);
  if FVt_Bstr <> '' then 
    Inc(ElemsAssigned);
  if FVt_Date <> 0 then 
    Inc(ElemsAssigned);
  if FVt_Filetime <> 0 then 
    Inc(ElemsAssigned);
  if FVt_Bool <> False then 
    Inc(ElemsAssigned);
  if FVt_Cy <> '' then 
    Inc(ElemsAssigned);
  if FVt_Error <> '' then 
    Inc(ElemsAssigned);
  if FVt_Stream <> 0 then 
    Inc(ElemsAssigned);
  if FVt_Ostream <> 0 then 
    Inc(ElemsAssigned);
  if FVt_Storage <> 0 then 
    Inc(ElemsAssigned);
  if FVt_Ostorage <> 0 then 
    Inc(ElemsAssigned);
  Inc(ElemsAssigned,FVt_Vstream.CheckAssigned);
  if FVt_Clsid <> '' then 
    Inc(ElemsAssigned);
  Inc(ElemsAssigned,FVt_Cf.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_Variant.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  case CalcHash_B(QName) of
    $B5C7FF21: begin
      if FVt_Variant = Nil then 
        FVt_Variant := TCT_Variant.Create(FOwner);
      Result := FVt_Variant;
    end;
    $81B62115: begin
      if FVt_Vector = Nil then 
        FVt_Vector := TCT_Vector.Create(FOwner);
      Result := FVt_Vector;
    end;
    $B440F481: begin
      if FVt_Array = Nil then 
        FVt_Array := TCT_Array.Create(FOwner);
      Result := FVt_Array;
    end;
    $967441F3: FVt_Blob := XmlStrToIntDef(AReader.Text,0);
    $98D69A10: FVt_Oblob := XmlStrToIntDef(AReader.Text,0);
    $7F5EEFD7: Result := FVt_Empty;
    $B853B495: Result := FVt_Null;
    $38DDD3DA: FVt_I1 := XmlStrToIntDef(AReader.Text,0);
    $38DDD3DB: FVt_I2 := XmlStrToIntDef(AReader.Text,0);
    $38DDD3DD: FVt_I4 := XmlStrToIntDef(AReader.Text,0);
    $38DDD3E1: FVt_I8 := XmlStrToIntDef(AReader.Text,0);
    $0A8080CD: FVt_Int := XmlStrToIntDef(AReader.Text,0);
    $B3221D13: FVt_Ui1 := XmlStrToIntDef(AReader.Text,0);
    $B3221D14: FVt_Ui2 := XmlStrToIntDef(AReader.Text,0);
    $B3221D16: FVt_Ui4 := XmlStrToIntDef(AReader.Text,0);
    $B3221D1A: FVt_Ui8 := XmlStrToIntDef(AReader.Text,0);
    $EF455F44: FVt_Uint := XmlStrToIntDef(AReader.Text,0);
    $4B44F38E: FVt_R4 := XmlStrToFloatDef(AReader.Text,0);
    $4B44F392: FVt_R8 := XmlStrToFloatDef(AReader.Text,0);
    $0BBD54E9: FVt_Decimal := XmlStrToFloatDef(AReader.Text,0);
    $1195B217: FVt_Lpstr := AReader.Text;
    $2B7BCD18: FVt_Lpwstr := AReader.Text;
    $77E8DE81: FVt_Bstr := AReader.Text;
    $5CBC47B4: FVt_Date := XmlStrToDateTime(AReader.Text);
    $3859BFE7: FVt_Filetime := XmlStrToDateTime(AReader.Text);
    $DB822482: FVt_Bool := XmlStrToBoolDef(AReader.Text,False);
    $81EE69AC: FVt_Cy := AReader.Text;
    $705AE430: FVt_Error := AReader.Text;
    $58E2571E: FVt_Stream := XmlStrToIntDef(AReader.Text,0);
    $DD4415E3: FVt_Ostream := XmlStrToIntDef(AReader.Text,0);
    $7788575F: FVt_Storage := XmlStrToIntDef(AReader.Text,0);
    $CFAA81EA: FVt_Ostorage := XmlStrToIntDef(AReader.Text,0);
    $8A564EC4: begin
      Result := FVt_Vstream;
      if AReader.HasText then 
        TCT_Vstream(Result).Content := XmlStrToIntDef(AReader.Text,0);
    end;
    $2E171597: FVt_Clsid := AReader.Text;
    $81EE6999: begin
      Result := FVt_Cf;
      if AReader.HasText then 
        TCT_Cf(Result).Content := XmlStrToIntDef(AReader.Text,0);
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,QName);
  end;
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_Variant.Write(AWriter: TXpgWriteXML);
begin
  if (FVt_Variant <> Nil) and FVt_Variant.Assigned then 
    if xaElements in FVt_Variant.FAssigneds then 
    begin
      AWriter.BeginTag('vt:variant');
      FVt_Variant.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('vt:variant');
  if (FVt_Vector <> Nil) and FVt_Vector.Assigned then 
  begin
    FVt_Vector.WriteAttributes(AWriter);
    if xaElements in FVt_Vector.FAssigneds then 
    begin
      AWriter.BeginTag('vt:vector');
      FVt_Vector.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('vt:vector');
  end;
  if (FVt_Array <> Nil) and FVt_Array.Assigned then 
  begin
    FVt_Array.WriteAttributes(AWriter);
    if xaElements in FVt_Array.FAssigneds then 
    begin
      AWriter.BeginTag('vt:array');
      FVt_Array.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('vt:array');
  end;
  if FVt_Blob <> 0 then 
    AWriter.SimpleTextTag('vt:blob',XmlIntToStr(FVt_Blob));
  if FVt_Oblob <> 0 then 
    AWriter.SimpleTextTag('vt:oblob',XmlIntToStr(FVt_Oblob));
  if FVt_Empty.Assigned then 
    AWriter.SimpleTag('vt:empty');
  if FVt_Null.Assigned then 
    AWriter.SimpleTag('vt:null');
  if FVt_I1 <> 0 then 
    AWriter.SimpleTextTag('vt:i1',XmlIntToStr(FVt_I1));
  if FVt_I2 <> 0 then 
    AWriter.SimpleTextTag('vt:i2',XmlIntToStr(FVt_I2));
  if FVt_I4 <> 0 then 
    AWriter.SimpleTextTag('vt:i4',XmlIntToStr(FVt_I4));
  if FVt_I8 <> 0 then 
    AWriter.SimpleTextTag('vt:i8',XmlIntToStr(FVt_I8));
  if FVt_Int <> 0 then 
    AWriter.SimpleTextTag('vt:int',XmlIntToStr(FVt_Int));
  if FVt_Ui1 <> 0 then 
    AWriter.SimpleTextTag('vt:ui1',XmlIntToStr(FVt_Ui1));
  if FVt_Ui2 <> 0 then 
    AWriter.SimpleTextTag('vt:ui2',XmlIntToStr(FVt_Ui2));
  if FVt_Ui4 <> 0 then 
    AWriter.SimpleTextTag('vt:ui4',XmlIntToStr(FVt_Ui4));
  if FVt_Ui8 <> 0 then 
    AWriter.SimpleTextTag('vt:ui8',XmlIntToStr(FVt_Ui8));
  if FVt_Uint <> 0 then 
    AWriter.SimpleTextTag('vt:uint',XmlIntToStr(FVt_Uint));
  if FVt_R4 <> 0 then 
    AWriter.SimpleTextTag('vt:r4',XmlFloatToStr(FVt_R4));
  if FVt_R8 <> 0 then 
    AWriter.SimpleTextTag('vt:r8',XmlFloatToStr(FVt_R8));
  if FVt_Decimal <> 0 then 
    AWriter.SimpleTextTag('vt:decimal',XmlFloatToStr(FVt_Decimal));
  if FVt_Lpstr <> '' then 
    AWriter.SimpleTextTag('vt:lpstr',FVt_Lpstr);
  if FVt_Lpwstr <> '' then 
    AWriter.SimpleTextTag('vt:lpwstr',FVt_Lpwstr);
  if FVt_Bstr <> '' then 
    AWriter.SimpleTextTag('vt:bstr',FVt_Bstr);
  if FVt_Date <> 0 then 
    AWriter.SimpleTextTag('vt:date',XmlDateTimeToStr(FVt_Date));
  if FVt_Filetime <> 0 then 
    AWriter.SimpleTextTag('vt:filetime',XmlDateTimeToStr(FVt_Filetime));
  if FVt_Bool <> False then 
    AWriter.SimpleTextTag('vt:bool',XmlBoolToStr(FVt_Bool));
  if FVt_Cy <> '' then 
    AWriter.SimpleTextTag('vt:cy',FVt_Cy);
  if FVt_Error <> '' then 
    AWriter.SimpleTextTag('vt:error',FVt_Error);
  if FVt_Stream <> 0 then 
    AWriter.SimpleTextTag('vt:stream',XmlIntToStr(FVt_Stream));
  if FVt_Ostream <> 0 then 
    AWriter.SimpleTextTag('vt:ostream',XmlIntToStr(FVt_Ostream));
  if FVt_Storage <> 0 then 
    AWriter.SimpleTextTag('vt:storage',XmlIntToStr(FVt_Storage));
  if FVt_Ostorage <> 0 then 
    AWriter.SimpleTextTag('vt:ostorage',XmlIntToStr(FVt_Ostorage));
  if FVt_Vstream.Assigned then 
  begin
    AWriter.Text := XmlIntToStr(FVt_Vstream.Content);
    FVt_Vstream.WriteAttributes(AWriter);
    AWriter.SimpleTag('vt:vstream');
  end;
  if FVt_Clsid <> '' then 
    AWriter.SimpleTextTag('vt:clsid',FVt_Clsid);
  if FVt_Cf.Assigned then 
  begin
    AWriter.Text := XmlIntToStr(FVt_Cf.Content);
    FVt_Cf.WriteAttributes(AWriter);
    AWriter.SimpleTag('vt:cf');
  end;
end;

constructor TCT_Variant.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 35;
  FAttributeCount := 0;
  FVt_Empty := TCT_Empty.Create(FOwner);
  FVt_Null := TCT_Null.Create(FOwner);
  FVt_Vstream := TCT_Vstream.Create(FOwner);
  FVt_Cf := TCT_Cf.Create(FOwner);
end;

destructor TCT_Variant.Destroy;
begin
  if FVt_Variant <> Nil then 
    FVt_Variant.Free;
  if FVt_Vector <> Nil then 
    FVt_Vector.Free;
  if FVt_Array <> Nil then 
    FVt_Array.Free;
  FVt_Empty.Free;
  FVt_Null.Free;
  FVt_Vstream.Free;
  FVt_Cf.Free;
end;

procedure TCT_Variant.Clear;
begin
  FAssigneds := [];
  if FVt_Variant <> Nil then 
    FVt_Variant.Clear;
  if FVt_Vector <> Nil then 
    FVt_Vector.Clear;
  if FVt_Array <> Nil then 
    FVt_Array.Clear;
  FVt_Blob := 0;
  FVt_Oblob := 0;
  FVt_Empty.Clear;
  FVt_Null.Clear;
  FVt_I1 := 0;
  FVt_I2 := 0;
  FVt_I4 := 0;
  FVt_I8 := 0;
  FVt_Int := 0;
  FVt_Ui1 := 0;
  FVt_Ui2 := 0;
  FVt_Ui4 := 0;
  FVt_Ui8 := 0;
  FVt_Uint := 0;
  FVt_R4 := 0;
  FVt_R8 := 0;
  FVt_Decimal := 0;
  FVt_Lpstr := '';
  FVt_Lpwstr := '';
  FVt_Bstr := '';
  FVt_Date := 0;
  FVt_Filetime := 0;
  FVt_Bool := False;
  FVt_Cy := '';
  FVt_Error := '';
  FVt_Stream := 0;
  FVt_Ostream := 0;
  FVt_Storage := 0;
  FVt_Ostorage := 0;
  FVt_Vstream.Clear;
  FVt_Clsid := '';
  FVt_Cf.Clear;
end;

function  TCT_Variant.CreateTCT_Variant: TCT_Variant;
begin
  Result := TCT_Variant.Create(FOwner);
end;

function  TCT_Variant.CreateTCT_Vector: TCT_Vector;
begin
  Result := TCT_Vector.Create(FOwner);
end;

function  TCT_Variant.CreateTCT_Array: TCT_Array;
begin
  Result := TCT_Array.Create(FOwner);
end;

{ TCT_VariantXpgList }

function  TCT_VariantXpgList.GetItems(Index: integer): TCT_Variant;
begin
  Result := TCT_Variant(inherited Items[Index]);
end;

function  TCT_VariantXpgList.Add: TCT_Variant;
begin
  Result := TCT_Variant.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_VariantXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_VariantXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
end;

{ TCT_Array }

function  TCT_Array.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FLBounds <> 0 then 
    Inc(AttrsAssigned);
  if FUBounds <> 0 then 
    Inc(AttrsAssigned);
  if Integer(FBaseType) <> XPG_UNKNOWN_ENUM then 
    Inc(AttrsAssigned);
  if FVt_VariantXpgList <> Nil then 
    Inc(ElemsAssigned,FVt_VariantXpgList.CheckAssigned);
  Inc(ElemsAssigned,FVt_I1XpgList.Count);
  Inc(ElemsAssigned,FVt_I2XpgList.Count);
  Inc(ElemsAssigned,FVt_I4XpgList.Count);
  Inc(ElemsAssigned,FVt_IntXpgList.Count);
  Inc(ElemsAssigned,FVt_Ui1XpgList.Count);
  Inc(ElemsAssigned,FVt_Ui2XpgList.Count);
  Inc(ElemsAssigned,FVt_Ui4XpgList.Count);
  Inc(ElemsAssigned,FVt_UintXpgList.Count);
  Inc(ElemsAssigned,FVt_R4XpgList.Count);
  Inc(ElemsAssigned,FVt_R8XpgList.Count);
  Inc(ElemsAssigned,FVt_DecimalXpgList.Count);
  Inc(ElemsAssigned,FVt_BstrXpgList.Count);
  Inc(ElemsAssigned,FVt_DateXpgList.Count);
  Inc(ElemsAssigned,FVt_BoolXpgList.Count);
  Inc(ElemsAssigned,FVt_ErrorXpgList.Count);
  Inc(ElemsAssigned,FVt_CyXpgList.Count);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_Array.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  case CalcHash_A(QName) of
    $00000419: begin
      if FVt_VariantXpgList = Nil then 
        FVt_VariantXpgList := TCT_VariantXpgList.Create(FOwner);
      Result := FVt_VariantXpgList.Add;
    end;
    $000001BE: FVt_I1XpgList.Add(AReader.Text);
    $000001BF: FVt_I2XpgList.Add(AReader.Text);
    $000001C1: FVt_I4XpgList.Add(AReader.Text);
    $0000026F: FVt_IntXpgList.Add(AReader.Text);
    $00000233: FVt_Ui1XpgList.Add(AReader.Text);
    $00000234: FVt_Ui2XpgList.Add(AReader.Text);
    $00000236: FVt_Ui4XpgList.Add(AReader.Text);
    $000002E4: FVt_UintXpgList.Add(AReader.Text);
    $000001CA: FVt_R4XpgList.Add(AReader.Text);
    $000001CE: FVt_R8XpgList.Add(AReader.Text);
    $000003F3: FVt_DecimalXpgList.Add(AReader.Text);
    $000002DF: FVt_BstrXpgList.Add(AReader.Text);
    $000002C2: FVt_DateXpgList.Add(AReader.Text);
    $000002D0: FVt_BoolXpgList.Add(AReader.Text);
    $0000034E: FVt_ErrorXpgList.Add(AReader.Text);
    $00000200: FVt_CyXpgList.Add(AReader.Text);
    else 
      FOwner.Errors.Error(xemUnknownElement,QName);
  end;
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_Array.Write(AWriter: TXpgWriteXML);
begin
  if FVt_VariantXpgList <> Nil then 
    FVt_VariantXpgList.Write(AWriter,'vt:variant');
  FVt_I1XpgList.WriteElements(AWriter,'vt:i1');
  FVt_I2XpgList.WriteElements(AWriter,'vt:i2');
  FVt_I4XpgList.WriteElements(AWriter,'vt:i4');
  FVt_IntXpgList.WriteElements(AWriter,'vt:int');
  FVt_Ui1XpgList.WriteElements(AWriter,'vt:ui1');
  FVt_Ui2XpgList.WriteElements(AWriter,'vt:ui2');
  FVt_Ui4XpgList.WriteElements(AWriter,'vt:ui4');
  FVt_UintXpgList.WriteElements(AWriter,'vt:uint');
  FVt_R4XpgList.WriteElements(AWriter,'vt:r4');
  FVt_R8XpgList.WriteElements(AWriter,'vt:r8');
  FVt_DecimalXpgList.WriteElements(AWriter,'vt:decimal');
  FVt_BstrXpgList.WriteElements(AWriter,'vt:bstr');
  FVt_DateXpgList.WriteElements(AWriter,'vt:date');
  FVt_BoolXpgList.WriteElements(AWriter,'vt:bool');
  FVt_ErrorXpgList.WriteElements(AWriter,'vt:error');
  FVt_CyXpgList.WriteElements(AWriter,'vt:cy');
end;

procedure TCT_Array.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('lBounds',XmlIntToStr(FLBounds));
  AWriter.AddAttribute('uBounds',XmlIntToStr(FUBounds));
  if Integer(FBaseType) <> XPG_UNKNOWN_ENUM then
    AWriter.AddAttribute('baseType',StrTST_ArrayBaseType[Integer(FBaseType)]);
end;

procedure TCT_Array.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
  S: AxUCString;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    S := AAttributes[i];
    case CalcHash_A(S) of
      $000002D7: FLBounds := XmlStrToIntDef(AAttributes.Values[i],0);
      $000002E0: FUBounds := XmlStrToIntDef(AAttributes.Values[i],0);
      $0000033D: FBaseType := TST_ArrayBaseType(StrToEnum('stabt' + AAttributes.Values[i]));
      else 
        FOwner.Errors.Error(xemUnknownAttribute,S);
    end;
  end
end;

constructor TCT_Array.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 17;
  FAttributeCount := 3;
  FVt_I1XpgList := TLongwordXpgList.Create;
  FVt_I2XpgList := TLongwordXpgList.Create;
  FVt_I4XpgList := TLongwordXpgList.Create;
  FVt_IntXpgList := TLongwordXpgList.Create;
  FVt_Ui1XpgList := TLongwordXpgList.Create;
  FVt_Ui2XpgList := TLongwordXpgList.Create;
  FVt_Ui4XpgList := TLongwordXpgList.Create;
  FVt_UintXpgList := TLongwordXpgList.Create;
  FVt_R4XpgList := TExtendedXpgList.Create;
  FVt_R8XpgList := TExtendedXpgList.Create;
  FVt_DecimalXpgList := TExtendedXpgList.Create;
  FVt_BstrXpgList := TStringXpgList.Create;
  FVt_DateXpgList := TTDateTimeXpgList.Create;
  FVt_BoolXpgList := TBooleanXpgList.Create;
  FVt_ErrorXpgList := TStringXpgList.Create;
  FVt_CyXpgList := TStringXpgList.Create;
  FBaseType := TST_ArrayBaseType(XPG_UNKNOWN_ENUM);
end;

destructor TCT_Array.Destroy;
begin
  if FVt_VariantXpgList <> Nil then 
    FVt_VariantXpgList.Free;
  FVt_I1XpgList.Free;
  FVt_I2XpgList.Free;
  FVt_I4XpgList.Free;
  FVt_IntXpgList.Free;
  FVt_Ui1XpgList.Free;
  FVt_Ui2XpgList.Free;
  FVt_Ui4XpgList.Free;
  FVt_UintXpgList.Free;
  FVt_R4XpgList.Free;
  FVt_R8XpgList.Free;
  FVt_DecimalXpgList.Free;
  FVt_BstrXpgList.Free;
  FVt_DateXpgList.Free;
  FVt_BoolXpgList.Free;
  FVt_ErrorXpgList.Free;
  FVt_CyXpgList.Free;
end;

procedure TCT_Array.Clear;
begin
  FAssigneds := [];
  if FVt_VariantXpgList <> Nil then 
    FVt_VariantXpgList.Clear;
  FVt_I1XpgList.DelimitedText := '';
  FVt_I2XpgList.DelimitedText := '';
  FVt_I4XpgList.DelimitedText := '';
  FVt_IntXpgList.DelimitedText := '';
  FVt_Ui1XpgList.DelimitedText := '';
  FVt_Ui2XpgList.DelimitedText := '';
  FVt_Ui4XpgList.DelimitedText := '';
  FVt_UintXpgList.DelimitedText := '';
  FVt_R4XpgList.DelimitedText := '';
  FVt_R8XpgList.DelimitedText := '';
  FVt_DecimalXpgList.DelimitedText := '';
  FVt_BstrXpgList.DelimitedText := '';
  FVt_DateXpgList.DelimitedText := '';
  FVt_BoolXpgList.DelimitedText := '';
  FVt_ErrorXpgList.DelimitedText := '';
  FVt_CyXpgList.DelimitedText := '';
  FLBounds := 0;
  FUBounds := 0;
  FBaseType := TST_ArrayBaseType(XPG_UNKNOWN_ENUM);
end;

function  TCT_Array.CreateTCT_VariantXpgList: TCT_VariantXpgList;
begin
  Result := TCT_VariantXpgList.Create(FOwner);
end;

{ TCT_Vector }

function  TCT_Vector.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if Integer(FBaseType) <> XPG_UNKNOWN_ENUM then
    Inc(AttrsAssigned);
  if FSize <> 0 then 
    Inc(AttrsAssigned);
  if FVt_VariantXpgList <> Nil then 
    Inc(ElemsAssigned,FVt_VariantXpgList.CheckAssigned);
  Inc(ElemsAssigned,FVt_I1XpgList.Count);
  Inc(ElemsAssigned,FVt_I2XpgList.Count);
  Inc(ElemsAssigned,FVt_I4XpgList.Count);
  Inc(ElemsAssigned,FVt_I8XpgList.Count);
  Inc(ElemsAssigned,FVt_Ui1XpgList.Count);
  Inc(ElemsAssigned,FVt_Ui2XpgList.Count);
  Inc(ElemsAssigned,FVt_Ui4XpgList.Count);
  Inc(ElemsAssigned,FVt_Ui8XpgList.Count);
  Inc(ElemsAssigned,FVt_R4XpgList.Count);
  Inc(ElemsAssigned,FVt_R8XpgList.Count);
  Inc(ElemsAssigned,FVt_LpstrXpgList.Count);
  Inc(ElemsAssigned,FVt_LpwstrXpgList.Count);
  Inc(ElemsAssigned,FVt_BstrXpgList.Count);
  Inc(ElemsAssigned,FVt_DateXpgList.Count);
  Inc(ElemsAssigned,FVt_FiletimeXpgList.Count);
  Inc(ElemsAssigned,FVt_BoolXpgList.Count);
  Inc(ElemsAssigned,FVt_CyXpgList.Count);
  Inc(ElemsAssigned,FVt_ErrorXpgList.Count);
  Inc(ElemsAssigned,FVt_ClsidXpgList.Count);
  Inc(ElemsAssigned,FVt_CfXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_Vector.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  case CalcHash_A(QName) of
    $00000419: begin
      if FVt_VariantXpgList = Nil then 
        FVt_VariantXpgList := TCT_VariantXpgList.Create(FOwner);
      Result := FVt_VariantXpgList.Add;
    end;
    $000001BE: FVt_I1XpgList.Add(AReader.Text);
    $000001BF: FVt_I2XpgList.Add(AReader.Text);
    $000001C1: FVt_I4XpgList.Add(AReader.Text);
    $000001C5: FVt_I8XpgList.Add(AReader.Text);
    $00000233: FVt_Ui1XpgList.Add(AReader.Text);
    $00000234: FVt_Ui2XpgList.Add(AReader.Text);
    $00000236: FVt_Ui4XpgList.Add(AReader.Text);
    $0000023A: FVt_Ui8XpgList.Add(AReader.Text);
    $000001CA: FVt_R4XpgList.Add(AReader.Text);
    $000001CE: FVt_R8XpgList.Add(AReader.Text);
    $00000359: FVt_LpstrXpgList.Add(AReader.Text);
    $000003D0: FVt_LpwstrXpgList.Add(AReader.Text);
    $000002DF: FVt_BstrXpgList.Add(AReader.Text);
    $000002C2: FVt_DateXpgList.Add(AReader.Text);
    $00000473: FVt_FiletimeXpgList.Add(AReader.Text);
    $000002D0: FVt_BoolXpgList.Add(AReader.Text);
    $00000200: FVt_CyXpgList.Add(AReader.Text);
    $0000034E: FVt_ErrorXpgList.Add(AReader.Text);
    $00000333: FVt_ClsidXpgList.Add(AReader.Text);
    $000001ED: begin
      Result := FVt_CfXpgList.Add;
      if AReader.HasText then 
        TCT_Cf(Result).Content := XmlStrToIntDef(AReader.Text,0);
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,QName);
  end;
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_Vector.Write(AWriter: TXpgWriteXML);
begin
  if FVt_VariantXpgList <> Nil then 
    FVt_VariantXpgList.Write(AWriter,'vt:variant');
  FVt_I1XpgList.WriteElements(AWriter,'vt:i1');
  FVt_I2XpgList.WriteElements(AWriter,'vt:i2');
  FVt_I4XpgList.WriteElements(AWriter,'vt:i4');
  FVt_I8XpgList.WriteElements(AWriter,'vt:i8');
  FVt_Ui1XpgList.WriteElements(AWriter,'vt:ui1');
  FVt_Ui2XpgList.WriteElements(AWriter,'vt:ui2');
  FVt_Ui4XpgList.WriteElements(AWriter,'vt:ui4');
  FVt_Ui8XpgList.WriteElements(AWriter,'vt:ui8');
  FVt_R4XpgList.WriteElements(AWriter,'vt:r4');
  FVt_R8XpgList.WriteElements(AWriter,'vt:r8');
  FVt_LpstrXpgList.WriteElements(AWriter,'vt:lpstr');
  FVt_LpwstrXpgList.WriteElements(AWriter,'vt:lpwstr');
  FVt_BstrXpgList.WriteElements(AWriter,'vt:bstr');
  FVt_DateXpgList.WriteElements(AWriter,'vt:date');
  FVt_FiletimeXpgList.WriteElements(AWriter,'vt:filetime');
  FVt_BoolXpgList.WriteElements(AWriter,'vt:bool');
  FVt_CyXpgList.WriteElements(AWriter,'vt:cy');
  FVt_ErrorXpgList.WriteElements(AWriter,'vt:error');
  FVt_ClsidXpgList.WriteElements(AWriter,'vt:clsid');
  FVt_CfXpgList.Write(AWriter,'vt:cf');
end;

procedure TCT_Vector.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if Integer(FBaseType) <> XPG_UNKNOWN_ENUM then
    AWriter.AddAttribute('baseType',StrTST_VectorBaseType[Integer(FBaseType)]);
  AWriter.AddAttribute('size',XmlIntToStr(FSize));
end;

procedure TCT_Vector.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
  S: AxUCString;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    S := AAttributes[i];
    case CalcHash_A(S) of
      $0000033D: FBaseType := TST_VectorBaseType(StrToEnum('stvbt' + AAttributes.Values[i]));
      $000001BB: FSize := XmlStrToIntDef(AAttributes.Values[i],0);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,S);
    end;
  end
end;

constructor TCT_Vector.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 21;
  FAttributeCount := 2;
  FVt_I1XpgList := TLongwordXpgList.Create;
  FVt_I2XpgList := TLongwordXpgList.Create;
  FVt_I4XpgList := TLongwordXpgList.Create;
  FVt_I8XpgList := TLongwordXpgList.Create;
  FVt_Ui1XpgList := TLongwordXpgList.Create;
  FVt_Ui2XpgList := TLongwordXpgList.Create;
  FVt_Ui4XpgList := TLongwordXpgList.Create;
  FVt_Ui8XpgList := TLongwordXpgList.Create;
  FVt_R4XpgList := TExtendedXpgList.Create;
  FVt_R8XpgList := TExtendedXpgList.Create;
  FVt_LpstrXpgList := TStringXpgList.Create;
  FVt_LpwstrXpgList := TStringXpgList.Create;
  FVt_BstrXpgList := TStringXpgList.Create;
  FVt_DateXpgList := TTDateTimeXpgList.Create;
  FVt_FiletimeXpgList := TTDateTimeXpgList.Create;
  FVt_BoolXpgList := TBooleanXpgList.Create;
  FVt_CyXpgList := TStringXpgList.Create;
  FVt_ErrorXpgList := TStringXpgList.Create;
  FVt_ClsidXpgList := TStringXpgList.Create;
  FVt_CfXpgList := TCT_CfXpgList.Create(FOwner);
  FBaseType := TST_VectorBaseType(XPG_UNKNOWN_ENUM);
end;

destructor TCT_Vector.Destroy;
begin
  if FVt_VariantXpgList <> Nil then
    FVt_VariantXpgList.Free;
  FVt_I1XpgList.Free;
  FVt_I2XpgList.Free;
  FVt_I4XpgList.Free;
  FVt_I8XpgList.Free;
  FVt_Ui1XpgList.Free;
  FVt_Ui2XpgList.Free;
  FVt_Ui4XpgList.Free;
  FVt_Ui8XpgList.Free;
  FVt_R4XpgList.Free;
  FVt_R8XpgList.Free;
  FVt_LpstrXpgList.Free;
  FVt_LpwstrXpgList.Free;
  FVt_BstrXpgList.Free;
  FVt_DateXpgList.Free;
  FVt_FiletimeXpgList.Free;
  FVt_BoolXpgList.Free;
  FVt_CyXpgList.Free;
  FVt_ErrorXpgList.Free;
  FVt_ClsidXpgList.Free;
  FVt_CfXpgList.Free;
end;

procedure TCT_Vector.Clear;
begin
  FAssigneds := [];
  if FVt_VariantXpgList <> Nil then
    FVt_VariantXpgList.Clear;
  FVt_I1XpgList.DelimitedText := '';
  FVt_I2XpgList.DelimitedText := '';
  FVt_I4XpgList.DelimitedText := '';
  FVt_I8XpgList.DelimitedText := '';
  FVt_Ui1XpgList.DelimitedText := '';
  FVt_Ui2XpgList.DelimitedText := '';
  FVt_Ui4XpgList.DelimitedText := '';
  FVt_Ui8XpgList.DelimitedText := '';
  FVt_R4XpgList.DelimitedText := '';
  FVt_R8XpgList.DelimitedText := '';
  FVt_LpstrXpgList.DelimitedText := '';
  FVt_LpwstrXpgList.DelimitedText := '';
  FVt_BstrXpgList.DelimitedText := '';
  FVt_DateXpgList.DelimitedText := '';
  FVt_FiletimeXpgList.DelimitedText := '';
  FVt_BoolXpgList.DelimitedText := '';
  FVt_CyXpgList.DelimitedText := '';
  FVt_ErrorXpgList.DelimitedText := '';
  FVt_ClsidXpgList.DelimitedText := '';
  FVt_CfXpgList.Clear;
  FBaseType := TST_VectorBaseType(XPG_UNKNOWN_ENUM);
  FSize := 0;
end;

function  TCT_Vector.CreateTCT_VariantXpgList: TCT_VariantXpgList;
begin
  if FVt_VariantXpgList <> Nil then
    FVt_VariantXpgList.Free;
  Result := TCT_VariantXpgList.Create(FOwner);
  FVt_VariantXpgList := Result;
end;

{ TCT_VectorVariant }

function  TCT_VectorVariant.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FVt_Vector.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_VectorVariant.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'vt:vector' then 
    Result := FVt_Vector
  else 
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_VectorVariant.Write(AWriter: TXpgWriteXML);
begin
  if FVt_Vector.Assigned then 
  begin
    FVt_Vector.WriteAttributes(AWriter);
    if xaElements in FVt_Vector.FAssigneds then 
    begin
      AWriter.BeginTag('vt:vector');
      FVt_Vector.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('vt:vector');
  end;
end;

constructor TCT_VectorVariant.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 0;
  FVt_Vector := TCT_Vector.Create(FOwner);
end;

destructor TCT_VectorVariant.Destroy;
begin
  FVt_Vector.Free;
end;

procedure TCT_VectorVariant.Clear;
begin
  FAssigneds := [];
  FVt_Vector.Clear;
end;

{ TCT_VectorLpstr }

function  TCT_VectorLpstr.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FVt_Vector.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_VectorLpstr.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'vt:vector' then 
    Result := FVt_Vector
  else 
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_VectorLpstr.Write(AWriter: TXpgWriteXML);
begin
  if FVt_Vector.Assigned then 
  begin
    FVt_Vector.WriteAttributes(AWriter);
    if xaElements in FVt_Vector.FAssigneds then 
    begin
      AWriter.BeginTag('vt:vector');
      FVt_Vector.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('vt:vector');
  end;
end;

constructor TCT_VectorLpstr.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 0;
  FVt_Vector := TCT_Vector.Create(FOwner);
end;

destructor TCT_VectorLpstr.Destroy;
begin
  FVt_Vector.Free;
end;

procedure TCT_VectorLpstr.Clear;
begin
  FAssigneds := [];
  FVt_Vector.Clear;
end;

{ TCT_DigSigBlob }

function  TCT_DigSigBlob.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FVt_Blob <> 0 then 
    Inc(ElemsAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_DigSigBlob.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'vt:blob' then 
    FVt_Blob := XmlStrToIntDef(AReader.Text,0)
  else 
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_DigSigBlob.Write(AWriter: TXpgWriteXML);
begin
  if FVt_Blob <> 0 then 
    AWriter.SimpleTextTag('vt:blob',XmlIntToStr(FVt_Blob));
end;

constructor TCT_DigSigBlob.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 0;
end;

destructor TCT_DigSigBlob.Destroy;
begin
end;

procedure TCT_DigSigBlob.Clear;
begin
  FAssigneds := [];
  FVt_Blob := 0;
end;

{ TCT_Properties }

function  TCT_Properties.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FTemplate <> '' then 
    Inc(ElemsAssigned);
  if FManager <> '' then 
    Inc(ElemsAssigned);
  if FCompany <> '' then 
    Inc(ElemsAssigned);
  if FPages <> 0 then 
    Inc(ElemsAssigned);
  if FWords <> 0 then 
    Inc(ElemsAssigned);
  if FCharacters <> 0 then 
    Inc(ElemsAssigned);
  if FPresentationFormat <> '' then 
    Inc(ElemsAssigned);
  if FLines <> 0 then 
    Inc(ElemsAssigned);
  if FParagraphs <> 0 then 
    Inc(ElemsAssigned);
  if FSlides <> 0 then 
    Inc(ElemsAssigned);
  if FNotes <> 0 then 
    Inc(ElemsAssigned);
  if FTotalTime <> 0 then 
    Inc(ElemsAssigned);
  if FHiddenSlides <> 0 then 
    Inc(ElemsAssigned);
  if FMMClips <> 0 then 
    Inc(ElemsAssigned);
  if FScaleCrop <> False then 
    Inc(ElemsAssigned);
  Inc(ElemsAssigned,FHeadingPairs.CheckAssigned);
  Inc(ElemsAssigned,FTitlesOfParts.CheckAssigned);
  if FLinksUpToDate <> False then 
    Inc(ElemsAssigned);
  if FCharactersWithSpaces <> 0 then 
    Inc(ElemsAssigned);
  if FSharedDoc <> False then 
    Inc(ElemsAssigned);
  if FHyperlinkBase <> '' then 
    Inc(ElemsAssigned);
  Inc(ElemsAssigned,FHLinks.CheckAssigned);
  if FHyperlinksChanged <> False then 
    Inc(ElemsAssigned);
  Inc(ElemsAssigned,FDigSig.CheckAssigned);
  if FApplication <> '' then 
    Inc(ElemsAssigned);
  if FAppVersion <> '' then 
    Inc(ElemsAssigned);
  if FDocSecurity <> 0 then 
    Inc(ElemsAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_Properties.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  case CalcHash_A(QName) of
    $0000033C: FTemplate := AReader.Text;
    $000002BB: FManager := AReader.Text;
    $000002D7: FCompany := AReader.Text;
    $000001F0: FPages := XmlStrToIntDef(AReader.Text,0);
    $0000020F: FWords := XmlStrToIntDef(AReader.Text,0);
    $00000400: FCharacters := XmlStrToIntDef(AReader.Text,0);
    $00000765: FPresentationFormat := AReader.Text;
    $000001FB: FLines := XmlStrToIntDef(AReader.Text,0);
    $00000409: FParagraphs := XmlStrToIntDef(AReader.Text,0);
    $00000264: FSlides := XmlStrToIntDef(AReader.Text,0);
    $00000209: FNotes := XmlStrToIntDef(AReader.Text,0);
    $00000393: FTotalTime := XmlStrToIntDef(AReader.Text,0);
    $000004B0: FHiddenSlides := XmlStrToIntDef(AReader.Text,0);
    $00000295: FMMClips := XmlStrToIntDef(AReader.Text,0);
    $0000037C: FScaleCrop := XmlStrToBoolDef(AReader.Text,False);
    $000004AF: Result := FHeadingPairs;
    $00000534: Result := FTitlesOfParts;
    $00000507: FLinksUpToDate := XmlStrToBoolDef(AReader.Text,False);
    $000007FB: FCharactersWithSpaces := XmlStrToIntDef(AReader.Text,0);
    $0000036D: FSharedDoc := XmlStrToBoolDef(AReader.Text,False);
    $00000531: FHyperlinkBase := AReader.Text;
    $00000249: Result := FHLinks;
    $000006D3: FHyperlinksChanged := XmlStrToBoolDef(AReader.Text,False);
    $00000237: Result := FDigSig;
    $00000474: FApplication := AReader.Text;
    $00000407: FAppVersion := AReader.Text;
    $0000046E: FDocSecurity := XmlStrToIntDef(AReader.Text,0);
    else 
      FOwner.Errors.Error(xemUnknownElement,QName);
  end;
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_Properties.Write(AWriter: TXpgWriteXML);
begin
  if FTemplate <> '' then 
    AWriter.SimpleTextTag('Template',FTemplate);
  if FManager <> '' then 
    AWriter.SimpleTextTag('Manager',FManager);
  if FCompany <> '' then 
    AWriter.SimpleTextTag('Company',FCompany);
  if FPages <> 0 then 
    AWriter.SimpleTextTag('Pages',XmlIntToStr(FPages));
  if FWords <> 0 then 
    AWriter.SimpleTextTag('Words',XmlIntToStr(FWords));
  if FCharacters <> 0 then 
    AWriter.SimpleTextTag('Characters',XmlIntToStr(FCharacters));
  if FPresentationFormat <> '' then 
    AWriter.SimpleTextTag('PresentationFormat',FPresentationFormat);
  if FLines <> 0 then 
    AWriter.SimpleTextTag('Lines',XmlIntToStr(FLines));
  if FParagraphs <> 0 then 
    AWriter.SimpleTextTag('Paragraphs',XmlIntToStr(FParagraphs));
  if FSlides <> 0 then 
    AWriter.SimpleTextTag('Slides',XmlIntToStr(FSlides));
  if FNotes <> 0 then 
    AWriter.SimpleTextTag('Notes',XmlIntToStr(FNotes));
  if FTotalTime <> 0 then 
    AWriter.SimpleTextTag('TotalTime',XmlIntToStr(FTotalTime));
  if FHiddenSlides <> 0 then 
    AWriter.SimpleTextTag('HiddenSlides',XmlIntToStr(FHiddenSlides));
  if FMMClips <> 0 then 
    AWriter.SimpleTextTag('MMClips',XmlIntToStr(FMMClips));
  if FScaleCrop <> False then 
    AWriter.SimpleTextTag('ScaleCrop',XmlBoolToStr(FScaleCrop));
  if FHeadingPairs.Assigned then 
    if xaElements in FHeadingPairs.FAssigneds then 
    begin
      AWriter.BeginTag('HeadingPairs');
      FHeadingPairs.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('HeadingPairs');
  if FTitlesOfParts.Assigned then 
    if xaElements in FTitlesOfParts.FAssigneds then 
    begin
      AWriter.BeginTag('TitlesOfParts');
      FTitlesOfParts.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('TitlesOfParts');
  if FLinksUpToDate <> False then 
    AWriter.SimpleTextTag('LinksUpToDate',XmlBoolToStr(FLinksUpToDate));
  if FCharactersWithSpaces <> 0 then 
    AWriter.SimpleTextTag('CharactersWithSpaces',XmlIntToStr(FCharactersWithSpaces));
  if FSharedDoc <> False then 
    AWriter.SimpleTextTag('SharedDoc',XmlBoolToStr(FSharedDoc));
  if FHyperlinkBase <> '' then 
    AWriter.SimpleTextTag('HyperlinkBase',FHyperlinkBase);
  if FHLinks.Assigned then 
    if xaElements in FHLinks.FAssigneds then 
    begin
      AWriter.BeginTag('HLinks');
      FHLinks.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('HLinks');
  if FHyperlinksChanged <> False then 
    AWriter.SimpleTextTag('HyperlinksChanged',XmlBoolToStr(FHyperlinksChanged));
  if FDigSig.Assigned then 
    if xaElements in FDigSig.FAssigneds then 
    begin
      AWriter.BeginTag('DigSig');
      FDigSig.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('DigSig');
  if FApplication <> '' then 
    AWriter.SimpleTextTag('Application',FApplication);
  if FAppVersion <> '' then 
    AWriter.SimpleTextTag('AppVersion',FAppVersion);
  if FDocSecurity <> 0 then 
    AWriter.SimpleTextTag('DocSecurity',XmlIntToStr(FDocSecurity));
end;

constructor TCT_Properties.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 27;
  FAttributeCount := 0;
  FHeadingPairs := TCT_VectorVariant.Create(FOwner);
  FTitlesOfParts := TCT_VectorLpstr.Create(FOwner);
  FHLinks := TCT_VectorVariant.Create(FOwner);
  FDigSig := TCT_DigSigBlob.Create(FOwner);
end;

destructor TCT_Properties.Destroy;
begin
  FHeadingPairs.Free;
  FTitlesOfParts.Free;
  FHLinks.Free;
  FDigSig.Free;
end;

procedure TCT_Properties.Clear;
begin
  FAssigneds := [];
  FTemplate := '';
  FManager := '';
  FCompany := '';
  FPages := 0;
  FWords := 0;
  FCharacters := 0;
  FPresentationFormat := '';
  FLines := 0;
  FParagraphs := 0;
  FSlides := 0;
  FNotes := 0;
  FTotalTime := 0;
  FHiddenSlides := 0;
  FMMClips := 0;
  FScaleCrop := False;
  FHeadingPairs.Clear;
  FTitlesOfParts.Clear;
  FLinksUpToDate := False;
  FCharactersWithSpaces := 0;
  FSharedDoc := False;
  FHyperlinkBase := '';
  FHLinks.Clear;
  FHyperlinksChanged := False;
  FDigSig.Clear;
  FApplication := '';
  FAppVersion := '';
  FDocSecurity := 0;
end;

{ T__ROOT__ }

function  T__ROOT__.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FProperties.CheckAssigned);
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
  case CalcHash_B(QName) of
    $DF9CD057: Result := FProperties;
    else 
      FOwner.Errors.Error(xemUnknownElement,QName);
  end;
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure T__ROOT__.Write(AWriter: TXpgWriteXML);
begin
  AWriter.Attributes := FRootAttributes.Text;
  if FProperties.Assigned then
    if xaElements in FProperties.FAssigneds then 
    begin
      AWriter.BeginTag('Properties');
      FProperties.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('Properties');
end;

constructor T__ROOT__.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FRootAttributes := TStringXpgList.Create;
  FElementCount := 36;
  FAttributeCount := 0;
  FProperties := TCT_Properties.Create(FOwner);
end;

destructor T__ROOT__.Destroy;
begin
  FRootAttributes.Free;
  FProperties.Free;
end;

procedure T__ROOT__.Clear;
begin
  FRootAttributes.Clear;
  FAssigneds := [];
  FProperties.Clear;
end;

{ TTXPGDocDocPropsApp }

function  TXPGDocDocPropsApp.GetProperties: TCT_Properties;
begin
  Result := FRoot.Properties;
end;

constructor TXPGDocDocPropsApp.Create;
begin
  FRoot := T__ROOT__.Create(Self);
  FErrors := TXpgPErrors.Create;
  FReader := TXPGReader.Create(FErrors,FRoot);
  FWriter := TXpgWriteXML.Create;
end;

destructor TXPGDocDocPropsApp.Destroy;
begin
  FRoot.Free;
  FReader.Free;
  FWriter.Free;
  FErrors.Free;
  inherited Destroy;
end;

procedure TXPGDocDocPropsApp.LoadFromFile(AFilename: AxUCString);
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

procedure TXPGDocDocPropsApp.LoadFromStream(AStream: TStream);
begin
  FReader.LoadFromStream(AStream);
end;

procedure TXPGDocDocPropsApp.SaveToFile(AFilename: AxUCString);
begin
  FWriter.SaveToFile(AFilename);
  FRoot.CheckAssigned;
  FRoot.Write(FWriter);
end;

procedure TXPGDocDocPropsApp.SaveToStream(AStream: TStream);
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
