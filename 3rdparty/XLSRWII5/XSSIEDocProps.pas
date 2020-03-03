unit XSSIEDocProps;

interface

{$B-}
{$H+}
{$R-}
{$I AxCompilers.inc}

uses { Delphi } Classes, SysUtils, Contnrs, vcl.Graphics,
{$ifdef DELPHI_XE3_OR_LATER}
     System.UITypes,
{$endif}
     XLSUtils5,
     { AXWord } XSSIEDefs, XSSIEUtils;

const AXW_COLOR_AUTOMATIC = $FF000000;
const AXW_COLOR_WHITE     = $00FFFFFF;

type TAXWChpId = (
axciBooleanStart,
axciBold,axciItalic,axciStrikeTrough,
axciBooleanEnd,

axciIntegerStart,
axciColor,axciUnderline,axciSubSuperscript,
axciIntegerEnd,

axciFloatStart,
axciSize,
axciFloatEnd,

axciComplexStart,
axciFontName,
axciComplexEnd);

type TAXWPapId = (
axpiBooleanStart,
axpiBooleanEnd,

axpiIntegerStart,
axpiFlags,axpiAlignment,axpiColor,
axpiIntegerEnd,

axpiFloatStart,
axpiIndent,axpiIndentRight,axpiIndentFirstLine,axpiIndentHanging,axpiLineSpaceing,axpiSpaceing,
axpiFloatEnd,

axpiComplexStart,
axpiComplexEnd);

type TAXWPropIdType = (axcitUnknown,axcitBoolean,axcitInteger,axcitFloat,axcitComplex);

// axcuSingle must be the first "visible" value.
type TAXWChpUnderline = (axcuNone,axcuSingle,axcuDouble);
type TAXWChpSubSuperscript = (axcssNone,axcssSubscript,axcssSuperscript);

type TAXWPapTextAlign = (axptaLeft,axptaCenter,axptaRight,axptaJustify);

type TAXWPapFlag = (axpfReadOnly);
     TAXWPapFlags = set of TAXWPapFlag;


type PAXWPropItem = ^TAXWPropItem;
     TAXWPropItem = record
     Id: integer;
     PData: PByteArray;
     case integer of
       0: (ValInt: integer);
       1: (ValBool: boolean);
       2: (ValFloat: double);
     end;

type TAXWTabStopAlignment = (atsaClear,atsaLeft,atsaCenter,atsaRight,atsaDecimal,atsaNum);
type TAXWTabStopLeader = (atslNone,atslDot,atslHyphen,atslUnderscore,atslHeavy,atslMiddleDot);

type TAXWTabStop = class(TObject)
protected
     FPosition : double;
     FAlignment: TAXWTabStopAlignment;
     FLeader   : TAXWTabStopLeader;
public
     property Position : double read FPosition write FPosition;
     property Alignment: TAXWTabStopAlignment read FAlignment write FAlignment;
     property Leader   : TAXWTabStopLeader read FLeader write FLeader;
     end;

type TAXWTabStops = class(TObjectList)
private
     function GetItems(Index: integer): TAXWTabStop;
protected
public
     constructor Create;

     function Add(const APosition: double): TAXWTabStop;

     property Items[Index: integer]: TAXWTabStop read GetItems; default;
     end;

type TAXWPropList = class(TList)
protected
     function  GetItems(Index: integer): PAXWPropItem;

     function  GetComplex(Index: integer): AxUCString;
     function  GetValueType(Index: integer): TAXWPropIdType;
     function  GetAsBoolean(Id: integer): boolean;
     function  GetAsInteger(Id: integer): integer;
     function  GetAsFloat(Id: integer): double;
     function  GetFloatVal(Index: integer): double;
     procedure SetAsBoolean(Id: integer; const Value: boolean);
     procedure SetAsInteger(Id: integer; const Value: integer);
     procedure SetAsFloat(Id: integer; const Value: double);
     procedure SetFloatVal(Index: integer; const Value: double);
     procedure SetComplex(Index: integer; const Value: AxUCString);

     function  IsBoolean(Id: integer): boolean; virtual; abstract;
     function  IsComplex(Id: integer): boolean; virtual; abstract;
     function  IsInteger(Id: integer): boolean; virtual; abstract;
     function  IsFloat(Id: integer): boolean; virtual; abstract;

     function  GetValue(AId: integer; out AValue: integer): boolean; overload;
     function  GetValue(AId: integer; out AValue: longword): boolean; overload;
     function  GetValue(AId: integer; out AValue: double): boolean; overload;
     function  GetValue(AId: integer; out AValue: boolean): boolean; overload;
     function  GetValue(AId: integer; out AValue: AxUCString): boolean; overload;

     procedure Notify(Ptr: Pointer; Action: TListNotification); override;
     function  Find(Id: integer): integer;

     property Items[Index: integer]: PAXWPropItem read GetItems; default;
     property ValueType[Index: integer]: TAXWPropIdType  read GetValueType;
     property Complex[Index: integer]: AxUCString read GetComplex write SetComplex;
     property FloatVal[Index: integer]: double read GetFloatVal write SetFloatVal;
public
     procedure Clear; override;

     procedure Toggle(Id: integer);
     procedure Assign(List: TAXWPropList);
     function  Equal(List: TAXWPropList): boolean;
     procedure Merge(List: TAXWPropList);
     procedure GetDebugList(List: TStrings); virtual; abstract;

     procedure AddInteger(Id: integer; Value: integer);
     procedure AddFloat(Id: integer; Value: double);
     procedure AddBoolean(Id: integer; Value: boolean);
     procedure AddComplex(Id: integer; Value: AxUCString);

     property AsBoolean[Id: integer]: boolean read GetAsBoolean write SetAsBoolean;
     property AsInteger[Id: integer]: integer read GetAsInteger write SetAsInteger;
     end;

type TAXWCHP = class;


     TAXWCHPX = class(TAXWPropList)
private
     function  GetUnderline: TAXWChpUnderline;
     function  GetFontName: AxUCString;
     function  GetBold: boolean;
     function  GetColor: longword;
     function  GetItalic: boolean;
     function  GetSize: double;
     function  GetStrikeTrough: boolean;
     function  GetSubSuperscript: TAXWChpSubSuperscript;
     procedure SetUnderline(const Value: TAXWChpUnderline);
     procedure SetFontName(const Value: AxUCString);
     procedure SetBold(const Value: boolean);
     procedure SetColor(const Value: longword);
     procedure SetItalic(const Value: boolean);
     procedure SetSize(const Value: double);
     procedure SetStrikeTrough(const Value: boolean);
     procedure SetSubSuperscript(const Value: TAXWChpSubSuperscript);
protected
     FCHP: TAXWCHP;

     function  IsBoolean(Id: integer): boolean; override;
     function  IsComplex(Id: integer): boolean; override;
     function  IsInteger(Id: integer): boolean; override;
     function  IsFloat(Id: integer): boolean; override;
public
     constructor Create(ACHP: TAXWCHP);

//     function  Equal(ACHPX:

     procedure GetDebugList(List: TStrings); override;

     property Bold: boolean read GetBold write SetBold;
     property Italic: boolean read GetItalic write SetItalic;
     property FontName: AxUCString read GetFontName write SetFontName;
     property Underline: TAXWChpUnderline read GetUnderline write SetUnderline;
     property Size: double read GetSize write SetSize;
     property Color: longword read GetColor write SetColor;
     property StrikeTrough: boolean read GetStrikeTrough write SetStrikeTrough;
     property SubSuperscript: TAXWChpSubSuperscript read GetSubSuperscript write SetSubSuperscript;
     end;

     TAXWCHP = class(TObject)
private
     FParentCHP: TAXWCHP;

     FUnderline     : TAXWChpUnderline;
     FColor         : longword;
     FFontName      : AxUCString;
     FSize          : double;
     FItalic        : boolean;
     FBold          : boolean;
     FStrikeTrough  : boolean;
     FSubSuperscript: TAXWChpSubSuperscript;
public
     constructor Create(ParentCHP: TAXWCHP);

     procedure Clear;

     procedure Assign(CHP: TAXWCHP); overload;
     procedure Assign(List: TAXWCHPX); overload;
     procedure AssignTFont(Font: TFont);
     procedure CopyToCHPX(List: TAXWCHPX);
     function  IsEmptyCHPX(List: TAXWCHPX): boolean;
     procedure CompactCHPX(List: TAXWCHPX);
     function  Equal(ACHP: TAXWCHP): boolean;

     property ParentCHP: TAXWCHP read FParentCHP write FParentCHP;

     property Bold: boolean read FBold write FBold;
     property Italic: boolean read FItalic write FItalic;
     property Underline: TAXWChpUnderline read FUnderline write FUnderline;
     property FontName: AxUCString read FFontName write FFontName;
     property Size: double read FSize write FSize;
     property Color: longword read FColor write FColor;
     property StrikeTrough: boolean read FStrikeTrough write FStrikeTrough;
     property SubSuperscript: TAXWChpSubSuperscript read FSubSuperscript write FSubSuperscript;
     end;

type TAXWPAPX = class;

     TAXWPAP = class(TObject)
private
     FIndent         : double;
     FIndentRight    : double;
     FIndentFirstLine: double;
     FIndentHanging  : double;
     FLineSpacing    : double;
     FSpacing        : double;
     FAlignment      : TAXWPapTextAlign;
     FFlags          : TAXWPapFlag; // TODO shall be TAXWPapFlags
     FColor          : longword;

     FParentPAP: TAXWPAP;
public
     constructor Create(ParentPAP: TAXWPAP);

     procedure Clear;

     procedure Assign(PAP: TAXWPAP); overload;
     procedure Assign(List: TAXWPAPX); overload;
     procedure CopyToPAPX(List: TAXWPAPX);
     function  IsEmptyPAPX(List: TAXWPAPX): boolean;
     procedure CompactPAPX(List: TAXWPAPX);

     property Alignment: TAXWPapTextAlign read FAlignment write FAlignment;
     property Indent: double read FIndent write FIndent;
     property IndentRight: double read FIndentRight write FIndentRight;
     property IndentFirstLine: double read FIndentFirstLine write FIndentFirstLine;
     property IndentHanging: double read FIndentHanging write FIndentHanging;
     property LineSpacing: double read FLineSpacing write FLineSpacing;
     property Spacing: double read FSpacing write FSpacing;
     property Color: longword read FColor write FColor;

     // TODO Check compatibility with Word.
     property Flags: TAXWPapFlag read FFlags write FFlags;
     end;

     TAXWPAPX = class(TAXWPropList)
private
     function  GetAlignment: TAXWPapTextAlign;
     function  GetColor: longword;
     procedure SetColor(const Value: longword);
     function  GetIndent: double;
     procedure SetIndent(const Value: double);
     function  GetIndentFirstLine: double;
     function  GetIndentHanging: double;
     function  GetIndentRight: double;
     procedure SetIndentFirstLine(const Value: double);
     procedure SetIndentHanging(const Value: double);
     procedure SetIndentRight(const Value: double);
protected
     FPAP: TAXWPAP;

     function  IsBoolean(Id: integer): boolean; override;
     function  IsComplex(Id: integer): boolean; override;
     function  IsInteger(Id: integer): boolean; override;
     function  IsFloat(Id: integer): boolean; override;
public
     constructor Create(APAP: TAXWPAP);

     procedure GetDebugList(List: TStrings); override;

     property Alignment: TAXWPapTextAlign read GetAlignment;
     property Indent: double read GetIndent write SetIndent;
     property IndentRight: double read GetIndentRight write SetIndentRight;
     property IndentFirstLine: double read GetIndentFirstLine write SetIndentFirstLine;
     property IndentHanging: double read GetIndentHanging write SetIndentHanging;
     property Color: longword read GetColor write SetColor;
     end;


type TAXWDOP = class(TObject)
protected
     FColor   : longword;
     FTabWidth: double;
     FDefTab  : TAXWTabStop;
public
     constructor Create;
     destructor Destroy; override;

     property Color: longword read FColor write FColor;
     property TabWidth: double read FTabWidth write FTabWidth;
     property DefTab: TAXWTabStop read FDefTab;
     end;

implementation

{ TAXWCHP }

procedure TAXWCHP.Assign(CHP: TAXWCHP);
begin
  FUnderline      := CHP.FUnderline;
  FSubSuperscript := CHP.FSubSuperscript;
  FStrikeTrough   := CHP.FStrikeTrough;
  FColor          := CHP.FColor;
  FFontName       := CHP.FFontName;
  FSize           := CHP.FSize;
  FItalic         := CHP.FItalic;
  FBold           := CHP.FBold;
end;

procedure TAXWCHP.Assign(List: TAXWCHPX);
var
  i: integer;
begin
  Clear;
  for i := 0 to List.Count - 1 do begin
    case TAXWChpId(List.Items[i].Id) of
      axciBold            : FBold := List.Items[i].ValBool;
      axciItalic          : FItalic := List.Items[i].ValBool;
      axciSize            : FSize := List.Items[i].ValFloat;
      axciColor           : FColor := List.Items[i].ValInt;
      axciUnderline       : FUnderline := TAXWChpUnderline(List.Items[i].ValInt);
      axciFontName        : FFontName := List.Complex[i];
      axciSubSuperscript  : FSubSuperscript := TAXWChpSubSuperscript(List.Items[i].ValInt);
      axciStrikeTrough    : FStrikeTrough := List.Items[i].ValBool;
    end;
  end;
end;

procedure TAXWCHP.AssignTFont(Font: TFont);
begin
  FUnderline      := axcuNone;
  FSubSuperscript := axcssNone;
  FStrikeTrough   := False;
  FItalic         := False;
  FBold           := False;

  if fsBold in Font.Style then FBold := True;
  if fsItalic in Font.Style then FItalic := True;
  if fsStrikeOut in Font.Style then FStrikeTrough := True;
  if fsUnderline in Font.Style then FUnderline := axcuSingle;

  FFontName := Font.Name;
  FColor    := Font.Color; // TODO correct value for automatic colors.
  FSize     := Font.Size;
end;

procedure TAXWCHP.CompactCHPX(List: TAXWCHPX);
var
  i1,i2: integer;
  DeleteItem: boolean;
begin
  i1 := 0;
  i2 := List.Count - 1;
  while i1 <= i2 do begin
    DeleteItem := False;
    case TAXWChpId(List[i1].Id) of
      axciBold           : DeleteItem := FBold = List.Items[i1].ValBool;
      axciItalic         : DeleteItem := FItalic = List.Items[i1].ValBool;
      axciStrikeTrough   : DeleteItem := FStrikeTrough = List.Items[i1].ValBool;
      axciSize           : DeleteItem := FSize = List.Items[i1].ValFloat;
      axciColor          : DeleteItem := FColor = Longword(List.Items[i1].ValInt);
      axciUnderline      : DeleteItem := FUnderline = TAXWChpUnderline(List.Items[i1].ValInt);
      axciSubSuperscript : DeleteItem := FSubSuperscript = TAXWChpSubSuperscript(List.Items[i1].ValInt);
      axciFontName       : DeleteItem := FFontName = List.Complex[i1];
    end;
    if DeleteItem then begin
      List.Delete(i1);
      Dec(i2);
    end
    else
      Inc(i1);
  end;
end;

procedure TAXWCHP.CopyToCHPX(List: TAXWCHPX);
begin
  List.Clear;

  if FParentCHP <> Nil then begin
    if FUnderline       <> FParentCHP.FUnderline        then List.AddInteger(Integer(axciUnderline),Integer(FUnderline));
    if FSubSuperscript  <> FParentCHP.FSubSuperscript   then List.AddInteger(Integer(axciSubSuperscript),Integer(FSubSuperscript));
    if FColor           <> FParentCHP.FColor            then List.AddInteger(Integer(axciColor),FColor);
    if FFontName        <> FParentCHP.FFontName         then List.AddComplex(Integer(axciFontName),FFontName);
    if FSize            <> FParentCHP.FSize             then List.AddFloat(Integer(axciSize),FSize);
    if FItalic          <> FParentCHP.FItalic           then List.AddBoolean(Integer(axciItalic),FItalic);
    if FBold            <> FParentCHP.FBold             then List.AddBoolean(Integer(axciBold),FBold);
    if FStrikeTrough    <> FParentCHP.FStrikeTrough     then List.AddBoolean(Integer(axciStrikeTrough),FStrikeTrough);
  end;
end;

constructor TAXWCHP.Create(ParentCHP: TAXWCHP);
begin
  FParentCHP := ParentCHP;
  Clear;
end;

function TAXWCHP.Equal(ACHP: TAXWCHP): boolean;
begin
  Result :=
  (FUnderline = ACHP.FUnderline) and
  (FColor = ACHP.FColor) and
   SameText(FFontName,ACHP.FFontName) and
  (FSize = ACHP.FSize) and
  (FItalic = ACHP.FItalic) and
  (FBold = ACHP.FBold) and
  (FStrikeTrough = ACHP.FStrikeTrough) and
  (FSubSuperscript = ACHP.FSubSuperscript);
end;

function TAXWCHP.IsEmptyCHPX(List: TAXWCHPX): boolean;
var
  i: integer;
begin
  Result := True;
  for i := 0 to List.Count - 1 do begin
    case TAXWChpId(List.Items[i].Id) of
      axciBold            : Result := FBold = List.Items[i].ValBool;
      axciItalic          : Result := FItalic = List.Items[i].ValBool;
      axciSize            : Result := FSize = List.Items[i].ValFloat;
      axciColor           : Result := FColor = Longword(List.Items[i].ValInt);
      axciUnderline       : Result := FUnderline = TAXWChpUnderline(List.Items[i].ValInt);
      axciFontName        : Result := FFontName = List.Complex[i];
      axciSubSuperscript  : Result := FSubSuperscript = TAXWChpSubSuperscript(List.Items[i].ValInt);
      axciStrikeTrough    : Result := FStrikeTrough = List.Items[i].ValBool;
    end;
    if not Result then
      Exit;
  end;
end;

procedure TAXWCHP.Clear;
begin
  if FParentCHP <> Nil then begin
    FColor := FParentCHP.FColor;
    FFontName := FParentCHP.FFontName;
    FSize := FParentCHP.FSize;
    FItalic := FParentCHP.FItalic;
    FBold := FParentCHP.FBold;
    FUnderline := FParentCHP.FUnderline;
    FColor := FParentCHP.FColor;
    FStrikeTrough := FParentCHP.FStrikeTrough;
    FSubSuperscript := FParentCHP.FSubSuperscript;
  end
  else begin
    FColor := clBlack;
    FFontName := 'Arial';
    FSize := 10;
    FItalic := False;
    FBold := False;
    FUnderline := axcuNone;
    FColor := clBlack;
    FStrikeTrough := False;
    FSubSuperscript := axcssNone;
  end;
end;

{ TAXWPropList }

procedure TAXWPropList.AddInteger(Id: integer; Value: integer);
var
  PropItem: PAXWPropItem;
begin
  New(PropItem);
  PropItem.Id := Id;
  PropItem.ValInt := Value;
  inherited Add(PropItem);
end;

procedure TAXWPropList.AddBoolean(Id: integer; Value: boolean);
var
  PropItem: PAXWPropItem;
begin
  New(PropItem);
  PropItem.Id := Id;
  PropItem.ValBool := Value;
  inherited Add(PropItem);
end;

procedure TAXWPropList.AddComplex(Id: integer; Value: AxUCString);
var
  PropItem: PAXWPropItem;
begin
  New(PropItem);
  PropItem.ValInt := Length(Value);
  GetMem(PropItem.PData,Length(Value) * 2);
  System.Move(Value[1],PropItem.PData[0],Length(Value) * 2);
  PropItem.Id := Id;
  inherited Add(PropItem);
end;

procedure TAXWPropList.AddFloat(Id: integer; Value: double);
var
  PropItem: PAXWPropItem;
begin
  New(PropItem);
  PropItem.ValFloat := Value;
  PropItem.Id := Id;
  inherited Add(PropItem);
end;

procedure TAXWPropList.Assign(List: TAXWPropList);
var
  i: integer;
begin
  if List = Nil then
    Exit;
  Clear;
  for i := 0 to List.Count - 1 do begin
    case List.ValueType[i] of
      axcitUnknown: raise XLSRWException.Create('Unknown CHPX');
      axcitBoolean: AddBoolean(List.Items[i].Id,List.Items[i].ValBool);
      axcitInteger: AddInteger(List.Items[i].Id,List.Items[i].ValInt);
      axcitFloat  : AddFloat(List.Items[i].Id,List.Items[i].ValFloat);
      axcitComplex: AddComplex(List.Items[i].Id,List.Complex[i]);
    end;
  end;
end;

procedure TAXWPropList.Clear;
var
  i: integer;
begin
  for i := 0 to Count - 1 do begin
    if IsComplex(Items[i].Id) then begin
      FreeMem(Items[i].PData);
      Items[i].PData := Nil;
    end;
//    FreeMem(Items[i]);
  end;
  inherited;
end;

function TAXWPropList.Equal(List: TAXWPropList): boolean;
var
  i: integer;
begin
  Result := False;
  if Count = List.Count then begin
    for i := 0 to List.Count - 1 do begin
      Result := Items[i].Id = List.Items[i].Id;
      if not Result then
        Exit;
      case ValueType[i] of
        axcitBoolean: Result := Items[i].ValBool = List.Items[i].ValBool;
        axcitInteger: Result := Items[i].ValInt = List.Items[i].ValInt;
        axcitComplex: Result := Complex[i] = List.Complex[i];
      end;
      if not Result then
        Exit;
    end;
  end;
end;

function TAXWPropList.Find(Id: integer): integer;
var
  First: Integer;
  Last: Integer;
begin
  First := 0;
  Last := Count - 1;

  while First <= Last do begin
    Result := (First + Last) div 2;
    if Id = Items[Result].Id then
      Exit;
    if Items[Result].Id > Id then
      Last := Result - 1
    else
      First := Result + 1;
  end;
  Result := -1;
end;

function TAXWPropList.GetAsBoolean(Id: integer): boolean;
var
  Index: integer;
begin
  Index := Find(Id);
  if Index >= 0 then
    Result := GetItems(Index).ValBool
  else
    Result := False;
end;

function TAXWPropList.GetAsFloat(Id: integer): double;
var
  Index: integer;
begin
  Index := Find(Id);
  if Index >= 0 then
    Result := GetItems(Index).ValFloat
  else
    Result := 0;
end;

function TAXWPropList.GetAsInteger(Id: integer): integer;
var
  Index: integer;
begin
  Index := Find(Id);
  if Index >= 0 then
    Result := GetItems(Index).ValInt
  else
    Result := 0;
end;

function TAXWPropList.GetComplex(Index: integer): AxUCString;
var
  CHPXItem: PAXWPropItem;
begin
  CHPXItem := Items[Index];
  SetLength(Result,CHPXItem.ValInt);
  System.Move(CHPXItem.PData[0],Result[1],CHPXItem.ValInt * 2);
end;

function TAXWPropList.GetFloatVal(Index: integer): double;
begin
  Result := Items[Index].ValFloat;
end;

function TAXWPropList.GetItems(Index: integer): PAXWPropItem;
begin
  Result := inherited Items[Index];
end;

function TAXWPropList.GetValue(AId: integer; out AValue: integer): boolean;
var
  i: integer;
begin
  i := Find(AId);
  Result := i >= 0;
  if Result then
    AValue := GetItems(i).ValInt;
end;

function TAXWPropList.GetValue(AId: integer; out AValue: double): boolean;
var
  i: integer;
begin
  i := Find(AId);
  Result := i >= 0;
  if Result then
    AValue := GetItems(i).ValFloat;
end;

function TAXWPropList.GetValue(AId: integer; out AValue: boolean): boolean;
var
  i: integer;
begin
  i := Find(AId);
  Result := i >= 0;
  if Result then
    AValue := GetItems(i).ValBool;
end;

function TAXWPropList.GetValueType(Index: integer): TAXWPropIdType;
begin
  if IsBoolean(Items[Index].Id)then
    Result := axcitBoolean
  else if IsInteger(Items[Index].Id) then
    Result := axcitInteger
  else if IsFloat(Items[Index].Id) then
    Result := axcitFloat
  else if IsComplex(Items[Index].Id) then
    Result := axcitComplex
  else
    Result := axcitUnknown;
end;

procedure TAXWPropList.Merge(List: TAXWPropList);
var
  i,j: integer;
begin
  if List = Nil then
    Exit;

  for i := 0 to List.Count - 1 do begin
    j := Find(List.Items[i].Id);
    case List.ValueType[i] of
      axcitBoolean: begin // Toggle(List[i].Id); { TODO }
        if j >= 0 then
          Items[j].ValBool := List.Items[i].ValBool
        else
          AddBoolean(List.Items[i].Id,List.Items[i].ValBool);
      end;
      axcitInteger: begin
        if j >= 0 then
          Items[j].ValInt := List.Items[i].ValInt
        else
          AddInteger(List.Items[i].Id,List.Items[i].ValInt);
      end;
      axcitFloat: begin
        if j >= 0 then
          Items[j].ValFloat := List.Items[i].ValFloat
        else
          AddFloat(List.Items[i].Id,List.Items[i].ValFloat);
      end;
      axcitComplex: begin
        if j >= 0 then
          Complex[j] := List.Complex[i]
        else
          AddComplex(List.Items[i].Id,List.Complex[i]);
      end;
    end;
  end;
end;

function PropListCompare(Item1, Item2: Pointer): Integer;
begin
  Result := Integer(PAXWPropItem(Item1).Id) - Integer(PAXWPropItem(Item2).Id);
end;

procedure TAXWPropList.Notify(Ptr: Pointer; Action: TListNotification);
var
  Item: PAXWPropItem;
begin
  inherited;
  if Action = lnDeleted then begin
    Item := Ptr;
    if IsComplex(Item.Id) and (Item.PData <> Nil) then
      FreeMem(Item.PData);
    FreeMem(Item);
  end
  else if Action = lnAdded then
    Sort(PropListCompare);
end;

procedure TAXWPropList.SetAsBoolean(Id: integer; const Value: boolean);
var
  Index: integer;
begin
  Index := Find(Id);
  if Index >= 0 then
    GetItems(Index).ValBool := Value
  else
    AddBoolean(Id,Value);
end;

procedure TAXWPropList.SetAsFloat(Id: integer; const Value: double);
var
  Index: integer;
begin
  Index := Find(Id);
  if Index >= 0 then
    GetItems(Index).ValFloat := Value
  else
    AddFloat(Id,Value);
end;

procedure TAXWPropList.SetAsInteger(Id: integer; const Value: integer);
var
  Index: integer;
begin
  Index := Find(Id);
  if Index >= 0 then
    GetItems(Index).ValInt := Value
  else
    AddInteger(Id,Value);
end;

procedure TAXWPropList.SetComplex(Index: integer; const Value: AxUCString);
begin
  raise XLSRWException.Create('Not implemented: SetComplex');
end;

procedure TAXWPropList.SetFloatVal(Index: integer; const Value: double);
begin
  Items[Index].ValFloat := Value;
end;

procedure TAXWPropList.Toggle(Id: integer);
var
  i: integer;
begin
  if IsBoolean(Id) then begin
    i := Find(Id);
    if i >= 0 then begin
      Delete(i);
    end
    else
      AddInteger(Id,1);
  end
  else
    raise XLSRWException.Create('Can not toggle this type');
end;

function TAXWPropList.GetValue(AId: integer; out AValue: AxUCString): boolean;
var
  i: integer;
begin
  i := Find(AId);
  Result := i >= 0;
  if Result then
    AValue := GetComplex(i);
end;

function TAXWPropList.GetValue(AId: integer; out AValue: longword): boolean;
var
  i: integer;
begin
  i := Find(AId);
  Result := i >= 0;
  if Result then
    AValue := GetItems(i).ValInt;
end;

{ TAXWCHPX }

constructor TAXWCHPX.Create(ACHP: TAXWCHP);
begin
  FCHP := ACHP;
end;

function TAXWCHPX.GetBold: boolean;
begin
  if not GetValue(Integer(axciBold),Result) then
    Result := FCHP.Bold;
end;

function TAXWCHPX.GetColor: longword;
begin
  if not GetValue(Integer(axciColor),Result) then
    Result := FCHP.Color;
end;

procedure TAXWCHPX.GetDebugList(List: TStrings);
var
  i: integer;

function StrBool(B: boolean): AxUCString;
begin
  if B then
    Result := 'True'
  else
    Result := 'False';
end;

begin
  for i := 0 to Count - 1 do begin
    case TAXWChpId(Items[i].Id) of
      axciBooleanStart    : List.Add('BooleanStart: ');
      axciBold            : List.Add('Bold: ' + StrBool(Items[i].ValBool));
      axciItalic          : List.Add('Italic: ' + StrBool(Items[i].ValBool));
      axciBooleanEnd      : List.Add('BooleanEnd: ');
      axciIntegerStart    : List.Add('IntegerStart: ');
      axciSize            : List.Add('Size: ' + FloatToStr(Items[i].ValFloat));
      axciColor           : List.Add('Color: ' + IntToStr(Items[i].ValInt));
      axciUnderline       : List.Add('Underline: ' + IntToStr(Items[i].ValInt));
      axciIntegerEnd      : List.Add('IntegerEnd: ');
      axciComplexStart    : List.Add('ComplexStart: ');
      axciFontName        : List.Add('FontName: ' + Complex[i]);
      axciComplexEnd      : List.Add('ComplexEnd: ');
    end;
  end;
end;

function TAXWCHPX.GetFontName: AxUCString;
begin
  if not GetValue(Integer(axciFontName),Result) then
    Result := FCHP.FontName;
end;

function TAXWCHPX.GetItalic: boolean;
begin
  if not GetValue(Integer(axciItalic),Result) then
    Result := FCHP.Italic;
end;

function TAXWCHPX.GetSize: double;
begin
  if not GetValue(Integer(axciSize),Result) then
    Result := FCHP.Size;
end;

function TAXWCHPX.GetStrikeTrough: boolean;
begin
  if not GetValue(Integer(axciStrikeTrough),Result) then
    Result := FCHP.StrikeTrough;
end;

function TAXWCHPX.GetSubSuperscript: TAXWChpSubSuperscript;
var
  Val: integer;
begin
  if not GetValue(Integer(axciSubSuperscript),Val) then
    Result := axcssNone
  else
    Result := FCHP.SubSuperscript;
end;

function TAXWCHPX.IsBoolean(Id: integer): boolean;
begin
  Result := TAXWChpId(Id) in [axciBooleanStart..axciBooleanEnd];
end;

function TAXWCHPX.IsComplex(Id: integer): boolean;
begin
  Result := TAXWChpId(Id) in [axciComplexStart..axciComplexEnd];
end;

function TAXWCHPX.IsFloat(Id: integer): boolean;
begin
  Result := TAXWChpId(Id) in [axciFloatStart..axciFloatEnd];
end;

function TAXWCHPX.IsInteger(Id: integer): boolean;
begin
  Result := TAXWChpId(Id) in [axciIntegerStart..axciIntegerEnd];
end;

function TAXWCHPX.GetUnderline: TAXWChpUnderline;
var
  Val: integer;
begin
  if GetValue(Integer(axciUnderline),Val) then
    Result := TAXWChpUnderline(Val)
  else
    Result := FCHP.Underline;
end;

procedure TAXWCHPX.SetBold(const Value: boolean);
begin
  if Value <> FCHP.Bold then
    SetAsBoolean(Integer(axciBold),Value);
end;

procedure TAXWCHPX.SetColor(const Value: longword);
begin
  if Value <> Longword(FCHP.Color) then
    SetAsInteger(Integer(axciColor),Value);
end;

procedure TAXWCHPX.SetFontName(const Value: AxUCString);
begin
  if Value <> FCHP.FontName then
    AddComplex(Integer(axciFontName),Value);
end;

procedure TAXWCHPX.SetItalic(const Value: boolean);
begin
  if Value <> FCHP.Italic then
    SetAsBoolean(Integer(axciItalic),Value);
end;

procedure TAXWCHPX.SetSize(const Value: double);
begin
  if Value <> FCHP.Size then
    SetAsFloat(Integer(axciSize),Value);
end;

procedure TAXWCHPX.SetStrikeTrough(const Value: boolean);
begin
  if Value <> FCHP.StrikeTrough then
    SetAsBoolean(Integer(axciStrikeTrough),Value);
end;

procedure TAXWCHPX.SetSubSuperscript(const Value: TAXWChpSubSuperscript);
begin
  if Value <> FCHP.SubSuperscript then
    SetAsInteger(Integer(axciSubSuperscript),Integer(Value));
end;

procedure TAXWCHPX.SetUnderline(const Value: TAXWChpUnderline);
begin
  if Value <> FCHP.Underline then
    SetAsInteger(Integer(axciUnderline),Integer(Value));
end;

{ TAXWPAPX }

constructor TAXWPAPX.Create(APAP: TAXWPAP);
begin
  FPAP := APAP;
end;

function TAXWPAPX.GetAlignment: TAXWPapTextAlign;
var
  i: integer;
begin
  i := Find(Integer(axpiAlignment));
  if i >= 0 then
    Result := TAXWPapTextAlign(Items[i].ValInt)
  else
    Result := axptaLeft;
end;

function TAXWPAPX.GetColor: longword;
begin
  if not GetValue(Integer(axpiColor),Result) then
    Result := FPAP.Color;
end;

procedure TAXWPAPX.GetDebugList(List: TStrings);
begin

end;

function TAXWPAPX.GetIndent: double;
begin
  if not GetValue(Integer(axpiIndent),Result) then
    Result := FPAP.Indent;
end;

function TAXWPAPX.GetIndentFirstLine: double;
begin
  if not GetValue(Integer(axpiIndentFirstLine),Result) then
    Result := FPAP.IndentFirstLine;
end;

function TAXWPAPX.GetIndentHanging: double;
begin
  if not GetValue(Integer(axpiIndentHanging),Result) then
    Result := FPAP.IndentHanging;
end;

function TAXWPAPX.GetIndentRight: double;
begin
  if not GetValue(Integer(axpiIndentRight),Result) then
    Result := FPAP.IndentRight;
end;

function TAXWPAPX.IsBoolean(Id: integer): boolean;
begin
  Result := TAXWPapId(Id) in [axpiBooleanStart..axpiBooleanEnd];
end;

function TAXWPAPX.IsComplex(Id: integer): boolean;
begin
  Result := TAXWPapId(Id) in [axpiComplexStart..axpiComplexEnd];
end;

function TAXWPAPX.IsFloat(Id: integer): boolean;
begin
  Result := TAXWPapId(Id) in [axpiFloatStart..axpiFloatEnd];
end;

function TAXWPAPX.IsInteger(Id: integer): boolean;
begin
  Result := TAXWPapId(Id) in [axpiIntegerStart..axpiIntegerEnd];
end;

procedure TAXWPAPX.SetColor(const Value: longword);
begin
  if Value <> FPAP.Color then
    SetAsInteger(Integer(axpiColor),Value);
end;

procedure TAXWPAPX.SetIndent(const Value: double);
begin
  if Value <> FPAP.Indent then
    SetAsFloat(Integer(axpiIndent),Value);
end;

procedure TAXWPAPX.SetIndentFirstLine(const Value: double);
begin
  if Value <> FPAP.IndentFirstLine then
    SetAsFloat(Integer(axpiIndentFirstLine),Value);
end;

procedure TAXWPAPX.SetIndentHanging(const Value: double);
begin
  if Value <> FPAP.IndentHanging then
    SetAsFloat(Integer(axpiIndentHanging),Value);
end;

procedure TAXWPAPX.SetIndentRight(const Value: double);
begin
  if Value <> FPAP.IndentRight then
    SetAsFloat(Integer(axpiIndentRight),Value);
end;

{ TAXWPAP }

procedure TAXWPAP.Assign(PAP: TAXWPAP);
begin
  FIndent := PAP.FIndent;
  FLineSpacing := PAP.FLineSpacing;
  FSpacing := PAP.FSpacing;
  FAlignment := PAP.FAlignment;
  FColor := PAP.Color;
  FFlags := PAP.FFlags;
end;

procedure TAXWPAP.Assign(List: TAXWPAPX);
var
  i: integer;
begin
  Clear;
  for i := 0 to List.Count - 1 do begin
    case TAXWPapId(List[i].Id) of
      axpiIndent       : FIndent := List.FloatVal[i];
      axpiLineSpaceing : FLineSpacing := List.FloatVal[i];
      axpiSpaceing     : FSpacing := List.FloatVal[i];
      axpiAlignment    : FAlignment := TAXWPapTextAlign(List.Items[i].ValInt);
      axpiColor        : FColor := List.Items[i].ValInt;
      axpiFlags        : FFlags := TAXWPapFlag(List.Items[i].ValInt);
    end;
  end;
end;

procedure TAXWPAP.Clear;
begin
  if FParentPAP <> Nil then
    Assign(FParentPAP)
  else begin
    FIndent := 0;
    FLineSpacing := 1;
    FSpacing := 0;
    FAlignment := axptaLeft;
    FColor := AXW_COLOR_AUTOMATIC;
  //  FFlags := [];
  end;
end;

procedure TAXWPAP.CompactPAPX(List: TAXWPAPX);
var
  i1,i2: integer;
  DeleteItem: boolean;
begin
  i1 := 0;
  i2 := List.Count - 1;
  while i1 <= i2 do begin
    DeleteItem := False;
    case TAXWPapId(List[i1].Id) of
      axpiIndent       : DeleteItem := FIndent = List.FloatVal[i1];
      axpiLineSpaceing : DeleteItem := FLineSpacing = List.FloatVal[i1];
      axpiSpaceing     : DeleteItem := FSpacing = List.FloatVal[i1];
      axpiAlignment    : DeleteItem := FAlignment = TAXWPapTextAlign(List.Items[i1].ValInt);
      axpiFlags        : DeleteItem := FFlags = TAXWPapFlag(List.Items[i1].ValInt);
    end;
    if DeleteItem then begin
      List.Delete(i1);
      Dec(i2);
    end
    else
      Inc(i1);
  end;
end;

procedure TAXWPAP.CopyToPAPX(List: TAXWPAPX);
begin
  if FParentPAP = Nil then
    raise XLSRWException.Create('Parent PAP is Nil');

  List.Clear;
  if FIndent      <> FParentPAP.FIndent      then List.AddFloat(Integer(axpiIndent),FIndent);
  if FLineSpacing <> FParentPAP.FLineSpacing then List.AddFloat(Integer(axpiLineSpaceing),FLineSpacing);
  if FSpacing     <> FParentPAP.FSpacing     then List.AddFloat(Integer(axpiSpaceing),FSpacing);
  if FAlignment   <> FParentPAP.FAlignment   then List.AddInteger(Integer(axpiAlignment),Integer(FAlignment));
  if FFlags       <> FParentPAP.FFlags       then List.AddInteger(Integer(axpiFlags),Integer(FFlags));
end;

constructor TAXWPAP.Create(ParentPAP: TAXWPAP);
begin
  FParentPAP := ParentPAP;

  Clear;
end;

{
     FIndent: double;
     FLineSpacing: double;
     FSpacing: double;
     FAlignment: TAXWPapTextAlign;
     FFlags: TAXWPapFlag;
}
function TAXWPAP.IsEmptyPAPX(List: TAXWPAPX): boolean;
var
  i: integer;
begin
  Result := True;
  for i := 0 to List.Count - 1 do begin
    case TAXWPapId(List.Items[i].Id) of
      axpiIndent          : Result := FIndent = List.FloatVal[i];
      axpiLineSpaceing    : Result := FLineSpacing = List.FloatVal[i];
      axpiSpaceing        : Result := FSpacing = List.FloatVal[i];
      axpiAlignment       : Result := FAlignment = TAXWPapTextAlign(List.Items[i].ValInt);
      axpiFlags           : Result := FFlags = TAXWPapFlag(List.Items[i].ValInt);
    end;
    if not Result then
      Exit;
  end;
end;

{ TAXWDOP }

constructor TAXWDOP.Create;
begin
  FColor := AXW_COLOR_WHITE;
  FTabWidth := (72 / 2.54) * 2.3; // 2.3 cm

  FDefTab := TAXWTabStop.Create;
  FDefTab.Position := 0;
  FDefTab.Alignment := atsaLeft;
  FDefTab.Leader := atslNone;
end;

destructor TAXWDOP.Destroy;
begin
  FDefTab.Free;

  inherited;
end;

{ TAXWTabStops }

function TAXWTabStops.Add(const APosition: double): TAXWTabStop;
begin
  Result := TAXWTabStop.Create;
  Result.Position := APosition;
  inherited Add(Result);
end;

constructor TAXWTabStops.Create;
begin
  inherited Create;
end;

function TAXWTabStops.GetItems(Index: integer): TAXWTabStop;
begin
  Result := TAXWTabStop(inherited Items[Index]);
end;

end.

