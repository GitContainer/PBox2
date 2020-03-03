unit XSSIEUtils;

{-
********************************************************************************
******* XLSSpreadSheet V3.00                                             *******
*******                                                                  *******
******* Copyright(C) 2006,2017 Lars Arvidsson, Axolot Data               *******
*******                                                                  *******
******* email: components@axolot.com                                     *******
******* URL:   http://www.axolot.com                                     *******
********************************************************************************
** Users of the XLSSpreadSheet component must accept the following            **
** disclaimer of warranty:                                                    **
**                                                                            **
** XLSSpreadSheet is supplied as is. The author disclaims all warranties,     **
** expressed or implied, including, without limitation, the warranties of     **
** merchantability and of fitness for any purpose. The author assumes no      **
** liability for damages, direct or consequential, which may result from the  **
** use of XLSSpreadSheet.                                                     **
********************************************************************************
}

{$B-}
{$H+}
{$R-}
{$I AxCompilers.inc}

interface

uses { Delphi  } Classes, SysUtils, Contnrs, vcl.Graphics, Types,
                 XLSUtils5,
     { AXWord  } XSSIEDefs {, XBookPaintGDI2};

//type TAXWGDI = TXBookGDI;

{$ifdef DELPHI_2009_OR_LATER}
type XLS8String = RawByteString;
type XLS8Char = AnsiChar;
type XLS8PChar = PAnsiChar;
type AxUCChar = char;
type AxUCString = string;
type AxPUCChar = PChar;
{$else}
type XLS8String = string;
type XLS8Char = Char;
type XLS8PChar = PChar;
type AxUCChar = WideChar;
type AxUCString = WideString;
type AxPUCChar = PWideChar;

type NativeInt = integer;
{$endif}

type PLongwordArray = ^TLongwordArray;
     TLongwordArray = array[0..0] of Longword;

type TAXWVertAlign = (avaBottom,avaCenter,avaTop,avaDistributed);

type TAXWFindTextOption = (axtoMatchCase,axtoWholeWords,axtoWildcards,axtoSoundex);

type TAXWCharType = (actWhitespace,actNotWhitespace,actPrintable);

type TBreakCharData = record
     Width: double;
     Char: AxUCChar;
     Percent: double;
     end;

//
//    ---+=(2)==========================+================+================+----------------+----------------+
//      ||                              |(3) ############|###############||(4) XXXXXXXXXXXX|XXXXXXXXXXXXXXXX|
//    ---+ (1),(3)                      +================+================+----------------+----------------+
//      ||                              |################|###############||XXXXXXXXXXXXXXXX|XXXXXXXXXXXXXXXX|
//    ---+==============================+================+================+----------------+----------------+
//       |XXXXXXXXXXXXXXXXXXXXXXXXXXXXXX|XXXXXXXXXXXXXXXX|XXXXXXXXXXXXXXXX|XXXXXXXXXXXXXXXX|XXXXXXXXXXXXXXXX|
//    ---+------------------------------+----------------+----------------+----------------+----------------+
//       |XXXXXXXXXXXXXXXXXXXXXXXXXXXXXX|XXXXXXXXXXXXXXXX|XXXXXXXXXXXXXXXX|XXXXXXXXXXXXXXXX|XXXXXXXXXXXXXXXX|
//    ---+------------------------------+----------------+----------------+----------------+----------------+
//       |XXXXXXXXXXXXXXXXXXXXXXXXXXXXXX|XXXXXXXXXXXXXXXX|XXXXXXXXXXXXXXXX|XXXXXXXXXXXXXXXX|XXXXXXXXXXXXXXXX|
//    ---+------------------------------+----------------+----------------+----------------+----------------+
//       |XXXXXXXXXXXXXXXXXXXXXXXXXXXXXX|XXXXXXXXXXXXXXXX|XXXXXXXXXXXXXXXX|XXXXXXXXXXXXXXXX|XXXXXXXXXXXXXXXX|
//    ---+------------------------------+----------------+----------------+----------------+----------------+
//
//
// 1. Cell area. The area the text shall be diplayed in. Can be more than one cell
//    if the cells are merged.
//
// 2. Editor area. The area that the editor is painted into. Will expand as more
//    text is typed.
//
// 3. Client area, equal to editor area minus margins. Will expand as the editor
//    area is expanded.
//
// 4. Max area. The maximum area the editor area can expand to. The text is
//    paginated into this area. Limited by the window size of the spreadsheet.
//
// 5. Padding. Extra margin outside the editor area, depending on the cell border
//    thickness.

type TAXWClientArea = class(TObject)
private
     procedure SetCX1(const Value: integer); virtual;
     procedure SetCX2(const Value: integer); virtual;
     procedure SetCY1(const Value: integer); virtual;
     procedure SetCY2(const Value: integer); virtual;
     procedure SetX1(const Value: integer); virtual;
     procedure SetX2(const Value: integer); virtual;
     procedure SetY1(const Value: integer); virtual;
     procedure SetY2(const Value: integer); virtual;
     function  GetPadX1: integer;
     function  GetPadX2: integer;
     function  GetPadY1: integer;
     function  GetPadY2: integer;
protected
     FX2           : integer;
     FY2           : integer;
     FX1           : integer;
     FY1           : integer;

     FCX2          : integer;
     FCY2          : integer;
     FCX1          : integer;
     FCY1          : integer;

     FPaddingLeft  : integer;
     FPaddingTop   : integer;
     FPaddingRight : integer;
     FPaddingBottom: integer;

     FGraphicWidth : integer;
public
     function Width: integer; {$ifdef D2006PLUS} inline; {$endif}
     function Height: integer; {$ifdef D2006PLUS} inline; {$endif}
     function ClientWidth: integer; {$ifdef D2006PLUS} inline; {$endif}
     function ClientHeight: integer; {$ifdef D2006PLUS} inline; {$endif}
     function PaginateWidth: integer; virtual; {$ifdef D2006PLUS} inline; {$endif}
     function PaginateHeight: integer; virtual;{$ifdef D2006PLUS} inline; {$endif}
     function PaperHeight: integer; virtual; {$ifdef D2006PLUS} inline; {$endif}

     property X1           : integer read FX1 write SetX1;
     property Y1           : integer read FY1 write SetY1;
     property X2           : integer read FX2 write SetX2;
     property Y2           : integer read FY2 write SetY2;

     property CX1          : integer read FCX1 write SetCX1;
     property CY1          : integer read FCY1 write SetCY1;
     property CX2          : integer read FCX2 write SetCX2;
     property CY2          : integer read FCY2 write SetCY2;

     property PaddingLeft  : integer read FPaddingLeft write FPaddingLeft;
     property PaddingTop   : integer read FPaddingTop write FPaddingTop;
     property PaddingRight : integer read FPaddingRight write FPaddingRight;
     property PaddingBottom: integer read FPaddingBottom write FPaddingBottom;

     property GraphicWidth : integer read FGraphicWidth write FGraphicWidth;

     property PadX1        : integer read GetPadX1;
     property PadY1        : integer read GetPadY1;
     property PadX2        : integer read GetPadX2;
     property PadY2        : integer read GetPadY2;
     end;

type TAXWEditorArea = class(TAXWClientArea)
private
     procedure SetMargBottom(const Value: integer);
     procedure SetMargLeft(const Value: integer);
     procedure SetMargRight(const Value: integer);
     procedure SetMargTop(const Value: integer);
     procedure SetX1(const Value: integer); override;
     procedure SetX2(const Value: integer); override;
     procedure SetY1(const Value: integer); override;
     procedure SetY2(const Value: integer); override;
     procedure SetCX1(const Value: integer); override;
     procedure SetCX2(const Value: integer); override;
     procedure SetCY1(const Value: integer); override;
     procedure SetCY2(const Value: integer); override;
protected
     FMargLeft     : integer;
     FMargTop      : integer;
     FMargRight    : integer;
     FMargBottom   : integer;

     FCellWidth    : integer;
     FCellHeight   : integer;

     FMinX         : integer;
     FMinY         : integer;
     FMaxX         : integer;
     FMaxY         : integer;
public
     procedure ClipToMax;

     function PaginateWidth: integer; override;
     function PaginateHeight: integer; override;
     function PaperHeight: integer; override;

     property MargLeft     : integer read FMargLeft write SetMargLeft;
     property MargTop      : integer read FMargTop write SetMargTop;
     property MargRight    : integer read FMargRight write SetMargRight;
     property MargBottom   : integer read FMargBottom write SetMargBottom;

     property CellWidth    : integer read FCellWidth write FCellWidth;
     property CellHeight   : integer read FCellHeight write FCellHeight;

     property MinX         : integer read FMinX write FMinX;
     property MinY         : integer read FMinY write FMinY;
     property MaxX         : integer read FMaxX write FMaxX;
     property MaxY         : integer read FMaxY write FMaxY;
     end;

type TAXWIndexObjectList = class;

     TAXWIndexObject = class(TObject)
protected
     FOwner: TAXWIndexObjectList;

     FIndex: Integer;
     FDirty: boolean;

     procedure SetDirty(const Value: boolean); virtual;
public
     constructor Create(AOwner: TAXWIndexObjectList);

     function Len: integer; virtual; abstract;

     property Index: integer read FIndex;
     property Dirty: boolean read FDirty write SetDirty;
     end;

     TAXWIndexObjectList = class(TObjectList)
private
     FOwner: TAXWIndexObject;
     FDirty: boolean;
protected
     procedure Notify(Ptr: Pointer; Action: TListNotification); override;
     procedure Reindex;
     function  GetItems(Index: integer): TAXWIndexObject;
     procedure SetDirty(const Value: boolean); virtual;
public
     constructor Create(AOwner: TAXWIndexObject);

     function Add(AObject: TAXWIndexObject): integer;

     property Dirty: boolean read FDirty write SetDirty;
     property Items[Index: integer]: TAXWIndexObject read GetItems; default;
     end;

type TListStack = class(TStack)
private
     function GetItems(Index: integer): Pointer;
protected
public
     property Items[Index: integer]: Pointer read GetItems; default;
     end;

{ TODO correct checking (Unicode lib) }
function CharIsWhitespace(C: AxUCChar): boolean; {$ifdef D2006PLUS} inline; {$endif}
function CharIsWhitespaceOrHardLB(C: AxUCChar): boolean; {$ifdef D2006PLUS} inline; {$endif}
function CharIsSpace(C: AxUCChar): boolean; {$ifdef D2006PLUS} inline; {$endif}
function CharIsBreakingChar(C: AxUCChar): boolean; {$ifdef D2006PLUS} inline; {$endif}
function CharIsPrintable(C: AxUCChar): boolean; {$ifdef D2006PLUS} inline; {$endif}
function CharIsAlphaNum(C: AxUCChar): boolean; {$ifdef D2006PLUS} inline; {$endif}
function WhitespacePos(S: AxUCString): integer;

function ValidateText(S: AxUCString): AxUCString; { inline; }

// Only to be used when reading external strings.
function ParaBreakPos(Pos: integer; S: AxUCString; var Sz: integer): integer; {$ifdef D2006PLUS} inline; {$endif}

{$ifndef DELPHI_2009_OR_LATER}
type TSysCharSet = set of AnsiChar;
function  CharInSet(C: WideChar; const CharSet: TSysCharSet): Boolean;
{$endif}
function  Fork(MinVal,MaxVal,Value: integer): integer; {$ifdef D2006PLUS} inline; {$endif}
function  Within(MinVal,MaxVal,Value: integer): boolean; {$ifdef D2006PLUS} inline; {$endif}
function  CPos(Pos: integer; C: AxUCChar; S: AxUCString): integer; {$ifdef D2006PLUS} inline; {$endif} overload;
//function  CPos(const C: AxUCChar; const S: AxUCString): integer; overload;
function  LeftText(const ASubstring,AString: AxUCString): boolean;
function  SplitAtChar(const C: AxUCChar; var S: AxUCString): AxUCString;
function  CountChars(C: AxUCChar; S: AxUCString): integer; {$ifdef D2006PLUS} inline; {$endif}
function  HexStrToInt(S: AxUCString): integer;
function  RevRGB(RGB: longword): longword; {$ifdef D2006PLUS} inline; {$endif}
procedure SplitRGB(const RGB: longword; out AR,AG,AB: byte); {$ifdef D2006PLUS} inline; {$endif}
function  IntToRoman(AValue: integer): AxUCString;

// ToDo Breaking chars
const TAXWWhitespaceChars = [' ',#9,#10];
// Breaking chars except default (space normally).
const TAXWBreakingChars = ['-'];
const TAXWSpaceChars = [' '];

implementation

{$ifndef DELPHI_2009_OR_LATER}
function CharInSet(C: WideChar; const CharSet: TSysCharSet): Boolean;
begin
  Result := AnsiChar(C) in CharSet;
end;
{$endif}

function Fork(MinVal,MaxVal,Value: integer): integer;
begin
  if Value < MinVal then
    Result := MinVal
  else if Value > MaxVal then
    Result := MaxVal
  else
    Result := Value;
end;

function Within(MinVal,MaxVal,Value: integer): boolean;
begin
  Result := (Value >= MinVal) and (Value <= MaxVal);
end;

function CPos(Pos: integer; C: AxUCChar; S: AxUCString): integer; overload;
begin
  for Result := Pos to Length(S) do begin
    if S[Result] = C then
      Exit;
  end;
  Result := -1;
end;

//function CPos(const C: AxUCChar; const S: AxUCString): integer;
//begin
//  for Result := 1 to Length(S) do begin
//    if S[Result] = C then
//      Exit;
//  end;
//  Result := -1;
//end;

function LeftText(const ASubstring,AString: AxUCString): boolean;
begin
  Result := Copy(AString,1,Length(ASubstring)) = ASubstring;
end;

function SplitAtChar(const C: AxUCChar; var S: AxUCString): AxUCString;
var
  p: integer;
begin
  p := CPos(C,S);
  if p > 0 then begin
    Result := Copy(S,1,p - 1);
    S := Copy(S,p + 1,MAXINT);
  end
  else begin
    Result := S;
    S := '';
  end;
end;

function CountChars(C: AxUCChar; S: AxUCString): integer;
var
  i: integer;
begin
  Result := 0;
  for i := 1 to Length(S) do begin
    if S[i] = C then
      Inc(Result);
  end;
end;

function IsAxUCChar(C: AxUCChar): boolean;
begin
  Result := (Longword(C) and $FF00) <> 0;
end;

function CharIsPrintable(C: AxUCChar): boolean;
begin
  Result := C >= #$0020;
end;

function CharIsWhitespace(C: AxUCChar): boolean;
begin
  Result := not IsAxUCChar(C) and CharInSet(C,TAXWWhitespaceChars);
end;

function CharIsWhitespaceOrHardLB(C: AxUCChar): boolean;
begin
  Result := not IsAxUCChar(C) and (CharInSet(C,TAXWWhitespaceChars) or (C = AXW_CHAR_SPECIAL_HARDLINEBREAK));
end;

function CharIsSpace(C: AxUCChar): boolean;
begin
  Result := not IsAxUCChar(C) and CharInSet(C,TAXWSpaceChars);
end;

function CharIsBreakingChar(C: AxUCChar): boolean;
begin
  Result := not IsAxUCChar(C) and CharInSet(C,TAXWBreakingChars);
end;

function WhitespacePos(S: AxUCString): integer;
begin
  for Result := 1 to Length(S) - 1 do begin
    if CharIsWhitespace(S[Result]) then
      Exit;
  end;
  Result := -1;
end;

// TODO
function ValidateText(S: AxUCString): AxUCString;
var
  i,j: integer;
begin
  SetLength(Result,Length(S));
  j := 0;
  for i := 1 to Length(S) do begin
    if not ((S[i] = AXW_CHAR_SPECIAL_ENDOFPARA) or
            (S[i] = AXW_CHAR_SPECIAL_CELLMARK) or
            (S[i] = AXW_CHAR_SPECIAL_PAGEBREAK) or
            (S[i] = AXW_CHAR_SPECIAL_COLUMNBREAK) or
            (S[i] = AXW_CHAR_SPECIAL_FIELDBEGIN) or
            (S[i] = AXW_CHAR_SPECIAL_FIELDEND) or
            (S[i] = AXW_CHAR_SPECIAL_FIELDSEP) or
            (S[i] = AXW_CHAR_SPECIAL_CELLMARK))
     then begin
       Inc(j);
       Result[j] := S[i];
     end;
  end;
  SetLength(Result,j);
end;

function ParaBreakPos(Pos: integer; S: AxUCString; var Sz: integer): integer;
begin
  Sz := 1;
  Result := CPos(Pos,AxUCChar(#13),S);
  if Result >= 1 then begin
    if (Result < Length(S)) and (S[Result + 1] = #10) then
      Inc(Sz)
    else if (Result > 1) and (S[Result + 1] = #10) then begin
      Dec(Result);
      Inc(Sz);
    end;
  end
  else begin
    Result := CPos(AxUCChar(#10),S);
    if Result >= 1 then begin
      if (Result < Length(S)) and (S[Result + 1] = #13) then
        Inc(Sz)
      else if (Result > 1) and (S[Result + 1] = #13) then begin
        Dec(Result);
        Inc(Sz);
      end;
    end;
  end;
end;

function CharIsAlphaNum(C: AxUCChar): boolean;
begin
  Result := CharInSet(C,['a'..'z','A'..'Z','0'..'9']);
end;

function HexStrToInt(S: AxUCString): integer;
begin
  Result := 0;
  if Length(S) <> 2 then
    Exit;
  S := Uppercase(S);
  if CharInSet(S[1],['0'..'9','A'..'F']) and CharInSet(S[2],['0'..'9','A'..'F']) then begin
    if CharInSet(S[1],['0'..'9']) then
      Result := (Ord(S[1]) - Ord('0')) * 16
    else
      Result := 160 + (Ord(S[1]) - Ord('A')) * 16;
    if CharInSet(S[2],['0'..'9']) then
      Inc(Result,Ord(S[2]) - Ord('0'))
    else
      Inc(Result,10 + (Ord(S[2]) - Ord('A')));
  end;
end;

function RevRGB(RGB: longword): longword; {$ifdef D2006PLUS} inline; {$endif}
begin
  Result := ((RGB and $FF0000) shr 16) + (RGB and $00FF00) + ((RGB and $0000FF) shl 16);
end;

procedure SplitRGB(const RGB: longword; out AR,AG,AB: byte);
begin
  AR := (RGB and $FF0000) shr 16;
  AG := (RGB and $00FF00) shr 8;
  AB := RGB and $0000FF;
end;

function IntToRoman(AValue: integer): AxUCString;
const
  Arabics: Array[0..12] of integer = (1,4,5,9,10,40,50,90,100,400,500,900,1000) ;
  Romans: Array[0..12] of string = ('I','IV','V','IX','X','XL','L','XC','C','CD','D','CM','M') ;
var
  i: integer;
begin
  Result := '';
  for i := High(Arabics) downto 0 do begin
    while (AValue >= Arabics[i]) do begin
      AValue := AValue - Arabics[i];
      Result := Result + Romans[i];
    end;
  end;
end;

{ TAXWArea }

function TAXWClientArea.ClientHeight: integer;
begin
  Result := FCY2 - FCY1;
end;

function TAXWClientArea.ClientWidth: integer;
begin
  Result := FCX2 - FCX1;
end;

function TAXWClientArea.GetPadX1: integer;
begin
  Result := FX1 - FPaddingLeft;
end;

function TAXWClientArea.GetPadX2: integer;
begin
  Result := FX2 + FPaddingRight;
end;

function TAXWClientArea.GetPadY1: integer;
begin
  Result := FY1 - FPaddingTop;
end;

function TAXWClientArea.GetPadY2: integer;
begin
  Result := FY2 + FPaddingBottom;
end;

function TAXWClientArea.Height: integer;
begin
  Result := FY2 - FY1;
end;

function TAXWClientArea.PaginateHeight: integer;
begin
  Result := ClientHeight;
end;

function TAXWClientArea.PaginateWidth: integer;
begin
  Result := ClientWidth - FGraphicWidth;
end;

function TAXWClientArea.PaperHeight: integer;
begin
  Result := ClientHeight;
end;

procedure TAXWClientArea.SetCX1(const Value: integer);
begin
  FCX1 := Value;
end;

procedure TAXWClientArea.SetCX2(const Value: integer);
begin
  FCX2 := Value;
end;

procedure TAXWClientArea.SetCY1(const Value: integer);
begin
  FCY1 := Value;
end;

procedure TAXWClientArea.SetCY2(const Value: integer);
begin
  FCY2 := Value;
end;

procedure TAXWClientArea.SetX1(const Value: integer);
begin
  FX1 := Value;
  if FX1 > FCX1 then
    FCX1 := FX1;
end;

procedure TAXWClientArea.SetX2(const Value: integer);
begin
  FX2 := Value;
  if FX2 < FCX2 then
    FCX2 := FX2;
end;

procedure TAXWClientArea.SetY1(const Value: integer);
begin
  FY1 := Value;
  if FY1 > FCY1 then
    FCY1 := FY1;
end;

procedure TAXWClientArea.SetY2(const Value: integer);
begin
  FY2 := Value;
  if FY2 < FCY2 then
    FCY2 := FY2;
end;

function TAXWClientArea.Width: integer;
begin
  Result := FX2 - FX1;
end;

{ TAXWIndexObjectList }

function TAXWIndexObjectList.Add(AObject: TAXWIndexObject): integer;
begin
  Result := inherited Add(AObject);
  AObject.FIndex := Result;
end;

constructor TAXWIndexObjectList.Create(AOwner: TAXWIndexObject);
begin
  inherited Create;
  FOwner := AOwner;
end;

function TAXWIndexObjectList.GetItems(Index: integer): TAXWIndexObject;
begin
  Result := TAXWIndexObject(inherited Items[Index]);
end;

procedure TAXWIndexObjectList.Notify(Ptr: Pointer; Action: TListNotification);
begin
  inherited;
  case Action of
    lnAdded    : Reindex;
    lnExtracted: Reindex;
    lnDeleted  : Reindex;
  end;
end;

procedure TAXWIndexObjectList.Reindex;
var
  i: integer;
begin
  for i := 0 to Count - 1 do
    TAXWIndexObject(Items[i]).FIndex := i;
end;

procedure TAXWIndexObjectList.SetDirty(const Value: boolean);
begin
  FDirty := Value;
  if Value and (FOwner <> Nil) then
    FOwner.Dirty := Value;
end;

{ TAXWIndexObject }

constructor TAXWIndexObject.Create(AOwner: TAXWIndexObjectList);
begin
  FOwner := AOwner;
end;

procedure TAXWIndexObject.SetDirty(const Value: boolean);
begin
  FDirty := Value;
  if Value and (FOwner <> Nil) then
    FOwner.Dirty := Value;
end;

{ TAXWEditorArea }

procedure TAXWEditorArea.ClipToMax;
begin
  if FX1 < FMinX then FX1 := FMinX;
  if FY1 < FMinY then FY1 := FMinY;
  if FX2 > FMaxX then FX2 := FMaxX;
  if FY1 > FMaxY then FY2 := FMaxY;
end;

function TAXWEditorArea.PaginateHeight: integer;
begin
  Result := FMaxY - FMinY - (FMargTop + FMargBottom);
end;

function TAXWEditorArea.PaginateWidth: integer;
begin
  Result := FMaxX - FMinX - (FMargLeft + FMargRight);
end;

function TAXWEditorArea.PaperHeight: integer;
begin
  Result := CellHeight;
end;

procedure TAXWEditorArea.SetCX1(const Value: integer);
begin
  raise XLSRWException.Create('Can not set client position. Set margins');
end;

procedure TAXWEditorArea.SetCX2(const Value: integer);
begin
  raise XLSRWException.Create('Can not set client position. Set margins');
end;

procedure TAXWEditorArea.SetCY1(const Value: integer);
begin
  raise XLSRWException.Create('Can not set client position. Set margins');
end;

procedure TAXWEditorArea.SetCY2(const Value: integer);
begin
  raise XLSRWException.Create('Can not set client position. Set margins');
end;

procedure TAXWEditorArea.SetMargBottom(const Value: integer);
begin
  FMargBottom := Value;
  FCY2 := FY2 - FMargBottom;
end;

procedure TAXWEditorArea.SetMargLeft(const Value: integer);
begin
  FMargLeft := Value;
  FCX1 := FX1 + FMargLeft;
end;

procedure TAXWEditorArea.SetMargRight(const Value: integer);
begin
  FMargRight := Value;
  FCX2 := FX2 - FMargRight;
end;

procedure TAXWEditorArea.SetMargTop(const Value: integer);
begin
  FMargTop := Value;
  FCY1 := FY1 + FMargTop;
end;

procedure TAXWEditorArea.SetX1(const Value: integer);
begin
  FX1 := Value;
  FCX1 := FX1 + FMargLeft;
end;

procedure TAXWEditorArea.SetX2(const Value: integer);
begin
  FX2 := Value;
  FCX2 := FX2 - FMargRight;
end;

procedure TAXWEditorArea.SetY1(const Value: integer);
begin
  FY1 := Value;
  FCY1 := FY1 + FMargTop;
end;

procedure TAXWEditorArea.SetY2(const Value: integer);
begin
  FY2 := Value;
  FCY2 := FY2 - FMargBottom;
end;

{ TListStack }

function TListStack.GetItems(Index: integer): Pointer;
begin
  Result := List[Index];
end;

end.
