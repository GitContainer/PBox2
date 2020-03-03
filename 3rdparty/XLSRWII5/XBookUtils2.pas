unit XBookUtils2;

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

uses {* Delphi  *} Classes, SysUtils, IniFiles,
     {* XLSRWII *} XLSUtils5, Xc12Utils5, Xc12DataStyleSheet5,
     {* XLSBook *} XBookTypes2;

const EX12_AUTO_COLOR = $FF000000;

const CellBorderWidthts: array[cbsNone..cbsSlantedDashDot] of integer = (0,1,2,1,1,3,3,1,2,1,2,1,2,2);

const XSSDefaultFontSizes: array[0..15] of integer = (8,9,10,11,12,14,16,18,20,22,24,26,28,36,48,72);

const PaperSizeNames: array[TXc12PaperSize] of string = (
  'None', 'Letter', 'LetterSmall', 'Tabloid', 'Ledger', 'Legal', 'Statement',
  'Executive', 'A3', 'A4', 'A4Small', 'A5', 'B4', 'B5', 'Folio', 'Quarto',
  '10X14', '11X17', 'Note', 'Env9', 'Env10', 'Env11', 'Env12', 'Env14',
  'CSheet', 'DSheet', 'ESheet', 'EnvDL', 'EnvC5', 'EnvC3', 'EnvC4', 'EnvC6',
  'EnvC65', 'EnvB4', 'EnvB5', 'EnvB6', 'EnvItaly', 'EnvMonarch', 'EnvPersonal',
  'FanfoldUS', 'FanfoldStdGerman', 'FanfoldLglGerman', 'ISO_B4', 'JapanesePostcard',
  '9X11', '10X11', '15X11', 'EnvInvite', 'Reserved48', 'Reserved49', 'LetterExtra',
  'LegalExtra', 'TabloidExtra', 'A4Extra', 'LetterTransverse', 'A4Transverse',
  'LetterExtraTransverse', 'APlus', 'BPlus', 'LetterPlus', 'A4Plus', 'A5Transverse',
  'B5transverse', 'A3Extra', 'A5Extra', 'B5Extra', 'A2', 'A3Transverse', 'A3ExtraTransverse');

type PLongwordArray = ^TLongwordArray;
     TLongwordArray = array [0..0] of longword;

type TPaintLineType = (pltHSize,pltVSize,pltHSplit,pltVSplit);

type TXBookNotification = (xbnNone,xbnCellLocked);

type TXLSProcEvent    = procedure of object;
type TXPaintLineEvent = procedure(Sender: TObject; LineType: TPaintLineType; var X,Y: integer; Execute: boolean) of object;
type TX2IntegerEvent  = procedure(Sender: TObject; Val1,Val2: integer) of object;
type TXStringEvent    = procedure(Sender: TObject; Value: AxUCString) of object;
type TXIntegerEvent   = procedure(Sender: TObject; Value: integer) of object;
type TX4IntegerEvent  = procedure(Sender: TObject; const Val1,Val2,Val3,Val4: integer) of object;
type TXBooleanEvent   = procedure(Sender: TObject; var Value: boolean) of object;
type TXNotifyEvent    = procedure(Sender: TObject; Notification: TXBookNotification) of object;
type TXColRowEvent    = procedure(Sender: TObject; Col,Row: integer) of object;


function  InDelta(Val,Pos,Delta: integer): boolean;
procedure OffsetXYRect(var Rect: TXYRect; DX,DY: integer);
function  PtInXYRect(X,Y: integer; Rect: TXYRect): boolean;
function  SizeXYRect(Rect: TXYRect; Delta: integer): TXYRect;
function  IntersectXYRect(R1, R2: TXYRect; out Dest: TXYRect): Boolean;
function  CellInArea(Col,Row: integer; Area: TXLSCellArea): boolean;
function  _Temp_XCColorToRGB(XColor: longword; AutoColor: longword): longword;

implementation

function InDelta(Val,Pos,Delta: integer): boolean;
begin
  Result := (Val >= (Pos - Delta)) and (Val <= (Pos + Delta));
end;

procedure OffsetXYRect(var Rect: TXYRect; DX,DY: integer);
begin
  Inc(Rect.X1,DX);
  Inc(Rect.Y1,DY);
  Inc(Rect.X2,DX);
  Inc(Rect.Y2,DY);
end;

function PtInXYRect(X,Y: integer; Rect: TXYRect): boolean;
begin
  Result := (X >= Rect.X1) and (X <= Rect.X2) and (Y >= Rect.Y1) and (Y <= Rect.Y2);
end;

function IntersectXYRect(R1, R2: TXYRect; out Dest: TXYRect): Boolean;
begin
  if R1.X1 >= R2.X1 then Dest.X1 := R1.X1 else Dest.X1 := R2.X1;
  if R1.X2 <= R2.X2 then Dest.X2 := R1.X2 else Dest.X2 := R2.X2;
  if R1.Y1 >= R2.Y1 then Dest.Y1 := R1.Y1 else Dest.Y1 := R2.Y1;
  if R1.Y2 <= R2.Y2 then Dest.Y2 := R1.Y2 else Dest.Y2 := R2.Y2;
  Result := (Dest.X2 >= Dest.X1) and (Dest.Y2 >= Dest.Y1);
end;

function CellInArea(Col,Row: integer; Area: TXLSCellArea): boolean;
begin
  Result := (Col >= Area.Col1) and (Col <= Area.Col2) and (Row >= Area.Row1) and (Row <= Area.Row2);
end;

function SizeXYRect(Rect: TXYRect; Delta: integer): TXYRect;
begin
  Result := Rect;
  Dec(Result.X1,Delta);
  Dec(Result.Y1,Delta);
  Inc(Result.X2,Delta);
  Inc(Result.Y2,Delta);
end;

function  _Temp_XCColorToRGB(XColor: longword; AutoColor: longword): longword;
begin
  if XColor = EX12_AUTO_COLOR then
    Result := AutoColor
  else
    Result := XColor;
//  Result := ((Result and $FF) shl 16) + (Result and $00FF00) + (Result shr 16);
end;

end.
