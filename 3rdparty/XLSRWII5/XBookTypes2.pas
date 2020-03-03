unit XBookTypes2;

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

uses Types;

type PXYRect = ^TXYRect;
     TXYRect = record
     X1,Y1,X2,Y2: integer;
     end;

type PXYDWRect = ^TXYDWRect;
     TXYDWRect = record
     X1,Y1,X2,Y2: longword;
     end;

type TXYPoint = record
     X,Y: integer;
     end;

type TXYFRect = record
     X1,Y1,X2,Y2: double;
     end;

type TXYFPoint = record
     X,Y: double;
     end;

type TPointArray = array[0..(MAXINT div SizeOf(TPoint) div 2)] of TPoint;
     PPointArray = ^TPointArray;

type TWordPoint = packed record
     X,Y: word;
     end;

type TWordPointArray = array[0..(MAXINT div SizeOf(TPoint) div 2)] of TWordPoint;
     PWordPointArray = ^TWordPointArray;

function  SetXYRect(const AX1,AY1,AX2,AY2: integer): TXYRect; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif} overload;
procedure SetXYRect(var Rect: TXYRect; X1,Y1,X2,Y2: integer); {$ifdef DELPHI_2006_OR_LATER} inline; {$endif} overload;
function HitXYRect(const AX,AY: integer; const ARect: TXYRect): boolean; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}

implementation

function SetXYRect(const AX1,AY1,AX2,AY2: integer): TXYRect;
begin
  Result.X1 := AX1;
  Result.Y1 := AY1;
  Result.X2 := AX2;
  Result.Y2 := AY2;
end;

procedure SetXYRect(var Rect: TXYRect; X1,Y1,X2,Y2: integer);
begin
  Rect.X1 := X1;
  Rect.Y1 := Y1;
  Rect.X2 := X2;
  Rect.Y2 := Y2;
end;

function HitXYRect(const AX,AY: integer; const ARect: TXYRect): boolean;
begin
  Result := (AX >= ARect.X1) and(AY >= ARect.Y1) and (AX <= ARect.X2) and (AY <= ARect.Y2)
end;

end.
