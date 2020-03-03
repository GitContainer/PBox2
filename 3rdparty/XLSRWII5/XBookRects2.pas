unit XBookRects2;

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

uses Classes, SysUtils, vcl.Controls,
     XBookWindows2, XBookPaintGDI2, XBookPaint2, XBookSkin2, XBookUtils2;

type TCornerRect = class(TXSSClientWindow)
protected
public
     procedure MouseMove(Shift: TXSSShiftState; X, Y: Integer); override;
     procedure Paint; override;
     end;

type TSplitterHandleType = (shtHorizontal,shtVertical,shtCenter);

type TSplitterHandle = class(TXSSClientWindow)
protected
     FSplitting: boolean;
     FSplitterType: TSplitterHandleType;
     FCurrX,FCurrY: integer;
     FPaintLineEvent: TXPaintLineEvent;
     FSmall: boolean;
public
     constructor Create(APArent: TXSSWindow; SplitterType: TSplitterHandleType);
     procedure Paint; override;
     procedure MouseDown(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer); override;
     procedure MouseMove(Shift: TXSSShiftState; X, Y: Integer); override;
     procedure MouseUp(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer); override;

     property OnPaintLine: TXPaintLineEvent read FPaintLineEvent write FPaintLineEvent;
     property Small: boolean read FSmall write FSmall;
     end;


implementation

{ TCornerRect }

procedure TCornerRect.MouseMove(Shift: TXSSShiftState; X, Y: Integer);
begin
  inherited MouseMove(Shift,X,Y);
  FSkin.GDI.SetCursor(xctCell);
end;

procedure TCornerRect.Paint;
begin
  inherited Paint;

  if FVisible then begin
    BeginPaint;

    FSkin.PaintHeader(FX1,FY1,FX2,FY2,hssNormal);
    EndPaint;
  end;
end;

{ TSplitterHandle }

constructor TSplitterHandle.Create(AParent: TXSSWindow; SplitterType: TSplitterHandleType);
begin
  inherited Create(AParent);
  FSplitterType := SplitterType;
  FSmall := True;
end;

procedure TSplitterHandle.MouseDown(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer);
begin
  inherited MouseDown(Button,Shift,X,Y);
  FSplitting := True;
  case FSplitterType of
    shtHorizontal: begin
      FCurrX := X;
      FPaintLineEvent(Self,pltHSplit,FCurrX,Y,False);
      FSkin.GDI.SizeLine(FCurrX,FSkin.CY1,FCurrX + SPLITTER_SIZE,FSkin.CY2);
    end;
    shtVertical: begin
      FCurrY := Y;
      FPaintLineEvent(Self,pltVSplit,X,FCurrY,False);
      FSkin.GDI.SizeLine(FSkin.CX1,FCurrY,FSkin.CX2,FCurrY + SPLITTER_SIZE);
    end;
    shtCenter: begin
      FCurrX := X;
      FPaintLineEvent(Self,pltHSplit,FCurrX,Y,False);
      FSkin.GDI.SizeLine(FCurrX,FSkin.CY1,FCurrX + SPLITTER_SIZE,FSkin.CY2);
      FCurrY := Y;
      FPaintLineEvent(Self,pltVSplit,X,FCurrY,False);
      FSkin.GDI.SizeLine(FSkin.CX1,FCurrY,FSkin.CX2,FCurrY + SPLITTER_SIZE);
    end;
  end;
end;

procedure TSplitterHandle.MouseMove(Shift: TXSSShiftState; X, Y: Integer);
begin
  inherited MouseMove(Shift,X,Y);;
  case FSplitterType of
    shtHorizontal: FSkin.GDI.SetCursor(xctHoriz2Split);
    shtVertical  : FSkin.GDI.SetCursor(xctVert2Split);
    shtCenter    : FSkin.GDI.SetCursor(xctCenter);
  end;
  if FSplitting then begin
    case FSplitterType of
      shtHorizontal: begin
        FSkin.GDI.SizeLine(FCurrX,FSkin.CY1,FCurrX + SPLITTER_SIZE,FSkin.CY2);
        FCurrX := X;
        FPaintLineEvent(Self,pltHSplit,FCurrX,Y,False);
        FSkin.GDI.SizeLine(FCurrX,FSkin.CY1,FCurrX + SPLITTER_SIZE,FSkin.CY2);
      end;
      shtVertical: begin
        FSkin.GDI.SizeLine(FSkin.CX1,FCurrY,FSkin.CX2,FCurrY + SPLITTER_SIZE);
        FCurrY := Y;
        FPaintLineEvent(Self,pltVSplit,X,FCurrY,False);
        FSkin.GDI.SizeLine(FSkin.CX1,FCurrY,FSkin.CX2,FCurrY + SPLITTER_SIZE);
      end;
      shtCenter: begin
        FSkin.GDI.SizeLine(FCurrX,FSkin.CY1,FCurrX + SPLITTER_SIZE,FSkin.CY2);
        FSkin.GDI.SizeLine(FSkin.CX1,FCurrY,FSkin.CX2,FCurrY + SPLITTER_SIZE);
        FCurrX := X;
        FCurrY := Y;
        FPaintLineEvent(Self,pltHSplit,FCurrX,Y,False);
        FPaintLineEvent(Self,pltVSplit,X,FCurrY,False);
        FSkin.GDI.SizeLine(FCurrX,FSkin.CY1,FCurrX + SPLITTER_SIZE,FSkin.CY2);
        FSkin.GDI.SizeLine(FSkin.CX1,FCurrY,FSkin.CX2,FCurrY + SPLITTER_SIZE);
      end;
    end;
  end;
end;

procedure TSplitterHandle.MouseUp(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer);
begin
  inherited MouseUp(Button,Shift,X,Y);;
  if FSplitting then begin
    FSplitting := False;
    case FSplitterType of
      shtHorizontal: begin
        FSkin.GDI.SizeLine(FCurrX,FSkin.CY1,FCurrX + SPLITTER_SIZE,FSkin.CY2);
        FPaintLineEvent(Self,pltHSplit,FCurrX,Y,True);
      end;
      shtVertical: begin
        FSkin.GDI.SizeLine(FSkin.CX1,FCurrY,FSkin.CX2,FCurrY + SPLITTER_SIZE);
        FPaintLineEvent(Self,pltVSplit,X,FCurrY,True);
      end;
      shtCenter: begin
        FSkin.GDI.SizeLine(FCurrX,FSkin.CY1,FCurrX + SPLITTER_SIZE,FSkin.CY2);
        FSkin.GDI.SizeLine(FSkin.CX1,FCurrY,FSkin.CX2,FCurrY + SPLITTER_SIZE);
        FPaintLineEvent(Self,pltHSplit,FCurrX,Y,True);
        FPaintLineEvent(Self,pltVSplit,X,FCurrY,True);
      end;
    end;
  end;
end;

procedure TSplitterHandle.Paint;
begin
  inherited Paint;

  BeginPaint;

  if FVisible then
    FSkin.PaintSplitterHandle(FX1,FY1,FX2,FY2,FSplitterType = shtVertical,FSmall);

  EndPaint;
end;

end.
