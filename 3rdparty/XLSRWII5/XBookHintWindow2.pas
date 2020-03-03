unit XBookHintWindow2;

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

uses Classes, SysUtils, vcl.Controls, Types, vcl.Graphics, Windows,
     XLSUtils5, 
     XBookPaint2;

type TXHintWindow = class(TObject)
private
     FParent: TControl;
     FHintWnd: THintWindow;
     FXPaint: TXLSBookPaint;

     function  GetText: AxUCString;
     procedure SetText(const Value: AxUCString);
public
     constructor Create(Parent: TControl; XPaint: TXLSBookPaint);
     destructor Destroy; override;
     procedure Show(X,Y: integer);
     procedure Hide;

     property Text: AxUCString read GetText write SetText;
     end;

implementation

{ TXHintWindow }

constructor TXHintWindow.Create(Parent: TControl; XPaint: TXLSBookPaint);
begin
  FParent := Parent;
  FXPaint := XPaint;
  FHintWnd := THintWindow.Create(FParent);
  FHintWnd.Color := clInfoBk;
end;

destructor TXHintWindow.Destroy;
begin
  FHintWnd.Free;
  inherited;
end;

function TXHintWindow.GetText: AxUCString;
begin
  Result := FHintWnd.Caption;
end;

procedure TXHintWindow.Hide;
begin
  ShowWindow(FHintWnd.Handle, SW_HIDE);
  FHintWnd.Invalidate;
end;

procedure TXHintWindow.SetText(const Value: AxUCString);
begin
  FHintWnd.Caption := Value;
  FHintWnd.Invalidate;
end;

procedure TXHintWindow.Show(X, Y: integer);
var
  R: TRect;
  Pt: TPoint;
begin
  R := FHintWnd.CalcHintRect(1024,FHintWnd.Caption,Nil);
  Pt := FParent.ClientToScreen(Point(X,Y));
  OffsetRect(R,Pt.X,Pt.Y);
  FHintWnd.ActivateHint(R,FHintWnd.Caption);
end;

end.
