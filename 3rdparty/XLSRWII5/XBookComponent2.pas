unit XBookComponent2;

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

uses Classes, SysUtils, vcl.Controls, vcl.Graphics, Windows, Messages,
     XLSUtils5,
     XBook_System_2, XBookPaintLayers2, XBookSkin2, XBookWindows2;

type TXSSComponent = class(TCustomControl)
protected
     FSystem    : TXSSSystem;
     FLayers    : TXPaintLayers;
     FSkin      : TXLSBookSkin;
     FRootWin   : TXSSRootWindow;
     FDebugEvent: TNotifyEvent;
     FDebugText : string;

     procedure CreateHandle; override;

     procedure WMSetFocus(var Message: TMessage); message WM_SETFOCUS;
     procedure WMKillFocus(var Message: TMessage); message WM_KILLFOCUS;
     procedure WMMouseLeave(var Message: TMessage); message WM_MOUSELEAVE;
     procedure WMTimer(var Message: TMessage); message WM_TIMER;
     procedure WMMouseWheel(var Message: TWMMouseWheel); message WM_MOUSEWHEEL;
     procedure WMGetDlgCode(var Msg: TWMNoParams); message WM_GETDLGCODE;
     procedure WMChar(var Message: TWMChar); message WM_CHAR;

     procedure CMWantSpecialKey(var Msg: TCMWantSpecialKey); message CM_WANTSPECIALKEY;
     function  CvtShift(const AShift: TShiftState): TXSSShiftState; {$ifdef D2006PLUS} inline; {$endif}
public
     constructor Create(AOwner: TComponent); override;
     destructor Destroy; override;

     procedure WndProc(var Message: TMessage); override;
     procedure MouseDown(Button: TMouseButton; Shift: TShiftState; X, Y: Integer); override;
     procedure MouseMove(Shift: TShiftState; X, Y: Integer); override;
     procedure MouseUp(Button: TMouseButton; Shift: TShiftState; X, Y: Integer); override;

     procedure SetBounds(ALeft, ATop, AWidth, AHeight: Integer); override;
     procedure Paint; override;

     // Public for debug.
     property RootWin: TXSSRootWindow read FRootWin;
     property OnDebug: TNotifyEvent read FDebugEvent write FDebugEvent;
     property DebugText: string read FDebugText;
     end;

implementation

{ TXSSComponent }

procedure TXSSComponent.CMWantSpecialKey(var Msg: TCMWantSpecialKey);
begin
  if Msg.CharCode in [VK_ESCAPE,VK_LEFT,VK_RIGHT,VK_UP,VK_DOWN,VK_PRIOR,VK_NEXT,VK_HOME,VK_END,VK_TAB] then
    Msg.Result := 1;
end;

constructor TXSSComponent.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  // Cursor must be set to anything except crDefault and crArrow. If not, the
  // cursor set by the component will flicker between the selected cursor and
  // the arrow cursor. There is probably a reason for this behaviour.
  Cursor := crCross;
//  ControlStyle := ControlStyle - [csAcceptsControls, csSetCaption] + [csClickEvents, csDoubleClicks, csFramed, csCaptureMouse];
end;

procedure TXSSComponent.CreateHandle;
begin
  inherited;

  if FLayers = Nil then begin
    FSystem := TXSSSystem.Create(Handle);
    FLayers := TXPaintLayers.Create(FSystem);
    FLayers.LayerMode := plmMulti;
    FSkin := TXLSBookSkin.Create(Handle,FLayers.SheetLayer);
    FRootWin := TXSSRootWindow.Create(FSystem,FSkin);
  end;
end;

function TXSSComponent.CvtShift(const AShift: TShiftState): TXSSShiftState;
begin
  Result := [];
  if ssShift  in AShift then Result := Result + [xssShift];
  if ssAlt    in AShift then Result := Result + [xssAlt];
  if ssCtrl   in AShift then Result := Result + [xssCtrl];
  if ssLeft   in AShift then Result := Result + [xssLeft];
  if ssRight  in AShift then Result := Result + [xssRight];
  if ssMiddle in AShift then Result := Result + [xssMiddle];
  if ssDouble in AShift then Result := Result + [xssDouble];
end;

destructor TXSSComponent.Destroy;
begin
  if FLayers <> Nil then begin
    FRootWin.Free;
    FSkin.Free;
    FLayers.Free;
    FSystem.Free;
  end;

  inherited;
end;

procedure TXSSComponent.MouseDown(Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  inherited;
  FRootWin.MouseDown(TXSSMouseButton(Button),CvtShift(Shift),X,Y);
  SetFocus;
end;

procedure TXSSComponent.MouseMove(Shift: TShiftState; X, Y: Integer);
var
  TME: TTrackMouseEvent;
begin
  inherited;
  FRootWin.MouseMove(CvtShift(Shift),X,Y);
  TME.cbSize := SizeOf(TTrackMouseEvent);
  TME.dwFlags := TME_LEAVE;
  TME.hwndTrack := Handle;
  Windows.TrackMouseEvent(TME);

  if Assigned(FDebugEvent) then begin
    FDebugText := Format('X: %d Y: %d',[X,Y]);
    FDebugEvent(Self);
  end;
end;

procedure TXSSComponent.MouseUp(Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  inherited;
  FRootWin.MouseUp(TXSSMouseButton(Button),CvtShift(Shift),X,Y);
end;

procedure TXSSComponent.Paint;
begin
  inherited;
  FRootWin.Paint;
end;

procedure TXSSComponent.SetBounds(ALeft, ATop, AWidth, AHeight: Integer);
begin
  inherited;

  if FLayers <> Nil then
    FRootWin.SetSize(0,0,AWidth,AHeight);
end;

procedure TXSSComponent.WMChar(var Message: TWMChar);
begin
  FRootWin.KeyPress(AxUCChar(Message.CharCode));
end;

procedure TXSSComponent.WMGetDlgCode(var Msg: TWMNoParams);
begin
  Msg.Result := DLGC_WANTARROWS;
end;

procedure TXSSComponent.WMKillFocus(var Message: TMessage);
begin
  FSystem.WinMessage(WM_KILLFOCUS,Message.LParam,Message.WParam);
  FRootWin.KillFocus;
end;

procedure TXSSComponent.WMMouseLeave(var Message: TMessage);
begin
  FRootWin.MouseLeave;
end;

procedure TXSSComponent.WMMouseWheel(var Message: TWMMouseWheel);
var
  Rect: TRect;
begin
  GetWindowRect(Handle,Rect);
  FRootWin.MouseWheel(Message.XPos - Rect.Left,Message.YPos - Rect.Top,Message.WheelDelta);
end;

procedure TXSSComponent.WMSetFocus(var Message: TMessage);
begin
  FSystem.WinMessage(WM_SETFOCUS,Message.LParam,Message.WParam);
  FRootWin.SetFocus;
end;

procedure TXSSComponent.WMTimer(var Message: TMessage);
begin
  FSystem.WinMessage(WM_TIMER,Message.LParam,Message.WParam);
end;

procedure TXSSComponent.WndProc(var Message: TMessage);
begin
  inherited;
  if not (csDestroying in ComponentState) then begin
    if FLayers <> Nil then
      FLayers.ReleaseHandle;
  end;
end;

end.
