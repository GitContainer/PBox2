unit XBookWindows2;

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

uses Classes, SysUtils, Contnrs,
     XLSUtils5,
     XBook_System_2, XBookPaintGDI2, XBookSkin2;

type TXSSMouseButton = (xmbLeft,xmbRight,xmbMiddle);
type TXSSShiftState = set of (xssShift,xssAlt,xssCtrl,xssLeft,xssRight,xssMiddle,xssDouble,xssTouch,xssPen,xssCommand);

type TXSSKey = (kyNone,kyEscape,kyTab,kyLeft,kyRight,kyUp,kyDown,kyPgUp,kyPgDown,kyHome,
                kyEnd,kyFirstCol,kyLastCol,kyFirstRow,kyLastRow,kyInplaceEdit,kyCopy,
                kyCut,kyPaste,kyPasteSpecial,kyDelete);

type TXSSShiftKey = (skShift,skCtrl,skAlt,skMouseLeft,skMouseRight,skMouseMiddle,skMouseDouble);
type TXSSShiftKeys = set of TXSSShiftKey;

type TXSSWindows = class;

     TXSSWindow = class(TObject)
protected
     FSystem  : TXSSSystem;
     FSkin    : TXLSBookSkin;
     FParent  : TXSSWindow;

     FChilds  : TXSSWindows;

     FX1,FY1,
     FX2,FY2  : integer;

     FVisible : boolean;
     FEnabled : boolean;
     FIsOnScr : boolean;
     FColor   : longword;
     FCursor  : TXLSCursorType;

     FId      : integer;

     procedure BeginPaint; virtual;
     procedure EndPaint; virtual;
     function  RootWindow: TXSSWindow;
public
     constructor Create; overload;
     constructor Create(ASkin: TXLSBookSkin); overload;
     constructor Create(AParent: TXSSWindow); overload;
     destructor Destroy; override;

     procedure DeleteChild(AChild: TXSSWindow);
     procedure DeleteChilds; virtual;

     function  FindWindow(const AX,AY: integer): TXSSWindow;
     procedure Add(AWin: TXSSWindow);

     procedure Paint; virtual;
     // Repaints the window and childs;
     procedure Repaint; virtual;
     // Same as Repaint, but inherited windows shall reload any data they depend on.
     procedure Invalidate; virtual;

     procedure SetSize(const AX1,AY1,AX2,AY2: integer); virtual;
     function  Hit(const AX,AY: integer): boolean; virtual;

     procedure CaptureMouse(const ACapture: boolean);

     procedure MouseEnter(X, Y: Integer); virtual;
     procedure MouseLeave; virtual;
     procedure MouseDown(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer); virtual;
     procedure MouseMove(Shift: TXSSShiftState; X, Y: Integer); virtual;
     procedure MouseUp(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer); virtual;
     procedure MouseWheel(const X,Y,Delta: Integer); virtual;

     procedure Focus;
     procedure SetFocus; virtual;
     procedure KillFocus; virtual;

     procedure KeyDown(Key: Word; Shift: TXSSShiftState); virtual;
     procedure KeyPress(Key: AxUCChar); virtual;

     procedure HandleKey(const AKey: TXSSKey; const AShift: TXSSShiftKeys); virtual;

     function  Width: integer;
     function  Height: integer;

     property X1: integer read FX1 write FX1;
     property Y1: integer read FY1 write FY1;
     property X2: integer read FX2 write FX2;
     property Y2: integer read FY2 write FY2;

     property Skin: TXLSBookSkin read FSkin;

     property Visible: boolean read FVisible write FVisible;
     property Enabled: boolean read FEnabled write FEnabled;
     property Color: longword read FColor write FColor;
     property Cursor: TXLSCursorType read FCursor write FCursor;

     property Id: integer read FId write FId;
//     property Childs: TXSSWindows read FChilds;
     end;

     TXSSWindows = class(TObject)
private
     function GetItems(const Index: integer): TXSSWindow;
protected
     FItems   : TObjectList;
     FSystem  : TXSSSystem;
     FSkin    : TXLSBookSkin;
public
     constructor Create(ASystem: TXSSSystem; ASkin: TXLSBookSkin);
     destructor Destroy; override;

     procedure Clear;

     procedure Add(AWin: TXSSWindow);
     procedure Delete(AWin: TXSSWindow);

     function  Count: integer;

     property Items[const Index: integer]: TXSSWindow read GetItems; default;
     end;

type TXSSClientWindow = class(TXSSWindow)
protected
     FClipHandle: longword;

     FCX1,FCY1,
     FCX2,FCY2: integer;
public
     procedure BeginPaint; override;
     procedure BeginPaintClipParent;
     procedure EndPaint; override;

     function  ClientHit(const AX,AY: integer): boolean; virtual;

     procedure SetSize(const AX1,AY1,AX2,AY2: integer); override;
     procedure SetClientSize(const AX1,AY1,AX2,AY2: integer); virtual;

     property CX1: integer read FCX1;
     property CY1: integer read FCY1;
     property CX2: integer read FCX2;
     property CY2: integer read FCY2;
     end;

type TXSSRootWindow = class(TXSSClientWindow)
protected
     FFocusedWin   : TXSSWindow;
     FMouseWin     : TXSSWindow;
     FMouseCaptured: TXSSWindow;
     FMouseIsDown  : boolean;

     // Don't use FindWindow as that may return the root window itself.
     function  FindChild(const AX,AY: integer): TXSSWindow;
     function  FindMouseChild(const AX,AY: integer): TXSSWindow;
     procedure RootMouseDown(AChilds: TXSSWindows; Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer);
     procedure RootPaint(AChilds: TXSSWindows);
     procedure RootCaptureMouse(const ACapture: boolean; AWin: TXSSWindow);
public
     constructor Create(ASystem: TXSSSystem; ASkin: TXLSBookSkin);

     procedure Clear;

     procedure KillFocus; override;
     procedure SetFocus; overload; override;
     procedure SetFocusWin(AWindow: TXSSWindow);

     procedure Paint; override;

     procedure MouseLeave; override;
     procedure MouseDown(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer); override;
     procedure MouseMove(Shift: TXSSShiftState; X, Y: Integer); override;
     procedure MouseUp(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer); override;
     procedure MouseWheel(const X,Y,Delta: Integer); override;

     procedure KeyDown(Key: Word; Shift: TXSSShiftState); override;
     procedure KeyPress(Key: AxUCChar); override;

     procedure HandleKey(const AKey: TXSSKey; const AShift: TXSSShiftKeys); override;

     property FocusedWin: TXSSWindow read FFocusedWin;
     end;

implementation

{$ifdef _AXOLOT_DEBUG}
var
  L_ActiveClip: boolean;
{$endif}

{ TXSSWindow }

procedure TXSSWindow.Add(AWin: TXSSWindow);
begin
  FChilds.Add(AWin);
end;

procedure TXSSWindow.BeginPaint;
begin

end;

procedure TXSSWindow.CaptureMouse(const ACapture: boolean);
var
  Win: TXSSWindow;
begin
  Win := RootWindow;
  if Win <> Nil then
    TXSSRootWindow(Win).RootCaptureMouse(ACapture,Self);
end;

constructor TXSSWindow.Create;
begin
  FChilds := TXSSWindows.Create(FSystem,FSkin);

  FVisible := True;
  FEnabled := True;
  FColor := XSS_SYS_COLOR_WHITE;
  FIsOnScr := True;
end;

constructor TXSSWindow.Create(ASkin: TXLSBookSkin);
begin
  FChilds := TXSSWindows.Create(FSystem,FSkin);

  FSkin := ASkin;

  FVisible := True;
  FEnabled := True;
  FColor := XSS_SYS_COLOR_WHITE;
  FIsOnScr := True;
end;

constructor TXSSWindow.Create(AParent: TXSSWindow);
begin
  FParent := AParent;

  FChilds := TXSSWindows.Create(FSystem,FSkin);

  FSystem := FParent.FSystem;
  FSkin := FParent.FSkin;

  FVisible := True;
  FEnabled := True;
  FColor := XSS_SYS_COLOR_WHITE;
  FIsOnScr := True;
end;

procedure TXSSWindow.DeleteChild(AChild: TXSSWindow);
var
  IsFocused: boolean;
begin
  IsFocused := AChild = TXSSRootWindow(RootWindow).FFocusedWin;
  FChilds.Delete(AChild);
  if IsFocused then begin
    TXSSRootWindow(RootWindow).FFocusedWin := Nil;
    Self.Focus;
  end;
end;

procedure TXSSWindow.DeleteChilds;
begin
  FChilds.Clear;
end;

destructor TXSSWindow.Destroy;
begin
  FChilds.Free;
  inherited;
end;

procedure TXSSWindow.EndPaint;
begin

end;

function TXSSWindow.FindWindow(const AX, AY: integer): TXSSWindow;
var
  i: integer;
begin
  for i := FChilds.Count - 1 downto 0 do begin
    Result := FChilds[i].FindWindow(AX,AY);
    if (Result <> Nil) and Result.Enabled and Result.Visible then
      Exit;
  end;
  if Hit(AX,AY) then
    Result := Self
  else
    Result := Nil;
end;

procedure TXSSWindow.Focus;
begin
  TXSSRootWindow(RootWindow).SetFocusWin(Self);
end;

procedure TXSSWindow.HandleKey(const AKey: TXSSKey; const AShift: TXSSShiftKeys);
begin

end;

function TXSSWindow.Height: integer;
begin
  Result := FY2 - FY1;
end;

function TXSSWindow.Hit(const AX, AY: integer): boolean;
begin
  Result := (AX >= FX1) and (AX <= FX2) and (AY >= FY1) and (AY <= FY2);
end;

procedure TXSSWindow.Repaint;
var
  Win: TXSSWindow;
begin
  Paint;

  Win := RootWindow;
  if Win <> Nil then
    TXSSRootWindow(Win).RootPaint(FChilds);

  if FSystem <> Nil then
    FSystem.ProcessRequests;
end;

procedure TXSSWindow.KeyDown(Key: Word; Shift: TXSSShiftState);
begin

end;

procedure TXSSWindow.KeyPress(Key: AxUCChar);
begin

end;

procedure TXSSWindow.KillFocus;
begin

end;

procedure TXSSWindow.MouseDown(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer);
begin

end;

procedure TXSSWindow.MouseEnter(X, Y: Integer);
begin

end;

procedure TXSSWindow.MouseLeave;
begin

end;

procedure TXSSWindow.MouseMove(Shift: TXSSShiftState; X, Y: Integer);
begin

end;

procedure TXSSWindow.MouseUp(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer);
begin

end;

procedure TXSSWindow.MouseWheel(const X, Y, Delta: Integer);
begin

end;

procedure TXSSWindow.Paint;
begin
//  FSkin.GDI.PaintColor := FColor;
//  FSkin.GDI.Rectangle(FX1,FY1,FX2,FY2);
end;

function TXSSWindow.RootWindow: TXSSWindow;
begin
  Result := Self;
  while Result.FParent <> Nil do
    Result := Result.FParent;
  if not (Result is TXSSRootWindow) then
    Result := Nil;
end;

procedure TXSSWindow.SetFocus;
begin

end;

procedure TXSSWindow.SetSize(const AX1, AY1, AX2, AY2: integer);
begin
  FX1 := AX1;
  FY1 := AY1;
  FX2 := AX2;
  FY2 := AY2;
end;

function TXSSWindow.Width: integer;
begin
  Result := FX2 - FX1;
end;

procedure TXSSWindow.Invalidate;
begin
  Repaint;
end;

{ TXSSRootWindow }

constructor TXSSRootWindow.Create(ASystem: TXSSSystem; ASkin: TXLSBookSkin);
begin
  FSystem := ASystem;
  FSkin := ASkin;

  inherited Create;
end;

function TXSSRootWindow.FindChild(const AX,AY: integer): TXSSWindow;
var
  i: integer;
begin
  for i := FChilds.Count - 1 downto 0 do begin
    Result := FChilds[i].FindWindow(AX,AY);
    if Result <> Nil then
      Exit;
  end;
  Result := Nil;
end;

function TXSSRootWindow.FindMouseChild(const AX, AY: integer): TXSSWindow;
begin
  if FMouseCaptured <> Nil then
    Result := FMouseCaptured
  else
    Result := FindChild(AX,AY);
end;

procedure TXSSRootWindow.HandleKey(const AKey: TXSSKey; const AShift: TXSSShiftKeys);
begin
  if FFocusedWin <> Nil then
    FFocusedWin.HandleKey(AKey,AShift);
end;

procedure TXSSRootWindow.KeyDown(Key: Word; Shift: TXSSShiftState);
begin
  if FFocusedWin <> Nil then
    FFocusedWin.KeyDown(Key,Shift);
end;

procedure TXSSRootWindow.KeyPress(Key: AxUCChar);
begin
  if FFocusedWin <> Nil then
    FFocusedWin.KeyPress(Key);
end;

procedure TXSSRootWindow.KillFocus;
begin
  if FFocusedWin <> Nil then
    FFocusedWin.KillFocus;
//  FFocusedWin := Nil;
  FMouseWin := Nil;
  FMouseCaptured := Nil;
end;

procedure TXSSRootWindow.MouseDown(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer);
var
  Win: TXSSWindow;
begin
  FMouseIsDown := True;

  Win := FindMouseChild(X,Y);
  if Win <> Nil then begin
    if (FFocusedWin <> Nil) and (FFocusedWin <> Win) then
      FFocusedWin.KillFocus;
    FFocusedWin := Win;
    FFocusedWin.MouseDown(Button,Shift,X,Y);
    FFocusedWin.SetFocus;
  end;

// Before uncomment this, check that the windows don't receive double clicks (TXTabSet).
//  RootMouseDown(FChilds,Button,Shift,X,Y);
end;

procedure TXSSRootWindow.RootCaptureMouse(const ACapture: boolean; AWin: TXSSWindow);
begin
  if ACapture then
    FMouseCaptured := AWin
  else
    FMouseCaptured := Nil;
end;

procedure TXSSRootWindow.RootMouseDown(AChilds: TXSSWindows; Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer);
var
  i: integer;
  Win: TXSSWindow;
begin
  FFocusedWin := Nil;
  for i := AChilds.Count - 1 downto 0 do begin
    Win := AChilds[i];
    if Win.Visible and Win.Enabled then begin
      if Win.FChilds.Count > 0 then
        RootMouseDown(Win.FChilds,Button,Shift,X,Y);
      if (FFocusedWin = Nil) and Win.Hit(X,Y) then begin
        FFocusedWin := Win;
        FFocusedWin.MouseDown(Button,Shift,X,Y);
      end
    end;
    if FFocusedWin <> Nil then
      Break;
  end;
end;

procedure TXSSRootWindow.MouseLeave;
begin
  if not FMouseIsDown and (FMouseWin <> Nil) then begin
    FMouseWin.MouseLeave;
    FMouseWin := Nil;
  end;
end;

procedure TXSSRootWindow.MouseMove(Shift: TXSSShiftState; X, Y: Integer);
var
  Win: TXSSWindow;
begin
  if not FMouseIsDown then begin
    Win := FindMouseChild(X,Y);
    if Win <> FMouseWin then begin
      if FMouseWin <> Nil then
        FMouseWin.MouseLeave;
      if (Win <> Nil) and Win.Enabled then
        Win.MouseEnter(X,Y);
      FMouseWin := Win;
    end;
  end;
  if (FMouseWin <> Nil) and FMouseWin.Enabled then
    FMouseWin.MouseMove(Shift,X,Y);
end;

procedure TXSSRootWindow.MouseUp(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer);
begin
  FMouseIsDown := False;

  if FFocusedWin <> Nil then begin
    FFocusedWin.MouseUp(Button,Shift,X,Y);
//    FClickWin := Nil;
  end;
end;

procedure TXSSRootWindow.MouseWheel(const X, Y, Delta: Integer);
var
  Win: TXSSWindow;
begin
  Win := FindMouseChild(X,Y);
  if Win <> Nil then
    Win.MouseWheel(X,Y,Delta);
end;

procedure TXSSRootWindow.Paint;
begin
  RootPaint(FChilds);
end;

procedure TXSSRootWindow.RootPaint(AChilds: TXSSWindows);
var
  i: integer;
  Win: TXSSWindow;
begin
  for i := 0 to AChilds.Count - 1 do begin
    Win := AChilds[i];
    if Win.Visible and Win.FIsOnScr then begin
      Win.Paint;
      if Win.FChilds.Count > 0 then
        RootPaint(Win.FChilds);
    end;
  end;
end;

procedure TXSSRootWindow.SetFocus;
begin
  if FFocusedWin <> Nil then
    FFocusedWin.SetFocus;
end;

procedure TXSSRootWindow.SetFocusWin(AWindow: TXSSWindow);
begin
  if FFocusedWin <> Nil then
    FFocusedWin.KillFocus;

  FFocusedWin := AWindow;
  FFocusedWin.SetFocus;
end;

procedure TXSSRootWindow.Clear;
begin
  FFocusedWin := Nil;
  FMouseWin := Nil;
  FMouseCaptured := Nil;
  FMouseIsDown := False;
end;

{ TXSSWindows }

procedure TXSSWindows.Add(AWin: TXSSWindow);
begin
  FItems.Add(AWin);
end;

procedure TXSSWindows.Clear;
begin
  FItems.Clear;
end;

function TXSSWindows.Count: integer;
begin
  Result := FItems.Count;

end;

constructor TXSSWindows.Create(ASystem: TXSSSystem; ASkin: TXLSBookSkin);
begin
  FSystem := ASystem;
  FSkin := ASkin;

  FItems := TObjectList.Create;
end;

procedure TXSSWindows.Delete(AWin: TXSSWindow);
var
  i: integer;
begin
  i := FItems.IndexOf(AWin);
  if i >= 0 then
    FItems.Delete(i);
end;

destructor TXSSWindows.Destroy;
begin
  FItems.Free;
  inherited;
end;

function TXSSWindows.GetItems(const Index: integer): TXSSWindow;
begin
  Result := TXSSWindow(FItems[Index]);
end;

{ TXSSClientWindow }

procedure TXSSClientWindow.BeginPaint;
begin
{$ifdef _AXOLOT_DEBUG}
  if L_ActiveClip then
    raise XLSRWException.Create('Active Clip');
{$endif}
  FClipHandle := FSkin.GDI.CreateClipRect(FCX1,FCY1,FCX2 + 1,FCY2 + 1);
//  FSkin.GDI.SelectClipRect(FClipHandle);
{$ifdef _AXOLOT_DEBUG}
  L_ActiveClip := True;
{$endif}
end;

procedure TXSSClientWindow.BeginPaintClipParent;
begin
{$ifdef _AXOLOT_DEBUG}
  if L_ActiveClip then
    raise XLSRWException.Create('Active Clip');
{$endif}
  FClipHandle := FSkin.GDI.CreateClipRect(TXSSClientWindow(FParent).CX1,TXSSClientWindow(FParent).FCY1,TXSSClientWindow(FParent).FCX2 + 1,TXSSClientWindow(FParent).FCY2 + 1);
//  FSkin.GDI.SelectClipRect(FClipHandle);
{$ifdef _AXOLOT_DEBUG}
  L_ActiveClip := True;
{$endif}
end;

function TXSSClientWindow.ClientHit(const AX, AY: integer): boolean;
begin
  Result := (AX >= FCX1) and (AX <= FCX2) and (AY >= FCY1) and (AY <= FCY2);
end;

procedure TXSSClientWindow.EndPaint;
begin
{$ifdef _AXOLOT_DEBUG}
  if not L_ActiveClip then
    raise XLSRWException.Create('No active Clip');
{$endif}
  FSkin.GDI.DeleteClipRect(FClipHandle);
  FClipHandle := 0;
{$ifdef _AXOLOT_DEBUG}
  L_ActiveClip := False;
{$endif}
end;

procedure TXSSClientWindow.SetClientSize(const AX1,AY1,AX2,AY2: integer);
begin
  FCX1 := AX1;
  FCY1 := AY1;
  FCX2 := AX2;
  FCY2 := AY2;
end;

procedure TXSSClientWindow.SetSize(const AX1, AY1, AX2, AY2: integer);
begin
  inherited SetSize(AX1,AY1,AX2,AY2);
  SetClientSize(AX1,AY1,AX2,AY2);
end;

{$ifdef _AXOLOT_DEBUG}
initialization
  L_ActiveClip := False;
{$endif}

end.
