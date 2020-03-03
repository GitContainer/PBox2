unit XBook_System_2;

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

uses Classes, SysUtils, Contnrs, Windows, Messages;

const XSS_SYS_COLOR_WHITE    = $00FFFFFF;
const XSS_SYS_COLOR_BLACK    = $00000000;
const XSS_SYS_COLOR_GRAY     = $00C0C0C0;

const XSS_SYS_MINTIMERID     = 2;
const XSS_SYS_MAXTIMERID     = High(Word);
const XSS_SYS_CARET_BLINK    = 300;

type TXSSTimerEvent = procedure (ASender: TObject; AData: Pointer) of object;

type TXSSCaret = record
     Width  : integer;
     Height : integer;
     X      : integer;
     Y      : integer;
     Visible: boolean;
     Enabled: boolean;
     end;

type TXSSTimerObject = class(TObject)
protected
     FOwner: TObject;
     FEvent: TXSSTimerEvent;
     FId   : integer;
     FData : Pointer;
public
     constructor Create(AOwner: TObject; AEvent: TXSSTimerEvent; AData: Pointer);

     property Owner: TObject read FOwner;
     property Event: TXSSTimerEvent read FEvent;
     property Id: integer read FId;
     property Data : Pointer read FData;
     end;

type TXSSTimerObjects = class(TObjectList)
private
     FCurrTimerId: integer;

     function GetItems(const Index: integer): TXSSTimerObject;
protected
public
     constructor Create;

     function  Add(AOwner: TObject; AEvent: TXSSTimerEvent; AData: Pointer): TXSSTimerObject;

     function  Find(AObject: TObject): TXSSTimerObject; overload;
     function  Find(const AId: integer): TXSSTimerObject; overload;
     function  FindIndex(AObject: TObject): integer; overload;
     function  FindIndex(AId: integer): integer; overload;

     property Items[const Index: integer]: TXSSTimerObject read GetItems; default;
     end;

type TXSSSystem = class(TObject)
protected
     FHandle: THandle;

     FTimers: TXSSTimerObjects;
     FCaret : TXSSCaret;
public
     constructor Create(const AHandle: THandle);
     destructor Destroy; override;

     function  WinMessage(const AMsg: integer; ALParam: longword; AWParam: word): longword;

     function  StartTimer(AOwner: TObject; AEvent: TXSSTimerEvent; const ATimeout: integer; AData: Pointer): integer;
     procedure StopTimer(AOwner: TObject); overload;
     procedure StopTimer(AId: integer); overload;
     procedure StopAllTimers;

     procedure CreateCaret;
     procedure DestroyCaret;
     procedure HideCaret;
     procedure ShowCaret; overload;
     procedure ShowCaret(const AWidth,AHeight,AX,AY: integer); overload;

     // "Process messages"
     procedure ProcessRequests;

     property Handle: THandle read FHandle;
     end;

implementation

{ TXSSSystem }

procedure TXSSSystem.StopAllTimers;
var
  i: integer;
begin
  for i := 0 to FTimers.Count - 1 do
    Windows.KillTimer(FHandle,FTimers[i].Id);
  FTimers.Clear;
end;

procedure TXSSSystem.StopTimer(AId: integer);
var
  i: integer;
begin
  i := FTimers.FindIndex(AId);
  if i >= 0 then begin
    Windows.KillTimer(FHandle,FTimers[i].Id);
    FTimers.Delete(i);
  end;
end;

constructor TXSSSystem.Create(const AHandle: THandle);
begin
  FHandle := AHandle;

  FTimers := TXSSTimerObjects.Create;

  FCaret.Width := 1;
  FCaret.Height := 16;
  FCaret.X := 0;
  FCaret.Y := 0;
  FCaret.Visible := False;
  FCaret.Enabled := False;
end;

procedure TXSSSystem.CreateCaret;
begin
  Windows.CreateCaret(FHandle,0,FCaret.Width,FCaret.Height);
  FCaret.Visible := False;
  FCaret.Enabled := True;
end;

destructor TXSSSystem.Destroy;
begin
  FTimers.Free;
  inherited;
end;

procedure TXSSSystem.DestroyCaret;
begin
  Windows.DestroyCaret;
  FCaret.Visible := False;
  FCaret.Enabled := False;
end;

procedure TXSSSystem.HideCaret;
begin
  if FCaret.Enabled and FCaret.Visible then begin
    Windows.HideCaret(FHandle);
    Windows.DestroyCaret;
    FCaret.Visible := False;
  end;
end;

procedure TXSSSystem.ProcessRequests;
begin
  // This will send a call to the message loop and pending requests such as
  // layer repaints will be executed.
  Windows.SendMessage(FHandle,WM_USER,0,0);
end;

procedure TXSSSystem.ShowCaret(const AWidth, AHeight, AX, AY: integer);
begin
  FCaret.Width := AWidth;
  FCaret.Height := AHeight;
  FCaret.X := AX;
  FCaret.Y := AY;
  if FCaret.Enabled and FCaret.Visible then begin
    if (FCaret.Width <> AWidth) or (FCaret.Height <> AHeight) then begin
      Windows.DestroyCaret;
      Windows.CreateCaret(FHandle,0,AWidth,AHeight);
      Windows.ShowCaret(FHandle);
    end;
    FCaret.Visible := True;

    Windows.SetCaretPos(FCaret.X,FCaret.Y);
  end;
end;

procedure TXSSSystem.ShowCaret;
begin
  if FCaret.Enabled and  not FCaret.Visible then begin
    FCaret.Visible := True;
    Windows.CreateCaret(FHandle,0,FCaret.Width,FCaret.Height);
    Windows.ShowCaret(FHandle);

    Windows.SetCaretPos(FCaret.X,FCaret.Y);
  end;
end;

function TXSSSystem.StartTimer(AOwner: TObject; AEvent: TXSSTimerEvent; const ATimeout: integer; AData: Pointer): integer;
var
  Timer: TXSSTimerObject;
begin
  if ATimeout > 0 then begin
    Timer := FTimers.Add(AOwner,AEvent,AData);

    Windows.SetTimer(FHandle,Timer.Id,ATimeout,Nil);

    Result := Timer.Id;
  end
  else
    Result := -1;
end;

procedure TXSSSystem.StopTimer(AOwner: TObject);
var
  i: integer;
begin
  i := FTimers.FindIndex(AOwner);
  if i >= 0 then begin
    Windows.KillTimer(FHandle,FTimers[i].Id);
    FTimers.Delete(i);
  end;
end;

function TXSSSystem.WinMessage(const AMsg: integer; ALParam: longword; AWParam: word): longword;
var
  Timer: TXSSTimerObject;
begin
  case AMsg of
    WM_SETFOCUS  : ShowCaret;
    WM_KILLFOCUS : HideCaret;
    WM_TIMER: begin
      Timer := FTimers.Find(AWParam);
      if Timer <> Nil then
        Timer.Event(Timer.Owner,Timer.Data);
    end;
  end;
  Result := 0;
end;

{ TXSSTimerObject }

constructor TXSSTimerObject.Create(AOwner: TObject; AEvent: TXSSTimerEvent; AData: Pointer);
begin
  FOwner := AOwner;
  FEvent := AEvent;
  FData := AData;
end;

{ TXSSTimerObjects }

function TXSSTimerObjects.Add(AOwner: TObject; AEvent: TXSSTimerEvent; AData: Pointer): TXSSTimerObject;
begin
  Inc(FCurrTimerId);
  if FCurrTimerId > XSS_SYS_MAXTIMERID then
    FCurrTimerId := XSS_SYS_MINTIMERID;

  Result := TXSSTimerObject.Create(AOwner,AEvent,AData);
  Result.FId := FCurrTimerId;
  inherited Add(Result);
end;

constructor TXSSTimerObjects.Create;
begin
  inherited Create;

  FCurrTimerId := XSS_SYS_MINTIMERID;
end;

function TXSSTimerObjects.Find(const AId: integer): TXSSTimerObject;
var
  i: integer;
begin
  for i := 0 to Count - 1 do begin
    if Items[i].Id = AId then begin
      Result := Items[i];
      Exit;
    end;
  end;
  Result := Nil;
end;

function TXSSTimerObjects.FindIndex(AId: integer): integer;
begin
  for Result := 0 to Count - 1 do begin
    if Items[Result].Id = AId then
      Exit;
  end;
  Result := -1;
end;

function TXSSTimerObjects.FindIndex(AObject: TObject): integer;
begin
  for Result := 0 to Count - 1 do begin
    if Items[Result].Owner = AObject then
      Exit;
  end;
  Result := -1;
end;

function TXSSTimerObjects.Find(AObject: TObject): TXSSTimerObject;
var
  i: integer;
begin
  for i := 0 to Count - 1 do begin
    if Items[i].Owner = AObject then begin
      Result := Items[i];
      Exit;
    end;
  end;
  Result := Nil;
end;

function TXSSTimerObjects.GetItems(const Index: integer): TXSSTimerObject;
begin
  Result := TXSSTimerObject(inherited Items[Index]);
end;

end.
