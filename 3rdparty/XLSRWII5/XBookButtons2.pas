unit XBookButtons2;

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

uses Classes, SysUtils, Contnrs, vcl.Controls,
     XBookUtils2, XBookPaintGDI2, XBookPaint2, XBookSkin2, XBookWindows2;

type TButtonRect = class(TObject)
protected
     FX1,FY1,FX2,FY2: integer;
     FIndex: integer;
     FButtonState: TBmpButtonState;
     FIdA,FIdB: integer;
public
     function  MouseHit(X,Y: integer): boolean;

     property X: integer read FX1 write FX1;
     property Y: integer read FY1 write FY1;
     property Index: integer read FIndex write FIndex;
     property ButtonState: TBmpButtonState read FButtonState write FButtonState;
     property IdA: integer read FIdA write FIdA;
     property IdB: integer read FIdB write FIdB;
     end;

type TButtonRectList = class(TObjectList)
private
     FButtonType: TBmpButtonType;
     FSkin: TXLSBookSkin;
     FCurrButton: integer;
     FButtonClickEvent: TX2IntegerEvent;

     function GetItems(Index: integer): TButtonRect;
public
     constructor Create(Skin: TXLSBookSkin; ButtonType: TBmpButtonType);
     procedure Add(pX,pY,Index: integer; IdA: integer = -1; IdB: integer = -1);
     procedure Paint;
     function  MouseHit(X,Y: integer): integer;
     procedure MouseDown(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer);
     procedure MouseMove(Shift: TXSSShiftState; X, Y: Integer);
     procedure MouseUp(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer);
     procedure MouseLeave;

     property Items[Index: integer]: TButtonRect read GetItems; default;
     property ButtonType: TBmpButtonType read FButtonType;
     property OnButtonClick: TX2IntegerEvent read FButtonClickEvent write FButtonClickEvent;
     end;

type TGroupButtons = class(TXSSClientWindow)
private
     FButtons: TButtonRectList;
     FHorizontal: boolean;

     function  GetLevels: integer;
     procedure SetLevels(const Value: integer);
public
     constructor Create(AParent: TXSSWindow; AHorizontal: boolean);
     destructor Destroy; override;
     procedure Clear; reintroduce;
     procedure Paint; override;
     procedure MouseDown(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer); override;
     procedure MouseMove(Shift: TXSSShiftState; X, Y: Integer); override;
     procedure MouseUp(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer); override;
     procedure MouseLeave; override;
     procedure SetSize(const pX1, pY1, pX2, pY2: integer); override;

     property Levels: integer read GetLevels write SetLevels;
     property Buttons: TButtonRectList read FButtons;
     end;

implementation

{ TButtonRectList }

procedure TButtonRectList.Add(pX,pY,Index: integer; IdA: integer = -1; IdB: integer = -1);
var
  W,H: integer;
  BR: TButtonRect;
begin
  FSkin.GetBmpBtnSize(FButtonType,H,W);
  BR := TButtonRect.Create;
  BR.FX1 := pX;
  BR.FY1 := pY;
  BR.FX2 := pX + W;
  BR.FY2 := pY + H;
  BR.IdA := IdA;
  BR.IdB := IdB;
  BR.Index := Index;
  inherited Add(BR);
end;

constructor TButtonRectList.Create(Skin: TXLSBookSkin; ButtonType: TBmpButtonType);
begin
  inherited Create;
  FSkin := Skin;
  FButtonType := ButtonType;
  FCurrButton := -1;
end;

function TButtonRectList.GetItems(Index: integer): TButtonRect;
begin
  Result := TButtonRect(inherited Items[Index]);
end;

procedure TButtonRectList.MouseDown(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer);
var
  i: integer;
begin
  i := MouseHit(X,Y);
  if i >= 0 then begin
    FCurrButton := i;
    Items[i].ButtonState := bbsClicked;
    FSkin.PaintBmpButton(FButtonType,Items[i].ButtonState,Items[i].Index,Items[i].X,Items[i].Y);
  end;
end;

function TButtonRectList.MouseHit(X, Y: integer): integer;
begin
  for Result := 0 to Count - 1 do begin
    if Items[Result].MouseHit(X,Y)then
      Exit;
  end;
  Result := -1;
end;

procedure TButtonRectList.MouseLeave;
begin
  if FCurrButton >= 0 then begin
    Items[FCurrButton].ButtonState := bbsNormal;
    FSkin.PaintBmpButton(FButtonType,Items[FCurrButton].ButtonState,Items[FCurrButton].Index,Items[FCurrButton].X,Items[FCurrButton].Y);
    FCurrButton := -1;
  end;
end;

procedure TButtonRectList.MouseMove(Shift: TXSSShiftState; X, Y: Integer);
var
  i: integer;
begin
  i := MouseHit(X,Y);
  if i >= 0 then begin
    if FCurrButton <> i then begin
      if FCurrButton >= 0 then
        MouseLeave;
      FCurrButton := i;
      Items[i].ButtonState := bbsFocused;
      FSkin.PaintBmpButton(FButtonType,Items[i].ButtonState,Items[i].Index,Items[i].X,Items[i].Y);
    end;
    Exit;
  end;
  if FCurrButton >= 0 then
    MouseLeave;
end;

procedure TButtonRectList.MouseUp(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer);
var
  i: integer;
begin
  i := MouseHit(X,Y);
  if i >= 0 then begin
    FCurrButton := i;
    if Items[i].ButtonState = bbsClicked then begin
      Items[i].ButtonState := bbsFocused;
      FSkin.PaintBmpButton(FButtonType,Items[i].ButtonState,Items[i].Index,Items[i].X,Items[i].Y);
      FCurrButton := -1;
      FButtonClickEvent(Self,Items[i].FIdA,Items[i].FIdB);
    end;
  end;
end;

procedure TButtonRectList.Paint;
var
  i: integer;
begin
  for i := 0 to Count - 1 do
    FSkin.PaintBmpButton(FButtonType,Items[i].ButtonState,Items[i].Index,Items[i].X,Items[i].Y);
end;

{ TButtonRect }

function TButtonRect.MouseHit(X, Y: integer): boolean;
begin
  Result := (X >= FX1) and (X <= FX2) and (Y >= FY1) and (Y <= FY2)
end;

{ TGroupButtons }

procedure TGroupButtons.Clear;
begin
  SetLevels(0);
end;

constructor TGroupButtons.Create(AParent: TXSSWindow; AHorizontal: boolean);
begin
  inherited Create(AParent);
  FHorizontal := AHorizontal;
  FButtons := TButtonRectList.Create(FSkin,bbtNumbers);
end;

destructor TGroupButtons.Destroy;
begin
  FButtons.Free;
  inherited;
end;

function TGroupButtons.GetLevels: integer;
begin
  Result := FButtons.Count;
end;

procedure TGroupButtons.MouseDown(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer);
begin
  inherited MouseDown(Button,Shift,X,Y);
  FButtons.MouseDown(Button, Shift, X, Y);
end;

procedure TGroupButtons.MouseLeave;
begin
  inherited;
  FButtons.MouseLeave;
end;

procedure TGroupButtons.MouseMove(Shift: TXSSShiftState; X, Y: Integer);
begin
  inherited MouseMove(Shift,X,Y);;
  FSkin.GDI.SetCursor(xctArrow);
  FButtons.MouseMove(Shift, X, Y);
end;

procedure TGroupButtons.MouseUp(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer);
begin
  inherited MouseUp(Button,Shift,X,Y);;
  FButtons.MouseUp(Button, Shift, X, Y);
end;

procedure TGroupButtons.Paint;
begin
  inherited;

  FSkin.PaintHeader(FCX1,FCY1,FCX2,FCY2,hssNormal,xshsGutter);

  BeginPaint;

  FButtons.Paint;

  EndPaint;
end;

procedure TGroupButtons.SetLevels(const Value: integer);
var
  p,i,d: integer;
begin
  FButtons.Clear;
  if FHorizontal then
    d := ((FCY2 - FCY1) div 2) - (GUTTER_SIZE div 2)
  else
    d := ((FCX2 - FCX1) div 2) - (GUTTER_SIZE div 2);
  p := 0;
  for i := FButtons.Count to Value - 1 do begin
    if FHorizontal then
      FButtons.Add(FCX1 + p,FCY1 + d,i,i + 1,i + 1)
    else
      FButtons.Add(FCX1 + d,FCY1 + p,i,i + 1,i + 1);
    Inc(p,GUTTER_SIZE);
  end;
end;

procedure TGroupButtons.SetSize(const pX1, pY1, pX2, pY2: integer);
begin
  inherited;
  SetLevels(FButtons.Count);
end;

end.
