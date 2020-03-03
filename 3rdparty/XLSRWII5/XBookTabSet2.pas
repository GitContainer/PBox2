unit XBookTabSet2;

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

uses { Delphi  } Classes, SysUtils, Contnrs, Math,
     { XLSRWII } Xc12Utils5, XLSUtils5,
     { XLSBook } XBookWindows2, XBookRects2, XBookPaintGDI2, XBookPaint2, XBookSkin2,
                 XBookUtils2, XBookTypes2;


const TABBUTTONWIDTH  = 15;
const TABBUTTONHEIGHT = 14;
const MAXTABTEXTLEN   = 31;

type TTabButtons = class(TXSSClientWindow)
private
     FOverButton: integer;
     FBtnRects: array[0..3] of TXYRect;
     FClickEvent: TXIntegerEvent;
protected
public
     constructor Create(AParent: TXSSWindow);
     procedure SetSize(const pX1,pY1,pX2,pY2: integer); override;
     procedure Paint; override;
     procedure MouseDown(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer); override;
     procedure MouseMove(Shift: TXSSShiftState; X, Y: Integer); override;
     procedure MouseUp(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer); override;
     procedure MouseLeave; override;

     property OnClick: TXIntegerEvent read FClickEvent write FClickEvent;
     end;

type TXTab = class(TObject)
private
     FText: AxUCString;
     FX1,FX2: integer;
     FColor: longword;
     FFocused: boolean;

     procedure SetText(const Value: AxUCString);
     procedure SetColor(const Value: longword);
public
     property Text: AxUCString read FText write SetText;
     property X1: integer read FX1 write FX1;
     property X2: integer read FX2 write FX2;
     property Color: longword read FColor write SetColor;
     property Focused: boolean read FFocused write FFocused;
     end;

type TXTabs = class(TObjectList)
private
     FSkin: TXLSBookSkin;
     FLeftTab: integer;
     FWidth: integer;
     FSelectedTab: integer;

     function  GetItems(Index: integer): TXTab;
     procedure SetClickTab(const Value: integer; State: TTabState; Range: boolean);
     procedure SetSelectedTab(const Value: integer);
     function  GetLeftStart: integer;
public
     constructor Create(Skin: TXLSBookSkin);
     procedure CalcSize;

     function Add: TXTab;
     function MouseDown(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer): integer;
     property Items[Index: integer]: TXTab read GetItems; default;

     property LeftTab: integer read FLeftTab write FLeftTab;
     property LeftStart: integer read GetLeftStart;
     property SelectedTab: integer read FSelectedTab write SetSelectedTab;
     property Width: integer read FWidth write FWidth;
     end;

type TXTabSet = class(TXSSClientWindow)
private
     FTabs: TXTabs;
     FSplitter: TSplitterHandle;
     FPercentWidth: double;
     FXSizeEvent: TX2IntegerEvent;
     FSplitterSizing: boolean;
     FTabButtons: TTabButtons;
     FLastClick: integer;
//     FTabEdit: TEdit;
     FClickEvent: TXIntegerEvent;

     procedure SetPercentWidth(const Value: double);
     function  GetSelectedTab: integer;
     procedure SetSelectedTab(const Value: integer);
     function  GetTexts(Index: integer): AxUCString;
     procedure SetTexts(Index: integer; const Value: AxUCString);
protected
     procedure ButtonClick(Sender: TObject; Value: integer);
     procedure ClientPaint;
     procedure CalcEditSize;
     procedure ShowEdit;
     procedure HideEdit;
     procedure TabEditKeyPress(Sender: TObject; var Key: Char);
public
     constructor Create(AParent: TXSSWindow);
     destructor Destroy; override;
     procedure SetSize(const pX1,pY1,pX2,pY2: integer); override;
     procedure Paint; override;
     procedure MouseDown(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer); override;
     procedure MouseMove(Shift: TXSSShiftState; X, Y: Integer); override;
     procedure MouseUp(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer); override;
     procedure MouseEnter(X, Y: Integer); override;
     procedure MouseLeave; override;
     function  SnapX(var X: integer): integer;
     procedure Exchange(Index1,Index2: integer);
     procedure DeleteChilds; override;
     procedure Add(TabText: AxUCString; TabColor: longword = XLSCOLOR_AUTO);

     property PercentWidth: double read FPercentWidth write SetPercentWidth;
     property OnSize: TX2IntegerEvent read FXSizeEvent write FXSizeEvent;
     property LastClick: integer read FLastClick write FLastClick;
     property SelectedTab: integer read GetSelectedTab write SetSelectedTab;
     property Texts[Index: integer]: AxUCString read GetTexts write SetTexts;

     property OnClick: TXIntegerEvent read FClickEvent write FClickEvent;
     end;

implementation

{ TXTabSet }

procedure TXTabSet.Add(TabText: AxUCString; TabColor: longword = XLSCOLOR_AUTO);
begin
  with FTabs.Add do begin
    Text := TabText;
    Color := TabColor;
  end;

  FTabs.CalcSize;
end;

procedure TXTabSet.ButtonClick(Sender: TObject; Value: integer);
var
  NewValue: integer;
begin
  NewValue := FTabs.LeftTab;
  case Value of
    0: NewValue := 0;
    1: if FTabs.LeftTab > 0 then
         NewValue := FTabs.LeftTab - 1;
    2: if (FTabs.LeftTab < (FTabs.Count - 1)) and ((FTabButtons.X2 + FTabs[FTabs.Count - 1].FX2) > FCX2) then
         NewValue := FTabs.LeftTab + 1;
    3: begin
      while (FTabs.LeftTab < (FTabs.Count - 1)) and ((FTabButtons.X2 + FTabs[FTabs.Count - 1].FX2) > FCX2) do begin
        FTabs.LeftTab := FTabs.LeftTab + 1;
        FTabs.CalcSize;
        NewValue := -1;
      end;
      if NewValue < 0 then
        Paint;
      Exit;
    end;
  end;
  if NewValue <> FTabs.LeftTab then begin
    FTabs.LeftTab := NewValue;
    FTabs.CalcSize;
    Paint;
  end;
end;

procedure TXTabSet.CalcEditSize;
//var
//  w: integer;
begin
//  w := TABSLANT * 3;
//  FTabEdit.Left := FCX1 + FEditTab.FX1 + w div 2;
//  FTabEdit.Top := FCY1;
//  FTabEdit.Width := FEditTab.FX2 - FEditTab.FX1 - w;
//
//  w := FTabEdit.Width;
//  if w <= 0 then
//    w := 1
//  else if w >= (FCX2 - FCX1 - TABSLANT) then
//    w := FCX2 - FCX1 - TABSLANT;
//  FTabEdit.Width := w;
//
//  if FSkin.Style >= xssExcel2007 then
//    FTabEdit.Width := FTabEdit.Width + TABSLANT * 2;
//
//  FTabEdit.Height := FSkin.GDI.TM_SYS.tmHeight - 2;
end;

procedure TXTabSet.DeleteChilds;
begin
  if FTabs <> Nil then
    FTabs.Clear;
end;

procedure TXTabSet.ClientPaint;
var
  i,j: integer;
begin
  FSkin.GDI.PaintColor := FSkin.Colors.WindowBkg;
  FSkin.GDI.Rectangle(FX1,FY1,FX2,FY2);

  FSkin.GDI.PenColor := FSkin.Colors.HeaderLine;
  FSkin.GDI.Line(FCX1,FCY1,FCX2,FCY1);

  j := -1;
  for i := FTabs.Count - 1 downto FTabs.LeftTab do begin
    if i = FTabs.SelectedTab then
      j := i
    else begin
      if FTabs[i].Focused then
        FSkin.PaintTab(FCX1 + FTabs[i].X1,FCX1 + FTabs[i].X2,FCY1,FTabs[i].Text,tsFocused,FTabs[i].Color)
      else
        FSkin.PaintTab(FCX1 + FTabs[i].X1,FCX1 + FTabs[i].X2,FCY1,FTabs[i].Text,tsNormal,FTabs[i].Color);
    end;
  end;
  if FTabs.LeftTab > 0 then
    FSkin.PaintTab(FCX1 - TABSLANT,FCX1 + TABSLANT,FCY1,'',tsNormal,FSkin.Colors.Header.Cl1);
  if j >= 0 then
    FSkin.PaintTab(FCX1 + FTabs[j].X1,FCX1 + FTabs[j].X2,FCY1,FTabs[j].Text,tsSelected,FTabs[j].Color);
end;

constructor TXTabSet.Create(AParent: TXSSWindow);
begin
  inherited Create(AParent);
  FLastClick := -1;
//  if False and not (csDesigning in FParentControl.ComponentState) then begin
//    FTabEdit := TEdit.Create(FParentControl);
//    FTabEdit.Visible := False;
//    FTabEdit.Parent := FParentControl;
//    FTabEdit.BorderStyle := bsNone;
//    FTabEdit.AutoSelect := True;
//    FTabEdit.HideSelection := False;
//    FTabEdit.OnKeyPress := TabEditKeyPress;
//    FTabEdit.MaxLength := MAXTABTEXTLEN;
//  end;
  FTabs := TXTabs.Create(FSkin);
  FTabs.Add.Text := 'Sheet1';
  FTabs.SelectedTab := 0;
  FTabs.CalcSize;
  FPercentWidth := 0.6;
  FTabButtons := TTabButtons.Create(Self);
  FTabButtons.OnClick := ButtonClick;
  FSplitter := TSplitterHandle.Create(Self,shtHorizontal);
end;

destructor TXTabSet.Destroy;
begin
  FreeAndNil(FTabButtons);
  FreeAndNil(FTabs);
  FreeAndNil(FSplitter);
//  if not ((FParentControl <> Nil) and (csDesigning in FParentControl.ComponentState)) and (FTabEdit <> Nil) then begin
////    FTabEdit.Parent := Nil;
//    FTabEdit.Free;
//  end;
  inherited;
end;

procedure TXTabSet.Exchange(Index1, Index2: integer);
begin
  FTabs.Exchange(Index1, Index2);
  FTabs.SelectedTab := Index1;
  FTabs.CalcSize;
  ClientPaint;
  EndPaint;
end;

function TXTabSet.GetSelectedTab: integer;
begin
  Result := FTabs.SelectedTab;
end;

function TXTabSet.GetTexts(Index: integer): AxUCString;
begin
  Result := FTabs[Index].Text;
end;

procedure TXTabSet.HideEdit;
begin
//  if (FTabEdit <> Nil) and FTabEdit.Visible then
//    FTabEdit.Visible := False;
//  FEditTab := Nil;
end;

procedure TXTabSet.MouseDown(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer);
begin
  inherited MouseDown(Button,Shift,X,Y);
  HideEdit;
  if xssLeft in Shift then begin
    if FTabButtons.Hit(X,Y) then
      FTabButtons.MouseDown(Button,Shift,X,Y)
    else if FSplitter.Hit(X,Y) then
      FSplitterSizing := True
    else if Hit(X,Y) then begin
      FLastClick := FTabs.MouseDown(Button,Shift,X - FCX1,Y);
      BeginPaint;
      ClientPaint;
      EndPaint;
      if xssDouble in Shift then
        ShowEdit
      else if FLastClick >= 0 then
        FClickEvent(Self,FLastClick);
    end;
  end;
end;

procedure TXTabSet.MouseEnter(X, Y: Integer);
begin
  MouseMove([],X,Y);
end;

procedure TXTabSet.MouseLeave;
begin
//  FSplitterSizing := False;
  FTabButtons.MouseLeave;
  FLastClick := -1;
//  HideEdit;
end;

procedure TXTabSet.MouseMove(Shift: TXSSShiftState; X, Y: Integer);
begin
  inherited MouseMove(Shift,X,Y);;
  if FSplitter.Hit(X,Y) then
    FSplitter.MouseMove(Shift,X,Y)
  else if FSplitterSizing then begin
    FSplitter.MouseMove(Shift,X,Y);
    FXSizeEvent(Self,X,Y)
  end
  else
    FTabButtons.MouseMove(Shift,X,Y);
end;

procedure TXTabSet.MouseUp(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer);
begin
  inherited MouseUp(Button,Shift,X,Y);;
  FSplitterSizing := False;
  FLastClick := -1;
  FTabButtons.MouseUp(Button,Shift,X,Y);
end;

procedure TXTabSet.Paint;
begin
  BeginPaint;

  ClientPaint;

  EndPaint;

  FSplitter.Paint;
  FTabButtons.Paint;
end;

procedure TXTabSet.SetPercentWidth(const Value: double);
begin
  if Value >1.0 then
    FPercentWidth := 1.0
  else if Value < 0.0 then
    FPercentWidth := 0.0
  else
    FPercentWidth := Value;
end;

procedure TXTabSet.SetSelectedTab(const Value: integer);
begin
  FTabs.SelectedTab := Value;
end;

procedure TXTabSet.SetSize(const pX1, pY1, pX2, pY2: integer);
var
  BtnX2: integer;
begin
  inherited SetSize(pX1, pY1, pX2, pY2);
  BtnX2 := Min(FCX1 + TABBUTTONWIDTH * 4,pX2 - SPLITTER_SIZE);
  SetClientSize(BtnX2 + 1, pY1, pX2 - SPLITTER_SIZE - 1, pY2);

  FTabButtons.SetSize(FX1,FY1,BtnX2,FY2);

  FSplitter.SetSize(pX2 - SPLITTER_SIZE,FY1,pX2 - 1,FY2);

  FTabs.Width := FCX2 - FCX1;

//  FTabs.SelectedTab := 0; //FTabs.IndexOf(FEditTab);
end;

procedure TXTabSet.SetTexts(Index: integer; const Value: AxUCString);
begin
  FTabs[Index].Text := Value;
  Paint;
end;

procedure TXTabSet.ShowEdit;
begin
{
  FEditTab := FTabs[FLastClick];
  FTabEdit.Font.Assign(FSkin.SystemFont);
  FTabEdit.Font.Style := [fsBold];
  CalcEditSize;
  FTabEdit.Text := FEditTab.FText;
  FTabEdit.SelectAll;
  FTabEdit.Visible := True;
  FTabEdit.SetFocus;
}  
end;

function TXTabSet.SnapX(var X: integer): integer;
var
  i: integer;
begin
  Dec(X,FCX1);
  if X <= 0 then begin
    X := FCX1;
    Result := 0;
    Exit;
  end
  else begin
    for i := FTabs.FLeftTab to FTabs.Count - 1 do begin
      if (FTabs[i].FX2 >= FTabs.Width) or ((X >= FTabs[i].FX1) and (X <= FTabs[i].FX2)) then begin
        X := FTabs[i].FX1 + FCX1 + TABSLANT div 2;
        Result := i;
        Exit;
      end;
    end;
  end;
  X := FTabs[FTabs.Count - 1].FX2 + FCX1;
  Result := FTabs.Count - 1;
end;

procedure TXTabSet.TabEditKeyPress(Sender: TObject; var Key: Char);
begin
//  FEditTab.FText := FTabEdit.Text + Key;
//  FTabs.SelectedTab := FTabs.IndexOf(FEditTab);
//  FTabs.CalcSize;
//  CalcEditSize;
//  inherited Paint;
//  ClientPaint;
//  EndPaint;
//  FTabEdit.Invalidate;
end;

{ TTabButtons }

constructor TTabButtons.Create(AParent: TXSSWindow);
begin
  inherited Create(AParent);
  FOverButton := -1;
end;

procedure TTabButtons.MouseDown(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer);
begin
  inherited MouseDown(Button,Shift,X,Y);
  if FOverButton >= 0 then begin
    BeginPaint;
    FSkin.PaintBmpButton(bbtTabset,bbsClicked,FOverButton,FBtnRects[FOverButton].X1,FBtnRects[FOverButton].Y1);
    EndPaint;
  end;
end;

procedure TTabButtons.MouseLeave;
begin
  if FOverButton >= 0 then begin
    BeginPaint;
    FSkin.PaintBmpButton(bbtTabset,bbsNormal,FOverButton,FBtnRects[FOverButton].X1,FBtnRects[FOverButton].Y1);
    FOverButton := -1;
    EndPaint;
  end;
end;

procedure TTabButtons.MouseMove(Shift: TXSSShiftState; X, Y: Integer);
var
  i: integer;
begin
  inherited MouseMove(Shift,X,Y);;
  for i := 0 to High(FBtnRects) do begin
    if PtInXYRect(X,Y,FBtnRects[i]) then begin
      if i <> FOverButton then begin
        BeginPaint;
        if FOverButton >= 0 then
          FSkin.PaintBmpButton(bbtTabset,bbsNormal,FOverButton,FBtnRects[FOverButton].X1,FBtnRects[FOverButton].Y1);
        FOverButton := i;
        FSkin.PaintBmpButton(bbtTabset,bbsFocused,FOverButton,FBtnRects[FOverButton].X1,FBtnRects[FOverButton].Y1);
        EndPaint;
      end;
      Exit;
    end;
  end;
  if FOverButton >= 0 then begin
    FSkin.PaintBmpButton(bbtTabset,bbsNormal,FOverButton,FBtnRects[FOverButton].X1,FBtnRects[FOverButton].Y1);
    FOverButton := -1;
  end;
end;

procedure TTabButtons.MouseUp(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer);
begin
  inherited MouseUp(Button,Shift,X,Y);;
  if FOverButton >= 0 then begin
    BeginPaint;
    FSkin.PaintBmpButton(bbtTabset,bbsNormal,FOverButton,FBtnRects[FOverButton].X1,FBtnRects[FOverButton].Y1);
    EndPaint;
    FClickEvent(Self,FOverButton);
    FOverButton := -1;
  end;
end;

procedure TTabButtons.Paint;
var
  i: integer;
begin
  BeginPaint;

  FSkin.GDI.PaintColor := FSkin.Colors.Header.Cl1;
  FSkin.GDI.Rectangle(FCX1,FCY1,FCX2,FCY2);
  FSkin.GDI.PenColor := FSkin.Colors.HeaderLine;
  FSkin.GDI.Line(FCX1,FCY1,FCX2,FCY1);

  for i := 0 to High(FBtnRects) do begin
    if i = FOverButton then
      FSkin.PaintBmpButton(bbtTabset,bbsFocused,i,FBtnRects[i].X1,FBtnRects[i].Y1 + 1)
    else
      FSkin.PaintBmpButton(bbtTabset,bbsNormal,i,FBtnRects[i].X1,FBtnRects[i].Y1 + 1);
  end;

  EndPaint;
end;

procedure TTabButtons.SetSize(const pX1, pY1, pX2, pY2: integer);
var
  i: integer;
begin
  inherited;
  FBtnRects[0].X1 := 0;
  FBtnRects[0].Y1 := 0;
  FBtnRects[0].X2 := TABBUTTONWIDTH;
  FBtnRects[0].Y2 := TABBUTTONHEIGHT;
  FBtnRects[1] := FBtnRects[0];
  FBtnRects[2] := FBtnRects[0];
  FBtnRects[3] := FBtnRects[0];
  for i := 0 to High(FBtnRects) do
    OffsetXYRect(FBtnRects[i],FCX1 + 1 + i * TABBUTTONWIDTH,FCY1 + 1);
end;

{ TXTabs }

function TXTabs.Add: TXTab;
begin
  Result := TXTab.Create;
  inherited Add(Result);
end;

procedure TXTabs.CalcSize;
var
  i,x: integer;
begin
  FSkin.SetTabFont(True);
  x := 0;
  for i := FLeftTab to Count - 1 do begin
    Items[i].FX1 := x;
    Items[i].FX2 := x + FSkin.GDI.TextWidth(Items[i].FText) + TABSLANT * 3;
    if FSkin.Style >= xssExcel2007 then
      Items[i].FX2 := Items[i].FX2 + TABSLANT;
    x := Items[i].FX2 - TABSLANT div 2 - 2;
  end;
  if FLeftTab > 0 then begin
    i := FLeftTab - 1;
    Items[i].FX1 := -(FSkin.GDI.TextWidth(Items[i].FText) + TABSLANT * 2);
    Items[i].FX2 := TABSLANT;
  end;
end;

constructor TXTabs.Create(Skin: TXLSBookSkin);
begin
  inherited Create;
  FSkin := Skin;
end;

function TXTabs.GetItems(Index: integer): TXTab;
begin
  Result := TXTab(inherited Items[Index]);
end;

function TXTabs.GetLeftStart: integer;
begin
  if FLeftTab > 0 then
    Result := FLeftTab - 1
  else
    Result := 0;
end;

function TXTabs.MouseDown(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer): integer;
var
  i: integer;
begin
  for i := LeftStart to Count - 1 do begin
    if (X >= Items[i].FX1) and (X <= Items[i].FX2) then begin
      Result := i;
      if xssCtrl in Shift then
        SetClickTab(i,tsFocused,False)
      else if xssShift in Shift then
        SetClickTab(i,tsFocused,True)
      else
        SetClickTab(i,tsSelected,False);
      Exit;
    end;
  end;
  Result := -1;
end;

procedure TXTabs.SetClickTab(const Value: integer; State: TTabState; Range: boolean);
var
  i: integer;

procedure ClearFocus;
var
  i: integer;
begin
  for i := 0 to Count - 1 do
    Items[i].FFocused := False;
end;

function AllFocused: boolean;
var
  i: integer;
begin
  Result := False;
  for i := 0 to Count - 1 do begin
    if not Items[i].FFocused then
      Exit;
  end;
  Result := True;
end;

begin
  if State = tsSelected then begin
    if (not Items[Value].FFocused) or AllFocused then
      ClearFocus;
    FSelectedTab := Value;
    Items[Value].FFocused := True;
  end
  else if Range then begin
    ClearFocus;
    for i := Min(Value,FSelectedTab) to Max(Value,FSelectedTab) do
      Items[i].FFocused := True;
  end
  else
    Items[Value].FFocused := not Items[Value].FFocused;
  if (Items[Value].FX1 < 0) and (FLeftTab > 0) then begin
    Dec(FLeftTab);
    CalcSize;
  end
  else begin
    while (Items[Value].FX2 > FWidth) and (Items[Value].FX1 > 0) and (FLeftTab < Count) do begin
      Inc(FLeftTab);
      CalcSize;
    end;
  end;
end;

procedure TXTabs.SetSelectedTab(const Value: integer);
begin
  SetClickTab(Value,tsSelected,False);
end;

{ TXTab }

procedure TXTab.SetColor(const Value: longword);
begin
  FColor := Value;
end;

procedure TXTab.SetText(const Value: AxUCString);
begin
  // Max length shall be taken from XLSReadWriteII
  FText := Copy(Value,1,MAXTABTEXTLEN);
end;

end.
