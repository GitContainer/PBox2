unit XBookHeaders2;

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
     XBookSkin2, XBookPaintGDI2, XBookPaint2, XBookUtils2, XBookSysVar2, XBookWindows2,
     XBookButtons2;

type THeaderOption = (hoHidden,hoCollapsed);
     THeaderOptions = set of THeaderOption;

type THeaderData = record
     Header: integer;
     Size,SizingSize: word;
     SizeFloat: double;
     Pos: integer;
     SelState: THeaderSelState;
     Hoover: boolean;
     OutlineLevel: integer;
     Options: THeaderOptions;
     end;

type THeaderScroll = (hsInc,hsDec,hsIncPage,hsDecPage,hsFirst,hsLast);

type TSizeHitType = (shtNone,shtSize,shtHidden);

type TOutlineLineType = (olpFirst,olpMiddle,olpLast,olpDot);

type THeaderClickEvent = procedure(Sender: TObject; Header: integer; SizeHit: TSizeHitType; Button: TXSSMouseButton; Shift: TXSSShiftState) of object;

type TXHeadersGutter = class(TXSSClientWindow)
protected
     FLevels: integer;
     FFirstLevel: integer;
     FFirstVisibleLevel: integer;
     FButtons: TButtonRectList;
     FCurrButton: integer;

     procedure PaintOutline(P1,P2,Level: integer; LineType: TOutlineLineType); virtual; abstract;
     procedure CentrePos(P1,P2,Delta: integer; var cx,cy: integer); virtual; abstract;
public
     constructor Create(AParent: TXSSWindow);
     destructor Destroy; override;
     procedure Paint(Headers: array of THeaderData); reintroduce;
     procedure CalcButtons(Headers: array of THeaderData; FirstHdr: integer);
     procedure MouseDown(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer); override;
     procedure MouseMove(Shift: TXSSShiftState; X, Y: Integer); override;
     procedure MouseUp(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer); override;
     procedure MouseLeave; override;

     property Levels: integer read FLevels write FLevels;
     property FirstLevel: integer read FFirstLevel write FFirstLevel;
     property FirstVisibleLevel: integer read FFirstVisibleLevel write FFirstVisibleLevel;
     property Buttons: TButtonRectList read FButtons;
     end;

type TXLSBookHeaders = class(TXSSClientWindow)
private
     function  GetOptions(Index: integer): THeaderOptions;
     function  GetVisibleSize(Index: integer): integer;
     function  GetOutlineLevel: integer;
     function  GetPos(Index: integer): integer;
     function  GetSelected(Index: integer): THeaderSelState;
     procedure SetSelected(Index: integer; const Value: THeaderSelState);
     function  GetHalfPos(Index: integer): integer;
     function  GetAbsPos(Index: integer): integer;
     function  GetAbsVisibleSize(Index: integer): integer;
     function  GetVisibleSizeFloat(Index: integer): double;
     function  GetAbsVisibleSizeEmu(Index: integer): integer;
protected
     FMaxHdr: integer;
     FFirstHdr,FLastHdr: integer;
     FFrozenHdr: integer;
     FFrozenFirstHdr: integer;
     FSizingHdr: integer;
     FCurrHeader: integer;
     FHooverHeader: integer;
     FMouseIsDown: boolean;
     FHeaders: array of THeaderData;
     FChangedEvent: TNotifyEvent;
     FClickP,FHeaderP: integer;
     FPaintLineEvent: TXPaintLineEvent;
     FGutter: TXHeadersGutter;
     FSyncButtonClickEvent: TX2IntegerEvent;
     FHeaderClickEvent: THeaderClickEvent;
     FSelHeaderChangedEvent: TX2IntegerEvent;
     FOffset: integer;

     procedure AddSizingHeaders;
     function  HeaderHit(P: integer; out HeaderPos: integer): integer;
     function  HdrSizeHit(P: integer; var Header,HeaderP: integer): TSizeHitType;
     function  FirstPos: integer; virtual; abstract;
     function  LastPos: integer; virtual; abstract;
     procedure GetHeaderData(Hdr: integer; var Size,Outline: integer; var Options: THeaderOptions); virtual; abstract;
     procedure GetHeaderDataFloat(Hdr: integer; var Size: double; var Outline: integer; var Options: THeaderOptions); virtual; abstract;
     function  XorY(X,Y: integer): integer; virtual; abstract;
     function  GetCursor(HitType: TSizeHitType): TXLSCursorType; virtual; abstract;
     function  DeltaPos: integer;
     function  FirstScrollHeader: integer; {$ifdef D2006PLUS} inline; {$endif}
     procedure PaintSizing; virtual; abstract;
     procedure PaintHeader(Hdr: integer); virtual; abstract;
     procedure CalcGutter;
     procedure HooverOff; {$ifdef D2006PLUS} inline; {$endif}
public
     constructor Create(Parent: TXSSClientWindow; MaxHdr: integer);
     procedure ClearSelected;
     procedure SetSelectedState(P1,P2: integer; IsSelected: boolean);
     procedure SetFocusedState(P1,P2: integer);
     function  Count: integer; {$ifdef D2006PLUS} inline; {$endif}
     procedure MouseDown(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer); override;
     procedure MouseMove(Shift: TXSSShiftState; X, Y: Integer); override;
     procedure MouseUp(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer); override;
     procedure MouseEnter(X, Y: Integer); override;
     procedure MouseLeave; override;
     procedure SetSize(const pX1,pY1,pX2,pY2: integer); override;
     procedure CalcHeaders(MaxHdrToCalc: integer = MAXINT); virtual;
     procedure CalcHeadersFromLast; virtual;
     function  Snap(P,EndP: integer; var Hdr: integer): integer;
     function  NextVisible(Hdr: integer): integer;
     function  PrevVisible(Hdr: integer): integer;
     function  PosToHeader(P: integer): integer;
     function  GetHeight(Hdr1,Hdr2: integer): integer;
     function  HdrFromSizeRev(Hdr,Size: integer): integer;
     function  IsFrozen: boolean; // inline;
     procedure Scroll(ScrollDirection: THeaderScroll);
     function  VisualSize: integer;
     function  ScrollSize(ScrollInc: boolean): integer;

     property Option[Index: integer]: THeaderOptions read GetOptions;
     property Selected[Index: integer]: THeaderSelState read GetSelected write SetSelected;
     property VisibleSize[Index: integer]: integer read GetVisibleSize;
     property VisibleSizeFloat[Index: integer]: double read GetVisibleSizeFloat;
     property AbsVisibleSize[Index: integer]: integer read GetAbsVisibleSize;
     property AbsVisibleSizeEmu[Index: integer]: integer read GetAbsVisibleSizeEmu;
     property Pos[Index: integer]: integer read GetPos;
     property AbsPos[Index: integer]: integer read GetAbsPos;
     property HalfPos[Index: integer]: integer read GetHalfPos;
     property OutlineLevel: integer read GetOutlineLevel;
     property FirstHeader: integer read FFirstHdr;
     property Offset: integer read FOffset write FOffset;

     property OnChanged: TNotifyEvent read FChangedEvent write FChangedEvent;
     property OnPaintLine: TXPaintLineEvent read FPaintLineEvent write FPaintLineEvent;
     property OnSyncButtonClick: TX2IntegerEvent read FSyncButtonClickEvent write FSyncButtonClickEvent;
     property OnHeaderClick: THeaderClickEvent read FHeaderClickEvent write FHeaderClickEvent;
     property OnSelHeaderChanged: TX2IntegerEvent read FSelHeaderChangedEvent write FSelHeaderChangedEvent;
     end;

implementation

{ TXLSBookHeaders }

procedure TXLSBookHeaders.AddSizingHeaders;
var
  i,p,Count: integer;
  Sz,Outline: integer;
  Options: THeaderOptions;
begin
  Count := Length(FHeaders);
  FSkin.SetSystemFont(False);
  FSkin.GetCellFontData;
  FSkin.TextAlign(ptaCenter);
  p := FirstPos;
  for i := 0 to High(FHeaders) do
    Inc(p,FHeaders[i].SizingSize);
  if p > LastPos then
    Exit;
  SetLength(FHeaders,Length(FHeaders) + 16);
  while (p <= LastPos) and (FLastHdr <= FMaxHdr) do begin
    GetHeaderData(FLastHdr,Sz,Outline,Options);
    FHeaders[Count].Size := Round(Sz);
    FHeaders[Count].SizingSize := FHeaders[Count].Size;
    FHeaders[Count].OutlineLevel := Outline;
    FHeaders[Count].Options := Options;
    FHeaders[Count].SelState := hssNormal;
    Inc(p,FHeaders[Count].Size);
    Inc(FLastHdr);
    Inc(Count);
    if Count > Length(FHeaders) then
      SetLength(FHeaders,Length(FHeaders) + 16);
  end;
  SetLength(FHeaders,Count)
end;

procedure TXLSBookHeaders.CalcHeaders(MaxHdrToCalc: integer = MAXINT);
var
  Count: integer;
  p: double;

procedure AddHeaders(H1,H2: integer);
var
  Sz: double;
  Outline: integer;
  Options: THeaderOptions;
begin
  if MaxHdrToCalc < MAXINT then
    Inc(MaxHdrToCalc);
  while (H1 <= H2) and (p <= LastPos) and (FLastHdr <= FMaxHdr) and (LastPos > 0) and (FLastHdr <= MaxHdrToCalc) do begin
    GetHeaderDataFloat(H1,Sz,Outline,Options);
    FHeaders[Count].Header := H1;
    FHeaders[Count].Size := Round(Sz);
    FHeaders[Count].SizeFloat := Sz;
    FHeaders[Count].SizingSize := FHeaders[Count].Size;
    FHeaders[Count].Pos := Round(p);
    FHeaders[Count].OutlineLevel := Outline;
    FHeaders[Count].Options := Options;
    FHeaders[Count].SelState := hssNormal;
    if not (hoHidden in FHeaders[Count].Options) then
      p := p + Sz;
    Inc(FLastHdr);

    Inc(Count);
    if Count >= Length(FHeaders) then
      SetLength(FHeaders,Length(FHeaders) + 64);
    Inc(H1);
  end;
end;

begin
  Count := 0;
  SetLength(FHeaders,64);
  FLastHdr := FirstScrollHeader;
  FSkin.SetSystemFont(False);
  FSkin.GetCellFontData;
  FSkin.TextAlign(ptaCenter);
  p := FirstPos;

  if IsFrozen then begin
    AddHeaders(FFirstHdr,FFrozenHdr);
    AddHeaders(FFrozenFirstHdr,MAXINT);
  end
  else
    AddHeaders(FFirstHdr,MAXINT);
  Dec(FLastHdr);
  SetLength(FHeaders,Count);
  FGutter.CalcButtons(FHeaders,FFirstHdr);
  CalcGutter;
end;

procedure TXLSBookHeaders.CalcHeadersFromLast;
var
  Sz: double;
  W: double;
  H: integer;
  Outline: integer;
  Options: THeaderOptions;
begin
  W := LastPos - FirstPos;
  H := FLastHdr;
  while (W > 0) and (H >= 0) do begin
    GetHeaderDataFloat(H,Sz,Outline,Options);
    if not (hoHidden in Options) then
      W := W - Sz;
    if W > 0 then
      Dec(H);
  end;
  FFirstHdr := H + 1;

  CalcHeaders;
end;

function TXLSBookHeaders.HdrFromSizeRev(Hdr, Size: integer): integer;
var
  Sz,Outline: integer;
  Options: THeaderOptions;
begin
  Result := Hdr;
  GetHeaderData(Result,Sz,Outline,Options);
  while Result > 0 do begin
    GetHeaderData(Result,Sz,Outline,Options);
    if not (hoHidden in Options) then begin
      if (Size - Sz) < 0 then
        Break;
      Dec(Size,Sz);
    end;
    Dec(Result);
  end;
end;

function TXLSBookHeaders.HdrSizeHit(P: integer; var Header,HeaderP: integer): TSizeHitType;
var
  i,Hidden: integer;
begin
  HeaderP := FirstPos;
  for i := 0 to High(FHeaders) - 1 do begin
    Inc(HeaderP,VisibleSize[i]);
    if InDelta(P,HeaderP,HEADERSIZINGHORIZ_MARG) then begin
      Hidden := i + 1;
      while (Hidden <= High(FHeaders)) and ((FHeaders[Hidden].Size = 0) or (hoHidden in FHeaders[Hidden].Options)) do
        Inc(Hidden);
      if (Hidden > (i + 1)) and ((P - HeaderP) > HEADERSIZINGHORIZHIDDEN_MARG) then begin
        Header := Hidden - 1;
        Result := shtHidden;
      end
      else begin
        Header := i;
        Result := shtSize;
      end;
      Exit;
    end;
  end;
  Header := -1;
  Result := shtNone;
end;

function TXLSBookHeaders.Count: integer;
begin
  Result := Length(FHeaders);
end;

constructor TXLSBookHeaders.Create(Parent: TXSSClientWindow; MaxHdr: integer);
begin
  inherited Create(Parent);
  FParent := Parent;
  FMaxHdr := MaxHdr;
  FFirstHdr := 0;
  FSizingHdr := -1;
  FFrozenHdr := -1;
  FHooverHeader := -1;
end;

function TXLSBookHeaders.HeaderHit(P: integer; out HeaderPos: integer): integer;
begin
  HeaderPos := FirstPos;
  for Result := 0 to High(FHeaders) do begin
    if P <= (HeaderPos + VisibleSize[Result]) then
      Exit;
    Inc(HeaderPos,VisibleSize[Result]);
  end;
  Result := -1;
end;

procedure TXLSBookHeaders.HooverOff;
begin
  if FHooverHeader >= 0 then begin
    FHeaders[FHooverHeader].Hoover := False;
    PaintHeader(FHooverHeader);
  end;
  FHooverHeader := -1;
end;

function TXLSBookHeaders.IsFrozen: boolean;
begin
  Result := FFrozenHdr >= 0;
end;

procedure TXLSBookHeaders.MouseDown(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer);
var
  P: integer;
  Hit: TSizeHitType;
begin
  inherited MouseDown(Button,Shift,X,Y);
  if FGutter.Hit(X,Y) then
    FGutter.MouseDown(Button,Shift,X,Y)
  else if not (xssDouble in Shift) then begin
    FClickP := XorY(X,Y);
    Hit := HdrSizeHit(XorY(X,Y),FSizingHdr,FHeaderP);

    FSkin.GDI.SetCursor(GetCursor(Hit));

    FCurrHeader := HeaderHit(FClickP,P);
    FMouseIsDown := FCurrHeader >= 0;

    if FMouseIsDown then
      FHeaderClickEvent(Self,FCurrHeader,Hit,Button,Shift);
  end;
end;

procedure TXLSBookHeaders.MouseEnter(X, Y: Integer);
begin
  inherited;

end;

procedure TXLSBookHeaders.MouseLeave;
begin
  inherited;
  HooverOff;
end;

procedure TXLSBookHeaders.MouseMove(Shift: TXSSShiftState; X, Y: Integer);
var
  i,p,p1,p2,w,Hdr,HdrP: integer;
  SHT: TSizeHitType;
begin
  inherited MouseMove(Shift,X,Y);
  p := XorY(X,Y);
  Dec(p,DeltaPos);
  if p < FirstPos then
    p := FirstPos
  else if p > LastPos then
    p := LastPos;
  if FSizingHdr >= 0 then begin
    w := p - FHeaderP;
    Hdr := HeaderHit(FHeaderP + w,HdrP);
    if (Hdr >= 0) and (Hdr < FSizingHdr) then begin
      for i := 0 to Hdr - 1 do
        FHeaders[i].SizingSize := VisibleSize[i];
      FHeaders[Hdr].SizingSize := p - HdrP;
      for i := Hdr + 1 to FSizingHdr do
        FHeaders[i].SizingSize := 0;
    end
    else begin
      for i := 0 to FSizingHdr - 1 do
        FHeaders[i].SizingSize := VisibleSize[i];
      FHeaders[FSizingHdr].SizingSize := VisibleSize[FSizingHdr] + w;
    end;
    PaintSizing;
  end
  else begin
    if FMouseIsDown then begin
      p := XorY(X,Y);
      p1 := XorY(FCX1,FCY1);
      p2 := XorY(FCX2,FCY2);
      if p < p1 then
        FSelHeaderChangedEvent(Self,FCurrHeader,-1)
      else if p > p2 then
        FSelHeaderChangedEvent(Self,FCurrHeader,1)
      else begin
        Hdr := HeaderHit(p,HdrP);
        if Hdr <> FCurrHeader then begin
          FCurrHeader := Hdr;
          FSelHeaderChangedEvent(Self,FCurrHeader,0);
        end;
      end;
    end
    else begin
      SHT := HdrSizeHit(XorY(X,Y),Hdr,HdrP);
      FSkin.GDI.SetCursor(GetCursor(SHT));
      if SHT <> shtNone then
        HooverOff
      else begin
        Hdr := HeaderHit(XorY(X,Y),HdrP);
        if Hdr <> FHooverHeader then begin
          if FHooverHeader >= 0 then begin
            FHeaders[FHooverHeader].Hoover := False;
            PaintHeader(FHooverHeader);
          end;
          FHooverHeader := Hdr;
          FHeaders[FHooverHeader].Hoover := True;
          PaintHeader(FHooverHeader);
        end;
      end;
    end;
  end;
end;

procedure TXLSBookHeaders.MouseUp(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer);
begin
  inherited MouseUp(Button,Shift,X,Y);
  if not FMouseIsDown then
    FHeaderClickEvent(Self,FCurrHeader,shtNone,Button,[xssDouble]);
  FMouseIsDown := False;
end;

procedure TXLSBookHeaders.SetSize(const pX1, pY1, pX2, pY2: integer);
begin
  inherited;
//  if (pX1 > 0) and (pY1 > 0) and (pX2 > 0) and (pY2 > 0) then
//    CalcHeaders;
end;

function TXLSBookHeaders.DeltaPos: integer;
begin
  Result := FClickP - FHeaderP;
end;

function TXLSBookHeaders.FirstScrollHeader: integer;
begin
  if IsFrozen then
    Result := FFrozenFirstHdr
  else
    Result := FFirstHdr
end;

function TXLSBookHeaders.Snap(P,EndP: integer; var Hdr: integer): integer;
var
  Sz,Tmp: integer;
  Options: THeaderOptions;
begin
  Hdr := FFirstHdr;
  Result := XorY(FCX1,FCY1);
  while (P < EndP) and (Hdr <= FMaxHdr) do begin
    if P < Result then begin
      GetHeaderData(Hdr - 1,Sz,Tmp,Options);
      if hoHidden in Options then
        Sz := 0;
      if P < (Result - (Sz div 2)) then begin
        Dec(Hdr);
        Dec(Result,Sz);
      end;
      Exit;
    end;
    GetHeaderData(Hdr,Sz,Tmp,Options);
    if hoHidden in Options then
      Sz := 0;
    Inc(Result,Sz);
    Inc(Hdr);
  end;
  Result := EndP;
end;

function TXLSBookHeaders.VisualSize: integer;
begin
  Result := Length(FHeaders) - 1;
end;

function TXLSBookHeaders.GetOptions(Index: integer): THeaderOptions;
begin
{$ifdef XDEBUG}
  if (Index < 0) or (Index > High(FHeaders)) then
    raise XLSRWException.Create('Header index out of range Options');
{$endif}
  Result := FHeaders[Index].Options;
end;

function TXLSBookHeaders.NextVisible(Hdr: integer): integer;
var
  Sz     : integer;
  Outline: integer;
  Options: THeaderOptions;
  Visible: boolean;
begin
  Result := Hdr;
  repeat
    GetHeaderData(Result,Sz,Outline,Options);
    Visible := (Sz > 0) and not (hoHidden in Options);
    if not Visible then
      Inc(Result);
  until Visible or (Result > FMaxHdr);
end;

function TXLSBookHeaders.PrevVisible(Hdr: integer): integer;
var
  Sz,Outline: integer;
  Options: THeaderOptions;
  Visible: boolean;
begin
  Result := Hdr;
  repeat
    GetHeaderData(Result,Sz,Outline,Options);
    Visible := (Sz > 0) and not (hoHidden in Options);
    if not Visible then
      Dec(Result);
  until Visible or (Result < 0);
end;

function TXLSBookHeaders.GetVisibleSize(Index: integer): integer;
begin
{$ifdef XDEBUG}
  if (Index < 0) or (Index > High(FHeaders)) then begin
    raise XLSRWException.CreateFmt('Header index out of range VisibleSize, [%d]',[Index]);
  end;
{$endif}
  if hoHidden in FHeaders[Index].Options then
    Result := 0
  else
    Result := FHeaders[Index].Size;
end;

function TXLSBookHeaders.GetVisibleSizeFloat(Index: integer): double;
begin
{$ifdef XDEBUG}
  if (Index < 0) or (Index > High(FHeaders)) then begin
    raise XLSRWException.CreateFmt('Header index out of range VisibleSize, [%d]',[Index]);
  end;
{$endif}
  if hoHidden in FHeaders[Index].Options then
    Result := 0
  else
    Result := FHeaders[Index].SizeFloat;
end;

procedure TXLSBookHeaders.CalcGutter;
var
  Sz,Outline,H: integer;
  Options: THeaderOptions;
begin
  if FFirstHdr > 0 then begin
    H := FFirstHdr - 1;
    GetHeaderData(H,Sz,Outline,Options);
    FGutter.FirstLevel := Outline;
    FGutter.FirstVisibleLevel := Outline;
    while (H > 0) and (Outline > 0) and (hoHidden in Options) do begin
      Dec(H);
      GetHeaderData(H,Sz,Outline,Options);
      FGutter.FirstVisibleLevel := Outline;
    end;
  end
  else
    FGutter.FirstLevel := 0;
end;

function TXLSBookHeaders.GetOutlineLevel: integer;
begin
  Result := FGutter.FLevels;
end;

function TXLSBookHeaders.GetPos(Index: integer): integer;
begin
{$ifdef XDEBUG}
  if (Index < 0) or (Index > High(FHeaders)) then begin
    raise XLSRWException.CreateFmt('Header index out of range Pos (%d)',[Index]);
  end;
{$endif}
  Result := FHeaders[Index].Pos;
end;

function TXLSBookHeaders.GetSelected(Index: integer): THeaderSelState;
begin
{$ifdef XDEBUG}
  if (Index < 0) or (Index > High(FHeaders)) then
    raise XLSRWException.Create('Header index out of range GetSelected');
{$endif}
  Result := FHeaders[Index].SelState;
end;

procedure TXLSBookHeaders.Scroll(ScrollDirection: THeaderScroll);
var
  V: integer;
begin
  V := 0;
  case ScrollDirection of
    hsInc:     V := 1;
    hsDec:     V := -1;
    hsIncPage: V := ScrollSize(True);
    hsDecPage: V := -ScrollSize(False);
    hsFirst: ;
    hsLast: ;
    else       V := 0;
  end;
  if IsFrozen then
    Inc(FFrozenFirstHdr,V)
  else
    Inc(FFirstHdr,V);
  CalcHeaders;
end;

function TXLSBookHeaders.ScrollSize(ScrollInc: boolean): integer;
var
  W,Sz: double;
  H,i: integer;
  Outline: integer;
  Options: THeaderOptions;
begin
  if ScrollInc then begin
    Result := VisualSize - 1;
    if IsFrozen then
      Dec(Result,FFrozenHdr + 1);
  end
  else begin
    W := LastPos - FirstPos;
    if IsFrozen then begin
      H := FFrozenFirstHdr;
      for i := 0 to FFrozenHdr do
        W := W - FHeaders[i].SizeFloat;
    end
    else
      H := FFirstHdr;
    repeat
      GetHeaderDataFloat(H,Sz,Outline,Options);
      if (W - Sz) < 0 then
        Break;
      W := W - Sz;
      Dec(H);
    until (H <= 0);
    if IsFrozen then
      Result := FFrozenFirstHdr - H
    else
      Result := FFirstHdr - H;
  end;
  if Result < 0 then
    Result := 0;
end;

procedure TXLSBookHeaders.SetFocusedState(P1, P2: integer);
var
  P: integer;
begin
  if (P2 < 0) or (P1 > High(FHeaders)) then
    Exit;
  if P1 < 0 then
    P1 := 0;
  if P2 > High(FHeaders) then
    P2 := High(FHeaders);
  for P := 0 to P1 - 1 do
    FHeaders[P].SelState := hssNormal;
  for P := P1 to P2 do
    FHeaders[P].SelState := hssFocused;
  for P := P2 + 1 to High(FHeaders) do
    FHeaders[P].SelState := hssNormal;
end;

procedure TXLSBookHeaders.SetSelected(Index: integer; const Value: THeaderSelState);
begin
{$ifdef XDEBUG}
  if (Index < 0) or (Index > High(FHeaders)) then
    raise XLSRWException.Create('Header index out of range SetSelected');
{$endif}
  FHeaders[Index].SelState := Value;
end;

procedure TXLSBookHeaders.ClearSelected;
var
  i: integer;
begin
  for i := 0 to High(FHeaders) do
    FHeaders[i].SelState := hssNormal;
end;

procedure TXLSBookHeaders.SetSelectedState(P1, P2: integer; IsSelected: boolean);
var
  P: integer;
begin
  if (P2 < 0) or (P1 > High(FHeaders)) then
    Exit;
  if P1 < 0 then
    P1 := 0;
  if P2 > High(FHeaders) then
    P2 := High(FHeaders);
  for P := P1 to P2 do begin
    if IsSelected then
      FHeaders[P].SelState := hssSelected
    else if FHeaders[P].SelState < hssSelected then
      FHeaders[P].SelState := hssFocused;
  end;
end;

function TXLSBookHeaders.PosToHeader(P: integer): integer;
var
  i: integer;
begin
  Result := FFirstHdr;
  if P < FHeaders[0].Pos then
    Exit;
  for i := 0 to High(FHeaders) do begin
    if (P >= FHeaders[i].Pos) and (P <= (FHeaders[i].Pos + GetVisibleSize(i) {FHeaders[i].Size})) then
      Exit;
    Inc(Result);
  end;
end;

function TXLSBookHeaders.GetHeight(Hdr1, Hdr2: integer): integer;
var
  h: integer;
  Sz,Outline: integer;
  Options: THeaderOptions;
begin
  Result := 0;
  for h := Hdr1 to Hdr2 do begin
    GetHeaderData(h,Sz,Outline,Options);
    Inc(Result,sz);
  end;
end;

function TXLSBookHeaders.GetHalfPos(Index: integer): integer;
begin
{$ifdef XDEBUG}
  if (Index < 0) or (Index > High(FHeaders)) then
    raise XLSRWException.Create('Header index out of range HalfPos');
{$endif}
  Result := FHeaders[Index].Pos + GetVisibleSize(Index) div 2;
end;

function TXLSBookHeaders.GetAbsPos(Index: integer): integer;
begin
  Dec(Index,FFirstHdr);
{$ifdef XDEBUG}
  if (Index < 0) or (Index > High(FHeaders)) then begin
    raise XLSRWException.CreateFmt('Header index out of range AbsPos %d(%d)',[Index,High(FHeaders)]);
  end;
{$endif}
  Result := FHeaders[Index].Pos;
end;

function TXLSBookHeaders.GetAbsVisibleSize(Index: integer): integer;
begin
  Dec(Index,FFirstHdr);
{$ifdef XDEBUG}
  if (Index < 0) or (Index > High(FHeaders)) then begin
    raise XLSRWException.Create('Header index out of range AbsVisibleSize');
  end;
{$endif}
  if hoHidden in FHeaders[Index].Options then
    Result := 0
  else
    Result := FHeaders[Index].Size;
end;

function TXLSBookHeaders.GetAbsVisibleSizeEmu(Index: integer): integer;
begin
  Dec(Index,FFirstHdr);
{$ifdef XDEBUG}
  if (Index < 0) or (Index > High(FHeaders)) then begin
    raise XLSRWException.Create('Header index out of range AbsVisibleSize');
  end;
{$endif}
  if hoHidden in FHeaders[Index].Options then
    Result := 0
  else
    Result := FHeaders[Index].Size;
end;

{ TXHeadersGutter }

procedure TXHeadersGutter.CalcButtons(Headers: array of THeaderData; FirstHdr: integer);
var
  i,j,p,cx,cy,Level,CurrLevel: integer;
begin
  FButtons.Clear;
  CurrLevel := FFirstLevel;
  p := 0;
  for i := 0 to High(Headers) do begin
    Level := Headers[i].OutlineLevel;
    if (CurrLevel > 0) or (Headers[i].OutlineLevel > 0) then begin
      if not (hoHidden in Headers[i].Options) then begin
        if Level < CurrLevel then begin
          for j := Level to CurrLevel - 1 do begin
            CentrePos(p,p + Headers[i].Size,j * GUTTER_SIZE,cx,cy);
            Dec(cx,(GUTTER_SIZE div 2));
            Dec(cy,(GUTTER_SIZE div 2));
            if hoCollapsed in Headers[i].Options then
              FButtons.Add(cx,cy,1,FirstHdr + i,Level)
            else
              FButtons.Add(cx,cy,0,FirstHdr + i,Level);
          end;
        end;
      end;
      CurrLevel := Level;
    end;
    if not (hoHidden in Headers[i].Options) then
      Inc(p,Headers[i].Size);
  end;
end;

constructor TXHeadersGutter.Create(AParent: TXSSWindow);
begin
  inherited Create(AParent);
  FButtons := TButtonRectList.Create(FSkin,bbtOutline);
  FCurrButton := -1;
end;

destructor TXHeadersGutter.Destroy;
begin
  FButtons.Free;
  inherited;
end;

procedure TXHeadersGutter.MouseDown(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer);
begin
  inherited MouseDown(Button,Shift,X,Y);
  FSkin.GDI.SetCursor(xctArrow);
  FButtons.MouseDown(Button,Shift,X,Y);
end;

procedure TXHeadersGutter.MouseLeave;
begin
  inherited;
  FButtons.MouseLeave;
end;

procedure TXHeadersGutter.MouseMove(Shift: TXSSShiftState; X, Y: Integer);
begin
  inherited MouseMove(Shift,X,Y);

  FSkin.GDI.SetCursor(xctArrow);
  FButtons.MouseMove(Shift,X,Y);
end;

procedure TXHeadersGutter.MouseUp(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer);
begin
  inherited MouseUp(Button,Shift,X,Y);
  FButtons.MouseUp(Button,Shift,X,Y);
end;

procedure TXHeadersGutter.Paint(Headers: array of THeaderData);
var
  i,j,x,w,Level,PrevLevel: integer;
begin
  BeginPaint;

  FSkin.PaintHeader(FX1,FY1,FX2,FY2,hssNormal,xshsGutter);

  FSkin.GDI.PaintColor := 0;
  PrevLevel := FFirstVisibleLevel;
  x := 0;
  for i := 0 to High(Headers) do begin
    Level := Headers[i].OutlineLevel;
    if (PrevLevel > 0) or (Level > 0) then begin
      if not (hoHidden in Headers[i].Options) then begin
        w := Headers[i].Size;
        if Level > PrevLevel then begin
          for j := PrevLevel to Level - 1 do
            PaintOutline(x,x + w,j,olpFirst);
          for j := 0 to Level - 2 do
            PaintOutline(x,x + w,j,olpMiddle);
          PaintOutline(x,x + w,Level,olpDot);
        end
        else if Level = PrevLevel then begin
          for j := 0 to Level - 1 do
            PaintOutline(x,x + w,j,olpMiddle);
          PaintOutline(x,x + w,Level,olpDot);
        end
        else if Level < PrevLevel then begin
          for j := PrevLevel - 1 downto Level do
            PaintOutline(x,x + w,j,olpLast);
          for j := 0 to Level - 1 do
            PaintOutline(x,x + w,j,olpMiddle);
        end;
        PrevLevel := Level;
      end;
    end;
    if not (hoHidden in Headers[i].Options) then
      Inc(x,Headers[i].Size);
  end;
  FButtons.Paint;

  EndPaint;
end;

end.
