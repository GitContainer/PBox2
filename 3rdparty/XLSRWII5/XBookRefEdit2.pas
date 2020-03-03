unit XBookRefEdit2;

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

uses { Delphi  } Classes, SysUtils, vcl.Controls, Contnrs,
     { XLSRWII } Xc12Utils5, Xc12DataWorkbook5, Xc12Manager5,
                 XLSUtils5, XLSFormulaTypes5, XLSEncodeFmla5, XLSFormula5,
                 XLSFmlaDebugData5,
                 XSSIEKeys,
     { XLSBook } XBookUtils2, XBookPaintGDI2, XBookPaint2, XBookSkin2, XBookColumns2,
                 XBookRows2, XBookWindows2, XBookTypes2, XBookInplaceEdit2;

type TRefEditMode = (remIdle,remMove,remSize);
type TRefScrollDirection = (rsdNone,rsdHoriz,rsdVert);
type TRefEditType = (retUndefined,retRef,retRef1d,retName,retSelect);

type TXBookEditRef = class(TObject)
private
     FVisible     : boolean;
     FType        : TRefEditType;
     FCol1,FRow1,
     FCol2,FRow2  : integer;
     FAbsC1,FAbsR1,
     FAbsC2,FAbsR2: boolean;
     FDeltaCol,
     FDeltaRow    : integer;
     // Index to char run in editor.
     FCRIndex     : integer;
     FName        : AxUCString;
     FRect        : TXYRect;
     FColor       : TXc12RGBColor;
     FHasGrips    : boolean;
     FGrips       : array[0..3] of TXYRect;
public
     constructor Create(C1,R1,C2,R2: integer);
     procedure Move;
     procedure Size(Corner: integer);
     procedure SetInside(var Col,Row: integer);
     function  Inside(Col,Row: integer): boolean;
     procedure FitToExtent;
     function  AsString: AxUCString;

     property Type_: TRefEditType read FType;
     property Col1: integer read FCol1 write FCol1;
     property Row1: integer read FRow1 write FRow1;
     property Col2: integer read FCol2 write FCol2;
     property Row2: integer read FRow2 write FRow2;
     property Rect: TXYRect read FRect write FRect;
     property AbsC1: boolean read FAbsC1;
     property AbsR1: boolean read FAbsR1;
     property AbsC2: boolean read FAbsC2;
     property AbsR2: boolean read FAbsR2;
     property CRIndex: integer read FCRIndex write FCRIndex;
     property Name: AxUCString read FName;
     property Color: TXc12RGBColor read FColor write FColor;
     property HasGrips: boolean read FHasGrips write FHasGrips;
     end;

type TXBookEditRefList = class(TObjectList)
private
     function  GetItems(Index: integer): TXBookEditRef;
protected
     FManager        : TXc12Manager;
     FEditor         : TXBookInplaceEditor;
     FFormula        : AxUCString;
     FWin            : TXSSClientWindow;
     FColumns        : TXLSBookColumns;
     FRows           : TXLSBookRows;
     FCurrColor      : integer;
     FFocused        : integer;
     FExcludePaint   : integer;
     FPaintCellsEvent: TX4IntegerEvent;
     FDoScrollEvent  : TXIntegerEvent;
     FEditMode       : TRefEditMode;
     FLastCol,
     FLastRow        : integer;
     FClickCol,
     FClickRow       : integer;
     FSizeCorner     : integer;
     FScrollDir      : TRefScrollDirection;
     FIsUpdating     : boolean;
     FDebugEvent     : TXStringEvent;

     procedure SetPos(ERef: TXBookEditRef; const AUpdateEditor: boolean = True);
     procedure SetSize(ERef: TXBookEditRef; Corner: integer);
     procedure PaintRef(Index: integer; Focused: boolean);
     function  RefHit(X,Y: integer): integer;
     procedure ClipToSheet(var C1,R1,C2,R2: integer);
     function  SetGripCursor(X,Y: integer): boolean;
     procedure ClipXYToWin(ERef: TXBookEditRef; C1,R1,C2,R2: integer);
     procedure UpdateEditor(AERef: TXBookEditRef; const C1,R1,C2,R2: integer);

     function  Add(const C,R: integer): TXBookEditRef; overload;
     function  Add(const C1,R1,C2,R2: integer; const AType: TRefEditType = retRef; const AName: AxUCString = ''): TXBookEditRef; overload;
     function  Add(const C,R: integer; const ASheetName: AxUCString): TXBookEditRef; overload;
     function  Add(const C1,R1,C2,R2: integer; const ASheetName: AxUCString): TXBookEditRef; overload;
     function  AddName(const C,R: integer; const AName: AxUCString): TXBookEditRef; overload;
     function  AddName(const C1,R1,C2,R2: integer; const AName: AxUCString): TXBookEditRef; overload;
     procedure ScanPtgs(ADebugItems: TFmlaDebugItems);
     procedure EditorTextChanged(const AText: AxUCString);
public
     constructor Create(AManager: TXc12Manager; AEditor: TXBookInplaceEditor; AWin: TXSSClientWindow; Columns: TXLSBookColumns; Rows: TXLSBookRows);
     destructor Destroy; override;

     procedure MouseDown(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y, Col, Row: Integer);
     procedure MouseMove(Shift: TXSSShiftState; X, Y, Col, Row: Integer);
     procedure MouseUp(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y, Col, Row: Integer);
     procedure Paint;
     procedure ScrollUpdate(Col,Row: integer);
     function  IsFocused: boolean;

     function  HasSelect: boolean;
     procedure AddSelect(const C,R: integer);

     procedure SetFormula(const AFormula: AxUCString); overload;
     procedure SetFormula(const AFormula: AxUCString; const AStrictMode: boolean); overload;

     property Items[Index: integer]: TXBookEditRef read GetItems; default;

     property OnPaintCells: TX4IntegerEvent read FPaintCellsEvent write FPaintCellsEvent;
     property OnDoScroll: TXIntegerEvent read FDoScrollEvent write FDoScrollEvent;
     property OnDebug: TXStringEvent read FDebugEvent write FDebugEvent;
     end;

implementation

const GRIPSZ_2 = 2;
      CLICKRANGE = 4;

var
  ColorPool: array[0..5] of TXc12RGBColor = ($0000FF,$008000,$9900CC,$800000,$FF6600,$CC0099);
///  ColorPool: array[0..5] of TXc12RGBColor = ($FF0000,$00FF00,$0000FF,$800000,$FFFF00,$FF00FF);

{ TXBookEditRefList }

function TXBookEditRefList.Add(const C1, R1, C2, R2: integer; const AType: TRefEditType; const AName: AxUCString): TXBookEditRef;
begin
  Result := TXBookEditRef.Create(C1,R1,C2,R2);
  Result.FType := AType;
  Result.FName := AName;
  Result.FColor := ColorPool[FCurrColor];
  Inc(FCurrColor);
  if FCurrColor > High(ColorPool) then
    FCurrColor := 0;
  Result.HasGrips := AType in [retRef,retRef1d,retSelect];
  SetPos(Result,False);
  inherited Add(Result);
end;

function TXBookEditRefList.Add(const C, R: integer): TXBookEditRef;
begin
  Result := Add(C,R,C,R,retRef);
end;

function TXBookEditRefList.AddName(const C, R: integer; const AName: AxUCString): TXBookEditRef;
begin
  Result := Add(C,R,C,R,retName,AName);
end;

function TXBookEditRefList.AddName(const C1, R1, C2, R2: integer; const AName: AxUCString): TXBookEditRef;
begin
  Result := Add(C1,R1,C2,R2,retName,AName);
end;

procedure TXBookEditRefList.AddSelect(const C, R: integer);
var
  S: AxUCString;
  Ref: TXBookEditRef;
begin
  Ref := Add(C,R,C,R,retSelect);
  FFocused := Count - 1;

  S := ColRowToRefStr(C,R);
  // Do NOT call BeginAddRefs. That will clear all editor text.
//  FEditor.BeginAddRefs;
  Ref.CRIndex := FEditor.AddRefText(S,RevRGB(Ref.Color));
  FEditor.EndAddRefs;

  PaintRef(FFocused,True);
end;

function TXBookEditRefList.Add(const C1, R1, C2, R2: integer; const ASheetName: AxUCString): TXBookEditRef;
begin
  Result := Add(C1,R1,C2,R2,retRef1d,ASheetName);
end;

function TXBookEditRefList.Add(const C, R: integer; const ASheetName: AxUCString): TXBookEditRef;
begin
  Result := Add(C,R,C,R,retRef1d,ASheetName);
end;

procedure TXBookEditRefList.ClipToSheet(var C1, R1, C2, R2: integer);
begin
  if C1 < FColumns.LeftCol then C1 := FColumns.LeftCol;
  if R1 < FRows.TopRow then R1 := FRows.TopRow;
  if C2 > FColumns.RightCol then C1 := FColumns.RightCol;
  if R2 > FRows.BottomRow then R1 := FRows.BottomRow;
end;

procedure TXBookEditRefList.ClipXYToWin(ERef: TXBookEditRef; C1, R1, C2, R2: integer);
begin
  if C1 < FColumns.LeftCol then
    ERef.FRect.X1 := FColumns.X1 - GRIPSZ_2 - 1
  else
    ERef.FRect.X1 := FColumns.AbsPos[C1];
  if C2 > FColumns.RightCol then
    ERef.FRect.X2 := FColumns.X2 + GRIPSZ_2 + 1
  else
    ERef.FRect.X2 := FColumns.AbsPos[C2] + FColumns.AbsVisibleSize[C2] - 2;

  if R1 < FRows.TopRow then
    ERef.FRect.Y1 := FRows.Y1 - GRIPSZ_2 - 1
  else
    ERef.FRect.Y1 := FRows.AbsPos[R1];
  if R2 > FRows.BottomRow then
    ERef.FRect.Y2 := FRows.Y2 + GRIPSZ_2 + 1
  else
    ERef.FRect.Y2 := FRows.AbsPos[R2] + FRows.AbsVisibleSize[R2] - 2;
end;

constructor TXBookEditRefList.Create(AManager: TXc12Manager; AEditor: TXBookInplaceEditor; AWin: TXSSClientWindow; Columns: TXLSBookColumns; Rows: TXLSBookRows);
begin
  inherited Create;

  FManager := AManager;
  FEditor := AEditor;
  FEditor.OnTextChanged := EditorTextChanged;
  FWin := AWin;
  FColumns := Columns;
  FRows := Rows;
  FFocused := -1;
  FExcludePaint := -1;
end;

destructor TXBookEditRefList.Destroy;
begin
  inherited;
end;

procedure TXBookEditRefList.EditorTextChanged(const AText: AxUCString);
begin
  FEditor.OnTextChanged := Nil;
//  try
    SetFormula(AText);
//  finally
//    FEditor.OnTextChanged := EditorTextChanged;
//  end;
end;

function TXBookEditRefList.GetItems(Index: integer): TXBookEditRef;
begin
  Result := TXBookEditRef(inherited Items[Index]);
end;

function TXBookEditRefList.HasSelect: boolean;
begin
  Result := (Count > 0) and (Items[Count - 1].Type_ = retSelect);
end;

function TXBookEditRefList.IsFocused: boolean;
begin
  Result := FFocused >= 0;
end;

procedure TXBookEditRefList.MouseDown(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y, Col, Row: Integer);
var
  i: integer;
begin
  i := RefHit(X,Y);
  if i >= 0 then begin
    FFocused := i;
    FLastCol := -1;
    FLastRow := -1;
    Items[FFocused].FDeltaCol := 0;
    Items[FFocused].FDeltaRow := 0;
    if SetGripCursor(X,Y) then begin
      FEditMode := remSize;
      case FSizeCorner of
        0: begin
          FClickCol := Items[FFocused].FCol1;
          FClickRow := Items[FFocused].FRow1;
        end;
        1: begin
          FClickCol := Items[FFocused].FCol2;
          FClickRow := Items[FFocused].FRow1;
        end;
        2: begin
          FClickCol := Items[FFocused].FCol2;
          FClickRow := Items[FFocused].FRow2;
        end;
        3: begin
          FClickCol := Items[FFocused].FCol1;
          FClickRow := Items[FFocused].FRow2;
        end;
      end;
    end
    else begin
      FClickCol := Col;
      FClickRow := Row;
      if Items[FFocused].Type_= retSelect then begin
        FEditMode := remSize;
        FWin.Skin.GDI.SetCursor(xctCell);
      end
      else begin
        FEditMode := remMove;
        FWin.Skin.GDI.SetCursor(xctArrow);
      end;
    end;
    Items[FFocused].SetInside(FClickCol,FClickRow);
  end;
end;

procedure TXBookEditRefList.MouseMove(Shift: TXSSShiftState; X, Y, Col, Row: Integer);
var
  i: integer;
  Ref: TXBookEditRef;
  C1,R1,C2,R2: integer;

procedure PaintPrev;
begin
  if FFocused >= 0 then begin
    FExcludePaint := FFocused;
    if Assigned(FPaintCellsEvent) then
      FPaintCellsEvent(Self,Items[FFocused].FCol1,Items[FFocused].FRow1,Items[FFocused].FCol2,Items[FFocused].FRow2);
    FExcludePaint := -1;
    PaintRef(FFocused,False);
    FFocused := -1;
  end;
end;

procedure CheckScroll;
begin
  if not FWin.Hit(X,Y) then begin
    if FScrollDir = rsdNone then begin
      if X < FWin.CX1 then begin
        FDoScrollEvent(Self,0);
        FScrollDir := rsdHoriz;
      end
      else if X > FWin.CX2 then begin
        FDoScrollEvent(Self,1);
        FScrollDir := rsdHoriz;
      end
      else if Y < FWin.CY1 then begin
        FDoScrollEvent(Self,2);
        FScrollDir := rsdVert;
      end
      else if Y > FWin.CY2 then begin
        FDoScrollEvent(Self,3);
        FScrollDir := rsdVert;
      end;
    end;
  end
  else if FScrollDir <> rsdNone then begin
    FDoScrollEvent(Self,-1);
    FScrollDir := rsdNone;
  end;
end;

begin
  case FEditMode of
    remIdle: begin
      PaintPrev;
      i := RefHit(X,Y);
      if i >= 0 then begin
        if i <> FFocused then begin
          PaintPrev;
          PaintRef(i,True);
          FFocused := i;
        end;
        if (FFocused >= 0) and not SetGripCursor(X,Y) then
          FWin.Skin.GDI.SetCursor(xctMoveCells);
      end;
    end;
    remMove: begin
      FWin.Skin.GDI.SetCursor(xctArrow);
      CheckScroll;
      if ((FScrollDir = rsdVert) and (Col = FLastCol)) or ((FScrollDir = rsdHoriz) and (Row = FLastRow)) then
        Exit;
      if (Col <> FLastCol) or (Row <> FLastRow) then begin
        C1 := (Items[FFocused].FCol1 + Items[FFocused].FDeltaCol) - 1;
        R1 := (Items[FFocused].FRow1 + Items[FFocused].FDeltaRow) - 1;
        C2 := (Items[FFocused].FCol2 + Items[FFocused].FDeltaCol) + 1;
        R2 := (Items[FFocused].FRow2 + Items[FFocused].FDeltaRow) + 1;

        FExcludePaint := FFocused;
        FPaintCellsEvent(Self,C1,R1,C2,R2);
        FExcludePaint := -1;

        Items[FFocused].FDeltaCol := Col - FClickCol;
        Items[FFocused].FDeltaRow := Row - FClickRow;
        Items[FFocused].FitToExtent;

        SetPos(Items[FFocused]);
        PaintRef(FFocused,True);
        FLastCol := Col;
        FLastRow := Row;
      end;
    end;
    remSize: begin
      if Items[FFocused].Type_ = retSelect then
        FWin.Skin.GDI.SetCursor(xctCell)
      else if FSizeCorner in [1,3] then
        FWin.Skin.GDI.SetCursor(xctNeSwArrow)
      else
        FWin.Skin.GDI.SetCursor(xctNwSeArrow);
      CheckScroll;
      if ((FScrollDir = rsdVert) and (Col = FLastCol)) or ((FScrollDir = rsdHoriz) and (Row = FLastRow)) then
        Exit;
      if ((Col <> FLastCol) or (Row <> FLastRow)) {and not Items[FFocused].Inside(Col,Row)} then begin
        Ref := Items[FFocused];
        C1 := Items[FFocused].FCol1 - 1;
        R1 := Items[FFocused].FRow1 - 1;
        C2 := Items[FFocused].FCol2 + 1;
        R2 := Items[FFocused].FRow2 + 1;
        case FSizeCorner of
          0: begin
            C1 := Ref.FCol1 + Items[FFocused].FDeltaCol - 1;
            R1 := Ref.FRow1 + Ref.FDeltaRow - 1;
          end;
          1: begin
            C2 := Ref.FCol2 + Ref.FDeltaCol + 1;
            R1 := Ref.FRow1 + Ref.FDeltaRow - 1;
          end;
          2: begin
            C2 := Ref.FCol2 + Ref.FDeltaCol + 1;
            R2 := Ref.FRow2 + Ref.FDeltaRow + 1;
          end;
          3: begin
            C1 := Ref.FCol1 + Ref.FDeltaCol - 1;
            R2 := Ref.FRow2 + Ref.FDeltaRow + 1;
          end;
        end;
        FExcludePaint := FFocused;
        FPaintCellsEvent(Self,C1,R1,C2,R2);
        FExcludePaint := -1;

        if (Col < Ref.FCol1) and (FSizeCorner in [1,2]) then begin
          FClickCol := Ref.FCol1;
          Ref.FCol2 := Ref.FCol1;
          if FSizeCorner = 1 then
            FSizeCorner := 0
          else
            FSizeCorner := 3;
        end
        else if (Col > Ref.FCol2) and (FSizeCorner in [0,3]) then begin
          FClickCol := Ref.FCol2;
          Ref.FCol1 := Ref.FCol2;
          if FSizeCorner = 0 then
            FSizeCorner := 1
          else
            FSizeCorner := 2;
        end;
        if (Row < Ref.FRow1) and (FSizeCorner in [2,3]) then begin
          FClickRow := Ref.FRow1;
          Ref.FRow2 := Ref.FRow1;
          if FSizeCorner = 2 then
            FSizeCorner := 1
          else
            FSizeCorner := 0;
        end
        else if (Row > Ref.FRow2) and (FSizeCorner in [0,1]) then begin
          FClickRow := Ref.FRow2;
          Ref.FRow1 := Ref.FRow2;
          if FSizeCorner = 0 then
            FSizeCorner := 3
          else
            FSizeCorner := 2;
        end;

        Ref.FDeltaCol := Col - FClickCol;
        Ref.FDeltaRow := Row - FClickRow;
        SetSize(Ref,FSizeCorner);
        PaintRef(FFocused,True);
        FLastCol := Col;
        FLastRow := Row;
      end;
    end;
  end;
end;

procedure TXBookEditRefList.MouseUp(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y, Col, Row: Integer);
begin
  case FEditMode of
    remMove: Items[FFocused].Move;
    remSize: Items[FFocused].Size(FSizeCorner);
  end;
  Items[FFocused].FDeltaCol := 0;
  Items[FFocused].FDeltaRow := 0;
  if Items[FFocused].Type_ = retSelect then begin
    Items[FFocused].FType := retRef;

    FEditor.AddRefText('');
    FEditor.EndAddRefs;
  end;
  if FEditMode <> remIdle then
    SetPos(Items[FFocused]);
//  FFocused := -1;
  FEditMode := remIdle;
end;

procedure TXBookEditRefList.Paint;
var
  i: integer;
begin
  for i := Count - 1 downto 0 do begin
    if i <> FExcludePaint then
      PaintRef(i,i = FFocused);
  end;
end;

procedure TXBookEditRefList.PaintRef(Index: integer; Focused: boolean);
begin
  if not Items[Index].FVisible then
    Exit;

  FWin.Skin.GDI.PaintColor:= Items[Index].FColor;
  if Focused then
    FWin.Skin.Border(Items[Index].Rect.X1,Items[Index].Rect.Y1,Items[Index].Rect.X2,Items[Index].Rect.Y2,2)
  else
    FWin.Skin.Border(Items[Index].Rect.X1,Items[Index].Rect.Y1,Items[Index].Rect.X2,Items[Index].Rect.Y2,1);
  if Items[Index].HasGrips then begin
    FWin.Skin.GDI.PaintColor := Items[Index].FColor;
    FWin.Skin.GDI.Rectangle(Items[Index].FGrips[0]);
    FWin.Skin.GDI.Rectangle(Items[Index].FGrips[1]);
    FWin.Skin.GDI.Rectangle(Items[Index].FGrips[2]);
    FWin.Skin.GDI.Rectangle(Items[Index].FGrips[3]);
  end;
end;

function TXBookEditRefList.RefHit(X, Y: integer): integer;
var
  R: TXYRect;
begin
  for Result := Count - 1 downto 0 do begin
    R := SizeXYRect(Items[Result].FRect,CLICKRANGE);
    if Items[Result].FHasGrips and PtInXYRect(X,Y,R) then begin
      if Items[Result].Type_ in [retRef,retRef1d] then begin
        R := SizeXYRect(Items[Result].FRect,-CLICKRANGE);
        if not PtInXYRect(X,Y,R) then
          Exit;
      end
      else
        Exit;
    end;
  end;
  Result := -1;
end;

procedure TXBookEditRefList.ScanPtgs(ADebugItems: TFmlaDebugItems);
var
  i   : integer;
  j   : integer;
  S   : AxUCString;
  Item: TFmlaDebugItem;
  ERef: TXBookEditRef;
  N   : TXc12DefinedName;
begin
  Clear;

  j := 1;

  FEditor.BeginAddRefs;
  for i := 0 to ADebugItems.Count - 1 do begin
    ERef := Nil;
    Item := ADebugItems[i];
    if Item._Ptgs = Nil then
      Continue;
    case Item._Ptgs.Id of
      xptgRef   : ERef := Add(PXLSPtgsRef(Item._Ptgs).Col,PXLSPtgsRef(Item._Ptgs).Row);
      xptgArea  : ERef := Add(PXLSPtgsArea(Item._Ptgs).Col1,PXLSPtgsArea(Item._Ptgs).Row1,PXLSPtgsArea(Item._Ptgs).Col2,PXLSPtgsArea(Item._Ptgs).Row2);
      xptgRef1d : begin
        if PXLSPtgsRef1d(Item._Ptgs).Sheet = FEditor.SheetIndex then
          ERef := Add(PXLSPtgsRef1d(Item._Ptgs).Col,PXLSPtgsRef1d(Item._Ptgs).Row,FManager.Worksheets[PXLSPtgsRef1d(Item._Ptgs).Sheet].Name);
      end;
      xptgArea1d: begin
        if PXLSPtgsArea1d(Item._Ptgs).Sheet = FEditor.SheetIndex then
          ERef := Add(PXLSPtgsArea1d(Item._Ptgs).Col1,PXLSPtgsArea1d(Item._Ptgs).Row1,PXLSPtgsArea1d(Item._Ptgs).Col2,PXLSPtgsArea1d(Item._Ptgs).Row2,FManager.Worksheets[PXLSPtgsArea1d(Item._Ptgs).Sheet].Name);
      end;
      xptgName  : begin
        N := FManager.Workbook.DefinedNames[PXLSPtgsName(Item._Ptgs).NameId];
        if N.SimpleArea.SheetIndex = FEditor.SheetIndex then begin
          case N.SimpleName of
            xsntNone : ;
            xsntRef  : AddName(N.SimpleArea.Col1,N.SimpleArea.Row1,N.Name);
            xsntArea : AddName(N.SimpleArea.Col1,N.SimpleArea.Row1,N.SimpleArea.Col2,N.SimpleArea.Row2,N.Name);
            xsntError: ;
          end;
        end;
      end;
    end;
    if ERef <> Nil then begin
      S := Copy(FFormula,j,Item.P1 - j);
      if S <> '' then
        FEditor.AddRefText(S);
      ERef.CRIndex := FEditor.AddRefText(Copy(FFormula,Item.P1,Item.P2 - Item.P1 + 1),RevRGB(ERef.Color));
      j := Item.P2 + 1;

      PaintRef(Count - 1,False);
    end;
  end;

  FEditor.AddRefText(Copy(FFormula,j,MAXINT));

  FEditor.EndAddRefs;
end;

procedure TXBookEditRefList.ScrollUpdate(Col, Row: integer);
var
  i: integer;
  Item: TXBookEditRef;
begin
  Item := Items[FFocused];
  Item.FDeltaCol := Item.FDeltaCol + Col;
  Item.FDeltaRow := Item.FDeltaRow + Row;
  if FEditMode = remMove then begin
    if (Item.FRow1 + Item.FDeltaRow) >= FRows.BottomRow then
      Item.FDeltaRow := FRows.BottomRow - Item.FRow1 - 1;
    if (Item.FCol1 + Item.FDeltaCol) >= FColumns.RightCol then
      Item.FDeltaCol := FColumns.RightCol - Item.FCol1 - 1;
    Item.FitToExtent;
    for i := 0 to Count - 1 do
      SetPos(Items[i]);
  end
  else begin
    if (Item.FRow2 + Item.FDeltaRow) >= FRows.BottomRow then
      Item.FDeltaRow := FRows.BottomRow - Item.FRow2 - 2;
    if (Item.FCol2 + Item.FDeltaCol) >= FColumns.RightCol then
      Item.FDeltaCol := FColumns.RightCol - Item.FCol2 - 2;
    Item.FitToExtent;
    for i := 0 to Count - 1 do
      SetSize(Items[i],FSizeCorner);
  end;
end;

procedure TXBookEditRefList.SetFormula(const AFormula: AxUCString; const AStrictMode: boolean);
var
  Ptgs: PXLSPtgs;
  PtgsSz: integer;
  DebugItems: TFmlaDebugItems;
begin
  if (AFormula <> '') and not FIsUpdating then begin
    FIsUpdating := True;
    FFormula := AFormula;

//    FEditor.BeginUpdate;

    DebugItems := TFmlaDebugItems.Create;
    try
      FManager.Errors.IgnoreErrors := True;
      try
        if XLSEncodeFormula(FManager,FFormula,Ptgs,PtgsSz,AStrictMode,DebugItems) then
          ScanPtgs(DebugItems);
      finally
        FManager.Errors.IgnoreErrors := False;
      end;
    finally
      DebugItems.Free;
      FreeMem(Ptgs);
    end;
//    FEditor.EndUpdate;
//    FEditor.Command(axcMoveEndOfDoc);

    FIsUpdating := False;
  end;
end;

procedure TXBookEditRefList.SetFormula(const AFormula: AxUCString);
begin
  SetFormula(AFormula,False);
end;

function TXBookEditRefList.SetGripCursor(X, Y: integer): boolean;
var
  i: integer;
begin
  for i := 0 to High(Items[FFocused].FGrips) do begin
    if PtInXYRect(X,Y,Items[FFocused].FGrips[i]) then begin
      FSizeCorner := i;
      if FSizeCorner in [1,3] then
        FWin.Skin.GDI.SetCursor(xctNeSwArrow)
      else
        FWin.Skin.GDI.SetCursor(xctNwSeArrow);
      Result := True;
      Exit;
    end;
  end;
  Result := False;
end;

procedure TXBookEditRefList.SetPos(ERef: TXBookEditRef; const AUpdateEditor: boolean = True);
var
  C1,R1,C2,R2: integer;
begin
  C1 := ERef.FCol1 + ERef.FDeltaCol;
  R1 := ERef.FRow1 + ERef.FDeltaRow;
  C2 := ERef.FCol2 + ERef.FDeltaCol;
  R2 := ERef.FRow2 + ERef.FDeltaRow;

  if AUpdateEditor then
    UpdateEditor(ERef,C1,R1,C2,R2);

  ERef.FVisible := not ((C1 > FColumns.RightCol) or (C2 < FColumns.LeftCol) or (R1 > FRows.BottomRow) or (R2 < FRows.TopRow));

  if not ERef.FVisible then
    Exit;

  if C1 > FColumns.RightCol then C1 := FColumns.RightCol;
  if C2 < FColumns.LeftCol  then C2 := FColumns.LeftCol;
  if R1 > FRows.BottomRow   then R1 := FRows.BottomRow;
  if R2 < FRows.TopRow      then R1 := FRows.TopRow;

  ClipXYToWin(ERef,C1,R1,C2,R2);

  SetXYRect(ERef.FGrips[0],ERef.FRect.X1 - GRIPSZ_2,ERef.FRect.Y1 - GRIPSZ_2,ERef.FRect.X1 + GRIPSZ_2,ERef.FRect.Y1 + GRIPSZ_2);
  SetXYRect(ERef.FGrips[1],ERef.FRect.X2 - GRIPSZ_2,ERef.FRect.Y1 - GRIPSZ_2,ERef.FRect.X2 + GRIPSZ_2,ERef.FRect.Y1 + GRIPSZ_2);
  SetXYRect(ERef.FGrips[2],ERef.FRect.X2 - GRIPSZ_2,ERef.FRect.Y2 - GRIPSZ_2,ERef.FRect.X2 + GRIPSZ_2,ERef.FRect.Y2 + GRIPSZ_2);
  SetXYRect(ERef.FGrips[3],ERef.FRect.X1 - GRIPSZ_2,ERef.FRect.Y2 - GRIPSZ_2,ERef.FRect.X1 + GRIPSZ_2,ERef.FRect.Y2 + GRIPSZ_2);
end;

procedure TXBookEditRefList.SetSize(ERef: TXBookEditRef; Corner: integer);
var
  C1,R1,C2,R2: integer;
begin
  C1 := ERef.FCol1;
  R1 := ERef.FRow1;
  C2 := ERef.FCol2;
  R2 := ERef.FRow2;
  case Corner of
    0: begin
      C1 := ERef.FCol1 + ERef.FDeltaCol;
      R1 := ERef.FRow1 + ERef.FDeltaRow;
    end;
    1: begin
      C2 := ERef.FCol2 + ERef.FDeltaCol;
      R1 := ERef.FRow1 + ERef.FDeltaRow;
    end;
    2: begin
      C2 := ERef.FCol2 + ERef.FDeltaCol;
      R2 := ERef.FRow2 + ERef.FDeltaRow;
    end;
    3: begin
      C1 := ERef.FCol1 + ERef.FDeltaCol;
      R2 := ERef.FRow2 + ERef.FDeltaRow;
    end;
  end;

  UpdateEditor(ERef,C1,R1,C2,R2);

  ERef.FVisible := not ((C1 > FColumns.RightCol) or (C2 < FColumns.LeftCol) or (R1 > FRows.BottomRow) or (R2 < FRows.TopRow));

  if not ERef.FVisible then
    Exit;

  if R1 < FRows.TopRow then
    ERef.FRect.Y1 := FRows.Y1 - GRIPSZ_2 - 1
  else
    ERef.FRect.Y1 := FRows.AbsPos[R1];
  if R2 > FRows.BottomRow then
    ERef.FRect.Y2 := FRows.Y2 + GRIPSZ_2 + 1
  else
    ERef.FRect.Y2 := FRows.AbsPos[R2] + FRows.AbsVisibleSize[R2] - 2;

  ClipXYToWin(ERef,C1,R1,C2,R2);

  SetXYRect(ERef.FGrips[0],ERef.FRect.X1 - GRIPSZ_2,ERef.FRect.Y1 - GRIPSZ_2,ERef.FRect.X1 + GRIPSZ_2,ERef.FRect.Y1 + GRIPSZ_2);
  SetXYRect(ERef.FGrips[1],ERef.FRect.X2 - GRIPSZ_2,ERef.FRect.Y1 - GRIPSZ_2,ERef.FRect.X2 + GRIPSZ_2,ERef.FRect.Y1 + GRIPSZ_2);
  SetXYRect(ERef.FGrips[2],ERef.FRect.X2 - GRIPSZ_2,ERef.FRect.Y2 - GRIPSZ_2,ERef.FRect.X2 + GRIPSZ_2,ERef.FRect.Y2 + GRIPSZ_2);
  SetXYRect(ERef.FGrips[3],ERef.FRect.X1 - GRIPSZ_2,ERef.FRect.Y2 - GRIPSZ_2,ERef.FRect.X1 + GRIPSZ_2,ERef.FRect.Y2 + GRIPSZ_2);
end;

procedure TXBookEditRefList.UpdateEditor(AERef: TXBookEditRef; const C1,R1,C2,R2: integer);
var
  S: AxUCString;
begin
  if AERef.Type_ in [retRef,retRef1d,retSelect] then begin
    if (C1 = C2) and (R1 = R2) then
      S := ColRowToRefStr(C1,R1,AERef.AbsC1,AERef.AbsR1)
    else
      S := AreaToRefStr(C1,R1,C2,R2,AERef.AbsC1,AERef.AbsR1,AERef.AbsC2,AERef.AbsR2);
    if AERef.Type_ = retRef1d then
      S := AERef.Name + '!' + S;
    FEditor.UpdateRefText(AEref.CRIndex,S);
    FEditor.EndAddRefs;
  end;
end;

{ TXBookEditRef }

function TXBookEditRef.AsString: AxUCString;
begin
  if (FCol1 = FCol2) and (FRow1 = FRow2) then
    Result := ColRowToRefStr(FCol1,FRow1,FAbsC1,FAbsR1)
  else
    Result := AreaToRefStr(FCol1,FRow1,FCol2,FRow2,FAbsC1,FAbsR1,FAbsC2,FAbsR2);
  case FType of
    retUndefined: ;
    retRef1d    : Result := FName + '!' + Result;
    retName     : Result := FName;
  end;
end;

constructor TXBookEditRef.Create(C1, R1, C2, R2: integer);
begin
  FCol1 := C1 and not COL_ABSFLAG;
  FRow1 := R1 and not ROW_ABSFLAG;
  FCol2 := C2 and not COL_ABSFLAG;
  FRow2 := R2 and not ROW_ABSFLAG;

  FAbsC1 := (C1 and COL_ABSFLAG) <> 0;
  FAbsR1 := (R1 and ROW_ABSFLAG) <> 0;
  FAbsC2 := (C2 and COL_ABSFLAG) <> 0;
  FAbsR2 := (R2 and ROW_ABSFLAG) <> 0;
end;

procedure TXBookEditRef.FitToExtent;
begin
  if (FCol1 + FDeltaCol) < 0 then
    FDeltaCol := -FCol1
  else if (FCol2 + FDeltaCol) > XLS_MAXCOL then
    FDeltaCol := XLS_MAXCOL - FCol2;
  if (FRow1 + FDeltaRow) < 0 then
    FDeltaRow := -FRow1
  else if (FRow2 + FDeltaRow) > XLS_MAXROW then
    FDeltaRow := XLS_MAXROW - FRow2;
end;

function TXBookEditRef.Inside(Col, Row: integer): boolean;
begin
  Result := (Col >= FCol1) and (Row >= FRow1) and (Col <= FCol2) and (Row <= FRow2);
end;

procedure TXBookEditRef.Move;
begin
  Inc(FCol1,FDeltaCol);
  Inc(FRow1,FDeltaRow);
  Inc(FCol2,FDeltaCol);
  Inc(FRow2,FDeltaRow);
end;

procedure TXBookEditRef.SetInside(var Col, Row: integer);
begin
       if Col < FCol1 then Col := FCol1
  else if Col > FCol2 then Col := FCol2;
       if Row < FRow1 then Row := Row1
  else if Row > FRow2 then Row := Row2;
end;

procedure TXBookEditRef.Size(Corner: integer);
begin
  case Corner of
    0: begin
      Inc(FCol1,FDeltaCol);
      Inc(FRow1,FDeltaRow);
    end;
    1: begin
      Inc(FCol2,FDeltaCol);
      Inc(FRow1,FDeltaRow);
    end;
    2: begin
      Inc(FCol2,FDeltaCol);
      Inc(FRow2,FDeltaRow);
    end;
    3: begin
      Inc(FCol1,FDeltaCol);
      Inc(FRow2,FDeltaRow);
    end;
  end;
end;

end.
