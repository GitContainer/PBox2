unit XBookDrawing2;

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

uses Classes, SysUtils, Contnrs, vcl.Graphics, Windows,
{$ifdef DELPHI_2009_OR_LATER}
     Vcl.Imaging.PNGImage, Vcl.Imaging.JPEG,
{$endif}
     Math,
     xpgParseDrawing, xpgParseDrawingCommon, xpgParseChart,
     Xc12Utils5, Xc12DataWorksheet5, Xc12Graphics, Xc12DataComments5, Xc12Manager5,
     Xc12DataStyleSheet5,
     XLSUtils5, XLSColumn5, XLSRow5, XLSComment5, XLSDrawing5, XLSTools5,
     XBookTypes2, XBookSysVar2, XBookUtils2, XBookPaintGDI2, XBookPaint2, XBookWindows2,
     XBookSkin2, XBookColumns2, XBookRows2, XBookRichPainter2, XBookGeometry2, XBookGeometryChart2,
     XLSEncodeFmla5, XLSFormulaTypes5;

type TXBookDrawingObject = class(TXSSClientWindow)
protected
public
     constructor Create(AParent: TXSSWindow);

     procedure GetPos(out AC1,AC1Offs,AR1,AR1Offs,AC2,AC2Offs,AR2,AR2Offs: integer); virtual; abstract;
     end;

type TXBookDrawingObjectComment = class(TXBookDrawingObject)
protected
     FComment: TXLSComment;
     FArrowStart: TXYPoint;
     FArrowEnd: TXYPoint;
public
     constructor Create(AParent: TXSSWindow; AComment: TXLSComment);

     procedure Paint; override;

     procedure GetPos(out AC1,AC1Offs,AR1,AR1Offs,AC2,AC2Offs,AR2,AR2Offs: integer); override;

     property Comment: TXLSComment read FComment;
     end;

type TXBookDrawingObjectAnchor = class(TXBookDrawingObject)
protected
     FAnchor: TCT_TwoCellAnchor;
public
     constructor Create(AParent: TXSSWindow; AAnchor: TCT_TwoCellAnchor);

     procedure GetPos(out AC1,AC1Offs,AR1,AR1Offs,AC2,AC2Offs,AR2,AR2Offs: integer); override;

     property Anchor: TCT_TwoCellAnchor read FAnchor write FAnchor;
     end;

type TXBookDrawingObjectImage = class(TXBookDrawingObjectAnchor)
protected
     FBitmap: vcl.Graphics.TBitmap;
     FXLSImage: TXLSDrawingImage;
public
     constructor Create(AParent: TXSSWindow; AXLSImage: TXLSDrawingImage);
     destructor Destroy; override;

     procedure Paint; override;

//     property Image: TXc12GraphicImage read FImage;
     end;

type TXBookDrawingObjectCommon = class(TXBookDrawingObjectAnchor)
protected
     procedure ApplySpPr(ASpPr: TCT_ShapeProperties);
public
     constructor Create(AParent: TXSSWindow; AAnchor: TCT_TwoCellAnchor);
     end;

type TXBookDrawingObjectTextBox = class(TXBookDrawingObjectCommon)
protected
     FXLSTextBox  : TXLSDrawingTextBox;
     FText        : TXLSFormattedText;
public
     constructor Create(AManager: TXc12Manager; AParent: TXSSWindow; AXLSTextBox: TXLSDrawingTextBox);
     destructor Destroy; override;

     procedure Paint; override;
     end;

type TXBookDrawingObjectChart = class(TXBookDrawingObjectCommon)
protected
     FDrwData   : TXLSDrawingChart;
     FChartSpace: TXSSChartItemChartSpace;
     FGeometry  : TXSSGeometryRect;
     FChartData : TXSSChartData;

     function  MapX(AX: double): double; {$ifdef D2006PLUS} inline; {$endif}
     function  MapY(AY: double): double; {$ifdef D2006PLUS} inline; {$endif}
     procedure PaintGeometry(AGeom: TXSSGeometry);
public
     constructor Create(AManager: TXc12Manager; AParent: TXSSWindow; AChartData: TXLSDrawingChart);
     destructor Destroy; override;

     procedure Paint; override;

     procedure Build;
     end;

type TXBookDrawing = class(TXSSClientWindow)
protected
     FManager   : TXc12Manager;
     FXLSDrawing: TXLSDrawing;

     FEnabled   : boolean;

     FColumns   : TXLSBookColumns;
     FRows      : TXLSBookRows;

     FCurrObj   : TXBookDrawingObject;

     function  AddImage(AXLSImage: TXLSDrawingImage): TXBookDrawingObjectImage;
     function  AddTextBox(AXLSTextBox: TXLSDrawingTextBox): TXBookDrawingObjectTextBox;
     function  AddChart(AChart: TXLSDrawingChart): TXBookDrawingObjectChart;
     procedure CalcObject(AChild: TXBookDrawingObject);
public
     constructor Create(AParent: TXSSClientWindow; AManager: TXc12Manager; AGDI: TAXWGDI; ASkinStyle: TXBookSkinStyle; AColumns: TXLSBookColumns; ARows: TXLSBookRows);
     destructor Destroy; override;

     procedure Clear;
     procedure Paint; override;
     procedure PaintPrint;
     function  Hit(const X,Y: integer): boolean; override;
     function  ClientHit(const X,Y: integer): boolean; override;
     procedure SetSize(const pX1,pY1,pX2,pY2: integer); override;

//     procedure MouseDown(Button: TMouseButton; Shift: TShiftState; X, Y: Integer); override;
//     procedure MouseMove(Shift: TShiftState; X, Y: Integer); override;
//     procedure MouseUp(Button: TMouseButton; Shift: TShiftState; X, Y: Integer); override;

     procedure SelectObject(const AIndex: integer);

     procedure AddComment(AComment: TXLSComment);
     procedure RemoveComment(AComment: TXLSComment);

     procedure CalcObjects;

     procedure LoadObjects(ASheet: TXc12DataWorksheet; AXLSDrawing: TXLSDrawing);
     end;

implementation

{ TXBookDrawing }

function TXBookDrawing.AddChart(AChart: TXLSDrawingChart): TXBookDrawingObjectChart;
begin
  Result := TXBookDrawingObjectChart.Create(FManager,Self,AChart);
  inherited Add(Result);
end;

procedure TXBookDrawing.AddComment(AComment: TXLSComment);
var
  i: integer;
  ii: integer;
  C: TXBookDrawingObjectComment;
  Pt1: TXYPoint;
// 0---------1
// |         |
// 2---------3
  Pts2: array[0..3] of TXYPoint;
  D: double;
  DMin: double;
begin
  C := TXBookDrawingObjectComment.Create(Self,AComment);
  CalcObject(C);

  Pt1.X := FColumns.AbsPos[AComment.Col] + FColumns.AbsVisibleSize[AComment.Col];
  Pt1.Y := FRows.AbsPos[AComment.Row];

  Pts2[0].X := C.X1;
  Pts2[0].Y := C.Y1;

  Pts2[1].X := C.X2;
  Pts2[1].Y := C.Y1;

  Pts2[2].X := C.X1;
  Pts2[2].Y := C.Y2;

  Pts2[3].X := C.X2;
  Pts2[3].Y := C.Y2;

  DMin := MAXINT;

  ii := 0;
  for i := 0 to 3 do begin
    D := Hypot(Pt1.X - Pts2[i].X,Pt1.Y - Pts2[i].Y);
    if D < DMin then begin
      DMin := D;
      ii := i;
    end;
  end;

  C.FArrowStart.X := Pts2[ii].X;
  C.FArrowStart.Y := Pts2[ii].Y;

  C.FArrowEnd.X := Pt1.X;
  C.FArrowEnd.Y := Pt1.Y;

  inherited Add(C);

  C.Paint;
end;

function TXBookDrawing.AddImage(AXLSImage: TXLSDrawingImage): TXBookDrawingObjectImage;
begin
  Result := TXBookDrawingObjectImage.Create(Self,AXLSImage);
  inherited Add(Result);
end;

function TXBookDrawing.AddTextBox(AXLSTextBox: TXLSDrawingTextBox): TXBookDrawingObjectTextBox;
begin
  Result := TXBookDrawingObjectTextBox.Create(FManager,Self,AXLSTextBox);
  inherited Add(Result);
end;

// TODO Picture are not drawn correct when partially outside and placed over hidden cols/rows.
procedure TXBookDrawing.CalcObject(AChild: TXBookDrawingObject);
var
  j: integer;
  C1,C1Offs,R1,R1Offs,C2,C2Offs,R2,R2Offs: integer;
  pX1,pY1,pX2,pY2: integer;
  V1,V2: integer;
  XCol: TXLSColumn;
  XRow: TXLSRow;
begin
//    FChilds[i].IsOnParentArea := not ((C2 < Columns.LeftCol) or (C1 > Columns.RightCol) or (R2 < Rows.TopRow) or (R1 > Rows.BottomRow));
  AChild.GetPos(C1,C1Offs,R1,R1Offs,C2,C2Offs,R2,R2Offs);
//  AChild.FIsOnScr := not ((C2 < FColumns.LeftCol) or (C1 > FColumns.RightCol) or (R2 < FRows.TopRow) or (R1 > FRows.BottomRow));
  AChild.FIsOnScr := not ((C2 < FColumns.LeftCol) or (C1 > FColumns.RightCol) or (R2 < FRows.TopRow) or (R1 > FRows.BottomRow));
  if AChild.FIsOnScr then begin
    C1Offs := FSkin.GDI.EMUToPixels(C1Offs);
    R1Offs := FSkin.GDI.EMUToPixels(R1Offs);
    C2Offs := FSkin.GDI.EMUToPixels(C2Offs);
    R2Offs := FSkin.GDI.EMUToPixels(R2Offs);
    AChild.Cursor := xctHandPoint;
    if (C1 >= FColumns.LeftCol) and (C2 <= FColumns.RightCol) and (R1 >= FRows.TopRow) and (R2 <= FRows.BottomRow) then begin
//      pX1 := FColumns.AbsPos[C1] + FColumns.AbsVisibleSize[C1] + C1Offs;
//      pY1 := FRows.AbsPos[R1]    + FRows.AbsVisibleSize[R1]    + R1Offs;
//      pX2 := FColumns.AbsPos[C2] + FColumns.AbsVisibleSize[C2] + C2Offs;
//      pY2 := FRows.AbsPos[R2]    + FRows.AbsVisibleSize[R2]    + R2Offs;

      pX1 := FColumns.AbsPos[C1] + C1Offs;
      pY1 := FRows.AbsPos[R1]    + R1Offs;
      pX2 := FColumns.AbsPos[C2] + C2Offs;
      pY2 := FRows.AbsPos[R2]    + R2Offs;
    end
    else begin
      if C1 < FColumns.LeftCol then begin
        V1 := 0;
        for j := FColumns.LeftCol - 1 downto C1 do begin
          XCol := FColumns.XLSColumns[j];
          if XCol <> Nil then
            V2 := FColumns.WidthToPixels(XCol.Width)
          else
            V2 := FColumns.WidthToPixels(FColumns.DefaultWidth);
          Inc(V1,V2);
        end;
//        pX1 := FColumns.Pos[0] + V2 + C1Offs - V1;
        pX1 := FColumns.Pos[0] + C1Offs - V1;
      end
      else
//        pX1 := FColumns.AbsPos[C1] + FColumns.AbsVisibleSize[C1] + C1Offs;
        pX1 := FColumns.AbsPos[C1] + C1Offs;
      if R1 < FRows.TopRow then begin
        V1 := 0;
        for j := FRows.TopRow - 1 downto R1 do begin
          XRow := FRows.XLSRows[j];
          if XRow <> Nil then
            V2 := FRows.HeightToPixels(XRow.Height)
          else
            V2 := FRows.HeightToPixels(FRows.DefaultHeight);
          Inc(V1,V2);
        end;
//        pY1 := FRows.Pos[0] + V2 + R1Offs - V1;
        pY1 := FRows.Pos[0] + R1Offs - V1;
      end
      else
//        pY1 := FRows.AbsPos[R1] + FRows.AbsVisibleSize[R1] + R1Offs;
        pY1 := FRows.AbsPos[R1] + R1Offs;
      if C2 > FColumns.RightCol then begin
        V1 := 0;
        for j := FColumns.RightCol + 1 to C2 do begin
          XCol := FColumns.XLSColumns[j];
          if XCol <> Nil then
            V2 := FColumns.WidthToPixels(XCol.Width)
          else
            V2 := FColumns.WidthToPixels(FColumns.DefaultWidth);
          Inc(V1,V2);
        end;
//        pX2 := FColumns.Pos[FColumns.Count - 1] + V2 + C2Offs + V1;
        pX2 := FColumns.Pos[FColumns.Count - 1] + C2Offs + V1;
      end
      else
//        pX2 := FColumns.AbsPos[C2] + FColumns.AbsVisibleSize[C2] + C2Offs;
        pX2 := FColumns.AbsPos[C2] + C2Offs;
      if R2 > FRows.BottomRow then begin
        V1 := 0;
        for j := FRows.BottomRow + 1 to R2 do begin
          XRow := FRows.XLSRows[j];
          if XRow <> Nil then
            V2 := FRows.HeightToPixels(XRow.Height)
          else
            V2 := FRows.HeightToPixels(FRows.DefaultHeight);
          Inc(V1,V2);
        end;
//        pY2 := FRows.Pos[FRows.Count - 1] + V2 + R2Offs + V1;
        pY2 := FRows.Pos[FRows.Count - 1] + R2Offs + V1;
      end
      else
//        pY2 := FRows.AbsPos[R2] + FRows.AbsVisibleSize[R2] + R2Offs;
        pY2 := FRows.AbsPos[R2] + R2Offs;
    end;
    Inc(pX1,FColumns.Offset);
    Inc(pY1,FRows.Offset);
    Inc(pX2,FColumns.Offset);
    Inc(pY2,FRows.Offset);

    AChild.SetSize(pX1,pY1,pX2,pY2);
  end;
end;

procedure TXBookDrawing.CalcObjects;
var
  i: integer;
begin
  for i := 0 to FChilds.Count - 1 do
    CalcObject(TXBookDrawingObject(FChilds[i]));
end;

procedure TXBookDrawing.Clear;
begin
  DeleteChilds;

  if FSkin <> Nil then
    FSkin.GDI.Clear;
end;

constructor TXBookDrawing.Create(AParent: TXSSClientWindow; AManager: TXc12Manager; AGDI: TAXWGDI; ASkinStyle: TXBookSkinStyle; AColumns: TXLSBookColumns; ARows: TXLSBookRows);
begin
  inherited Create(AParent);

  FManager := AManager;

  FSkin := TXLSBookSkin.Create(0,AGDI);
  FSkin.Style := ASkinStyle;
//  FSkin.GDI.PrintStatus := ParentSkin.GDI.PrintStatus;
//  FSkin.GDI.GetTextMetricSys;

  FColumns := AColumns;
  FRows := ARows;

  FEnabled := True;
end;

destructor TXBookDrawing.Destroy;
begin
  FreeAndNil(FSkin);
  inherited;
end;

procedure TXBookDrawing.LoadObjects(ASheet: TXc12DataWorksheet; AXLSDrawing: TXLSDrawing);
var
  i: integer;
begin
  Clear;

  FXLSDrawing := AXLSDrawing;

  for i := 0 to FXLSDrawing.Images.Count - 1 do
    AddImage(FXLSDrawing.Images[i]);

  for i := 0 to FXLSDrawing.TextBoxes.Count - 1 do
    AddTextBox(FXLSDrawing.TextBoxes[i]);

  for i := 0 to FXLSDrawing.Charts.Count - 1 do
    AddChart(FXLSDrawing.Charts[i]);
end;

function TXBookDrawing.ClientHit(const X, Y: integer): boolean;
begin
  Result := False;
end;

function TXBookDrawing.Hit(const X, Y: integer): boolean;
begin
  Result := False;
end;

procedure TXBookDrawing.Paint;
begin
  FSkin.GDI.Clear;
  FSkin.GDI.ReleaseHandle;
end;

procedure TXBookDrawing.RemoveComment(AComment: TXLSComment);
var
  i: integer;
begin
  for i := 0 to FChilds.Count - 1 do begin
    if (FChilds[i] is TXBookDrawingObjectComment) and (TXBookDrawingObjectComment(FChilds[i]).Comment = AComment) then begin
      FChilds.Delete(FChilds[i]);
      Exit;
    end;
  end;
end;

procedure TXBookDrawing.SelectObject(const AIndex: integer);
begin
  if AIndex >= 0 then
    FCurrObj := TXBookDrawingObject(FChilds[AIndex])
  else
    FCurrObj := Nil;
end;

procedure TXBookDrawing.SetSize(const pX1, pY1, pX2, pY2: integer);
begin
  inherited;

end;

{ TXBookDrawingObject }

constructor TXBookDrawingObject.Create(AParent: TXSSWindow);
begin
  inherited Create(AParent);
end;

{ TXBookDrawingObjectImage }

constructor TXBookDrawingObjectImage.Create(AParent: TXSSWindow; AXLSImage: TXLSDrawingImage);
begin
  inherited Create(AParent,AXLSImage.Anchor);

  FXLSImage := AXLSImage;
  FBitmap := FXLSImage.CreateBitmap;
end;

destructor TXBookDrawingObjectImage.Destroy;
begin
  if FBitmap <> Nil then
    FBitmap.Free;
  inherited;
end;

procedure TXBookDrawingObjectImage.Paint;
begin
  BeginPaintClipParent;

  if (FBitmap <> Nil) and not FBitmap.Empty then
    FSkin.GDI.PaintBMPImage(FBitmap,FX1,FY1,FX2,FY2)
  else begin
    FSkin.GDI.BrushColor := $A0A0FF;
    FSkin.GDI.Rectangle(FX1,FY1,FX2,FY2);
  end;

  EndPaint;
end;

{ TXBookDrawingObjectComment }

constructor TXBookDrawingObjectComment.Create(AParent: TXSSWindow; AComment: TXLSComment);
begin
  inherited Create(AParent);

  FComment := AComment;
end;

procedure TXBookDrawingObjectComment.GetPos(out AC1, AC1Offs, AR1, AR1Offs, AC2, AC2Offs, AR2, AR2Offs: integer);
begin
  AC1     := FComment.Xc12Comment.Col1;
  AC1Offs := FComment.Xc12Comment.Col1Offs;
  AR1     := FComment.Xc12Comment.Row1;
  AR1Offs := FComment.Xc12Comment.Row1Offs;
  AC2     := FComment.Xc12Comment.Col2;
  AC2Offs := FComment.Xc12Comment.Col2Offs;
  AR2     := FComment.Xc12Comment.Row2;
  AR2Offs := FComment.Xc12Comment.Row2Offs;
end;

procedure TXBookDrawingObjectComment.Paint;
begin
  inherited;
  BeginPaintClipParent;

  FSkin.GDI.PenColor := $000000;
  FSkin.GDI.BrushColor := FComment.Xc12Comment.Color;
  FSkin.GDI.Rectangle(FX1,FY1,FX2,FY2);

  RichTextRect(FSkin,FX1 + COMMENT_MARGIN,FY1 + COMMENT_MARGIN,FX2 - COMMENT_MARGIN,FY2 - COMMENT_MARGIN,@FComment.Xc12Comment.FontRuns[0],Length(FComment.Xc12Comment.FontRuns),FComment.Xc12Comment.Text);

  FSkin.GDI.Line(FArrowStart.X,FArrowStart.Y,FArrowEnd.X,FArrowEnd.Y);

  FSkin.GDI.BrushColor := $000000;
  FSkin.PaintArrow(FArrowEnd,FArrowStart,1);

  EndPaint;
end;

{ TXBookDrawingObjectAnchor }

constructor TXBookDrawingObjectAnchor.Create(AParent: TXSSWindow; AAnchor: TCT_TwoCellAnchor);
begin
  inherited Create(AParent);

  FAnchor := AAnchor;
end;

procedure TXBookDrawingObjectAnchor.GetPos(out AC1, AC1Offs, AR1, AR1Offs, AC2,AC2Offs, AR2, AR2Offs: integer);
begin
//  AC1     := FAnchor.From.Col - 1;
//  AC1Offs := FAnchor.From.ColOff;
//  AR1     := FAnchor.From.Row - 1;
//  AR1Offs := FAnchor.From.RowOff;
//  AC2     := FAnchor.To_.Col - 1;
//  AC2Offs := FAnchor.To_.ColOff;
//  AR2     := FAnchor.To_.Row - 1;
//  AR2Offs := FAnchor.To_.RowOff;

  AC1     := FAnchor.From.Col;
  AC1Offs := FAnchor.From.ColOff;
  AR1     := FAnchor.From.Row;
  AR1Offs := FAnchor.From.RowOff;
  AC2     := FAnchor.To_.Col;
  AC2Offs := FAnchor.To_.ColOff;
  AR2     := FAnchor.To_.Row;
  AR2Offs := FAnchor.To_.RowOff;
end;

{ TXBookDrawingObjectGraphic }

procedure TXBookDrawingObjectCommon.ApplySpPr(ASpPr: TCT_ShapeProperties);
begin
  FSkin.GDI.BrushColor := SpPrFillToRGB(ASpPr);
end;

constructor TXBookDrawingObjectCommon.Create(AParent: TXSSWindow; AAnchor: TCT_TwoCellAnchor);
begin
  inherited Create(AParent,AAnchor);
end;

{ TXBookDrawingObjectTextBox }

constructor TXBookDrawingObjectTextBox.Create(AManager: TXc12Manager; AParent: TXSSWindow; AXLSTextBox: TXLSDrawingTextBox);
var
  i,j    : integer;
  TxtBody: TCT_TextBody;
  Para   : TCT_TextParagraph;
  Run    : TCT_RegularTextRun;
  Font   : TXc12Font;
begin
  inherited Create(AParent,AXLSTextBox.Anchor);

  FXLSTextBox := AXLSTextBox;

  FText := TXLSFormattedText.Create;

  TxtBody := FXLSTextBox.Anchor.Objects.Sp.TxBody;

  if TxtBody <> Nil then begin
    for i := 0 to TxtBody.Paras.Count - 1 do begin
      Para := TxtBody.Paras[i];
      for j := 0 to Para.TextRuns.Count - 1 do begin
        Font := TXc12Font.Create(Nil);
        Font.Assign(AManager.StyleSheet.Fonts.DefaultFont);

        Run := Para.TextRuns[j].Run;
        if Run.RPr.Available then begin
          Font.Assign(Run.RPr);
          if Run.RPr.FillProperties.SolidFill <> Nil then
            Font.PColor.ARGB := RevRGB(SolidFillToRGB(Run.RPr.FillProperties.SolidFill));
        end;

        FText.Add(Run.T,Font);
      end;
    end;
  end;
end;

destructor TXBookDrawingObjectTextBox.Destroy;
var
  i: integer;
begin
  for i := 0 to FText.Count - 1 do
    FText[i].Font.Free;

  FText.Free;

  inherited;
end;

procedure TXBookDrawingObjectTextBox.Paint;
begin
  BeginPaintClipParent;

  FSkin.GDI.PaintColor := $FFFFF0;
  FSkin.GDI.Rectangle(FX1,FY1,FX2,FY2);

  FSkin.GDI.PenColor := $000000;
  ApplySpPr(FXLSTextbox.Anchor.Objects.Sp.SpPr);
  FSkin.GDI.Rectangle(FX1,FY1,FX2,FY2);

  RichTextRect(FSkin,FX1 + COMMENT_MARGIN,FY1 + COMMENT_MARGIN,FX2 - COMMENT_MARGIN,FY2 - COMMENT_MARGIN,FText);

  EndPaint;
end;

procedure TXBookDrawing.PaintPrint;
var
  i: integer;
begin
  FSkin.GDI.Clear;
  FSkin.GDI.ReleaseHandle;

  FSkin.GDI.BrushColor := $FF0000;

  if FVisible then begin
    for i := 0 to FChilds.Count - 1 do
      FChilds[i].Paint;
  end;
end;

{ TXBookDrawingObjectChart }

procedure TXBookDrawingObjectChart.Build;
begin
  FGeometry.Clear;

  FChartData.PtWidth := FSkin.GDI.PixToPtX(FX2 - FX1);
  FChartData.PtHeight := FSkin.GDI.PixToPtY(FY2 - FY1);

  FChartSpace.Build(FGeometry,FDrwData.ChartSpace);
end;

constructor TXBookDrawingObjectChart.Create(AManager: TXc12Manager; AParent: TXSSWindow; AChartData: TXLSDrawingChart);
begin
  inherited Create(AParent,AChartData.Anchor);

  FDrwData := AChartData;

  FChartData := TXSSChartData.Create;
  FChartData.Font := AManager.StyleSheet.Fonts[0];
  FChartSpace := TXSSChartItemChartSpace.Create(FSkin.GDI,FChartData);
  FGeometry := TXSSGeometryRect.Create(FSkin.GDI);
  FGeometry.SetSize(0,0,1,1);
end;

destructor TXBookDrawingObjectChart.Destroy;
begin
  FChartSpace.Free;
  FGeometry.Free;
  FChartData.Free;

  inherited;
end;

function TXBookDrawingObjectChart.MapX(AX: double): double;
begin
  Result := FX1 + (FX2 - FX1) * AX;
end;

function TXBookDrawingObjectChart.MapY(AY: double): double;
begin
  Result := FY1 + (FY2 - FY1) * AY;
end;

procedure TXBookDrawingObjectChart.Paint;
begin
  Build;

  BeginPaintClipParent;

//  FSkin.GDI.PaintColor := $FFFFF0;
//  FSkin.GDI.Rectangle(FX1,FY1,FX2,FY2);
//
//  FSkin.GDI.PenColor := $000000;
//
//  FSkin.GDI.Rectangle(FX1,FY1,FX2,FY2);

  PaintGeometry(FGeometry);

  EndPaint;
end;

procedure TXBookDrawingObjectChart.PaintGeometry(AGeom: TXSSGeometry);
var
  i,j     : integer;
  X,Y     : double;
  W,H     : double;
  Marker  : TXSSGeometryMarker;
  Rect    : TXSSGeometryRect;
  Circle  : TXSSGeometryCircle;
  Slice   : TXSSGeometryPizzaSlice;
  Line    : TXSSGeometryLine;
  PLine   : TXSSGeometryPolyline;
  PGon    : TXSSGeometryPolygon;
  Text    : TXSSGeometryText;
  TxtAlign: TDrawTextOptions;
  RichText: TXSSGeometryRichText;
  Pts     : array of TPoint;

procedure ApplyColors(AStyle: TXSSGeometryStyle);
begin
  FSkin.GDI.BrushColor := AStyle.Color;
  FSkin.GDI.PenColor := AStyle.LineColor;
  FSkin.GDI.PenWidth := FSkin.GDI.PtToPixX(AStyle.LineWidth);
end;

begin
  for i := 0 to AGeom.Count - 1 do begin
    case AGeom[i].Type_ of
      xgtMarker: begin
        Marker := TXSSGeometryMarker(AGeom[i]);
        W := FSkin.GDI.PtToPixX(Marker.Size);
        FSkin.GDI.PaintColor := Marker.Style.Color;
        FSkin.GDI.PenWidthF := FSkin.GDI.PtToPixX(1.5);

        case Marker.MStyle of
          xgmsNone    : ;
          xgmsCircle  : FSkin.GDI.CircleF(MapX(Marker.Pt.X),MapY(Marker.Pt.Y),W * 2);
          xgmsDash    : FSkin.GDI.LineF(MapX(Marker.Pt.X) - W,MapY(Marker.Pt.Y),MapX(Marker.Pt.X) + W,MapY(Marker.Pt.Y));
          xgmsDiamond : begin
            SetLength(Pts,5);

            Pts[0].X := Round(MapX(Marker.Pt.X) - W);
            Pts[0].Y := Round(MapY(Marker.Pt.Y));

            Pts[1].X := Round(MapX(Marker.Pt.X));
            Pts[1].Y := Round(MapY(Marker.Pt.Y) - W);

            Pts[2].X := Round(MapX(Marker.Pt.X) + W);
            Pts[2].Y := Round(MapY(Marker.Pt.Y));

            Pts[3].X := Round(MapX(Marker.Pt.X));
            Pts[3].Y := Round(MapY(Marker.Pt.Y) + W);

            Pts[4] := Pts[0];

            FSkin.GDI.Polygon(Pts);
          end;
          xgmsDot     : FSkin.GDI.CircleF(MapX(Marker.Pt.X),MapY(Marker.Pt.Y),W);
          xgmsPlus    : begin
            FSkin.GDI.LineF(MapX(Marker.Pt.X) - W,MapY(Marker.Pt.Y),MapX(Marker.Pt.X) + W,MapY(Marker.Pt.Y));
            FSkin.GDI.LineF(MapX(Marker.Pt.X),MapY(Marker.Pt.Y) - W,MapX(Marker.Pt.X),MapY(Marker.Pt.Y) + W);
          end;
          xgmsSquare  : FSkin.GDI.RectangleF(MapX(Marker.Pt.X) - W,MapY(Marker.Pt.Y) - W,MapX(Marker.Pt.X) + W,MapY(Marker.Pt.Y) + W);
          xgmsStar    : begin
            FSkin.GDI.LineF(MapX(Marker.Pt.X) - W,MapY(Marker.Pt.Y) - W,MapX(Marker.Pt.X) + W,MapY(Marker.Pt.Y) + W);
            FSkin.GDI.LineF(MapX(Marker.Pt.X) - W,MapY(Marker.Pt.Y) + W,MapX(Marker.Pt.X) + W,MapY(Marker.Pt.Y) - W);
            FSkin.GDI.LineF(MapX(Marker.Pt.X),MapY(Marker.Pt.Y) - W,MapX(Marker.Pt.X),MapY(Marker.Pt.Y) + W);
          end;
          xgmsTriangle: begin
            SetLength(Pts,4);

            Pts[0].X := Round(MapX(Marker.Pt.X) - W);
            Pts[0].Y := Round(MapY(Marker.Pt.Y) + W);

            Pts[1].X := Round(MapX(Marker.Pt.X));
            Pts[1].Y := Round(MapY(Marker.Pt.Y) - W);

            Pts[2].X := Round(MapX(Marker.Pt.X) + W);
            Pts[2].Y := Round(MapY(Marker.Pt.Y) + W);

            Pts[3] := Pts[0];

            FSkin.GDI.Polygon(Pts);
          end;
          xgmsX       : begin
            FSkin.GDI.LineF(MapX(Marker.Pt.X) - W,MapY(Marker.Pt.Y) - W,MapX(Marker.Pt.X) + W,MapY(Marker.Pt.Y) + W);
            FSkin.GDI.LineF(MapX(Marker.Pt.X) - W,MapY(Marker.Pt.Y) + W,MapX(Marker.Pt.X) + W,MapY(Marker.Pt.Y) - W);
          end;
        end;
      end;
      xgtLine: begin
        Line := TXSSGeometryLine(AGeom[i]);
        if Line.Style <> Nil then begin
          FSkin.GDI.PenColor := Line.Style.LineColor;
          FSkin.GDI.PenWidth := FSkin.GDI.PtToPixX(Line.Style.LineWidth);
        end;

        FSkin.GDI.LineF(MapX(Line.Pt1.X),MapY(Line.Pt1.Y),MapX(Line.Pt2.X),MapY(Line.Pt2.Y));
      end;
      xgtPolygon: begin
        PGon := TXSSGeometryPolyGon(AGeom[i]);
        if (PGon.Count > 0) and (PGon.Style <> Nil) then begin
          FSkin.GDI.PenColor := PGon.Style.LineColor;
          FSkin.GDI.BrushColor := PGon.Style.Color;
          FSkin.GDI.PenWidth := FSkin.GDI.PtToPixX(PGon.Style.LineWidth);

          SetLength(Pts,PGon.Count);
          for j := 0 to PGon.Count - 1 do begin
            Pts[j].X := Round(MapX(PGon.Pts[j].X));
            Pts[j].Y := Round(MapY(PGon.Pts[j].Y));
          end;

          FSkin.GDI.Polygon(Pts);
        end;
      end;
      xgtPolyline: begin
        PLine := TXSSGeometryPolyline(AGeom[i]);
        if (PLine.Count > 0) and (PLine.Style <> Nil) then begin
          FSkin.GDI.PenColor := PLine.Style.LineColor;
          FSkin.GDI.PenWidth := FSkin.GDI.PtToPixX(PLine.Style.LineWidth);

          FSkin.GDI.MoveToF(MapX(PLine.Pts[0].X),MapY(PLine.Pts[0].Y));
          for j := 1 to PLine.Count - 1 do
            FSkin.GDI.LineToF(MapX(PLine.Pts[j].X),MapY(PLine.Pts[j].Y));
        end;
      end;
      xgtRect: begin
        Rect := TXSSGeometryRect(AGeom[i]);

        ApplyColors(Rect.Style);

        if Rect.Style.Color <> XLS_COLOR_NONE then
          FSkin.GDI.RectangleF(MapX(Rect.Pt1.X),MapY(Rect.Pt1.Y),MapX(Rect.Pt2.X),MapY(Rect.Pt2.Y));
      end;
      xgtCircle: begin
        Circle  := TXSSGeometryCircle(AGeom[i]);

        ApplyColors(Circle.Style);

        W := Min(MapX(Circle.RefRect.Pt2.X) - MapX(Circle.RefRect.Pt1.X),MapY(Circle.RefRect.Pt2.Y) - MapY(Circle.RefRect.Pt1.Y));

        FSkin.GDI.CircleF(MapX(Circle.Pt.X),MapY(Circle.Pt.Y),W * Circle.MMScale);
      end;
      xgtPizzaSlice: begin
        Slice := TXSSGeometryPizzaSlice(AGeom[i]);

        ApplyColors(Slice.Style);

        W := MapX(Slice.Pt2.X) - MapX(Slice.Pt1.X);
        H := MapY(Slice.Pt2.Y) - MapY(Slice.Pt1.Y);

        if H > W then
          FSkin.GDI.PizzaSliceF(MapX(Slice.MidX),MapY(Slice.MidY),MapX(Slice.MMRadius) - FX1,Slice.StartAngle,Slice.EndAngle)
        else
          FSkin.GDI.PizzaSliceF(MapX(Slice.MidX),MapY(Slice.MidY),MapY(Slice.MMRadius) - FY1,Slice.StartAngle,Slice.EndAngle);
      end;
      xgtText: begin
        Text := TXSSGeometryText(AGeom[i]);

        FSkin.GDI.SetFont(Text.Font);

        TxtAlign := [];

        case Text.VertAlign of
          cvaTop   : TxtAlign := [dtoTop];
          cvaCenter: TxtAlign := [dtoVCenter];
          cvaBottom: TxtAlign := [dtoBottom];
        end;

        case Text.HorizAlign of
          chaGeneral,
          chaLeft   : TxtAlign := TxtAlign + [dtoLeft];
          chaCenter : TxtAlign := TxtAlign + [dtoHCenter];
          chaRight  : TxtAlign := TxtAlign + [dtoRight];
        end;

        FSkin.GDI.TextOutF(MapX(Text.Pt.X),MapY(Text.Pt.Y),Text.Text,TxtAlign);
      end;
      xgtRichText: begin
        FSkin.GDI.SetTextAlign(xhtaLeft,xvtaBaseline);

        RichText := TXSSGeometryRichText(AGeom[i]);

        X := MapX(RichText.Pt.X);
        Y := MapY(RichText.Pt.Y);

        case RichText.HorizAlign of
          chaCenter: RichTextRect(FSkin,X - 50,Y,X + 50,Y,RichText.Text,RichText.HorizAlign,RichText.VertAlign);
          else       RichTextRect(FSkin,X,Y,X + 5000,Y + 5000,RichText.Text,RichText.HorizAlign,RichText.VertAlign);
        end;

        FSkin.GDI.SetTextAlign(xhtaLeft,xvtaTop);
      end;
    end;

    if AGeom[i].HasChilds then
      PaintGeometry(AGeom[i]);
  end;
end;

end.
