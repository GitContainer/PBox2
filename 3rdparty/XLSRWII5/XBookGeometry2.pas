unit XBookGeometry2;

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

uses Classes, SysUtils, Contnrs, Math,
     Xc12Utils5, Xc12DataStyleSheet5,
     XLSUtils5, XLSTools5,
     XBookPaintGDI2;

type TXSSGeometryType = (xgtMarker,xgtRect,xgtCircle,xgtPizzaSlice,xgtLine,xgtPolyline,xgtPolygon,xgtText,xgtRichText);

type TXSSGeometryPoint = class(TObject)
protected
     FX: double;
     FY: double;
     FZ: double;
public
     procedure Assign(APoint: TXSSGeometryPoint);

     procedure Move(ADX,ADY: double);

     property X: double read FX write FX;
     property Y: double read FY write FY;
     property Z: double read FZ write FZ;
     end;

type TXSSGeometryStyle = class(TObject)
protected
     FColor    : longword;
     FLineColor: longword;
     FLineWidth: double;
public
     constructor Create;

     property Color    : longword read FColor write FColor;
     property LineColor: longword read FLineColor write FLineColor;
     property LineWidth: double read FLineWidth write FLineWidth;
     end;

type TXSSGeometryMarker = class;
     TXSSGeometryLine = class;
     TXSSGeometryPolyline = class;
     TXSSGeometryPolygon = class;
     TXSSGeometryRect = class;
     TXSSGeometryCircle = class;
     TXSSGeometryPizzaSlice = class;
     TXSSGeometryText = class;
     TXSSGeometryRichText = class;

     TXSSGeometry = class(TObject)
private
     function  GetItems(Index: integer): TXSSGeometry;
protected
     FGDI   : TAXWGDI;
     FChilds: TObjectList;
     FStyle : TXSSGeometryStyle;
public
     constructor Create(AGDI: TAXWGDI);
     destructor Destroy; override;

     procedure Clear;

     function Type_: TXSSGeometryType; virtual; abstract;

     function  Count: integer;
     function  HasChilds: boolean;

     function  Add(AType: TXSSGeometryType): TXSSGeometry;

     function  AddMarker: TXSSGeometryMarker;
     function  AddRect: TXSSGeometryRect;
     function  AddLine: TXSSGeometryLine;
     function  AddPolyline: TXSSGeometryPolyline;
     function  AddPolygon: TXSSGeometryPolygon;
     function  AddCircle: TXSSGeometryCircle;
     function  AddPizzaSlice: TXSSGeometryPizzaSlice;
     function  AddText: TXSSGeometryText; overload;
     function  AddText(AFont: TXc12Font; AVertAlign: TXc12VertAlignment; AHorizAlign: TXc12HorizAlignment; AX,AY: double; AText: AxUCSTring): TXSSGeometryText; overload;
     function  AddRichText: TXSSGeometryRichText;

     function  AsRect: TXSSGeometryRect; 

     procedure Move(ADX,ADY: double); virtual;

     property Style : TXSSGeometryStyle read FStyle;
     property Items[Index: integer]: TXSSGeometry read GetItems; default;
     end;

     TXSSGeometryMarkerStyle = (xgmsNone,xgmsCircle,xgmsDash,xgmsDiamond,xgmsDot,xgmsPlus,xgmsSquare,xgmsStar,xgmsTriangle,xgmsX);

     TXSSGeometryMarker = class(TXSSGeometry)
protected
     FPt    : TXSSGeometryPoint;
     FMStyle: TXSSGeometryMarkerStyle;
     FSize  : double;
public
     constructor Create(AGDI: TAXWGDI);
     destructor Destroy; override;

     function Type_: TXSSGeometryType; override;
     procedure Move(ADX,ADY: double); override;

     property Pt   : TXSSGeometryPoint read FPt;
     property MStyle: TXSSGeometryMarkerStyle read FMStyle write FMStyle;
     property Size : double read FSize write FSize;
     end;

     TXSSGeometryLine = class(TXSSGeometry)
protected
     FPt1,FPt2: TXSSGeometryPoint;
public
     constructor Create(AGDI: TAXWGDI);
     destructor Destroy; override;

     function Type_: TXSSGeometryType; override;

     procedure Assign(ALine: TXSSGeometryLine);

     procedure SetSize(AX1,AY1,AX2,AY2: double);

     procedure Move(ADX,ADY: double); override;

     property Pt1   : TXSSGeometryPoint read FPt1;
     property Pt2   : TXSSGeometryPoint read FPt2;
     end;

     TXSSGeometryPolyline = class(TXSSGeometry)
private
     function  GetPts(Index: integer): TXSSGeometryPoint;
protected
     FPts: TObjectList;
public
     constructor Create(AGDI: TAXWGDI);
     destructor Destroy; override;

     function Type_: TXSSGeometryType; override;

     procedure Add(AX,AY: double);
     procedure Close;

     procedure Assign(APolyline: TXSSGeometryPolyline);

     procedure Move(ADX,ADY: double); override;

     function Count: integer;

     property Pts[Index: integer]: TXSSGeometryPoint read GetPts; default;
     end;

     TXSSGeometryPolygon = class(TXSSGeometryPolyline)
protected
public
     function Type_: TXSSGeometryType; override;
     end;

     TXSSGeometryRect = class(TXSSGeometryLine)
protected
public
     constructor Create(AGDI: TAXWGDI);
     destructor Destroy; override;

     function Type_: TXSSGeometryType; override;

     procedure Assign(ARect: TXSSGeometryRect);

     function  Width: double;
     function  Height: double;

     function  MidX: double;
     function  MidY: double;

     procedure CenterHoriz(AWidth: double);
     procedure CenterVert(AHeight: double);

     procedure Shrink(ALeftRight,ATopBottom: double);
     end;

     TXSSGeometryCircle = class(TXSSGeometry)
protected
     FPt     : TXSSGeometryPoint;
     FRefRect: TXSSGeometryRect;
     FMMScale: double;
public
     constructor Create(AGDI: TAXWGDI);
     destructor Destroy; override;

     function Type_: TXSSGeometryType; override;

     property Pt     : TXSSGeometryPoint read FPt;
     property RefRect: TXSSGeometryRect read FRefRect write FRefRect;
     property MMScale: double read FMMScale write FMMScale;
     end;

     TXSSGeometryPizzaSlice = class(TXSSGeometryRect)
protected
     FStartAngle: double;
     FEndAngle  : double;

     function  GetRadius: double;
public
     function Type_: TXSSGeometryType; override;

     property StartAngle: double read FStartAngle write FStartAngle;
     property EndAngle  : double read FEndAngle write FEndAngle;

     property MMRadius : double read GetRadius;
     end;

     TXSSGeometryText = class(TXSSGeometry)
protected
     FPt        : TXSSGeometryPoint;
     FText      : AxUCString;
     FFont      : TXc12Font;
     FHorizAlign: TXc12HorizAlignment;
     FVertAlign : TXc12VertAlignment;
public
     constructor Create(AGDI: TAXWGDI);
     destructor Destroy; override;

     function Type_: TXSSGeometryType; override;

     procedure Assign(AText: TXSSGeometryText);

     procedure Move(ADX,ADY: double); override;

     property Pt        : TXSSGeometryPoint read FPt;
     property Text      : AxUCString read FText write FText;
     property Font      : TXc12Font read FFont write FFont;
     property HorizAlign: TXc12HorizAlignment read FHorizAlign write FHorizAlign;
     property VertAlign : TXc12VertAlignment read FVertAlign write FVertAlign;
     end;

     TXSSGeometryRichText = class(TXSSGeometry)
protected
     FText      : TXLSFormattedText;
     FPt        : TXSSGeometryPoint;
     FHorizAlign: TXc12HorizAlignment;
     FVertAlign : TXc12VertAlignment;
public
     constructor Create(AGDI: TAXWGDI);
     destructor Destroy; override;

     function Type_: TXSSGeometryType; override;

     procedure Add(AText: AxUCString; AFont: TXc12Font);

     procedure Move(ADX,ADY: double); override;

     property Pt        : TXSSGeometryPoint read FPt;
     property Text      : TXLSFormattedText read FText;
     property HorizAlign: TXc12HorizAlignment read FHorizAlign write FHorizAlign;
     property VertAlign : TXc12VertAlignment read FVertAlign write FVertAlign;
     end;

implementation

{ TXSSGeometryPoint }

procedure TXSSGeometryPoint.Assign(APoint: TXSSGeometryPoint);
begin
  FX := APoint.FX;
  FY := APoint.FY;
  FZ := APoint.FZ;
end;

procedure TXSSGeometryPoint.Move(ADX, ADY: double);
begin
  FX := FX + ADX;
  FY := FY + ADY;
end;

{ TXSSGeometry }

function TXSSGeometry.Add(AType: TXSSGeometryType): TXSSGeometry;
begin
  if FChilds = Nil then
    FChilds := TObjectList.Create;

  case AType of
    xgtMarker    : Result := TXSSGeometryMarker.Create(FGDI);
    xgtRect      : Result := TXSSGeometryRect.Create(FGDI);
    xgtLine      : Result := TXSSGeometryLine.Create(FGDI);
    xgtPolyline  : Result := TXSSGeometryPolyline.Create(FGDI);
    xgtPolygon   : Result := TXSSGeometryPolygon.Create(FGDI);
    xgtCircle    : Result := TXSSGeometryCircle.Create(FGDI);
    xgtPizzaSlice: Result := TXSSGeometryPizzaSlice.Create(FGDI);
    xgtText      : Result := TXSSGeometryText.Create(FGDI);
    xgtRichText  : Result := TXSSGeometryRichText.Create(FGDI);
    else         Result := Nil;
  end;

  FChilds.Add(Result);
end;

function TXSSGeometry.AddCircle: TXSSGeometryCircle;
begin
  Result := TXSSGeometryCircle(Add(xgtCircle));
end;

function TXSSGeometry.AddLine: TXSSGeometryLine;
begin
  Result := TXSSGeometryLine(Add(xgtLine));
end;

function TXSSGeometry.AddMarker: TXSSGeometryMarker;
begin
  Result := TXSSGeometryMarker(Add(xgtMarker));
end;

function TXSSGeometry.AddPizzaSlice: TXSSGeometryPizzaSlice;
begin
  Result := TXSSGeometryPizzaSlice(Add(xgtPizzaSlice));
end;

function TXSSGeometry.AddPolygon: TXSSGeometryPolygon;
begin
  Result := TXSSGeometryPolygon(Add(xgtPolygon));
end;

function TXSSGeometry.AddPolyline: TXSSGeometryPolyline;
begin
  Result := TXSSGeometryPolyline(Add(xgtPolyline));
end;

function TXSSGeometry.AddRect: TXSSGeometryRect;
begin
  Result := TXSSGeometryRect(Add(xgtRect));
end;

function TXSSGeometry.AddRichText: TXSSGeometryRichText;
begin
  Result := TXSSGeometryRichText(Add(xgtRichText));
end;

function TXSSGeometry.AddText(AFont: TXc12Font; AVertAlign: TXc12VertAlignment; AHorizAlign: TXc12HorizAlignment; AX, AY: double; AText: AxUCSTring): TXSSGeometryText;
begin
  Result := AddText;
  Result.Font := AFont;
  Result.VertAlign := AVertAlign;
  Result.HorizAlign := AHorizAlign;
  Result.Pt.X := AX;
  Result.Pt.Y := AY;
  Result.Text := AText;
end;

function TXSSGeometry.AddText: TXSSGeometryText;
begin
  Result := TXSSGeometryText(Add(xgtText));
end;

function TXSSGeometry.AsRect: TXSSGeometryRect;
begin
  Result := TXSSGeometryRect(Self);
end;

procedure TXSSGeometry.Clear;
begin
  if FChilds <> Nil then
    FChilds.Clear;
end;

function TXSSGeometry.Count: integer;
begin
  if FChilds <> Nil then
    Result := FChilds.Count
  else
    Result := 0;
end;

constructor TXSSGeometry.Create(AGDI: TAXWGDI);
begin
  FGDI := AGDI;

  FStyle := TXSSGeometryStyle.Create;
end;

destructor TXSSGeometry.Destroy;
begin
  if FChilds <> Nil then
    FChilds.Free;

  FStyle.Free;

  inherited;
end;

function TXSSGeometry.GetItems(Index: integer): TXSSGeometry;
begin
  Result := TXSSGeometry(FChilds.Items[Index]);
end;

function TXSSGeometry.HasChilds: boolean;
begin
  Result := FChilds <> Nil;
end;

procedure TXSSGeometry.Move(ADX, ADY: double);
var
  i: integer;
begin
  if FChilds <> Nil then begin
    for i  := 0 to FChilds.Count - 1 do
      Items[i].Move(ADX,ADY);
  end;
end;

{ TXSSGeometryRect }

procedure TXSSGeometryRect.Assign(ARect: TXSSGeometryRect);
begin
  inherited Assign(TXSSGeometryLine(ARect));

  FPt1.Assign(ARect.FPt1);
  FPt2.Assign(ARect.FPt2);
end;

procedure TXSSGeometryRect.CenterHoriz(AWidth: double);
var
  dX: double;
begin
  dX := (AWidth / 2) - (Width / 2);

  Move(dX,0);
end;

procedure TXSSGeometryRect.CenterVert(AHeight: double);
var
  dY: double;
begin
  dY := (AHeight / 2) - (Height / 2);

  Move(0,dY);
end;

constructor TXSSGeometryRect.Create(AGDI: TAXWGDI);
begin
  inherited Create(AGDI);
end;

destructor TXSSGeometryRect.Destroy;
begin
  inherited;
end;

function TXSSGeometryRect.Height: double;
begin
  Result := FPt2.Y - FPt1.Y;
end;

function TXSSGeometryRect.MidX: double;
begin
  Result := FPt1.X + ((FPt2.X - FPt1.X) / 2);
end;

function TXSSGeometryRect.MidY: double;
begin
  Result := FPt1.Y + ((FPt2.Y - FPt1.Y) / 2);
end;

procedure TXSSGeometryRect.Shrink(ALeftRight, ATopBottom: double);
begin
  FPt1.X := FPt1.X + ALeftRight;
  FPt2.X := FPt2.X - ALeftRight;

  FPt1.Y := FPt1.Y + ATopBottom;
  FPt2.Y := FPt2.Y - ATopBottom;
end;

function TXSSGeometryRect.Type_: TXSSGeometryType;
begin
  Result := xgtRect;
end;

function TXSSGeometryRect.Width: double;
begin
  Result := FPt2.X - FPt1.X;
end;

{ TXSSGeometryLine }

procedure TXSSGeometryLine.Assign(ALine: TXSSGeometryLine);
begin
  FPt1.Assign(ALine.FPt1);
  FPt2.Assign(ALine.FPt2);
end;

constructor TXSSGeometryLine.Create(AGDI: TAXWGDI);
begin
  inherited Create(AGDI);

  FPt1 := TXSSGeometryPoint.Create;
  FPt2 := TXSSGeometryPoint.Create;
end;

destructor TXSSGeometryLine.Destroy;
begin
  FPt1.Free;
  FPt2.Free;

  inherited;
end;

procedure TXSSGeometryLine.Move(ADX, ADY: double);
begin
  inherited Move(ADX,ADY);

  FPt1.Move(ADX,ADY);
  FPt2.Move(ADX,ADY);
end;

procedure TXSSGeometryLine.SetSize(AX1, AY1, AX2, AY2: double);
begin
  FPt1.X := AX1;
  FPt1.Y := AY1;
  FPt2.X := AX2;
  FPt2.Y := AY2;
end;

function TXSSGeometryLine.Type_: TXSSGeometryType;
begin
  Result := xgtLine;
end;

{ TXSSGeometryText }

procedure TXSSGeometryText.Assign(AText: TXSSGeometryText);
begin
  FPt.Assign(AText.FPt);

  FText := AText.FText;
end;

constructor TXSSGeometryText.Create(AGDI: TAXWGDI);
begin
  inherited Create(AGDI);

  FPt := TXSSGeometryPoint.Create;

  FHorizAlign := chaLeft;
  FVertAlign := cvaTop;
end;

destructor TXSSGeometryText.Destroy;
begin
  FPt.Free;

  inherited;
end;

procedure TXSSGeometryText.Move(ADX, ADY: double);
begin
  inherited Move(ADX,ADY);

  FPt.Move(ADX,ADY);
end;

function TXSSGeometryText.Type_: TXSSGeometryType;
begin
  Result := xgtText;
end;

{ TXSSGeometryRichText }

procedure TXSSGeometryRichText.Add(AText: AxUCString; AFont: TXc12Font);
begin
  FText.Add(AText,AFont);
end;

constructor TXSSGeometryRichText.Create(AGDI: TAXWGDI);
begin
  inherited Create(AGDI);

  FPt := TXSSGeometryPoint.Create;
  FText := TXLSFormattedText.Create;

  FHorizAlign := chaLeft;
  FVertAlign := cvaTop;
end;

destructor TXSSGeometryRichText.Destroy;
begin
  FText.Free;
  FPt.Free;

  inherited;
end;

procedure TXSSGeometryRichText.Move(ADX, ADY: double);
begin
  inherited Move(ADX,ADY);

  FPt.Move(ADX,ADY);
end;

function TXSSGeometryRichText.Type_: TXSSGeometryType;
begin
  Result := xgtRichText;
end;

{ TXSSGeometryPolyine }

procedure TXSSGeometryPolyline.Add(AX, AY: double);
var
  Pt: TXSSGeometryPoint;
begin
  Pt := TXSSGeometryPoint.Create;

  Pt.X := AX;
  Pt.Y := AY;

  FPts.Add(Pt);
end;

procedure TXSSGeometryPolyline.Assign(APolyline: TXSSGeometryPolyline);
var
  i: integer;
begin
  FPts.Clear;

  for i := 0 to APolyline.Count - 1 do
    Add(APolyline[i].X,APolyline[i].Y);
end;

procedure TXSSGeometryPolyline.Close;
begin
  if Count > 0 then
    Add(Pts[0].X,pts[0].Y);
end;

function TXSSGeometryPolyline.Count: integer;
begin
  Result := FPts.Count;
end;

constructor TXSSGeometryPolyline.Create(AGDI: TAXWGDI);
begin
  inherited Create(AGDI);

  FPts := TObjectList.Create;
end;

destructor TXSSGeometryPolyline.Destroy;
begin
  FPts.Free;

  inherited;
end;

function TXSSGeometryPolyline.GetPts(Index: integer): TXSSGeometryPoint;
begin
  Result := TXSSGeometryPoint(FPts[Index]);
end;

procedure TXSSGeometryPolyline.Move(ADX, ADY: double);
var
  i: integer;
begin
  inherited Move(ADX,ADY);

  for i := 0 to FPts.Count - 1 do
    TXSSGeometryPoint(FPts[i]).Move(ADX,ADY);
end;

function TXSSGeometryPolyline.Type_: TXSSGeometryType;
begin
  Result := xgtPolyline;
end;

{ TXSSGeometryMarker }

constructor TXSSGeometryMarker.Create(AGDI: TAXWGDI);
begin
  inherited Create(AGDI);

  FPt := TXSSGeometryPoint.Create;
end;

destructor TXSSGeometryMarker.Destroy;
begin
  FPt.Free;

  inherited;
end;

procedure TXSSGeometryMarker.Move(ADX, ADY: double);
begin
  inherited Move(ADX,ADY);

  FPt.Move(ADX,ADY);
end;

function TXSSGeometryMarker.Type_: TXSSGeometryType;
begin
  Result := xgtMarker;
end;

{ TXSSGeometryCircle }

constructor TXSSGeometryCircle.Create(AGDI: TAXWGDI);
begin
  inherited Create(AGDI);

  FMMScale := 1.0;

  FPt := TXSSGeometryPoint.Create;
end;

destructor TXSSGeometryCircle.Destroy;
begin
  FPt.Free;

  inherited;
end;

function TXSSGeometryCircle.Type_: TXSSGeometryType;
begin
  Result := xgtCircle;
end;

{ TXSSGeometryPizzaSlice }

function TXSSGeometryPizzaSlice.GetRadius: double;
begin
  Result := Min(Width,Height) / 2;
end;

function TXSSGeometryPizzaSlice.Type_: TXSSGeometryType;
begin
  Result := xgtPizzaSlice;
end;

{ TXSSGeometryStyle }

constructor TXSSGeometryStyle.Create;
begin
  FColor := $FFFFFF;
  FLineColor := $000000;
  FLineWidth := 0.75;
end;

{ TXSSGeometryPolygon }

function TXSSGeometryPolygon.Type_: TXSSGeometryType;
begin
  Result := xgtPolygon;
end;

end.
