unit FormFormatChartItem;

{$I AxCompilers.inc}

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics,
{$ifdef DELPHI_2006_OR_LATER}
  PNGImage,
{$endif}
  Controls, Forms, Dialogs, xpgParseDrawingCommon, StdCtrls, JPEG,
  ExtCtrls, ComCtrls, XLSDrawing5, ExtDlgs, Xc12Utils5, FrmColor, XLSUtils5;


type
  TfrmFormatChartItem = class(TForm)
    Button2: TButton;
    rgFill: TRadioGroup;
    rgLine: TRadioGroup;
    pcFill: TPageControl;
    tsSolid: TTabSheet;
    tsGradient: TTabSheet;
    tsPicture: TTabSheet;
    btnGradAdd: TButton;
    btnGradRemove: TButton;
    cbGradStops: TComboBox;
    btnGradColor: TButton;
    pbFillGrad: TPaintBox;
    Label1: TLabel;
    Label2: TLabel;
    edFillFilename: TEdit;
    Button4: TButton;
    dlgOpenPict: TOpenPictureDialog;
    cbFillPictTile: TCheckBox;
    btnFillColor: TButton;
    tbGradTransp: TTrackBar;
    lblGradTransp: TLabel;
    tbPictTransp: TTrackBar;
    lblPictTransp: TLabel;
    tbSolidTransp: TTrackBar;
    lblSolidTransp: TLabel;
    pbFillPict: TPaintBox;
    pbFillSolid: TPaintBox;
    pcLine: TPageControl;
    tsLineSolid: TTabSheet;
    Label3: TLabel;
    Label4: TLabel;
    udLineWidth: TUpDown;
    edLineWidth: TEdit;
    Button3: TButton;
    pbLineSolid: TPaintBox;
    procedure pbFillGradPaint(Sender: TObject);
    procedure rgFillClick(Sender: TObject);
    procedure btnGradAddClick(Sender: TObject);
    procedure btnGradRemoveClick(Sender: TObject);
    procedure btnGradColorClick(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure cbFillPictTileClick(Sender: TObject);
    procedure btnFillColorClick(Sender: TObject);
    procedure tbGradTranspChange(Sender: TObject);
    procedure tbPictTranspChange(Sender: TObject);
    procedure tbSolidTranspChange(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure pbFillPictPaint(Sender: TObject);
    procedure pbFillSolidPaint(Sender: TObject);
    procedure cbFillGradHorizClick(Sender: TObject);
    procedure rgLineClick(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure pbLineSolidPaint(Sender: TObject);
    procedure edLineWidthChange(Sender: TObject);
  private
    FShapeProps: TXLSDrwShapeProperies;
    FBitmap   : TBitmap;

    procedure LoadPicture(AFilename: AxUCString);
    procedure PaintPicture;
    procedure SetupGradient;
    procedure GradientRect(x1, y1, x2, y2: integer; StartColor, EndColor: longword; AAlpha: byte; HorizontalFill: boolean);
  public
    procedure Execute(AShapeProps: TXLSDrwShapeProperies);
  end;

implementation

{$R *.dfm}

{ TfrmFormatChartItem }

procedure TfrmFormatChartItem.btnGradAddClick(Sender: TObject);
begin
  FShapeProps.Fill.AsGradient.Stops.Add;

  SetupGradient;

  cbGradStops.ItemIndex := cbGradStops.Items.Count - 1;
end;

procedure TfrmFormatChartItem.btnGradRemoveClick(Sender: TObject);
begin
  FShapeProps.Fill.AsGradient.Stops.Remove(cbGradStops.Items.Count - 1);

  SetupGradient;

  cbGradStops.ItemIndex := cbGradStops.Items.Count - 1;
end;

procedure TfrmFormatChartItem.btnGradColorClick(Sender: TObject);
var
  GS: TXLSDrwFillGradientStop;
  Cl: TXc12Color;
begin
  GS := FShapeProps.Fill.AsGradient.Stops[cbGradStops.ItemIndex];

  Cl := RGBColorToXc12(GS.Color.AsRGB);
  if TfrmSelectColor.Create(Self).Execute(Cl) then begin
    GS.Color.AsRGB := Xc12ColorToRGB(Cl);
    pbFillGrad.Repaint;
  end;
end;

procedure TfrmFormatChartItem.Button3Click(Sender: TObject);
var
  Cl: TXc12Color;
begin
  Cl := RGBColorToXc12(FShapeProps.Line.Fill.AsSolid.Color.AsRGB);
  if TfrmSelectColor.Create(Self).Execute(Cl) then begin
    FShapeProps.Line.Fill.AsSolid.Color.AsRGB := Xc12ColorToRGB(Cl);
    pbLineSolid.Repaint;
  end;
end;

procedure TfrmFormatChartItem.Button4Click(Sender: TObject);
begin
  dlgOpenPict.FileName := edFillFilename.Text;
  if dlgOpenPict.Execute then begin
    edFillFilename.Text := dlgOpenPict.FileName;
    LoadPicture(edFillFilename.Text);
    pbFillPict.Repaint;
  end;
end;

procedure TfrmFormatChartItem.btnFillColorClick(Sender: TObject);
var
  Cl: TXc12Color;
begin
  Cl := RGBColorToXc12(FShapeProps.Fill.AsSolid.Color.AsRGB);
  if TfrmSelectColor.Create(Self).Execute(Cl) then begin
    FShapeProps.Fill.AsSolid.Color.AsRGB := Xc12ColorToRGB(Cl);
    pbFillSolid.Repaint;
  end;
end;

procedure TfrmFormatChartItem.cbFillGradHorizClick(Sender: TObject);
begin
  FShapeProps.Fill.AsGradient.Direction
end;

procedure TfrmFormatChartItem.cbFillPictTileClick(Sender: TObject);
begin
  FShapeProps.Fill.AsPicture.TilePicture := cbFillPictTile.Checked;
end;

procedure TfrmFormatChartItem.edLineWidthChange(Sender: TObject);
begin
  FShapeProps.Line.Width := StrToIntDef(edLineWidth.Text,Round(FShapeProps.Line.Width));

  pbLineSolid.Repaint;
end;

procedure TfrmFormatChartItem.Execute(AShapeProps: TXLSDrwShapeProperies);
begin
  FShapeProps := AShapeProps;

  case FShapeProps.Fill.FillType of
    xdftNone    : rgFill.ItemIndex := 0;
    xdftSolid   : begin
      rgFill.ItemIndex := 1;

      tbSolidTransp.Position := Round(FShapeProps.Fill.AsSolid.Color.Transparency * 100);
      lblSolidTransp.Caption := 'Transparency: ' + IntToStr(tbSolidTransp.Position) + '%';
    end;
    xdftGradient: begin
      rgFill.ItemIndex := 2;

      SetupGradient;
    end;
    xdftPicture : begin
      rgFill.ItemIndex := 3;

      edFillFilename.Text := FShapeProps.Fill.AsPicture.Filename;
      LoadPicture(edFillFilename.Text);
      tbPictTransp.Position := Round(FShapeProps.Fill.AsPicture.Transparency * 100);
      lblPictTransp.Caption := 'Transparency: ' + IntToStr(tbPictTransp.Position) + '%';
    end;
    xdftAutomatic : rgFill.ItemIndex := 4;
  end;

  case FShapeProps.Line.Fill.FillType of
    xdftNone : rgLine.ItemIndex := 0;
    xdftSolid: begin
      rgLine.ItemIndex := 1;

      udLineWidth.Position := Round(FShapeProps.Line.Width);
    end;
    xdftAutomatic: rgLine.ItemIndex := 2;
  end;

  rgFillClick(rgFill);
  rgLineClick(rgLine);

  ShowModal;

  // Don't load the picture when the file is selected in the dialog box as
  // the picture is stored in the XLSX then. And this happends every time the
  // picture is changed. You may then end up with an XLSX filled with unused
  // pictures.
  // Use FShapeProps.Fill.AsPicture.LoadPicture(No args) to load the picture.
  if FShapeProps.Fill.FillType = xdftPicture then
    FShapeProps.Fill.AsPicture.Filename := edFillFilename.Text;
end;

procedure TfrmFormatChartItem.FormCreate(Sender: TObject);
begin
  FBitmap := TBitmap.Create;
end;

procedure TfrmFormatChartItem.FormDestroy(Sender: TObject);
begin
  FBitmap.Free;
end;

procedure TfrmFormatChartItem.GradientRect(x1, y1, x2, y2: integer; StartColor, EndColor: longword; AAlpha: byte; HorizontalFill: boolean);
var
  Vertex: array[0..1] of TRIVERTEX;
  GRect: GRADIENT_RECT;
begin
  Vertex[0].x := x1;
  Vertex[0].y := y1;
  Vertex[0].Red :=  (StartColor and $00FF0000) shr 8;
  Vertex[0].Green := StartColor and $0000FF00;
  Vertex[0].Blue := (StartColor and $000000FF) shl 8;
  Vertex[0].Alpha := AAlpha;
  Vertex[1].x := x2 + 1;
  Vertex[1].y := y2 + 1;
  Vertex[1].Red :=  (EndColor and $00FF0000) shr 8;
  Vertex[1].Green := EndColor and $0000FF00;
  Vertex[1].Blue := (EndColor and $000000FF) shl 8;
  Vertex[1].Alpha := AAlpha;

  GRect.UpperLeft := 0;
  GRect.LowerRight := 1;

  // Declaration of GradientFill in Delphi 7 (Windows.pas) is wrong.
{$ifdef DELPHI_2006_OR_LATER}
  if HorizontalFill then
    GradientFill(pbFillGrad.Canvas.Handle,@Vertex,2,@GRect,1,GRADIENT_FILL_RECT_H)
  else
    GradientFill(pbFillGrad.Canvas.Handle,@Vertex,2,@GRect,1,GRADIENT_FILL_RECT_V);
{$else}
  pbFillSolid.Canvas.Brush.Color := EndColor;
  pbFillSolid.Canvas.Rectangle(x1, y1, x2 + 1, y2 + 1);
{$endif}
end;

procedure TfrmFormatChartItem.LoadPicture(AFilename: AxUCString);
var
{$ifdef DELPHI_2006_OR_LATER}
  PNG: TPNGImage;
{$endif}
  JPG: TJPEGImage;
  ext: AxUCString;
begin
  FShapeProps.Fill.AsPicture.LoadPicture(AFilename);

  Ext := Lowercase(ExtractFileExt(AFilename));
  if Ext = '.bmp' then
    FBitmap.LoadFromFile(AFilename)
{$ifdef DELPHI_2006_OR_LATER}
  else if Ext = '.png' then begin
    PNG := TPNGImage.Create;
    try
      PNG.LoadFromFile(AFilename);
      FBitmap.Assign(PNG);
    finally
      PNG.Free;
    end;
  end
{$endif}  
  else if (Ext = '.jpg') or (Ext = '.jpeg') then begin
    JPG := TJPEGImage.Create;
    try
      JPG.LoadFromFile(AFilename);
      FBitmap.Assign(JPG);
    finally
      JPG.Free;
    end;
  end
  else
    ShowMessage('Can not read picture of type ' + Ext);
end;

procedure TfrmFormatChartItem.PaintPicture;
var
  Blendfunc: TBLENDFUNCTION;
  RW,RH: double;
  W,H: integer;
begin
//  pbFillPict.Canvas.Draw(0,0,FBitmap);
  if (FBitmap <> Nil) then begin
    Blendfunc.BlendOp := AC_SRC_OVER;
    Blendfunc.BlendFlags := 0;
    Blendfunc.SourceConstantAlpha := Round($FF * ((100 - tbPictTransp.Position) / 100));
    BlendFunc.AlphaFormat := AC_SRC_OVER;

    if (FBitmap.Width > pbFillPict.Width) or (FBitmap.Height > pbFillPict.Height) then begin
      RW := FBitmap.Width / pbFillPict.Width;
      RH := FBitmap.Height / pbFillPict.Height;

      if RW > RH then begin
        W := Round(FBitmap.Width / RW);
        H := Round(FBitmap.Height / RW);
      end
      else begin
        W := Round(FBitmap.Width / RH);
        H := Round(FBitmap.Height / RH);
      end;
    end
    else begin
      W := FBitmap.Width;
      H := FBitmap.Height;
    end;

{$ifdef DELPHI_2006_OR_LATER}
    Windows.AlphaBlend(pbFillPict.Canvas.Handle,0,0,W,H,FBitmap.Canvas.Handle,0,0,FBitmap.Width,FBitmap.Height,Blendfunc);
{$else}
    Windows.AlphaBlend(pbFillPict.Canvas.Handle,0,0,W,H,FBitmap.Canvas.Handle,0,0,FBitmap.Width,FBitmap.Height,Blendfunc);
{$endif}
  end;
end;

procedure TfrmFormatChartItem.pbFillGradPaint(Sender: TObject);
begin
  GradientRect(0,0,pbFillGrad.Width,pbFillGrad.Height,FShapeProps.Fill.AsGradient.Stops[0].Color.AsRGB,FShapeProps.Fill.AsGradient.Stops[1].Color.AsRGB,Round(tbGradTransp.Position * 2.55),False);
end;

procedure TfrmFormatChartItem.pbFillPictPaint(Sender: TObject);
begin
  PaintPicture;
end;

procedure TfrmFormatChartItem.pbFillSolidPaint(Sender: TObject);
var
  Blendfunc: TBLENDFUNCTION;
  BMP: TBitmap;
begin
  BMP := TBitmap.Create;
  try
{$ifdef DELPHI_2006_OR_LATER}
    BMP.SetSize(pbFillSolid.Width,pbFillSolid.Height);
{$else}
    BMP.Width := pbFillSolid.Width;
    BMP.Height := pbFillSolid.Height;
{$endif}
    BMP.Canvas.Brush.Color := FShapeProps.Fill.AsSolid.Color.AsTColor;
    BMP.Canvas.Rectangle(0,0,BMP.Width,BMP.Height);

    Blendfunc.BlendOp := AC_SRC_OVER;
    Blendfunc.BlendFlags := 0;
    Blendfunc.SourceConstantAlpha := Round($FF * ((100 - tbSolidTransp.Position) / 100));
    BlendFunc.AlphaFormat := AC_SRC_OVER;

{$ifdef DELPHI_2006_OR_LATER}
    Windows.AlphaBlend(pbFillSolid.Canvas.Handle,0,0,BMP.Width,BMP.Height,BMP.Canvas.Handle,0,0,BMP.Width,BMP.Height,Blendfunc);
{$else}
    Windows.AlphaBlend(pbFillSolid.Canvas.Handle,0,0,BMP.Width,BMP.Height,BMP.Canvas.Handle,0,0,BMP.Width,BMP.Height,Blendfunc);
{$endif}
  finally
    BMP.Free;
  end;
end;

procedure TfrmFormatChartItem.pbLineSolidPaint(Sender: TObject);
var
  Y: integer;
begin
  pbLineSolid.Canvas.Pen.Color := FShapeProps.Line.Fill.AsSolid.Color.AsTColor;
  // Points -> Pixels. Only correct on 96 dpi monitors.
  pbLineSolid.Canvas.Pen.Width := Round(FShapeProps.Line.Width * 0.75);

  Y := pbLineSolid.Height div 2;
  pbLineSolid.Canvas.MoveTo(0,Y);
  pbLineSolid.Canvas.LineTo(pbLineSolid.Width,Y);
end;

procedure TfrmFormatChartItem.rgFillClick(Sender: TObject);
begin
  case rgFill.ItemIndex of
    0: FShapeProps.Fill.FillType := xdftNone;
    1: begin
      if FShapeProps.Fill.FillType <> xdftSolid then begin
        FShapeProps.Fill.FillType := xdftSolid;
        FShapeProps.Fill.AsSolid.Color.AsRGB := $00FFFFFF;
      end;

      pcFill.ActivePage := tsSolid;
    end;
    2: begin
      FShapeProps.Fill.FillType := xdftGradient;

      pcFill.ActivePage := tsGradient;

      SetupGradient;
    end;
    3: begin
      FShapeProps.Fill.FillType := xdftPicture;

      cbFillPictTile.Checked := FShapeProps.Fill.AsPicture.TilePicture;

      pcFill.ActivePage := tsPicture;
    end;
    4: FShapeProps.Fill.FillType := xdftAutomatic;
  end;

  pcFill.Visible := (rgFill.ItemIndex > 0) and (rgFill.ItemIndex < (rgFill.Items.Count - 1));

  pbFillGrad.Repaint;
end;

procedure TfrmFormatChartItem.rgLineClick(Sender: TObject);
begin
  case rgLine.ItemIndex of
    0 : FShapeProps.Line.Fill.FillType := xdftNone;
    1 : begin
      if FShapeProps.Line.Fill.FillType <> xdftSolid then begin
        FShapeProps.Line.Fill.FillType := xdftSolid;
        FShapeProps.Line.Fill.AsSolid.Color.AsRGB := $00000000;

        FShapeProps.Line.Width := 1;
      end;

      udLineWidth.Position := Round(FShapeProps.Line.Width);
    end;
    2 : FShapeProps.Line.Fill.FillType := xdftAutomatic;
  end;

  pcLine.Visible := FShapeProps.Line.Fill.FillType = xdftSolid;
end;

procedure TfrmFormatChartItem.SetupGradient;
var
  i: integer;
begin
  cbGradStops.Clear;

  for i := 1 to FShapeProps.Fill.AsGradient.Stops.Count do
    cbGradStops.Items.Add('Stop ' + IntToStr(i));

  cbGradStops.ItemIndex := 0;

//  tbGradPos.Position := Round(FShapeProps.Fill.AsGradient.Stops[0].StopPosition * 100);
//  lblGradPos.Caption := 'Position: ' + IntToStr(tbGradPos.Position) + '%';

  tbGradTransp.Position := Round(FShapeProps.Fill.AsGradient.Stops[0].Transparency * 100);
  lblGradTransp.Caption := 'Transparency: ' + IntToStr(tbGradTransp.Position) + '%';

  btnGradRemove.Enabled := FShapeProps.Fill.AsGradient.Stops.Count > 2;
end;

procedure TfrmFormatChartItem.tbGradTranspChange(Sender: TObject);
var
  i: integer;
begin
  for i := 0 to FShapeProps.Fill.AsGradient.Stops.Count - 1 do
    FShapeProps.Fill.AsGradient.Stops[i].Transparency := tbGradTransp.Position / 100;

  lblGradTransp.Caption := 'Transparency: ' + IntToStr(tbGradTransp.Position) + '%';

  pbFillGrad.Repaint;
end;

procedure TfrmFormatChartItem.tbPictTranspChange(Sender: TObject);
begin
  FShapeProps.Fill.AsPicture.Transparency := tbPictTransp.Position / 100;

  lblPictTransp.Caption := 'Transparency: ' + IntToStr(tbPictTransp.Position) + '%';

  pbFillPict.Repaint;
end;

procedure TfrmFormatChartItem.tbSolidTranspChange(Sender: TObject);
begin
  FShapeProps.Fill.AsSolid.Color.Transparency := tbSolidTransp.Position / 100;

  lblSolidTransp.Caption := 'Transparency: ' + IntToStr(tbSolidTransp.Position) + '%';

  pbFillSolid.Repaint;
end;

end.
