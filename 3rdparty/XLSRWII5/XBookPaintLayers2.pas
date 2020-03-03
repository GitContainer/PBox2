unit XBookPaintLayers2;

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
     Windows,
     XBook_System_2, XBookPaintGDI2,
     XLSUtils5;

type TXPaintLayerMode = (plmNone,plmDestination,plmSingle,plmMulti,plmMetafile);

type TXPaintLayers = class(TObject)
private
     FDestLayer: TAXWGDI;
     FSheetLayer: TAXWGDI;
     FDrawingLayer: TAXWGDI;

     FLayerMode: TXPaintLayerMode;
     FSystem: TXSSSystem;

     procedure SetLayerMode(const Value: TXPaintLayerMode);
     function  GetMetafileLayer: TAXWGDIMetafile;
public
     constructor Create(ASystem: TXSSSystem);
     destructor Destroy; override;
     procedure ReleaseHandle;

     property LayerMode: TXPaintLayerMode read FLayerMode write SetLayerMode;
     property DestLayer: TAXWGDI read FDestLayer;
     property SheetLayer: TAXWGDI read FSheetLayer;
     property DrawingLayer: TAXWGDI read FDrawingLayer;
     property MetafileLayer: TAXWGDIMetaFile read GetMetafileLayer;
     end;

implementation

{ TXPaintPaintLayers }

constructor TXPaintLayers.Create(ASystem: TXSSSystem);
begin
  FSystem := ASystem;
//  FGDI3D := TXBookGDI3D.Create;
end;

destructor TXPaintLayers.Destroy;
begin
  case FLayerMode of
    plmMetafile,
    plmDestination: begin
      FDestLayer.Free;
    end;
    plmSingle: raise XLSRWException.Create('Unsupported layer mode (Single)');
    plmMulti: begin
      FDestLayer.Free;
      FSheetLayer.Free;
      FDrawingLayer.Free;
    end;
  end;
  inherited;
end;

function TXPaintLayers.GetMetafileLayer: TAXWGDIMetafile;
begin
  Result := TAXWGDIMetafile(FDestLayer);
end;

procedure TXPaintLayers.ReleaseHandle;
var
  Dirty: boolean;
begin
  Dirty := FSheetLayer.Dirty;
  if Dirty then
    FSystem.HideCaret;

  if FLayerMode = plmMulti then begin
    FSheetLayer.RenderDirty(FDestLayer);
    FDrawingLayer.RenderDirty(FDestLayer);
    // DrawingLayer.ReleaseHandle is called when the drawing is repainted in XBookDrawing2.
    FSheetLayer.ReleaseHandle;
//    FDrawingLayer.ReleaseHandle;
  end;

  if Dirty then
    FSystem.ShowCaret;
end;

procedure TXPaintLayers.SetLayerMode(const Value: TXPaintLayerMode);
begin
  if FLayerMode <> plmNone then
    raise XLSRWException.Create('Layers are allready initiated');
  FLayerMode := Value;
  case FLayerMode of
    plmNone       : raise XLSRWException.Create('Layers can not be initiated to None');
    plmDestination: FDestLayer := TAXWGDI.Create(FSystem.Handle);
    plmSingle     : raise XLSRWException.Create('Unsupported layer mode (Single)');
    plmMulti      : begin
      FDestLayer := TAXWGDI.Create(FSystem.Handle);
      FSheetLayer := TAXWGDIBMP.Create;
      FDrawingLayer := TAXWGDIBMPTrans.Create;
    end;
    plmMetafile  : begin
      FDestLayer := TAXWGDIMetafile.Create;
      FDrawingLayer := FDestLayer;
    end;
  end;
end;

end.
