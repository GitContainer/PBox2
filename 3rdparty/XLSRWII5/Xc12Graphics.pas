unit Xc12Graphics;

{-
********************************************************************************
******* XLSReadWriteII V6.00                                             *******
*******                                                                  *******
******* Copyright(C) 1999,2017 Lars Arvidsson, Axolot Data               *******
*******                                                                  *******
******* email: components@axolot.com                                     *******
******* URL:   http://www.axolot.com                                     *******
********************************************************************************
** Users of the XLSReadWriteII component must accept the following            **
** disclaimer of warranty:                                                    **
**                                                                            **
** XLSReadWriteII is supplied as is. The author disclaims all warranties,     **
** expressedor implied, including, without limitation, the warranties of      **
** merchantability and of fitness for any purpose. The author assumes no      **
** liability for damages, direct or consequential, which may result from the  **
** use of XLSReadWriteII.                                                     **
********************************************************************************
}

{$B-}
{$H+}
{$R-}
{$I AxCompilers.inc}
{$I XLSRWII.inc}

interface

uses Classes, SysUtils, Contnrs,
{$ifdef BABOON}
  {$ifdef DELPHI_XE5_OR_LATER}
       FMX.Graphics,
  {$else}
       FMX.Types,
  {$endif}
{$else}
     vcl.Graphics,
{$endif}
{$ifdef XLS_HAS_JPG_SUPPORT}
     Vcl.Imaging.jpeg,
{$endif}
{$ifdef XLS_HAS_PNG_SUPPORT}
     Vcl.Imaging.PNGImage,
{$endif}
     Xc12Utils5, XLSUtils5, XLSReadWriteOPC5, XLSRelCells5;

type  TXc12ImageType = (x12itJPEG,x12itPNG,x12itEMF,x12itWMF,x12itGIF,x12itUnknown);
const Xc12ImageTypeStr: array[TXc12ImageType] of AxUCString = ('jpeg','png','emf','wmf','GIF','???');

type TXc12GraphicImage = class(TIndexObject)
protected
     FDrawingRId: AxUCString;
     FRId       : AxUCString;
     FRefCount  : integer;
     FImageType : TXc12ImageType;
     FImage     : Pointer;
     FImageSize : integer;
     FFileId    : AxUCString;
     FDescr     : AxUCString;
     FWritten   : boolean;
     FWidth     : integer;
     FHeight    : integer;
public
     constructor Create;
     destructor Destroy; override;

     function  CreateBitmap: TBitmap;

     procedure IncRefCount;
     procedure DecRefCount;

     function  TypeExt: AxUCString;
     // Used by (html) export.
     function  UniqueName: AxUCString;

     property DrawingRId: AxUCString read FDrawingRId write FDrawingRId;
     property RId: AxUCString read FRId write FRId;
     property ImageType: TXc12ImageType read FImageType;
     property Image: Pointer read FImage;
     property ImageSize: integer read FImageSize;
     property FileId: AxUCString read FFileId;
     property Descr: AxUCString read FDescr write FDescr;
     property Width: integer read FWidth write FWidth;
     property Heigh: integer read FHeight write FHeight;
     // For temporary use, as when html export is written.
     property Written: boolean read FWritten write FWritten;
     end;

     TXc12GraphicImages = class(TIndexObjectList)
private
     function  GetItems(Index: integer): TXc12GraphicImage;
protected
     FErrors: TXLSErrorManager;

     function  GetImageSize(AImage: TStream; AType: TXc12ImageType; out AWidth,AHeight: integer): boolean;
public
     constructor Create(AErrors: TXLSErrorManager);

     function  FindByRId(const ADrawingId,AId: AxUCString): TXc12GraphicImage;
     function  FindByFileId(const AId: AxUCString): TXc12GraphicImage;
     function  FindByDescr(const ADescr: AxUCString): TXc12GraphicImage;
     function  TypeFromExt(AExt: AxUCString): TXc12ImageType;
     function  ExtFromType(AType: TXc12ImageType): AxUCString;

     // TODO Storing common images in a common list don't works correct.
     // Solve the problem by creating a list with only the imgaes where
     // the images are refrereerd by OPC Target.

     // Used when adding images from file.
     procedure Xc12_Add(AStream: TStream; AImageExt,AFileId,ADrawingRId,ARId: AxUCString);
     // Used when adding images from code.
     function  Add(AStream: TStream; AImageType: TXc12ImageType): TXc12GraphicImage; overload;

     function  LoadFromFile(const AFilename: AxUCString): TXc12GraphicImage;
     function  LoadFromStream(AStream: TStream; const AFilename: AxUCString): TXc12GraphicImage;
     procedure Delete(AImage: TXc12GraphicImage);
     procedure SetWritten(const AValue: boolean);

     property Items[Index: integer]: TXc12GraphicImage read GetItems; default;
     end;

     TXc12GraphicManager = class(TObject)
protected
     FErrors: TXLSErrorManager;
     FImages: TXc12GraphicImages;

     FXLSOPC: TOPC_XLSX;
public
     constructor Create(AErrors: TXLSErrorManager);
     destructor Destroy; override;

     procedure Clear;

     procedure BeforeWrite;

     function SaveImageToOPC(AImage: TXc12GraphicImage): AxUCString;

     function CreateRelativeCells(ARef: AxUCString): TXLSRelCells; virtual;

     property Images: TXc12GraphicImages read FImages;
     property XLSOPC: TOPC_XLSX read FXLSOPC write FXLSOPC;
     end;

implementation

function ReadBigEndian(AStream: TStream): Word;
type
  TBigEndian = record
    case Byte of
      0: (Value: Word);
      1: (Byte1, Byte2: Byte);
  end;
var
  BE: TBigEndian;
begin
  AStream.Read(BE.Byte2, SizeOf(Byte));
  AStream.Read(BE.Byte1, SizeOf(Byte));
  Result := BE.Value;
end;

function GetJpegSize(jpeg: TStream; out Width, Height: integer): boolean;
var
  n: integer;
  b: byte;
  w: Word;
begin
  Result := False;
  n := jpeg.Size - 8;
  jpeg.Position := 0;
  if n <= 0 then
    Exit;
  jpeg.Read(w,2);
  if w <> $D8FF then
    Exit; // invalid format
  jpeg.Read(b,1);
  while (jpeg.Position < n) and (b = $FF) do begin
    jpeg.Read(b,1);
    case b of
      $C0..$C3: begin
        jpeg.Seek(3,soFromCurrent);
        jpeg.Read(w,2);
        Height := Swap(w);
        jpeg.Read(w,2);
        Width := Swap(w);
        jpeg.Read(b,1);
//        BitDepth := b * 8;
        Result := True; // JPEG format OK
        Exit;
      end;
      $FF: jpeg.Read(b,1);
      $D0..$D9, $01: begin
        jpeg.Seek(1,soFromCurrent);
        jpeg.Read(b,1);
      end;
      else begin
        jpeg.Read(w,2);
        jpeg.Seek(swap(w)-2, soFromCurrent);
        jpeg.Read(b,1);
      end;
    end;
  end;
end;

function GetPNGSize(png: TStream; var wWidth, wHeight: integer): boolean;
type
  TPNGSig = array[0..7] of Byte;
const
  ValidSig: TPNGSig = (137,80,78,71,13,10,26,10);
var
  Sig: TPNGSig;
  x: integer;
begin
  Result := False;
  FillChar(Sig, SizeOf(Sig), #0);

  png.Read(Sig[0], SizeOf(Sig));
  for x := Low(Sig) to High(Sig) do begin
    if Sig[x] <> ValidSig[x] then
      Exit;
  end;
  png.Seek(18, 0);
  wWidth := ReadBigEndian(png);
  png.Seek(22, 0);
  wHeight := ReadBigEndian(png);
  Result := True;
end;

function GetMetafileSize(AMetafile: TStream; var wWidth, wHeight: integer): boolean;
{$ifdef XLS_HAS_METAFILE_SUPPORT}
var
  WMF: TMetafile;
{$endif}
begin
{$ifdef XLS_HAS_METAFILE_SUPPORT}
  WMF := TMetafile.Create;
  try
    try
      AMetafile.Seek(0,soFromBeginning);

      WMF.LoadFromStream(AMetafile);
    except
      Result := False;
      Exit;
    end;
    wWidth := WMF.Width;
    wHeight := WMF.Height;
    Result := True;
  finally
    WMF.Free;
  end;
{$else}
  Result := False;
{$endif}
end;

{$ifdef XLS_HAS_PNG_SUPPORT}
procedure BmpToPng(AStreamBMP,AStreamPNG: TStream);
var
  BMP: TBitmap;
  PNG: TPNGImage;
begin
  BMP := TBitmap.Create;
  try
    BMP.LoadFromStream(AStreamBMP);
    PNG := TPNGImage.Create;
    try
      PNG.Assign(BMP);
      PNG.SaveToStream(AStreamPNG);
      AStreamPNG.Seek(0,soFromBeginning);
    finally
      PNG.Free;
    end;
  finally
    BMP.Free;
  end;
end;
{$endif}

{ TXc12DrawingImages }

procedure TXc12GraphicImages.Xc12_Add(AStream: TStream; AImageExt,AFileId,ADrawingRId,ARId: AxUCString);
var
  Image: TXc12GraphicImage;
begin
//  Image := FindByFileId(AFileId);
//  if Image <> Nil then begin
//    Image.IncRefCount;
//  end
//  else begin
    Image := TXc12GraphicImage.Create;

    Image.FImageSize := AStream.Size;
    GetMem(Image.FImage,Image.FImageSize);
    AStream.Read(Image.FImage^,Image.FImageSize);

    Image.FDrawingRId := ADrawingRId;
    Image.FRId := ARId;
    Image.FFileId := AFileId;
    Image.FImageType := TypeFromExt(AImageExt);
    Image.FRefCount := 1;

    inherited Add(Image);
//  end;
end;

function TXc12GraphicImages.Add(AStream: TStream; AImageType: TXc12ImageType): TXc12GraphicImage;
var
  i: integer;
  S: AxUCString;
begin
  i := Count;
  repeat
    Inc(i);
    S := Format('Image%d.%s',[i,ExtFromType(AImageType)]);
    Result := FindByFileId(S);
  until Result = Nil;

  Result := TXc12GraphicImage.Create;

  Result.FImageSize := AStream.Size;
  GetMem(Result.FImage,Result.FImageSize);
  AStream.Seek(0,soFromBeginning);
  AStream.Read(Result.FImage^,Result.FImageSize);

  Result.FFileId := S;
  Result.FImageType := AImageType;
  Result.FRefCount := 1;

  inherited Add(Result);
end;

procedure TXc12GraphicImages.SetWritten(const AValue: boolean);
var
  i: integer;
begin
  for i := 0 to Count - 1 do
    Items[i].FWritten := AValue;
end;

constructor TXc12GraphicImages.Create(AErrors: TXLSErrorManager);
begin
  inherited Create;
  FErrors := AErrors;
end;

procedure TXc12GraphicImages.Delete(AImage: TXc12GraphicImage);
var
  i: integer;
begin
  for i := 0 to Count - 1 do begin
    if Items[i] = AImage then begin
      Items[i].DecRefCount;
      if Items[i].FRefCount <= 0 then
        inherited Delete(i);
      Exit;
    end;
  end;
end;

function TXc12GraphicImages.ExtFromType(AType: TXc12ImageType): AxUCString;
begin
  case AType of
    x12itJPEG   : Result := 'jpeg';
    x12itPNG    : Result := 'png';
    x12itEMF    : Result := 'emf';
    x12itWMF    : Result := 'wmf';
    else          raise XLSRWException.Create('Unknown image type');
  end;
end;

function TXc12GraphicImages.FindByDescr(const ADescr: AxUCString): TXc12GraphicImage;
var
  i: integer;
  S: AxUCString;
begin
  S := AnsiLowercase(ADescr);
  for i := 0 to Count - 1 do begin
    if AnsiLowercase(Items[i].Descr) = S then begin
      Result := Items[i];
      Exit;
    end;
  end;
  Result := Nil;
end;

function TXc12GraphicImages.FindByFileId(const AId: AxUCString): TXc12GraphicImage;
var
  i: integer;
begin
  for i := 0 to Count - 1 do begin
    if Items[i].FFileId = AId then begin
      Result := Items[i];
      Exit;
    end;
  end;
  Result := Nil;
end;

function TXc12GraphicImages.FindByRId(const ADrawingId,AId: AxUCString): TXc12GraphicImage;
var
  i: integer;
begin
  for i := 0 to Count - 1 do begin
    if (Items[i].FDrawingRId = ADrawingId) and (Items[i].FRId = AId) then begin
      Result := Items[i];
      Exit;
    end;
  end;
  Result := Nil;
end;

function TXc12GraphicImages.GetImageSize(AImage: TStream; AType: TXc12ImageType; out AWidth, AHeight: integer): boolean;
begin
  case AType of
    x12itJPEG   : Result := GetJpegSize(AImage,AWidth,AHeight);
    x12itPNG    : Result := GetPNGSize(AImage,AWidth,AHeight);
    x12itEMF,
    x12itWMF    : Result := GetMetafileSize(AImage,AWidth,AHeight);
    else          Result := False;
  end;
  AImage.Seek(0,soFromBeginning);
end;

function TXc12GraphicImages.GetItems(Index: integer): TXc12GraphicImage;
begin
  Result := TXc12GraphicImage(inherited Items[Index]);
end;

function TXc12GraphicImages.LoadFromFile(const AFilename: AxUCString): TXc12GraphicImage;
var
  S: AxUCString;
  Stream: TFileStream;
{$ifdef XLS_HAS_PNG_SUPPORT}
  Stream2: TMemoryStream;
{$endif}
  Ext: AxUCString;
begin
  S := AFilename;
  Ext := AnsiLowercase(ExtractFileExt(S));

  Stream := TFileStream.Create(AFilename,fmOpenRead{$ifdef DELPHI_6_OR_LATER},fmShareDenyNone{$endif});
  try
    if Ext = '.bmp' then begin
{$ifdef XLS_HAS_PNG_SUPPORT}
      S := ChangeFileExt(S,'.png');
      Stream2 := TMemoryStream.Create;
      try
        BmpToPng(Stream,Stream2);
        Result := LoadFromStream(Stream2,S);
      finally
        Stream2.Free;
      end;
{$else}
      FErrors.Warning('BMP',XLSWARN_UNKNOWNIMAGE);
      Result := Nil;
{$endif}
    end
    else
      Result := LoadFromStream(Stream,AFilename);
  finally
    Stream.Free;
  end;
end;

function TXc12GraphicImages.LoadFromStream(AStream: TStream; const AFilename: AxUCString): TXc12GraphicImage;
var
  IType: TXc12ImageType;
  W,H: integer;
  Stream: TPointerMemoryStream;
begin
  Result := Nil;

  IType := TypeFromExt(AFilename);

  if IType = x12itUnknown then begin
    FErrors.Warning(AFilename,XLSWARN_UNKNOWNIMAGE);
    Exit;
  end;

  Result := FindByDescr(AFilename);
  if Result <> Nil then begin
    if (Result.Width > 0) and (Result.Heigh > 0) then begin
      W := Result.Width;
      H := Result.Heigh;
    end
    else begin
      Stream := TPointerMemoryStream.Create;
      try
        Stream.SetStreamData(Result.Image,Result.ImageSize);
        if not GetImageSize(Stream,Result.ImageType,W,H) then begin
          FErrors.Warning('',XLSWARN_IMAGEERROR);
          Exit;
        end;
      finally
        Stream.Free;
      end;
    end;
  end
  else begin
    try
      if not GetImageSize(AStream,IType,W,H) then begin
        FErrors.Warning('',XLSWARN_IMAGEERROR);
        Exit;
      end;
      Result := Add(AStream,IType);
      Result.Descr := AFilename;
      Result.Width := W;
      Result.Heigh := H;
    except
      Result.Width := 0;
      Result.Heigh := 0;
    end;
  end;
end;

function TXc12GraphicImages.TypeFromExt(AExt: AxUCString): TXc12ImageType;
var
  p: integer;
begin
  p := RCPos('.',AExt);
  if p > 0 then
    AExt := Copy(AExt,p + 1,MAXINT);
  AExt := Lowercase(AExt);
  if (AExt = 'jpg') or (AExt = 'jpeg') then
    Result := x12itJPEG
  else if AExt = 'png' then
    Result := x12itPNG
  else if AExt = 'gif' then
    Result := x12itGIF
  else if AExt = 'emf' then
    Result := x12itEMF
  else if AExt = 'wmf' then
    Result := x12itWMF
  else
    Result := x12itUnknown;
end;

{ TXc12DrawingImage }

constructor TXc12GraphicImage.Create;
begin

end;

function TXc12GraphicImage.CreateBitmap: TBitmap;
var
{$ifdef XLS_HAS_PNG_SUPPORT}
  PNG: TPNGImage;
{$endif}
{$ifdef XLS_HAS_JPG_SUPPORT}
  JPG: TJPEGImage;
{$endif}
{$ifdef XLS_HAS_METAFILE_SUPPORT}
  Meta: TMetafile;
  MFCanvas:TMetaFileCanvas;
{$endif}
  Stream: TPointerMemoryStream;
begin
  Result := Nil;
  Stream := TPointerMemoryStream.Create;
  try
    Stream.SetStreamData(FImage,FImageSize);
    case FImageType of
      x12itJPEG   : begin
{$ifdef XLS_HAS_JPG_SUPPORT}
        JPG := TJPEGImage.Create;
        try
          JPG.LoadFromStream(Stream);
          Result := TBitmap.Create;
          Result.Assign(JPG);
        finally
          JPG.Free;
        end;
{$else}
        Result := Nil;
{$endif}
      end;
      x12itPNG    : begin
{$ifdef XLS_HAS_PNG_SUPPORT}
        PNG := TPNGImage.Create;
        try
          PNG.LoadFromStream(Stream);
          Result := TBitmap.Create;
          Result.Assign(PNG);
        finally
          PNG.Free;
        end;
{$else}
        Result := Nil;
{$endif}
      end;
      x12itEMF    : begin
{$ifdef XLS_HAS_METAFILE_SUPPORT}
        Meta := TMetafile.Create;
        try
          Meta.LoadFromStream(Stream);
          MFCanvas := TMetaFileCanvas.Create(Meta,0);
          Result := TBitmap.Create;
          MFCanvas.Draw(0,0,Result);
        finally
          Meta.Free;
        end;
{$else}
       Result := Nil;
{$endif}
      end;
      x12itWMF    : Result := Nil;
      x12itUnknown: Result := Nil;
    end;
  finally
    Stream.Free;
  end;
end;

procedure TXc12GraphicImage.DecRefCount;
begin
  Dec(FRefCount);
end;

destructor TXc12GraphicImage.Destroy;
begin
  if FImage <> Nil then
    FreeMem(FImage);
  inherited;
end;

function TXc12GraphicImage.UniqueName: AxUCString;
begin
  Result := Format('Image%.3d.%s',[Index + 1,TypeExt]);
end;

procedure TXc12GraphicImage.IncRefCount;
begin
  Inc(FRefCount);
end;

function TXc12GraphicImage.TypeExt: AxUCString;
begin
  Result := Xc12ImageTypeStr[FImageType];
end;

{ TXc12DrawingManager }

procedure TXc12GraphicManager.BeforeWrite;
begin
  FImages.SetWritten(False);
end;

procedure TXc12GraphicManager.Clear;
begin
  FImages.Clear;
end;

constructor TXc12GraphicManager.Create(AErrors: TXLSErrorManager);
begin
  FErrors := AErrors;

  FImages := TXc12GraphicImages.Create(FErrors);
end;

destructor TXc12GraphicManager.Destroy;
begin
  FImages.Free;
  inherited;
end;

function TXc12GraphicManager.CreateRelativeCells(ARef: AxUCString): TXLSRelCells;
begin
  Result := Nil;
end;

function TXc12GraphicManager.SaveImageToOPC(AImage: TXc12GraphicImage): AxUCString;
var
  OPC: TOPCItem;
  Stream: TPointerMemoryStream;
begin
  OPC := FXLSOPC.CreateImage(AImage.Index + 1,AImage.TypeExt);
  Result := OPC.Id;

  if not AImage.FWritten then begin
    Stream := TPointerMemoryStream.Create;
    try
      Stream.SetStreamData(AImage.Image,AImage.ImageSize);
      FXLSOPC.ItemWrite(OPC,Stream);
      FXLSOPC.ItemClose(OPC);
    finally
      Stream.Free;
    end;
    AImage.FWritten := True;
  end;
end;

end.
