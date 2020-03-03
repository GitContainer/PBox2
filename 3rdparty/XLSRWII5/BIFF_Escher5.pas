unit BIFF_Escher5;

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

uses Classes, SysUtils,
{$ifdef BABOON}
  {$ifdef DELPHI_XE5_OR_LATER}
     FMX.Graphics,
  {$else}
     FMX.Types,
  {$endif}
{$else}
     Windows, vcl.Graphics,
{$endif}
     Math, Contnrs,
     BIFF_Utils5, BIFF_Stream5, BIFF_RecsII5, BIFF_RecordStorage5, BIFF_EscherTypes5,
     BIFF_MD4Hash_5, BIFF_ExcelFuncII5, BIFF_Names5,
     BIFF_RecordStorageChart5, BIFF_EscherCtrlObj5, BIFF_EscherFileObjs5,
     Xc12DataStylesheet5,
     XLSUtils5, XLSZlibPas5;

{$ifdef BABOON}
type RGBQUAD = packed record
     rgbBlue    : byte;
     rgbGreen   : byte;
     rgbRed     : byte;
     rgbReserved: byte;
     end;
{$endif}

type TBasicShapeType = (bstLine,bstArrow,bstEllipse,bstRectangle);

//* Auto shape types.
type TAutoShapeType = (astSun,astMoon,astLeftArrow,astDownArrow,
astUpArrow,astLeftRightArrow,astUpDownArrow,astQuadArrow,astLeftArrowCallout,
astRightArrowCallout,astUpArrowCallout,astDownArrowCallout,
astLeftRightArrowCallout,astUpDownArrowCallout,astQuadArrowCallout);

//* ~exclude
type TOptValType = (ovtLongword,ovtString,ovtPointer);

//* ~exclude
type POPTData = ^TOPTData;
     TOPTData = record
     Options: word;
     Value: longword;
     Extra: PByteArray;
     end;

//* ~exclude
type PMSORecord = ^TMSORecord;
     TMSORecord = record
     VerInst: word;
     FBT: word;
     Length: longword;
     Data: array[0..MAXINT div 2] of byte;
     end;

//* ~exclude
type TEscherDrawing = class;
//* ~exclude
     TEscherGroup = class;

//* ~exclude
     TMSORecords = class(TList)
private
     function  GetData(Index: integer): PByteArray;
     function  GetFBT(Index: integer): word;
     function  GetInstance(Index: integer): word;
     function  GetLength(Index: integer): longword;
     function  GetVersion(Index: integer): byte;
     function  GetItems(Index: integer): PMSORecord;
public
     destructor Destroy; override;
     procedure Clear; override;
     procedure Add(Header: TMSOHeader; Data: PByteArray);
     procedure Write(Stream: TXLSStream);
     function  Size: integer;
     procedure Assign(Records: TMSORecords);

     property Items[Index: integer]: PMSORecord read GetItems; default;
     property Version[Index: integer]: byte read GetVersion;
     property Instance[Index: integer]: word read GetInstance;
     property FBT[Index: integer]: word read GetFBT;
     property Length[Index: integer]: longword read GetLength;
     property Data[Index: integer]: PByteArray read GetData;
     end;

//* ~exclude
     TOPT = class(TList)
private
     FItemCount: integer;

     function  GetId(Index: integer): word;
     function  GetComplex(Index: integer): boolean;
     function  GetValue(Index: integer): longword;
     procedure SetValue(Index: integer; const Value: longword);
     function  GetExtra(Index: integer): PByteArray;
public
     destructor Destroy; override;
     procedure Clear; override;
     procedure Assign(OPT: TOPT);
     procedure AddValue(Id: word; Value: longword; IsBlip: boolean = False);
     procedure AddString(Id: word; Value: AxUCString);
     procedure AddExtra(Id: word; Value: integer; IsBlip: boolean = False);
     procedure LoadFromStream(Stream: TXLSStream; Length: integer; PBuf: PByteArray; ItemCount: integer);
     procedure LoadFromBuffer(Length: integer; PBuf: PByteArray; ItemCount: integer);
     procedure WriteToStream(Stream: TXLSStream; PBuf: PByteArray);
     procedure WriteToBuffer(PBuf: PByteArray; IncludeHeader: boolean = True);
     function  Size: integer;
     function  TxId: longword;
     function  BlipId: integer;
     procedure SetBlipId(Id: integer);
     function  Find(Id: integer): integer;
     function  FindValue(Id: integer): longword; overload;
     function  FindValue(Id: integer; Default: longword): longword; overload;
     function  FindString(Id: integer): AxUCString;
     procedure UpdateValue(Id: integer; Value: longword);
     procedure UpdateString(Id: integer; Value: AxUCString);
     procedure UpdateExtra(Id: integer; Size: integer);
     procedure SetItemCount;

     property Id[Index: integer]: word read GetId; default;
     property Value[Index: integer]: longword read GetValue write SetValue;
     property Complex[Index: integer]: boolean read GetComplex;
     property Extra[Index: integer]: PByteArray read GetExtra;
     end;

//* ~exclude
     TShapeOutsideMsoChart = class(TShapeOutsideMso)
protected
     FRecords: TChartRecordUpdate;
     FDrawing: TEscherDrawing;
     FDrawingGroup: TEscherGroup;
     FFonts: TXc12Fonts;
public
     constructor Create(ObjId: integer; DrawingGroup: TEscherGroup; Fonts: TXc12Fonts);
     destructor Destroy; override;
     procedure Assign(Shape: TShapeOutsideMso); override;
     procedure Read(Stream: TXLSStream; PBuf: PByteArray); override;
     procedure Write(Stream: TXLSStream); override;

     property Records: TChartRecordUpdate read FRecords;
     property Drawing: TEscherDrawing read FDrawing write FDrawing;
     property DrawingGroup: TEscherGroup read FDrawingGroup;
     property Fonts: TXc12Fonts read FFonts;
     end;

//* ~exclude
     TShape = class(TObject)
private
     procedure SetShapeType(const Value: integer);
     function  GetFlipHorizontal: boolean;
     function  GetFlipVertical: boolean;
     procedure SetFlipHorizontal(const Value: boolean);
     procedure SetFlipVertical(const Value: boolean);
protected
     FSpId: longword;
     FShapeType: integer;
     FOptions: longword;
     FOPT: TOPT;
     // TODO Write these opts.
     FFOPT: TOPT;
     FUnknown: TMSORecords;
     FOBJ: TBaseRecordStorage;
     FInterface: TShapeOutsideMso;
public
     constructor Create;
     destructor Destroy; override;
     function  Size: integer; virtual;
     procedure WriteToStream(Stream: TXLSStream; PBuf: PByteArray; WriteMSODRAWING: boolean); virtual;
     procedure Assign(Shape: TShape); virtual;
     procedure MakeMacro;

     property SpId: longword read FSpId;
     property ShapeType: integer read FShapeType write SetShapeType;
     property Options: longword read FOptions;
     property OPT: TOPT read FOPT;
     property OBJ: TBaseRecordStorage read FOBJ;
     property ExShape: TShapeOutsideMso read FInterface write FInterface;
     property FlipHorizontal: boolean read GetFlipHorizontal write SetFlipHorizontal;
     property FlipVertical: boolean read GetFlipVertical write SetFlipVertical;

     property ShpInterface: TShapeOutsideMso read FInterface;
     end;

//* ~exclude
     TShapeClientAnchor = class(TShape)
private
     FVisible: boolean;
protected
     FAnchor: TMSOObjCLIENTANCHOR;
public
     constructor Create;
     destructor Destroy; override;
     function Size: integer; override;
     procedure Assign(Shape: TShape); override;
     procedure GetPos(var C1,C1Offs,R1,R1Offs,C2,C2Offs,R2,R2Offs: integer);

     property Anchor: TMSOObjCLIENTANCHOR read FAnchor;
     property Visible: boolean read FVisible write FVisible;
     end;

//* ~exclude
     TShapeChildAnchor = class(TShape)
protected
     FCHILDANCHOR: TMSORecCHILDANCHOR;
public
     function Size: integer; override;
     procedure Assign(Shape: TShape); override;

     property X1: longword read FCHILDANCHOR.X1 write FCHILDANCHOR.X1;
     property Y1: longword read FCHILDANCHOR.Y1 write FCHILDANCHOR.Y1;
     property X2: longword read FCHILDANCHOR.X2 write FCHILDANCHOR.X2;
     property Y2: longword read FCHILDANCHOR.Y2 write FCHILDANCHOR.Y2;
     end;

//* ~exclude
     TShapeGroup = class(TShape {ClientAnchor})
private
     FDrawing: TEscherDrawing;
     FList: TList;
     FSPGR: TMSORecSPGR;

     function GetItems(Index: integer): TShape;
public
     constructor Create(Drawing: TEscherDrawing);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(Shape: TShape); override;
//     function  Size: integer; override;
     function  PrivateSize: integer; virtual; abstract;
//     procedure WriteToStream(Stream: TXLSStream; PBuf: PByteArray; WriteMSODRAWING: boolean); override;
     procedure Add(Shape: TShape);
     procedure Delete(Index: integer);
     procedure DeleteBySpId(SpId: integer);
     function  Count: integer;

     property Items[Index: integer]: TShape read GetItems; default;
     property SPGR: TMSORecSPGR read FSPGR;
     end;

//* ~exclude
     TShapeGroupClientAnchor = class(TShapeGroup)
private
     FAnchor: TMSOObjCLIENTANCHOR;
public
     constructor Create(Drawing: TEscherDrawing);
     destructor Destroy; override;
     procedure Assign(Shape: TShape); override;
     procedure GetPos(var C1,C1Offs,R1,R1Offs,C2,C2Offs,R2,R2Offs: Integer);
     function  Size: integer; override;
     function  PrivateSize: integer; override;
     procedure WriteToStream(Stream: TXLSStream; PBuf: PByteArray; WriteMSODRAWING: boolean); override;

     property Anchor: TMSOObjCLIENTANCHOR read FAnchor;
     end;

//* ~exclude
     TShapeGroupChildAnchor = class(TShapeGroup)
private
     FCHILDANCHOR: TMSORecCHILDANCHOR;
public
     procedure Assign(Shape: TShape); override;
     function  Size: integer; override;
     function  PrivateSize: integer; override;
     procedure WriteToStream(Stream: TXLSStream; PBuf: PByteArray; WriteMSODRAWING: boolean); override;
     end;

//* Not used
//* ~param Sender Not used.
//* ~param Shape Not used.
     TFileReadShapeEvent = procedure (Sender: TObject; Shape: TShape) of object;

//* ~exclude
     TEscherDrawing = class(TObject)
private
     FParent: TEscherGroup;
     FGroup: TShapeGroup;
     FDgId: word;
     FDG: TMSORecDG;
     FFileReGroupItems: array of TMSOFileReGroupItem;
     FSolverContainer: TMSORecords;
     FFileReadShapeEvent: TFileReadShapeEvent;
     FFonts: TXc12Fonts;
     FInternalNames: TInternalNames;
     FMaxObjId: integer;
protected
     function  ReadRoot(Stream: TXLSStream; PBuf: PByteArray): integer;
     procedure ReadOBJ(Shape: TShape; Stream: TXLSStream; PBuf: PByteArray);
     function  GetSpId: integer;
     function  MaxSpId: integer;
     function  GetObjId: integer;
     procedure SetRootData;
public
     constructor Create(Parent: TEscherGroup; Fonts: TXc12Fonts; InternalNames: TInternalNames);
     destructor Destroy; override;
     procedure Clear;
     function  ShapeCount: integer;
     procedure LoadFromStream(Stream: TXLSStream; PBuf: PByteArray);
     procedure SaveToStream(Stream: TXLSStream; PBuf: PByteArray);
//     procedure SaveToStream(Stream: TXLSStream; PBuf: PByteArray; DrawingId: integer);
     procedure SetNoteData(Col,Row,Options,ObjId: word; Author: AxUCString);
     procedure AssignBlipIds(Blips: TList);
     procedure SetBlipRefCount;
     procedure AddEmpty;
     procedure Copy(Index: integer; DestCol,DestRow: word); overload;
     procedure Copy(Col1,Row1,Col2,Row2,DestCol,DestRow: word); overload;
     procedure CopyList(List: TList; Col1,Row1,Col2,Row2: integer);
     procedure InsertList(List: TList; DestCol,DestRow: integer);
     procedure DeleteList(List: TList);
     procedure Move(Index: integer; DestCol,DestRow: longword); overload;
     procedure Move(Col1,Row1,Col2,Row2,DestCol,DestRow: longword); overload;
     procedure Delete(Col1,Row1,Col2,Row2: longword);

     // Never call any of these methods directly. TDrawingObjects and
     // TControlObjects will go out of sync.
     procedure DeleteShape(SpId: integer);
     function  AddTextBox: TShapeClientAnchor;
     function  AddNote: TShapeClientAnchor;
     function  AddBasic: TShapeClientAnchor;
     function  AddAutoShape: TShapeClientAnchor;
     function  AddPicture: TShapeClientAnchor;
     function  AddChart: TShapeClientAnchor;

     function  AddListComboBox: TShapeClientAnchor;
     function  AddListBox: TShapeClientAnchor;
     function  AddComboBox: TShapeClientAnchor;
     function  AddComboBoxAutoFilter: TShapeClientAnchor;
     function  AddButton: TShapeClientAnchor;
     function  AddCheckBox: TShapeClientAnchor;
     function  AddRadioButton: TShapeClientAnchor;

     property Group: TShapeGroup read FGroup;
     property EscherGroup: TEscherGroup read FParent;
     property DgId: word read FDgId;
     property OnReadShape: TFileReadShapeEvent read FFileReadShapeEvent write FFileReadShapeEvent;
     end;

//* ~exclude
     TDGGData = class(TObject)
private
     FFIDCL: TDGGRecFIDCL;
     FDrawing: TEscherDrawing;
     FNewDrawing: boolean;

     procedure SetDrawing(const Value: TEscherDrawing);
public
     procedure LoadFromStream(Stream: TXLSStream);
     procedure WriteToStream(Stream: TXLSStream);

     property FIDCL: TDGGRecFIDCL read FFIDCL;
     property Drawing: TEscherDrawing read FDrawing write SetDrawing;
     property NewDrawing: boolean read FNewDrawing write FNewDrawing;
     end;

//* ~exclude
     TDGG = class(TList)
private
     FMaxSpId: longword;
     FSavedShapes: longword;
     FSavedDrawings: longword;

     function GetItems(Index: integer): TDGGData;
public
     destructor Destroy; override;
     procedure Clear; override;
     function Add(Item: Pointer): Integer;
     procedure LoadFromStream(Stream: TXLSStream; var Length: integer; PBuf: PByteArray);
     procedure WriteToStream(Stream: TXLSStream; PBuf: PByteArray);

     property Items[Index: integer]: TDGGData read GetItems;
     end;

     TMSOPictures = class;

//* TMSOPicture is a storage for pictures in worksheets. All pictures are
//* loaded into common TMSOPicture objects, and when inserted into a worksheet,
//* a link to the TMSOPicture object is created. This in order to save space.
//* Bitmap (BMP) pictures in size over 32kb are not accepted by Excel. Use
//* JPEG or PNG formats for pictures over 32kb. If TPNGImage support is enabled
//* (defined in XLSRWII2.inc), bitmap (BMP) files larger than 256 bytes are
//* automatically converted to PNG images.
     TMSOPicture = class(TObject)
private
     FOwner: TMSOPictures;
     FPicture: Pointer;
     FCompressedPicture: Pointer;
     FPictSize: integer;
     FPictType: TMSOBlipType;
     FRefCount: integer;
     FBlipName: AxUCString;
     FCompressedSize: integer;
     FHash: array[0..15] of byte;
     FFileBlipId: integer;

     FName: string;
     procedure SetFilename(const Value: AxUCString);
protected
     function  LoadFromStream(Stream: TXLSStream; var Length: integer; PBuf: PByteArray): boolean; overload;
     procedure WriteToStream(Stream: TXLSStream; PBuf: PByteArray);
     procedure MakeHash;
     procedure CompressPict;
     procedure FreeCompressedPict;

     property Picture: Pointer read FPicture;
     property PictSize: integer read FPictSize;
     property RefCount: integer read FRefCount write FRefCount;
     property BlipName: AxUCString read FBlipName write FBlipName;
public
     //* ~exclude
     constructor Create(AOwner: TMSOPictures);
     //* ~exclude
     destructor Destroy; override;
     //* Read a picture from a TStream.
     //* ~param Stream The stream to read the picture from.
     //* ~param Name Name of the picture.
     //* ~param PictType The type of the picture. As the file is in a stream (and not a disk file), there is
     //* no file extension from where the picture type can be detected.
     procedure LoadFromStream(Stream: TStream; Name: AxUCString; PictType: TMSOBlipType); overload;
     //* Read a picture from a file.
     //* ~param Filename The name of the disk file with the picture.
     procedure LoadFromFile(Filename: AxUCString);
     //* Save a picture to a TStream.
     //* ~param Stream The stream to write the picture to.
     procedure SaveToStream(Stream: TStream);
     //* Save a picture To a file.
     //* ~param Filename The name of the disk file that the picture shall be saved to.
     procedure SaveToFile(Filename: AxUCString);
     //* Copies the picture to a TBitmap object.
     //* ~param Bitmap The TBitmap object to copy the picture to.
     //* ~result True if the picture could be converted to a TBitmap.
     function CopyToBitmap(Bitmap: TBitmap): boolean;
     //* Copies the picture to a TMetafile object.
     //* ~param Metafile The TMetafile object to copy the picture to.
     //* ~result True if the picture could be converted to a TMetafile.
{$ifdef WINDOWS}
     function CopyToMetafile(Metafile: TMetafile): boolean;
{$endif}
     //* Checks if the picture is a raster (not a metafile) picture.
     //* ~result True if the picture is a raster picture. 
     function IsRaster: boolean;

     //* The type of the picture.
     property PictType: TMSOBlipType read FPictType;
     //* The size of the picture.
     property PictureSize: integer read FPictSize;
     //* ~exclude
     property PictureBuf: Pointer read FPicture;
     //* ~exclude
     property FileBlipId: integer read FFileBlipId write FFileBlipId;

     //* The (file)name of the picture. This name is used to identify pictures
     //* when inserted into worksheets.
     property Filename: AxUCString read FBlipName write SetFilename;
     end;

//* ~exclude
     TEscherGroup = class(TObjectList)
private
     function GetMSOPicture(Index: integer): TMSOPicture;
protected
     FDGG: TDGG;
     FOPT: TOPT;
     FSplitMenuColors: array of RGBQUAD;
     FColorMRU: array of RGBQUAD;
     FUnknown: TMSORecords;
     FCurrSpIdBlock: integer;

     procedure ReadPictures(Stream: TXLSStream; Length: integer; PBuf: PByteArray);
     function  GetSpIdBlock: integer;
     procedure AssignDrawing(Drawing: TEscherDrawing);
     procedure UpdateDrawing(DgId: integer);
     procedure DeleteDrawing(DgId: integer);
public
     constructor Create;
     destructor Destroy; override;
     procedure Clear; override;
     procedure AddDrawing(Drawing: TEscherDrawing; var DgId: word; var SpIdBlock: longword);
     procedure LoadFromStream(Stream: TXLSStream; PBuf: PByteArray);
     procedure WriteToStream(Stream: TXLSStream; PBuf: PByteArray);
     function  HasData: boolean;

     property Items[Index: integer]: TMSOPicture read GetMSOPicture; default;
     end;

//* List TMSOPicture.
     TMSOPictures = class(TEscherGroup)
private
protected
public
     //* Add a new TMSOPicture.
     //* ~result The new TMSOPicture object.
     function  Add: TMSOPicture;
     //* ~exclude
     procedure GetBlipIds(List: TList);
     //* ~exclude
     procedure ResetBlipRefCount;
     end;

//* ~exclude
function BlipFromExt(Filename: AxUCString): TMSOBlipType;

implementation

var
  DebugCount: integer;

const SP_ID_BLOCK_SIZE = $0400;

const SIZE_WMF_HEADER = 22;

//* ~exclude
type PBMPHeader = ^TBMPHeader;
     TBMPHeader = packed record
     ID: word;     // Shall always be $4D42
     Size: word;   // Picture size.
     Blocks: byte; // If Size > $FFFF then this is No. of 64k blocks. Size is the remaining.
     Res1: array[0..4] of byte;
     BMPType: word; // Seems to be $0436 or $0036
     Res2: word;
     end;

function GetTxId: integer;
begin
  // Don't know the correct value for TxId...
  Result := {$ifndef BABOON} Integer(GetTickCount) + {$endif} Random($FFFF);
end;

function BlipFromExt(Filename: AxUCString): TMSOBlipType;
var
  Ext: AxUCString;
begin
  Ext := MyWideUpperCase(Copy(ExtractFileExt(Filename),2,MAXINT));
       if Ext = 'WMF'  then Result := msoblipWMF
  else if Ext = 'EMF'  then Result := msoblipEMF
  else if Ext = 'PICT' then Result := msoblipPICT
  else if Ext = 'JPG'  then Result := msoblipJPEG
  else if Ext = 'JPEG' then Result := msoblipJPEG
  else if Ext = 'PNG'  then Result := msoblipPNG
  else if Ext = 'BMP'  then Result := msoblipDIB
  else raise XLSRWException.Create('Unknown picture type ' + Ext);
end;


{ TShape }

procedure TShape.Assign(Shape: TShape);
begin
  Shape.FSpId := 0;
  Shape.FShapeType := FShapeType;
  Shape.FOptions := FOptions;
  FOPT.Assign(Shape.FOPT);
  if FUnknown <> Nil then begin
    Shape.FUnknown := TMSORecords.Create;
    FUnknown.Assign(Shape.FUnknown);
  end;
  FOBJ.Assign(Shape.FOBJ);
  if FInterface <> Nil then begin
    if FInterface is TShapeOutsideMsoNote then
      Shape.FInterface := TShapeOutsideMsoNote.Create(0,mstNote)
    else if FInterface is TShapeOutsideMsoBaseText then
      Shape.FInterface := TShapeOutsideMsoBaseText.Create(0,mstTextBox)
    else if FInterface is TShapeOutsideMsoChart then begin
// Note_JKQWNH
//  Created in calling method, as Shape.FInterface constructor require parameters not available here. See Note_JKQWNH
    end
    else
      raise XLSRWException.Create('Can not assign this interface');
    FInterface.Assign(Shape.FInterface);
    Shape.FOPT.UpdateValue(msooptLTxid,GetTxId);
  end;
end;

constructor TShape.Create;
begin
  FOPT := TOPT.Create;
  FFOPT := TOPT.Create;
  FOBJ := TBaseRecordStorage.Create;
end;

destructor TShape.Destroy;
begin
  FOPT.Free;
  FFOPT.Free;
  FOBJ.Free;
  FUnknown.Free;
  FInterface.Free;
  inherited;
end;

function TShape.GetFlipHorizontal: boolean;
begin
  Result := (FOptions and SpOptFlipH) = SpOptFlipH;
end;

function TShape.GetFlipVertical: boolean;
begin
  Result := (FOptions and SpOptFlipV) = SpOptFlipV;
end;

procedure TShape.MakeMacro;
begin
  FOPT.AddValue(msooptFNoFillHitTest,$00010001);
  FOPT.AddValue(msooptFPrint,$00080008);
  FOPT.FItemCount := FOPT.Count;
end;

procedure TShape.SetFlipHorizontal(const Value: boolean);
begin
  FOptions := FOptions or SpOptFlipH;
end;

procedure TShape.SetFlipVertical(const Value: boolean);
begin
  FOptions := FOptions or SpOptFlipV;
end;

procedure TShape.SetShapeType(const Value: integer);
begin
  FShapeType := Value;
  case FShapeType of
    msosptLine,msosptArrow:
      PObjCMO(@FOBJ[0].Data).ObjType := OBJTYPE_LINE;
    msosptEllipse:
      PObjCMO(@FOBJ[0].Data).ObjType := OBJTYPE_OVAL;
    msosptRectangle:
      PObjCMO(@FOBJ[0].Data).ObjType := OBJTYPE_RECTANGLE;
  end;
end;

function TShape.Size: integer;
begin
            // SPCONTAINER
  Result := SizeOf(TMSOHeader) +
            // SP
            SizeOf(TMSOHeader) + SizeOf(TMSORecSP) +
            // OPT
            SizeOf(TMSOHeader) + FOPT.Size +
            // CLIENTDATA
            SizeOf(TMSOHeader);
  if FInterface is TShapeOutsideMsoBaseText then
    Inc(Result,SizeOf(TMSOHeader));
  if FUnknown <> Nil then
    Inc(Result,FUnknown.Size);
end;

procedure TShape.WriteToStream(Stream: TXLSStream; PBuf: PByteArray; WriteMSODRAWING: boolean);
var
  Sz: integer;
begin
  Sz := Size;
  if WriteMSODRAWING then begin
    if FInterface is TShapeOutsideMsoBaseText then
      Stream.WriteHeader(BIFFRECID_MSODRAWING,Sz - SizeOf(TMSOHeader))
    else
      Stream.WriteHeader(BIFFRECID_MSODRAWING,Sz);
  end;
  Dec(Sz,SizeOf(TMSOHeader));
  Stream.WriteMSOHeader($0F,$0000,MSO_SPCONTAINER,Sz);
  Stream.WriteMSOHeader($02,FShapeType,MSO_SP,SizeOf(TMSORecSP));
  Stream.Write(FSpId,SizeOf(FSpId));
  Stream.Write(FOptions,SizeOf(FOptions));
  FOPT.WriteToStream(Stream,PBuf);

  if FUnknown <> Nil then
    FUnknown.Write(Stream);
  if Self is TShapeClientAnchor then
    TShapeClientAnchor(Self).Anchor.Write(Stream)
  else if Self is TShapeChildAnchor then begin
    Stream.WriteMSOHeader($00,$0000,MSO_CHILDANCHOR,SizeOf(TMSORecCHILDANCHOR));
    Stream.Write(TShapeChildAnchor(Self).FCHILDANCHOR,SizeOf(TMSORecCHILDANCHOR));
  end;
  Stream.WriteMSOHeader($00,$0000,MSO_CLIENTDATA,0);

{
    Stream.WriteHeader(BIFFRECID_OBJ,FOBJ.Size);
    FOBJ.WriteAllRecs(Stream);
}

  if (FInterface <> Nil) and FInterface.OwnsObjData then
    FInterface.Write(Stream)
  else begin
    Stream.WriteHeader(BIFFRECID_OBJ,FOBJ.Size);
    FOBJ.WriteAllRecs(Stream);
    if FInterface <> Nil then
      FInterface.Write(Stream);
  end;
end;

{ TShapeGroup }

procedure TShapeGroup.Add(Shape: TShape);
begin
  FList.Add(Shape);
end;

procedure TShapeGroup.Assign(Shape: TShape);
var
  i: integer;
  Shp: TShape;
begin
  inherited Assign(Shape);
  TShapeGroup(Shape).FSPGR := FSPGR;
  for i := 0 to Count - 1 do begin
    if Items[i] is TShapeGroupClientAnchor then
      Shp := TShapeGroupClientAnchor.Create(FDrawing)
    else if Items[i] is TShapeGroupChildAnchor then
      Shp := TShapeGroupChildAnchor.Create(FDrawing)
    else if Items[i] is TShapeClientAnchor then
      raise XLSRWException.Create('ClientAnchor in ShapeGroup')
    else if Items[i] is TShapeChildAnchor then
      Shp := TShapeChildAnchor.Create
    else
      raise XLSRWException.Create('Dont know how to Copy this shape');
    Items[i].Assign(Shp);
    Shp.FSpId := FDrawing.GetSpId;
    TShapeGroup(Shape).Add(Shp);
  end;
end;

procedure TShapeGroup.Clear;

procedure ClearGroup(Group: TShapeGroup);
var
  i: integer;
begin
  for i := 0 to Group.Count - 1 do begin
    if Group[i] is TShapeGroup then
      ClearGroup(TShapeGroup(Group[i]));
    Group[i].Free;
  end;
  Group.FList.Clear;
end;

begin
  ClearGroup(Self);
end;

function TShapeGroup.Count: integer;
begin
  Result := FList.Count;
end;

constructor TShapeGroup.Create(Drawing: TEscherDrawing);
begin
  inherited Create;
  FDrawing := Drawing;
  FList := TList.Create;
end;

procedure TShapeGroup.Delete(Index: integer);
begin
  TShape(FList.Items[Index]).Free;
  FList.Delete(Index);
end;

procedure TShapeGroup.DeleteBySpId(SpId: integer);
var
  i: integer;
begin
  for i := 0 to FList.Count - 1 do begin
    if Integer(TShape(FList.Items[i]).FSpId) = SpId then begin
      Delete(i);
      Exit;
    end;
  end;
end;

destructor TShapeGroup.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

function TShapeGroup.GetItems(Index: integer): TShape;
begin
  Result := TShape(FList.Items[Index]);
end;

{
function TShapeGroup.PrivateSize: integer;
begin
  Result := inherited Size +
            // SPGRCONTAINER
            SizeOf(TMSOHeader) +
            // SPGR
            SizeOf(TMSOHeader) + SizeOf(TMSORecSPGR);
end;
}

{
function TShapeGroup.Size: integer;
var
  i: integer;
begin
  Result := PrivateSize;
  for i := 0 to FList.Count - 1 do
    Inc(Result,Items[i].Size);
end;
}

{
procedure TShapeGroup.WriteToStream(Stream: TXLSStream; PBuf: PByteArray; WriteMSODRAWING: boolean);
var
  Sz: integer;
  i: integer;
begin
  Sz := PrivateSize;
  if WriteMSODRAWING then
    Stream.WriteHeader(BIFFRECID_MSODRAWING,Sz);
  Stream.WriteMSOHeader($0F,$0000,MSO_SPGRCONTAINER,Size - SizeOf(TMSOHeader));

  Stream.WriteMSOHeader($0F,$0000,MSO_SPCONTAINER,PrivateSize - (SizeOf(TMSOHeader) * 2));

  Stream.WriteMSOHeader($01,$0000,MSO_SPGR,SizeOf(TMSORecSPGR));
  Stream.Write(FSPGR,SizeOf(TMSORecSPGR));

  Stream.WriteMSOHeader($02,FShapeType,MSO_SP,SizeOf(TMSORecSP));
  Stream.Write(FSpId,SizeOf(FSpId));
  Stream.Write(FOptions,SizeOf(FOptions));

  FOPT.WriteToStream(Stream,PBuf);
  if FUnknown <> Nil then
    FUnknown.Write(Stream);
  Stream.WriteMSOHeader($00,$0000,MSO_CLIENTANCHOR,SizeOf(TMSORecCLIENTANCHOR));
  Stream.Write(FCLIENTANCHOR,SizeOf(TMSORecCLIENTANCHOR));
  Stream.WriteMSOHeader($00,$0000,MSO_CLIENTDATA,0);

  Stream.WriteHeader(BIFFRECID_OBJ,FOBJ.Size);
  FOBJ.WriteAllRecs(Stream);

  if FInterface <> Nil then
    FInterface.Write(Stream);
  for i := 0 to FList.Count - 1 do
    Items[i].WriteToSTream(Stream,PBuf,True);
end;
}

{ TEscherDrawing }

procedure TEscherDrawing.Clear;
begin
  FGroup.Clear;
  if FDgId > 0 then
    FParent.DeleteDrawing(FDgId);
  FDgId := 0;
end;

constructor TEscherDrawing.Create(Parent: TEscherGroup; Fonts: TXc12Fonts; InternalNames: TInternalNames);
begin
  FParent := Parent;
  FFonts := Fonts;
  FInternalNames := InternalNames;
  FGroup := TShapeGroupClientAnchor.Create(Self);
end;

destructor TEscherDrawing.Destroy;
begin
  FGroup.Free;
  FSolverContainer.Free;
  if FDgId > 0 then
    FParent.DeleteDrawing(FDgId);
  inherited;
end;

function TEscherDrawing.ReadRoot(Stream: TXLSStream; PBuf: PByteArray): integer;
var
  i,Count: longword;
  Header: TMSOHeader;
begin
  Stream.Read(Header,SizeOf(TMSOHeader));
  if Header.FBT <> MSO_DG then
    raise XLSRWException.Create('Expected record missing: DG');

  Stream.Read(FDG,SizeOf(TMSORecDG));
  FDgId := (Header.VerInst shr 4) and $0FFF;
  FParent.AssignDrawing(Self);

  Stream.Read(Header,SizeOf(TMSOHeader));
  if Header.FBT = MSO_REGROUPITEMS then begin
    Count := (Header.VerInst shr 4) and $0FFF;
    SetLength(FFileReGroupItems,Count);
    for i := 0 to Count - 1 do
      Stream.Read(FFileReGroupItems[i],SizeOf(TMSOFileReGroupItem));
    Stream.Read(Header,SizeOf(TMSOHeader));
  end;
  if Header.FBT <> MSO_SPGRCONTAINER then
    raise XLSRWException.Create('Expected record missing: SPGRCONTAINER');
  Result := Header.Length;

  Stream.Read(Header,SizeOf(TMSOHeader));
  if Header.FBT <> MSO_SPCONTAINER then
    raise XLSRWException.Create('Expected record missing: SPCONTAINER');
  Dec(Result,Header.Length + SizeOf(TMSOHeader));

  Stream.Read(Header,SizeOf(TMSOHeader));
  if Header.FBT <> MSO_SPGR then
    raise XLSRWException.Create('Expected record missing: SPGR');
  Stream.Read(FGroup.FSPGR,SizeOf(TMSORecSPGR));

  Stream.Read(Header,SizeOf(TMSOHeader));
  if Header.FBT <> MSO_SP then
    raise XLSRWException.Create('Expected record missing: SP');
  Stream.Read(FGroup.FSpId,SizeOf(FGroup.FSpId));
  Stream.Read(FGroup.FOptions,SizeOf(FGroup.FOptions));
  if (FGroup.FOptions and SpOptPatriarch) = 0 then
    raise XLSRWException.Create('SP is not root');
end;

procedure TEscherDrawing.ReadOBJ(Shape: TShape; Stream: TXLSStream; PBuf: PByteArray);
var
  ObjType: word;
  Len: integer;
  Header: TBIFFHeader;

function GetLbsSize: integer;
var
  p: integer;
begin
  p := Stream.Pos;
  Stream.Read(PBuf[0],8);
  Stream.Read(PBuf[8],PWordArray(PBuf)[1]);
  Stream.Read(PBuf[8 + PWordArray(PBuf)[1]],7);
  // Two extra bytes for invalid END record
  if (PWordArray(@PBuf[8 + PWordArray(PBuf)[1] + 5])[0] and $0030) <> 0 then
    Stream.Read(PBuf[8 + PWordArray(PBuf)[1] + 7],PWordArray(@PBuf[8 + PWordArray(PBuf)[1] + 1])[0] + 2)
  else
    Stream.Read(PBuf[8 + PWordArray(PBuf)[1] + 7],2);
  Result := Stream.Pos - p - 2;
end;

begin
  ObjType := 0;
  Stream.ReadOtherHeader(Header);
//  Stream.ReadHeader(Header);
  // This may occure if the shape is deleted. Fix it.
  if Header.RecID <> BIFFRECID_OBJ then
    raise XLSRWException.Create('Expected record missing: OBJ');
  Len := Header.Length;
  while Len > 0 do begin
    Stream.Read(Header,SizeOf(TBIFFHeader));
    // This record always has invalid length.
    // As it always seems to be the second last record (before OBJREC_END),
    // calculate the length.
    if Header.RecID = OBJREC_LBSDATA then begin
      Header.Length := Len - SizeOf(TBIFFHeader);
      Stream.Read(PBuf^,Header.Length);
      Shape.FOBJ.AddRec(Header,PBuf);

      if ObjType in [OBJTYPE_LIST_BOX,OBJTYPE_COMBO_BOX] then begin
        // LBS list box do not have a END record.
        Header.RecID := OBJREC_END;
        Header.Length := 0;
        Shape.FOBJ.AddRec(Header,PBuf);
        Break;
      end;
    end;
    if not (Header.RecID in [OBJREC_END,OBJREC_MACRO..OBJREC_CMO]) then
      raise XLSRWException.CreateFmt('Unknown OBJ record %.2X',[Header.RecId]);
    // OBJREC_END may have invalid length...
    if Header.RecID = OBJREC_END then begin
      // The OBJ record (005D) may be padded with FF and have a size larger than
      // the parts of the OBJ Data.
      if Len > 4 then
        Stream.Seek(Len + 4,soFromCurrent);
      Header.Length := 0;
      Shape.FOBJ.AddRec(Header,PBuf);
      Break;
    end;
    if Header.RecID <> OBJREC_LBSDATA then
      Stream.Read(PBuf^,Header.Length);
    Dec(Len,Header.Length + SizeOf(TBIFFHeader));
    if Header.RecID = OBJREC_CMO then begin
      ObjType := PObjCMO(PBuf).ObjType;
      if PObjCMO(PBuf).ObjId > FMaxObjId then
        FMaxObjId := PObjCMO(PBuf).ObjId;
    end;
    if Header.RecID <> OBJREC_LBSDATA then
      Shape.FOBJ.AddRec(Header,PBuf);
  end;
end;

procedure TEscherDrawing.LoadFromStream(Stream: TXLSStream; PBuf: PByteArray);
var
  Count: integer;
  Header: TMSOHeader;
  BIFFHeader: TBIFFHeader;
  StreamPos,L: integer;

procedure ReadMSODRAWING;
begin
  Stream.ResetCONTINUERead;
  Stream.ReadHeader(BIFFHeader);
  // BUG in excel: when writing notes MSODRAWING may be replaced with CONTINUE
  if not (BIFFHeader.RecID in [BIFFRECID_MSODRAWING,BIFFRECID_CONTINUE]) then begin
    if BIFFHeader.RecID = BIFFRECID_IMDATA then begin
      Stream.Read(PBuf^,BIFFHeader.Length);
      while Stream.PeekHeader in [BIFFRECID_IMDATA,BIFFRECID_CONTINUE] do begin
        Stream.ReadHeader(BIFFHeader);
        Stream.Read(PBuf^,BIFFHeader.Length);
      end;
      ReadMSODRAWING;
    end
    else
      raise XLSRWException.Create('Expected record missing: MSODRAWING');
  end;
end;

procedure ReadShape(Group: TShapeGroup; GroupLen: integer);
var
  V1,V2: longword;
  Instance: word;
  Len,ChildLen: integer;
  Shape: TShape;
  SPGR: TMSORecSPGR;
  MSODRAWING_Read: boolean;
begin
  while GroupLen > 0 do begin
    Inc(Count);
    Stream.Read(Header,SizeOf(TMSOHeader));
    Len := Header.Length;

    // Some problem with the file if this is true. The file can be read, but not written.
    if Len > GroupLen then
      Exit;

    ChildLen := Len;
    if Header.FBT = MSO_SPGRCONTAINER then begin
      ChildLen := Header.Length;
      Stream.Read(Header,SizeOf(TMSOHeader));
      Len := Header.Length;
      Dec(ChildLen,Len + SizeOf(TMSOHeader));
    end;
    MSODRAWING_Read := False;
    Shape := Nil;
    Dec(GroupLen,Len + SizeOf(TMSOHeader));
    if Header.FBT <> MSO_SPCONTAINER then
      raise XLSRWException.Create('Expected record missing: SPCONTAINER');
    while Len > 0 do begin
      // Some files may have an extra MSODRAWING record after the OBJ record.
      // This record continues the parent MSODRAWING record.
      // Also, some files have a USERDEFPROP record that belongs to the first
      // record, written to the following MSODRAWING record.
      if Stream.PeekHeader = BIFFRECID_MSODRAWING then begin
        Stream.Seek(SizeOf(TBIFFHeader),soFromCurrent);
        MSODRAWING_Read := True;
      end;

      Inc(DebugCount);

      Stream.Read(Header,SizeOf(TMSOHeader));
      Instance := (Header.VerInst shr 4) and $0FFF;
      case Header.FBT of
        MSO_SPGR: begin
          Stream.Read(SPGR,SizeOf(TMSORecSPGR));
        end;
        MSO_SP: begin
          Stream.Read(V1,SizeOf(V1));
          Stream.Read(V2,SizeOf(V2));
          if (V2 and SpOptGroup) = SpOptGroup then begin
            if (V2 and SpOptChild) = SpOptChild then
              Shape := TShapeGroupChildAnchor.Create(Self)
            else
              Shape := TShapeGroupClientAnchor.Create(Self);
            TShapeGroup(Shape).FSPGR := SPGR;
          end
          else if (V2 and SpOptChild) = SpOptChild then
            Shape := TShapeChildAnchor.Create
          else
            Shape := TShapeClientAnchor.Create;
          Shape.FSpId := V1;
          Shape.FOptions := V2;
          Shape.FShapeType := Instance;
        end;
        MSO_OPT: begin
          Shape.FOPT.LoadFromStream(Stream,Header.Length,PBuf,Instance);
        end;
{
        MSO_USERDEFPROP: begin
          raise XLSRWException.Create('MSO_USERDEFPROP');
          Stream.Read(PBuf^,Header.Length);
        end;
}
        MSO_CHILDANCHOR: begin
          if Shape is TShapeChildAnchor then
            Stream.Read(TShapeChildAnchor(Shape).FCHILDANCHOR,SizeOf(TMSORecCHILDANCHOR))
          else if Shape is TShapeGroupChildAnchor then
            Stream.Read(TShapeGroupChildAnchor(Shape).FCHILDANCHOR,SizeOf(TMSORecCHILDANCHOR))
          else
            raise XLSRWException.Create('CHILDANCHOR when shape not is ChildAnchor');
        end;
        MSO_CLIENTANCHOR: begin
          if Shape is TShapeClientAnchor then
            TShapeClientAnchor(Shape).Anchor.Read(Stream)
          else if Shape is TShapeGroupClientAnchor then
            TShapeGroupClientAnchor(Shape).Anchor.Read(Stream)
          else
            raise XLSRWException.Create('CLIENTANCHOR when shape not is ClientAnchor');
        end;
        MSO_CLIENTDATA: begin
          if Header.Length > 0 then
            raise XLSRWException.Create('Size of CLIENTDATA > 0');
          // Required for comments larger than ~8k
          Stream.EndCONTINUERead;
          // Not sure if this is the best place to read OBJ...
          ReadOBJ(Shape,Stream,PBuf);
          case PObjCMO(@Shape.FOBJ[0].Data).ObjType of
            OBJTYPE_COMMENT: begin
              Shape.FInterface := TShapeOutsideMsoNote.Create(0,mstNote);
              Shape.FInterface.Read(Stream,PBuf);
              Dec(Len,SizeOf(TMSOHeader));
            end;
            OBJTYPE_TEXT: begin
              // TxId seems to be missing on "empty/deleted" text boxes.
              // These text boxes don't have these records avter OBJ:
              // MOSODRAWING, TXO, CONTINUE, CONTINUE
              if Shape.OPT.TxId <> 0 then begin
                Shape.FInterface := TShapeOutsideMsoBaseText.Create(0,mstTextBox);
                Shape.FInterface.Read(Stream,PBuf);
                Dec(Len,SizeOf(TMSOHeader));
              end;
            end;
            OBJTYPE_BUTTON: begin
              Shape.FInterface := TShapeControlButton.Create(0,Self.FInternalNames);
              if Shape.FInterface.ParseOBJ(Shape.FOBJ) then
                Shape.FOBJ.Clear;
              Shape.FInterface.Read(Stream,PBuf);
              Dec(Len,SizeOf(TMSOHeader));
            end;
            OBJTYPE_CHART: begin
              Shape.FInterface := TShapeOutsideMsoChart.Create(0,FParent,FFonts);
              Shape.FInterface.Read(Stream,PBuf);
            end;
            OBJTYPE_SCROLL_BAR: begin
              Shape.FInterface := TShapeControlScrollBar.Create(0);
              if Shape.FInterface.ParseOBJ(Shape.FOBJ) then
                Shape.FOBJ.Clear;
            end;
            OBJTYPE_LIST_BOX: begin
              Shape.FInterface := TShapeControlListBox.Create(0,mstCtrlListBox);
              if Shape.FInterface.ParseOBJ(Shape.FOBJ) then
                Shape.FOBJ.Clear;
            end;
            OBJTYPE_COMBO_BOX: begin
              Shape.FInterface := TShapeControlComboBox.Create(0,mstCtrlComboBox);
              if Shape.FInterface.ParseOBJ(Shape.FOBJ) then
                Shape.FOBJ.Clear;
            end;
            OBJTYPE_CHECK_BOX: begin
              Shape.FInterface := TShapeControlCheckBox.Create(0,Self.FInternalNames);
              if Shape.FInterface.ParseOBJ(Shape.FOBJ) then
                Shape.FOBJ.Clear;
              Shape.FInterface.Read(Stream,PBuf);
              Dec(Len,SizeOf(TMSOHeader));
            end;
            OBJTYPE_OPTION_BUTTON: begin
              Shape.FInterface := TShapeControlRadioButton.Create(0,Self.FInternalNames);
              if Shape.FInterface.ParseOBJ(Shape.FOBJ) then
                Shape.FOBJ.Clear;
              Shape.FInterface.Read(Stream,PBuf);
              Dec(Len,SizeOf(TMSOHeader));
            end;
            else begin
              if Shape.FOPT.TxId <> 0 then begin
                case PObjCMO(@Shape.FOBJ[0].Data).ObjType of
                  OBJTYPE_LABEL    : Shape.FInterface := TShapeOutsideMsoBaseText.Create(0,mstCtrlLabel);
                  OBJTYPE_GROUP_BOX: Shape.FInterface := TShapeOutsideMsoBaseText.Create(0,mstCtrlGroupBox);
                  OBJTYPE_MICROSOFT_OFFICE_DRAWING: Shape.FInterface := TShapeOutsideMsoBaseText.Create(0,mstTextShape);
                  else               Shape.FInterface := TShapeOutsideMsoBaseText.Create(0,mstUnknown);
                end;
                Shape.FInterface.Read(Stream,PBuf);
                Dec(Len,SizeOf(TMSOHeader));
              end;
            end;
          end;
        end;
        MSO_TERTIARY_FOPT: begin
          Shape.FFOPT.LoadFromStream(Stream,Header.Length,PBuf,Instance);
        end
        else begin
          if Shape.FUnknown = Nil then
            Shape.FUnknown := TMSORecords.Create;
          if Header.Length > MAXRECSZ_97 then begin
            ReallocMem(PBuf,Header.Length);
          end;
          Stream.Read(PBuf^,Header.Length);
          Shape.FUnknown.Add(Header,PBuf);
        end;
      end;
      Dec(Len,Header.Length + SizeOf(TMSOHeader));
    end;
    if Len <> 0 then
      raise XLSRWException.Create('Error while reading Shape. Length <> 0');
    Group.Add(Shape);
    if Assigned(FFileReadShapeEvent) then
      FFileReadShapeEvent(Self,Shape);
    if (GroupLen > 0) and not MSODRAWING_Read then
      ReadMSODRAWING;
    if Shape is TShapeGroup then begin
      Dec(GroupLen,ChildLen + SizeOf(TMSOHeader));
      ReadShape(TShapeGroup(Shape),ChildLen);
      if GroupLen > 0 then
        ReadMSODRAWING;
    end;
  end;
  if GroupLen <> 0 then
    raise XLSRWException.Create('Error while reading Shape. Group Length <> 0');
end;

begin
  Stream.Read(Header,SizeOf(TMSOHeader));
  if Header.FBT <> MSO_DGCONTAINER then
    raise XLSRWException.Create('Expected record missing: DGCONTAINER');
  FMaxObjId := 0;
  Count := 0;
  try
    StreamPos := Stream.Pos;
    ReadShape(FGroup,ReadRoot(Stream,PBuf));

    if Stream.PeekHeader = BIFFRECID_MSODRAWING then begin
      Stream.ReadHeader(BIFFHeader);
      FSolverContainer := TMSORecords.Create;
      while BIFFHeader.Length > 0 do begin
        Stream.Read(Header,SizeOf(TMSOHeader));
        if (Header.VerInst and $0F) = $0F then
          Dec(BIFFHeader.Length,SizeOf(TMSOHeader))
        else begin
          Stream.Read(PBuf^,Header.Length);
          Dec(BIFFHeader.Length,Header.Length + SizeOf(TMSOHeader));
        end;
        FSolverContainer.Add(Header,PBuf);
      end;
    end
    else begin
      // This is not optimal.
      // Files with deleted (empty) drawings may contain other records
      // (SOLVERCONTAINER) after the shapes. Ths will ensure that thease
      // records are read.
      // The correct method is to decrement the MSODRAWING record length
      // when anything in the drawing is read.
      L := Integer(Header.Length) - (Stream.Pos - StreamPos);
      if L > 0 then
        Stream.Seek(L,soFromCurrent);
    end;
  except
    on E: XLSRWException do
      raise XLSRWException.CreateFmt('Error on reading shape #%d' + #13 + E.Message,[Count]);
  end;
end;

procedure TEscherDrawing.SetNoteData(Col, Row, Options, ObjId: word; Author: AxUCString);
var
  i: integer;
begin
  for i := 0 to FGroup.Count - 1 do begin
    if (FGroup[i].FInterface <> Nil) and (FGroup[i].FInterface is TShapeOutsideMsoNote) and (PObjCMO(@FGroup[i].FOBJ[0].Data).ObjId = ObjId) then begin
      TShapeOutsideMsoNote(FGroup[i].FInterface).CellCol := Col;
      TShapeOutsideMsoNote(FGroup[i].FInterface).CellRow := Row;
      TShapeOutsideMsoNote(FGroup[i].FInterface).Options := Options;
      TShapeOutsideMsoNote(FGroup[i].FInterface).Author := Author;
      TShapeClientAnchor(FGroup[i]).FVisible := (Options and $0002) = $0002;
      Exit;
    end;
  end;
  raise XLSRWException.CreateFmt('Can not find shape for note #%d',[ObjId]);
end;

procedure TEscherDrawing.SaveToStream(Stream: TXLSStream; PBuf: PByteArray {; DrawingId: integer});
var
  i,Sz: integer;

// WriteWideString in TXLSStream don't writes the unicode prefix byte on
// empty strings. This is required for notes.
procedure X_WriteWideString(Value: AxUCString);
begin
  Stream.WWord(Length(Value));
  if Value <> '' then begin
    Stream.WByte(1);
    Stream.Write(Pointer(Value)^,Length(Value) * 2);
  end
  else
    Stream.WByte(0);
end;

procedure WriteRoot;
var
  i,FileSz: integer;
begin
        // DGCONTAINER
  Sz := SizeOf(TMSOHeader) +
        // DG
        SizeOf(TMSOHeader) + SizeOf(TMSORecDG) +
        // REGROUPITEMS
        // Conditional, added below.
        // SPGRCONTAINER
        SizeOf(TMSOHeader) +
        // SPCONTAINER
        SizeOf(TMSOHeader) +
        // SPGR
        SizeOf(TMSOHeader) + SizeOf(TMSORecSPGR) +
        // SP
        SizeOf(TMSOHeader) + SizeOf(TMSORecSP);
  // REGROUPITEMS
  if Length(FFileReGroupItems) > 0 then
    Inc(Sz,SizeOf(TMSOHeader) + (Length(FFileReGroupItems) * SizeOf(TMSOFileReGroupItem)));
  FileSz := Sz;
  if FGroup.Count > 0 then begin
    if FGroup[0] is TShapeGroup then
      Inc(FileSz,TShapeGroup(FGroup[0]).PrivateSize)
    else
      Inc(FileSz,FGroup[0].Size);
  end;
  if (FGroup.Count > 0) and (FGroup[0].FInterface is TShapeOutsideMsoBaseText) then
    Stream.WriteHeader(BIFFRECID_MSODRAWING,FileSz - SizeOf(TMSOHeader))
  else
    Stream.WriteHeader(BIFFRECID_MSODRAWING,FileSz);

  for i := 0 to FGroup.Count - 1 do
    Inc(Sz,FGroup[i].Size);

  Dec(Sz,SizeOf(TMSOHeader));
  if FSolverContainer <> Nil then
    Stream.WriteMSOHeader($0F,$0000,MSO_DGCONTAINER,Sz + FSolverContainer.Size)
  else
    Stream.WriteMSOHeader($0F,$0000,MSO_DGCONTAINER,Sz);
  Dec(Sz,SizeOf(TMSOHeader));

  Stream.WriteMSOHeader($00,FDgId,MSO_DG,SizeOf(TMSORecDG));

  FDG.ShapeCount_ := 1 + ShapeCount;
  Stream.Write(FDG,SizeOf(TMSORecDG));

  Dec(Sz,SizeOf(TMSOHeader) + SizeOf(TMSORecDG));

  if Length(FFileReGroupItems) > 0 then begin
    Stream.WriteMSOHeader($00,Length(FFileReGroupItems),MSO_REGROUPITEMS,Length(FFileReGroupItems) * SizeOf(TMSOFileReGroupItem));
    for i := 0 to High(FFileReGroupItems) do
      Stream.Write(FFileReGroupItems[i],SizeOf(TMSOFileReGroupItem));
    Dec(Sz,SizeOf(TMSOHeader) + (Length(FFileReGroupItems) * SizeOf(TMSOFileReGroupItem)));
  end;

  Stream.WriteMSOHeader($0F,$0000,MSO_SPGRCONTAINER,Sz);

  Sz := SizeOf(TMSOHeader) + SizeOf(TMSORecSPGR) +
        SizeOf(TMSOHeader) + SizeOf(TMSORecSP);
  Stream.WriteMSOHeader($0F,$0000,MSO_SPCONTAINER,Sz);

  Stream.WriteMSOHeader($01,$0000,MSO_SPGR,SizeOf(TMSORecSPGR));
  Stream.Write(FGroup.FSPGR,SizeOf(TMSORecSPGR));

  Stream.WriteMSOHeader($02,$0000,MSO_SP,SizeOf(TMSORecSP));
  Stream.Write(FGroup.SpId,SizeOf(FGroup.SpId));
  Stream.Write(FGroup.Options,SizeOf(FGroup.Options));
end;

begin
  WriteRoot;
  for i := 0 to FGroup.Count - 1 do
    FGroup[i].WriteToStream(Stream,PBuf,i > 0);

  if FSolverContainer <> Nil then begin
    Sz := FSolverContainer.Size;
    Stream.WriteHeader(BIFFRECID_MSODRAWING,Sz);
    FSolverContainer.Write(Stream);
  end;
  for i := 0 to FGroup.Count - 1 do begin
    if FGroup[i].FInterface is TShapeOutsideMsoNote then begin
      Stream.WriteHeader(BIFFRECID_NOTE,TRecNOTE_FIXEDLEN + 1 + 2 + Length(TShapeOutsideMsoNote(FGroup[i].FInterface).Author) * 2);
      Stream.Write(TShapeOutsideMsoNote(FGroup[i].FInterface).CellRow,2);
      Stream.Write(TShapeOutsideMsoNote(FGroup[i].FInterface).CellCol,2);
      Stream.Write(TShapeOutsideMsoNote(FGroup[i].FInterface).Options,2);
      Stream.Write(PObjCMO(@FGroup[i].FOBJ[0].Data).ObjId,2);
      X_WriteWideString(TShapeOutsideMsoNote(FGroup[i].FInterface).Author);
    end;
  end;
end;

function TEscherDrawing.ShapeCount: integer;
begin
  Result := FGroup.Count;
end;

procedure TEscherDrawing.AssignBlipIds(Blips: TList);

procedure SetId(Shape: TShape);
var
  i,j,Id: integer;
begin
  if Shape is TShapeGroup then begin
    for i := 0 to TShapeGroup(Shape).Count - 1 do
      SetId(TShapeGroup(Shape)[i]);
  end
  else begin
    Id := Shape.FOPT.BlipId;
    if Id >= 0 then begin
      j := Blips.IndexOf(Pointer(Id - 1));
      if j < 0 then
        raise XLSRWException.CreateFmt('Can not find Blip Id #%d',[Id]);
      Shape.FOPT.SetBlipId(j + 1);
    end;
  end;
end;

begin
  SetId(FGroup);
end;

function TEscherDrawing.AddTextBox: TShapeClientAnchor;
var
  PBuf: PByteArray;
  Header: TBIFFHeader;
begin
  SetRootData;

  Result := TShapeClientAnchor.Create;
  Result.FShapeType := msosptTextBox;
  Result.FInterface := TShapeOutsideMsoBaseText.Create(GetObjId,mstTextBox);

  TShapeOutsideMsoBaseText(Result.FInterface).SetDefaultTxo;

  Result.FSpId := GetSpId;
  Result.FOptions := SpOptHaveAnchor + SpOptHaveSpt;

  Result.FAnchor.Col1 := 1;
  Result.FAnchor.Row1 := 1;
  Result.FAnchor.Col2 := 3;
  Result.FAnchor.Row2 := 5;

  Result.FOPT.AddValue(msooptLTxid,GetTxId);
  Result.FOPT.AddValue(msooptTxdir,2);
  Result.FOPT.AddValue(msooptFFitTextToShape,$00080008);
  Result.FOPT.AddValue(msooptFillColor,$08000041);
  Result.FOPT.AddValue(msooptLineColor,$08000040);
  Result.FOPT.FItemCount := Result.FOPT.Count;

  FGroup.Add(Result);
  GetMem(PBuf,MAXRECSZ_97);
  try
    Header.RecID := OBJREC_CMO;
    Header.Length := SizeOf(TObjCMO);
    FillChar(PBuf^,SizeOf(TObjCMO),#0);
    PObjCMO(PBuf).ObjType := OBJTYPE_TEXT;
    PObjCMO(PBuf).ObjId := GetObjId;
    PObjCMO(PBuf).Options := $6011;
    Result.FOBJ.AddRec(Header,PBuf);

    Header.RecID := OBJREC_END;
    Header.Length := 0;
    Result.FOBJ.AddRec(Header,PBuf);
  finally
    FreeMem(PBuf);
  end;
end;

function TEscherDrawing.AddBasic: TShapeClientAnchor;
var
  PBuf: PByteArray;
  Header: TBIFFHeader;
begin
  SetRootData;

  Result := TShapeClientAnchor.Create;
  Result.FShapeType := msosptLine;
  Result.FInterface := Nil;
  Result.FSpId := GetSpId;
  Result.FOptions := SpOptHaveAnchor + SpOptHaveSpt {+ SpOptFlipH + SpOptFlipV};

  Result.FAnchor.Col1 := 1;
  Result.FAnchor.Row1 := 1;
  Result.FAnchor.Col2 := 3;
  Result.FAnchor.Row2 := 5;

  Result.FOPT.AddValue(msooptFFitTextToShape,$00080008);
  Result.FOPT.AddValue(msooptShapePath,msoshapeComplex);
  Result.FOPT.AddValue(msooptFFillOk,$00010000);
  Result.FOPT.AddValue(msooptFNoFillHitTest,$00100000);
  Result.FOPT.AddValue(msooptLineColor,$08000040);
  Result.FOPT.AddValue(msooptFNoLineDrawDash,$00100010);

  Result.FOPT.FItemCount := Result.FOPT.Count;

  FGroup.Add(Result);
  GetMem(PBuf,MAXRECSZ_97);
  try
    Header.RecID := OBJREC_CMO;
    Header.Length := SizeOf(TObjCMO);
    FillChar(PBuf^,SizeOf(TObjCMO),#0);
    PObjCMO(PBuf).ObjType := OBJTYPE_LINE;
    PObjCMO(PBuf).ObjId := GetObjId;
    PObjCMO(PBuf).Options := $6011;
    Result.FOBJ.AddRec(Header,PBuf);

    Header.RecID := OBJREC_END;
    Header.Length := 0;
    Result.FOBJ.AddRec(Header,PBuf);
  finally
    FreeMem(PBuf);
  end;
end;

function TEscherDrawing.MaxSpId: integer;
begin
  Result := FDG.CurSpId;
end;

procedure TEscherDrawing.SetRootData;
begin
  if ShapeCount <= 0 then begin
    FParent.AddDrawing(Self,FDgId,FDG.CurSpId);
    FGroup.FOptions := SpOptGroup + SpOptPatriarch;
    FGroup.FSpId := FDG.CurSpId;
    Inc(FDG.CurSpId);
  end
  else
    FParent.UpdateDrawing(FDgId);
end;

function TEscherDrawing.GetSpId: integer;
begin
  Result := FDG.CurSpId;
  Inc(FDG.CurSpId);
end;

procedure TEscherDrawing.DeleteShape(SpId: integer);
begin
  FGroup.DeleteBySpId(SpId);
  if FGroup.Count <= 0 then begin
    FParent.DeleteDrawing(FDgId);
    FDgId := 0;
  end
  else
    FParent.UpdateDrawing(FDgId);
end;

function TEscherDrawing.AddListBox: TShapeClientAnchor;
begin
  Result := AddListComboBox;
  Result.FInterface := TShapeControlListBox.Create(GetObjId,mstCtrlListBox);
end;

function TEscherDrawing.AddListComboBox: TShapeClientAnchor;
begin
  SetRootData;

  Result := TShapeClientAnchor.Create;
  Result.FShapeType := msosptHostControl;
  Result.FSpId := GetSpId;
  Result.FOptions := SpOptHaveAnchor + SpOptHaveSpt;

  Result.FAnchor.Col1 := 3;
  Result.FAnchor.Row1 := 1;
  Result.FAnchor.Col2 := 5;
  Result.FAnchor.Row2 := 5;

  Result.FOPT.AddValue(msooptFLockAgainstGrouping,$01040104);
  Result.FOPT.AddValue(msooptFFitTextToShape,$00080008);
  Result.FOPT.AddValue(msooptLineColor,$08000040);
  Result.FOPT.AddValue(msooptFNoLineDrawDash,$00100010);
  Result.FOPT.FItemCount := Result.FOPT.Count;

  FGroup.Add(Result);
end;

function TEscherDrawing.AddComboBox: TShapeClientAnchor;
begin
  Result := AddListComboBox;
  Result.FInterface := TShapeControlComboBox.Create(GetObjId,mstCtrlComboBox);
  TShapeControlListBox(Result.FInterface).IsComboBox := True;
  Result.FAnchor.Row2 := Result.FAnchor.Row1 + 1;
end;

function TEscherDrawing.AddComboBoxAutoFilter: TShapeClientAnchor;
begin
  Result := AddComboBox;
  TShapeControlComboBox(Result.FInterface).IsAutoFilter := True;
end;

function TEscherDrawing.AddButton: TShapeClientAnchor;
begin
  SetRootData;

  Result := TShapeClientAnchor.Create;
  Result.FShapeType := msosptHostControl;
  Result.FInterface := TShapeControlButton.Create(GetObjId,Self.FInternalNames);
  Result.FSpId := GetSpId;
  Result.FOptions := SpOptHaveAnchor + SpOptHaveSpt;

  TShapeOutsideMsoBaseText(Result.FInterface).SetDefaultTxo;
  TShapeOutsideMsoBaseText(Result.FInterface).Text := 'Button';

  Result.FAnchor.Col1 := 3;
  Result.FAnchor.Row1 := 1;
  Result.FAnchor.Col2 := 4;
  Result.FAnchor.Row2 := 2;

  Result.FOPT.AddValue(msooptFLockAgainstGrouping,$01040104);
  Result.FOPT.AddValue(msooptLTxId,GetTxId);
  Result.FOPT.AddValue(msooptWrapText,$00000001);
  Result.FOPT.AddValue(msooptTxDir,$00000002);
  Result.FOPT.AddValue(msooptFFitTextToShape,$001A0008);
  Result.FOPT.AddValue(msooptFillColor,$08000043);
  Result.FOPT.AddValue(msooptFillBackColor,$08000043);
  Result.FOPT.AddValue(msooptFNoFillHitTest,$00110011);
  Result.FOPT.AddValue(msooptLineColor,$08000040);
  Result.FOPT.AddValue(msooptFPrint,$00080008);
  Result.FOPT.FItemCount := Result.FOPT.Count;

  FGroup.Add(Result);
end;

function TEscherDrawing.AddNote: TShapeClientAnchor;
var
  PBuf: PByteArray;
  Header: TBIFFHeader;
begin
  SetRootData;

  Result := TShapeClientAnchor.Create;
  Result.FShapeType := msosptTextBox;
  Result.FInterface := TShapeOutsideMsoNote.Create(GetObjId,mstNote);

  TShapeOutsideMsoNote(Result.FInterface).SetDefaultTxo;
  TShapeOutsideMsoNote(Result.FInterface).TXORuns[0].FontIndex := FFonts.DefaultFont.Index;

  Result.FSpId := GetSpId;
  Result.FOptions := SpOptHaveAnchor + SpOptHaveSpt;

  Result.FAnchor.Col1 := 1;
  Result.FAnchor.Row1 := 1;
  Result.FAnchor.Col2 := 3;
  Result.FAnchor.Row2 := 5;

  Result.FOPT.AddValue(msooptLTxid,GetTxId);
  Result.FOPT.AddValue(msooptTxdir,2);
  Result.FOPT.AddValue(msooptFFitTextToShape,$00080008);
  Result.FOPT.AddValue($0158,$00000000);
  Result.FOPT.AddValue(msooptFillColor,$08000050);
  Result.FOPT.AddValue(msooptFillBackColor,$08000050);
  Result.FOPT.AddValue(msooptFillCrMod,$100000F4);
  Result.FOPT.AddValue(msooptFNoFillHitTest,$00100010);
  Result.FOPT.AddValue(msooptLineCrMod,$100000F4);
  Result.FOPT.AddValue(msooptShadowColor,$00000000);
  Result.FOPT.AddValue(msooptShadowCrMod,$100000F4);
  Result.FOPT.AddValue(msooptShadowObscured,$00030003);
  Result.FOPT.AddValue(msooptFPrint,$00020002);
  Result.FOPT.FItemCount := Result.FOPT.Count;

  FGroup.Add(Result);
  GetMem(PBuf,MAXRECSZ_97);
  try
    Header.RecID := OBJREC_CMO;
    Header.Length := SizeOf(TObjCMO);
    FillChar(PBuf^,SizeOf(TObjCMO),#0);
    PObjCMO(PBuf).ObjType := OBJTYPE_COMMENT;
    PObjCMO(PBuf).ObjId := GetObjId;
    PObjCMO(PBuf).Options := $4011;
    Result.FOBJ.AddRec(Header,PBuf);

    Header.RecID := OBJREC_NTS;
    // Note data. Must be present for Excel 2003 and later.
    Header.Length := 22;
    FillChar(PBuf[0],22,#0);
    Result.FOBJ.AddRec(Header,PBuf);

    Header.RecID := OBJREC_END;
    Header.Length := 0;
    Result.FOBJ.AddRec(Header,PBuf);
  finally
    FreeMem(PBuf);
  end;
end;

function TEscherDrawing.AddAutoShape: TShapeClientAnchor;
var
  PBuf: PByteArray;
  Header: TBIFFHeader;
begin
  SetRootData;

  Result := TShapeClientAnchor.Create;
  Result.FShapeType := msosptSun;
  Result.FInterface := Nil;
  Result.FSpId := GetSpId;
  Result.FOptions := SpOptHaveAnchor + SpOptHaveSpt {+ SpOptFlipH + SpOptFlipV};

  Result.FAnchor.Col1 := 1;
  Result.FAnchor.Row1 := 1;
  Result.FAnchor.Col2 := 2;
  Result.FAnchor.Row2 := 5;

  Result.FOPT.AddValue(msooptFFitTextToShape,$00080008);
  Result.FOPT.AddValue(msooptFillColor,$08000041);
  Result.FOPT.AddValue(msooptLineColor,$08000040);
  Result.FOPT.FItemCount := Result.FOPT.Count;

  FGroup.Add(Result);
  GetMem(PBuf,MAXRECSZ_97);
  try
    Header.RecID := OBJREC_CMO;
    Header.Length := SizeOf(TObjCMO);
    FillChar(PBuf^,SizeOf(TObjCMO),#0);
    PObjCMO(PBuf).ObjType := OBJTYPE_MICROSOFT_OFFICE_DRAWING;
    PObjCMO(PBuf).ObjId := GetObjId;
    PObjCMO(PBuf).Options := $6011;
    Result.FOBJ.AddRec(Header,PBuf);

    Header.RecID := OBJREC_END;
    Header.Length := 0;
    Result.FOBJ.AddRec(Header,PBuf);
  finally
    FreeMem(PBuf);
  end;
end;

function TEscherDrawing.AddPicture: TShapeClientAnchor;
var
  PBuf: PByteArray;
  Header: TBIFFHeader;
begin
  SetRootData;

  Result := TShapeClientAnchor.Create;
  Result.FShapeType := msosptPictureFrame;
  Result.FInterface := Nil;
  Result.FSpId := GetSpId;
  Result.FOptions := SpOptHaveAnchor + SpOptHaveSpt {+ SpOptFlipH + SpOptFlipV};

  Result.FAnchor.Col1 := 1;
  Result.FAnchor.Row1 := 1;
  Result.FAnchor.Col2 := 3;
  Result.FAnchor.Row2 := 7;

//  Result.FOPT.AddValue(msooptFLockAgainstGrouping,$00800080);
  Result.FOPT.AddValue(msooptPib,$00000000,True);
  Result.FOPT.AddString(msooptPibName,'[???]');
//  Result.FOPT.AddValue(msooptFBackground,$00100010);
  Result.FOPT.FItemCount := Result.FOPT.Count;

  FGroup.Add(Result);
  GetMem(PBuf,MAXRECSZ_97);
  try
    Header.RecID := OBJREC_CMO;
    Header.Length := SizeOf(TObjCMO);
    FillChar(PBuf^,SizeOf(TObjCMO),#0);
    PObjCMO(PBuf).ObjType := OBJTYPE_PICTURE;
    PObjCMO(PBuf).ObjId := GetObjId;
    PObjCMO(PBuf).Options := $6011;
    Result.FOBJ.AddRec(Header,PBuf);

    Header.RecID := OBJREC_CF;
    Header.Length := 2;
    PWordArray(PBuf)[0] := $FFFF;
    Result.FOBJ.AddRec(Header,PBuf);

    Header.RecID := OBJREC_PIOGRBIT;
    Header.Length := 2;
    PWordArray(PBuf)[0] := $0000;
    Result.FOBJ.AddRec(Header,PBuf);

    Header.RecID := OBJREC_END;
    Header.Length := 0;
    Result.FOBJ.AddRec(Header,PBuf);
  finally
    FreeMem(PBuf);
  end;
end;

function TEscherDrawing.AddRadioButton: TShapeClientAnchor;
begin
  SetRootData;

  Result := TShapeClientAnchor.Create;
  Result.FShapeType := msosptHostControl;
  Result.FInterface := TShapeControlRadioButton.Create(GetObjId,Self.FInternalNames);
  Result.FSpId := GetSpId;
  Result.FOptions := SpOptHaveAnchor + SpOptHaveSpt;

  TShapeOutsideMsoBaseText(Result.FInterface).SetDefaultTxo;
  TShapeOutsideMsoBaseText(Result.FInterface).Text := 'RadioButton';

  Result.FAnchor.Col1 := 3;
  Result.FAnchor.Row1 := 1;
  Result.FAnchor.Col2 := 4;
  Result.FAnchor.Row2 := 2;

  Result.FOPT.AddValue(msooptFLockAgainstGrouping,$01000100);
  Result.FOPT.AddValue(msooptLTxId,GetTxId);
  Result.FOPT.AddValue(msooptWrapText,$00000001);
  Result.FOPT.AddValue(msooptTxDir,$00000002);
  Result.FOPT.AddValue(msooptFFitTextToShape,$001A0008);
  Result.FOPT.AddValue(msooptFFillOk,$00290029);
  Result.FOPT.AddValue(msooptFillColor,$08000041);
//  Result.FOPT.AddValue(msooptFillBackColor,$08000043);
  Result.FOPT.AddValue(msooptFNoFillHitTest,$00100000);
  Result.FOPT.AddValue(msooptLineColor,$08000040);
  Result.FOPT.AddValue(msooptFNoLineDrawDash,$00080000);
  Result.FOPT.FItemCount := Result.FOPT.Count;

  FGroup.Add(Result);
end;

function TEscherDrawing.AddChart: TShapeClientAnchor;
var
  PBuf: PByteArray;
  Header: TBIFFHeader;
begin
  SetRootData;

  Result := TShapeClientAnchor.Create;
  Result.FShapeType := msosptHostControl;
  Result.FInterface := TShapeOutsideMsoChart.Create(GetObjId,FParent,FFonts);
  Result.FSpId := GetSpId;
  Result.FOptions := SpOptHaveAnchor + SpOptHaveSpt;

  Result.FAnchor.Col1 := 2;
  Result.FAnchor.Row1 := 1;
  Result.FAnchor.Col2 := 8;
  Result.FAnchor.Row2 := 16;

  Result.FOPT.AddValue(msooptFLockAgainstGrouping,$01040104);
  Result.FOPT.AddValue(msooptFFitTextToShape,$00080008);
  Result.FOPT.AddValue(msooptFillColor,$0800004E);
  Result.FOPT.AddValue(msooptFillBackColor,$0800004D);
  Result.FOPT.AddValue(msooptFNoFillHitTest,$00100010);
  Result.FOPT.AddValue(msooptLineColor,$0800004D);
  Result.FOPT.AddValue(msooptFNoLineDrawDash,$00100010);
  Result.FOPT.AddValue(msooptShadowObscured,$00020000);
  Result.FOPT.FItemCount := Result.FOPT.Count;

  FGroup.Add(Result);

  GetMem(PBuf,MAXRECSZ_97);
  try
    Header.RecID := OBJREC_CMO;
    Header.Length := SizeOf(TObjCMO);
    FillChar(PBuf^,SizeOf(TObjCMO),#0);
    PObjCMO(PBuf).ObjType := OBJTYPE_CHART;
    PObjCMO(PBuf).ObjId := GetObjId;
    PObjCMO(PBuf).Options := $6011;
    Result.FOBJ.AddRec(Header,PBuf);

    Header.RecID := OBJREC_END;
    Header.Length := 0;
    Result.FOBJ.AddRec(Header,PBuf);
  finally
    FreeMem(PBuf);
  end;
end;

function TEscherDrawing.AddCheckBox: TShapeClientAnchor;
begin
  SetRootData;

  Result := TShapeClientAnchor.Create;
  Result.FShapeType := msosptHostControl;
  Result.FInterface := TShapeControlCheckBox.Create(GetObjId,Self.FInternalNames);
  Result.FSpId := GetSpId;
  Result.FOptions := SpOptHaveAnchor + SpOptHaveSpt;

  TShapeOutsideMsoBaseText(Result.FInterface).SetDefaultTxo;
  TShapeOutsideMsoBaseText(Result.FInterface).Text := 'CheckBox';

  Result.FAnchor.Col1 := 3;
  Result.FAnchor.Row1 := 1;
  Result.FAnchor.Col2 := 4;
  Result.FAnchor.Row2 := 2;

  Result.FOPT.AddValue(msooptFLockAgainstGrouping,$01000100);
  Result.FOPT.AddValue(msooptLTxId,GetTxId);
  Result.FOPT.AddValue(msooptWrapText,$00000001);
  Result.FOPT.AddValue(msooptTxDir,$00000002);
  Result.FOPT.AddValue(msooptFFitTextToShape,$001A0008);
  Result.FOPT.AddValue(msooptFFillOk,$00290029);
  Result.FOPT.AddValue(msooptFillColor,$08000041);
//  Result.FOPT.AddValue(msooptFillBackColor,$08000043);
  Result.FOPT.AddValue(msooptFNoFillHitTest,$00100000);
  Result.FOPT.AddValue(msooptLineColor,$08000040);
  Result.FOPT.AddValue(msooptFNoLineDrawDash,$00080000);
  Result.FOPT.FItemCount := Result.FOPT.Count;

  FGroup.Add(Result);
end;

procedure TEscherDrawing.SetBlipRefCount;

procedure SetRefCount(Shape: TShape);
var
  i,Id: integer;
begin
  if Shape is TShapeGroup then begin
    for i := 0 to TShapeGroup(Shape).Count - 1 do
      SetRefCount(TShapeGroup(Shape)[i]);
  end
  else if Shape.FShapeType = msosptPictureFrame then begin
    Id := Shape.FOPT.FindValue(msooptPib);
    if Id > TMSOPictures(FParent).Count then
      raise XLSRWException.Create('Unassigned picture in file.');
    // This is wrong, but will ensure that all added pictures are valid.
    if (Id > 0) and (TMSOPictures(FParent)[Id - 1].FRefCount = 0) then
//      TMSOPictures(FParent)[Id - 1].FRefCount := TMSOPictures(FParent)[Id - 1].FRefCount + 1;
      TMSOPictures(FParent)[Id - 1].FRefCount := 1;
  end;
end;

begin
  SetRefCount(FGroup);
end;

function TEscherDrawing.GetObjId: integer;
begin
  Inc(FMaxObjId);
  Result := FMaxObjId;
end;

procedure TEscherDrawing.AddEmpty;
begin
  SetRootData;
end;

procedure TEscherDrawing.Move(Index: integer; DestCol, DestRow: longword);
var
  dCol,dRow: longword;
  Shape: TShapeClientAnchor;
begin
  if FGroup[Index] is TShapeClientAnchor then
    Shape := TShapeClientAnchor(FGroup[Index])
  else if FGroup[Index] is TShapeChildAnchor then
    raise XLSRWException.Create('Not implemented: Move shape child anchor')
  else
    raise XLSRWException.Create('Dont know how to Move this shape');
  dCol := Shape.Anchor.Col2 - Shape.Anchor.Col1;
  dRow := Shape.Anchor.Row2 - Shape.Anchor.Row1;
  Shape.Anchor.Col1 := DestCol;
  Shape.Anchor.Col2 := dCol + DestCol;
  Shape.Anchor.Row1 := DestRow;
  Shape.Anchor.Row2 := dRow + DestRow;
  if (Shape.FInterface <> Nil) and (Shape.FInterface is TShapeOutsideMsoNote) then begin
    TShapeOutsideMsoNote(Shape.FInterface).CellCol := DestCol;
    TShapeOutsideMsoNote(Shape.FInterface).CellRow := DestRow;
  end;
end;

procedure TEscherDrawing.Copy(Index: integer; DestCol, DestRow: word);
var
  Shape: TShape;
begin
  if FGroup[Index] is TShapeGroupClientAnchor then
    Shape := TShapeGroupClientAnchor.Create(Self)
  else if FGroup[Index] is TShapeGroupChildAnchor then
    Shape := TShapeGroupChildAnchor.Create(Self)
  else if FGroup[Index] is TShapeClientAnchor then
    Shape := TShapeClientAnchor.Create
  else if FGroup[Index] is TShapeChildAnchor then
    Shape := TShapeChildAnchor.Create
  else
    raise XLSRWException.Create('Dont know how to Copy this shape');
  FGroup[Index].Assign(Shape);
  Shape.FSpId := GetSpId;
  if Shape.FInterface <> Nil then begin
    Shape.FInterface.ObjId := GetObjId;
    PObjCMO(@Shape.FOBJ[0].Data).ObjId := Shape.FInterface.ObjId;
  end
  else
    PObjCMO(@Shape.FOBJ[0].Data).ObjId := GetObjId;
  FGroup.Add(Shape);
  if not(FGroup[Index] is TShapeChildAnchor) then
    Move(FGroup.Count - 1,DestCol,DestRow);
end;

procedure TEscherDrawing.Copy(Col1, Row1, Col2, Row2, DestCol, DestRow: word);
var
  i: integer;
  Shp: TShapeClientAnchor;
  Note: TShapeOutsideMsoNote;
begin
  for i := 0 to FGroup.Count - 1 do begin
    if (FGroup[i] is TShapeClientAnchor) then begin
      Shp := TShapeClientAnchor(FGroup[i]);
      if (Shp.FInterface <> Nil) and (Shp.FInterface is TShapeOutsideMsoNote) then begin
        Note := TShapeOutsideMsoNote(Shp.FInterface);
        if (Note.CellCol >= Col1) and (Note.CellRow >= Row1) and (Note.CellCol <= Col2) and (Note.CellRow <= Row2) then
          Copy(i,(Note.CellCol - Col1) + DestCol,(Note.CellRow - Row1) + DestRow);
      end
      else if (Shp.Anchor.Col1 >= Col1) and (Shp.Anchor.Row1 >= Row1) and (Shp.Anchor.Col1 <= Col2) and (Shp.Anchor.Row1 <= Row2) then begin
{$ifdef AXOLOT_DEBUG}
{$message WARN 'Charts not copied. Causes XLSRWException}
{$endif}
        if (Shp.FInterface <> Nil) and (Shp.FInterface.ShapeType = mstChart) then
          Continue;
        Copy(i,(Shp.Anchor.Col1 - Col1) + DestCol,(Shp.Anchor.Row1 - Row1) + DestRow);
      end
    end;
  end;
end;

procedure TEscherDrawing.CopyList(List: TList; Col1, Row1, Col2, Row2: integer);
var
  i: integer;
  Shp: TShapeClientAnchor;
  Note: TShapeOutsideMsoNote;
begin
  for i := 0 to FGroup.Count - 1 do begin
    if (FGroup[i] is TShapeClientAnchor) then begin
      Shp := TShapeClientAnchor(FGroup[i]);
      if (Shp.FInterface <> Nil) and (Shp.FInterface is TShapeOutsideMsoNote) then begin
        Note := TShapeOutsideMsoNote(Shp.FInterface);
        if (Note.CellCol >= Col1) and (Note.CellRow >= Row1) and (Note.CellCol <= Col2) and (Note.CellRow <= Row2) then
          List.Add(Shp);
      end
      else if (Shp.Anchor.Col1 >= Longword(Col1)) and (Shp.Anchor.Row1 >= Longword(Row1)) and (Shp.Anchor.Col1 <= Longword(Col2)) and (Shp.Anchor.Row1 <= Longword(Row2)) then begin
        List.Add(Shp);
      end
    end;
  end;
end;

procedure TEscherDrawing.InsertList(List: TList; DestCol,DestRow: integer);
var
  i: integer;
  Shape: TShape;
begin
  for i := 0 to List.Count - 1 do begin
    if TObject(List[i]) is TShapeGroupClientAnchor then
      Shape := TShapeGroupClientAnchor.Create(Self)
    else if TObject(List[i]) is TShapeGroupChildAnchor then
      Shape := TShapeGroupChildAnchor.Create(Self)
    else if TObject(List[i]) is TShapeClientAnchor then
      Shape := TShapeClientAnchor.Create
    else if TObject(List[i]) is TShapeChildAnchor then
      Shape := TShapeChildAnchor.Create
    else
      raise XLSRWException.Create('Dont know how to Copy this shape');
    if FGroup.Count <= 0 then
      SetRootData;

// Note_JKQWNH (Search for this).
    if TShape(List[i]).FInterface is TShapeOutsideMsoChart then
      Shape.FInterface := TShapeOutsideMsoChart.Create(0,FParent,FFonts);
    TShape(List[i]).Assign(Shape);

    Shape.FSpId := GetSpId;
    PObjCMO(@Shape.FOBJ[0].Data).ObjId := GetObjId;
    if Shape.FInterface <> Nil then
      Shape.FInterface.ObjId := PObjCMO(@Shape.FOBJ[0].Data).ObjId;
    FGroup.Add(Shape);
    if Assigned(FFileReadShapeEvent) then
      FFileReadShapeEvent(Self,Shape);
{
    if not(Shape is TShapeChildAnchor) then
      Move(FGroup.Count - 1,TShapeClientAnchor(Shape).Col1 + DestCol,TShapeClientAnchor(Shape).Row1 + DestRow);
}
  end;
end;

procedure TEscherDrawing.DeleteList(List: TList);
var
  i: integer;
begin
  for i := 0 to List.Count - 1 do
    FGroup.DeleteBySpId(TShape(List[i]).SpId);
end;

procedure TEscherDrawing.Move(Col1, Row1, Col2, Row2, DestCol, DestRow: longword);
var
  i: integer;
  C,R: word;
  Shp: TShapeClientAnchor;
begin
  for i := 0 to FGroup.Count - 1 do begin
    if (FGroup[i] is TShapeClientAnchor) then begin
      Shp := TShapeClientAnchor(FGroup[i]);
      if (Shp.FInterface <> Nil) and (Shp.FInterface is TShapeOutsideMsoNote) then begin
        C := TShapeOutsideMsoNote(Shp.FInterface).CellCol;
        R := TShapeOutsideMsoNote(Shp.FInterface).CellRow;
      end
      else begin
        C := TShapeClientAnchor(FGroup[i]).Anchor.Col1;
        R := TShapeClientAnchor(FGroup[i]).Anchor.Row1;
      end;
      if (C >= Col1) and (R >= Row1) and (C <= Col2) and (R <= Row2) then begin
        Move(i,(C - Col1) + DestCol,(R - Row1) + DestRow);
        if (Shp.FInterface <> Nil) and (Shp.FInterface is TShapeOutsideMsoNote) then begin
          TShapeOutsideMsoNote(Shp.FInterface).CellCol := (C - Col1) + DestCol;
          TShapeOutsideMsoNote(Shp.FInterface).CellRow := (R - Row1) + DestRow;
        end
      end;
    end;
  end;
end;

procedure TEscherDrawing.Delete(Col1, Row1, Col2, Row2: longword);
begin
  // TODO Not implemented
end;

{ TMSORecords }

procedure TMSORecords.Add(Header: TMSOHeader; Data: PByteArray);
var
  P,P2: PByteArray;
begin
  if (Header.VerInst and $0F) = $0F then begin
    GetMem(P,SizeOf(TMSOHeader));
    System.Move(Header,P^,SizeOf(TMSOHeader));
  end
  else begin
    GetMem(P,Header.Length + SizeOf(TMSOHeader));
    System.Move(Header,P^,SizeOf(TMSOHeader));
    P2 := PByteArray(NativeInt(P) + SizeOf(TMSOHeader));
    System.Move(Data^,P2^,Header.Length);
  end;
  inherited Add(P);
end;

procedure TMSORecords.Assign(Records: TMSORecords);
var
  i: integer;
  Header: TMSOHeader;
begin
  for i := 0 to Count - 1 do begin
    Header.VerInst := Version[i] + (Instance[i] shl 4);
    Header.FBT := FBT[i];
    Header.Length := Length[i];
    if (Header.VerInst and $0F) = $0F then
      Records.Add(Header,Nil)
    else
      Records.Add(Header,Data[i]);
  end;
end;

procedure TMSORecords.Clear;
var
  i: integer;
begin
  for i := 0 to Count - 1 do
    FreeMem(inherited Items[i]);
  inherited Clear;
end;

destructor TMSORecords.Destroy;
begin
  Clear;
  inherited;
end;

function TMSORecords.GetData(Index: integer): PByteArray;
begin
  Result := @PMSORecord(inherited Items[Index]).Data;
end;

function TMSORecords.GetFBT(Index: integer): word;
begin
  Result := PMSORecord(inherited Items[Index]).FBT;
end;

function TMSORecords.GetInstance(Index: integer): word;
begin
  Result := (PMSORecord(inherited Items[Index]).VerInst and $FFF0) shr 4;
end;

function TMSORecords.GetItems(Index: integer): PMSORecord;
begin
  Result := PMSORecord(inherited Items[Index]);
end;

function TMSORecords.GetLength(Index: integer): longword;
begin
  Result := PMSORecord(inherited Items[Index]).Length;
end;

function TMSORecords.GetVersion(Index: integer): byte;
begin
  Result := PMSORecord(inherited Items[Index]).VerInst and $000F;
end;

function TMSORecords.Size: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do begin
    Inc(Result,SizeOf(TMSOHeader));
    if Version[i] < $0F then
      Inc(Result,Items[i].Length)
  end;
end;

procedure TMSORecords.Write(Stream: TXLSStream);
var
  i: integer;
begin
  for i := 0 to Count - 1 do begin
    Stream.WriteMSOHeader(Version[i],Instance[i],FBT[i],Length[i]);
    if Version[i] < $0F then
      Stream.Write(Data[i][0],Length[i]);
  end;
end;

{ TEscherGroup }

procedure TEscherGroup.AddDrawing(Drawing: TEscherDrawing; var DgId: word; var SpIdBlock: longword);
const
  DefSplitMenuColors: array[0..15] of byte = ($0D,$00,$00,$08,$0C,$00,$00,$08,$17,$00,$00,$08,$F7,$00,$00,$10);
var
  Data: TDGGData;
begin
  if FDGG.Count <= 0 then begin
    // This is a minimum, but not correct.
    // OPT values depends of shapes in drawings and blip's
    FOPT.AddValue(msooptFFitTextToShape,$00080008);
    FOPT.AddValue(msooptFillColor,$08000041);
    FOPT.AddValue(msooptLineColor,$08000040);
    FOPT.FItemCount := FOPT.Count;
    SetLength(FSplitMenuColors,4);
    System.Move(DefSplitMenuColors,FSplitMenuColors[0],SizeOf(DefSplitMenuColors));
    // Root shape
    FDGG.FSavedShapes := 1;
  end;

  Data := TDGGData.Create;
  Data.NewDrawing := True;
  Data.FFIDCL.DgId := FDGG.Count + 1;
  Data.FFIDCL.SpIdUsed := 0;
  Data.FDrawing := Drawing;
  DgId := Data.FFIDCL.DgId;
  SpIdBlock := GetSpIdBlock;
  FDGG.Add(Data);
  Inc(FDGG.FSavedDrawings);
end;

procedure TEscherGroup.AssignDrawing(Drawing: TEscherDrawing);
var
  i: integer;
  Data: TDGGData;
begin
  for i := 0 to FDGG.Count - 1 do begin
    if FDGG.Items[i].FFIDCL.DgId = Drawing.FDgId then begin
      FDGG.Items[i].FDrawing := Drawing;
      FDGG.Items[i].NewDrawing := False;
      Exit;
    end;
  end;
// ATTN: There may be drawings that not have an entry in the IdClust list.
// Add it.
  Data := TDGGData.Create;
  Data.FDrawing := Drawing;
  Data.FFIDCL.DgId := Drawing.FDgId;
  Data.FFIDCL.SpIdUsed := Drawing.ShapeCount;
  FDGG.Add(Data);
  if FDGG.Count > Integer(FDGG.FSavedDrawings) then
    FDGG.FSavedDrawings := FDGG.Count;
end;

procedure TEscherGroup.UpdateDrawing(DgId: integer);
var
  i: integer;
begin
  for i := 0 to FDGG.Count - 1 do begin
    if Integer(FDGG.Items[i].FFIDCL.DgId) = DgId then begin
      FDGG.Items[i].NewDrawing := True;
      Exit;
    end;
  end;
  raise XLSRWException.Create('Can not update FIDCL for DgId #' + IntToStr(DgId));
end;

procedure TEscherGroup.Clear;
begin
  FDGG.Clear;
  FOPT.Clear;
  SetLength(FSplitMenuColors,0);
  FUnknown.Clear;
  inherited;
end;

constructor TEscherGroup.Create;
begin
  inherited Create;
  FOPT := TOPT.Create;
  FDGG := TDGG.Create;
  FUnknown := TMSORecords.Create;
end;

destructor TEscherGroup.Destroy;
begin
  // Clear must be called before objects are free'd
  inherited;

  FDGG.Free;
  FOPT.Free;
  SetLength(FSplitMenuColors,0);
  FUnknown.Free;
end;

function TEscherGroup.GetMSOPicture(Index: integer): TMSOPicture;
begin
  Result := TMSOPicture(inherited Items[Index]);
end;

function TEscherGroup.GetSpIdBlock: integer;
begin
  Inc(FCurrSpIdBlock,SP_ID_BLOCK_SIZE);
  Result := FCurrSpIdBlock;
end;

function TEscherGroup.HasData: boolean;
begin
  Result := FOPT.Count > 0;
end;

procedure TEscherGroup.LoadFromStream(Stream: TXLSStream; PBuf: PByteArray);
var
  Length: integer;
  Header: TMSOHeader;
  Instance: word;
begin
  Stream.Read(Header,SizeOf(TMSOHeader));
  if Header.FBT <> MSO_DGGCONTAINER then
    raise XLSRWException.Create('Expected record missing: DGGContainer');

  Length := Header.Length;
  while Length > 0 do begin
    if Length < SizeOf(TMSOHeader) then
      raise XLSRWException.Create('MSOHeader split in MSODRAWINGGROUP');
    Dec(Length,Stream.Read(Header,SizeOf(TMSOHeader)));
    Instance := (Header.VerInst shr 4) and $0FFF;
    case Header.FBT of
      MSO_DGG: begin
        FDGG.LoadFromStream(Stream,Length,PBuf);
        FCurrSpIdBlock := (FDGG.FMaxSpId div SP_ID_BLOCK_SIZE) * SP_ID_BLOCK_SIZE;
      end;
      MSO_CLSID: raise XLSRWException.Create('CLSID in MSODRAWINGGROUP');
      MSO_OPT: begin
        FOPT.LoadFromStream(Stream,Header.Length,PBuf,Instance);
        Dec(Length,Header.Length);
      end;
      MSO_COLORMRU: begin
        SetLength(FColorMRU,Instance);
        Dec(Length,Stream.Read(FColorMRU[0],Instance * SizeOf(RGBQUAD)));
      end;
      MSO_SPLITMENUCOLORS: begin
        SetLength(FSplitMenuColors,Instance);
        Dec(Length,Stream.Read(FSplitMenuColors[0],Instance * SizeOf(RGBQUAD)));
      end;
      MSO_BSTORECONTAINER: begin
        ReadPictures(Stream,Header.Length,PBuf);
        Dec(Length,Header.Length);
      end;
      else begin
        if Header.Length > MAXRECSZ_97 then
          raise XLSRWException.Create('Unknown FBT in MSODRAWINGGROUP to large');
        Dec(Length,Stream.Read(PBuf^,Header.Length));
        FUnknown.Add(Header,PBuf);            
      end;
    end;
  end;
  if Length < 0 then
    raise XLSRWException.Create('Error while reading MSODRAWINGGROUP. Length < 0');
end;

procedure TEscherGroup.ReadPictures(Stream: TXLSStream; Length: integer; PBuf: PByteArray);
var
  i,Hdr: integer;
  Header: TMSOHeader;
  Pict: TMSOPicture;
  Buf: PByteArray;
begin
  i := 0;
  while Length > 0 do begin
    Hdr := Stream.PeekMSOHeader;
    if (Hdr >= MSO_BLIPSTART) and (Hdr <= MSO_BLIPEND) then begin
      Dec(Length,Stream.Read(Header,SizeOf(TMSOHeader)));
      Dec(Length,Stream.Read(PBuf^,16));
      Dec(Length,Stream.Read(PBuf^,1));
      GetMem(Buf,Header.Length);
      try
        Dec(Length,Stream.Read(Buf^,Header.Length - 17));
      finally
        FreeMem(Buf);
      end;
      if Length <= 0 then
        Break;

{
      Dec(Length,Stream.Seek(soFromCurrent,16));
      Dec(Length,Stream.Seek(soFromCurrent,1));
      Dec(Length,Stream.Seek(soFromCurrent,Header.Length - 17));
}
    end;


    Pict := TMSOPictures(Self).Add;
    if not Pict.LoadFromStream(Stream,Length,PBuf) then
      Delete(Count - 1)
    else
      Pict.FileBlipId := i;
    Inc(i);
  end;
end;

procedure TEscherGroup.WriteToStream(Stream: TXLSStream; PBuf: PByteArray);
var
  i,Size,BlipSize: integer;

function CalcBlipSize: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do begin
    // BSE
    Inc(Result,SizeOf(TMSOHeader) + SizeOf(TMSORecBSE));
    //BLIP
    Inc(Result,SizeOf(TMSOHeader));
    Inc(Result,16);
    if TMSOPicture(Items[i]).FPictType in [msoblipEMF,msoblipWMF,msoblipPICT] then
      Inc(Result,SizeOf(TMSOBSEMetafile))
    else
      Inc(Result);
    if TMSOPicture(Items[i]).FPictType in [msoblipEMF,msoblipWMF,msoblipPICT] then begin
      TMSOPicture(Items[i]).CompressPict;
      Inc(Result,TMSOPicture(Items[i]).FCompressedSize);
    end
    else
      Inc(Result,TMSOPicture(Items[i]).FPictSize);
  end;
end;

function UnknownSize: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to FUnknown.Count - 1 do begin
    if (PMSORecord(FUnknown[i]).VerInst and $000F) = $000F then
      Inc(Result,SizeOf(TMSOHeader))
    else
      Inc(Result,SizeOf(TMSOHeader) + PMSORecord(FUnknown[i]).Length);
  end;
end;

begin
  BlipSize := CalcBlipSize;

  Size := // DGG
          SizeOf(TMSOHeader) + SizeOf(TMSORecDGG) + (FDGG.Count * SizeOf(TDGGRecFIDCL)) +
          // Blip's
          BlipSize +
          // OPT
          SizeOf(TMSOHeader) + FOPT.Size +
          // SPLITMENUCOLORS
          SizeOf(TMSOHeader) + (Length(FSplitMenuColors) * SizeOf(RGBQUAD)) +
          UnknownSize;
  if Length(FColorMRU) > 0 then
    // COLORMRU
    Inc(Size,SizeOf(TMSOHeader) + (Length(FColorMRU) * SizeOf(RGBQUAD)));

  if Count > 0 then
    // BSTORECONTAINER
    Inc(Size,SizeOf(TMSOHeader));

  // Record size is written after EndCONTINUE
  Stream.WriteHeader(BIFFRECID_MSODRAWINGGROUP,0);
  Stream.WriteMSOHeader($0F,$0000,MSO_DGGCONTAINER,Size);

  FDGG.WriteToStream(Stream,PBuf);
  if Count > 0 then begin
    Stream.WriteMSOHeader($0F,Count,MSO_BSTORECONTAINER,BlipSize);
    for i := 0 to Count - 1 do
      TMSOPicture(Items[i]).WriteToStream(Stream,PBuf);
  end;
  FOPT.WriteToStream(Stream,PBuf);

  for i := 0 to FUnknown.Count - 1 do begin
    if (PMSORecord(FUnknown[i]).VerInst and $000F) = $000F then
      Stream.Write(FUnknown[i]^,SizeOf(TMSOHeader))
    else
      Stream.Write(FUnknown[i]^,SizeOf(TMSOHeader) + PMSORecord(FUnknown[i]).Length);
  end;

  if Length(FColorMRU) > 0 then begin
    Stream.WriteMSOHeader($00,Length(FColorMRU),MSO_COLORMRU,Length(FColorMRU) * SizeOf(RGBQUAD));
    Stream.Write(FColorMRU[0],Length(FColorMRU) * SizeOf(RGBQUAD));
  end;

  if Length(FSplitMenuColors) > 0 then begin
    Stream.WriteMSOHeader($00,Length(FSplitMenuColors),MSO_SPLITMENUCOLORS,Length(FSplitMenuColors) * SizeOf(RGBQUAD));
    Stream.Write(FSplitMenuColors[0],Length(FSplitMenuColors) * SizeOf(RGBQUAD));
  end;
end;

procedure TEscherGroup.DeleteDrawing(DgId: integer);
var
  i: integer;
begin
  if FDGG.Count > 0 then begin
    for i := 0 to FDGG.Count - 1 do begin
      if Integer(FDGG.Items[i].FFIDCL.DgId) = DgId then begin
        FDGG.Items[i].Free;
        FDGG.Delete(i);
        Dec(FDGG.FSavedDrawings);
        Exit;
      end;
    end;
    raise XLSRWException.Create('Can not delete FIDCL for DgId #' + IntToStr(DgId));
  end;
end;

{ TMSOPictures }

function TMSOPictures.Add: TMSOPicture;
begin
  Result := TMSOPicture.Create(Self);
  Result.FFileBlipId := Count + 1;
  inherited Add(Result);
end;

procedure TMSOPictures.GetBlipIds(List: TList);
var
  i: integer;
begin
  for i := 0 to Count - 1 do
    List.Add(Pointer(Items[i].FileBlipId));
end;

procedure TMSOPictures.ResetBlipRefCount;
var
  i: integer;
begin
  for i := 0 to Count - 1 do
    Items[i].FRefCount := 0;
end;

{ TMSOPicture }

constructor TMSOPicture.Create(AOwner: TMSOPictures);
begin
  FOwner := AOwner;
  FName := 'Picture ' + IntToStr(AOwner.Count + 1);
end;

destructor TMSOPicture.Destroy;
begin
  FreeMem(FPicture);
  inherited;
end;

procedure TMSOPicture.FreeCompressedPict;
begin
  if FCompressedPicture <> Nil then begin
    FreeMem(FCompressedPicture);
    FCompressedPicture := Nil;
  end;
end;

function TMSOPicture.LoadFromStream(Stream: TXLSStream; var Length: integer; PBuf: PByteArray): boolean;
var
  v: integer;
  Header: TMSOHeader;
  MetafileRec: TMSOBSEMetafile;
  POut: Pointer;
//  sOut: TFileStream;
begin
  Result := False;
  Dec(Length,Stream.Read(Header,SizeOf(TMSOHeader)));
  if Header.FBT <> MSO_BSE then
    raise XLSRWException.Create('Expected record missing: BSE');
  Dec(Length,Stream.Read(PBuf^,SizeOf(TMSORecBSE)));
  if PMSORecBSE(PBuf).NameLen > 0 then
    Dec(Length,Stream.ReadWideString(PMSORecBSE(PBuf).NameLen,FBlipName));
  FPictType := PMSORecBSE(PBuf).WinType;
  FRefCount := PMSORecBSE(PBuf).RefCount;

  // RefCount is zero for unused pictures (BSE without picture data).
  // The BLIP (picture) data may however be stored at the end of the DGG.
  if (FRefCount <= 0) or (FPictType = msoblipERROR) and (Header.Length = SizeOf(TMSORecBSE)) then
    Exit;

  Dec(Length,Stream.Read(Header,SizeOf(TMSOHeader)));

  if (Header.FBT < MSO_BLIPSTART) or (Header.FBT > MSO_BLIPEND) then
    raise XLSRWException.Create('Expected record missing: BLIP');
  if (Header.Length - (PMSORecBSE(PBuf).Size - SizeOf(TMSOHeader))) <> 0 then
    raise XLSRWException.Create('Blip size in BSE don''t match actually blip');
  Dec(Length,Stream.Read(FHash,16));
  if FPictType in [msoblipEMF,msoblipWMF,msoblipPICT] then begin
    Dec(Length,Stream.Read(MetafileRec,SizeOf(TMSOBSEMetafile)));
    GetMem(FPicture,MetafileRec.Size);
    Dec(Length,Stream.Read(FPicture^,MetafileRec.Size));

    v := (Integer(Header.Length) - 16) - SizeOf(TMSOBSEMetafile) - MetafileRec.Size;
    // Don't know what this extra space is for.
    while v > 0 do begin
      Dec(Length,Stream.Read(PBuf^,Min(v,MAXRECSZ_97)));
      Dec(v,MAXRECSZ_97);
    end;

    FPictSize := MetafileRec.Size;
    if MetafileRec.Compressed = 0 then begin
      DecompressBuf(FPicture,MetafileRec.Size,MetafileRec.CacheSize,POut,FPictSize);

      FreeMem(FPicture);
      FPicture := POut;
    end;
    MakeHash;
  end
  else begin
    GetMem(FPicture,Header.Length - 17);
    Dec(Length,Stream.Read(PBuf^,1));
    Dec(Length,Stream.Read(FPicture^,Header.Length - 17));
    FPictSize := Header.Length - 17;
    MakeHash;
  end;
{
  case FPictType of
    msoblipEMF:
      sOut := TFileStream.Create('d:\temp\temp.emf',fmCreate);
    msoblipWMF:
      sOut := TFileStream.Create('d:\temp\temp.wmf',fmCreate);
    msoblipPICT: ;
    msoblipJPEG:
      sOut := TFileStream.Create('d:\temp\temp.jpg',fmCreate);
    msoblipPNG:
      sOut := TFileStream.Create('d:\temp\temp.png',fmCreate);
    msoblipDIB:
      sOut := TFileStream.Create('d:\temp\temp.bmp',fmCreate);
  end;
  try
    sOut.Write(FPicture^,Header.Length - 17);
  finally
    sOut.Free;
  end;
}
  Result := FPictType <> msoblipERROR;
end;

procedure TMSOPicture.LoadFromFile(Filename: AxUCString);
var
  PictType: TMSOBlipType;
  Stream: TFileStream;
begin
  PictType := BlipFromExt(Filename);
  Stream := TFileStream.Create(Filename,fmOpenRead);
  try
    LoadFromStream(Stream,Filename,PictType);
  finally
    Stream.Free;
  end;
end;

procedure TMSOPicture.LoadFromStream(Stream: TStream; Name: AxUCString; PictType: TMSOBlipType);
var
  Strm: TStream;
begin
  Strm := Stream;
  if MyWideUppercase(ExtractFileExt(Name)) = '.BMP' then begin
    if Stream.Size > $7FFF then
      raise XLSRWException.Create('Bitmap (BMP) pictures can be max 32767 bytes in size');
  end;
  FBlipName := Name;
  FPictType := PictType;
  FRefCount := 0;

  FPictSize := Strm.Size;
  GetMem(FPicture,FPictSize);
  // Strip header.
  case FPictType of
    msoblipDIB: begin
      Strm.Seek(SizeOf(TBMPHeader),soFromBeginning);
      Dec(FPictSize,SizeOf(TBMPHeader));
    end;
    msoblipWMF: begin
      Strm.Seek(SIZE_WMF_HEADER,soFromBeginning);
      Dec(FPictSize,SIZE_WMF_HEADER);
    end;
  end;
  Strm.Read(FPicture^,FPictSize);
  MakeHash;
  if Stream = Nil then
    Strm.Free;
end;

procedure TMSOPicture.MakeHash;
var
  MD4: TMD4Hash;
begin
  MD4 := TMD4Hash.Create;
  try
    MD4.Init;
    MD4.Update(FPicture^,FPictSize);
    MD4.Final(FHash);
  finally
    MD4.Free;
  end;
end;

procedure TMSOPicture.WriteToStream(Stream: TXLSStream; PBuf: PByteArray);
var
  Size,Id: integer;
begin
  if FPictType in [msoblipEMF,msoblipWMF,msoblipPICT] then
    Size := FCompressedSize + 16 + SizeOf(TMSOBSEMetafile)
  else
    Size := FPictSize + 16 + 1;
  Stream.WriteMSOHeader($02,Word(FPictType),MSO_BSE,SizeOf(TMSORecBSE) + SizeOf(TMSOHeader) + Size);
  FillChar(PBuf^,SizeOf(TMSORecBSE),#0);
  PMSORecBSE(PBuf).WinType := FPictType;
  PMSORecBSE(PBuf).MacType := FPictType;
  Move(FHash[0],PMSORecBSE(PBuf).BlipId[0],16);
  PMSORecBSE(PBuf).Size := Size + SizeOf(TMSOHeader);
  PMSORecBSE(PBuf).RefCount := FRefCount;
  PMSORecBSE(PBuf).Tag := $FF;
  PMSORecBSE(PBuf).Unused := $0288;
  Stream.Write(PBuf^,SizeOf(TMSORecBSE));
  case FPictType of
    msoblipERROR:   Id := 0;
    msoblipUNKNOWN: Id := 0;
    msoblipEMF:     Id := BLIP_EMF;
    msoblipWMF:     Id := BLIP_WMF;
    msoblipPICT:    Id := BLIP_PICT;
    msoblipJPEG:    Id := BLIP_JPEG;
    msoblipPNG:     Id := BLIP_PNG;
    msoblipDIB:     Id := BLIP_DIB;
    else
      Id := 0;
  end;
  Stream.WriteMSOHeader($00,Id,MSO_BLIPSTART + Word(FPictType),Size);
  Stream.Write(PMSORecBSE(PBuf).BlipId,SizeOf(PMSORecBSE(PBuf).BlipId));
  if FPictType in [msoblipEMF,msoblipWMF,msoblipPICT] then begin
    FillChar(PBuf^,SizeOf(TMSOBSEMetafile),#0);

    PMSOBSEMetafile(PBuf).Size := FCompressedSize;
    PMSOBSEMetafile(PBuf).CacheSize := FPictSize;

    // Any value except zero seems to be ok.
    PMSOBSEMetafile(PBuf).Rect.Right := 100;
    PMSOBSEMetafile(PBuf).Rect.Bottom := 100;

    PMSOBSEMetafile(PBuf).Compressed := 0;

    PMSOBSEMetafile(PBuf).Filter := 254;
    Stream.Write(PBuf^,SizeOf(TMSOBSEMetafile));
    Stream.Write(FCompressedPicture^,FCompressedSize);
    FreeCompressedPict;
  end
  else begin
    Stream.Write(PMSORecBSE(PBuf).Tag,1);
    Stream.Write(FPicture^,FPictSize);
  end;
end;

procedure TMSOPicture.SetFilename(const Value: AxUCString);
var
  i: integer;
begin
  for i := 0 to FOwner.Count - 1 do begin
    if MyWideUppercase(Value) = MyWideUppercase(FOwner[i].FBlipName) then
      raise XLSRWException.Create('A picture with this name already exists.');
  end;
  LoadFromFile(Value);
end;

procedure TMSOPicture.SaveToFile(Filename: AxUCString);
var
  Stream: TFileStream;
begin
  Stream := TFileStream.Create(Filename,fmCreate);
  try
    SaveToStream(Stream);
  finally
    Stream.Free;
  end;
end;

procedure TMSOPicture.SaveToStream(Stream: TStream);
var
  L: integer;
  BMPHeader: TBMPHeader;
begin
  if FPictType = msoblipDIB then begin
    FillChar(BMPHeader,SizeOf(TBMPHeader),#0);
    L := FPictSize + SizeOf(TBMPHeader);
    BMPHeader.ID := $4D42;
    BMPHeader.Size := L - ((L div $FFFF) * $FFFF);
    BMPHeader.Blocks := (L div $FFFF) * $FFFF;
    BMPHeader.BMPType := $0036;
    Stream.Write(BMPHeader,SizeOf(TBMPHeader));
  end;
  Stream.Write(Picture^,FPictSize);
end;

procedure TMSOPicture.CompressPict;
begin
//  FCompressedPicture := FPicture;
//  FCompressedSize := FPictSize;

//  GetMem(FCompressedPicture,FPictSize);
  CompressBuf(FPicture,FPictSize,FCompressedPicture,FCompressedSize);
end;

function TMSOPicture.CopyToBitmap(Bitmap: TBitmap): boolean;
var
  Stream: TMemoryStream;
begin
  Result := True;
  Stream := TMemoryStream.Create;
  try
    case FPictType of
      msoblipERROR,
      msoblipUNKNOWN: Result := False;
      msoblipWMF,
      msoblipEMF:     Result := False;
      msoblipPICT:    Result := False;
      msoblipJPEG: begin
        Result := False;
      end;
      msoblipPNG: begin
        Result := False;
      end;
      msoblipDIB: begin
        Stream.Read(FPicture,FPictSize);
        Stream.Seek(0,soFromBeginning);
        Bitmap.LoadFromStream(Stream);
      end;
    end;
  finally
    Stream.Free;
  end;
end;

{$ifdef WINDOWS}
function TMSOPicture.CopyToMetafile(Metafile: TMetafile): boolean;
var
  Stream: TMemoryStream;
begin
  Stream := TMemoryStream.Create;
  try
    if FPictType in [msoblipWMF,msoblipEMF] then begin
      Stream.Write(FPicture^,FPictSize);
      Stream.Seek(0,soFromBeginning);
      try
        Metafile.LoadFromStream(Stream);
        Result := (Metafile.Width > 0) and (Metafile.Height > 0);
      except
        Result := False;
      end;
    end
    else
      Result := False;
  finally
    Stream.Free;
  end;
end;
{$endif}

function TMSOPicture.IsRaster: boolean;
begin
  Result := FPictType in [msoblipPICT,msoblipJPEG,msoblipPNG,msoblipDIB];
end;

{ TOPT }

procedure TOPT.AddExtra(Id: word; Value: integer; IsBlip: boolean = False);
var
  P: POPTData;
begin
  GetMem(P,SizeOf(TOPTData));
  P.Value := Value;
  P.Options := Id or $8000;
  if IsBlip then
    P.Options := P.Options or $4000;
  GetMem(P.Extra,P.Value);
  inherited Add(P);
end;

procedure TOPT.AddString(Id: word; Value: AxUCString);
var
  P: POPTData;
begin
  GetMem(P,SizeOf(TOPTData));
  P.Value := (Length(Value) * 2) + 2;
  P.Options := Id or ($4000 + $8000);
  GetMem(P.Extra,P.Value);
  System.Move(Pointer(Value)^,P.Extra^,Length(Value) * 2);
  P.Extra[P.Value - 2] := 0;
  P.Extra[P.Value - 1] := 0;
  inherited Add(P);
end;

procedure TOPT.AddValue(Id: word; Value: longword; IsBlip: boolean = False);
var
  P: POPTData;
begin
  GetMem(P,SizeOf(TOPTData));
  P.Value := Value;
  P.Options := Id;
  if IsBlip then
    P.Options := P.Options or $4000;
  P.Extra := Nil;
  // The order of the values are importatnt. If Rotation not is first, Excel craches.
  case Id of
    msooptRotation: Insert(0,P);
    else inherited Add(P);
  end;
end;

procedure TOPT.Assign(OPT: TOPT);
var
  Sz: integer;
  P: PByteArray;
begin
  Sz := Size;
  GetMem(P,Sz);
  WriteToBuffer(P,False);
  OPT.LoadFromBuffer(Sz,P,Count);
end;

function TOPT.BlipId: integer;
var
  i: integer;
begin
  Result := -1;
  for i := 0 to Count - 1 do begin
    if ((POPTData(Items[i]).Options and $8000) = 0) and ((POPTData(Items[i]).Options and $4000) <> 0) then begin
      Result := Value[i];
      Exit;
    end;
  end;
end;

procedure TOPT.Clear;
var
  i: integer;
begin
  for i := 0 to Count - 1 do begin
    FreeMem(POPTData(Items[i]).Extra);
    FreeMem(Items[i]);
  end;
  inherited Clear;
end;

destructor TOPT.Destroy;
begin
  Clear;
  inherited;
end;

function TOPT.Find(Id: integer): integer;
begin
  for Result := 0 to Count - 1 do begin
    if (POPTData(Items[Result]).Options and $3FFF) = Id then
      Exit;
  end;
  Result := -1;
end;

function TOPT.FindString(Id: integer): AxUCString;
var
  i: integer;
begin
  i := Find(Id);
  if (i >= 0) and (POPTData(Items[i]).Extra <> Nil) then begin
    SetLength(Result,(POPTData(Items[i]).Value - 2) div 2);
    System.Move(POPTData(Items[i]).Extra^,Pointer(Result)^,POPTData(Items[i]).Value - 2);
  end
  else
    Result := '';
end;

function TOPT.FindValue(Id: integer; Default: longword): longword;
var
  i: integer;
begin                                      
  i := Find(Id);
  if i >= 0 then
    Result := POPTData(Items[i]).Value
  else
    Result := Default;          
end;

function TOPT.FindValue(Id: integer): longword;
var
  i: integer;
begin
  i := Find(Id);
  if i >= 0 then
    Result := POPTData(Items[i]).Value
  else
    Result := $FFFFFFFF;
end;

function TOPT.GetComplex(Index: integer): boolean;
begin
  Result := (POPTData(Items[Index]).Options and $8000) <> 0;
end;

function TOPT.GetExtra(Index: integer): PByteArray;
begin
  Result := POPTData(Items[Index]).Extra;
end;

function TOPT.GetId(Index: integer): word;
begin
  Result := POPTData(Items[Index]).Options and $3FFF;
end;

function TOPT.GetValue(Index: integer): longword;
begin
  Result := POPTData(Items[Index]).Value;
end;

procedure TOPT.LoadFromBuffer(Length: integer; PBuf: PByteArray; ItemCount: integer);
var
  i: integer;
  Data: POPTData;
  IsComplex: boolean;
begin
  FItemCount := ItemCount;
  while Length > 0 do begin
    IsComplex := (PMSORecOPTValue(PBuf).Options and $8000) <> 0;
    New(Data);
    Data.Options := PMSORecOPTValue(PBuf).Options;
    Data.Value := PMSORecOPTValue(PBuf).Value;
    if IsComplex and (Data.Value > 0) then begin
      Data.Extra := Pointer(1);
      Dec(Length,Data.Value);
    end
    else
      Data.Extra := Nil;
    if Count < FItemCount then
      Add(Data);
    Dec(Length,SizeOf(TMSORecOPTValue));
    PBuf := PByteArray(NativeInt(PBuf) + SizeOf(TMSORecOPTValue));
  end;
  if Length <> 0 then
    raise XLSRWException.Create('Error when reading OPT items, Length <> 0');
  for i := 0 to Count - 1 do begin
    if POPTData(Items[i]).Extra <> Nil then begin
      with POPTData(Items[i])^ do begin
        GetMem(Extra,Value);
        System.Move(PBuf^,Extra^,Value);
//        Dec(Length,Value);
      end;
    end;
  end;
end;

{

}

procedure TOPT.LoadFromStream(Stream: TXLSStream; Length: integer; PBuf: PByteArray; ItemCount: integer);
var
  i: integer;
  Data: POPTData;
//  OPTId: integer;
//  IsBlip: boolean;
  IsComplex: boolean;
begin
  // BUG in Excel: Item count (Instance in OPT) don't match actually items.
  // Length works. But actually (Instance) count of items has to be written
  // to the file, otherwise crashes excel. Therefore is ItemCount saved.
  // The last item (the one that is extra) don't seems to contain valid data,
  // so it may not be a real item.
  FItemCount := ItemCount;
  while Length > 0 do begin
    Dec(Length,Stream.Read(PBuf^,SizeOf(TMSORecOPTValue)));
//    OPTId := PMSORecOPTValue(PBuf).Options and $3FFF;
//    IsBlip := (PMSORecOPTValue(PBuf).Options and $4000) <> 0;
    IsComplex := (PMSORecOPTValue(PBuf).Options and $8000) <> 0;
    New(Data);
    Data.Options := PMSORecOPTValue(PBuf).Options;
    Data.Value := PMSORecOPTValue(PBuf).Value;
    if IsComplex and (Data.Value > 0) then begin
      Data.Extra := Pointer(1);
      Dec(Length,Data.Value);
    end
    else
      Data.Extra := Nil;
    // Only add valid items, See above.
    if Count < FItemCount then
      Add(Data);
  end;
  if Length <> 0 then
    raise XLSRWException.Create('Error when reading OPT items, Length <> 0');
  for i := 0 to Count - 1 do begin
    if POPTData(Items[i]).Extra <> Nil then begin
      with POPTData(Items[i])^ do begin
        GetMem(Extra,Value);
        Stream.Read(Extra^,Value);
      end;
    end;
  end;
end;

procedure TOPT.SetBlipId(Id: integer);
var
  i: integer;
begin
  for i := 0 to Count - 1 do begin
    if ((POPTData(Items[i]).Options and $8000) = 0) and ((POPTData(Items[i]).Options and $4000) <> 0) then begin
      Value[i] := Id;
      Exit;
    end;
  end;
  raise XLSRWException.Create('Shape has no Blip Id');
end;

procedure TOPT.SetItemCount;
begin
  FItemCount := Count;
end;

procedure TOPT.SetValue(Index: integer; const Value: longword);
begin
  if ((POPTData(Items[Index]).Options and $8000) <> 0) then
    raise XLSRWException.Create('OPT value is complex');
  POPTData(Items[Index]).Value := Value;
end;

function TOPT.Size: integer;
var
  i: integer;
begin
  Result := Count * SizeOf(TMSORecOPTValue);
  for i := 0 to Count - 1 do begin
    if POPTData(Items[i]).Extra <> Nil then
      Inc(Result,POPTData(Items[i]).Value);
  end;
end;

function TOPT.TxId: longword;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do begin
    if Id[i] = msooptLTxid then begin
      Result := Value[i];
      Exit;
    end;
  end;
end;

procedure TOPT.UpdateExtra(Id, Size: integer);
var
  i: integer;
begin
  i := Find(Id);
  if i >= 0 then begin
    POPTData(Items[i]).Value := Size;
    ReAllocMem(POPTData(Items[i]).Extra,Size);
  end;
end;

procedure TOPT.UpdateString(Id: integer; Value: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do begin
    if (POPTData(Items[i]).Options and $3FFF) = Id then begin
      FreeMem(POPTData(Items[i]).Extra);
      POPTData(Items[i]).Value := (Length(Value) * 2) + 2;
      GetMem(POPTData(Items[i]).Extra,POPTData(Items[i]).Value + 2);
      System.Move(Pointer(Value)^,POPTData(Items[i]).Extra^,Length(Value) * 2);
      POPTData(Items[i]).Extra[POPTData(Items[i]).Value - 1] := 0;
      POPTData(Items[i]).Extra[POPTData(Items[i]).Value - 2] := 0;
      Exit;
    end;
  end;
end;

procedure TOPT.UpdateValue(Id: integer; Value: longword);
var
  i: integer;
begin
  for i := 0 to Count - 1 do begin
    if (POPTData(Items[i]).Options and $3FFF) = Id then begin
      POPTData(Items[i]).Value := Value;
      Exit;
    end;
  end;
end;

procedure TOPT.WriteToBuffer(PBuf: PByteArray; IncludeHeader: boolean);
var
  i: integer;
begin
  if IncludeHeader then begin
    PMSOHeader(PBuf).VerInst := $03 + (FItemCount shl 4);
    PMSOHeader(PBuf).FBT := MSO_OPT;
    PMSOHeader(PBuf).Length := Size;
    PBuf := PByteArray(NativeInt(PBuf) + SizeOf(TMSOHeader));
  end;
  for i := 0 to Count - 1 do begin
    PWordArray(PBuf)[0] := POPTData(Items[i]).Options;
    PBuf := PByteArray(NativeInt(PBuf) + 2);
    PLongWordArray(PBuf)[0] := POPTData(Items[i]).Value;
    PBuf := PByteArray(NativeInt(PBuf) + 4);
  end;
  for i := 0 to Count - 1 do begin
    if POPTData(Items[i]).Extra <> Nil then begin
      System.Move(POPTData(Items[i]).Extra^,PBuf^,POPTData(Items[i]).Value);
      PBuf := PByteArray(NativeInt(PBuf) + NativeInt(POPTData(Items[i]).Value));
    end;
  end;
end;

procedure TOPT.WriteToStream(Stream: TXLSStream; PBuf: PByteArray);
var
  i: integer;
begin
  Stream.WriteMSOHeader($03,FItemCount,MSO_OPT,Size);
  for i := 0 to Count - 1 do begin
    Stream.Write(POPTData(Items[i]).Options,SizeOf(word));
    Stream.Write(POPTData(Items[i]).Value,SizeOf(longword));
  end;
  for i := 0 to Count - 1 do begin
    if POPTData(Items[i]).Extra <> Nil then
      Stream.Write(POPTData(Items[i]).Extra^,POPTData(Items[i]).Value);
  end;
end;

{ TDGGData }

procedure TDGGData.LoadFromStream(Stream: TXLSStream);
begin
  Stream.Read(FFIDCL.DgId,4);
  Stream.Read(FFIDCL.SpIdUsed,4);
end;

procedure TDGGData.SetDrawing(const Value: TEscherDrawing);
begin
  FDrawing := Value;
end;

procedure TDGGData.WriteToStream(Stream: TXLSStream);
begin
  Stream.Write(FFIDCL.DgId,4);
  Stream.Write(FFIDCL.SpIdUsed,4);
end;

{ TDGG }

function TDGG.Add(Item: Pointer): Integer;
begin
  Result := inherited Add(Item);
end;

procedure TDGG.Clear;
var
  i: integer;
begin
  for i := 0 to Count - 1 do
    Items[i].Free;
  inherited Clear;
end;

destructor TDGG.Destroy;
begin
  Clear;
  inherited;
end;

function TDGG.GetItems(Index: integer): TDGGData;
begin
  Result := TDGGData(inherited Items[Index]);
end;

procedure TDGG.LoadFromStream(Stream: TXLSStream; var Length: integer; PBuf: PByteArray);
var
  i: integer;
  Data: TDGGData;
begin
  Dec(Length,Stream.Read(PBuf^,SizeOf(TMSORecDGG)));
  FMaxSpId := PMSORecDGG(PBuf).MaxSpId;
  FSavedShapes := PMSORecDGG(PBuf).SavedShapes;
  FSavedDrawings := PMSORecDGG(PBuf).SavedDrawings;
  // FIdClusterCount seems to be # of FIDCL + 1.
  for i := 1 to PMSORecDGG(PBuf).IdClusterCount - 1 do begin
    Data := TDGGData.Create;
    Data.LoadFromStream(Stream);
    inherited Add(Data);
  end;
  Dec(Length,Count * SizeOf(TDGGRecFIDCL));
end;

procedure TDGG.WriteToStream(Stream: TXLSStream; PBuf: PByteArray);
var
  i: integer;
  MaxId: integer;
begin
  MaxId := 0;
  FSavedShapes := 0;
  for i := 0 to Count - 1 do begin
    if Items[i].Drawing <> Nil then begin
      Inc(FSavedShapes,1 + Items[i].Drawing.ShapeCount);
      if Items[i].NewDrawing then begin
        Items[i].FFIDCL.SpIdUsed := 1 + Items[i].Drawing.ShapeCount;
        if Items[i].Drawing.MaxSpId > MaxId then
          MaxId := Items[i].Drawing.MaxSpId;
      end;
    end;
  end;
  if MaxId > Integer(FMaxSpId) then
    FMaxSpId := MaxId;
                                           
  Stream.WriteMSOHeader($00,$0000,MSO_DGG,SizeOf(TMSORecDGG) + (Count * SizeOf(TDGGRecFIDCL)));
  PMSORecDGG(PBuf).MaxSpId := FMaxSpId;
  PMSORecDGG(PBuf).IdClusterCount := Count + 1;
  PMSORecDGG(PBuf).SavedShapes := FSavedShapes;
  PMSORecDGG(PBuf).SavedDrawings := FSavedDrawings;
  Stream.Write(PBuf^,SizeOf(TMSORecDGG));
  for i := 0 to Count - 1 do
    Items[i].WriteToStream(Stream);
end;

{ TShapeInterfaceChart }

procedure TShapeOutsideMsoChart.Assign(Shape: TShapeOutsideMso);
var
  PBuf: PByteArray;
  Stream: TXLSStream;
begin
  if not (Shape is TShapeOutsideMsoChart) then
    raise XLSRWException.Create('Wrong type of shape in Assign');

  Stream := TXLSStream.Create(Nil);
  try
    Stream.OpenRawMemStreamWrite;

    FRecords.Write(Stream);
    Stream.Seek(0,soFromBeginning);
    GetMem(PBuf,MAXRECSZ_97);
    try
      TShapeOutsideMsoChart(Shape).FRecords.Read(Stream,PBuf,FFonts);
    finally
      FreeMem(PBuf);
    end;
  finally
    Stream.Free;
  end;
//  raise XLSRWException.Create('Chart Assign not implemented');
end;

constructor TShapeOutsideMsoChart.Create(ObjId: integer; DrawingGroup: TEscherGroup; Fonts: TXc12Fonts);
begin
  inherited Create(ObjId,mstChart);
  FDrawingGroup := DrawingGroup;
  FRecords := TChartRecordUpdate.Create(TChartRecord(Self),ID_CHARTRECORDROOT,0);
  FFonts := Fonts;
end;

destructor TShapeOutsideMsoChart.Destroy;
begin
  FRecords.Free;
  if FDrawing <> Nil then
    FDrawing.Free;
  inherited;
end;

procedure TShapeOutsideMsoChart.Read(Stream: TXLSStream; PBuf: PByteArray);
begin
  FRecords.Read(Stream,PBuf,FFonts);
end;

procedure TShapeOutsideMsoChart.Write(Stream: TXLSStream);
begin
  FRecords.Write(Stream);
end;

{ TShapeClientAnchor }

procedure TShapeClientAnchor.Assign(Shape: TShape);
begin
  inherited Assign(Shape);
  FAnchor.Assign(TShapeClientAnchor(Shape).Anchor);
end;

constructor TShapeClientAnchor.Create;
begin
  inherited Create;
  FAnchor := TMSOObjCLIENTANCHOR.Create;
  FVisible := True;
end;

destructor TShapeClientAnchor.Destroy;
begin
  FAnchor.Free;
  inherited;
end;

procedure TShapeClientAnchor.GetPos(var C1,C1Offs,R1,R1Offs,C2,C2Offs,R2,R2Offs: Integer);
begin
  with FAnchor do begin
    C1 := Col1;
    C1Offs := Col1Offset;
    R1 := Row1;
    R1Offs := Row1Offset;
    C2 := Col2;
    C2Offs := Col2Offset;
    R2 := Row2;
    R2Offs := Row2Offset;
  end;
end;

function TShapeClientAnchor.Size: integer;
begin
  Result := inherited Size + SizeOf(TMSOHeader) + SizeOf(TMSORecCLIENTANCHOR)
end;

{ TShapeChildAnchor }

procedure TShapeChildAnchor.Assign(Shape: TShape);
begin
  inherited Assign(Shape);
  Move(FCHILDANCHOR,TShapeChildAnchor(Shape).FCHILDANCHOR,SizeOf(TMSORecCHILDANCHOR));
end;

function TShapeChildAnchor.Size: integer;
begin
  Result := inherited Size + SizeOf(TMSOHeader) + SizeOf(TMSORecCHILDANCHOR)
end;

{ TShapeGroupClientAnchor }

procedure TShapeGroupClientAnchor.Assign(Shape: TShape);
begin
  inherited Assign(Shape);
  FAnchor.Assign(TShapeGroupClientAnchor(Shape).FAnchor);
end;

constructor TShapeGroupClientAnchor.Create(Drawing: TEscherDrawing);
begin
  inherited Create(Drawing);
  FAnchor := TMSOObjCLIENTANCHOR.Create;
end;

destructor TShapeGroupClientAnchor.Destroy;
begin
  FAnchor.Free;
  inherited;
end;

procedure TShapeGroupClientAnchor.GetPos(var C1,C1Offs,R1,R1Offs,C2,C2Offs,R2,R2Offs: Integer);
begin
  with FAnchor do begin
    C1 := Col1;
    C1Offs := Col1Offset;
    R1 := Row1;
    R1Offs := Row1Offset;
    C2 := Col2;
    C2Offs := Col2Offset;
    R2 := Row2;
    R2Offs := Row2Offset;
  end;
end;

function TShapeGroupClientAnchor.PrivateSize: integer;
begin
  Result := inherited Size +
            // SPGRCONTAINER
            SizeOf(TMSOHeader) +
            // CLIENTANCHOR
            SizeOf(TMSOHeader) + SizeOf(TMSORecCLIENTANCHOR) +
            // SPGR
            SizeOf(TMSOHeader) + SizeOf(TMSORecSPGR);
end;

function TShapeGroupClientAnchor.Size: integer;
var
  i: integer;
begin
  Result := PrivateSize;
  for i := 0 to FList.Count - 1 do
    Inc(Result,Items[i].Size);
end;

procedure TShapeGroupClientAnchor.WriteToStream(Stream: TXLSStream; PBuf: PByteArray; WriteMSODRAWING: boolean);
var
  Sz: integer;
  i: integer;
begin
  Sz := PrivateSize;
  if WriteMSODRAWING then
    Stream.WriteHeader(BIFFRECID_MSODRAWING,Sz);
  Stream.WriteMSOHeader($0F,$0000,MSO_SPGRCONTAINER,Size - SizeOf(TMSOHeader));

  Stream.WriteMSOHeader($0F,$0000,MSO_SPCONTAINER,PrivateSize - (SizeOf(TMSOHeader) * 2));

  Stream.WriteMSOHeader($01,$0000,MSO_SPGR,SizeOf(TMSORecSPGR));
  Stream.Write(FSPGR,SizeOf(TMSORecSPGR));

  Stream.WriteMSOHeader($02,FShapeType,MSO_SP,SizeOf(TMSORecSP));
  Stream.Write(FSpId,SizeOf(FSpId));
  Stream.Write(FOptions,SizeOf(FOptions));

  FOPT.WriteToStream(Stream,PBuf);
  if FUnknown <> Nil then
    FUnknown.Write(Stream);

  FAnchor.Write(Stream);

  Stream.WriteMSOHeader($00,$0000,MSO_CLIENTDATA,0);

  Stream.WriteHeader(BIFFRECID_OBJ,FOBJ.Size);
  FOBJ.WriteAllRecs(Stream);

  if FInterface <> Nil then
    FInterface.Write(Stream);
  for i := 0 to FList.Count - 1 do
    Items[i].WriteToSTream(Stream,PBuf,True);
end;

{ TShapeGroupChildAnchor }

procedure TShapeGroupChildAnchor.Assign(Shape: TShape);
begin
  inherited Assign(Shape);
  Move(FCHILDANCHOR,TShapeChildAnchor(Shape).FCHILDANCHOR,SizeOf(TMSORecCHILDANCHOR));
end;

function TShapeGroupChildAnchor.PrivateSize: integer;
begin
  Result := inherited Size +
            // SPGRCONTAINER
            SizeOf(TMSOHeader) +
            // CHILDANCHOR
            SizeOf(TMSOHeader) + SizeOf(TMSORecCHILDANCHOR) +
            // SPGR
            SizeOf(TMSOHeader) + SizeOf(TMSORecSPGR);
end;

function TShapeGroupChildAnchor.Size: integer;
var
  i: integer;
begin
  Result := PrivateSize;
  for i := 0 to FList.Count - 1 do
    Inc(Result,Items[i].Size);
end;

procedure TShapeGroupChildAnchor.WriteToStream(Stream: TXLSStream; PBuf: PByteArray; WriteMSODRAWING: boolean);
var
  Sz: integer;
  i: integer;
begin
  Sz := PrivateSize;
  if WriteMSODRAWING then
    Stream.WriteHeader(BIFFRECID_MSODRAWING,Sz);
  Stream.WriteMSOHeader($0F,$0000,MSO_SPGRCONTAINER,Size - SizeOf(TMSOHeader));

  Stream.WriteMSOHeader($0F,$0000,MSO_SPCONTAINER,PrivateSize - (SizeOf(TMSOHeader) * 2));

  Stream.WriteMSOHeader($01,$0000,MSO_SPGR,SizeOf(TMSORecSPGR));
  Stream.Write(FSPGR,SizeOf(TMSORecSPGR));

  Stream.WriteMSOHeader($02,FShapeType,MSO_SP,SizeOf(TMSORecSP));
  Stream.Write(FSpId,SizeOf(FSpId));
  Stream.Write(FOptions,SizeOf(FOptions));

  FOPT.WriteToStream(Stream,PBuf);
  if FUnknown <> Nil then
    FUnknown.Write(Stream);
  Stream.WriteMSOHeader($00,$0000,MSO_CHILDANCHOR,SizeOf(TMSORecCHILDANCHOR));
  Stream.Write(FCHILDANCHOR,SizeOf(TMSORecCHILDANCHOR));
  Stream.WriteMSOHeader($00,$0000,MSO_CLIENTDATA,0);

  Stream.WriteHeader(BIFFRECID_OBJ,FOBJ.Size);
  FOBJ.WriteAllRecs(Stream);

  if FInterface <> Nil then
    FInterface.Write(Stream);
  for i := 0 to FList.Count - 1 do
    Items[i].WriteToSTream(Stream,PBuf,True);
end;

initialization
  DebugCount := 0;

end.
