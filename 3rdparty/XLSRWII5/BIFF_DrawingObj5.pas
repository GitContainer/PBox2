unit BIFF_DrawingObj5;

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
{$else}
     vcl.Graphics,
{$endif}
     BIFF_DrawingObjAnchor5, BIFF_RecsII5, BIFF_Escher5, BIFF_EscherCtrlObj5,
     BIFF_Stream5, BIFF_EscherTypes5, BIFF_FormulaHandler5, BIFF_Utils5,
     BIFF_ExcelFuncII5,
     Xc12DataStylesheet5,
     XLSUtils5, XLSFormulaTypes5;

//* Base class for drawing objects.
type TDrwCustom = class(TDrwAnchor)
private
     function  GetRotation: double;
     procedure SetRotation(const Value: double);
     function  GetMacroName: AxUCString;
     procedure SetMacroName(const Value: AxUCString); virtual;
     function  GetLineColor: TColor;
     procedure SetLineColor(const Value: TColor);
     function  GetFillColor: TColor;
     procedure SetFillColor(const Value: TColor);
     function  GetMoveWithCells: boolean;
     procedure SetMoveWithCells(const Value: boolean);
public
     //* Macro that is executed when the drawing onject is clicked.
     //* Cell notes can not have macros assigned.
     property MacroName: AxUCString read GetMacroName write SetMacroName;
     //* Rotation, in degrees, of the drawing object.
     property Rotation: double read GetRotation write SetRotation;
     //* Line color of the object.
     property LineColor: TColor read GetLineColor write SetLineColor;
     //* Fill color of the object. Only for objects that have a fill color,
     //* such as ellipses and rectangles.
     property FillColor: TColor read GetFillColor write SetFillColor;
     //* Graphic is moved and sized when cells (rows/columns) are resized.
     property MoveWithCells: boolean read GetMoveWithCells write SetMoveWithCells;
     end;

type TDrwCollection = class(TCollection)
protected
     FOwner: TPersistent;
     FDrawing: TEscherDrawing;
     FFileAdd: boolean;

     function  GetOwner: TPersistent; override;
public
     //* ~exclude
     constructor Create(ItemClass: TCollectionItemClass; AOwner: TPersistent; Drawing: TEscherDrawing);
     //* Cleare the list of all objects.
     procedure Clear;
     //* Deletes an object in the list.
     //* ~param Index Index of the object to be deleted.
     procedure Delete(Index: integer);
     end;

type
//* Base class for drawing objects with text, such ax text boxes and notes.
    TCustomDrwText = class(TDrwCustom)
private
     FName: string;
     procedure SetName(const Value: AxUCString);
     function  GetName: AxUCString;
     function  GetText: AxUCString;
     procedure SetText(const Value: AxUCString);
     function  GetFormatting(Index: integer): TRecTXORUN;
     function  GetMarginLeft: double;
     function  GetMarginBottom: double;
     function  GetMarginRight: double;
     function  GetMarginTop: double;
     procedure SetMarginLeft(const Value: double);
     procedure SetMarginBottom(const Value: double);
     procedure SetMarginRight(const Value: double);
     procedure SetMarginTop(const Value: double);
protected
public
     //* ~exclude
     constructor Create(Collection: TCollection); override;
     //* ~exclude
     destructor Destroy; override;
     //* ~exclude
     procedure SetShape(const Value: TShapeClientAnchor); override;
     //* ~exclude
     procedure LoadFromStream(Stream: TStream);
     //* ~exclude
     procedure LoadFromFile(Filename: AxUCString);
     //* ~exclude
     property Formatting[Index: integer]: TRecTXORUN read GetFormatting;
published
     //* The name of the object.
     property Name: AxUCString read GetName write SetName;
     //* The text. Texts can be max 32767 characters long.
     property Text: AxUCString read GetText write SetText;
     // Left margin, in centimeters, for the text in the object.
     property MarginLeft: double read GetMarginLeft write SetMarginLeft;
     // Top margin, in centimeters, for the text in the object.
     property MarginTop: double read GetMarginTop write SetMarginTop;
     // Right margin, in centimeters, for the text in the object.
     property MarginRight: double read GetMarginRight write SetMarginRight;
     // Bottom margin, in centimeters, for the text in the object.
     property MarginBottom: double read GetMarginBottom write SetMarginBottom;
     end;

type
//* Text box drawing object.
    TDrwText = class(TCustomDrwText)
private
protected
public
     //* ~exclude
     constructor Create(Collection: TCollection); override;
published
     end;

type
//* List for text box drawing objects.
    TDrwTexts = class(TDrwCollection)
private
     function  GetDrwText(Index: integer): TDrwText;
protected
     FFonts: TXc12Fonts;

     function  GetOwner: TPersistent; override;
public
     //* ~exclude
     constructor Create(AOwner: TPersistent; Drawing: TEscherDrawing; Fonts: TXc12Fonts);
     //* Add a new TDrwText.
     //* ~result The new TDrwText object.
     function  Add: TDrwText;
     //* ~exclude
     procedure AddFromFile(Shape: TShapeClientAnchor);

     //* Items in the list.
     property Items[Index: integer]: TDrwText read GetDrwText; default;
     end;

type
//* Cell note drawing object.
//* ~[br]
//* ~[br]
//* ~[b Example:]
//* ~[br]
//* ~[sample
//*  XLS.Sheets[0~[].AsString[3,3~[] := 'Cell with note';
//*  with XLS.Sheets[0~[].DrawingObjects.Notes.Add do begin
//*    CellCol := 3;
//*    CellRow := 3;
//*    Col1 := 3;
//*    Row1 := 2;
//*    Col2 := 5;
//*    Row2 := 8;
//*    FillColor := clYellow;
//*    LineColor := clRed;
//*    MarginLeft := 0.5;
//*    MarginRight := 0.5;
//*    Text := 'Cell note sample.';
//*  end;
//* ]
    TDrwNote = class(TCustomDrwText)
private
     procedure SetCellCol(const Value: integer);
     procedure SetCellRow(const Value: integer);
     function  GetCellCol: integer;
     function  GetCellRow: integer;
     function  GetAuthor: AxUCString;
     procedure SetAuthor(const Value: AxUCString);
     procedure SetMacroName(const Value: AxUCString); override;
     function  GetAlwaysVisible: boolean;
     procedure SetAlwaysVisible(const Value: boolean);
protected
public
     //* ~exclude
     constructor Create(Collection: TCollection); override;
     //* ~exclude
     destructor Destroy; override;
published
     //* The cell column that the note is attached to.
     property CellCol: integer read GetCellCol write SetCellCol;
     //* The cell row that the note is attached to.
     property CellRow: integer read GetCellRow write SetCellRow;
     //* Author text of the note.
     property Author: AxUCString read GetAuthor write SetAuthor;
     //* This property is True if the note always is visible.
     property AlwaysVisible: boolean read GetAlwaysVisible write SetAlwaysVisible;
     end;

type
//* List for notes drawing objects.
    TDrwNotes = class(TDrwCollection)
private
     function  GetDrwNote(Index: integer): TDrwNote;
protected
     FFonts: TXc12Fonts;

     function  GetOwner: TPersistent; override;
public
     //* ~exclude
     constructor Create(AOwner: TPersistent; Drawing: TEscherDrawing; Fonts: TXc12Fonts);
     //* Add a new TDrwNote.
     //* ~result The new TDrwNote object.
     function  Add: TDrwNote;
     //* ~exclude
     procedure AddFromFile(Shape: TShapeClientAnchor);
     //* Find a note by the cell column and row it is attached to. If found,
     //* the index in the list is return. When not found, -1 is returned.
     //* ~param Col Column where the note shall be.
     //* ~param Row Row where the note shall be.
     //* ~result Index in the list of the found note. If not found, -1 is returned.
     function  FindByColRow(Col,Row: integer): integer;

     //* Items in the list.
     property Items[Index: integer]: TDrwNote read GetDrwNote; default;
     end;

type
//* Basic drawing objects. Basic objects are such as lines, arrows, rectangles and circles.
//* Default object type when creating a new object is bstArrow.
    TDrwBasic = class(TDrwCustom)
private
     function  GetShapeType: TBasicShapeType;
     procedure SetShapeType(const Value: TBasicShapeType);
protected
public
     //* ~exclude
     constructor Create(Collection: TCollection); override;
published
     //* Type of drawing object. See ~[link TBasicShapeType].
     property ShapeType: TBasicShapeType read GetShapeType write SetShapeType;
     end;

type
//* List of basic drawing objects.
    TDrwBasics = class(TDrwCollection)
private
     function  GetDrwBasic(Index: integer): TDrwBasic;
protected
     function  GetOwner: TPersistent; override;
public
     //* ~exclude
     constructor Create(AOwner: TPersistent; Drawing: TEscherDrawing);
     //* Add a new TDrwBasic.
     //* ~result The new TDrwBasic object.
     function  Add: TDrwBasic;
     //* ~exclude
     procedure AddFromFile(Shape: TShapeClientAnchor);
     //* Items in the list.
     property Items[Index: integer]: TDrwBasic read GetDrwBasic; default;
     end;

type
//* Autoshape drawing objects. Autoshapes are drawing object such as  sun, moon, arrows, callouts etc.
//* Default object type when creating a new object is astSun.
//* ~[br]
//* ~[br]
//* ~[b Example:]
//* ~[br]
//* ~[sample
//*  with XLS.Sheets[0~[].DrawingObjects.AutoShapes.Add do begin
//*    Col1 := 3;
//*    Row1 := 2;
//*    Col2 := 5;
//*    Row2 := 8;
//*    ShapeType := astLeftArrowCallout;
//*    FillColor := clYellow;
//*    LineColor := clRed;
//*  end;
//* ]
    TDrwAutoShape = class(TDrwCustom)
private
     function  GetAutoShapeType: TAutoShapeType;
     function  GetFillColor: TColor;
     function  GetLineColor: TColor;
     procedure SetAutoShapeType(const Value: TAutoShapeType);
     procedure SetFillColor(const Value: TColor);
     procedure SetLineColor(const Value: TColor);
protected
public
     //* ~exclude
     constructor Create(Collection: TCollection); override;
     procedure SetAutoShapeTypeInteger(Value: integer);
published
     //* Type of autoshape object. See ~[link TAutoShapeType]
     property ShapeType: TAutoShapeType read GetAutoShapeType write SetAutoShapeType;
     //* Line color of the object.
     property LineColor: TColor read GetLineColor write SetLineColor;
     //* Fill color of the object.
     property FillColor: TColor read GetFillColor write SetFillColor;
     end;

type
//* List of TAutoShape objects.
    TDrwAutoShapes = class(TDrwCollection)
private
     function  GetDrwAutoShape(Index: integer): TDrwAutoShape;
protected

     function  GetOwner: TPersistent; override;
public
     //* ~exclude
     constructor Create(AOwner: TPersistent; Drawing: TEscherDrawing);
     //* Add a new TDrwAutoShape.
     //* ~result The new TDrwAutoShape object.
     function  Add: TDrwAutoShape;
     //* ~exclude
     procedure AddFromFile(Shape: TShapeClientAnchor);
     //* Items in the list.
     property Items[Index: integer]: TDrwAutoShape read GetDrwAutoShape; default;
     end;

type
//* Picture drawing object. In order to save space, the picture itself is
//* in ~[link XLSReadWriteII4.TXLSReadWriteII4.MSOPictures]. The TDrwPicture object is only a
//* link to a TMSOPicture object. Before a picture can be inserted into a
//* worksheet, it must first have been loaded into a TMSOPicture object.
//* ~[br]
//* ~[br]
//* ~[b Example:]
//* ~[br]
//* ~[sample
//*  // IMPORTANT! As XLSReadWriteII not can check that the syntax of the macro
//*  // code is correct, try to test run it in Excel first.
//*  // If there are errors in the code, Excel will open the file with an
//*  // error message.
//*  EditVBA must be set to True in order to edit macros.
//*  XLS.VBA.EditVBA := True;
//*  // Add a VBA module and a procedure.
//*  with XLS.VBA.AddModule('Module1',vmtMacro) do begin
//*    Source.Text :=
//*    'Sub MyTestMacro()' + #13#10 +
//*    'MsgBox "MyTestMacro executed"' + #13#10 +
//*    'End Sub';
//*  end;
//*
//*  // All pictures are globally stored, in order to save space.
//*  // (Change the file name to a picture file on your disk)
//*  XLS.MSOPictures.Add.Filename := 'd:\temp\pig.jpg';
//*
//*  // Insert the picture into the sheet.
//*  with XLS.Sheets[0~[].DrawingObjects.Pictures.Add do begin
//*    // IMPORTANT! If:
//*    // a. The file is a BMP file, and
//*    // b. TPNGImage is used, and
//*    // c. The size is more than 255 bytes, then then picture will be converted
//*    // to a PNG picture, and the filename's extension is changed to PNG.
//*    // Excel only accepts BMP pictures smaller than 32k.
//*    PictureName := XLS.MSOPictures[XLS.MSOPictures.Count - 1~[].Filename;
//*    Col1 := 1;
//*    Row1 := 1;
//*    Col2 := 4;
//*    Col2Offset := 0.25;
//*    Row2 := 9;
//*    Row2Offset := 0.5;
//*    Rotation := 45;
//*    // The name of the previous added macro.
//*    MacroName := 'MyTestMacro';
//*  end;
//* ]
    TDrwPicture = class(TDrwCustom)
private
     function  GetPictureId: integer;
     function  GetPictureName: AxUCString;
     procedure SetPictureId(const Value: integer);
     procedure SetPictureName(const Value: AxUCString);
protected
public
     //* ~exclude
     constructor Create(Collection: TCollection); override;
published
     //* ~exclude
     property PictureId: integer read GetPictureId write SetPictureId;
     //* The name of the picture. This shall be the same name as the
     //* filename in ~[link XLSReadWriteII2.TXLSReadWriteII2.MSOPictures].
     property PictureName: AxUCString read GetPictureName write SetPictureName;
     end;

type
//* List of TDrwPicture objects.
    TDrwPictures = class(TDrwCollection)
private
     function  GetDrwPicture(Index: integer): TDrwPicture;
protected
     function  GetOwner: TPersistent; override;
public
     //* ~exclude
     constructor Create(AOwner: TPersistent; Drawing: TEscherDrawing);
     //* Add a new TDrwPicture.
     //* ~result The new TDrwPicture object.
     function  Add: TDrwPicture;
     //* ~exclude
     procedure AddFromFile(Shape: TShapeClientAnchor);
     //* Items in the list.
     property Items[Index: integer]: TDrwPicture read GetDrwPicture; default;
     end;

type
// List of all lists with drawing objects on a worksheet.
   TDrawingObjects = class(TPersistent)
private
     FOwner: TPersistent;
     FDrawing: TEscherDrawing;
     FormulaHandler: TFormulaHandler;

     FTexts: TDrwTexts;
     FNotes: TDrwNotes;
     FBasics: TDrwBasics;
     FAutoShapes: TDrwAutoShapes;
     FPictures: TDrwPictures;
protected
     function  GetOwner: TPersistent; override;
public
     //* ~exclude
     constructor Create(AOwner: TPersistent; Drawing: TEscherDrawing; FmlaHandler: TFormulaHandler; Fonts: TXc12Fonts);
     //* ~exclude
     destructor Destroy; override;
     //* Clear all drawing objects.
     procedure Clear;
published
     // Text box drawing objects.
     property Texts: TDrwTexts read FTexts write FTexts;
     // Cell notes drawing objects.
     property Notes: TDrwNotes read FNotes write FNotes;
     // Basic drawing objects.
     property Basics: TDrwBasics read FBasics write FBasics;
     // Autoshape drawing objects.
     property AutoShapes: TDrwAutoShapes read FAutoShapes write FAutoShapes;
     // Picture drawing objects.
     property Pictures: TDrwPictures read FPictures write FPictures;
     end;

implementation

uses BIFF_SheetData5, BIFF5;

{ TCustomDrwText }

constructor TCustomDrwText.Create(Collection: TCollection);
begin
  inherited;
end;

destructor TCustomDrwText.Destroy;
begin
  inherited;
end;

function TCustomDrwText.GetFormatting(Index: integer): TRecTXORUN;
begin
  if (Index < 0) or (Index > High(TShapeOutsideMsoBaseText(FShape.ExShape).Formatting)) then
    raise XLSRWException.Create('Index out of range.');
  Result := TShapeOutsideMsoBaseText(FShape.ExShape).Formatting[Index];
end;

function TCustomDrwText.GetMarginBottom: double;
begin
  Result := EmuToCm(FShape.OPT.FindValue(msooptDyTextBottom));
end;

function TCustomDrwText.GetMarginLeft: double;
begin
  Result := EmuToCm(FShape.OPT.FindValue(msooptDxTextLeft));
end;

function TCustomDrwText.GetMarginRight: double;
begin
  Result := EmuToCm(FShape.OPT.FindValue(msooptDxTextRight));
end;

function TCustomDrwText.GetMarginTop: double;
begin
  Result := EmuToCm(FShape.OPT.FindValue(msooptDyTextTop));
end;

function TCustomDrwText.GetName: AxUCString;
begin
  if (FShape <> Nil) and (FShape.OPT.FindString(msooptWzName) <> '') then
    Result := FShape.OPT.FindString(msooptWzName)
  else
    Result := FName;
end;

function TCustomDrwText.GetText: AxUCString;
begin
  Result := TShapeOutsideMsoBaseText(FShape.ExShape).Text;
end;

procedure TCustomDrwText.LoadFromFile(Filename: AxUCString);
var
  Stream: TFileStream;
begin
  Stream := TFileStream.Create(Filename,fmOpenRead);
  try
    LoadFromStream(Stream);
  finally
    Stream.Free;
  end;
end;

procedure TCustomDrwText.LoadFromStream(Stream: TStream);
var
  S: XLS8String;
begin
  SetLength(S,Stream.Size);
  Stream.Read(Pointer(S)^,Stream.Size);
  SetText(AxUCString(S));
end;

procedure TCustomDrwText.SetMarginBottom(const Value: double);
begin
  FShape.OPT.UpdateValue(msooptDyTextBottom,CmToEmu(Value));
end;

procedure TCustomDrwText.SetMarginLeft(const Value: double);
begin
  FShape.OPT.UpdateValue(msooptDxTextLeft,CmToEmu(Value));
end;

procedure TCustomDrwText.SetMarginRight(const Value: double);
begin
  FShape.OPT.UpdateValue(msooptDxTextRight,CmToEmu(Value));
end;

procedure TCustomDrwText.SetMarginTop(const Value: double);
begin
  FShape.OPT.UpdateValue(msooptDyTextTop,CmToEmu(Value));
end;

procedure TCustomDrwText.SetName(const Value: AxUCString);
begin
  if (FShape <> Nil) and (FShape.OPT.FindString(msooptWzName) <> '') then
    FShape.OPT.UpdateString(msooptWzName,Value);
  FName := Value;
end;

procedure TCustomDrwText.SetShape(const Value: TShapeClientAnchor);
begin
  inherited SetShape(Value);
//   TDrwTextStrings(FStrings).FOBJ := TOBJNote(Value);
end;

procedure TCustomDrwText.SetText(const Value: AxUCString);
begin
  if Length(Value) > MAX_EXCEL_STRSZ then
    TShapeOutsideMsoBaseText(FShape.ExShape).Text := Copy(Value,1,MAX_EXCEL_STRSZ)
  else
    TShapeOutsideMsoBaseText(FShape.ExShape).Text := Value;
end;

{ TDrwText }

constructor TDrwText.Create(Collection: TCollection);
begin
  if not TDrwTexts(Collection).FFileAdd then
    FShape := TDrwTexts(Collection).FDrawing.AddTextBox;
  inherited Create(Collection);
  FName := 'Text ' + IntToStr(ID);
end;

{ TDrwTexts }

function TDrwTexts.Add: TDrwText;
begin
  Result := TDrwText(inherited Add);
end;

procedure TDrwTexts.AddFromFile(Shape: TShapeClientAnchor);
begin
  FFileAdd := True;
  try
    TDrwText(inherited Add).FShape := Shape;
  finally
    FFileAdd := False;
  end;
end;

constructor TDrwTexts.Create(AOwner: TPersistent; Drawing: TEscherDrawing; Fonts: TXc12Fonts);
begin
  inherited Create(TDrwText,AOwner,Drawing);
  FFonts := Fonts;
end;

function TDrwTexts.GetDrwText(Index: integer): TDrwText;
begin
  Result := TDrwText(inherited Items[Index]);
end;

function TDrwTexts.GetOwner: TPersistent;
begin
  Result := FOwner;
end;

{ TDrwNote }

procedure TDrwNote.SetCellCol(const Value: integer);
begin
  TShapeOutsideMsoNote(FShape.ExShape).CellCol := Value;
end;

procedure TDrwNote.SetCellRow(const Value: integer);
begin
  TShapeOutsideMsoNote(FShape.ExShape).CellRow := Value;
end;

procedure TDrwNote.SetMacroName(const Value: AxUCString);
begin
  raise XLSRWException.Create('Notes can not have macros assigned');
end;

function TDrwNote.GetCellCol: integer;
begin
  Result := TShapeOutsideMsoNote(FShape.ExShape).CellCol;
end;

function TDrwNote.GetCellRow: integer;
begin
  Result := TShapeOutsideMsoNote(FShape.ExShape).CellRow;
end;

constructor TDrwNote.Create(Collection: TCollection);
begin
  if not TDrwNotes(Collection).FFileAdd then
    FShape := TDrwNotes(Collection).FDrawing.AddNote;
  inherited Create(Collection);
  FName := 'Note ' + IntToStr(ID);
end;

destructor TDrwNote.Destroy;
begin
  inherited;
end;

function TDrwNote.GetAlwaysVisible: boolean;
begin
  Result := (TShapeOutsideMsoNote(FShape.ExShape).Options and SHP_NOTE_OPTION_ALWAYS_VISIBLE) <> 0;
end;

function TDrwNote.GetAuthor: AxUCString;
begin
  Result := TShapeOutsideMsoNote(FShape.ExShape).Author;
end;

procedure TDrwNote.SetAlwaysVisible(const Value: boolean);
begin
  if Value then
    TShapeOutsideMsoNote(FShape.ExShape).Options := TShapeOutsideMsoNote(FShape.ExShape).Options or SHP_NOTE_OPTION_ALWAYS_VISIBLE
  else
    TShapeOutsideMsoNote(FShape.ExShape).Options := TShapeOutsideMsoNote(FShape.ExShape).Options and not SHP_NOTE_OPTION_ALWAYS_VISIBLE;
end;

procedure TDrwNote.SetAuthor(const Value: AxUCString);
begin
  TShapeOutsideMsoNote(FShape.ExShape).Author := Value;
end;

{ TDrwNotes }

function TDrwNotes.Add: TDrwNote;
begin
  Result := TDrwNote(inherited Add);
end;

procedure TDrwNotes.AddFromFile(Shape: TShapeClientAnchor);
begin
  FFileAdd := True;
  try
    TDrwNote(inherited Add).FShape := Shape;
  finally
    FFileAdd := False;
  end;
end;

constructor TDrwNotes.Create(AOwner: TPersistent; Drawing: TEscherDrawing; Fonts: TXc12Fonts);
begin
  inherited Create(TDrwNote,AOwner,Drawing);
  FFonts := Fonts;
end;

function TDrwNotes.FindByColRow(Col,Row: integer): integer;
{
var
  hi,lo: integer;
}  
begin

  for Result := 0 to Count - 1 do begin
    if (Items[Result].CellCol = Col) and (Items[Result].CellRow = Row) then
      Exit;
  end;


  // Notes are not sorted by RowCol...
{
  lo := 0;
  hi := Count - 1;
  while lo <= hi do begin
    Result := (lo + hi) shr 1;
    if RowCol > Items[Result].RowCol then
      lo := Result + 1
    else if RowCol < Items[Result].RowCol then
      hi := Result - 1
    else
      Exit;
  end;
}  
  Result := -1;
end;

function TDrwNotes.GetDrwNote(Index: integer): TDrwNote;
begin
  Result := TDrwNote(inherited Items[Index]);
end;

function TDrwNotes.GetOwner: TPersistent;
begin
  Result := FOwner;
end;

{ TDrawingObjects }

procedure TDrawingObjects.Clear;
begin
  FTexts.Clear;
  FNotes.Clear;
  FBasics.Clear;
  FAutoShapes.Clear;
  FPictures.Clear;
end;

constructor TDrawingObjects.Create(AOwner: TPersistent; Drawing: TEscherDrawing; FmlaHandler: TFormulaHandler; Fonts: TXc12Fonts);
begin
  FOwner := AOwner;
  FDrawing := Drawing;
  FormulaHandler := FmlaHandler;
  FTexts := TDrwTexts.Create(Self,FDrawing,Fonts);
  FNotes := TDrwNotes.Create(Self,FDrawing,Fonts);
  FBasics := TDrwBasics.Create(Self,FDrawing);
  FAutoShapes := TDrwAutoShapes.Create(Self,FDrawing);
  FPictures := TDrwPictures.Create(Self,FDrawing);
end;

destructor TDrawingObjects.Destroy;
begin
  FTexts.Free;
  FNotes.Free;
  FBasics.Free;
  FAutoShapes.Free;
  FPictures.Free;
  inherited;
end;

function TDrawingObjects.GetOwner: TPersistent;
begin
  Result := FOwner;
end;

{ TDrwBasic }

constructor TDrwBasic.Create(Collection: TCollection);
begin
  if not TDrwBasics(Collection).FFileAdd then
    FShape := TDrwBasics(Collection).FDrawing.AddBasic;
  inherited Create(Collection);
  FName := 'Basic ' + IntToStr(ID);
end;

function TDrwBasic.GetShapeType: TBasicShapeType;
begin
  case FShape.ShapeType of
    msosptLine: begin
      if FShape.OPT.FindValue(msooptLineEndArrowHead) = msolineArrowEnd then
        Result := bstArrow
      else
        Result := bstLine;
    end;
    msosptEllipse:                Result := bstEllipse;
    msosptRectangle:              Result := bstRectangle;
  else
    Result := bstLine;
  end;
end;

procedure TDrwBasic.SetShapeType(const Value: TBasicShapeType);
begin
  case Value of
    bstLine,bstArrow:          FShape.ShapeType := msosptLine;
    bstEllipse:                FShape.ShapeType := msosptEllipse;
    bstRectangle:              FShape.ShapeType := msosptRectangle;
  else
    FShape.ShapeType := msosptLine;
  end;
  FShape.OPT.Clear;

  case FShape.ShapeType of
    msosptLine,msosptArrow: begin
      FShape.OPT.AddValue(msooptFFitTextToShape,$00080008);
      FShape.OPT.AddValue(msooptShapePath,msoshapeComplex);
      FShape.OPT.AddValue(msooptFFillOk,$00010000);
      FShape.OPT.AddValue(msooptFNoFillHitTest,$00100000);
      FShape.OPT.AddValue(msooptLineColor,$08000040);
      if Value = bstArrow then
        FShape.OPT.AddValue(msooptLineEndArrowHead,msolineArrowEnd)
      else
        FShape.OPT.AddValue(msooptLineEndArrowHead,msolineNoEnd);
      FShape.OPT.AddValue(msooptFNoLineDrawDash,$00100010);
    end;
    msosptEllipse,msosptRectangle: begin
      FShape.OPT.AddValue(msooptFFitTextToShape,$00080008);
      FShape.OPT.AddValue(msooptFillColor,$0800000D);
      FShape.OPT.AddValue(msooptFNoFillHitTest,$00100010);
      FShape.OPT.AddValue(msooptLineColor,$08000040);
    end;
  end;
  FShape.OPT.SetItemCount;
end;

{ TDrwBasics }

function TDrwBasics.Add: TDrwBasic;
begin
  Result := TDrwBasic(inherited Add);
end;

procedure TDrwBasics.AddFromFile(Shape: TShapeClientAnchor);
begin
  FFileAdd := True;
  try
    TDrwBasic(inherited Add).FShape := Shape;
  finally
    FFileAdd := False;
  end;
end;

constructor TDrwBasics.Create(AOwner: TPersistent; Drawing: TEscherDrawing);
begin
  inherited Create(TDrwBasic,AOwner,Drawing);
end;

function TDrwBasics.GetDrwBasic(Index: integer): TDrwBasic;
begin
  Result := TDrwBasic(inherited Items[Index]);
end;

function TDrwBasics.GetOwner: TPersistent;
begin
  Result := FOwner;
end;

{ TDrwAutoShape }

constructor TDrwAutoShape.Create(Collection: TCollection);
begin
  if not TDrwAutoShapes(Collection).FFileAdd then
    FShape := TDrwAutoShapes(Collection).FDrawing.AddAutoShape;
  inherited Create(Collection);
  FName := 'AutoShape ' + IntToStr(ID);
end;

function TDrwAutoShape.GetAutoShapeType: TAutoShapeType;
begin
  case FShape.ShapeType of
    msosptSun:                   Result := astSun;
    msosptMoon:                  Result := astMoon;
    msosptLeftArrow:             Result := astLeftArrow;
    msosptDownArrow:             Result := astDownArrow;
    msosptUpArrow:               Result := astUpArrow;
    msosptLeftRightArrow:        Result := astLeftRightArrow;
    msosptUpDownArrow:           Result := astUpDownArrow;
    msosptQuadArrow:             Result := astQuadArrow;
    msosptLeftArrowCallout:      Result := astLeftArrowCallout;
    msosptRightArrowCallout:     Result := astRightArrowCallout;
    msosptUpArrowCallout:        Result := astUpArrowCallout;
    msosptDownArrowCallout:      Result := astDownArrowCallout;
    msosptLeftRightArrowCallout: Result := astLeftRightArrowCallout;
    msosptUpDownArrowCallout:    Result := astUpDownArrowCallout;
    msosptQuadArrowCallout:      Result := astQuadArrowCallout;
    else
      Result := astSun;
  end;
end;

function TDrwAutoShape.GetFillColor: TColor;
begin
  Result := FShape.OPT.FindValue(msooptFillColor);
end;

function TDrwAutoShape.GetLineColor: TColor;
begin
  Result := FShape.OPT.FindValue(msooptLineColor);
end;

procedure TDrwAutoShape.SetAutoShapeType(const Value: TAutoShapeType);
begin
  case Value of
    astSun:                      FShape.ShapeType := msosptSun;
    astMoon:                     FShape.ShapeType := msosptMoon;
    astLeftArrow:                FShape.ShapeType := msosptLeftArrow;
    astDownArrow:                FShape.ShapeType := msosptDownArrow;
    astUpArrow:                  FShape.ShapeType := msosptUpArrow;
    astLeftRightArrow:           FShape.ShapeType := msosptLeftRightArrow;
    astUpDownArrow:              FShape.ShapeType := msosptUpDownArrow;
    astQuadArrow:                FShape.ShapeType := msosptQuadArrow;
    astLeftArrowCallout:         FShape.ShapeType := msosptLeftArrowCallout;
    astRightArrowCallout:        FShape.ShapeType := msosptRightArrowCallout;
    astUpArrowCallout:           FShape.ShapeType := msosptUpArrowCallout;
    astDownArrowCallout:         FShape.ShapeType := msosptDownArrowCallout;
    astLeftRightArrowCallout:    FShape.ShapeType := msosptLeftRightArrowCallout;
    astUpDownArrowCallout:       FShape.ShapeType := msosptUpDownArrowCallout;
    astQuadArrowCallout:         FShape.ShapeType := msosptQuadArrowCallout;
    else
      FShape.ShapeType := msosptSun;
  end;
end;

procedure TDrwAutoShape.SetAutoShapeTypeInteger(Value: integer);
begin
  if (Value > 0) and (Value < msosptDoubleWave) then
    FShape.ShapeType := Value;
end;

procedure TDrwAutoShape.SetFillColor(const Value: TColor);
begin
  FShape.OPT.UpdateValue(msooptFillColor,Value)
end;

procedure TDrwAutoShape.SetLineColor(const Value: TColor);
begin
  FShape.OPT.UpdateValue(msooptLineColor,Value)
end;

{ TDrwAutoShapes }

function TDrwAutoShapes.Add: TDrwAutoShape;
begin
  Result := TDrwAutoShape(inherited Add);
end;

procedure TDrwAutoShapes.AddFromFile(Shape: TShapeClientAnchor);
begin
  FFileAdd := True;
  try
    TDrwAutoShape(inherited Add).FShape := Shape;
  finally
    FFileAdd := False;
  end;
end;

constructor TDrwAutoShapes.Create(AOwner: TPersistent; Drawing: TEscherDrawing);
begin
  inherited Create(TDrwAutoShape,AOwner,Drawing);
end;

function TDrwAutoShapes.GetDrwAutoShape(Index: integer): TDrwAutoShape;
begin
  Result := TDrwAutoShape(inherited Items[Index]);
end;

function TDrwAutoShapes.GetOwner: TPersistent;
begin
  Result := FOwner;
end;

{ TDrwPicture }

constructor TDrwPicture.Create(Collection: TCollection);
begin
  if not TDrwPictures(Collection).FFileAdd then
    FShape := TDrwPictures(Collection).FDrawing.AddPicture;
  inherited Create(Collection);
  FName := 'Picture ' + IntToStr(ID);
  SetMoveWithCells(False);
end;

function TDrwPicture.GetPictureId: integer;
begin
  Result := FShape.OPT.FindValue(msooptPib);
end;

function TDrwPicture.GetPictureName: AxUCString;
begin
  Result := FShape.OPT.FindString(msooptPibName);
end;

procedure TDrwPicture.SetPictureId(const Value: integer);
begin
  if TDrwPictures(Collection).FDrawing.EscherGroup.Count >  0 then begin
    if (Value < 0) or (Value > TDrwPictures(Collection).FDrawing.EscherGroup.Count) then
      raise XLSRWException.Create('Picture Id out of range.');

    if Value = 0 then
      FShape.OPT.UpdateString(msooptPibName,'')
    else
      FShape.OPT.UpdateString(msooptPibName,ExtractFilename(TDrwPictures(Collection).FDrawing.EscherGroup[Value - 1].Filename));
  end;
  FShape.OPT.UpdateValue(msooptPib,Value)
end;

procedure TDrwPicture.SetPictureName(const Value: AxUCString);
var
  i: integer;
  XLS: TBIFF5;
  Sheet: TSheet;
  Sheets: TSheets;
  DrawingObjects: TDrawingObjects;
begin
  DrawingObjects := TDrawingObjects(TDrwPictures(Collection).FOwner);
  Sheet := TSheet(DrawingObjects.FOwner);
  Sheets := TSheets(Sheet.Collection);
{$ifdef DELPHI_5}
  raise XLSRWException.Create('Can not do this in Delphi 5');
{$else}
  XLS := TBIFF5(Sheets.Owner);
{$endif}
  for i := 0 to XLS.MSOPictures.Count - 1 do begin
    if MyWideUppercase(Value) = MyWideUppercase(XLS.MSOPictures[i].Filename) then begin
      SetPictureId(XLS.MSOPictures[i].FileBlipId);
      Exit;
    end;
  end;
  raise XLSRWException.Create('Can not find picture: ' + Value);
end;

{ TDrwPictures }

function TDrwPictures.Add: TDrwPicture;
begin
  Result := TDrwPicture(inherited Add);
end;

procedure TDrwPictures.AddFromFile(Shape: TShapeClientAnchor);
begin
  FFileAdd := True;
  try
    TDrwPicture(inherited Add).FShape := Shape;
  finally
    FFileAdd := False;
  end;
end;

constructor TDrwPictures.Create(AOwner: TPersistent; Drawing: TEscherDrawing);
begin
  inherited Create(TDrwPicture,AOwner,Drawing);
end;

function TDrwPictures.GetDrwPicture(Index: integer): TDrwPicture;
begin
  Result := TDrwPicture(inherited Items[Index]);
end;

function TDrwPictures.GetOwner: TPersistent;
begin
  Result := FOwner;
end;

{ TDrwCollection }

procedure TDrwCollection.Clear;
var
  i: integer;
begin
  for i := 0 to Count - 1 do
    FDrawing.DeleteShape(TDrwAnchor(Items[i]).SpId);
  inherited Clear;
end;

constructor TDrwCollection.Create(ItemClass: TCollectionItemClass; AOwner: TPersistent; Drawing: TEscherDrawing);
begin
  inherited Create(ItemClass);
  FOwner := AOwner;
  FDrawing := Drawing;
end;

procedure TDrwCollection.Delete(Index: integer);
begin
  FDrawing.DeleteShape(TDrwAnchor(Items[Index]).SpId);
  inherited Delete(Index);
end;

function TDrwCollection.GetOwner: TPersistent;
begin
  Result := FOwner;
end;

{ TDrwMacro }

function TDrwCustom.GetFillColor: TColor;
begin
  Result := FShape.OPT.FindValue(msooptFillColor);
end;

function TDrwCustom.GetLineColor: TColor;
begin
  Result := FShape.OPT.FindValue(msooptLineColor);
end;

function TDrwCustom.GetMacroName: AxUCString;
{
var
  i: integer;
  P: PByteArray;
}
begin
  Result := '';
{
  i := FShape.OBJ.FindRecord(OBJREC_PICTFMLA);
  if i > 0 then begin
    P := @FShape.OBJ[i].Data;
  end;
}
end;

function TDrwCustom.GetMoveWithCells: boolean;
begin
  Result := (FShape.Anchor.Options or $0002) = 0;
end;

function TDrwCustom.GetRotation: double;
var
  i: integer;
  V: longword;
begin
  i := FShape.OPT.Find(msooptRotation);
  if i >= 0 then begin
    V := FShape.OPT.Value[i];
    Result := Smallint(V shr 16) + ((V and $0000FFFF) / 10000);
  end
  else
    Result := 0;
end;

procedure TDrwCustom.SetFillColor(const Value: TColor);
begin
  FShape.OPT.UpdateValue(msooptFillColor,Value);
end;

procedure TDrwCustom.SetLineColor(const Value: TColor);
begin
  FShape.OPT.UpdateValue(msooptLineColor,Value);
end;

procedure TDrwCustom.SetMacroName(const Value: AxUCString);
var
  i,Sz,NameIndex: integer;
  FH: TFormulaHandler;
  Macro: PByteArray;
  Header: TBIFFHeader;
begin
  FH := TDrawingObjects(TDrwCollection(Collection).FOwner).FormulaHandler;
  with FH.InternalNames.Add do begin
    Name := Value;
    IsVBProcedure := True;
    NameIndex := FH.InternalNames.Count - 1;
  end;
  FH.ExternalNames.AddSelf($FFFE,1);

  Sz := 6 + 1 + SizeOf(TPTGNameX8) + 1;
  GetMem(Macro,Sz);
  try
    FillChar(Macro^,Sz,#0);
    PWordArray(@Macro[0])[0] := 1 + SizeOf(TPTGNameX8);
    Macro[6] := xptgNameX97;

    PPTGNameX8(@Macro[7]).ExtSheet := 0;
    PPTGNameX8(@Macro[7]).NameIndex := NameIndex + 1;
    PPTGNameX8(@Macro[7]).Reserved := 0;

    i := FShape.OBJ.FindRecord(OBJREC_MACRO);
    if i > 0 then
      FShape.OBJ.UpdateRec(i,Sz,Macro)
    else begin
      FShape.OBJ.Delete(FShape.OBJ.Count - 1);
      Header.RecID := OBJREC_MACRO;
      Header.Length := Sz;
      FShape.OBJ.AddRec(Header,Macro);
      Header.RecID := OBJREC_END;
      Header.Length := 0;
      FShape.OBJ.AddRec(Header,Nil);
    end;
    FShape.MakeMacro;
//    i := FShape.OBJ.FindRecord(OBJREC_PIOGRBIT);
//    if i > 0 then begin
//      PWordArray(@FShape.OBJ[i].Data)[0] := $0001;
//      FShape.OBJ.UpdateRec(i,2,Macro)
//    end;
  finally
    FreeMem(Macro);
  end;
end;

procedure TDrwCustom.SetMoveWithCells(const Value: boolean);
begin
  if FShape <> Nil then begin
    if Value then
      FShape.Anchor.Options := FShape.Anchor.Options and ($FFFF - $0002)
    else
      FShape.Anchor.Options := FShape.Anchor.Options or $0002;
  end;
end;

procedure TDrwCustom.SetRotation(const Value: double);
var
  i: integer;
  V: longword;
begin
  if (Value < -180) or (Value > 180) then
    raise XLSRWException.Create('Rotation must be between -180 and 180');
  V := (Round(Int(Value)) shl 16) + Round(Frac(Value) * 10000);
  i := FShape.OPT.Find(msooptRotation);
  if i < 0 then begin
    FShape.OPT.AddValue(msooptRotation,V);
    FShape.OPT.SetItemCount;
  end
  else
    FShape.OPT.Value[i] := V;
end;

end.
