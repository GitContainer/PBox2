unit XLSRange5;

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
{$ifdef MSWINDOWS}
  {$ifndef BABOON}
     vcl.Graphics,
   {$endif}
{$endif}
     Xc12Utils5, Xc12DataStylesheet5, Xc12DataWorksheet5, Xc12Manager5,
     XLSUtils5, XLSCellAreas5, XLSFormattedObj5, XLSCmdFormat5;

//* What type of are the range covers.
type TXLSRangeType = (xrtCell,   //* A single cell.
                      xrtColumn, //* A whole column.
                      xrtRow,    //* A whole row.
                      xrtArea    //* An area.
                      );

//* Use TXLSRange to manipulate a whole area of cells. The area can be
//* formatted, merged, moved, copied etc. When formatting an area, blank
//* cells are added automatically if there are no cells in the area.
//* ~[br]
//* ~[br]
//* ~[b Example:]
//* ~[br]
//* ~[sample
//* // Write a string cell.
//* XLS.Sheet[0~[].AsStringRef['C2'~[] := 'Hello';
//* // Set the font size of the cells in the area.
//* XLS.Sheet[0~[].Range.Items[1,0,3,3~[].FontSize := 14;
//* // Set the color of the cells.
//* XLS.Sheet[0~[].Range.ItemsRef['B1:D4'~[].FillPatternForeColor := xcYellow;
//* // Set a outline border.
//* XLS.Sheet[0~[].Range.ItemsRef['B1:D4'~[].BorderOutlineStyle := cbsThick;
//* // Set color of the outline border.
//* XLS.Sheet[0~[].Range.ItemsRef['B1:D4'~[].BorderOutlineColor := xcRed;
//* // Make a copy of the cells.
//* XLS.Sheet[0~[].Range.ItemsRef['B1:D4'~[].Copy(8,10);
//* // Move the cells.
//* XLS.Sheet[0~[].Range.ItemsRef['B1:D4'~[].Move(8,2);
//* ]
type TXLSRange = class(TCellAreaImpl)
private
     FManager: TXc12Manager;
     FRangeType: TXLSRangeType;
     FSheet: TIndexObject;
     FCmdFormat: TXLSCmdFormat;

     procedure SetBorderInsideHorizColor(const Value: TXc12RGBColor);
     procedure SetBorderInsideHorizStyle(const Value: TXc12CellBorderStyle);
     procedure SetBorderInsideVertColor(const Value: TXc12RGBColor);
     procedure SetBorderInsideVertStyle(const Value: TXc12CellBorderStyle);
     procedure SetBorderOutlineColor(const Value: TXc12RGBColor);
     procedure SetBorderOutlineStyle(const Value: TXc12CellBorderStyle);
     procedure SetBorderBottomColor(const Value: TXc12RGBColor);
     procedure SetBorderBottomStyle(const Value: TXc12CellBorderStyle);
     procedure SetBorderDiagColor(const Value: TXc12RGBColor);
//     procedure SetBorderDiagLines(const Value: TDiagLines);
     procedure SetBorderDiagStyle(const Value: TXc12CellBorderStyle);
     procedure SetBorderLeftColor(const Value: TXc12RGBColor);
     procedure SetBorderLeftStyle(const Value: TXc12CellBorderStyle);
     procedure SetBorderRightColor(const Value: TXc12RGBColor);
     procedure SetBorderRightStyle(const Value: TXc12CellBorderStyle);
     procedure SetBorderTopColor(const Value: TXc12RGBColor);
     procedure SetBorderTopStyle(const Value: TXc12CellBorderStyle);
//     procedure SetFormatOptions(const Value: TFormatOptions);
     procedure SetFillPatternBackColor(const Value: TXc12RGBColor);
     procedure SetFillPatternForeColor(const Value: TXc12RGBColor);
     procedure SetFillPatternPattern(const Value: TXc12FillPattern);
{$ifndef BABOON}
     procedure SetFontCharset(const Value: TFontCharset);
{$endif}
     procedure SetFontColor(const Value: longword);
     procedure SetFontFamily(const Value: integer);
     procedure SetFontName(const Value: AxUCString);
     procedure SetFontSize(const Value: double);
     procedure SetFontSize20(const Value: integer);
     procedure SetFontSubSuperScript(const Value: TXc12SubSuperscript);
     procedure SetFontUnderline(const Value: TXc12Underline);
     procedure SetHorizAlignment(const Value: TXc12HorizAlignment);
     procedure SetIndent(const Value: integer);
     procedure SetNumberFormat(const Value: AxUCString);
     procedure SetProtection(const Value: TXc12CellProtections);
     procedure SetRotation(const Value: integer);
     procedure SetVertAlignment(const Value: TXc12VertAlignment);
     procedure SetFontStyle(const Value: TXc12FontStyles);

     function  GetItems(C1, R1, C2, R2: integer): TXLSRange;
     function  GetItemsRef(Ref: AxUCString): TXLSRange;
     procedure SetShrinkToFit(const Value: boolean);
     procedure SetWrapText(const Value: boolean);
protected
     function GetDefaultFormat(const ACol,ARow: integer): integer;
public
     //* ~exclude
     constructor Create(AManager: TXc12Manager; ASheet: TIndexObject);
     destructor Destroy; override;

     //* Protection of cells.
     property Protection: TXc12CellProtections        write SetProtection;
     //* Horizontal alignment of text in cells.
     property HorizAlignment: TXc12HorizAlignment write SetHorizAlignment;
     //* Vertical alignment of text in cells.
     property VertAlignment: TXc12VertAlignment   write SetVertAlignment;
     //* Text indent
     //* The indent values can range from 0-15. A value of 0 (zero) is no
     //* indent. An increment with one increases the indent with about one character.
     property Indent: integer                     write SetIndent;
     //* Rotation of cell text.
     //* Rotation, in degrees. ~[br]
     //* The value 0 ?90 is rotation up 0 ?90 deg. The value 91 ?180 is
     //* rotation down 1 ?90 deg. The value 255 is vertical text.
     property Rotation: integer                   write SetRotation;

     property WrapText: boolean                   write SetWrapText;
     property ShrinkToFit: boolean                write SetShrinkToFit;

     //* Cell format options.
//     property FormatOptions: TFormatOptions       write SetFormatOptions;
     //* The foreground fill color.
     //* Use this property to set the color of the foreground fill pattern. ~[br]
     //* Note: This is the cell color. If you want to have just one color
     //* for the cell, use this property. ~[br]
     //* ~[link FillPatternBackColor], ~[link FillPatternPattern]
     property FillPatternForeColor: TXc12RGBColor   write SetFillPatternForeColor;
     //* The background fill color.
     //* Use this property to set the color of the background fill pattern. ~[br]
     //* <b>Note:</b> If no fill pattern is assigned to the format, this
     //* property has no effect. If you just want to set the cell color, use
     //* FillPatternForeColor.
     //* ~[link FillPatternForeColor], ~[link FillPatternPattern]
     property FillPatternBackColor: TXc12RGBColor   write SetFillPatternBackColor;
     //* The fill pattern for the cell.
     //* Use this property to set the fill pattern for the cell. Values can
     //* range between 0-127. 0 (zero) is no fill pattern.
     //* Fill patterns and there corresponding numbers, that are used in Excel:~[br]
     //* ~[image ..\help\PatternsNum.bmp]
     //* ~[link FillPatternForeColor], ~[link FillPatternBackColor]
     property FillPatternPattern: TXc12FillPattern write SetFillPatternPattern;
     //* Format mask used to format numbers.
     //* NumberFormat uses the same options as in the Format dilaog inExcel.~[br]
     //* NumberFormat is also used to create cells with date and time values,
     //* as there are no specific date or time cells in Excel.
     property NumberFormat: AxUCString            write SetNumberFormat;

     //* The color of the top cell border.
     property BorderTopColor:    TXc12RGBColor      write SetBorderTopColor;
     //* The style of the top cell border;
     //* ~[link CellFormats4.TXc12CellBorderStyle TXc12CellBorderStyle]
     property BorderTopStyle:    TXc12CellBorderStyle write SetBorderTopStyle;
     //* The color of the left cell border.
     property BorderLeftColor:   TXc12RGBColor      write SetBorderLeftColor;
     //* The style of the left cell border;
     //* ~[link CellFormats4.TXc12CellBorderStyle TXc12CellBorderStyle]
     property BorderLeftStyle:   TXc12CellBorderStyle write SetBorderLeftStyle;
     //* The color of the right cell border.
     property BorderRightColor:  TXc12RGBColor      write SetBorderRightColor;
     //* The style of the right cell border;
     //* ~[link CellFormats4.TXc12CellBorderStyle TXc12CellBorderStyle]
     property BorderRightStyle:  TXc12CellBorderStyle write SetBorderRightStyle;
     //* The color of the bottom cell border.
     property BorderBottomColor: TXc12RGBColor      write SetBorderBottomColor;
     //* The style of the bottom cell border;
     //* ~[link CellFormats4.TXc12CellBorderStyle TXc12CellBorderStyle]
     property BorderBottomStyle: TXc12CellBorderStyle write SetBorderBottomStyle;
     //* The color of the lines in a cell with diagonal lines.
     property BorderDiagColor:   TXc12RGBColor      write SetBorderDiagColor;
     //* The style of diagonal lines in a cell;
     //* ~[link CellFormats4.TXc12CellBorderStyle TXc12CellBorderStyle]
     property BorderDiagStyle:   TXc12CellBorderStyle write SetBorderDiagStyle;
     //* If the cell shall have diagonal lines.
//     property BorderDiagLines:   TDiagLines       write SetBorderDiagLines;

     //* Set the outline border color of the cells in the area.
     property BorderOutlineColor:TXc12RGBColor      write SetBorderOutlineColor;
     //* Set the outline border style of the cells in the area.
     property BorderOutlineStyle:TXc12CellBorderStyle write SetBorderOutlineStyle;
     //* Set the color of vertical lines between the cells in the area. The outline of the
     //* cell area is not affected. See ~[link BorderOutlineColor] and ~[link BorderOutlineStyle]
     property BorderInsideVertColor:TXc12RGBColor   write SetBorderInsideVertColor;
     //* Set the style of vertical lines between the cells in the area. The outline of the
     //* cell area is not affected. See ~[link BorderOutlineColor] and ~[link BorderOutlineStyle]
     property BorderInsideVertStyle:TXc12CellBorderStyle write SetBorderInsideVertStyle;
     //* Set the color of horizontal lines between the cells in the area. The outline of the
     //* cell area is not affected. See ~[link BorderOutlineColor] and ~[link BorderOutlineStyle]
     property BorderInsideHorizColor:TXc12RGBColor   write SetBorderInsideHorizColor;
     //* Set the style of horizontal lines between the cells in the area. The outline of the
     //* cell area is not affected. See ~[link BorderOutlineColor] and ~[link BorderOutlineStyle]
     property BorderInsideHorizStyle:TXc12CellBorderStyle write SetBorderInsideHorizStyle;

     //* The name of the font.
     property FontName: AxUCString                write SetFontName;
     //* The character set of the font.
{$ifndef BABOON}
     property FontCharset: TFontCharset           write SetFontCharset;
{$endif}
     //* The family of the font.
     //* The FontFamily can have the following values:~[br]
     //* 0 = None (unknown/don't care).~[br]
     //* 1 = Roman.~[br]
     //* 2 = Swiss.~[br]
     //* 3 = Modern (fixed width).~[br]
     //* 4 = Script.~[br]
     //* 5 = Decorative.~[br]
     property FontFamily: integer                 write SetFontFamily;
     //* The color of the font.
     property FontColor: longword                 write SetFontColor;
     //* The size of the font in points.
     //* ~[link FontSize20]
     property FontSize: double                    write SetFontSize;
     //* Size of the font in units of 1/20th of a point.
     //* ~[link FontSize]
     property FontSize20: integer                 write SetFontSize20;
     //* Font style.
     property FontStyle: TXc12FontStyles             write SetFontStyle;
     //* Use FontSubSuperScript to set if the font shall be subscript or superscript.
     property FontSubSuperScript: TXc12SubSuperscript write SetFontSubSuperScript;
     //* Underline style of the font.
     property FontUnderline: TXc12Underline           write SetFontUnderline;
     //* Defines the cell area to manipulate.
     //* C1 = Left column. ~[br]
     //* R1 = Top row. ~[br]
     //* C2 = Right column. ~[br]
     //* R2 = bottom row. ~[br]
     property Items[C1,R1,C2,R2: integer]: TXLSRange read GetItems; default;
     //* Defines the cell area to manipulate.
     //* Ref = The area as a string, like: 'A1:D4'.
     property ItemsRef[Ref: AxUCString]: TXLSRange read GetItemsRef;
     end;

implementation


{ TXLSRange }

constructor TXLSRange.Create(AManager: TXc12Manager; ASheet: TIndexObject);
begin
  FManager := AManager;
  FCmdFormat := TXLSCmdFormat.Create(FManager);
  FSheet := ASheet;
end;

destructor TXLSRange.Destroy;
begin
  FCmdFormat.Free;

  inherited;
end;

function TXLSRange.GetDefaultFormat(const ACol, ARow: integer): integer;
begin
  Result := XLS_STYLE_DEFAULT_XF;
end;

function TXLSRange.GetItems(C1, R1, C2, R2: integer): TXLSRange;
begin
  if not ClipAreaToSheet(C1,R1,C2,R2) then
    raise XLSRWException.Create('Range is outside sheet limits');
  NormalizeArea(C1,R1,C2,R2);
  FCol1 := C1;
  FRow1 := R1;
  FCol2 := C2;
  FRow2 := R2;
  if (FCol1 = FCol2) and (FRow1 = FRow2) then
    FRangeType := xrtCell
  else if FCol1 = FCol2 then
    FRangeType := xrtColumn
  else if FRow1 = FRow2 then
    FRangeType := xrtRow
  else
    FRangeType := xrtArea;
  Result := Self;
end;

function TXLSRange.GetItemsRef(Ref: AxUCString): TXLSRange;
begin
  AreaStrToColRow(Ref,FCol1,FRow1,FCol2,FRow2);
  if (FCol1 = FCol2) and (FRow1 = FRow2) then
    FRangeType := xrtCell
  else if FCol1 = FCol2 then
    FRangeType := xrtColumn
  else if FRow1 = FRow2 then
    FRangeType := xrtRow
  else
    FRangeType := xrtArea;
  Result := Self;
end;

procedure TXLSRange.SetBorderOutlineColor(const Value: TXc12RGBColor);
begin
  FCmdFormat.BeginEdit(FSheet);
  FCmdFormat.Border.Color.RGB := Value;
  FCmdFormat.Border.Preset(cbspOutline);
  FCmdFormat.Apply(FCol1,FRow1,FCol2,FRow2);
end;

procedure TXLSRange.SetBorderOutlineStyle(const Value: TXc12CellBorderStyle);
begin
  FCmdFormat.BeginEdit(FSheet);
  FCmdFormat.Border.Style := Value;
  FCmdFormat.Border.Preset(cbspOutline);
  FCmdFormat.Apply(FCol1,FRow1,FCol2,FRow2);
end;

procedure TXLSRange.SetBorderBottomColor(const Value: TXc12RGBColor);
begin
  FCmdFormat.BeginEdit(FSheet);
  FCmdFormat.Border.Color.RGB := Value;
  FCmdFormat.Border.Side[cbsBottom] := True;
  FCmdFormat.Apply(FCol1,FRow1,FCol2,FRow2);
end;

procedure TXLSRange.SetBorderBottomStyle(const Value: TXc12CellBorderStyle);
begin
  FCmdFormat.BeginEdit(FSheet);
  FCmdFormat.Border.Style := Value;
  FCmdFormat.Border.Side[cbsBottom] := True;
  FCmdFormat.Apply(FCol1,FRow1,FCol2,FRow2);
end;

procedure TXLSRange.SetBorderDiagColor(const Value: TXc12RGBColor);
begin
  FCmdFormat.BeginEdit(FSheet);
  FCmdFormat.Border.Color.RGB := Value;
  FCmdFormat.Border.Side[cbsDiagLeft] := True;
  FCmdFormat.Border.Side[cbsDiagRight] := True;
  FCmdFormat.Apply(FCol1,FRow1,FCol2,FRow2);
end;

//procedure TXLSRange.SetBorderDiagLines(const Value: TDiagLines);
//var
//  C,R: integer;
//begin
//  AddBlanks;
//  for R := FRow1 to FRow2 do begin
//    for C := FCol1 to FCol2 do
//      FCells[C,R].BorderDiagLines := Value;
//  end;
//end;
//
procedure TXLSRange.SetBorderDiagStyle(const Value: TXc12CellBorderStyle);
begin
  FCmdFormat.BeginEdit(FSheet);
  FCmdFormat.Border.Style := Value;
  FCmdFormat.Border.Side[cbsDiagLeft] := True;
  FCmdFormat.Border.Side[cbsDiagRight] := True;
  FCmdFormat.Apply(FCol1,FRow1,FCol2,FRow2);
end;

procedure TXLSRange.SetBorderInsideHorizColor(const Value: TXc12RGBColor);
begin
  FCmdFormat.BeginEdit(FSheet);
  FCmdFormat.Border.Color.RGB := Value;
  FCmdFormat.Border.Side[cbsInsideHoriz] := True;
  FCmdFormat.Apply(FCol1,FRow1,FCol2,FRow2);
end;

procedure TXLSRange.SetBorderInsideHorizStyle(const Value: TXc12CellBorderStyle);
begin
  FCmdFormat.BeginEdit(FSheet);
  FCmdFormat.Border.Style := Value;
  FCmdFormat.Border.Side[cbsInsideHoriz] := True;
  FCmdFormat.Apply(FCol1,FRow1,FCol2,FRow2);
end;

procedure TXLSRange.SetBorderInsideVertColor(const Value: TXc12RGBColor);
begin
  FCmdFormat.BeginEdit(FSheet);
  FCmdFormat.Border.Color.RGB := Value;
  FCmdFormat.Border.Side[cbsInsideVert] := True;
  FCmdFormat.Apply(FCol1,FRow1,FCol2,FRow2);
end;

procedure TXLSRange.SetBorderInsideVertStyle(const Value: TXc12CellBorderStyle);
begin
  FCmdFormat.BeginEdit(FSheet);
  FCmdFormat.Border.Style := Value;
  FCmdFormat.Border.Side[cbsInsideVert] := True;
  FCmdFormat.Apply(FCol1,FRow1,FCol2,FRow2);
end;

procedure TXLSRange.SetBorderLeftColor(const Value: TXc12RGBColor);
begin
  FCmdFormat.BeginEdit(FSheet);
  FCmdFormat.Border.Color.RGB := Value;
  FCmdFormat.Border.Side[cbsLeft] := True;
  FCmdFormat.Apply(FCol1,FRow1,FCol2,FRow2);
end;

procedure TXLSRange.SetBorderLeftStyle(const Value: TXc12CellBorderStyle);
begin
  FCmdFormat.BeginEdit(FSheet);
  FCmdFormat.Border.Style := Value;
  FCmdFormat.Border.Side[cbsLeft] := True;
  FCmdFormat.Apply(FCol1,FRow1,FCol2,FRow2);
end;

procedure TXLSRange.SetBorderRightColor(const Value: TXc12RGBColor);
begin
  FCmdFormat.BeginEdit(FSheet);
  FCmdFormat.Border.Color.RGB := Value;
  FCmdFormat.Border.Side[cbsRight] := True;
  FCmdFormat.Apply(FCol1,FRow1,FCol2,FRow2);
end;

procedure TXLSRange.SetBorderRightStyle(const Value: TXc12CellBorderStyle);
begin
  FCmdFormat.BeginEdit(FSheet);
  FCmdFormat.Border.Style := Value;
  FCmdFormat.Border.Side[cbsRight] := True;
  FCmdFormat.Apply(FCol1,FRow1,FCol2,FRow2);
end;

procedure TXLSRange.SetBorderTopColor(const Value: TXc12RGBColor);
begin
  FCmdFormat.BeginEdit(FSheet);
  FCmdFormat.Border.Color.RGB := Value;
  FCmdFormat.Border.Side[cbsTop] := True;
  FCmdFormat.Apply(FCol1,FRow1,FCol2,FRow2);
end;

procedure TXLSRange.SetBorderTopStyle(const Value: TXc12CellBorderStyle);
begin
  FCmdFormat.BeginEdit(FSheet);
  FCmdFormat.Border.Style := Value;
  FCmdFormat.Border.Side[cbsTop] := True;
  FCmdFormat.Apply(FCol1,FRow1,FCol2,FRow2);
end;

//procedure TXLSRange.SetFormatOptions(const Value: TFormatOptions);
//var
//  C,R: integer;
//begin
//  AddBlanks;
//  for R := FRow1 to FRow2 do begin
//    for C := FCol1 to FCol2 do
//      FCell[C,R].FormatOptions := Value;
//  end;
//end;

procedure TXLSRange.SetFillPatternBackColor(const Value: TXc12RGBColor);
begin
  FCmdFormat.BeginEdit(FSheet);
  FCmdFormat.Fill.PatternColor.RGB := Value;
  FCmdFormat.Apply(FCol1,FRow1,FCol2,FRow2);
end;

procedure TXLSRange.SetFillPatternForeColor(const Value: TXc12RGBColor);
begin
  FCmdFormat.BeginEdit(FSheet);
  FCmdFormat.Fill.BackgroundColor.RGB := Value;
  FCmdFormat.Apply(FCol1,FRow1,FCol2,FRow2);
end;

procedure TXLSRange.SetFillPatternPattern(const Value: TXc12FillPattern);
begin
  FCmdFormat.BeginEdit(FSheet);
  FCmdFormat.Fill.PatternStyle := Value;
  FCmdFormat.Apply(FCol1,FRow1,FCol2,FRow2);
end;

{$ifndef BABOON}
procedure TXLSRange.SetFontCharset(const Value: TFontCharset);
begin
  FCmdFormat.BeginEdit(FSheet);
  FCmdFormat.Font.Charset := Value;
  FCmdFormat.Apply(FCol1,FRow1,FCol2,FRow2);
end;
{$endif}

procedure TXLSRange.SetFontColor(const Value: longword);
begin
  FCmdFormat.BeginEdit(FSheet);
  FCmdFormat.Font.Color.RGB := Value;
  FCmdFormat.Apply(FCol1,FRow1,FCol2,FRow2);
end;

procedure TXLSRange.SetFontFamily(const Value: integer);
begin
end;

procedure TXLSRange.SetFontName(const Value: AxUCString);
begin
  FCmdFormat.BeginEdit(FSheet);
  FCmdFormat.Font.Name := Value;
  FCmdFormat.Apply(FCol1,FRow1,FCol2,FRow2);
end;

procedure TXLSRange.SetFontSize(const Value: double);
begin
  FCmdFormat.BeginEdit(FSheet);
  FCmdFormat.Font.Size := Value;
  FCmdFormat.Apply(FCol1,FRow1,FCol2,FRow2);
end;

procedure TXLSRange.SetFontSize20(const Value: integer);
begin
  FCmdFormat.BeginEdit(FSheet);
  FCmdFormat.Font.Size := Value / 20;
  FCmdFormat.Apply(FCol1,FRow1,FCol2,FRow2);
end;

procedure TXLSRange.SetFontStyle(const Value: TXc12FontStyles);
begin
  FCmdFormat.BeginEdit(FSheet);
  FCmdFormat.Font.Style := Value;
  FCmdFormat.Apply(FCol1,FRow1,FCol2,FRow2);
end;

procedure TXLSRange.SetFontSubSuperScript(const Value: TXc12SubSuperscript);
begin
  FCmdFormat.BeginEdit(FSheet);
  FCmdFormat.Font.SubSuperscript := Value;
  FCmdFormat.Apply(FCol1,FRow1,FCol2,FRow2);
end;

procedure TXLSRange.SetFontUnderline(const Value: TXc12Underline);
begin
  FCmdFormat.BeginEdit(FSheet);
  FCmdFormat.Font.Underline := Value;
  FCmdFormat.Apply(FCol1,FRow1,FCol2,FRow2);
end;

procedure TXLSRange.SetHorizAlignment(const Value: TXc12HorizAlignment);
begin
  FCmdFormat.BeginEdit(FSheet);
  FCmdFormat.Alignment.Horizontal := Value;
  FCmdFormat.Apply(FCol1,FRow1,FCol2,FRow2);
end;

procedure TXLSRange.SetIndent(const Value: integer);
begin
  FCmdFormat.BeginEdit(FSheet);
  FCmdFormat.Alignment.Indent := Value;
  FCmdFormat.Apply(FCol1,FRow1,FCol2,FRow2);
end;

procedure TXLSRange.SetNumberFormat(const Value: AxUCString);
begin
  FCmdFormat.BeginEdit(FSheet);
  FCmdFormat.Number.Format := Value;
  FCmdFormat.Apply(FCol1,FRow1,FCol2,FRow2);
end;

procedure TXLSRange.SetProtection(const Value: TXc12CellProtections);
begin
  FCmdFormat.BeginEdit(FSheet);
  if cpLocked in Value then
    FCmdFormat.Protection.Locked := True;
  if cpHidden in Value then
    FCmdFormat.Protection.Hidden := True;
  FCmdFormat.Apply(FCol1,FRow1,FCol2,FRow2);
end;

procedure TXLSRange.SetRotation(const Value: integer);
begin
  FCmdFormat.BeginEdit(FSheet);
  FCmdFormat.Alignment.Rotation := Value;
  FCmdFormat.Apply(FCol1,FRow1,FCol2,FRow2);
end;

procedure TXLSRange.SetShrinkToFit(const Value: boolean);
begin
  FCmdFormat.BeginEdit(FSheet);
  FCmdFormat.Alignment.ShrinkToFit := Value;
  FCmdFormat.Apply(FCol1,FRow1,FCol2,FRow2);
end;

procedure TXLSRange.SetVertAlignment(const Value: TXc12VertAlignment);
begin
  FCmdFormat.BeginEdit(FSheet);
  FCmdFormat.Alignment.Vertical := Value;
  FCmdFormat.Apply(FCol1,FRow1,FCol2,FRow2);
end;

procedure TXLSRange.SetWrapText(const Value: boolean);
begin
  FCmdFormat.BeginEdit(FSheet);
  FCmdFormat.Alignment.WrapText := Value;
  FCmdFormat.Apply(FCol1,FRow1,FCol2,FRow2);
end;

end.
