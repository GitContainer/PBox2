unit XLSFormattedObj5;

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
     vcl.Graphics,
{$endif}
     Xc12Utils5, Xc12Common5, Xc12DataStylesheet5, Xc12Manager5,
     XLSUtils5, XLSMMU5, XLSCellMMU5;

type TXLSFormattedCell = class(TObject)
private
     function  GetBorderBottomColor: TXc12Color;
     function  GetBorderBottomStyle: TXc12CellBorderStyle;
     function  GetBorderDiagColor: TXc12Color;
     function  GetBorderDiagStyle: TXc12CellBorderStyle;
     function  GetBorderLeftColor: TXc12Color;
     function  GetBorderLeftStyle: TXc12CellBorderStyle;
     function  GetBorderRightColor: TXc12Color;
     function  GetBorderRightStyle: TXc12CellBorderStyle;
     function  GetBorderTopColor: TXc12Color;
     function  GetBorderTopStyle: TXc12CellBorderStyle;
     function  GetFillPatternBackColor: TXc12Color;
     function  GetFillPatternForeColor: TXc12Color;
     function  GetFillPatternPattern: TXc12FillPattern;
     function  GetFontColor: TXc12Color;
     function  GetFontFamily: integer;
     function  GetFontName: AxUCString;
     function  GetFontSize: double;
     function  GetFontStyle: TXc12FontStyles;
     function  GetFontSubSuperScript: TXc12SubSuperscript;
     function  GetFontUnderline: TXc12Underline;
     function  GetHorizAlignment: TXc12HorizAlignment;
     function  GetIndent: integer;
     function  GetNumberFormat: AxUCString;
     function  GetProtection: TXc12CellProtections;
     function  GetRotation: integer;
     function  GetShrinkToFit: boolean;
     function  GetVertAlignment: TXc12VertAlignment;
     function  GetWrapText: boolean;
protected
     FXF: TXc12XF;
public
     procedure SetXF(AXF: TXc12XF);

     property Protection: TXc12CellProtections        read GetProtection;
     property HorizAlignment: TXc12HorizAlignment     read GetHorizAlignment;
     property VertAlignment: TXc12VertAlignment       read GetVertAlignment;
     property Indent: integer                         read GetIndent;
     property Rotation: integer                       read GetRotation;
     property WrapText: boolean                       read GetWrapText;
     property ShrinkToFit: boolean                    read GetShrinkToFit;
     property FillPatternForeColor: TXc12Color        read GetFillPatternForeColor;
     property FillPatternBackColor: TXc12Color        read GetFillPatternBackColor;
     property FillPatternPattern: TXc12FillPattern    read GetFillPatternPattern;
     property NumberFormat: AxUCString                read GetNumberFormat;
     property BorderTopColor:    TXc12Color           read GetBorderTopColor;
     property BorderTopStyle:    TXc12CellBorderStyle read GetBorderTopStyle;
     property BorderLeftColor:   TXc12Color           read GetBorderLeftColor;
     property BorderLeftStyle:   TXc12CellBorderStyle read GetBorderLeftStyle;
     property BorderRightColor:  TXc12Color           read GetBorderRightColor;
     property BorderRightStyle:  TXc12CellBorderStyle read GetBorderRightStyle;
     property BorderBottomColor: TXc12Color           read GetBorderBottomColor;
     property BorderBottomStyle: TXc12CellBorderStyle read GetBorderBottomStyle;
     property BorderDiagColor:   TXc12Color           read GetBorderDiagColor;
     property BorderDiagStyle:   TXc12CellBorderStyle read GetBorderDiagStyle;
     property FontName: AxUCString                    read GetFontName;
     property FontFamily: integer                     read GetFontFamily;
     property FontColor: TXc12Color                   read GetFontColor;
     property FontSize: double                        read GetFontSize;
     property FontStyle: TXc12FontStyles              read GetFontStyle;
     property FontSubSuperScript: TXc12SubSuperscript read GetFontSubSuperScript;
     property FontUnderline: TXc12Underline           read GetFontUnderline;
     end;

type TXLSFormattedObj = class(TObject)
private
     function  GetBorderBottomColor: TXc12IndexColor;
     function  GetBorderBottomColorRGB: TXc12RGBColor;
     function  GetBorderBottomStyle: TXc12CellBorderStyle;
     function  GetBorderDiagColor: TXc12IndexColor;
     function  GetBorderDiagColorRGB: TXc12RGBColor;
     function  GetBorderDiagStyle: TXc12CellBorderStyle;
     function  GetBorderLeftColor: TXc12IndexColor;
     function  GetBorderLeftColorRGB: TXc12RGBColor;
     function  GetBorderLeftStyle: TXc12CellBorderStyle;
     function  GetBorderRightColor: TXc12IndexColor;
     function  GetBorderRightColorRGB: TXc12RGBColor;
     function  GetBorderRightStyle: TXc12CellBorderStyle;
     function  GetBorderTopColor: TXc12IndexColor;
     function  GetBorderTopColorRGB: TXc12RGBColor;
     function  GetBorderTopStyle: TXc12CellBorderStyle;
     function  GetCellColorRGB: TXc12RGBColor;
     function  GetFillPatternBackColor: TXc12IndexColor;
     function  GetFillPatternBackColorRGB: TXc12RGBColor;
     function  GetFillPatternForeColor: TXc12IndexColor;
     function  GetFillPatternPattern: TXc12FillPattern;
     function  GetFillPatternForeColorRGB: TXc12RGBColor;
{$ifndef BABOON}
     function  GetFontCharset: TFontCharset;
{$endif}
     function  GetFontColor: TXc12RGBColor;
     function  GetFontFamily: integer;
     function  GetFontName: AxUCString;
     function  GetFontSize: double;
     function  GetFontSize20: integer;
     function  GetFontStyle: TXc12FontStyles;
     function  GetFontSubSuperScript: TXc12SubSuperscript;
     function  GetFontUnderline: TXc12Underline;
     function  GetHorizAlignment: TXc12HorizAlignment;
     function  GetIndent: integer;
     function  GetNumberFormat: AxUCString;
     function  GetProtection: TXc12CellProtections;
     function  GetRotation: integer;
     function  GetVertAlignment: TXc12VertAlignment;
     function  GetWrapText: boolean;
     function  GetShrinkToFit: boolean;

     procedure SetBorderBottomColor(const Value: TXc12IndexColor);
     procedure SetBorderBottomColorRGB(const Value: TXc12RGBColor);
     procedure SetBorderBottomStyle(const Value: TXc12CellBorderStyle);
     procedure SetBorderDiagColor(const Value: TXc12IndexColor);
     procedure SetBorderDiagColorRGB(const Value: TXc12RGBColor);
     procedure SetBorderDiagStyle(const Value: TXc12CellBorderStyle);
     procedure SetBorderLeftColor(const Value: TXc12IndexColor);
     procedure SetBorderLeftColorRGB(const Value: TXc12RGBColor);
     procedure SetBorderLeftStyle(const Value: TXc12CellBorderStyle);
     procedure SetBorderRightColor(const Value: TXc12IndexColor);
     procedure SetBorderRightColorRGB(const Value: TXc12RGBColor);
     procedure SetBorderRightStyle(const Value: TXc12CellBorderStyle);
     procedure SetBorderTopColor(const Value: TXc12IndexColor);
     procedure SetBorderTopColorRGB(const Value: TXc12RGBColor);
     procedure SetBorderTopStyle(const Value: TXc12CellBorderStyle);

     procedure SetCellColorRGB(const Value: TXc12RGBColor);

     procedure SetFillPatternBackColor(const Value: TXc12IndexColor);
     procedure SetFillPatternBackColorRGB(const Value: TXc12RGBColor);
     procedure SetFillPatternForeColorRGB(const Value: TXc12RGBColor);
     procedure SetFillPatternForeColor(const Value: TXc12IndexColor);
     procedure SetFillPatternPattern(const Value: TXc12FillPattern);
{$ifndef BABOON}
     procedure SetFontCharset(const Value: TFontCharset);
{$endif}
     procedure SetFontColor(const Value: TXc12RGBColor);
     procedure SetFontFamily(const Value: integer);
     procedure SetFontName(const Value: AxUCString);
     procedure SetFontSize(const Value: double);
     procedure SetFontSize20(const Value: integer);
     procedure SetFontStyle(const Value: TXc12FontStyles);
     procedure SetFontSubSuperScript(const Value: TXc12SubSuperscript);
     procedure SetFontUnderline(const Value: TXc12Underline);

     procedure SetProtection(const Value: TXc12CellProtections);

     procedure SetHorizAlignment(const Value: TXc12HorizAlignment);
     procedure SetIndent(const Value: integer);
     procedure SetNumberFormat(const Value: AxUCString);
     procedure SetRotation(const Value: integer);
     procedure SetVertAlignment(const Value: TXc12VertAlignment);
     procedure SetWrapText(const Value: boolean);
     procedure SetShrinkToFit(const Value: boolean);
     function  GetFontColorXc12: TXc12Color;
     procedure SetFontColorXc12(const Value: TXc12Color);
protected
     FStyles    : TXc12DataStyleSheet;
     FXF        : TXc12XF;

     procedure StyleChanged; virtual; abstract;
public
     constructor Create(AStyles: TXc12DataStyleSheet);
     destructor Destroy; override;

     //* Resets the format to default settings.
     procedure SetDefaultFormat;
     //* Copies the font (TXFont) of the format to a TFont object.
     //* ~param Font The TFont object to copy the properties to.
     procedure CopyToTFont(Font: TFont);
     //* Returns True if the object has formatting that differs from the default.
     //* ~result True if the object is formatted.
     function  IsFormatted: boolean;

     property Style: TXc12XF read FXF;

     //* Protection of cells.
     property Protection: TXc12CellProtections    read GetProtection           write SetProtection;
     //* Horizontal alignment of text in cells.
     property HorizAlignment: TXc12HorizAlignment read GetHorizAlignment       write SetHorizAlignment;
     //* Vertical alignment of text in cells.
     property VertAlignment: TXc12VertAlignment   read GetVertAlignment        write SetVertAlignment;
     //* Text indent
     //* The indent values can range from 0-15. A value of 0 (zero) is no
     //* indent. An increase with one increases the indent with about one character.
     property Indent: integer                     read GetIndent               write SetIndent;
     //* Rotation of cell text.
     //* Rotation, in degrees. ~[br]
     //* The value 0 ?90 is rotation up 0 ?90 deg. The value 91 ?180 is
     //* rotation down 1 ?90 deg. The value 255 is vertical text.
     property Rotation: integer                   read GetRotation             write SetRotation;

     property WrapText: boolean                   read GetWrapText             write SetWrapText;
     property ShrinkToFit: boolean                read GetShrinkToFit          write SetShrinkToFit;
     //* The foreground fill color.
     //* Use this property to set the color of the foreground fill pattern. ~[br]
     //* Note: This is the cell color. If you want to have just one color
     //* for the cell, use this property. ~[br]
     //* ~[link FillPatternBackColor] ~[link FillPatternPattern]
     property FillPatternForeColor: TXc12IndexColor    read GetFillPatternForeColor write SetFillPatternForeColor;
     property FillPatternForeColorRGB: TXc12RGBColor    read GetFillPatternForeColorRGB write SetFillPatternForeColorRGB;
     //* The background fill color.
     //* Use this property to set the color of the background fill pattern. ~[br]
     //* Note: If no fill pattern is assigned to the format, this
     //* property has no effect. If you just want to set the cell color, use
     //* FillPatternForeColor.
     //* ~[link FillPatternForeColor] ~[link FillPatternPattern]
     property FillPatternBackColor: TXc12IndexColor    read GetFillPatternBackColor write SetFillPatternBackColor;
     property FillPatternBackColorRGB: TXc12RGBColor    read GetFillPatternBackColorRGB write SetFillPatternBackColorRGB;
     //* The fill pattern for the cell.
     //* Use this property to set the fill pattern for the cell. Values can
     //* range between 0-127. 0 (zero) is no fill pattern.
     //* Fill patterns and there corresponding numbers, that are used in Excel:~[br]
     //* ~[image ..\help\PatternsNum.bmp]
     //* ~[link FillPatternForeColor] ~[link FillPatternBackColor]
     property CellColorRGB: TXc12RGBColor               read GetCellColorRGB         write SetCellColorRGB;
     property FillPatternPattern: TXc12FillPattern read GetFillPatternPattern write SetFillPatternPattern;
     //* Format mask used to format numbers.
     //* NumberFormat uses the same options as in the Format dilaog inExcel.~[br]
     //* NumberFormat is also used to create cells with date and time values,
     //* as there are no specific date or time cells in Excel.
     property NumberFormat: AxUCString            read GetNumberFormat         write SetNumberFormat;

     //* The color of the top cell border.
     //* ~[link TXc12CellBorderStyle]
     property BorderTopColor:    TXc12IndexColor      read GetBorderTopColor       write SetBorderTopColor;
     property BorderTopColorRGB: TXc12RGBColor        read GetBorderTopColorRGB    write SetBorderTopColorRGB;
     //* The style of the top cell border.
     property BorderTopStyle:    TXc12CellBorderStyle read GetBorderTopStyle       write SetBorderTopStyle;
     //* The color of the left cell border.
     property BorderLeftColor:   TXc12IndexColor      read GetBorderLeftColor      write SetBorderLeftColor;
     property BorderLeftColorRGB: TXc12RGBColor       read GetBorderLeftColorRGB   write SetBorderLeftColorRGB;
     //* The style of the left cell border.
     property BorderLeftStyle:   TXc12CellBorderStyle read GetBorderLeftStyle      write SetBorderLeftStyle;
     //* The color of the right cell border.
     property BorderRightColor:  TXc12IndexColor      read GetBorderRightColor     write SetBorderRightColor;
     property BorderRightColorRGB: TXc12RGBColor      read GetBorderRightColorRGB  write SetBorderRightColorRGB;
     //* The style of the right cell border.
     property BorderRightStyle:  TXc12CellBorderStyle read GetBorderRightStyle     write SetBorderRightStyle;
     //* The color of the bottom cell border.
     property BorderBottomColor: TXc12IndexColor      read GetBorderBottomColor    write SetBorderBottomColor;
     property BorderBottomColorRGB: TXc12RGBColor     read GetBorderBottomColorRGB write SetBorderBottomColorRGB;
     //* The style of the bottom cell border.
     property BorderBottomStyle: TXc12CellBorderStyle read GetBorderBottomStyle    write SetBorderBottomStyle;
     //* The color of the lines in a cell with diagonal lines.
     property BorderDiagColor:   TXc12IndexColor      read GetBorderDiagColor      write SetBorderDiagColor;
     property BorderDiagColorRGB: TXc12RGBColor       read GetBorderDiagColorRGB   write SetBorderDiagColorRGB;
     //* The style of diagonal lines in a cell.
     property BorderDiagStyle:   TXc12CellBorderStyle read GetBorderDiagStyle      write SetBorderDiagStyle;
     //* If the cell shall have diagonal lines.
//     property BorderDiagLines:   TDiagLines       read GetBorderDiagLines      write SetBorderDiagLines;

     //* The name of the font.
     property FontName: AxUCString read GetFontName write SetFontName;
     //* The character set of the font.
{$ifndef BABOON}
     property FontCharset: TFontCharset read GetFontCharset write SetFontCharset;
{$endif}
     //* The family of the font.
     //* The FontFamily can have the following values: ~[br]
     //* 0 = None (unknown/don't care). ~[br]
     //* 1 = Roman. ~[br]
     //* 2 = Swiss. ~[br]
     //* 3 = Modern (fixed width). ~[br]
     //* 4 = Script. ~[br]
     //* 5 = Decorative. ~[br]
     property FontFamily: integer read GetFontFamily write SetFontFamily;
     //* The color of the font.
     property FontColor: TXc12RGBColor read GetFontColor write SetFontColor;
     property FontColorXc12: TXc12Color read GetFontColorXc12 write SetFontColorXc12;
     //* The size of the font in points.
     //* ~[link FontSize20]
     property FontSize: double read GetFontSize write SetFontSize;
     //* Size of the font in units of 1/20th of a point.
     //* ~[link FontSize]
     property FontSize20: integer read GetFontSize20 write SetFontSize20;
     //* Font style.
     property FontStyle: TXc12FontStyles read GetFontStyle write SetFontStyle;
     //* Use FontSubSuperScript to set if the font shall be subscript or superscript.
     property FontSubSuperScript: TXc12SubSuperscript read GetFontSubSuperScript write SetFontSubSuperScript;
     //* Underline style of the font.
     property FontUnderline: TXc12Underline read GetFontUnderline write SetFontUnderline;

     property XF: TXc12XF read FXF;
     end;

     TXLSCell = class(TXLSFormattedObj)
private
     function GetItems(const ACol, ARow: integer): TXLSCell;
protected
     FCol  : integer;
     FRow  : integer;
     FCells: TXLSCellMMU;
     FRef  : TXLSCellItem;

     procedure StyleChanged; override;
public
     constructor Create(AStyles: TXc12DataStyleSheet; ACells: TXLSCellMMU);

     property Cells: TXLSCellMMU read FCells;
     property Items[const ACol,ARow: integer]: TXLSCell read GetItems; default;
     end;


implementation

{ TXLSFormattedObj }

procedure TXLSFormattedObj.CopyToTFont(Font: TFont);
begin

end;

constructor TXLSFormattedObj.Create(AStyles: TXc12DataStyleSheet);
begin
  FStyles := AStyles;
end;

destructor TXLSFormattedObj.Destroy;
begin
  inherited;
end;

function TXLSFormattedObj.GetBorderBottomColor: TXc12IndexColor;
begin
  if Xc12ColorIsIndexColor(FXF.Border.Bottom.Color) then
    Result := FXF.Border.Bottom.Color.Indexed
  else
    Result := XLSCOLOR_INDEX_NOT_INDEX;
end;

function TXLSFormattedObj.GetBorderBottomColorRGB: TXc12RGBColor;
begin
  Result := Xc12ColorToRGB(FXF.Border.Bottom.Color);
end;

function TXLSFormattedObj.GetBorderBottomStyle: TXc12CellBorderStyle;
begin
  Result := FXF.Border.Bottom.Style;
end;

function TXLSFormattedObj.GetBorderDiagColor: TXc12IndexColor;
begin
  if Xc12ColorIsIndexColor(FXF.Border.Diagonal.Color) then
    Result := FXF.Border.Diagonal.Color.Indexed
  else
    Result := XLSCOLOR_INDEX_NOT_INDEX;
end;

function TXLSFormattedObj.GetBorderDiagColorRGB: TXc12RGBColor;
begin
  Result := Xc12ColorToRGB(FXF.Border.Diagonal.Color);
end;

function TXLSFormattedObj.GetBorderDiagStyle: TXc12CellBorderStyle;
begin
  Result := FXF.Border.Diagonal.Style;
end;

function TXLSFormattedObj.GetBorderLeftColor: TXc12IndexColor;
begin
  if Xc12ColorIsIndexColor(FXF.Border.Left.Color) then
    Result := FXF.Border.Left.Color.Indexed
  else
    Result := XLSCOLOR_INDEX_NOT_INDEX;
end;

function TXLSFormattedObj.GetBorderLeftColorRGB: TXc12RGBColor;
begin
  Result := Xc12ColorToRGB(FXF.Border.Left.Color);
end;

function TXLSFormattedObj.GetBorderLeftStyle: TXc12CellBorderStyle;
begin
  Result := FXF.Border.Left.Style;
end;

function TXLSFormattedObj.GetBorderRightColor: TXc12IndexColor;
begin
  if Xc12ColorIsIndexColor(FXF.Border.Right.Color) then
    Result := FXF.Border.Right.Color.Indexed
  else
    Result := XLSCOLOR_INDEX_NOT_INDEX;
end;

function TXLSFormattedObj.GetBorderRightColorRGB: TXc12RGBColor;
begin
  Result := Xc12ColorToRGB(FXF.Border.Right.Color);
end;

function TXLSFormattedObj.GetBorderRightStyle: TXc12CellBorderStyle;
begin
  Result := FXF.Border.Right.Style;
end;

function TXLSFormattedObj.GetBorderTopColor: TXc12IndexColor;
begin
  if Xc12ColorIsIndexColor(FXF.Border.Top.Color) then
    Result := FXF.Border.Top.Color.Indexed
  else
    Result := XLSCOLOR_INDEX_NOT_INDEX;
end;

function TXLSFormattedObj.GetBorderTopColorRGB: TXc12RGBColor;
begin
  Result := Xc12ColorToRGB(FXF.Border.Top.Color);
end;

function TXLSFormattedObj.GetBorderTopStyle: TXc12CellBorderStyle;
begin
  Result := FXF.Border.Top.Style;
end;

function TXLSFormattedObj.GetCellColorRGB: TXc12RGBColor;
begin
  Result := GetFillPatternForeColorRGB;
end;

function TXLSFormattedObj.GetFillPatternBackColor: TXc12IndexColor;
begin
  if Xc12ColorIsIndexColor(FXF.Fill.BgColor) then
    Result := FXF.Fill.BgColor.Indexed
  else
    Result := XLSCOLOR_INDEX_NOT_INDEX;
end;

function TXLSFormattedObj.GetFillPatternBackColorRGB: TXc12RGBColor;
begin
  Result := Xc12ColorToRGB(FXF.Fill.BgColor);
end;

function TXLSFormattedObj.GetFillPatternForeColor: TXc12IndexColor;
begin
  if Xc12ColorIsIndexColor(FXF.Fill.FgColor) then
    Result := FXF.Fill.FgColor.Indexed
  else
    Result := XLSCOLOR_INDEX_NOT_INDEX;
end;

function TXLSFormattedObj.GetFillPatternForeColorRGB: TXc12RGBColor;
begin
  Result := Xc12ColorToRGB(FXF.Fill.FgColor);
end;

function TXLSFormattedObj.GetFillPatternPattern: TXc12FillPattern;
begin
  Result := FXF.Fill.PatternType;
end;

{$ifndef BABOON}
function TXLSFormattedObj.GetFontCharset: TFontCharset;
begin
  Result := FXF.Font.Charset;
end;
{$endif}

function TXLSFormattedObj.GetFontColor: TXc12RGBColor;
begin
  Result := Xc12ColorToRGB(FXF.Font.Color);
end;

function TXLSFormattedObj.GetFontColorXc12: TXc12Color;
begin
  Result := FXF.Font.Color;
end;

function TXLSFormattedObj.GetFontFamily: integer;
begin
  Result := FXF.Font.Family;
end;

function TXLSFormattedObj.GetFontName: AxUCString;
begin
  Result := FXF.Font.Name;
end;

function TXLSFormattedObj.GetFontSize: double;
begin
  Result := FXF.Font.Size;
end;

function TXLSFormattedObj.GetFontSize20: integer;
begin
  Result := Round(FXF.Font.Size * 20);
end;

function TXLSFormattedObj.GetFontStyle: TXc12FontStyles;
begin
  Result := FXF.Font.Style;
end;

function TXLSFormattedObj.GetFontSubSuperScript: TXc12SubSuperscript;
begin
  Result := FXF.Font.SubSuperscript;
end;

function TXLSFormattedObj.GetFontUnderline: TXc12Underline;
begin
  Result := FXF.Font.Underline;
end;

function TXLSFormattedObj.GetHorizAlignment: TXc12HorizAlignment;
begin
  Result := FXF.Alignment.HorizAlignment;
end;

function TXLSFormattedObj.GetIndent: integer;
begin
  Result := FXF.Alignment.Indent;
end;

function TXLSFormattedObj.GetNumberFormat: AxUCString;
begin
  Result := FXF.NumFmt.Value;
end;

function TXLSFormattedObj.GetProtection: TXc12CellProtections;
begin
  Result := FXF.Protection;
end;

function TXLSFormattedObj.GetRotation: integer;
begin
  Result := FXF.Alignment.Rotation;
end;

function TXLSFormattedObj.GetShrinkToFit: boolean;
begin
  Result := foShrinkToFit in FXF.Alignment.Options;
end;

function TXLSFormattedObj.GetVertAlignment: TXc12VertAlignment;
begin
  Result := FXF.Alignment.VertAlignment;
end;

function TXLSFormattedObj.GetWrapText: boolean;
begin
  Result := foWrapText in FXF.Alignment.Options;
end;

function TXLSFormattedObj.IsFormatted: boolean;
begin
  Result := FXF.Index <> XLS_STYLE_DEFAULT_XF;
end;

procedure TXLSFormattedObj.SetBorderBottomColor(const Value: TXc12IndexColor);
begin
  FStyles.XFEditor.BeginEdit(FXF);
  FStyles.XFEditor.BorderBottomColor := IndexColorToXc12(Value);
  FXF := FStyles.XFEditor.EndEdit;
  StyleChanged;
end;

procedure TXLSFormattedObj.SetBorderBottomColorRGB(const Value: TXc12RGBColor);
begin
  FStyles.XFEditor.BeginEdit(FXF);
  FStyles.XFEditor.BorderBottomColorRGB := Value;
  FXF := FStyles.XFEditor.EndEdit;
  StyleChanged;
end;

procedure TXLSFormattedObj.SetBorderBottomStyle(const Value: TXc12CellBorderStyle);
begin
  FStyles.XFEditor.BeginEdit(FXF);
  FStyles.XFEditor.BorderBottomStyle := Value;
  FXF := FStyles.XFEditor.EndEdit;
  StyleChanged;
end;

procedure TXLSFormattedObj.SetBorderDiagColor(const Value: TXc12IndexColor);
begin
  FStyles.XFEditor.BeginEdit(FXF);
  FStyles.XFEditor.BorderDiagColor := IndexColorToXc12(Value);
  FXF := FStyles.XFEditor.EndEdit;
  StyleChanged;
end;

procedure TXLSFormattedObj.SetBorderDiagColorRGB(const Value: TXc12RGBColor);
begin
  FStyles.XFEditor.BeginEdit(FXF);
  FStyles.XFEditor.BorderDiagColorRGB := Value;
  FXF := FStyles.XFEditor.EndEdit;
  StyleChanged;
end;

procedure TXLSFormattedObj.SetBorderDiagStyle(const Value: TXc12CellBorderStyle);
begin
  FStyles.XFEditor.BeginEdit(FXF);
  FStyles.XFEditor.BorderDiagStyle := Value;
  FXF := FStyles.XFEditor.EndEdit;
  StyleChanged;
end;

procedure TXLSFormattedObj.SetBorderLeftColor(const Value: TXc12IndexColor);
begin
  FStyles.XFEditor.BeginEdit(FXF);
  FStyles.XFEditor.BorderLeftColor := IndexColorToXc12(Value);
  FXF := FStyles.XFEditor.EndEdit;
  StyleChanged;
end;

procedure TXLSFormattedObj.SetBorderLeftColorRGB(const Value: TXc12RGBColor);
begin
  FStyles.XFEditor.BeginEdit(FXF);
  FStyles.XFEditor.BorderLeftColorRGB := Value;
  FXF := FStyles.XFEditor.EndEdit;
  StyleChanged;
end;

procedure TXLSFormattedObj.SetBorderLeftStyle(const Value: TXc12CellBorderStyle);
begin
  FStyles.XFEditor.BeginEdit(FXF);
  FStyles.XFEditor.BorderLeftStyle := Value;
  FXF := FStyles.XFEditor.EndEdit;
  StyleChanged;
end;

procedure TXLSFormattedObj.SetBorderRightColor(const Value: TXc12IndexColor);
begin
  FStyles.XFEditor.BeginEdit(FXF);
  FStyles.XFEditor.BorderRightColor := IndexColorToXc12(Value);
  FXF := FStyles.XFEditor.EndEdit;
  StyleChanged;
end;

procedure TXLSFormattedObj.SetBorderRightColorRGB(const Value: TXc12RGBColor);
begin
  FStyles.XFEditor.BeginEdit(FXF);
  FStyles.XFEditor.BorderRightColorRGB := Value;
  FXF := FStyles.XFEditor.EndEdit;
  StyleChanged;
end;

procedure TXLSFormattedObj.SetBorderRightStyle(const Value: TXc12CellBorderStyle);
begin
  FStyles.XFEditor.BeginEdit(FXF);
  FStyles.XFEditor.BorderRightStyle := Value;
  FXF := FStyles.XFEditor.EndEdit;
  StyleChanged;
end;

procedure TXLSFormattedObj.SetBorderTopColor(const Value: TXc12IndexColor);
begin
  FStyles.XFEditor.BeginEdit(FXF);
  FStyles.XFEditor.BorderTopColor := IndexColorToXc12(Value);
  FXF := FStyles.XFEditor.EndEdit;
  StyleChanged;
end;

procedure TXLSFormattedObj.SetBorderTopColorRGB(const Value: TXc12RGBColor);
begin
  FStyles.XFEditor.BeginEdit(FXF);
  FStyles.XFEditor.BorderTopColorRGB := Value;
  FXF := FStyles.XFEditor.EndEdit;
  StyleChanged;
end;

procedure TXLSFormattedObj.SetBorderTopStyle(const Value: TXc12CellBorderStyle);
begin
  FStyles.XFEditor.BeginEdit(FXF);
  FStyles.XFEditor.BorderTopStyle := Value;
  FXF := FStyles.XFEditor.EndEdit;
  StyleChanged;
end;

procedure TXLSFormattedObj.SetCellColorRGB(const Value: TXc12RGBColor);
begin
  SetFillPatternForeColorRGB(Value);
end;

procedure TXLSFormattedObj.SetDefaultFormat;
begin
  if FXF.Index <> XLS_STYLE_DEFAULT_XF then begin
    FStyles.XFEditor.FreeStyle(FXF);
    FXF := FStyles.XFs[XLS_STYLE_DEFAULT_XF];
    StyleChanged;
  end;
end;

procedure TXLSFormattedObj.SetFillPatternBackColor(const Value: TXc12IndexColor);
begin
  FStyles.XFEditor.BeginEdit(FXF);
  FStyles.XFEditor.FillBgColor := IndexColorToXc12(Value);
  FXF := FStyles.XFEditor.EndEdit;
  StyleChanged;
end;

procedure TXLSFormattedObj.SetFillPatternBackColorRGB(const Value: TXc12RGBColor);
begin
  FStyles.XFEditor.BeginEdit(FXF);
  FStyles.XFEditor.FillBgColorRGB := Value;
  FXF := FStyles.XFEditor.EndEdit;
  StyleChanged;
end;

procedure TXLSFormattedObj.SetFillPatternForeColor(const Value: TXc12IndexColor);
begin
  FStyles.XFEditor.BeginEdit(FXF);
  FStyles.XFEditor.FillColor := IndexColorToXc12(Value);
  FXF := FStyles.XFEditor.EndEdit;
  StyleChanged;
end;

procedure TXLSFormattedObj.SetFillPatternForeColorRGB(const Value: TXc12RGBColor);
begin
  FStyles.XFEditor.BeginEdit(FXF);
  FStyles.XFEditor.FillColorRGB := Value;
  FXF := FStyles.XFEditor.EndEdit;
  StyleChanged;
end;

procedure TXLSFormattedObj.SetFillPatternPattern(const Value: TXc12FillPattern);
begin
  FStyles.XFEditor.BeginEdit(FXF);
  FStyles.XFEditor.FillPattern := Value;
  FXF := FStyles.XFEditor.EndEdit;
  StyleChanged;
end;

{$ifndef BABOON}
procedure TXLSFormattedObj.SetFontCharset(const Value: TFontCharset);
begin
  FStyles.XFEditor.BeginEdit(FXF);
  FStyles.XFEditor.FontCharset := Value;
  FXF := FStyles.XFEditor.EndEdit;
  StyleChanged;
end;
{$endif}

procedure TXLSFormattedObj.SetFontColor(const Value: TXc12RGBColor);
begin
  FStyles.XFEditor.BeginEdit(FXF);
  FStyles.XFEditor.FontColorRGB := Value;
  FXF := FStyles.XFEditor.EndEdit;
  StyleChanged;
end;

procedure TXLSFormattedObj.SetFontColorXc12(const Value: TXc12Color);
begin
  FXF.Font.Color := Value;
end;

procedure TXLSFormattedObj.SetFontFamily(const Value: integer);
begin
  FStyles.XFEditor.BeginEdit(FXF);
  FStyles.XFEditor.FontFamily := Value;
  FXF := FStyles.XFEditor.EndEdit;
  StyleChanged;
end;

procedure TXLSFormattedObj.SetFontName(const Value: AxUCString);
begin
  FStyles.XFEditor.BeginEdit(FXF);
  FStyles.XFEditor.FontName := Value;
  FXF := FStyles.XFEditor.EndEdit;
  StyleChanged;
end;

procedure TXLSFormattedObj.SetFontSize(const Value: double);
begin
  FStyles.XFEditor.BeginEdit(FXF);
  FStyles.XFEditor.FontSize := Value;
  FXF := FStyles.XFEditor.EndEdit;
  StyleChanged;
end;

procedure TXLSFormattedObj.SetFontSize20(const Value: integer);
begin
  FStyles.XFEditor.BeginEdit(FXF);
  FStyles.XFEditor.FontSize := Value / 20;
  FXF := FStyles.XFEditor.EndEdit;
  StyleChanged;
end;

procedure TXLSFormattedObj.SetFontStyle(const Value: TXc12FontStyles);
begin
  FStyles.XFEditor.BeginEdit(FXF);
  FStyles.XFEditor.FontStyle := Value;
  FXF := FStyles.XFEditor.EndEdit;
  StyleChanged;
end;

procedure TXLSFormattedObj.SetFontSubSuperScript(const Value: TXc12SubSuperscript);
begin
  FStyles.XFEditor.BeginEdit(FXF);
  FStyles.XFEditor.FontSubSuperscript := Value;
  FXF := FStyles.XFEditor.EndEdit;
  StyleChanged;
end;

procedure TXLSFormattedObj.SetFontUnderline(const Value: TXc12Underline);
begin
  FStyles.XFEditor.BeginEdit(FXF);
  FStyles.XFEditor.FontUnderline := Value;
  FXF := FStyles.XFEditor.EndEdit;
  StyleChanged;
end;

procedure TXLSFormattedObj.SetHorizAlignment(const Value: TXc12HorizAlignment);
begin
  FStyles.XFEditor.BeginEdit(FXF);
  FStyles.XFEditor.AlignHoriz := Value;
  FXF := FStyles.XFEditor.EndEdit;
  StyleChanged;
end;

procedure TXLSFormattedObj.SetIndent(const Value: integer);
begin
  FStyles.XFEditor.BeginEdit(FXF);
  FStyles.XFEditor.AlignIndent := Value;
  FXF := FStyles.XFEditor.EndEdit;
  StyleChanged;
end;

procedure TXLSFormattedObj.SetNumberFormat(const Value: AxUCString);
begin
  FStyles.XFEditor.BeginEdit(FXF);
  FStyles.XFEditor.NumberFormat := Value;
  FXF := FStyles.XFEditor.EndEdit;
  StyleChanged;
end;

procedure TXLSFormattedObj.SetProtection(const Value: TXc12CellProtections);
begin
  FStyles.XFEditor.BeginEdit(FXF);
  FStyles.XFEditor.ProtectionLocked := cpLocked in Value;
  FStyles.XFEditor.ProtectionHidden := cpHidden in Value;
  FXF := FStyles.XFEditor.EndEdit;
  StyleChanged;
end;

procedure TXLSFormattedObj.SetRotation(const Value: integer);
begin
  FStyles.XFEditor.BeginEdit(FXF);
  FStyles.XFEditor.AlignRotation := Value;
  FXF := FStyles.XFEditor.EndEdit;
  StyleChanged;
end;

procedure TXLSFormattedObj.SetShrinkToFit(const Value: boolean);
var
  Opt: TXc12AlignmentOptions;
begin
  if Value then
    Opt := FXF.Alignment.Options + [foShrinkToFit]
  else
    Opt := FXF.Alignment.Options - [foShrinkToFit];

  FStyles.XFEditor.BeginEdit(FXF);
  FStyles.XFEditor.AlignOptions := Opt;
  FXF := FStyles.XFEditor.EndEdit;
  StyleChanged;
end;

procedure TXLSFormattedObj.SetVertAlignment(const Value: TXc12VertAlignment);
begin
  FStyles.XFEditor.BeginEdit(FXF);
  FStyles.XFEditor.AlignVert := Value;
  FXF := FStyles.XFEditor.EndEdit;
  StyleChanged;
end;

procedure TXLSFormattedObj.SetWrapText(const Value: boolean);
var
  Opt: TXc12AlignmentOptions;
begin
  if Value then
    Opt := FXF.Alignment.Options + [foWrapText]
  else
    Opt := FXF.Alignment.Options - [foWrapText];

  FStyles.XFEditor.BeginEdit(FXF);
  FStyles.XFEditor.AlignOptions := Opt;
  FXF := FStyles.XFEditor.EndEdit;
  StyleChanged;
end;

{ TXLSCell }

constructor TXLSCell.Create(AStyles: TXc12DataStyleSheet; ACells: TXLSCellMMU);
begin
  inherited Create(AStyles);

  FCells := ACells;
end;

function TXLSCell.GetItems(const ACol, ARow: integer): TXLSCell;
begin
  FCol := ACol;
  FRow := ARow;
  if FCells.FindCell(FCol,FRow,FRef) then begin
    FXF := FStyles.XFs[FCells.GetStyle(@FRef)];
    Result := Self;
  end
  else begin
    FXF := Nil;
    Result := Nil;
  end;
end;

procedure TXLSCell.StyleChanged;
begin
  if FRef.Data <> Nil then
    FCells.SetStyle(FCol,FRow,@FRef,FXF.Index);
end;

{ TXLSFormattedCell }

function TXLSFormattedCell.GetBorderBottomColor: TXc12Color;
begin
  Result := FXF.Border.Bottom.Color;
end;

function TXLSFormattedCell.GetBorderBottomStyle: TXc12CellBorderStyle;
begin
  Result := FXF.Border.Bottom.Style;
end;

function TXLSFormattedCell.GetBorderDiagColor: TXc12Color;
begin
  Result := FXF.Border.Diagonal.Color;
end;

function TXLSFormattedCell.GetBorderDiagStyle: TXc12CellBorderStyle;
begin
  Result := FXF.Border.Diagonal.Style;
end;

function TXLSFormattedCell.GetBorderLeftColor: TXc12Color;
begin
  Result := FXF.Border.Left.Color;
end;

function TXLSFormattedCell.GetBorderLeftStyle: TXc12CellBorderStyle;
begin
  Result := FXF.Border.Left.Style;
end;

function TXLSFormattedCell.GetBorderRightColor: TXc12Color;
begin
  Result := FXF.Border.Right.Color;
end;

function TXLSFormattedCell.GetBorderRightStyle: TXc12CellBorderStyle;
begin
  Result := FXF.Border.Right.Style;
end;

function TXLSFormattedCell.GetBorderTopColor: TXc12Color;
begin
  Result := FXF.Border.Top.Color;
end;

function TXLSFormattedCell.GetBorderTopStyle: TXc12CellBorderStyle;
begin
  Result := FXF.Border.Top.Style;
end;

function TXLSFormattedCell.GetFillPatternBackColor: TXc12Color;
begin
  Result := FXF.Fill.BgColor;
end;

function TXLSFormattedCell.GetFillPatternForeColor: TXc12Color;
begin
  Result := FXF.Fill.FgColor;
end;

function TXLSFormattedCell.GetFillPatternPattern: TXc12FillPattern;
begin
  Result := FXF.Fill.PatternType;
end;

function TXLSFormattedCell.GetFontColor: TXc12Color;
begin
  Result := FXF.Font.Color;
end;

function TXLSFormattedCell.GetFontFamily: integer;
begin
  Result := FXF.Font.Family;
end;

function TXLSFormattedCell.GetFontName: AxUCString;
begin
  Result := FXF.Font.Name;
end;

function TXLSFormattedCell.GetFontSize: double;
begin
  Result := FXF.Font.Size;
end;

function TXLSFormattedCell.GetFontStyle: TXc12FontStyles;
begin
  Result := FXF.Font.Style;
end;

function TXLSFormattedCell.GetFontSubSuperScript: TXc12SubSuperscript;
begin
  Result := FXF.Font.SubSuperscript;
end;

function TXLSFormattedCell.GetFontUnderline: TXc12Underline;
begin
  Result := FXF.Font.Underline;
end;

function TXLSFormattedCell.GetHorizAlignment: TXc12HorizAlignment;
begin
  Result := FXF.Alignment.HorizAlignment;
end;

function TXLSFormattedCell.GetIndent: integer;
begin
  Result := FXF.Alignment.Indent
end;

function TXLSFormattedCell.GetNumberFormat: AxUCString;
begin
  Result := FXF.NumFmt.Value;
end;

function TXLSFormattedCell.GetProtection: TXc12CellProtections;
begin
  Result := FXF.Protection;
end;

function TXLSFormattedCell.GetRotation: integer;
begin
  Result := FXF.Alignment.Rotation;
end;

function TXLSFormattedCell.GetShrinkToFit: boolean;
begin
  Result := foShrinkToFit in FXF.Alignment.Options;
end;

function TXLSFormattedCell.GetVertAlignment: TXc12VertAlignment;
begin
  Result := FXF.Alignment.VertAlignment;
end;

function TXLSFormattedCell.GetWrapText: boolean;
begin
  Result := FXF.Alignment.IsWrapText;
end;

procedure TXLSFormattedCell.SetXF(AXF: TXc12XF);
begin
  FXF := AXF;
end;

end.
