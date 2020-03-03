unit XLSCmdFormat5;

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

uses Classes, SysUtils, Contnrs, IniFiles, Math, 
{$ifdef DELPHI_XE5_OR_LATER}
     UITypes,
{$endif}
{$ifdef BABOON}
  {$ifdef DELPHI_XE5_OR_LATER}
     FMX.Graphics,
  {$else}
     FMX.Types,
  {$endif}
{$else}
     Windows, vcl.Graphics,
{$endif}
     Xc12Utils5, Xc12Common5, Xc12DataStylesheet5, Xc12DataWorksheet5, Xc12Manager5,
     Xc12DataSST5,
     XLSUtils5, XLSMMU5, XLSCellMMU5, XLSColumn5, XLSRow5, XLSCellAreas5, XLSCmdFormatValues5;

type TXLSCmdFormatMode = (xcfmMerge,xcfmReplace);

type TCmdFmtXFProp   = (cfxfAlignHoriz,cfxfAlignVert,cfxfIndent,cfxfJustifyDist,cfxfWrapText,cfxfShrinkToFit,cfxfMergeCells,cfxfRotation,cfxfTextDirection);
type TCmdFmtFontProp = (cffpName,cffpSize,cffpColor,cffpStyle,cffpUnderline,cffpSubSuperscript,cffpCharset);
type TCmdFmtFillProp = (cfipBgColor,cfipPatColor,cfipPattern);
type TCmdFmtObject   = (cfoXFAlignment,cfoXFProtection,cfoNumFmt,cfoFill,cfoFont,cfoBorder);
type TCmdTriState    = (ctsTrue,ctsFalse,ctsUnassigned);

type TCmdBorderSide = (cbsTop,cbsLeft,cbsRight,cbsBottom,cbsInsideHoriz,cbsInsideVert,cbsDiagLeft,cbsDiagRight);
     TCmdBorderSides = set of TCmdBorderSide;
type TCmdBorderSidePreset = (cbspNone,cbspOutline,cbspInside,cbspOutlineAndInside);

type TXLSCmdFormatColor = class(TObject)
private
     function  GetRGB: longword;
     procedure SetRGB(const Value: TXc12RGBColor);
     function  GetIndexColor: TXc12IndexColor;
     procedure SetIndexColor(const Value: TXc12IndexColor);
     procedure SetTheme(AScheme: TXc12ClrSchemeColor; const Value: double);
     function  GetColor: TXc12Color;
     procedure SetColor(const Value: TXc12Color);
     function  GetTColor: TColor;
     procedure SetTColor(const Value: TColor);
protected
     FChanged: TNotifyEvent;

     FColor: PXc12Color;

     property Changed: TNotifyEvent read FChanged write FChanged;
public
     constructor Create(AColor: PXc12Color);

     procedure Clear;

     procedure ExcelSwatch(AColor,ATint: integer);

     property Color: TXc12Color read GetColor write SetColor;
     property TColor: TColor read GetTColor write SetTColor;
     property RGB: longword read GetRGB write SetRGB;
     property Theme[AScheme: TXc12ClrSchemeColor]: double write SetTheme;
     property IndexColor: TXc12IndexColor read GetIndexColor write SetIndexColor;
     end;

type TXLSCmdFormat = class;

     TXLSCmdFormatAlignment = class(TObject)
private
     function  GetHorizontal: TXc12HorizAlignment;
     function  GetIndent: integer;
     function  GetJustifyDistributed: boolean;
     function  GetMergeCells: boolean;
     function  GetRotation: integer;
     function  GetShrinkToFit: boolean;
     function  GetTextDirection: TXc12ReadOrder;
     function  GetVertical: TXc12VertAlignment;
     function  GetWrapText: boolean;
     procedure SetHorizontal(const Value: TXc12HorizAlignment);
     procedure SetIndent(const Value: integer);
     procedure SetJustifyDistributed(const Value: boolean);
     procedure SetMergeCells(const Value: boolean);
     procedure SetRotation(const Value: integer);
     procedure SetShrinkToFit(const Value: boolean);
     procedure SetTextDirection(const Value: TXc12ReadOrder);
     procedure SetVertical(const Value: TXc12VertAlignment);
     procedure SetWrapText(const Value: boolean);
protected
     FOwner     : TXLSCmdFormat;

     FXFs       : TXc12XFs;
     FXF        : TXc12XF;

     FMergeCells: boolean;

     FAssigneds : array[TCmdFmtXFProp] of boolean;

     procedure SetResult(ASrc,ADest: TXc12CellAlignment);
public
     constructor Create(AOwner: TXLSCmdFormat; AXFs: TXc12XFs);
     destructor Destroy; override;

     procedure Clear;

     property Horizontal: TXc12HorizAlignment read GetHorizontal write SetHorizontal;
     property Vertical: TXc12VertAlignment read GetVertical write SetVertical;
     property Indent: integer read GetIndent write SetIndent;
     property JustifyDistributed: boolean read GetJustifyDistributed write SetJustifyDistributed;
     property Rotation: integer read GetRotation write SetRotation;
     property WrapText: boolean read GetWrapText write SetWrapText;
     property ShrinkToFit: boolean read GetShrinkToFit write SetShrinkToFit;
     property MergeCells: boolean read GetMergeCells write SetMergeCells;
     property TextDirection: TXc12ReadOrder read GetTextDirection write SetTextDirection;
     end;

     TXLSNumFmtFractions = (xnffNone,xnffOneDigit,xnffTwoDigits,xnffThreeDigits,xnffAsHalves,xnffAsQuarters,
                            xnffAsEights,xnffAsSixtheenths,xnffAsTenths,xnffAsHundredths);

     TXLSCmdFormatNumber = class(TObject)
private
     function  GetFormat: AxUCString;
     procedure SetFormat(const Value: AxUCString);
     procedure SetDecimals(const Value: integer);
     procedure SetThousands(const Value: boolean);
     procedure SetNegColor(const Value: TColor);
     procedure SetNegColorSign(const Value: boolean);
     procedure SetCurrencySymbolLCID(const Value: integer);
     procedure SetPercentage(const Value: boolean);
     procedure SetFractions(const Value: TXLSNumFmtFractions);
     procedure SetScientific(const Value: boolean);
protected
     FOwner       : TXLSCmdFormat;

     FNumFmts     : TXc12NumberFormats;
     FNumFmt      : TXc12NumberFormat;

     FDecimals    : integer;
     FThousands   : boolean;
     FNegColor    : TColor;
     FNegColorSign: boolean;
     FCurrencyLCID: integer;
     FPercentage  : boolean;
     FFractions   : TXLSNumFmtFractions;
     FScientific  : boolean;

     FAssigned    : boolean;

     procedure SetResult(ADest: TXc12NumberFormat);
     function  GetDecimalsStr: AxUCString;
     procedure MakeFormatString;
public
     constructor Create(AOwner: TXLSCmdFormat; ANumFmts: TXc12NumberFormats);
     destructor Destroy; override;

     procedure Clear;

     property Format: AxUCString read GetFormat write SetFormat;
     property Decimals: integer read FDecimals write SetDecimals;
     property Thousands: boolean read FThousands write SetThousands;
     property NegativeColor: TColor read FNegColor write SetNegColor;
     // If True, sign is shown on negative numbers when they have a different
     // color set by NegativeColor.
     property NegativeColorShowSign: boolean read FNegColorSign write SetNegColorSign;
     // If LCID is 0, the currency symbol of the default locale is used.
     // If LCID is < 0, no currency symbol is used.
     property CurrencySymbolLCID: integer read FCurrencyLCID write SetCurrencySymbolLCID;
     property Percentage: boolean read FPercentage write SetPercentage;
     property Fractions: TXLSNumFmtFractions read FFractions write SetFractions;
     property Scientific: boolean read FScientific write SetScientific;
     end;

// Uses the notation from the Excel Format Cells dialog.
// Note that there is the cell color named Background color. This is the
// revere of what the cell properties uses, FillPatternForeColor. The cell
// color is stored as Foreground Color in the files. Confusing, isn't it?
     TXLSCmdFormatFill = class(TObject)
private
     procedure SetPatternStyle(const Value: TXc12FillPattern);
     function  GetPatternStyle: TXc12FillPattern;
     procedure SetNoColor(const Value: boolean);
protected
     FOwner    : TXLSCmdFormat;

     FFills    : TXc12Fills;
     FFill     : TXc12Fill;

     FBgColor  : TXLSCmdFormatColor;
     FPatColor : TXLSCmdFormatColor;

     FNoColor  : boolean;

     FAssigneds: array[TCmdFmtFillProp] of boolean;

     procedure BgColorChanged(ASender: TObject);
     procedure PatColorChanged(ASender: TObject);

     procedure SetResult(ASrc,ADest: TXc12Fill);
public
     constructor Create(AOwner: TXLSCmdFormat; AFills: TXc12Fills);
     destructor Destroy; override;

     procedure Clear;

     property BackgroundColor: TXLSCmdFormatColor read FBgColor;
     property PatternColor: TXLSCmdFormatColor read FPatColor;
     property PatternStyle: TXc12FillPattern read GetPatternStyle write SetPatternStyle;
     property NoColor: boolean read FNoColor write SetNoColor;
     end;

     TXLSCmdFormatFont = class(TObject)
private
     procedure SetCharset(const Value: integer);
     procedure SetName(const Value: AxUCString);
     procedure SetSize(const Value: double);
     procedure SetStyle(const Value: TXc12FontStyles);
     procedure SetSubSuperscript(const Value: TXc12SubSuperscript);
     procedure SetUnderline(const Value: TXc12Underline);
     function  GetCharset: integer;
     function  GetName: AxUCString;
     function  GetSize: double;
     function  GetStyle: TXc12FontStyles;
     function  GetSubSuperscript: TXc12SubSuperscript;
     function  GetUnderline: TXc12Underline;
     function  GetBold: boolean;
     function  GetItalic: boolean;
     procedure SetBold(const Value: boolean);
     procedure SetItalic(const Value: boolean);
protected
     FOwner     : TXLSCmdFormat;

     FFonts     : TXc12Fonts;
     FFont      : TXc12Font;
     FColor     : TXLSCmdFormatColor;

     FAssigneds : array[TCmdFmtFontProp] of boolean;

     procedure SetAssigned(const AValue: TCmdFmtFontProp);
     procedure Clear;
     procedure ColorChanged(ASender: TObject);

     procedure SetResult(ASrc,ADest: TXc12Font);
public
     constructor Create(AOwner: TXLSCmdFormat; AFonts: TXc12Fonts);
     destructor Destroy; override;

     procedure SetDefault;

     procedure ClearAssigned;
     procedure Assign(AFont: TFont); overload;
     procedure Assign(AFont: TXc12Font); overload;
     procedure CopyToTFont(ATFont: TFont);

     property Name: AxUCString read GetName write SetName;
     property Size: double read GetSize write SetSize;
     property Color: TXLSCmdFormatColor read FColor;
     property Style: TXc12FontStyles read GetStyle write SetStyle;
     property Bold: boolean read GetBold write SetBold;
     property Italic: boolean read GetItalic write SetItalic;
     property Underline: TXc12Underline read GetUnderline write SetUnderline;
     property SubSuperscript: TXc12SubSuperscript read GetSubSuperscript write SetSubSuperscript;
     property Charset: integer read GetCharset write SetCharset;
     end;

     TXLSCmdFormatBorder = class(TObject)
private
     function  GetSide(Index: TCmdBorderSide): boolean;
     procedure SetSide(Index: TCmdBorderSide; const Value: boolean);
protected
     FOwner   : TXLSCmdFormat;

     FOptions : TXc12CellBorderOptions;

     FStyle   : TXc12CellBorderStyle;
     FColor   : TXc12Color;

     FBorders : array[TCmdBorderSide] of TXc12BorderPr;
     FSides   : array[TCmdBorderSide] of TCmdTriState;

     FCmdColor: TXLSCmdFormatColor;

     procedure SetResult(ASrc,ADest: TXc12Border; const ASides: TCmdBorderSides);
     procedure SetResultPr(ASrc,ADest: TXc12BorderPr; const ASide: TCmdBorderSide);
public
     constructor Create(AOwner: TXLSCmdFormat);
     destructor Destroy; override;

     procedure Clear;

     procedure Preset(const ASides: TCmdBorderSidePreset);

     property Color: TXLSCmdFormatColor read FCmdColor;
     property Style: TXc12CellBorderStyle read FStyle write FStyle;
     property Side[Index: TCmdBorderSide]: boolean read GetSide write SetSide;
     end;

     TXLSCmdFormatProtection = class(TObject)
private
     function  GetHidden: boolean;
     function  GetLocked: boolean;
     procedure SetHidden(const Value: boolean);
     procedure SetLocked(const Value: boolean);
protected
     FOwner       : TXLSCmdFormat;

     FXFs         : TXc12XFs;
     FXF          : TXc12XF;

     FAssignLocked: TCmdTriState;
     FAssignHidden: TCmdTriState;

     procedure SetResult(ASrc,ADest: TXc12XF);
public
     constructor Create(AOwner: TXLSCmdFormat; AXFs: TXc12XFs);
     destructor Destroy; override;

     procedure Clear;

     property Locked: boolean read GetLocked write SetLocked;
     property Hidden: boolean read GetHidden write SetHidden;
     end;

     TXLSDefaultFormat = class(TObject)
protected
     FXF       : TXc12XF;
     FName     : AxUCString;
public
     procedure UseByDirectWrite(ACell: TXLSEventCell);

     property Name: AxUCString read FName write FName;
     property XF: TXc12XF read FXF;
     end;

     TXLSDefaultFormats = class(TObject)
private
     function GetItems(Index: integer): TXLSDefaultFormat;
protected
{$ifdef DELPHI_5}
     FItems: TStringList;
{$else}
     FItems: THashedStringList;
{$endif}

     function Add(AXF: TXc12XF; const AName: AxUCString): TXLSDefaultFormat;
public
     constructor Create;
     destructor Destroy; override;

     procedure Clear;
     function  Count: integer;
     function  Find(const AName: AxUCString): TXLSDefaultFormat;

     property Items[Index: integer]: TXLSDefaultFormat read GetItems; default;
     end;

     TXLSCmdFormat = class(TObject)
private
     procedure SetSetRowHeight(const Value: boolean);
protected
     FManager     : TXc12Manager;
     FStyles      : TXc12DataStyleSheet;
     FSheet       : TXc12DataWorksheet;
     FColumns     : TXLSColumns;
     FRows        : TXLSRows;
     FAreas       : TCellAreas;
     FCells       : TXLSCellMMU;

     FMode        : TXLSCmdFormatMode;
     FSetRowHeight: boolean;
     FMaxFont     : TXc12Font;

     FXF          : TXc12XF;

     FAlignment   : TXLSCmdFormatAlignment;
     FNumFmt      : TXLSCmdFormatNumber;
     FFont        : TXLSCmdFormatFont;
     FFill        : TXLSCmdFormatFill;
     FBorder      : TXLSCmdFormatBorder;
     FProtect     : TXLSCmdFormatProtection;

     FSearchXF    : TXc12XF;

     FAssigneds   : array[TCmdFmtObject] of boolean;

     FDefaults    : TXLSDefaultFormats;

     // Do not add blank cells when formatting rows and columns. Only update
     // format on existing cells.
     FAddBlank    : boolean;

     FCommands    : TXLSCmdFormatValues;

     function  MakeXF(ASrcXF: TXc12XF; const ASides: TCmdBorderSides): TXc12XF;
     procedure DoSSTRichString(ASSTStr: PXLSString);
     procedure DoReplaceCell(const ACol,ARow: integer; AXF: TXc12XF);
     procedure DoMergeCell(const ACol,ARow: integer; const ASides: TCmdBorderSides); overload;
     procedure DoMergeCell(ACell: PXLSCellItem; const ACol,ARow: integer; const ASides: TCmdBorderSides); overload;
     procedure DoApplyReplace(ACol1,ARow1,ACol2,ARow2: integer);
     procedure DoApplyReplaceBorder(ACol1,ARow1,ACol2,ARow2: integer);
     procedure DoApplyMerge(ACol1,ARow1,ACol2,ARow2: integer);
     procedure DoApplyCommands;
public
     constructor Create(AManager: TXc12Manager);
     destructor Destroy; override;

     // Defaults are not clear'd here.
     procedure Clear;
     procedure BeginEdit(ASheet: TIndexObject);

     //* Apply formatting to current selection.
     procedure Apply; overload;
     procedure Apply(const ACol,ARow: integer); overload;
     procedure Apply(ACol1,ARow1,ACol2,ARow2: integer); overload;
     procedure ApplyCols(ACol1,ACol2: integer);
     procedure ApplyRows(ARow1,ARow2: integer);
     procedure ApplySheet;

     function  AddAsDefault(const AName: AxUCString): TXLSDefaultFormat;

     property Mode: TXLSCmdFormatMode read FMode write FMode;
     //* If True row height is set to the height of the (max) font.
     //* Assumes that the row heigh not it manually set.
     property SetRowHeight: boolean read FSetRowHeight write SetSetRowHeight;

     property Alignment : TXLSCmdFormatAlignment read FAlignment;
     property Number    : TXLSCmdFormatNumber read FNumFmt;
     property Font      : TXLSCmdFormatFont read FFont;
     property Fill      : TXLSCmdFormatFill read FFill;
     property Border    : TXLSCmdFormatBorder read FBorder;
     property Protection: TXLSCmdFormatProtection read FProtect;

     property Defaults: TXLSDefaultFormats read FDefaults;

     property Commands: TXLSCmdFormatValues read FCommands;
     end;

implementation

{ TXLSCmdFormat }

procedure TXLSCmdFormat.Apply;
var
 i: integer;
 CA: TCellArea;
begin
  DoApplyCommands;

  for i := 0 to FAreas.Count - 1 do begin
    CA := FAreas[i];
    if CA.IsColumn then
      ApplyCols(CA.Col1,CA.Col2)
    else if CA.IsRow then
      ApplyRows(CA.Row1,CA.Row2)
    else
      Apply(CA.Col1,CA.Row1,CA.Col2,CA.Row2);
  end;
end;

procedure TXLSCmdFormat.BeginEdit(ASheet: TIndexObject);
begin
  FAddBlank := True;

  FAlignment.MergeCells := False;

  if ASheet <> Nil then begin
    // The idea with this is that TXLSCmdFormat don't has to know about TXLSSheet.
    // XLSSheetData5 can't be in the interface units section as that would cause a
    // circular ref. Having uses calues in the implementaion section is not an option.
    FSheet := TXc12DataWorksheet(ASheet.RequestObject(TXc12DataWorksheet));
    FColumns := TXLSColumns(ASheet.RequestObject(TXLSColumns));
    FRows := TXLSRows(ASheet.RequestObject(TXLSRows));
    FAreas := TCellAreas(ASheet.RequestObject(TCellAreas));
    FCells := FSheet.Cells;
  end;
  Clear;
end;

procedure TXLSCmdFormat.Apply(const ACol, ARow: integer);
begin
  Apply(ACol,ARow,ACol,ARow);
end;

procedure TXLSCmdFormat.Apply(ACol1, ARow1, ACol2, ARow2: integer);
var
  R: integer;
  H: double;
begin
  if FSheet = Nil then
    raise XLSRWException.Create('No worksheet defined fo Apply');

  DoApplyCommands;

  NormalizeArea(ACol1,ARow1,ACol2,ARow2);
  ClipAreaToExtent(ACol1,ARow1,ACol2,ARow2);

  if FMode = xcfmReplace then begin
    if FAssigneds[cfoBorder] then
      DoApplyReplaceBorder(ACol1, ARow1, ACol2, ARow2)
    else
      DoApplyReplace(ACol1, ARow1, ACol2, ARow2);
  end
  else
    DoApplyMerge(ACol1, ARow1, ACol2, ARow2);

  if (FRows <> Nil) and FSetRowHeight and (FMaxFont <> Nil) then begin
    for R := ARow1 to ARow2 do begin
      // This is not an exact calculation but is reasonably accurate.
      H := FMaxFont.Size + 4;
      if not FRows[R].CustomHeight and (H > FRows[R].HeightPt) then
        FRows[R].HeightPt := H;
    end;
  end;
end;

procedure TXLSCmdFormat.ApplyCols(ACol1, ACol2: integer);
var
  i,k: integer;
  j: TCmdBorderSide;
  C,R: integer;
  XF: TXc12XF;
  List: TXc12Columns;
  Sides: TCmdBorderSides;
  OldAddBlank: boolean;
  Cell: TXLSCellItem;
  Cols: array[0..XLS_MAXCOL] of word;
  ColCount: integer;
begin
  DoApplyCommands;

  if ACol1 < 0 then ACol1 := 0;
  if ACol2 < 0 then ACol2 := 0;
  if ACol1 > XLS_MAXCOL then ACol1 := XLS_MAXCOL;
  if ACol2 > XLS_MAXCOL then ACol2 := XLS_MAXCOL;

  Sides := [];
  for j := Low(TCmdBorderSide) to High(TCmdBorderSide) do begin
    if FBorder.FSides[j] = ctsTrue then
      Sides := Sides + [j];
  end;

  if FMode = xcfmReplace then begin
    XF := MakeXF(FStyles.XFs.DefaultXF,Sides);
    List := FColumns.BeginSetStyle(ACol1,ACol2,XF);
    if List <> Nil then begin
      for i := 0 to List.Count - 1 do begin
        FStyles.XFEditor.FreeStyle(List[i].Style);
        FStyles.XFEditor.UseStyle(XF);
        List[i].Style := XF;
      end;
    end;
  end
  else begin
    XF := MakeXF(FStyles.XFs.DefaultXF,Sides);
    List := FColumns.BeginSetStyle(ACol1,ACol2,XF);
    if List <> Nil then begin
      try
        for i := 0 to List.Count - 1 do begin
          XF := MakeXF(List[i].Style,Sides);
          FStyles.XFEditor.FreeStyle(List[i].Style);
          FStyles.XFEditor.UseStyle(XF);
          List[i].Style := XF;
        end;
        FColumns.EndSetStyle(List);
      finally
        List.Free;
      end;
    end;
  end;
  OldAddBlank := FAddBlank;
  FAddBlank := False;
  if FMode = xcfmReplace then begin
    Apply(ACol1,0,ACol2,XLS_MAXROW);
  end
  else begin
    FCells.CalcDimensions;
    i := FCells.Dimension.Row1;
    // Cache the columns as the iterated cell value not can be used as the
    // iterate data might be overwritten by DoMergeCell.
    while i <= FCells.Dimension.Row2 do begin
      ColCount := 0;
      FCells.BeginIterate(i);
      while FCells.IterateNextCol do begin
        Cols[ColCount] := FCells.IterCellCol;
        Inc(ColCount);
      end;

      i := Max(0,FCells.IterCellRow) + 1;

      for k := 0 to ColCount - 1 do begin
        C := Cols[k];
        R := i - 1;

        Cell := FCells.FindCell(C,R);

        if C = ACol1 then begin
          if R = 0 then
            DoMergeCell(@Cell,ACol1,0,[cbsTop,cbsLeft,cbsInsideVert,cbsInsideHoriz])
          else if R = XLS_MAXROW then
            DoMergeCell(@Cell,ACol1,XLS_MAXROW,[cbsLeft,cbsBottom,cbsInsideVert])
          else
            DoMergeCell(@Cell,ACol1,R,[cbsLeft,cbsInsideVert,cbsInsideHoriz]);
        end
        else if C = ACol2 then begin
          if R = 0 then
            DoMergeCell(@Cell,ACol2,0,[cbsTop,cbsRight,cbsInsideHoriz])
          else if R = XLS_MAXROW then
            DoMergeCell(@Cell,ACol2,XLS_MAXROW,[cbsRight,cbsBottom])
          else
            DoMergeCell(@Cell,ACol2,R,[cbsRight,cbsInsideHoriz]);
        end
        else if (C > ACol1) and (C < ACol2) then begin
          if R = 0 then
            DoMergeCell(@Cell,C,0,[cbsTop,cbsInsideVert,cbsInsideHoriz])
          else if R = XLS_MAXROW then
            DoMergeCell(@Cell,C,XLS_MAXROW,[cbsBottom,cbsInsideVert]);
          DoMergeCell(@Cell,C,R,[cbsInsideVert,cbsInsideHoriz]);
        end;
      end;
    end;
  end;
  FAddBlank := OldAddBlank;
end;

procedure TXLSCmdFormat.ApplyRows(ARow1, ARow2: integer);
var
  R: integer;
  i: TCmdBorderSide;
  XF: TXc12XF;
  Row: PXLSMMURowHeader;
  Sides: TCmdBorderSides;
begin
  DoApplyCommands;

  if ARow1 < 0 then ARow1 := 0;
  if ARow2 < 0 then ARow2 := 0;
  if ARow1 > XLS_MAXROW then ARow1 := XLS_MAXROW;
  if ARow2 > XLS_MAXROW then ARow2 := XLS_MAXROW;

  Sides := [];
  for i := Low(TCmdBorderSide) to High(TCmdBorderSide) do begin
    if FBorder.FSides[i] = ctsTrue then
      Sides := Sides + [i];
  end;

  if FMode = xcfmReplace then begin
    XF := MakeXF(FStyles.XFs.DefaultXF,Sides);
    for R := ARow1 to ARow2 do begin
      FStyles.XFEditor.UseStyle(XF);
      Row := FCells.FindRow(R);
      if Row <> Nil then begin
        FStyles.XFEditor.FreeStyle(Row.Style);
        Row.Style := XF.Index;
      end
      else
        FCells.AddRow(R,XF.Index);
    end;
  end
  else begin
    for R := ARow1 to ARow2 do begin
      Row := FCells.FindRow(R);
      if Row <> Nil then begin
        XF := MakeXF(FStyles.XFs[Row.Style],Sides);
        FStyles.XFEditor.FreeStyle(Row.Style);
        FStyles.XFEditor.UseStyle(XF);
        Row.Style := XF.Index;
      end
      else begin
        XF := MakeXF(FStyles.XFs.DefaultXF,Sides);
        FStyles.XFEditor.UseStyle(XF);
        FCells.AddRow(R,XF.Index);
      end;
    end;
  end;
  FAddBlank := False;
  Apply(0,ARow1,XLS_MAXCOL,ARow2);
  FAddBlank := True;
end;

procedure TXLSCmdFormat.ApplySheet;
begin
  if FManager.Version >= xvExcel2007 then
    ApplyCols(0,XLS_MAXCOL)
  else
    ApplyCols(0,XLS_MAXCOL_97);
end;

procedure TXLSCmdFormat.Clear;
begin
  FAlignment.Clear;
  FNumFmt.Clear;
  FFont.Clear;
  FFill.Clear;
  FBorder.Clear;
  FProtect.Clear;
  FXF.Assign(FStyles.XFs.DefaultXF);

  FMaxFont := Nil;
end;

constructor TXLSCmdFormat.Create(AManager: TXc12Manager);
begin
  FManager := AManager;
  FStyles := FManager.StyleSheet;

  FMode := xcfmMerge;

  FAlignment := TXLSCmdFormatAlignment.Create(Self,FStyles.XFs);
  FNumFmt := TXLSCmdFormatNumber.Create(Self,FStyles.NumFmts);
  FFont := TXLSCmdFormatFont.Create(Self,FStyles.Fonts);
  FFill := TXLSCmdFormatFill.Create(Self,FStyles.Fills);
  FBorder := TXLSCmdFormatBorder.Create(Self);
  FProtect := TXLSCmdFormatProtection.Create(Self,FStyles.XFs);

  FXF := TXc12XF.Create(Nil);

  FSearchXF := TXc12XF.Create(Nil);
  FSearchXF.NumFmt := TXc12NumberFormat.Create(Nil);
  FSearchXF.Font := TXc12Font.Create(Nil);
  FSearchXF.Fill := TXc12Fill.Create(Nil);
  FSearchXF.Border := TXc12Border.Create(Nil);

  FDefaults := TXLSDefaultFormats.Create;

  FCommands := TXLSCmdFormatValues.Create;

  Clear;
end;

destructor TXLSCmdFormat.Destroy;
begin
  FXF.Free;
  FAlignment.Free;
  FNumFmt.Free;
  FFont.Free;
  FFill.Free;
  FBorder.Free;
  FProtect.Free;

  FSearchXF.NumFmt.Free;
  FSearchXF.Font.Free;
  FSearchXF.Fill.Free;
  FSearchXF.Border.Free;
  FSearchXF.Free;

  FDefaults.Free;

  FCommands.Free;

  inherited;
end;

procedure TXLSCmdFormat.DoApplyCommands;
var
  i:  integer;
begin
  for i := 0 to FCommands.Count - 1 do begin
    case FCommands[i].Command of
      xcfcCellColor        : FFill.BackgroundColor.RGB := FCommands[i].AsInteger;
      xcfcBorderBottomColor: begin
        FBorder.Style := cbsMedium;
        FBorder.Color.RGB := FCommands[i].AsInteger;
        FBorder.Side[cbsBottom] := True;
      end;
      xcfcFontSize         : FFont.Size := FCommands[i].AsFloat;
      xcfcFontBold         : FFont.Style := FFont.Style + [xfsBold];
      xcfcFontItalic       : FFont.Style := FFont.Style + [xfsItalic];
      xcfcIndent           : FAlignment.Indent := FCommands[i].AsInteger;
    end;
  end;

  FCommands.Clear;
end;

procedure TXLSCmdFormat.DoApplyMerge(ACol1, ARow1, ACol2, ARow2: integer);
var
  C,R: integer;
  Cols,Rows: integer;
  Diag: TCmdBorderSides;
begin
//  for R := ARow1 to ARow2 do begin
//    for C := ACol1 to ACol2 do
//      DoMergeCell(C,R,[cbsTop,cbsLeft,cbsRight,cbsBottom,cbsInsideVert,cbsInsideHoriz])
//  end;
//  Exit;

  Cols := ACol2 - ACol1 + 1;
  Rows := ARow2 - ARow1 + 1;

 Diag := [];
 if FBorder.Side[cbsDiagLeft] then
   Diag := Diag + [cbsDiagLeft];
 if FBorder.Side[cbsDiagRight] then
   Diag := Diag + [cbsDiagRight];

// +-------+
// |   1   |
// +-------+
  if (Cols = 1) and (Rows = 1) then
    DoMergeCell(ACol1,ARow1,[cbsTop,cbsLeft,cbsRight,cbsBottom] + Diag)
// +-------+-------+
// |   1   |   3   |
// +-------+-------+
  else if (Cols = 2) and (Rows = 1) then begin
    DoMergeCell(ACol1,ARow1,[cbsTop,cbsLeft,cbsInsideVert,cbsBottom] + Diag);
    DoMergeCell(ACol2,ARow1,[cbsTop,cbsRight,cbsBottom] + Diag);
  end
// +-------+-------+-------+
// |   1   |   2   |   3   |
// +-------+-------+-------+
  else if (Cols > 2) and (Rows = 1) then begin
    DoMergeCell(ACol1,ARow1,[cbsTop,cbsLeft,cbsInsideVert,cbsBottom] + Diag);
    for C := ACol1 + 1 to ACol2 - 1 do
      DoMergeCell(C,ARow1,[cbsTop,cbsInsideVert,cbsBottom] + Diag);
    DoMergeCell(ACol2,ARow1,[cbsTop,cbsRight,cbsBottom] + Diag);
  end
// +-------+
// |   1   |
// +-------+
// |   7   |
// +-------+
  else if (Cols = 1) and (Rows = 2) then begin
    DoMergeCell(ACol1,ARow1,[cbsTop,cbsLeft,cbsRight,cbsInsideHoriz] + Diag);
    DoMergeCell(ACol1,ARow2,[cbsLeft,cbsRight,cbsBottom] + Diag);
  end
// +-------+
// |   1   |
// +-------+
// |   4   |
// +-------+
// |   7   |
// +-------+
  else if (Cols = 1) and (Rows > 2) then begin
    DoMergeCell(ACol1,ARow1,[cbsTop,cbsLeft,cbsRight,cbsInsideHoriz] + Diag);
    for R := ARow1 + 1 to ARow2 - 1 do
      DoMergeCell(ACol1,R,[cbsLeft,cbsRight,cbsInsideHoriz] + Diag);
    DoMergeCell(ACol1,ARow2,[cbsLeft,cbsRight,cbsBottom] + Diag);
  end
// +-------+-------+
// |   1   |   3   |
// +-------+-------+
// |   7   |   9   |
// +-------+-------+
  else if (Cols = 2) and (Rows = 2) then begin
    DoMergeCell(ACol1,ARow1,[cbsTop,cbsLeft,cbsInsideHoriz,cbsInsideVert] + Diag);
    DoMergeCell(ACol1,ARow2,[cbsLeft,cbsBottom,cbsInsideVert] + Diag);
    DoMergeCell(ACol2,ARow1,[cbsTop,cbsRight,cbsInsideHoriz] + Diag);
    DoMergeCell(ACol2,ARow2,[cbsRight,cbsBottom] + Diag);
  end
// +-------+-------+-------+
// |   1   |   2   |   3   |
// +-------+-------+-------+
// |   7   |   8   |   9   |
// +-------+-------+-------+
  else if (Cols > 2) and (Rows = 2) then begin
    DoMergeCell(ACol1,ARow1,[cbsTop,cbsLeft,cbsInsideHoriz,cbsInsideVert] + Diag);
    DoMergeCell(ACol1,ARow2,[cbsLeft,cbsBottom,cbsInsideVert] + Diag);
    for C := ACol1 + 1 to ACol2 - 1 do begin
      DoMergeCell(C,ARow1,[cbsTop,cbsInsideHoriz,cbsInsideVert] + Diag);
      DoMergeCell(C,ARow2,[cbsBottom,cbsInsideVert] + Diag);
    end;
    DoMergeCell(ACol2,ARow1,[cbsTop,cbsRight,cbsInsideHoriz] + Diag);
    DoMergeCell(ACol2,ARow2,[cbsRight,cbsBottom] + Diag);
  end
// +-------+-------+
// |   1   |   3   |
// +-------+-------+
// |   4   |   6   |
// +-------+-------+
// |   7   |   9   |
// +-------+-------+
  else if (Cols = 2) and (Rows > 2) then begin
    DoMergeCell(ACol1,ARow1,[cbsTop,cbsLeft,cbsInsideHoriz,cbsInsideVert] + Diag);
    for R := ARow1 + 1 to ARow2 - 1 do
      DoMergeCell(ACol1,R,[cbsLeft,cbsInsideVert,cbsInsideHoriz] + Diag);
    DoMergeCell(ACol1,ARow2,[cbsLeft,cbsBottom,cbsInsideVert] + Diag);
    DoMergeCell(ACol2,ARow1,[cbsTop,cbsRight,cbsInsideHoriz] + Diag);
    for R := ARow1 + 1 to ARow2 - 1 do
      DoMergeCell(ACol2,R,[cbsRight,cbsInsideHoriz] + Diag);
    DoMergeCell(ACol2,ARow2,[cbsRight,cbsBottom] + Diag);
  end
// +-------+-------+-------+
// |   1   |   2   |   3   |
// +-------+-------+-------+
// |   4   |   5   |   6   |
// +-------+-------+-------+
// |   7   |   8   |   9   |
// +-------+-------+-------+
  else begin
    DoMergeCell(ACol1,ARow1,[cbsTop,cbsLeft,cbsInsideVert,cbsInsideHoriz] + Diag);
    for R := ARow1 + 1 to ARow2 - 1 do
      DoMergeCell(ACol1,R,[cbsLeft,cbsInsideVert,cbsInsideHoriz] + Diag);
    DoMergeCell(ACol1,ARow2,[cbsLeft,cbsBottom,cbsInsideVert] + Diag);

    for C := ACol1 + 1 to ACol2 - 1 do
      DoMergeCell(C,ARow1,[cbsTop,cbsInsideVert,cbsInsideHoriz] + Diag);
    for R := ARow1 + 1 to ARow2 - 1 do begin
      for C := ACol1 + 1 to ACol2 - 1 do
        DoMergeCell(C,R,[cbsInsideVert,cbsInsideHoriz] + Diag);
    end;
    for C := ACol1 + 1 to ACol2 - 1 do
      DoMergeCell(C,ARow2,[cbsBottom,cbsInsideVert] + Diag);

    DoMergeCell(ACol2,ARow1,[cbsTop,cbsRight,cbsInsideHoriz] + Diag);
    for R := ARow1 + 1 to ARow2 - 1 do
      DoMergeCell(ACol2,R,[cbsRight,cbsInsideHoriz] + Diag);
    DoMergeCell(ACol2,ARow2,[cbsRight,cbsBottom] + Diag);
  end;
end;

procedure TXLSCmdFormat.DoApplyReplace(ACol1, ARow1, ACol2, ARow2: integer);
var
  C,R: integer;
  Cell: TXLSCellItem;
  XF: TXc12XF;
begin
  XF := MakeXF(FStyles.XFs.DefaultXF,[]);

  for R := ARow1 to ARow2 do begin
    for C := ACol1 to ACol2 do begin
      Cell := FCells.FindCell(C,R);

      if Cell.Data = Nil then begin
        if FAddBlank then
          FCells.StoreBlank(C,R,XF.Index);
      end
      else
        FCells.SetStyle(C,R,XF.Index);
    end;
  end;
end;

// ************  CELLS  ************
// +-------+-------+-------+-------+
// |   1   |   2   |   2   |   3   |
// +-------+-------+-------+-------+
// |   4   |   5   |   5   |   6   |
// +-------+-------+-------+-------+
// |   4   |   5   |   5   |   6   |
// +-------+-------+-------+-------+
// |   7   |   8   |   8   |   9   |
// +-------+-------+-------+-------+

procedure TXLSCmdFormat.DoApplyReplaceBorder(ACol1, ARow1, ACol2, ARow2: integer);
var
  C,R: integer;
  Cols,Rows: integer;
  XFs: array[1..9] of TXc12XF;
begin
  Cols := ACol2 - ACol1 + 1;
  Rows := ARow2 - ARow1 + 1;

// +-------+
// |   1   |
// +-------+
  if (Cols = 1) and (Rows = 1) then begin
    XFs[1] := MakeXF(FStyles.XFs.DefaultXF,[cbsTop,cbsLeft,cbsRight,cbsBottom]);

    DoReplaceCell(ACol1,ARow1,XFs[1]);
  end
// +-------+-------+
// |   1   |   3   |
// +-------+-------+
  else if (Cols = 2) and (Rows = 1) then begin
    XFs[1] := MakeXF(FStyles.XFs.DefaultXF,[cbsTop,cbsLeft,cbsInsideVert,cbsBottom]);
    XFs[3] := MakeXF(FStyles.XFs.DefaultXF,[cbsTop,cbsRight,cbsBottom]);

    DoReplaceCell(ACol1,ARow1,XFs[1]);
    DoReplaceCell(ACol2,ARow1,XFs[3]);
  end
// +-------+-------+-------+
// |   1   |   2   |   3   |
// +-------+-------+-------+
  else if (Cols > 2) and (Rows = 1) then begin
    XFs[1] := MakeXF(FStyles.XFs.DefaultXF,[cbsTop,cbsLeft,cbsInsideVert,cbsBottom]);
    XFs[2] := MakeXF(FStyles.XFs.DefaultXF,[cbsTop,cbsInsideVert,cbsBottom]);
    XFs[3] := MakeXF(FStyles.XFs.DefaultXF,[cbsTop,cbsRight,cbsBottom]);

    DoReplaceCell(ACol1,ARow1,XFs[1]);
    for C := ACol1 + 1 to ACol2 - 1 do
      DoReplaceCell(C,ARow1,XFs[2]);
    DoReplaceCell(ACol2,ARow1,XFs[3]);
  end
// +-------+
// |   1   |
// +-------+
// |   7   |
// +-------+
  else if (Cols = 1) and (Rows = 2) then begin
    XFs[1] := MakeXF(FStyles.XFs.DefaultXF,[cbsTop,cbsLeft,cbsRight,cbsInsideHoriz]);
    XFs[7] := MakeXF(FStyles.XFs.DefaultXF,[cbsLeft,cbsRight,cbsBottom]);

    DoReplaceCell(ACol1,ARow1,XFs[1]);
    DoReplaceCell(ACol1,ARow2,XFs[7]);
  end
// +-------+
// |   1   |
// +-------+
// |   4   |
// +-------+
// |   7   |
// +-------+
  else if (Cols = 1) and (Rows > 2) then begin
    XFs[1] := MakeXF(FStyles.XFs.DefaultXF,[cbsTop,cbsLeft,cbsRight,cbsInsideHoriz]);
    XFs[4] := MakeXF(FStyles.XFs.DefaultXF,[cbsLeft,cbsRight,cbsInsideHoriz]);
    XFs[7] := MakeXF(FStyles.XFs.DefaultXF,[cbsLeft,cbsRight,cbsBottom]);

    DoReplaceCell(ACol1,ARow1,XFs[1]);
    for R := ARow1 + 1 to ARow2 - 1 do
      DoReplaceCell(ACol1,R,XFs[4]);
    DoReplaceCell(ACol1,ARow2,XFs[7]);
  end
// +-------+-------+
// |   1   |   3   |
// +-------+-------+
// |   7   |   9   |
// +-------+-------+
  else if (Cols = 2) and (Rows = 2) then begin
    XFs[1] := MakeXF(FStyles.XFs.DefaultXF,[cbsTop,cbsLeft,cbsInsideHoriz,cbsInsideVert]);
    XFs[7] := MakeXF(FStyles.XFs.DefaultXF,[cbsLeft,cbsBottom,cbsInsideVert]);
    XFs[3] := MakeXF(FStyles.XFs.DefaultXF,[cbsTop,cbsRight,cbsInsideHoriz]);
    XFs[9] := MakeXF(FStyles.XFs.DefaultXF,[cbsRight,cbsBottom]);

    DoReplaceCell(ACol1,ARow1,XFs[1]);
    DoReplaceCell(ACol1,ARow2,XFs[7]);
    DoReplaceCell(ACol2,ARow1,XFs[3]);
    DoReplaceCell(ACol2,ARow2,XFs[9]);
  end
// +-------+-------+-------+
// |   1   |   2   |   3   |
// +-------+-------+-------+
// |   7   |   8   |   9   |
// +-------+-------+-------+
  else if (Cols > 2) and (Rows = 2) then begin
    XFs[1] := MakeXF(FStyles.XFs.DefaultXF,[cbsTop,cbsLeft,cbsInsideHoriz,cbsInsideVert]);
    XFs[7] := MakeXF(FStyles.XFs.DefaultXF,[cbsLeft,cbsBottom,cbsInsideVert]);
    XFs[2] := MakeXF(FStyles.XFs.DefaultXF,[cbsTop,cbsInsideHoriz,cbsInsideVert]);
    XFs[8] := MakeXF(FStyles.XFs.DefaultXF,[cbsBottom,cbsInsideVert]);
    XFs[3] := MakeXF(FStyles.XFs.DefaultXF,[cbsTop,cbsRight,cbsInsideHoriz]);
    XFs[9] := MakeXF(FStyles.XFs.DefaultXF,[cbsRight,cbsBottom]);

    DoReplaceCell(ACol1,ARow1,XFs[1]);
    DoReplaceCell(ACol1,ARow2,XFs[7]);
    for C := ACol1 + 1 to ACol2 - 1 do begin
      DoReplaceCell(C,ARow1,XFs[2]);
      DoReplaceCell(C,ARow2,XFs[8]);
    end;
    DoReplaceCell(ACol2,ARow1,XFs[3]);
    DoReplaceCell(ACol2,ARow2,XFs[9]);
  end
// +-------+-------+
// |   1   |   3   |
// +-------+-------+
// |   4   |   6   |
// +-------+-------+
// |   7   |   9   |
// +-------+-------+
  else if (Cols = 2) and (Rows > 2) then begin
    XFs[1] := MakeXF(FStyles.XFs.DefaultXF,[cbsTop,cbsLeft,cbsInsideHoriz,cbsInsideVert]);
    XFs[4] := MakeXF(FStyles.XFs.DefaultXF,[cbsLeft,cbsInsideVert,cbsInsideHoriz]);
    XFs[7] := MakeXF(FStyles.XFs.DefaultXF,[cbsLeft,cbsBottom,cbsInsideVert]);
    XFs[3] := MakeXF(FStyles.XFs.DefaultXF,[cbsTop,cbsRight,cbsInsideHoriz]);
    XFs[6] := MakeXF(FStyles.XFs.DefaultXF,[cbsRight,cbsInsideHoriz]);
    XFs[9] := MakeXF(FStyles.XFs.DefaultXF,[cbsRight,cbsBottom]);

    DoReplaceCell(ACol1,ARow1,XFs[1]);
    DoReplaceCell(ACol2,ARow1,XFs[3]);
    for R := ARow1 + 1 to ARow2 - 1 do begin
      DoReplaceCell(ACol1,R,XFs[4]);
      DoReplaceCell(ACol2,R,XFs[6]);
    end;
    DoReplaceCell(ACol1,ARow2,XFs[7]);
    DoReplaceCell(ACol2,ARow2,XFs[9]);
  end
// +-------+-------+-------+
// |   1   |   2   |   3   |
// +-------+-------+-------+
// |   4   |   5   |   6   |
// +-------+-------+-------+
// |   7   |   8   |   9   |
// +-------+-------+-------+
  else begin
    XFs[1] := MakeXF(FStyles.XFs.DefaultXF,[cbsTop,cbsLeft,cbsInsideVert,cbsInsideHoriz]);
    XFs[4] := MakeXF(FStyles.XFs.DefaultXF,[cbsLeft,cbsInsideVert,cbsInsideHoriz]);
    XFs[7] := MakeXF(FStyles.XFs.DefaultXF,[cbsLeft,cbsBottom,cbsInsideVert]);

    XFs[2] := MakeXF(FStyles.XFs.DefaultXF,[cbsTop,cbsInsideVert,cbsInsideHoriz]);
    XFs[5] := MakeXF(FStyles.XFs.DefaultXF,[cbsInsideVert,cbsInsideHoriz]);
    XFs[8] := MakeXF(FStyles.XFs.DefaultXF,[cbsBottom,cbsInsideVert]);

    XFs[3] := MakeXF(FStyles.XFs.DefaultXF,[cbsTop,cbsRight,cbsInsideHoriz]);
    XFs[6] := MakeXF(FStyles.XFs.DefaultXF,[cbsRight,cbsInsideHoriz]);
    XFs[9] := MakeXF(FStyles.XFs.DefaultXF,[cbsRight,cbsBottom]);

    DoReplaceCell(ACol1,ARow1,XFs[1]);

    for R := ARow1 + 1 to ARow2 - 1 do
      DoReplaceCell(ACol1,R,XFs[4]);

    DoReplaceCell(ACol1,ARow2,XFs[7]);

    for C := ACol1 + 1 to ACol2 - 1 do
      DoReplaceCell(C,ARow1,XFs[2]);

    DoReplaceCell(ACol2,ARow1,XFs[3]);

    for R := ARow1 + 1 to ARow2 - 1 do
      DoReplaceCell(ACol2,R,XFs[6]);

    DoReplaceCell(ACol2,ARow2,XFs[9]);

    for C := ACol1 + 1 to ACol2 - 1 do
      DoReplaceCell(C,ARow2,XFs[8]);

    for R := ARow1 + 1 to ARow2 - 1 do begin
      for C := ACol1 + 1 to ACol2 - 1 do
        DoReplaceCell(C,R,XFs[5]);
    end;
  end;
end;

procedure TXLSCmdFormat.DoMergeCell(ACell: PXLSCellItem; const ACol,ARow: integer; const ASides: TCmdBorderSides);
var
  i       : integer;
  XF      : TXc12XF;
  P       : PXLSString;
begin
  if ACell.Data <> Nil then begin
    XF := MakeXF(FStyles.XFs[FCells.GetStyle(ACell)],ASides);

    if FAssigneds[cfoFont] and (FCells.CellType(ACell) = xctString) then begin
      i := FCells.GetStringSST(ACell);
      P := FManager.SST[i];
      if SSTStringIsRichStr(P) then
        DoSSTRichString(P);
    end;

    FCells.SetStyle(ACol,ARow,ACell,XF.Index);
  end
  else if FAddBlank then begin
    XF := MakeXF(FStyles.XFs.DefaultXF,ASides);
    FCells.StoreBlank(ACol,ARow,XF.Index)
  end;
end;

procedure TXLSCmdFormat.DoMergeCell(const ACol, ARow: integer; const ASides: TCmdBorderSides);
var
  i       : integer;
  XF      : TXc12XF;
  Cell    : TXLSCellItem;
  P       : PXLSString;

function GetDefaultFormat(const ACol, ARow: integer): TXc12XF;
var
  R: PXLSMMURowHeader;
begin
  if not FColumns[ACol].Style.IsDefault then
    Result := FStyles.XFs[FColumns[ACol].Style.Index]
  else begin
    R := FCells.FindRow(ARow);
    if (R <> Nil) and (R.Style <> XLS_STYLE_DEFAULT_XF) then
      Result := FStyles.XFs[R.Style]
    else
      Result := FStyles.XFs[XLS_STYLE_DEFAULT_XF];
  end;
end;

begin
  Cell := FCells.FindCell(ACol,ARow);
  if Cell.Data <> Nil then begin
    XF := MakeXF(FStyles.XFs[FCells.GetStyle(@Cell)],ASides);

    if FAssigneds[cfoFont] and (FCells.CellType(@Cell) = xctString) then begin
      i := FCells.GetStringSST(@Cell);
      P := FManager.SST[i];
      if SSTStringIsRichStr(P) then
        DoSSTRichString(P);
    end;

    FCells.SetStyle(ACol,ARow,XF.Index);
  end
  else if FAddBlank then begin
    XF := MakeXF(GetDefaultFormat(ACol,ARow),ASides);
//    XF := MakeXF(FStyles.XFs.DefaultXF,ASides);
    FCells.StoreBlank(ACol,ARow,XF.Index)
  end;
end;

procedure TXLSCmdFormat.DoReplaceCell(const ACol, ARow: integer; AXF: TXc12XF);
var
  Cell: TXLSCellItem;
begin
  Cell := FCells.FindCell(ACol,ARow);
  if Cell.Data = Nil then begin
    if FAddBlank then
      FCells.StoreBlank(ACol,ARow,AXF.Index);
  end
  else
    FCells.SetStyle(ACol,ARow,AXF.Index);
end;

procedure TXLSCmdFormat.DoSSTRichString(ASSTStr: PXLSString);
var
  i       : integer;
  Cnt     : integer;
  FontRuns: PXc12FontRunArray;
  Font    : TXc12Font;
begin
  Cnt := FManager.SST.GetFontRunsCount(ASSTStr);
  FontRuns := FManager.SST.GetFontRuns(ASSTStr);

  for i := 0 to Cnt - 1 do begin
    FFont.SetResult(FontRuns[i].Font,FSearchXF.Font);

    Font := FStyles.Fonts.Find(FSearchXF.Font);
    if Font = Nil then begin
      Font := FStyles.Fonts.Add;
      Font.Assign(FSearchXF.Font);
    end;
    FManager.StyleSheet.XFEditor.FreeFont(FontRuns[i].Font);
    FManager.StyleSheet.XFEditor.UseFont(Font);
    FontRuns[i].Font := Font;
  end;
end;

function TXLSCmdFormat.MakeXF(ASrcXF: TXc12XF; const ASides: TCmdBorderSides): TXc12XF;
begin
  if FAssigneds[cfoNumFmt] then
    FNumFmt.SetResult(FSearchXF.NumFmt)
  else
    FSearchXF.NumFmt.Assign(ASrcXF.NumFmt);

  if FAssigneds[cfoXFAlignment] then
    FAlignment.SetResult(ASrcXF.Alignment,FSearchXF.Alignment)
  else
    FSearchXF.Alignment.Assign(ASrcXF.Alignment);

  if FAssigneds[cfoFill] then begin
    if FFill.NoColor then
      FSearchXF.Fill.Assign(FStyles.Fills.DefaultFill)
    else
      FFill.SetResult(ASrcXF.Fill,FSearchXF.Fill);
  end
  else
    FSearchXF.Fill.Assign(ASrcXF.Fill);

  if FAssigneds[cfoFont] then
    FFont.SetResult(ASrcXF.Font,FSearchXF.Font)
  else
    FSearchXF.Font.Assign(ASrcXF.Font);

  if FAssigneds[cfoBorder] then
    FBorder.SetResult(ASrcXF.Border,FSearchXF.Border,ASides)
  else
    FSearchXF.Border.Assign(ASrcXF.Border);

  if FAssigneds[cfoXFProtection] then
    FProtect.SetResult(ASrcXF,FSearchXF)
  else
    FSearchXF.Protection := ASrcXF.Protection;

  Result := FStyles.XFs.Find(FSearchXF);
  if Result = Nil then begin
    Result := FStyles.XFs.AddOrGetFree;

    if FAssigneds[cfoXFAlignment] or FAssigneds[cfoXFProtection] then begin
      Result.AssignProps(FSearchXF);
      Result.Apply := Result.Apply + [eafAlignment];
    end
    else
      Result.AssignProps(ASrcXF);

    if FAssigneds[cfoXFProtection] then begin
      Result.AssignProps(FSearchXF);
      Result.Apply := Result.Apply + [eafProtection];
    end
    else
      Result.Protection := ASrcXF.Protection;

    if FAssigneds[cfoNumFmt] then begin
      Result.NumFmt := FStyles.NumFmts.Find(FSearchXF.NumFmt);
      if Result.NumFmt = Nil then begin
        Result.NumFmt := FStyles.NumFmts.Add;
        Result.NumFmt.Assign(FSearchXF.NumFmt);
      end;
      Result.Apply := Result.Apply + [eafNumberFormat];
    end
    else
      Result.NumFmt := ASrcXF.NumFmt;

    if FAssigneds[cfoFill] then begin
      Result.Fill := FStyles.Fills.Find(FSearchXF.Fill);
      if Result.Fill = Nil then begin
        Result.Fill := FStyles.Fills.Add;
        Result.Fill.Assign(FSearchXF.Fill);
      end;
      Result.Apply := Result.Apply + [eafFill];
    end
    else
      Result.Fill := ASrcXF.Fill;

    if FAssigneds[cfoFont] then begin
      Result.Font := FStyles.Fonts.Find(FSearchXF.Font);
      if Result.Font = Nil then begin
        Result.Font := FStyles.Fonts.Add;
        Result.Font.Assign(FSearchXF.Font);
      end;
      Result.Apply := Result.Apply + [eafFont];
    end
    else
      Result.Font := ASrcXF.Font;

    if FAssigneds[cfoBorder] then begin
      Result.Border := FStyles.Borders.Find(FSearchXF.Border);
      if Result.Border = Nil then begin
        Result.Border := FStyles.Borders.Add;
        Result.Border.Assign(FSearchXF.Border);
      end;
      Result.Apply := Result.Apply + [eafBorder];
    end
    else
      Result.Border := ASrcXF.Border;
  end;

  if FAssigneds[cfoFont] and FSetRowHeight then begin
    if FMaxFont = nil then
      FMaxFont := Result.Font
    else if Result.Font.Size > FMaxFont.Size then
      FMaxFont := Result.Font;
  end;
end;

function TXLSCmdFormat.AddAsDefault(const AName: AxUCString): TXLSDefaultFormat;
var
  XF: TXc12XF;
begin
  XF := MakeXF(FStyles.XFs.DefaultXF,[cbsTop,cbsLeft,cbsRight,cbsBottom,cbsInsideHoriz,cbsInsideVert,cbsDiagLeft,cbsDiagRight]);
  Result := FDefaults.Add(XF,AName);
  Result.XF.Locked := True;
  Result.XF.NumFmt.Locked := True;
  Result.XF.Fill.Locked := True;
  Result.XF.Border.Locked := True;
  Result.XF.Font.Locked := True;
end;

procedure TXLSCmdFormat.SetSetRowHeight(const Value: boolean);
begin
  FSetRowHeight := Value;
end;

{ TXLSCmdFormatFont }

procedure TXLSCmdFormatFont.Assign(AFont: TFont);
begin
  ClearAssigned;
  SetDefault;

{$ifdef BABOON}
  if AFont.Family <> Name                then Name := AFont.Family;
  if AFont.Size <> Size                  then Size := AFont.Size;
//  if AFont.Color <> Color.TColor         then Color.TColor := AFont.Color;
  if (TFontStyle.fsBold in AFont.Style) <> Bold     then Bold := (TFontStyle.fsBold in AFont.Style);
  if (TFontStyle.fsItalic in AFont.Style) <> Italic then Italic := (TFontStyle.fsItalic in AFont.Style);
  if TFontStyle.fsUnderline in AFont.Style          then Underline := xulSingle else Underline := xulNone;
{$else}
  if AFont.Name <> Name                  then Name := AFont.Name;
  if AFont.Size <> Size                  then Size := AFont.Size;
  if AFont.Color <> Color.TColor         then Color.TColor := AFont.Color;
  if (fsBold in AFont.Style) <> Bold     then Bold := (fsBold in AFont.Style);
  if (fsItalic in AFont.Style) <> Italic then Italic := (fsItalic in AFont.Style);
  if fsUnderline in AFont.Style          then Underline := xulSingle else Underline := xulNone;
{$endif}
end;

procedure TXLSCmdFormatFont.Assign(AFont: TXc12Font);
begin
  ClearAssigned;
  SetDefault;

  if AFont.Name <> Name                       then Name := AFont.Name;
  if AFont.Size <> Size                       then Size := AFont.Size;
  if Xc12ColorToRGB(AFont.Color) <> Color.RGB then Color.RGB := Xc12ColorToRGB(AFont.Color);
  if (xfsBold in AFont.Style) <> Bold         then Bold := (xfsBold in AFont.Style);
  if (xfsItalic in AFont.Style) <> Italic     then Italic := (xfsItalic in AFont.Style);
  if AFont.Underline <> xulNone               then Underline := xulSingle else Underline := xulNone;

  FFont.Assign(AFont);
end;

procedure TXLSCmdFormatFont.Clear;
begin
  FFont.Assign(FFonts.DefaultFont);
  FColor.Clear;

  ClearAssigned;
end;

procedure TXLSCmdFormatFont.ClearAssigned;
var
  i: TCmdFmtFontProp;
begin
  for i := Low(FAssigneds) to High(FAssigneds) do
    FAssigneds[i] := False;

  FOwner.FAssigneds[cfoFont] := False;
end;

procedure TXLSCmdFormatFont.ColorChanged(ASender: TObject);
begin
  SetAssigned(cffpColor);
end;

procedure TXLSCmdFormatFont.CopyToTFont(ATFont: TFont);
begin
  FFont.CopyToTFont(ATfont);
end;

constructor TXLSCmdFormatFont.Create(AOwner: TXLSCmdFormat; AFonts: TXc12Fonts);
begin
  FOwner := AOwner;
  FFonts := AFonts;

  FFont := TXc12Font.Create(Nil);
  FColor := TXLSCmdFormatColor.Create(FFont.PColor);
  FColor.Changed := ColorChanged;
end;

destructor TXLSCmdFormatFont.Destroy;
begin
  FColor.Free;
  FFont.Free;
  inherited;
end;

function TXLSCmdFormatFont.GetBold: boolean;
begin
  Result := xfsBold in FFont.Style;
end;

function TXLSCmdFormatFont.GetCharset: integer;
begin
  Result := FFont.Charset;
end;

function TXLSCmdFormatFont.GetItalic: boolean;
begin
  Result := xfsItalic in FFont.Style;
end;

function TXLSCmdFormatFont.GetName: AxUCString;
begin
  Result := FFont.Name;
end;

function TXLSCmdFormatFont.GetSize: double;
begin
  Result := FFont.Size;
end;

function TXLSCmdFormatFont.GetStyle: TXc12FontStyles;
begin
  Result := FFont.Style;
end;

function TXLSCmdFormatFont.GetSubSuperscript: TXc12SubSuperscript;
begin
  Result := FFont.SubSuperscript;
end;

function TXLSCmdFormatFont.GetUnderline: TXc12Underline;
begin
  Result := FFont.Underline;
end;

procedure TXLSCmdFormatFont.SetAssigned(const AValue: TCmdFmtFontProp);
begin
  FAssigneds[AValue] := True;
  FOwner.FAssigneds[cfoFont] := True;
end;

procedure TXLSCmdFormatFont.SetBold(const Value: boolean);
begin
  if Value then
    FFont.Style := FFont.Style + [xfsBold]
  else
    FFont.Style := FFont.Style - [xfsBold];
  SetAssigned(cffpStyle);
end;

procedure TXLSCmdFormatFont.SetCharset(const Value: integer);
begin
  FFont.Charset := Value;
  SetAssigned(cffpCharset);
end;

procedure TXLSCmdFormatFont.SetDefault;
begin
  FFont.Assign(FFonts.DefaultFont);
end;

procedure TXLSCmdFormatFont.SetItalic(const Value: boolean);
begin
  if Value then
    FFont.Style := FFont.Style + [xfsItalic]
  else
    FFont.Style := FFont.Style - [xfsItalic];
  SetAssigned(cffpStyle);
end;

procedure TXLSCmdFormatFont.SetName(const Value: AxUCString);
begin
  if Trim(Value) <> '' then begin
    FFont.Name := Trim(Value);
    SetAssigned(cffpName);
  end;
end;

procedure TXLSCmdFormatFont.SetResult(ASrc,ADest: TXc12Font);
var
  i: TCmdFmtFontProp;
begin
  for i := Low(TCmdFmtFontProp) to High(TCmdFmtFontProp) do begin
    if FAssigneds[i] then begin
      case i of
        cffpName          : ADest.Name := FFont.Name;
        cffpSize          : ADest.Size := FFont.Size;
        cffpColor         : ADest.Color := FFont.Color;
        cffpStyle         : ADest.Style := FFont.Style;
        cffpUnderline     : ADest.Underline := FFont.Underline;
        cffpSubSuperscript: ADest.SubSuperscript := FFont.SubSuperscript;
        cffpCharset       : ADest.Charset := FFont.Charset;
      end;
    end
    else begin
      case i of
        cffpName          : ADest.Name := ASrc.Name;
        cffpSize          : ADest.Size := ASrc.Size;
        cffpColor         : ADest.Color := ASrc.Color;
        cffpStyle         : ADest.Style := ASrc.Style;
        cffpUnderline     : ADest.Underline := ASrc.Underline;
        cffpSubSuperscript: ADest.SubSuperscript := ASrc.SubSuperscript;
        cffpCharset       : ADest.Charset := ASrc.Charset;
      end;
    end;
  end;
end;

procedure TXLSCmdFormatFont.SetSize(const Value: double);
begin
  FFont.Size := Value;
  SetAssigned(cffpSize);
end;

procedure TXLSCmdFormatFont.SetStyle(const Value: TXc12FontStyles);
begin
  FFont.Style := Value;
  SetAssigned(cffpStyle);
end;

procedure TXLSCmdFormatFont.SetSubSuperscript(const Value: TXc12SubSuperscript);
begin
  FFont.SubSuperscript := Value;
  SetAssigned(cffpSubSuperscript);
end;

procedure TXLSCmdFormatFont.SetUnderline(const Value: TXc12Underline);
begin
  FFont.Underline := Value;
  SetAssigned(cffpUnderline);
end;

{ TXLSCmdFormatColor }

procedure TXLSCmdFormatColor.Clear;
begin
  FColor.ColorType := exctAuto;
  FColor.ARGB := $000000;
end;

constructor TXLSCmdFormatColor.Create(AColor: PXc12Color);
begin
  FColor := AColor
end;

function TXLSCmdFormatColor.GetColor: TXc12Color;
begin
  Result := FColor^;
end;

function TXLSCmdFormatColor.GetIndexColor: TXc12IndexColor;
begin
  if FColor.ColorType = exctIndexed then
    Result := FColor.Indexed
  else
    Result := xcAutomatic;
end;

function TXLSCmdFormatColor.GetRGB: longword;
begin
  Result := Xc12ColorToRGB(FColor);
end;

function TXLSCmdFormatColor.GetTColor: TColor;
begin
  Result := RevRGB(GetRGB);
end;

procedure TXLSCmdFormatColor.SetColor(const Value: TXc12Color);
begin
  if not Xc12ColorEqual(Value,FColor^) then begin
    FColor^ := Value;
    if Assigned(FChanged) then
      FChanged(Self);
  end;
end;

procedure TXLSCmdFormatColor.SetIndexColor(const Value: TXc12IndexColor);
begin
  if (FColor.ColorType = exctIndexed) and (FColor.Indexed = Value) then
    Exit;

  FColor.ColorType := exctIndexed;
  FColor.ARGB := Xc12IndexColorPalette[Integer(Value)];
  FColor.Indexed := Value;

  if Assigned(FChanged) then
    FChanged(Self);
end;

procedure TXLSCmdFormatColor.SetRGB(const Value: TXc12RGBColor);
begin
  if Xc12ColorEqual(FColor,Value) then
    Exit;

  RGBColorToXc12(Value,FColor);

  if Assigned(FChanged) then
    FChanged(Self);
end;

procedure TXLSCmdFormatColor.SetTColor(const Value: TColor);
begin
  SetRGB(RevRGB(Value));
end;

procedure TXLSCmdFormatColor.SetTheme(AScheme: TXc12ClrSchemeColor; const Value: double);
begin
  ThemeColorToXc12(FColor,AScheme,Fork(Value,-1,1));

  if Assigned(FChanged) then
    FChanged(Self);
end;

procedure TXLSCmdFormatColor.ExcelSwatch(AColor, ATint: integer);
var
  Tint: double;
begin
  AColor := Fork(AColor,0,9);
  ATint := Fork(ATint,0,4);

  case ATint of
    0: Tint := 0.0;
    1: Tint := 0.8;
    2: Tint := 0.6;
    3: Tint := 0.4;
    4: Tint := -0.25;
    5: Tint := -0.5;
    else Tint := 0;
  end;

  ThemeColorToXc12(FColor,TXc12ClrSchemeColor(AColor),Tint);

  if Assigned(FChanged) then
    FChanged(Self);
end;

{ TXLSCmdFormatFill }

procedure TXLSCmdFormatFill.BgColorChanged(ASender: TObject);
begin
  FAssigneds[cfipBgColor] := True;
  if PatternStyle = efpNone then
    SetPatternStyle(efpSolid)
  else
    FOwner.FAssigneds[cfoFill] := True;
end;

procedure TXLSCmdFormatFill.Clear;
var
  i: TCmdFmtFillProp;
begin
  for i := Low(TCmdFmtFillProp) to High(TCmdFmtFillProp) do
    FAssigneds[i] := False;

  FFill.Assign(FFills.DefaultFill);

  FOwner.FAssigneds[cfoFill] := False;

  FNoColor := False;
end;

constructor TXLSCmdFormatFill.Create(AOwner: TXLSCmdFormat; AFills: TXc12Fills);
begin
  FOwner := AOwner;
  FFills := AFills;
  FFill := TXc12Fill.Create(Nil);

  FBgColor := TXLSCmdFormatColor.Create(FFill.PFgColor);
  FBgColor.Changed := BgColorChanged;
  FPatColor := TXLSCmdFormatColor.Create(FFill.PBgColor);
  FPatColor.Changed := PatColorChanged;
end;

destructor TXLSCmdFormatFill.Destroy;
begin
  FBgColor.Free;
  FPatColor.Free;

  FFill.Free;
  inherited;
end;

function TXLSCmdFormatFill.GetPatternStyle: TXc12FillPattern;
begin
  Result := FFill.PatternType;
end;

procedure TXLSCmdFormatFill.PatColorChanged(ASender: TObject);
begin
  FAssigneds[cfipPatColor] := True;
  FAssigneds[cfipBgColor] := True;
  FAssigneds[cfipPatColor] := True;
  FAssigneds[cfipPattern] := True;
  FOwner.FAssigneds[cfoFill] := True;
end;

procedure TXLSCmdFormatFill.SetNoColor(const Value: boolean);
begin
  FNoColor := Value;
  FOwner.FAssigneds[cfoFill] := True;
end;

procedure TXLSCmdFormatFill.SetPatternStyle(const Value: TXc12FillPattern);
begin
  FFill.PatternType := Value;
  FAssigneds[cfipPattern] := True;
  FOwner.FAssigneds[cfoFill] := True;
end;

procedure TXLSCmdFormatFill.SetResult(ASrc,ADest: TXc12Fill);
begin
  ADest.PatternType := efpSolid;
  if FAssigneds[cfipBgColor] then
    ADest.FgColor := FFill.FgColor
  else
    ADest.FgColor := ASrc.FgColor;
  if FAssigneds[cfipPatColor] then
    ADest.BgColor := FFill.BgColor
  else
    ADest.BgColor := ASrc.BgColor;
  if FAssigneds[cfipPattern] then
    ADest.PatternType := FFill.PatternType
  else
    ADest.PatternType := ASrc.PatternType;
end;

{ TXLSCmdFormatBorder }

procedure TXLSCmdFormatBorder.Clear;
var
  i: TCmdBorderSide;
begin
  for i := Low(TCmdBorderSide) to High(TCmdBorderSide) do
    FSides[i] := ctsUnassigned;

  FColor := FOwner.FStyles.Borders.DefaultBorder.Left.Color;
  FStyle := FOwner.FStyles.Borders.DefaultBorder.Left.Style;

  FOwner.FAssigneds[cfoBorder] := False;

  FOptions := [];
end;

constructor TXLSCmdFormatBorder.Create(AOwner: TXLSCmdFormat);
var
  i: TCmdBorderSide;
begin
  FOwner := AOwner;

  FCmdColor := TXLSCmdFormatColor.Create(@FColor);

  for i := Low(TCmdBorderSide) to High(TCmdBorderSide) do
    FBorders[i] := TXc12BorderPr.Create;
end;

destructor TXLSCmdFormatBorder.Destroy;
var
  i: TCmdBorderSide;
begin
  for i := Low(TCmdBorderSide) to High(TCmdBorderSide) do
    FBorders[i].Free;

  FCmdColor.Free;

  inherited;
end;

function TXLSCmdFormatBorder.GetSide(Index: TCmdBorderSide): boolean;
begin
  Result := FSides[Index] = ctsTrue;
end;

procedure TXLSCmdFormatBorder.Preset(const ASides: TCmdBorderSidePreset);
var
  i: TCmdBorderSide;
begin
  case ASides of
    cbspNone   : begin
      for i := Low(TCmdBorderSide) to High(TCmdBorderSide) do
        SetSide(i,False);
    end;
    cbspOutline: begin
      SetSide(cbsLeft,True);
      SetSide(cbsTop,True);
      SetSide(cbsRight,True);
      SetSide(cbsBottom,True);
    end;
    cbspInside : begin
      SetSide(cbsInsideVert,True);
      SetSide(cbsInsideHoriz,True);
    end;
    cbspOutlineAndInside: begin
      SetSide(cbsLeft,True);
      SetSide(cbsTop,True);
      SetSide(cbsRight,True);
      SetSide(cbsBottom,True);

      SetSide(cbsInsideVert,True);
      SetSide(cbsInsideHoriz,True);
    end;
  end;
end;

procedure TXLSCmdFormatBorder.SetResult(ASrc,ADest: TXc12Border; const ASides: TCmdBorderSides);
begin
  ADest.Left.Style := ASrc.Left.Style;
  ADest.Top.Style := ASrc.Top.Style;
  ADest.Right.Style := ASrc.Right.Style;
  ADest.Bottom.Style := ASrc.Bottom.Style;

  if cbsLeft in ASides then
    SetResultPr(ASrc.Left,ADest.Left,cbsLeft);

  if cbsTop in ASides then
    SetResultPr(ASrc.Top,ADest.Top,cbsTop);

  if cbsRight in ASides then
    SetResultPr(ASrc.Right,ADest.Right,cbsRight);

  if cbsBottom in ASides then
    SetResultPr(ASrc.Bottom,ADest.Bottom,cbsBottom);

  if (cbsInsideHoriz in ASides) and not (cbsBottom in ASides) then
    SetResultPr(ASrc.Bottom,ADest.Bottom,cbsInsideHoriz);

  if (cbsInsideVert in ASides) and not (cbsRight in ASides) then
    SetResultPr(ASrc.Right,ADest.Right,cbsInsideVert);

  if cbsDiagLeft in ASides then begin
    SetResultPr(ASrc.Diagonal,ADest.Diagonal,cbsDiagLeft);

    ADest.Options := ADest.Options + [ecboDiagonalDown];
  end;

  if cbsDiagRight in ASides then begin
    SetResultPr(ASrc.Diagonal,ADest.Diagonal,cbsDiagRight);

    ADest.Options := ADest.Options + [ecboDiagonalUp];
  end;
end;

procedure TXLSCmdFormatBorder.SetResultPr(ASrc,ADest: TXc12BorderPr; const ASide: TCmdBorderSide);
begin
  case FSides[ASide] of
    ctsTrue      : begin
      ADest.Assign(FBorders[ASide]);
      if (FBorders[ASide].Color.ColorType  <> exctAuto) and (ASrc.Style <> cbsNone) then
        ADest.Style := ASrc.Style;
    end;
    ctsFalse     : ADest.Clear;
    ctsUnassigned: ADest.Assign(ASrc);
  end;

  case ASide of
    cbsDiagLeft : FOptions := FOptions + [ecboDiagonalDown];
    cbsDiagRight: FOptions := FOptions + [ecboDiagonalDown];
  end;
end;

procedure TXLSCmdFormatBorder.SetSide(Index: TCmdBorderSide; const Value: boolean);
begin
  if Value then
    FSides[Index] := ctsTrue
  else
    FSides[Index] := ctsFalse;
  FBorders[Index].Color := FColor;
  FBorders[Index].Style := FStyle;
  FOwner.FAssigneds[cfoBorder] := True;
end;

{ TXLSCmdFormatNumberFormat }

procedure TXLSCmdFormatNumber.Clear;
begin
  FNumFmt.Assign(FNumFmts.DefaultNumFmt);
  FAssigned := False;

  FDecimals := 0;
  FThousands := False;
  FNegColor := clDefault;
  FNegColorSign := False;
  FCurrencyLCID := -1;
  FPercentage := False;
  FFractions := xnffNone;
  FScientific := False;

  FOwner.FAssigneds[cfoNumFmt] := False;
end;

constructor TXLSCmdFormatNumber.Create(AOwner: TXLSCmdFormat; ANumFmts: TXc12NumberFormats);
begin
  FOwner := AOwner;

  FNumFmts := ANumFmts;
  FNumFmt := TXc12NumberFormat.Create(Nil);

  FDecimals := -1;
  FNegColor := clDefault;
  FCurrencyLCID := -1;
end;

destructor TXLSCmdFormatNumber.Destroy;
begin
  FNumFmt.Free;
  inherited;
end;

function TXLSCmdFormatNumber.GetDecimalsStr: AxUCString;
var
  i: integer;
begin
  Result := '';
  if FDecimals <= 0 then
    Result := Result + '0'
  else if FDecimals > 0 then begin
    Result := Result + '0.';
    for i := 1 to FDecimals do
      Result := Result + '0';
  end;
end;

function TXLSCmdFormatNumber.GetFormat: AxUCString;
begin
  Result := FNumFmt.Value;
end;

procedure TXLSCmdFormatNumber.MakeFormatString;
var
  CurrFmt: integer;
  S: AxUCString;
  sCurrency: AxUCString;
  sColor: AxUCString;
  Fmt: TFormatSettings;
begin
  if FPercentage then
    S := GetDecimalsStr + '%'
  else if FScientific then
    S := GetDecimalsStr + 'E+00'
  else if FFractions > xnffNone then begin
    case FFractions of
      xnffOneDigit     : S := '#" "?/?';
      xnffTwoDigits    : S := '#" "??/??';
      xnffThreeDigits  : S := '#" "???/???';
      xnffAsHalves     : S := '#" "?/2';
      xnffAsQuarters   : S := '#" "?/4';
      xnffAsEights     : S := '#" "?/8';
      xnffAsSixtheenths: S := '#" "?/16';
      xnffAsTenths     : S := '#" "?/10';
      xnffAsHundredths : S := '#" "?/100';
    end;
  end
  else begin
    sCurrency := '';
    CurrFmt := 0;
    if FCurrencyLCID >= 0 then begin
      if FCurrencyLCID = 0 then begin
        sCurrency := FormatSettings.CurrencyString;
        CurrFmt := FormatSettings.CurrencyFormat;
      end
      else begin
{$ifdef DELPHI_XE_OR_LATER}
{$WARN SYMBOL_PLATFORM OFF}
        Fmt := TFormatSettings.Create({$ifndef BABOON}FCurrencyLCID{$endif});
{$WARN SYMBOL_PLATFORM ON}
{$else}
        Fmt := TFormatSettings.Create;
{$endif}
        sCurrency := Fmt.CurrencyString;
        CurrFmt := Fmt.CurrencyFormat;
      end;
    end;

    S := '';
    if FThousands then
      S := S + '#,##';

    S := S + GetDecimalsStr;

    case FNegColor of
      clBlack  : sColor := ';[Black]';
      clFuchsia: sColor := ';[Fuchsia]';
      clPurple : sColor := ';[Purple]';
      clWhite  : sColor := ';[White]';
      clBlue   : sColor := ';[Blue]';
      clGreen  : sColor := ';[Green]';
      clRed    : sColor := ';[Red]';
      clYellow : sColor := ';[Yellow]';
    end;

    // "#,##0.000_ ;[Red]\-#,##0.000\ "
    if Longword(FNegColor) <> Longword(clDefault) then begin
      if FNegColorSign then
        S := S + '_ ' + sColor + '\-' + S + '\ '
      else
        S := S + sColor + S;
    end;

    if sCurrency <> '' then begin
      case CurrFmt of
        0: S := '\' + sCurrency + S;
        1: S := S + '\' + sCurrency;
        2: S := '\' + sCurrency + ' ' + S;
        3: S := S + '\' + ' ' + sCurrency;
      end;
    end;
  end;

  SetFormat(S);
end;

procedure TXLSCmdFormatNumber.SetCurrencySymbolLCID(const Value: integer);
begin
  FCurrencyLCID := Value;
  MakeFormatString;
end;

procedure TXLSCmdFormatNumber.SetDecimals(const Value: integer);
begin
  FDecimals := Value;
  MakeFormatString;
end;

procedure TXLSCmdFormatNumber.SetFormat(const Value: AxUCString);
begin
  FNumFmt.Value := Value;
  FAssigned := True;
  FOwner.FAssigneds[cfoNumFmt] := True;
end;


procedure TXLSCmdFormatNumber.SetFractions(const Value: TXLSNumFmtFractions);
begin
  FFractions := Value;
  MakeFormatString;
end;

procedure TXLSCmdFormatNumber.SetNegColor(const Value: TColor);
begin
  if (Longword(Value) = Longword(clBlack)) or (Longword(Value) = Longword(clFuchsia)) or (Longword(Value) = Longword(clPurple)) or (Longword(Value) = Longword(clWhite)) or
     (Longword(Value) = Longword(clBlue)) or (Longword(Value) = Longword(clGreen)) or (Longword(Value) = Longword(clRed)) or (Longword(Value) = Longword(clYellow)) or (Longword(Value) = Longword(clDefault)) then begin
    FNegColor := Value;
    MakeFormatString;
  end
  else
    raise XLSRWException.Create('Invalid negative color');
end;

procedure TXLSCmdFormatNumber.SetNegColorSign(const Value: boolean);
begin
  FNegColorSign := Value;
  MakeFormatString;
end;

procedure TXLSCmdFormatNumber.SetPercentage(const Value: boolean);
begin
  FPercentage := Value;
  MakeFormatString;
end;

procedure TXLSCmdFormatNumber.SetResult(ADest: TXc12NumberFormat);
begin
  if FAssigned then
    ADest.Value := FNumFmt.Value;
end;

procedure TXLSCmdFormatNumber.SetScientific(const Value: boolean);
begin
  FScientific := Value;
  MakeFormatString;
end;

procedure TXLSCmdFormatNumber.SetThousands(const Value: boolean);
begin
  FThousands := Value;
  MakeFormatString;
end;

{ TXLSCmdFormatAlignment }

procedure TXLSCmdFormatAlignment.Clear;
var
  i: TCmdFmtXFProp;
begin
  for i := Low(TCmdFmtXFProp) to High(TCmdFmtXFProp) do
    FAssigneds[i] := False;

  FXF.Alignment.Assign(FXFs.DefaultXF.Alignment);

  FOwner.FAssigneds[cfoXFAlignment] := False;
end;

constructor TXLSCmdFormatAlignment.Create(AOwner: TXLSCmdFormat; AXFs: TXc12XFs);
begin
  FOwner := AOwner;
  FXFs := AXFs;

  FXF := TXc12XF.Create(Nil);
end;

destructor TXLSCmdFormatAlignment.Destroy;
begin
  FXF.Free;
  inherited;
end;

function TXLSCmdFormatAlignment.GetHorizontal: TXc12HorizAlignment;
begin
  Result := FXF.Alignment.HorizAlignment;
end;

function TXLSCmdFormatAlignment.GetIndent: integer;
begin
  Result := FXF.Alignment.Indent;
end;

function TXLSCmdFormatAlignment.GetJustifyDistributed: boolean;
begin
  Result := foJustifyLastLine in FXF.Alignment.Options;
end;

function TXLSCmdFormatAlignment.GetMergeCells: boolean;
begin
  Result := FMergeCells;
end;

function TXLSCmdFormatAlignment.GetRotation: integer;
begin
  Result := FXF.Alignment.Rotation;
end;

function TXLSCmdFormatAlignment.GetShrinkToFit: boolean;
begin
  Result := foShrinkToFit in FXF.Alignment.Options;
end;

function TXLSCmdFormatAlignment.GetTextDirection: TXc12ReadOrder;
begin
  Result := FXF.Alignment.ReadingOrder;
end;

function TXLSCmdFormatAlignment.GetVertical: TXc12VertAlignment;
begin
  Result := FXF.Alignment.VertAlignment;
end;

function TXLSCmdFormatAlignment.GetWrapText: boolean;
begin
  Result := foWrapText in FXF.Alignment.Options;
end;

procedure TXLSCmdFormatAlignment.SetHorizontal(const Value: TXc12HorizAlignment);
begin
  FXF.Alignment.HorizAlignment := Value;
  FAssigneds[cfxfAlignHoriz] := True;
  FOwner.FAssigneds[cfoXFAlignment] := True;
end;

procedure TXLSCmdFormatAlignment.SetIndent(const Value: integer);
begin
  if Value >= 0 then begin
    FXF.Alignment.Indent := Value;
    FAssigneds[cfxfIndent] := True;
    FOwner.FAssigneds[cfoXFAlignment] := True;
  end;
end;

procedure TXLSCmdFormatAlignment.SetJustifyDistributed(const Value: boolean);
begin
  if Value then
    FXF.Alignment.Options := FXF.Alignment.Options + [foJustifyLastLine]
  else
    FXF.Alignment.Options := FXF.Alignment.Options - [foJustifyLastLine];
  FAssigneds[cfxfJustifyDist] := True;
  FOwner.FAssigneds[cfoXFAlignment] := True;
end;

procedure TXLSCmdFormatAlignment.SetMergeCells(const Value: boolean);
begin
  FMergeCells := Value;
  FAssigneds[cfxfMergeCells] := True;
  FOwner.FAssigneds[cfoXFAlignment] := True;
end;

procedure TXLSCmdFormatAlignment.SetResult(ASrc,ADest: TXc12CellAlignment);
var
  i: TCmdFmtXFProp;
begin
  ADest.Options := [];
  for i := Low(TCmdFmtXFProp) to High(TCmdFmtXFProp) do begin
    if FAssigneds[i] then begin
      case i of
        cfxfAlignHoriz   : ADest.HorizAlignment := FXF.Alignment.HorizAlignment;
        cfxfAlignVert    : ADest.VertAlignment := FXF.Alignment.VertAlignment;
        cfxfIndent       : ADest.Indent := FXF.Alignment.Indent;
        cfxfJustifyDist  : begin
          if foJustifyLastLine in FXF.Alignment.Options then
            ADest.Options := ADest.Options + [foJustifyLastLine]
          else
            ADest.Options := ADest.Options - [foJustifyLastLine];
        end;
        cfxfWrapText     : begin
          if foWrapText in FXF.Alignment.Options then
            ADest.Options := ADest.Options + [foWrapText]
          else
            ADest.Options := ADest.Options - [foWrapText];
        end;
        cfxfShrinkToFit  : begin
          if foShrinkToFit in FXF.Alignment.Options then
            ADest.Options := ADest.Options + [foShrinkToFit]
          else
            ADest.Options := ADest.Options - [foShrinkToFit];
        end;
        cfxfMergeCells   : ;
        cfxfRotation     : ADest.Rotation := FXF.Alignment.Rotation;
        cfxfTextDirection: ADest.ReadingOrder := FXF.Alignment.ReadingOrder;
      end;
    end
    else begin
      case i of
        cfxfAlignHoriz   : ADest.HorizAlignment := ASrc.HorizAlignment;
        cfxfAlignVert    : ADest.VertAlignment := ASrc.VertAlignment;
        cfxfIndent       : ADest.Indent := ASrc.Indent;
        cfxfJustifyDist  : if foJustifyLastLine in ASrc.Options then ADest.Options := ADest.Options + [foJustifyLastLine];
        cfxfWrapText     : if foWrapText in ASrc.Options then ADest.Options := ADest.Options + [foWrapText];
        cfxfShrinkToFit  : if foShrinkToFit in ASrc.Options then ADest.Options := ADest.Options + [foShrinkToFit];
        cfxfMergeCells   : ;
        cfxfRotation     : ADest.Rotation := ASrc.Rotation;
        cfxfTextDirection: ADest.ReadingOrder := ASrc.ReadingOrder;
      end;
    end;
  end;
end;

procedure TXLSCmdFormatAlignment.SetRotation(const Value: integer);
begin
  FXF.Alignment.Rotation := Fork(Value,-180,180);
  FAssigneds[cfxfRotation] := True;
  FOwner.FAssigneds[cfoXFAlignment] := True;
end;

procedure TXLSCmdFormatAlignment.SetShrinkToFit(const Value: boolean);
begin
  if Value then
    FXF.Alignment.Options := FXF.Alignment.Options + [foShrinkToFit]
  else
    FXF.Alignment.Options := FXF.Alignment.Options - [foShrinkToFit];
  FAssigneds[cfxfShrinkToFit] := True;
  FOwner.FAssigneds[cfoXFAlignment] := True;
end;

procedure TXLSCmdFormatAlignment.SetTextDirection(const Value: TXc12ReadOrder);
begin
  FXF.Alignment.ReadingOrder := Value;
  FAssigneds[cfxfTextDirection] := True;
  FOwner.FAssigneds[cfoXFAlignment] := True;
end;

procedure TXLSCmdFormatAlignment.SetVertical(const Value: TXc12VertAlignment);
begin
  FXF.Alignment.VertAlignment := Value;
  FAssigneds[cfxfAlignVert] := True;
  FOwner.FAssigneds[cfoXFAlignment] := True;
end;

procedure TXLSCmdFormatAlignment.SetWrapText(const Value: boolean);
begin
  if Value then
    FXF.Alignment.Options := FXF.Alignment.Options + [foWrapText]
  else
    FXF.Alignment.Options := FXF.Alignment.Options - [foWrapText];
  FAssigneds[cfxfWrapText] := True;
  FOwner.FAssigneds[cfoXFAlignment] := True;
end;

{ TXLSCmdFormatProtection }

procedure TXLSCmdFormatProtection.Clear;
begin
  FAssignLocked := ctsUnassigned;
  FAssignHidden := ctsUnassigned;

  FOwner.FAssigneds[cfoXFProtection] := False;
end;

constructor TXLSCmdFormatProtection.Create(AOwner: TXLSCmdFormat; AXFs: TXc12XFs);
begin
  FOwner := AOwner;
  FXFs := AXFs;
  FXF := TXc12XF.Create(Nil);
end;

destructor TXLSCmdFormatProtection.Destroy;
begin
  FXF.Free;
  inherited;
end;

function TXLSCmdFormatProtection.GetHidden: boolean;
begin
  Result := cpHidden in FXF.Protection;
end;

function TXLSCmdFormatProtection.GetLocked: boolean;
begin
  Result := cpLocked in FXF.Protection;
end;

procedure TXLSCmdFormatProtection.SetHidden(const Value: boolean);
begin
  if Value then
    FAssignHidden := ctsTrue
  else
    FAssignHidden := ctsFalse;
  FOwner.FAssigneds[cfoXFProtection] := True;
end;

procedure TXLSCmdFormatProtection.SetLocked(const Value: boolean);
begin
  if Value then
    FAssignLocked := ctsTrue
  else
    FAssignLocked := ctsFalse;
  FOwner.FAssigneds[cfoXFProtection] := True;
end;

procedure TXLSCmdFormatProtection.SetResult(ASrc,ADest: TXc12XF);
begin
  ADest.Protection := [];
  if FAssignHidden = ctsTrue then
    ADest.Protection := ADest.Protection + [cpHidden]
  else if FAssignHidden = ctsFalse then
    ADest.Protection := ADest.Protection - [cpHidden]
  else if cpHidden in ASrc.Protection then
    ADest.Protection := ADest.Protection + [cpHidden];

  if FAssignLocked = ctsTrue then
    ADest.Protection := ADest.Protection + [cpLocked]
  else if FAssignLocked = ctsFalse then
    ADest.Protection := ADest.Protection - [cpLocked]
  else if cpHidden in ASrc.Protection then
    ADest.Protection := ADest.Protection + [cpLocked];
end;

{ TXLSDefaultFormats }

function TXLSDefaultFormats.Add(AXF: TXc12XF; const AName: AxUCString): TXLSDefaultFormat;
begin
  Result := TXLSDefaultFormat.Create;
  Result.FXF := AXF;
  Result.Name := AName;
  FItems.AddObject(AName,Result);
end;

procedure TXLSDefaultFormats.Clear;
var
  i: integer;
begin
  for i := 0 to FItems.Count - 1 do
    Items[i].XF.Locked := False;
  FItems.Clear;
end;

function TXLSDefaultFormats.Count: integer;
begin
  Result := FItems.Count;
end;

constructor TXLSDefaultFormats.Create;
begin
{$ifdef DELPHI_5}
  FItems := TStringList.Create;
{$else}
  FItems := THashedStringList.Create;
{$endif}
end;

destructor TXLSDefaultFormats.Destroy;
var
  i: integer;
begin
  for i := 0 to FItems.Count - 1 do
    TXLSDefaultFormat(FItems.Objects[i]).Free;

  FItems.Free;
  inherited;
end;

function TXLSDefaultFormats.Find(const AName: AxUCString): TXLSDefaultFormat;
var
  i: integer;
begin
  i := FItems.IndexOf(AName);
  if i >= 0 then
    Result := Items[i]
  else
    Result := Nil;
end;

function TXLSDefaultFormats.GetItems(Index: integer): TXLSDefaultFormat;
begin
  Result := TXLSDefaultFormat(FItems.Objects[Index]);
end;

{ TXLSDefaultFormat }

procedure TXLSDefaultFormat.UseByDirectWrite(ACell: TXLSEventCell);
begin
  ACell.XF := FXF;
end;

end.
