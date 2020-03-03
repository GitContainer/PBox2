unit XLSSheetData5;

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

uses Classes, SysUtils, Contnrs, Math,
{$ifdef MSWINDOWS}
     Windows,
  {$ifndef BABOON}
     vcl.Graphics, vcl.ComCtrls,
  {$endif}
{$endif}
{$ifdef DELPHI_6_OR_LATER}
     Variants,
{$endif}
{$ifdef XLS_BIFF}
     BIFF5, BIFF_EncodeFormulaII5, BIFF_SheetData5, BIFF_DrawingObj5, BIFF_VBA5,
{$endif}
     xpgParseDrawing, xpgParserPivot,
     Xc12Utils5, Xc12Common5, Xc12DataSST5, Xc12DataWorkbook5, Xc12DataWorksheet5,
     Xc12Manager5, Xc12FileData5,
     XLSUtils5, XLSMMU5, XLSCellMMU5, XLSRow5, XLSColumn5, XLSCellAreas5,
     Xc12DataStyleSheet5,Xc12DataAutofilter5,XLSValidate5, XLSMergedCells5, XLSHyperlinks5, XLSCondFormat5,
     XLSAutofilter5, XLSDrawing5, XLSDecodeFmla5, XLSRTFReadWrite5,
     XLSFormula5, XLSMask5, XLSFormulaTypes5, XLSNames5, XLSClassFactory5, XLSComment5,
     XLSFormattedObj5, XLSRange5, XLSCmdFormat5, XLSRelCells5, XLSSharedItems5, XLSPivotTables5,
     XLSAXWEditor, XLSRTFUtils, XLSReadRTF, XLSWriteRTF, XLSRichPainter5,
{$ifdef BABOON}
{$else}
     XLSClipboard5,
{$endif}
     XLSTools5;

type TXLSDeleteOption = (xdoCellValues,xdoFormats,xdoComments,xdoDataValidations,xdoConditionalFormats,xdoAutofilters,xdoDrawing);
     TXLSDeleteOptions = set of TXLSDeleteOption;

const DefaultDeleteOptions = [xdoCellValues,xdoFormats,xdoComments,xdoDataValidations,xdoConditionalFormats,xdoAutofilters,xdoDrawing];

//* Pane type when a worksheet has splitted panes.
type TPaneType = (ptNone,    //* No panes.
                  ptFrozen,  //* Frozen panes,
                  ptSplit,   //* Splitted panes.
                  ptFrozenSplit
                  );

type TActivePane = (apBottomRight,apTopRight,apBottomLeft,apTopLeft);

type TWorkspaceOption = (woShowAutoBreaks, //* Automatic page breaks are visible.
                         woApplyStyles,    //* Automatic styles are applied to an outline
                         woRowSumsBelow,   //* Summary rows appear below detail in an outline
                         woColSumsRight,   //* Summary columns appear to the right of detail in an outline
                         woFitToPage,      //* Page fit option is on.
                         woOutlineSymbols  //* Outline symbols are displayed.
                         );
     TWorkspaceOptions = set of TWorkspaceOption;

type TWorkbookOption = (woHidden,   //* The window is hidden.
                        woIconized, //* The window is displayed as an icon.
                        woHScroll,  //* The horizontal scroll bar is displayed.
                        woVScroll,  //* The vertical scroll bar is displayed.
                        woTabs      //* The workbook tabs are displayed.
                        );
     TWorkbookOptions = set of TWorkbookOption;

//* Options when copying and moving cells.
type TCopyCellsOption = (ccoAdjustCells,      //* Adjust relative cells according to the new location.
                         ccoLockStartRow,     //* Don't adjust the first row.
                         ccoForceAdjust,      //* Adjust absolute cell references as well.
                         ccoCopyValues,       //* Copy cell values.
                         ccoCopyClearFormulas,//* Remove formulas. Formula results will be converted to cell values.
                         ccoCopyShapes,       //* Copy shape objects.
                         ccoCopyNotes,        //* Copy cell notes.
                         ccoCopyCondFmt,      //* Copy conditional formats.
                         ccoCopyValidations,  //* Copy cell validations.
                         ccoCopyMerged        //* Copy merged cells.
                         );
     TCopyCellsOptions = set of TCopyCellsOption;

//* What parts of the data on a worksheet that shall be copied/moved.
const CopyAllCells = [ccoCopyValues,       //* Cell values
                      ccoCopyShapes,       //* Drawing shapes.
                      ccoCopyNotes,        //* Cell notes.
                      ccoCopyCondFmt,      //* Conditional formats.
                      ccoCopyValidations,  //* Cell validations.
                      ccoCopyMerged        //* Merged cells.
                      ];

type TPrintSetupOption = (psoLeftToRight, psoPortrait, psoNoColor, psoDraftQuality,
                          psoNotes, psoRowColHeading, psoGridlines, psoHorizCenter,
                          psoVertCenter);
type TPrintSetupOptions = set of TPrintSetupOption;

type TXLSSheetOption = (soGridlines,      //* Show gridlines.
                        soRowColHeadings, //* Show row and column headings.
                        soShowFormulas,   //* Show formulas instead of formula result.
                        soShowZeros      //* Show zero values.
                        );

type TXLSSheetOptions = set of TXLSSheetOption;

type TSelectedAreaHit = (sahNo,sahInside,sahEdge,sahActiveCell);

type TSelectedEdgeHit = (sehLeft,sehTop,sehRight,sehBottom);
type TSelectedEdgeHits = set of TSelectedEdgeHit;

type TXLSCellSortItem = record
     Cell: TXLSCellItem;
     Row: integer;
     end;

type TXLSSelectedArea = class(TCellArea)
protected
     FCol1,FRow1,FCol2,FRow2: integer;

     function  GetCol1: integer; override;
     function  GetRow1: integer; override;
     function  GetCol2: integer; override;
     function  GetRow2: integer; override;
     procedure SetCol1(const AValue: integer); override;
     procedure SetRow1(const AValue: integer); override;
     procedure SetCol2(const AValue: integer); override;
     procedure SetRow2(const AValue: integer); override;
public
     function  Hit(ACol,ARow: integer): boolean;
     function  AsRect: TXLSCellArea;
     procedure SetArea(C1,R1,C2,R2: integer);
     procedure Intersect(C1,R1,C2,R2: integer);

     //* True if full columns are selected.
     function  IsColumns: boolean;
     end;

type TXLSSelectedAreas = class(TCellAreas)
private
     function GetItems(Index: integer): TXLSSelectedArea;
protected
     FActiveCol: integer;
     FActiveRow: integer;
     FActiveArea: integer;

     procedure SetActiveCol(const Value: integer);
     procedure SetActiveRow(const Value: integer);
public
     constructor Create;

     procedure Init(ACol: integer = 0; ARow: integer = 0); overload;
     procedure Init(C1,R1,C2,R2,ActC,ActR: integer); overload;

     procedure CursorCell(ACol,ARow: integer);

     function  Add: TXLSSelectedArea; overload;
     function  Add(C1,R1,C2,R2: integer): TXLSSelectedArea; overload;

     function  First: TXLSSelectedArea;
     function  Last: TXLSSelectedArea;

     function  CellInAreas(ACol,ARow: integer; var EdgeHit: TSelectedEdgeHits; var AreaHit: TSelectedAreaHit): integer;
     function  CursorVisible: boolean;

     property ActiveCol: integer read FActiveCol write SetActiveCol;
     property ActiveRow: integer read FActiveRow write SetActiveRow;
     property ActiveArea: integer read FActiveArea write FActiveArea;
     property Items[Index: integer]: TXLSSelectedArea read GetItems; default;
     end;

type TXLSHorizPagebreak = class(TObject)
protected
     FXc12PageBreak: TXc12Break;

     function  GetRow: integer;
     function  GetCol1: integer;
     function  GetCol2: integer;
     procedure SetRow(const Value: integer);
     procedure SetCol1(const Value: integer);
     procedure SetCol2(const Value: integer);
public
     constructor Create(APageBreak: TXc12Break);

     property Row: integer read GetRow write SetRow;
     property Col1: integer read GetCol1 write SetCol1;
     property Col2: integer read GetCol2 write SetCol2;
     end;

type TXLSHorizPagebreaks = class(TObjectList)
private
     function GetItems(Index: integer): TXLSHorizPagebreak;
protected
     FXc12PageBreaks: TXc12PageBreaks;
public
     constructor Create(AXc12PageBreaks: TXc12PageBreaks);

     function Add: TXLSHorizPagebreak; overload;
     function Add(const ARow: integer): TXLSHorizPagebreak; overload;

     function Find(const ARow: integer): TXLSHorizPagebreak;
     function FindIndex(const ARow: integer): integer;

     property Items[Index: integer]: TXLSHorizPagebreak read GetItems; default;
     end;

type TXLSVertPagebreak = class(TObject)
protected
     FXc12PageBreak: TXc12Break;

     function  GetCol: integer;
     function  GetRow1: integer;
     function  GetRow2: integer;
     procedure SetCol(const Value: integer);
     procedure SetRow1(const Value: integer);
     procedure SetRow2(const Value: integer);
public
     constructor Create(APageBreak: TXc12Break);

     property Col: integer read GetCol write SetCol;
     property Row1: integer read GetRow1 write SetRow1;
     property Row2: integer read GetRow2 write SetRow2;
     end;

type TXLSVertPagebreaks = class(TObjectList)
private
     function GetItems(Index: integer): TXLSVertPagebreak;
protected
     FXc12PageBreaks: TXc12PageBreaks;
public
     constructor Create(AXc12PageBreaks: TXc12PageBreaks);

     function Add: TXLSVertPagebreak; overload;
     function Add(const ACol: integer): TXLSVertPagebreak; overload;

     function Find(const ACol: integer): TXLSVertPagebreak;
     function FindIndex(const ACol: integer): integer;

     property Items[Index: integer]: TXLSVertPagebreak read GetItems; default;
     end;

// Formatting codes for heder/footer.
// There is no required order in which these codes must appear.
// The first occurrence of the following codes turns the formatting ON, the second occurrence turns it OFF again:
//   - strikethrough
//   - superscript
//   - subscript
// Superscript and subscript cannot both be ON at same time. Whichever comes first wins and the other is ignored,
// while the first is ON.

// &L - code for "left section" (there are three header / footer locations, "left", "center", and "right"). When
//   two or more occurrences of this section marker exist, the contents from all markers are concatenated, in the
//   order of appearance, and placed into the left section.
// &P - code for "current page #"
// &N - code for "total pages"
// &font size - code for "text font size", where font size is a font size in points.
// &K - code for "text font color"
//   RGB Color is specified as RRGGBB
//   Theme Color is specifed as TTSNN where TT is the theme color Id, S is either "+" or "-" of the tint/shade
//   value, NN is the tint/shade value.
// &S - code for "text strikethrough" on / off
// &X - code for "text super script" on / off
// &Y - code for "text subscript" on / off
// &C - code for "center section". When two or more occurrences of this section marker exist, the contents
//   from all markers are concatenated, in the order of appearance, and placed into the center section.
// &D - code for "date"
// &T - code for "time"
// &G - code for "picture as background"
// &U - code for "text single underline"
// &E - code for "double underline"
// &R - code for "right section". When two or more occurrences of this section marker exist, the contents
//   from all markers are concatenated, in the order of appearance, and placed into the right section.
// &Z - code for "this workbook's file path"
// &F - code for "this workbook's file name"
// &A - code for "sheet tab name"
// &+ - code for add to page #.
// &- - code for subtract from page #.
// &"font name,font type" - code for "text font name" and "text font type", where font name and font type
//   are strings specifying the name and type of the font, separated by a comma. When a hyphen appears in font
//   name, it means "none specified". Both of font name and font type can be localized values.
// &"-,Bold" - code for "bold font style"
// &B - also means "bold font style".
// &"-,Regular" - code for "regular font style"
// &"-,Italic" - code for "italic font style"
// &I - also means "italic font style"
// &"-,Bold Italic" code for "bold italic font style"
// &O - code for "outline style"
// &H - code for "shadow style"

type TXLSHeaderFooter = class(TObject)
private
     function  GetAlignWithMargins: boolean;
     function  GetDifferentFirst: boolean;
     function  GetDifferentOddEven: boolean;
     function  GetEvenFooter: AxUCString;
     function  GetEvenHeader: AxUCString;
     function  GetFirstFooter: AxUCString;
     function  GetFirstHeader: AxUCString;
     function  GetOddFooter: AxUCString;
     function  GetOddHeader: AxUCString;
     function  GetScaleWithDoc: boolean;
     procedure SetAlignWithMargins(const Value: boolean);
     procedure SetDifferentFirst(const Value: boolean);
     procedure SetDifferentOddEven(const Value: boolean);
     procedure SetEvenFooter(const Value: AxUCString);
     procedure SetEvenHeader(const Value: AxUCString);
     procedure SetFirstFooter(const Value: AxUCString);
     procedure SetFirstHeader(const Value: AxUCString);
     procedure SetOddFooter(const Value: AxUCString);
     procedure SetOddHeader(const Value: AxUCString);
     procedure SetScaleWithDoc(const Value: boolean);
protected
     FXc12HeaderFooter: TXc12HeaderFooter;
public
     constructor Create(AXc12HeaderFooter: TXc12HeaderFooter);

     procedure Clear;

     property DifferentOddEven: boolean read GetDifferentOddEven write SetDifferentOddEven;
     property DifferentFirst: boolean read GetDifferentFirst write SetDifferentFirst;
     property ScaleWithDoc: boolean read GetScaleWithDoc write SetScaleWithDoc;
     property AlignWithMargins: boolean read GetAlignWithMargins write SetAlignWithMargins;

     ///	<summary>
     ///	  Default header when DifferentOddEven is False.
     ///	</summary>
     property OddHeader: AxUCString read GetOddHeader write SetOddHeader;
     ///	<summary>
     ///	  Default footer when DifferentOddEven is False.
     ///	</summary>
     property OddFooter: AxUCString read GetOddFooter write SetOddFooter;
     property EvenHeader: AxUCString read GetEvenHeader write SetEvenHeader;
     property EvenFooter: AxUCString read GetEvenFooter write SetEvenFooter;
     property FirstHeader: AxUCString read GetFirstHeader write SetFirstHeader;
     property FirstFooter: AxUCString read GetFirstFooter write SetFirstFooter;
     end;

type TXLSPrintSettings = class(TObject)
private
     function  GetOptions: TPrintSetupOptions;
     procedure SetOptions(const Value: TPrintSetupOptions);
     function  GetCopies: integer;
     function  GetFooterMargin: double;
     function  GetHeaderMargin: double;
     function  GetPaperSize: TXc12PaperSize;
     function  GetScalingFactor: integer;
     function  GetStartingPage: integer;
     procedure SetCopies(const Value: integer);
     procedure SetFooterMargin(const Value: double);
     procedure SetHeaderMargin(const Value: double);
     procedure SetPaperSize(const Value: TXc12PaperSize);
     procedure SetScalingFactor(const Value: integer);
     procedure SetStartingPage(const Value: integer);
     function  GetResolution: integer;
     procedure SetResolution(const Value: integer);
     function  GetMarginBottomCm: double;
     function  GetMarginLeftCm: double;
     function  GetMarginRightCm: double;
     function  GetMarginTopCm: double;
     procedure SetMarginBottomCm(const Value: double);
     procedure SetMarginLeftCm(const Value: double);
     procedure SetMarginRightCm(const Value: double);
     procedure SetMarginTopCm(const Value: double);
     function  GetFitHeight: integer;
     function  GetFitWidth: integer;
     procedure SetFitHeight(const Value: integer);
     procedure SetFitWidth(const Value: integer);
     function  GetFooterMarginCm: double;
     function  GetHeaderMarginCm: double;
     procedure SetFooterMarginCm(const Value: double);
     procedure SetHeaderMarginCm(const Value: double);
     function  GetMarginBottom: double;
     function  GetMarginLeft: double;
     function  GetMarginRight: double;
     function  GetMarginTop: double;
     procedure SetMarginBottom(const Value: double);
     procedure SetMarginLeft(const Value: double);
     procedure SetMarginRight(const Value: double);
     procedure SetMarginTop(const Value: double);
protected
     FXc12Sheet      : TXc12DataWorksheet;
     FNames          : TXLSNames;

     FXLSHeaderFooter: TXLSHeaderFooter;
     FVertPagebreaks : TXLSVertPagebreaks;
     FHorizPagebreaks: TXLSHorizPagebreaks;
public
     constructor Create(AXc12Sheet: TXc12DataWorksheet; ANames: TXLSNames);
     destructor Destroy; override;

     procedure PrintArea(const ACol1,ARow1,ACol2,ARow2: integer);
     function  GetPrintArea(out ACol1,ARow1,ACol2,ARow2: integer): boolean;
     procedure ClearPrintArea;

     //* Set columns (to the left) and rows (at the top)
     //* that are repeated on each page. To set only one of them, set
     //* the other to -1.
     procedure PrintTitles(const ACol1,ACol2,ARow1,ARow2: integer);
     function  GetPrintTitles(out ACol1,ACol2,ARow1,ARow2: integer): boolean;
     function  GetPrintTitlesRows(out ARow1,ARow2: integer): boolean;
     function  GetPrintTitlesCols(out ACol1,ACol2: integer): boolean;
     procedure ClearPrintTitles;

     procedure Clear;

     function PaperSizeDim: TXLSPointF;

     property Copies: integer read GetCopies write SetCopies;
     property HeaderFooter: TXLSHeaderFooter read FXLSHeaderFooter;
     property FooterMargin: double read GetFooterMargin write SetFooterMargin;
     property FooterMarginCm: double read GetFooterMarginCm write SetFooterMarginCm;
     property HeaderMargin: double read GetHeaderMargin write SetHeaderMargin;
     property HeaderMarginCm: double read GetHeaderMarginCm write SetHeaderMarginCm;
     property MarginBottom: double read GetMarginBottom write SetMarginBottom;
     property MarginLeft: double read GetMarginLeft write SetMarginLeft;
     property MarginRight: double read GetMarginRight write SetMarginRight;
     property MarginTop: double read GetMarginTop write SetMarginTop;
     property MarginBottomCm: double read GetMarginBottomCm write SetMarginBottomCm;
     property MarginLeftCm: double read GetMarginLeftCm write SetMarginLeftCm;
     property MarginRightCm: double read GetMarginRightCm write SetMarginRightCm;
     property MarginTopCm: double read GetMarginTopCm write SetMarginTopCm;
     property Options: TPrintSetupOptions read GetOptions write SetOptions;
     property PaperSize: TXc12PaperSize read GetPaperSize write SetPaperSize;
     property ScalingFactor: integer read GetScalingFactor write SetScalingFactor;
     property StartingPage: integer read GetStartingPage write SetStartingPage;
     property HorizPagebreaks: TXLSHorizPagebreaks read FHorizPagebreaks;
     property VertPagebreaks: TXLSVertPagebreaks read FVertPagebreaks;
     property Resolution: integer read GetResolution write SetResolution;
     property FitWidth: integer read GetFitWidth write SetFitWidth;
     property FitHeight: integer read GetFitHeight write SetFitHeight;
     end;

type TXLSPane = class(TObject)
private
     FXc12Pane: TXc12Pane;

     function  GetActivePane: TActivePane;
     function  GetLeftCol: integer;
     function  GetPaneType: TPaneType;
     function  GetSplitColX: double;
     function  GetSplitRowY: double;
     function  GetTopRow: integer;
     procedure SetActivePane(const Value: TActivePane);
     procedure SetLeftCol(const Value: integer);
     procedure SetPaneType(const Value: TPaneType);
     procedure SetSplitColX(const Value: double);
     procedure SetSplitRowY(const Value: double);
     procedure SetTopRow(const Value: integer);
public
     constructor Create(const AXc12Pane: TXc12Pane);

     //* ~exclude
     procedure Clear;
     //* Which pane that is active, i.e the pane with the cursor.
     property ActivePane: TActivePane read GetActivePane write SetActivePane;
     //* Pane style.
     property PaneType: TPaneType read GetPaneType write SetPaneType;
     //* Vertical split.
     //* Split pane (PaneType = ptSplit): width of left pane in units of 1/20
     //* of a point.~[br]
     //* Frozen pane (PaneType = ptFrozen): Number of visible columns in left pane.~[br]
     //* If there not shall be a vertical split, set SplitColX to zero.
     property SplitColX: double read GetSplitColX write SetSplitColX;
     //* Horizontal split.
     //* Split pane (PaneType = ptSplit): height of top pane in units of 1/20
     //* of a point.~[br]
     //* Frozen pane (PaneType = ptFrozen): Number of visible rows in top pane.
     //* If there not shall be a horizontal split, set SplitColY to zero.
     property SplitRowY: double read GetSplitRowY write SetSplitRowY;
     //* First visible column in right pane.
     property LeftCol: integer read GetLeftCol write SetLeftCol;
     //* First visible row in bottom pane.
     property TopRow: integer read GetTopRow write SetTopRow;
     end;

type TXLSSheetProtection = class(TObject)
private
     function  GetAutoFilter: boolean;
     function  GetDeleteColumns: boolean;
     function  GetDeleteRows: boolean;
     function  GetFormatCells: boolean;
     function  GetFormatColumns: boolean;
     function  GetFormatRows: boolean;
     function  GetInsertColumns: boolean;
     function  GetInsertHyperlinks: boolean;
     function  GetInsertRows: boolean;
     function  GetObjects: boolean;
     function  GetPasswordString: AxUCString;
     function  GetPivotTables: boolean;
     function  GetScenarios: boolean;
     function  GetSelectLockedCells: boolean;
     function  GetSelectUnlockedCells: boolean;
     function  GetSheet: boolean;
     function  GetSort: boolean;
     procedure SetAutoFilter(const Value: boolean);
     procedure SetDeleteColumns(const Value: boolean);
     procedure SetDeleteRows(const Value: boolean);
     procedure SetFormatCells(const Value: boolean);
     procedure SetFormatColumns(const Value: boolean);
     procedure SetFormatRows(const Value: boolean);
     procedure SetInsertColumns(const Value: boolean);
     procedure SetInsertHyperlinks(const Value: boolean);
     procedure SetInsertRows(const Value: boolean);
     procedure SetObjects(const Value: boolean);
     procedure SetPasswordString(const Value: AxUCString);
     procedure SetPivotTables(const Value: boolean);
     procedure SetScenarios(const Value: boolean);
     procedure SetSelectLockedCells(const Value: boolean);
     procedure SetSelectUnlockedCells(const Value: boolean);
     procedure SetSheet(const Value: boolean);
     procedure SetSort(const Value: boolean);
protected
     FProtection: TXc12SheetProtection;
     FAssigned  : boolean;
public
     constructor Create(AProtection: TXc12SheetProtection);

     procedure Clear;

     property Password: AxUCString read GetPasswordString write SetPasswordString;
     property Sheet: boolean read GetSheet write SetSheet;
     property Objects: boolean read GetObjects write SetObjects;
     property Scenarios: boolean read GetScenarios write SetScenarios;
     property FormatCells: boolean read GetFormatCells write SetFormatCells;
     property FormatColumns: boolean read GetFormatColumns write SetFormatColumns;
     property FormatRows: boolean read GetFormatRows write SetFormatRows;
     property InsertColumns: boolean read GetInsertColumns write SetInsertColumns;
     property InsertRows: boolean read GetInsertRows write SetInsertRows;
     property InsertHyperlinks: boolean read GetInsertHyperlinks write SetInsertHyperlinks;
     property DeleteColumns: boolean read GetDeleteColumns write SetDeleteColumns;
     property DeleteRows: boolean read GetDeleteRows write SetDeleteRows;
     property SelectLockedCells: boolean read GetSelectLockedCells write SetSelectLockedCells;
     property Sort: boolean read GetSort write SetSort;
     property AutoFilter: boolean read GetAutoFilter write SetAutoFilter;
     property PivotTables: boolean read GetPivotTables write SetPivotTables;
     property SelectUnlockedCells: boolean read GetSelectUnlockedCells write SetSelectUnlockedCells;
     end;

type TXLSWorkbook = class;
     TXLSWorksheet = class;

     TXLSRelCellsImpl = class(TXLSRelCells)
protected
     FSheet: TXLSWorksheet;

     function  Clone: TXLSRelCells; override;
     function  GetAsFloat(ACol,ARow: integer): double; override;
     function  GetAsString(ACol,ARow: integer): AxUCString; override;
     function  GetAsBoolean(ACol, ARow: integer): boolean; override;
     function  GetAsError(ACol, ARow: integer): TXc12CellError; override;
     function  GetAsBlank(ACol, ARow: integer): boolean; override;
     function  GetAsEmpty(ACol,ARow: integer): boolean; override;
     function  GetAsFmtString(ACol,ARow: integer): AxUCString; override;
     function  GetAsSharedItem(ACol,ARow: integer): TXLSSharedItemsValue; override;
     function  GetCellType(ACol, ARow: integer): TXLSCellType; override;
     function  GetIsDateTime(Col, Row: integer): boolean; override;
     procedure SetAsBoolean(ACol, ARow: integer; const AValue: boolean); override;
     procedure SetAsFloat(ACol, ARow: integer; const AValue: double); override;
     procedure SetAsString(ACol, ARow: integer; const AValue: AxUCString); override;
     procedure SetAsError(ACol, ARow: integer; const AValue: TXc12CellError); override;
     procedure SetAsBlank(ACol, ARow: integer; const Value: boolean); override;
     procedure SetAsSharedItem(ACol,ARow: integer; Value: TXLSSharedItemsValue); override;
public
     constructor Create(ASheet: TXLSWorksheet);

     procedure ClearAll(ACol1,ARow1,ACol2,ARow2: integer); override;

     function  SheetName: AxUCString; override;
     function  SheetIndex: integer; override;

     procedure AutoWidthCol(ACol: integer); override;

     procedure BeginCommand; override;
     procedure ApplyCommand; overload; override;
     procedure ApplyCommand(const ACol,ARow: integer); overload; override;
     procedure ApplyCommand(ACol1,ARow1,ACol2,ARow2: integer); overload; override;

     property Sheet: TXLSWorksheet read FSheet write FSheet;
     end;

     TXLSWorksheet = class(TIndexObject)
protected
     FManager       : TXc12Manager;
     FClassFactory  : TXLSClassFactory;

     FOwner         : TXLSWorkbook;

     FXc12Sheet     : TXc12DataWorksheet;
     FXc12SheetView : TXc12SheetView;
     FOwnsSheetView : boolean;
     // Help variable for FXc12Sheet.Cells
     FCells         : TXLSCellMMU;

     FTheCell       : TXLSCell;
     FTheCellFormat : TXLSFormattedCell;
     FRange         : TXLSRange;

     FRows          : TXLSRows;
     FColumns       : TXLSColumns;

     FMergedCells   : TXLSMergedCells;
     FValidations   : TXLSDataValidations;
     FHyperlinks    : TXLSHyperlinks;
     FCondFormats   : TXLSConditionalFormats;
     FAutofilter    : TXLSAutofilter;
     FComments      : TXLSComments;

     FDrawing       : TXLSDrawing;
     FPivotTables   : TXLSPivotTables;

     FSelectedAreas : TXLSSelectedAreas;
     FVisibleAreas  : TCellAreas;
     FPrintSettings : TXLSPrintSettings;
     FPane          : TXLSPane;

     FProtection    : TXLSSheetProtection;

     FTabColor      : TColor;

     FCurrFindPos   : integer;
     FCurrFindCol   : integer;
     FCurrFindRow   : integer;

     FDefDecimals   : integer;

     FXSSCellInvalidateEvent: TXc12RefEvent;

     function  GetRows: TXLSRows;
     function  GetColumns: TXLSColumns;

     function  GetAsFloat(ACol,ARow: integer): double;
     procedure SetAsFloat(ACol,ARow: integer; AValue: double);
     function  GetAsBlank(ACol, ARow: integer): boolean;
     function  GetAsEmpty(ACol,ARow: integer): boolean;
     procedure SetAsBlank(ACol, ARow: integer; const Value: boolean);
     function  GetAutofilter: TXLSAutofilter;
     function  GetDefaultRowHeight: integer;
     procedure SetDefaultRowHeight(const Value: integer);
     function  GetHyperlinks: TXLSHyperlinks;
     function  GetSelectedAreas: TXLSSelectedAreas;
     function  GetMergedCells: TXLSMergedCells;
     function  GetPrintSettings: TXLSPrintSettings;
     function  GetOptions: TXLSSheetOptions;
     procedure SetOptions(const Value: TXLSSheetOptions);
     function  GetDrawing: TXLSDrawing;
     function  GetVisibleAreas: TCellAreas;
     function  GetCell(ACol, ARow: integer): TXLSCell;
     function  GetConditionalFormats: TXLSConditionalFormats;
     function  GetLeftCol: integer;
     function  GetTopRow: integer;
     procedure SetLeftCol(const Value: integer);
     procedure SetTopRow(const Value: integer);
     function  GetZoom: integer;
     procedure SetZoom(const Value: integer);
     function  GetName: AxUCString;
     function  GetTabColor: longword;
     procedure SetName(const Value: AxUCString);
     procedure SetTabColor(const Value: longword);
     function  GetAsBlankRef(const ARef: AxUCString): boolean;
     function  GetAsBoolean(ACol, ARow: integer): boolean;
     function  GetAsBooleanRef(const ARef: AxUCString): boolean;
     function  GetAsBoolFormulaValue(ACol, ARow: integer): boolean;
     function  GetAsBoolFormulaValueRef(const ARef: AxUCString): boolean;
     function  GetAsDateTime(ACol, ARow: integer): TDateTime;
     function  GetAsDateTimeRef(const ARef: AxUCString): TDateTime;
     function  GetAsError(ACol, ARow: integer): TXc12CellError;
     function  GetAsErrorRef(const ARef: AxUCString): TXc12CellError;
     function  GetAsFloatRef(const ARef: AxUCString): double;
     function  GetAsFormula(ACol, ARow: integer): AxUCString; overload;
     function  GetAsFormulaRef(const ARef: AxUCString): AxUCString;
     function  GetAsHTML(ACol, ARow: integer): AxUCString;
     function  GetAsHTMLRef(const ARef: AxUCString): AxUCString;
     function  GetAsHyperlink(ACol, ARow: integer): AxUCString;
     function  GetAsInteger(ACol, ARow: integer): integer;
     function  GetAsIntegerRef(const ARef: AxUCString): integer;
     function  GetAsNumFormulaValue(ACol, ARow: integer): double;
     function  GetAsNumFormulaValueRef(const ARef: AxUCString): double;
     function  GetAsRichText(ACol, ARow: integer): AnsiString;
     function  GetAsRichTextRef(const ARef: AxUCString): AxUCString;
     function  GetAsSimpleTags(ACol, ARow: integer): AxUCString;
     function  GetAsSimpleTagsRef(const ARef: AxUCString): AxUCString;
     function  GetAsStrFormulaValue(ACol, ARow: integer): AxUCString;
     function  GetAsStrFormulaValueRef(const ARef: AxUCString): AxUCString;
     function  GetAsString(ACol, ARow: integer): AxUCString;
     function  GetAsStringRef(const ARef: AxUCString): AxUCString;
     function  GetAsVariant(ACol, ARow: integer): Variant;
     function  GetAsVariantRef(const ARef: AxUCString): Variant;
     function  GetCellType(ACol, ARow: integer): TXLSCellType;
     function  GetFormulaType(ACol, ARow: integer): TXLSCellFormulaType;
     function  GetDefaultColWidth: integer;
     function  GetAsFmtString(ACol, ARow: integer): AxUCString;
     function  GetAsFmtStringRef(const ARef: AxUCString): AxUCString;
     function  GetIsDateTime(ACol, ARow: integer): boolean;
     function  GetRecalcFormulas: boolean;
     function  GetValidations: TXLSDataValidations;
     function  GetWorkspaceOptions: TWorkspaceOptions;
     function  GetZoomPreview: integer;

     procedure SetAsBlankRef(const ARef: AxUCString; const Value: boolean);
     procedure SetAsBoolean(ACol, ARow: integer; const Value: boolean);
     procedure SetAsBooleanRef(const ARef: AxUCString; const Value: boolean);
     procedure SetAsBoolFormulaValue(ACol, ARow: integer; const Value: boolean);
     procedure SetAsBoolFormulaValueRef(const ARef: AxUCString; const Value: boolean);
     procedure SetAsDateTime(ACol, ARow: integer; const Value: TDateTime);
     procedure SetAsDateTimeRef(const ARef: AxUCString; const Value: TDateTime);
     procedure SetAsError(ACol, ARow: integer; const Value: TXc12CellError);
     procedure SetAsErrorRef(const ARef: AxUCString; const Value: TXc12CellError);
     procedure SetAsFloatRef(const ARef: AxUCString; const Value: double);
     procedure SetAsFormula(ACol, ARow: integer; const Value: AxUCString);
     procedure SetAsFormulaRef(const ARef: AxUCString; const Value: AxUCString);
     procedure SetAsHyperlink(ACol, ARow: integer; const Value: AxUCString);
     procedure SetAsInteger(ACol, ARow: integer; const Value: integer);
     procedure SetAsIntegerRef(const ARef: AxUCString; const Value: integer);
     procedure SetAsNumFormulaValue(ACol, ARow: integer; const Value: double);
     procedure SetAsNumFormulaValueRef(const ARef: AxUCString; const Value: double);
     procedure SetAsRichText(ACol, ARow: integer; const Value: AnsiString);
     procedure SetAsRichTextRef(const ARef: AxUCString; const Value: AxUCString);
     procedure SetAsSimpleTags(ACol, ARow: integer; const Value: AxUCString);
     procedure SetAsSimpleTagsRef(const ARef, Value: AxUCString);
     procedure SetAsStrFormulaValue(ACol, ARow: integer; const Value: AxUCString);
     procedure SetAsStrFormulaValueRef(const ARef: AxUCString; const Value: AxUCString);
     procedure SetAsString(ACol, ARow: integer; const Value: AxUCString);
     procedure SetAsStringRef(const ARef: AxUCString; const Value: AxUCString);
     procedure SetAsVariant(ACol, ARow: integer; const Value: Variant);
     procedure SetAsVariantRef(const ARef: AxUCString; const Value: Variant);

     procedure SetDefaultColWidth(const Value: integer);
     procedure SetRecalcFormulas(const Value: boolean);
     procedure SetWorkspaceOptions(const Value: TWorkspaceOptions);
     procedure SetZoomPreview(const Value: integer);
     function  GetPane: TXLSPane;
     function  GetCellFormat(const ACol, ARow: integer): TXLSFormattedCell;
     function  GetAsArray(Col1, Row1, Col2, Row2, MaxLength: integer): AxUCString;
     function  GetVisibility: TXc12Visibility;
     procedure SetVisibility(const Value: TXc12Visibility);
     procedure SetIgnoreErrorNumbersAsText(const Value: boolean);

     procedure MoveCellObjects(ACol1, ARow1, ACol2, ARow2, ADeltaCol, ADeltaRow: integer);
     procedure DeleteCellObjects(ACol1, ARow1, ACol2, ARow2: integer);
     procedure CompileFormula(const ACol,ARow: integer; ACell: PXLSCellItem);
     procedure CompileFormulas;
     function  GetDefaultFormat(const ACol, ARow: integer): integer;

     procedure ColWidthChanged(ASender: TObject; const ACol1,ACol2: integer);
     procedure RowHeightChanged(ASender: TObject; const ARow: integer);

     function  CompareCells(const ACell1,ACell2: PXLSCellItem; const ACaseSencitive: boolean): integer;
     function  CompareRowCells(const ACell1,ACell2: PXLSCellItem; const ARow1,ARow2,ACol1,ACol2: integer; const AAscending,ACaseSencitive: boolean): integer;

     procedure AdjustCell(ACol,ARow,ADestCol,ADestRow: integer; ALockStartRow: boolean);
     procedure AdjustColumnsFormulas(const ACol,ADeltaCols: integer);
     procedure AdjustRowsFormulas(const ARow,ADeltaRows: integer);

     procedure AfterRead;

     function  DrawingAddImage97(AAnchor: TCT_TwoCellAnchor): boolean;
public
     constructor Create(AClassFactory: TXLSClassFactory; AOwner: TXLSWorkbook; AManager: TXc12Manager; AXc12Sheet: TXc12DataWorksheet);
     destructor Destroy; override;

     //* For internal use.
     function  RequestObject(AClass: TClass): TObject; override;
     //* For internal use.
     function  MakeFormattedText(const ACell: TXLSCellItem): TXLSFormattedText; overload;
     //* For internal use.
     function  MakeFormattedText(const ACol,ARow: integer; var AFormattedText: TXLSFormattedText): boolean; overload;

{$ifdef _AXOLOT_DEBUG}
     function  CountFormulas: integer;
     procedure DumpMem(const AList: TStrings);
{$endif}

     procedure Assign(ASource: TXLSWorksheet);

     //* For internal use.
     function  MMUCells: TXLSCellMMU; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}

     //* Clear all data of the worksheet and set it to default values.
     procedure Clear;
     procedure ClearData; {$ifdef DELPHI_2007_OR_LATER} deprecated 'Use Clear'; {$endif}
     //* Deletes all cell values and drawing objects. Equivalent to the Excel command Editing->Clear->Clear All.
     procedure ClearWorksheet;

     procedure MakePivotTable(ACol1,ARow1,ACol2,ARow2: integer);

     //* Fills an area with random numbers, from 0 to Value - 1. Usefull for test purpose.
     //* ~param Ref The cell area as a string, like: A1:D9
     //* ~param Value Max random value - 1.
     procedure FillRandom(const AArea: AxUCString;const AValue: integer);

     //* Sets the width of the column to the widest text of the cells in the column.
     //* See also $link(AutoWidthCols).
     function  AutoWidthCol(const ACol: integer; AMinWidth: integer = 0; AMaxWidth: integer = MAXINT): integer;
     //* Sets the width of the columns between Col1 and Col2 to the widest text of the cells in the column.
     //* See also $link(AutoWidthCol).
     procedure AutoWidthCols(const Col1,Col2: integer; AMinWidth: integer = 0; AMaxWidth: integer = MAXINT);
     //* Sets the height of the row to the text max height of the cells in the row.
     //* See also $link(AutoHeightRows).
     function  AutoHeightRow(const ARow: integer; AMinHeight: integer = 0; AMaxHeight: integer = MAXINT): boolean;
     //* Sets the height of the rows between Row1 and Row2 to the text max height of the cells in the row.
     //* See also $link(AutoHeightRow).
     function  AutoHeightRows(const Row1,Row2: integer; AMinHeight: integer = 0; AMaxHeight: integer = MAXINT): boolean;
     procedure AutoSizeCell(const ACol,ARow: integer; const ASizeCol,ASizeRow: boolean);

     //* Removes frozen panes created with @link(FreezePanes).
     procedure UnfreezePanes;
     //* Create frozen panes, that is, top rows and/or Left columns are locked.
     //* AColumns is the number of locked columns. Set to zero fo no locked columns.
     //* ARows is the number of locked rows. Set to zero fo no locked rows.
     //* Remove frozen panes with @linkt(UnfreezePanes).
     procedure FreezePanes(const AColumns,ARows: integer);

     //* Removes splitted panes created with @link(SplitPanes).
     procedure UnsplitPanes;
     //* Creates splitted panes.
     //* The width and height is in character points.
     //* Remove splitted panes with @linkt(UnsplitPanes).
     procedure SplitPanes(const APointsWidth,APointsHeight: integer);

     //* Calculate the dimensions of the worksheet.
     //* CalcDimensions will calculate the first and last row, first and last column.
     // That is, it sets the values for @link(FirstCol), @link(FirstRow), @link(LastCol) and @link(LastRow).
     procedure CalcDimensions;
     //* Calculate the dimensions of the worksheet.
     //* CalcDimensions will calculate the first and last, first and last column.
     //* Only cells with values are included. Cells with only formatting (blank)
     //* are excluded.
     //* Consider using CalcDimensions instead as CalcDimensionsEx id *MUCH* slower.
     procedure CalcDimensionsEx;

     function  Calculate(const ACol, ARow: integer): Variant;
     function  CalculateEx(Col, Row: integer; CalculateOptions: boolean): Variant; {deprecated 'Use Calculate'; {$endif}
     function  CalculateAsString(const ACol, ARow: integer): AxUCString;
     procedure CalculateSheet; {$ifdef DELPHI_2007_OR_LATER} deprecated 'Use Workbook.Calculate'; {$endif}
     procedure Calculated(const IsCalculated: boolean); {$ifdef DELPHI_2007_OR_LATER} deprecated 'Not used anymore'; {$endif}

     procedure CopyCell(const ACol, ARow, ADestCol, ADestRow: integer; ALockStartRow: boolean; AClearFmla: boolean = False);
     procedure CopyCells(const ACol1, ARow1, ACol2, ARow2, ADestCol, ADestRow: integer; const ACopyOptions: TCopyCellsOptions = CopyAllCells);

     procedure DeleteCell(const ACol, ARow: integer; const AOptions: TXLSDeleteOptions = DefaultDeleteOptions);
     procedure DeleteCells(const ACol1, ARow1, ACol2, ARow2: integer; const AOptions: TXLSDeleteOptions = DefaultDeleteOptions);
     //*        Deletes just the cell values
     procedure ClearCell(const ACol, ARow: integer);
     procedure ClearCells(const ACol1, ARow1, ACol2, ARow2: integer);

     procedure MoveCell(const ACol, ARow, ADestCol, ADestRow: integer; AAdjustCell: boolean = False; ALockStartRow: boolean = False);
     procedure MoveCells(const ACol1, ARow1, ACol2, ARow2, ADestCol, ADestRow: integer; const ACopyOptions: TCopyCellsOptions = CopyAllCells);

     procedure MergeCells; overload;
     function  MergeCells(Col1, Row1, Col2, Row2: integer): boolean; overload;
     procedure UnMergeCells; overload;
     procedure UnMergeCells(Col1, Row1, Col2, Row2: integer); overload;

     //* Deletes the rows between ARow1 and ARow2 and shiftes the rows below up.
     //* See also @link(ClearRows).
     procedure DeleteRows(const ARow1,ARow2: integer);
     //* Deletes the rows between ARow1 and ARow2.
     //* See also @link(DeleteRows).
     procedure ClearRows(const ARow1,ARow2: integer);
     procedure InsertRows(const ARow,ARowCount: integer);
     procedure CopyRows(const ARow1,ARow2,ADestRow: integer);
     procedure MoveRows(const ARow1,ARow2,ADestRow: integer);

     //* Deletes the columns between ACol1 and ACol2 and shiftes the columns to the right left.
     //* See also @link(ClearColumns).
     procedure DeleteColumns(const ACol1,ACol2: integer);
     //* Deletes the columns between ACol1 and ACol2.
     //* See also @link(ClearColumns).
     procedure ClearColumns(const ACol1,ACol2: integer);
     procedure InsertColumns(const ACol,AColCount: integer);
     procedure CopyColumns(const ACol1,ACol2,ADestCol: integer);
     procedure MoveColumns(const ACol1,ACol2,ADestCol: integer);

     function  IsEmpty: boolean;

     //* Returns True if the sheet is protected and the cell is locked.
     function  IsCellLocked(const ACol,ARow: integer): boolean;

     // Sorts on selected area.
     procedure Sort(const AAscending,ACaseSencitive: boolean); overload;
     procedure Sort(ACol1, ARow1, ACol2, ARow2: integer; const AAscending,ACaseSencitive: boolean); overload;

     procedure GroupRows(const ARow1,ARow2: integer; ACollapsed: boolean = False);
     procedure GroupColumns(const ACol1, ACol2: integer; ACollapsed: boolean = False);
     //! Ungroups the rows betweem ARow1 and Row2. All rows must have the same
     //! grouping (outline) level. If the level is higher than one, the level
     //! is shifted one step down. That is, if the grouping level is three, it
     //! will become two.
     function  UngroupRows(const ARow1,ARow2: integer): boolean;
     //! Ungroups the columnss betweem ACol1 and ACol2. All columnss must have the same
     //! grouping (outline) level. If the level is higher than one, the level
     //! is shifted one step down. That is, if the grouping level is three, it
     //! will become two.
     function  UngroupColumns(const ACol1, ACol2: integer): boolean;

     procedure RichTextCell(const ACol, ARow: integer; const AText: AxUCString; AFontRuns: TXc12DynFontRunArray);
     procedure RichTextLoadFromFile(const ACol, ARow: integer; const Filename: AxUCString; const AllText: boolean = True);
     procedure RichTextLoadFromStream(const ACol, ARow: integer; Stream: TStream; const AllText: boolean = True);
     procedure RichTextSaveToFile(const ACol, ARow: integer; const Filename: AxUCString);
     procedure RichTextSaveToStream(const ACol, ARow: integer; Stream: TStream);
{$ifndef BABOON}
     procedure CopyToRichEdit(const ACol,ARow: integer; var ARichEdit: TRichEdit);
{$endif}

     procedure ImportRichText(ACol, ARow: integer; AStream: TStream); overload;
     procedure ImportRichText(ACol, ARow: integer; AFilename: AxUCString); overload;

     function  FindCell(const ACol,ARow: integer): boolean;

     procedure BeginFindText;
     function  FindText(const AText: AxUCString; CaseInsensitive: boolean = True): boolean;
     procedure GetFindData(var ACol,ARow,ATextPos: integer; var AText: AxUCString);

     function  GetFormatString(ACell: TXLSCellItem; AColor: PLongword): AxUCString;

     procedure InsertFloatRowValues(const ACol,ARow: integer; const AValues: array of double); overload;
     procedure InsertStringRowValues(const ACol,ARow: integer; const AValues: array of AxUCString); overload;
     procedure InsertRowValues(const ACol,ARow: integer; const AValues: array of Variant); overload;
     procedure InsertFloatColValues(const ACol,ARow: integer; const AValues: array of double); overload;
     procedure InsertStringColValues(const ACol,ARow: integer; const AValues: array of AxUCString); overload;
     procedure InsertColValues(const ACol,ARow: integer; const AValues: array of Variant); overload;

     procedure MakeHyperlink(const ACol,ARow: integer; const AAddress, AText: AxUCString); overload;
     procedure MakeHyperlink(const ACol,ARow: integer; const AAddress, AText,ATooltip: AxUCString); overload;

     procedure GetShapePixelSize(AShape: TXLSDrwTwoPosShape; out AWidth,AHeight: integer);

     procedure BeginIterate;
     function  IterateNext: boolean;
     procedure IteratePos(out ACol,ARow: integer);

     function  FirstCol: integer;
     function  LastCol: integer;
     function  FirstRow: integer;
     function  LastRow: integer;

     procedure ApplyAutofilter;

     procedure EditPivotTables;

     //! Create a relative cells object.
     function  CreateRelativeCells: TXLSRelCells; overload;
     function  CreateRelativeCells(ARefStr: AxUCString): TXLSRelCells; overload;
     function  CreateRelativeCells(ACol,ARow: integer): TXLSRelCells; overload;
     function  CreateRelativeCells(ACol1,ARow1,ACol2,ARow2: integer): TXLSRelCells; overload;
     function  CreateRelativeCells(AArea: TXLSCellArea): TXLSRelCells; overload;

     property LeftCol: integer read GetLeftCol write SetLeftCol;
     property TopRow: integer read GetTopRow write SetTopRow;
     property DefaultColWidth: integer read GetDefaultColWidth write SetDefaultColWidth;
     property DefaultRowHeight: integer read GetDefaultRowHeight write SetDefaultRowHeight;

     property AsBlank[Col,Row: integer]: boolean read GetAsBlank write SetAsBlank;
     property AsInteger[Col,Row: integer]: integer read GetAsInteger write SetAsInteger;
     property AsFloat[Col,Row: integer]: double read GetAsFloat write SetAsFloat;
     property AsDateTime[Col,Row: integer]: TDateTime read GetAsDateTime write SetAsDateTime;
     property AsString[Col,Row: integer]: AxUCString read GetAsString write SetAsString;
     property AsRichText[Col,Row: integer]: AnsiString read GetAsRichText write SetAsRichText;
    //* AsSimpleTags let you easily create formatted text with html-like tags.
    //* There are no end tags. A tag sets an option, and that option is used
    //* until a new tags replaces it. In order to write the "<", character,
    //* write an empty tag "<>" in the text. Example: 'Two <> one'
    //* The following tags can be uses:
    //*  <b>  Turn bold on.
    //* </b>  Turn bold off.
    //*  <i>  Turn italic on.
    //* </i>  Turn italic off.
    //*  <u>  Turn underline on.
    //* </u>  Turn underline off.
    //* <f:Font name> Changes the font name.
    //* <s:nn> Changes the font size. Setting the size to zero sets the font size to the default size.
    //* <c:xxxxxx> Changes the font color. The value xxxxxx is the hexadecimal RGB color value.
    //*            If xxxxxx is set to 'Auto', the default color will be used.
     property AsSimpleTags[ACol,ARow: integer]: AxUCString read GetAsSimpleTags write SetAsSimpleTags;
     property AsFmtString[Col,Row: integer]: AxUCString read GetAsFmtString;
     property AsHyperlink[Col,Row: integer]: AxUCString read GetAsHyperlink write SetAsHyperlink;
     property AsHTML[Col,Row: integer]: AxUCString read GetAsHTML;
     property AsBoolean[Col,Row: integer]: boolean read GetAsBoolean write SetAsBoolean;
     property AsError[Col,Row: integer]: TXc12CellError read GetAsError write SetAsError;
     property AsVariant[Col,Row: integer]: Variant read GetAsVariant write SetAsVariant;
     property AsEmpty[Col,Row: integer]: boolean read GetAsEmpty;

     property AsFormula[Col,Row: integer]: AxUCString read GetAsFormula write SetAsFormula;
     property AsNumFormulaValue[Col,Row: integer]: double read GetAsNumFormulaValue write SetAsNumFormulaValue;
     property AsStrFormulaValue[Col,Row: integer]: AxUCString read GetAsStrFormulaValue write SetAsStrFormulaValue;
     property AsBoolFormulaValue[Col,Row: integer]: boolean read GetAsBoolFormulaValue write SetAsBoolFormulaValue;

     property IsDateTime[Col,Row: integer]: boolean read GetIsDateTime;

     property AsArray[Col1,Row1,Col2,Row2,MaxLength: integer]: AxUCString read GetAsArray;

     property AsBlankRef[const ARef: AxUCString]: boolean read GetAsBlankRef write SetAsBlankRef;
     property AsIntegerRef[const ARef: AxUCString]: integer read GetAsIntegerRef write SetAsIntegerRef;
     property AsFloatRef[const ARef: AxUCString]: double read GetAsFloatRef write SetAsFloatRef;
     property AsDateTimeRef[const ARef: AxUCString]: TDateTime read GetAsDateTimeRef write SetAsDateTimeRef;
     property AsStringRef[const ARef: AxUCString]: AxUCString read GetAsStringRef write SetAsStringRef;
     property AsRichTextRef[const ARef: AxUCString]: AxUCString read GetAsRichTextRef write SetAsRichTextRef;
     property AsSimpleTagsRef[const ARef: AxUCString]: AxUCString read GetAsSimpleTagsRef write SetAsSimpleTagsRef;
     property AsFmtStringRef[const ARef: AxUCString]: AxUCString read GetAsFmtStringRef;
     property AsHTMLRef[const ARef: AxUCString]: AxUCString read GetAsHTMLRef;
     property AsBooleanRef[const ARef: AxUCString]: boolean read GetAsBooleanRef write SetAsBooleanRef;
     property AsErrorRef[const ARef: AxUCString]: TXc12CellError read GetAsErrorRef write SetAsErrorRef;
     property AsFormulaRef[const ARef: AxUCString]: AxUCString read GetAsFormulaRef write SetAsFormulaRef;
     property AsNumFormulaValueRef[const ARef: AxUCString]: double read GetAsNumFormulaValueRef write SetAsNumFormulaValueRef;
     property AsStrFormulaValueRef[const ARef: AxUCString]: AxUCString read GetAsStrFormulaValueRef write SetAsStrFormulaValueRef;
     property AsBoolFormulaValueRef[const ARef: AxUCString]: boolean read GetAsBoolFormulaValueRef write SetAsBoolFormulaValueRef;
     property AsVariantRef[const ARef: AxUCString]: Variant read GetAsVariantRef write SetAsVariantRef;

     // Number of decimals used when formatting numbers to strings. Default is -1,
     // Which means that the FloatToStr function is used.
     property DefDecimals: integer read FDefDecimals write FDefDecimals;

     property Cell[Col,Row: integer]: TXLSCell read GetCell;
     property CellFormat[const ACol,ARow: integer]: TXLSFormattedCell read GetCellFormat;
     property CellType[Col,Row: integer]: TXLSCellType read GetCellType;
     property FormulaType[Col,Row: integer]: TXLSCellFormulaType read GetFormulaType;

     property Rows: TXLSRows read GetRows;
     property Columns: TXLSColumns read GetColumns;
     property Autofilter: TXLSAutofilter read GetAutofilter;
     property Hyperlinks: TXLSHyperlinks read GetHyperlinks;
     property MergedCells: TXLSMergedCells read GetMergedCells;
     property ConditionalFormats: TXLSConditionalFormats read GetConditionalFormats;
     property Validations: TXLSDataValidations read GetValidations;
     property Comments: TXLSComments read FComments;

     property PrintSettings: TXLSPrintSettings read GetPrintSettings;
     property SelectedAreas: TXLSSelectedAreas read GetSelectedAreas;
     property VisibleAreas: TCellAreas read GetVisibleAreas;

     property Drawing: TXLSDrawing read GetDrawing;
     property PivotTables: TXLSPivotTables read FPivotTables;

     property Range: TXLSRange read FRange;
     property Pane: TXLSPane read GetPane;
     property Protection: TXLSSheetProtection read FProtection;

     property IgnoreErrorNumbersAsText: boolean write SetIgnoreErrorNumbersAsText;

     property Name: AxUCString read GetName write SetName;
     property TabColor: longword read GetTabColor write SetTabColor;
     property Options: TXLSSheetOptions read GetOptions write SetOptions;
     property WorkspaceOptions: TWorkspaceOptions read GetWorkspaceOptions write SetWorkspaceOptions;
     property Zoom: integer read GetZoom write SetZoom;
     property ZoomPreview: integer read GetZoomPreview write SetZoomPreview;
     property RecalcFormulas: boolean read GetRecalcFormulas write SetRecalcFormulas;
     property Visibility: TXc12Visibility read GetVisibility write SetVisibility;

     property Xc12Sheet: TXc12DataWorksheet read FXc12Sheet;
     property XSSCellInvalidateEvent: TXc12RefEvent read FXSSCellInvalidateEvent write FXSSCellInvalidateEvent;
     end;

     TXLSWorkbookData = class(TPersistent)
private
     function  ReadHeight: integer;
     function  ReadLeft: integer;
     function  ReadOptions: TWorkbookOptions;
     function  ReadSelectedTab: integer;
     function  ReadTop: integer;
     function  ReadWidth: integer;
     procedure WriteHeight(const Value: integer);
     procedure WriteLeft(const Value: integer);
     procedure WriteOptions(const Value: TWorkbookOptions);
     procedure WriteSelectedTab(const Value: integer);
     procedure WriteTop(const Value: integer);
     procedure WriteWidth(const Value: integer);
protected
     FBookViews: TXc12BookViews;
public
     constructor Create(ABookViews: TXc12BookViews);

     //* Left position of the window, in units if 1/20th of a point.
     property Left: integer read ReadLeft write WriteLeft;
     //* Top position of the window, in units if 1/20th of a point.
     property Top: integer read ReadTop write WriteTop;
     //* Width of the window, in units if 1/20th of a point.
     property Width: integer read ReadWidth write WriteWidth;
     //* Height of the window, in units if 1/20th of a point.
     property Height: integer read ReadHeight write WriteHeight;
     //* Index of selected workbook tab, zero based.
     property SelectedTab: integer read ReadSelectedTab write WriteSelectedTab;
     //* Options for the window.
     property Options: TWorkbookOptions read ReadOptions write WriteOptions;
     end;

     TXLSOptionsDialog = class(TPersistent)
private
     FCalcPr: TXc12CalcPr;
     FWorkbookPr: TXc12WorkbookPr;
     // Is transfered to Worksheet.Sheetviews in BeforeWrite.
     FRightToLeftAssigned: boolean;
     FRightToLeft: boolean;
     // Only for Excel 97.
     FSaveRecalc: boolean;

     function  GetShowObjects: TXc12Objects;
     procedure SetShowObjects(const Value: TXc12Objects);
     function  GetPrecisionAsDisplayed: boolean;
     procedure SetPrecisionAsDisplayed(const Value: boolean);
     function  GetSaveExtLinkVal: boolean;
     procedure SetSaveExtLinkVal(const Value: boolean);
     procedure SetMaxIterations(const Value: integer);
     function  GetMaxIterations: integer;
     function  GetCalcMode: TXc12CalcMode;
     function  GetDelta: double;
     function  GetIteration: boolean;
     function  GetR1C1Mode: boolean;
     function  GetRecalcBeforeSave: boolean;
     function  GetRightToLeft: boolean;
     function  GetSaveRecalc: boolean;
     function  GetUncalced: boolean;
     procedure SetCalcMode(const Value: TXc12CalcMode);
     procedure SetDelta(const Value: double);
     procedure SetIteration(const Value: boolean);
     procedure SetR1C1Mode(const Value: boolean);
     procedure SetRecalcBeforeSave(const Value: boolean);
     procedure SetRightToLeft(const Value: boolean);
     procedure SetSaveRecalc(const Value: boolean);
     procedure SetUncalced(const Value: boolean);
     function  GetDate1904: boolean;
     procedure SetDate1904(const Value: boolean);
public
     constructor Create(ACalcPr: TXc12CalcPr; AWorkbookPr: TXc12WorkbookPr);

     procedure Clear;

     //* TRue if the Save External Link Values option is turned off (Options dialog box, Calculation tab).
     property SaveExtLinkVal: boolean read GetSaveExtLinkVal write SetSaveExtLinkVal;
     //* Maximum Iterations option from the Options dialog box, Calculation tab.
     property CalcCount: integer read GetMaxIterations write SetMaxIterations;
     //* Maximum Iterations option from the Options dialog box, Calculation tab.
     property MaxIterations: integer read GetMaxIterations write SetMaxIterations;
     //* Calculation mode.
     property CalcMode: TXc12CalcMode read GetCalcMode write SetCalcMode;
     //* Maximum Change value from the Options dialog box, Calculation tab.
     property Delta: double read GetDelta write SetDelta;
     //* Which object to show.
     property ShowObjects: TXc12Objects read GetShowObjects write SetShowObjects;
     //* True if the iteration option is on.
     property Iteration: boolean read GetIteration write SetIteration;
     //* True if Precision As Displayed option is selected.
     property PrecisionAsDisplayed: boolean read GetPrecisionAsDisplayed write SetPrecisionAsDisplayed;
     //* True if RIC1 mode is on.
     property R1C1Mode: boolean read GetR1C1Mode write SetR1C1Mode;
     //* True if recalc before save is on.
     property RecalcBeforeSave: boolean read GetRecalcBeforeSave write SetRecalcBeforeSave;
     //* True if the workbook is marked as uncalced.
     property Uncalced: boolean read GetUncalced write SetUncalced;
     //* True if recalculate before saving is on.
     property SaveRecalc: boolean read GetSaveRecalc write SetSaveRecalc;
     //* Returns True if the 1904 date system is used.
     property Date1904: boolean read GetDate1904 write SetDate1904;

     property RightToLeft: boolean read GetRightToLeft write SetRightToLeft;
     end;

     TXLSWorkbook = class(TComponent)
private
     procedure SetDefaultName(const Value: AxUCString);
     function  GetFont: TXc12Font;
     function  GetNameIsSimpeName(const AName: AxUCString): boolean;
     function  GetNameAsBoolean(const AName: AxUCString): boolean;
     function  GetNameAsBoolFormulaValue(const AName: AxUCString): boolean;
     function  GetNameAsError(const AName: AxUCString): TXc12CellError;
     function  GetNameAsFloat(const AName: AxUCString): double;
     function  GetNameAsFmtString(const AName: AxUCString): AxUCString;
     function  GetNameAsFormula(const AName: AxUCString): AxUCString;
     function  GetNameAsNumFormulaValue(const AName: AxUCString): double;
     function  GetNameAsStrFormulaValue(const AName: AxUCString): AxUCString;
     function  GetNameAsString(const AName: AxUCString): AxUCString;
     procedure SetNameAsBoolean(const AName: AxUCString; const Value: boolean);
     procedure SetNameAsBoolFormulaValue(const AName: AxUCString; const Value: boolean);
     procedure SetNameAsError(const AName: AxUCString; const Value: TXc12CellError);
     procedure SetNameAsFloat(const AName: AxUCString; const Value: double);
     procedure SetNameAsFormula(const AName: AxUCString; const Value: AxUCString);
     procedure SetNameAsNumFormulaValue(const AName: AxUCString; const Value: double);
     procedure SetNameAsStrFormulaValue(const AName: AxUCString; const Value: AxUCString);
     procedure SetNameAsString(const AName: AxUCString; const Value: AxUCString);
     function  GetBackup: boolean;
     procedure SetBackup(const Value: boolean);
     function  GetRefreshAll: boolean;
     procedure SetRefreshAll(const Value: boolean);
     function  GetUserName: AxUCString;
     procedure SetUserName(const Value: AxUCString);
     function  GetPalette(Index: integer): TColor;
     procedure SetPalette(Index: integer; const Value: TColor);
{$ifdef XLS_BIFF}
     function  GetVBA: TXLSVBA;
{$endif}
     function  GetNames: TXLSNames; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     function  GetInternalNames: TXLSNames; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     function  GetSelectedTab: integer;
     procedure SetSelectedTab(const Value: integer);
     function  GetDocProps: TXc12DocProps;
     function  GetFormulasUncalced: boolean;
     procedure SetFormulasUncalced(const Value: boolean);
protected
     FItems            : TIndexObjectList;

     FClassFactory     : TXLSClassFactory;

     FManager          : TXc12Manager;
     FDefaultName      : AxUCString;
     FFormulas         : TXLSFormulaHandler;

     FCmdFormat        : TXLSCmdFormat;
     FDefaultFormat    : TXLSDefaultFormat;

     FNames            : TXLSNames_Int;
     FOptionsDialog    : TXLSOptionsDialog;
     FWorkbookData     : TXLSWorkbookData;
{$ifdef XLS_BIFF}
     FBIFF             : TBIFF5;
{$endif}
{$ifndef BABOON}
     FXLSClipboard     : TXLSClipboard;
{$endif}
     FDefaultPaperSize : TXc12PaperSize;

     FShowFormulas     : boolean;
     FSelectedTab      : integer;

     FCurrFindSheet    : integer;
     FCompiled         : boolean;

     FLocalizedFormulas: boolean;

     function GetItems(Index: integer): TXLSWorksheet;

     procedure AfterReadAdd(AIndex: integer);
     // ASheetIndex = Index of first changed sheet.
     // ACount? Number of sheets, positive if added, negative if deleted.
     procedure AdjustSheetIndex(const ASheetIndex, ACount: integer);
     function  GetUniqueSheetname: AxUCString;

     function  AddImage97(AAnchor: TCT_TwoCellAnchor; const ASheetIndex: integer): boolean;  virtual;

     procedure CopyColumnHeaders(const ASrcSheet,ACol1,ACol2,ADestSheet,ADestCol: integer);
     procedure GetSheetExtent(const ASheet: integer; out ACol1,Arow1,ACol2,ARow2: integer);
     procedure CreateMembers;
     procedure AdjustColumnsFormulas1d(const AActiveSheet, ACol, ADeltaCols: integer);
     procedure AdjustRowsFormulas1d(const AActiveSheet, ARow, ADeltaRows: integer);
public
     destructor Destroy; override;

     function  Count: integer;

     procedure Delete(const AIndex: integer; ACount: integer = 1);
     procedure Insert(const AIndex: integer; ACount: integer = 1);

     procedure Clear;
     procedure ClearCells; {$ifdef DELPHI_2007_OR_LATER}  deprecated 'Use Clear'; {$endif}

     function  Add: TXLSWorksheet;

     //* Searches for a worksheet by it's name.
     //* ~param Name The name of the worksheet.
     //* ~result The TSheet object. If not found, Nil is returned.
     function  SheetByName(const Name: AxUCString): TXLSWorksheet;

     procedure AfterRead;
     procedure BeforeWrite;

     procedure CalcDimensions;

     procedure CompileFormulas;
     procedure Calculate;

{$ifdef _AXOLOT_DEBUG}
     procedure CalculateAndVerify;
     procedure CheckIntegrity(const AList: TStrings);
     function  CountFormulas: integer;
{$endif}

      //* Returns a password that can be used to unprotect worksheets.
      //* As the password is calculated, this may not be the same password used to protect it.
      //* ~result A valid password that can unprotect the file. If there is no protection, an empty string is returned.
     function  GetWeakPassword: AXUCString;

      //* The max number of rows in a worksheet. 65536 for Excel 97 and 1048576 for Excel 2007.
     function  MaxRowCount: integer;
      //* The max number of columns in a worksheet. 256 for Excel 97 and 16384 for Excel 2007.
     function  MaxColCount: integer;

     function  CopyCells(const ASrcSheet: integer; ASrcAreas: TBaseCellAreas; const ADestSheet,ADestCol,ADestRow: integer; const ANoFormat: boolean = False): boolean; overload;
     procedure CopyCells(SrcSheet,Col1,Row1,Col2,Row2,DestSheet,DestCol,DestRow: integer; const CopyOptions: TCopyCellsOptions = [ccoAdjustCells] + CopyAllCells; const ANoFormat: boolean = False); overload;
     procedure CopyCells(ASrcXLS: TXLSWorkbook; const ASrcSheet,ACol1,ARow1,ACol2,ARow2,ADestSheet,ADestCol,ADestRow: integer; const ANoFormat: boolean = False); overload;
     procedure MoveCells(const SrcSheet,Col1,Row1,Col2,Row2,DestSheet,DestCol,DestRow: integer; const CopyOptions: TCopyCellsOptions = [ccoAdjustCells] + CopyAllCells);
     procedure DeleteCells(const ASheet: integer; AAreas: TBaseCellAreas; AOptions: TXLSDeleteOptions = DefaultDeleteOptions); overload;
     procedure DeleteCells(const ASheet: integer; AArea: TCellArea; AOptions: TXLSDeleteOptions = DefaultDeleteOptions); overload;
     procedure DeleteCells(const ASheet, ACol1, ARow1, ACol2, ARow2: integer; AOptions: TXLSDeleteOptions = DefaultDeleteOptions); overload;

     procedure CopyColumns(const ASrcSheet,ACol1,ACol2,ADestSheet,ADestCol: integer);
     procedure MoveColumns(const ASrcSheet,ACol1,ACol2,ADestSheet,ADestCol: integer);
     procedure DeleteColumns(const ASheet,ACol1,ACol2: integer);
     procedure ClearColumns(const ASheet,ACol1,ACol2: integer);
     procedure InsertColumns(const ASheet,ACol,AColCount: integer);


     procedure CopyRows(const ASrcSheet,ARow1,ARow2,ADestSheet,ADestRow: integer);
     procedure MoveRows(const ASrcSheet,ARow1,ARow2,ADestSheet,ADestRow: integer);
     procedure DeleteRows(const ASheet,ARow1,ARow2: integer);
     procedure ClearRows(const ASheet,ARow1,ARow2: integer);
     procedure InsertRows(const ASheet,ARow,ARowCount: integer);

     procedure CopySheet(const ASrcSheet,ADestSheet: integer); overload;
     procedure CopySheet(ASrcXLS: TXLSWorkbook; const ASrcSheet,ADestSheet: integer); overload;

     procedure BeginFindText;
     function  FindText(const Text: AxUCString; const CaseInsensitive: boolean = True): boolean;
     procedure GetFindData(var Sheet,Col,Row,TextPos: integer; var Text: AxUCString);

     //* The Backup property specifies whether Microsoft Excel should save
     //* backup versions of a file, when the file is opened with Excel.
     property Backup: boolean read GetBackup write SetBackup;
     //* Set the RefreshAll property to True if all external data should be
     //* refreshed when the workbook is loaded by Excel.
     property RefreshAll: boolean read GetRefreshAll write SetRefreshAll;
     //* The name of the creator of the workbook.
     property UserName: AxUCString read GetUserName write SetUserName;

     property NameIsSimpeName       [const AName: AxUCString]: boolean read GetNameIsSimpeName;
     property NameAsFloat           [const AName: AxUCString]: double read GetNameAsFloat write SetNameAsFloat;
     property NameAsString          [const AName: AxUCString]: AxUCString read GetNameAsString write SetNameAsString;
     property NameAsFmtString       [const AName: AxUCString]: AxUCString read GetNameAsFmtString;
     property NameAsBoolean         [const AName: AxUCString]: boolean read GetNameAsBoolean write SetNameAsBoolean;
     property NameAsError           [const AName: AxUCString]: TXc12CellError read GetNameAsError write SetNameAsError;
     property NameAsFormula         [const AName: AxUCString]: AxUCString read GetNameAsFormula write SetNameAsFormula;
     property NameAsNumFormulaValue [const AName: AxUCString]: double read GetNameAsNumFormulaValue write SetNameAsNumFormulaValue;
     property NameAsStrFormulaValue [const AName: AxUCString]: AxUCString read GetNameAsStrFormulaValue write SetNameAsStrFormulaValue;
     property NameAsBoolFormulaValue[const AName: AxUCString]: boolean read GetNameAsBoolFormulaValue write SetNameAsBoolFormulaValue;

     property DefaultName: AxUCString read FDefaultName write SetDefaultName;
     property ShowFormulas: boolean  read FShowFormulas write FShowFormulas;

     property OptionsDialog: TXLSOptionsDialog read FOptionsDialog;
     property WorkbookData: TXLSWorkbookData read FWorkbookData;

     property Formulas: TXLSFormulaHandler read FFormulas;
     property Names: TXLSNames read GetNames;
     // For v4 compatibility.
     property InternalNames: TXLSNames read GetInternalNames;
     //* The palette for the colors used by Excel. The palette has 64 entries,
     //* but the first 8 are fixed and can not be changed.
     property Palette[Index: integer]: TColor read GetPalette write SetPalette;

{$ifdef XLS_BIFF}
     property BIFF: TBIFF5 read FBIFF;
     property VBA: TXLSVBA read GetVBA;
{$endif}
{$ifndef BABOON}
     property Clipboard: TXLSClipboard read FXLSClipboard;
{$endif}
     property CmdFormat: TXLSCmdFormat read FCmdFormat;
     property DefaultFormat: TXLSDefaultFormat read FDefaultFormat write FDefaultFormat;

     property DocProps: TXc12DocProps read GetDocProps;

     property Font: TXc12Font read GetFont;
     property SelectedTab: integer read GetSelectedTab write SetSelectedTab;
     property DefaultPaperSize: TXc12PaperSize read FDefaultPaperSize write FDefaultPaperSize;
     // Tag formulas as uncalced. Excel will calculate formulas when the file is loaded.
     property FormulasUncalced: boolean read GetFormulasUncalced write SetFormulasUncalced;
     property LocalizedFormulas: boolean read FLocalizedFormulas write FLocalizedFormulas;
     property Items[Index: integer]: TXLSWorksheet read GetItems; default;
     property Sheets[Index: integer]: TXLSWorksheet read GetItems;
     end;

implementation

{ TXLSWorksheet }

function TXLSWorksheet.CreateRelativeCells: TXLSRelCells;
begin
  Result := TXLSRelCellsImpl.Create(Self);
end;

function TXLSWorksheet.CreateRelativeCells(ACol, ARow: integer): TXLSRelCells;
begin
  Result := TXLSRelCellsImpl.Create(Self);
  Result.Col1 := ACol;
  Result.Row1 := ARow;
  Result.Col2 := ACol;
  Result.Row2 := ARow;
end;

procedure TXLSWorksheet.AdjustCell(ACol, ARow, ADestCol, ADestRow: integer; ALockStartRow: boolean);
var
  Sz  : integer;
  Cell: TXLSCellItem;
  Ptgs: PXLSPtgs;
begin
  if FCells.FindCell(ADestCol,ADestRow,Cell) and (Cell.Data <> Nil) and (FCells.CellType(@Cell) in XLSCellTypeFormulas) then begin
    if not FCells.IsFormulaCompiled(@Cell) then begin
      CompileFormula(ADestCol,ADestRow,@Cell);
      // Cell is not valid after the cell is compiled.
      Cell := FCells.FindCell(ADestCol,ADestRow);
    end;
    Sz := FCells.GetFormulaPtgs(@Cell,Ptgs);
    if Sz > 0 then
      FOwner.FFormulas.AdjustCell(Ptgs,Sz,ADestCol - ACol,ADestRow - ARow,ALockStartRow);
  end;
end;

procedure TXLSWorksheet.AdjustColumnsFormulas(const ACol, ADeltaCols: integer);
var
  i: integer;
  Cell: PXLSCellItem;
  C: TXLSCellItem;
  Ptgs: PXLSPtgs;
  Sz: integer;
  Name: TXc12DefinedName;
begin
  if ADeltaCols <> 0 then begin
    FCells.BeginIterate;
    while FCells.IterateNext do begin
      Cell := FCells.IterCell;

      if FCells.CellType(Cell) in XLSCellTypeFormulas then begin
        if not FCells.IsFormulaCompiled(Cell) then begin
          CompileFormula(FCells.IterCellCol,FCells.IterCellRow,Cell);
          // Cell is not valid after the cell is compiled.
          C := FCells.FindCell(FCells.IterCellCol,FCells.IterCellRow);
          Cell := @C;
        end;
        Sz := FCells.GetFormulaPtgs(Cell,Ptgs);
        FOwner.FFormulas.AdjustCellColsChanged(Ptgs,Sz,Index,ACol,ADeltaCols);
      end;
    end;
  end;

  FOwner.AdjustColumnsFormulas1d(Index,ACol,ADeltaCols);

  for i := 0 to FOwner.Names.Count - 1 do begin
    Name := TXc12DefinedName(FOwner.Names[i]);
    FOwner.FFormulas.AdjustCellColsChanged(Name.Ptgs,Name.PtgsSz,Index,ACol,ADeltaCols);
    TXLSName(Name).Update;
  end;
end;

procedure TXLSWorksheet.AdjustRowsFormulas(const ARow, ADeltaRows: integer);
var
  i: integer;
  Cell: PXLSCellItem;
  C: TXLSCellItem;
  Ptgs: PXLSPtgs;
  Sz: integer;
  Name: TXc12DefinedName;
begin
  FCells.BeginIterate;
//  FCells.BeginIterate(Max(ARow + ADeltaRows,0));
  while FCells.IterateNext do begin
    Cell := FCells.IterCell;

    if FCells.CellType(Cell) in XLSCellTypeFormulas then begin
      if not FCells.IsFormulaCompiled(Cell) then begin
        CompileFormula(FCells.IterCellCol,FCells.IterCellRow,Cell);
        // Cell is not valid after the cell is compiled.
        C := FCells.FindCell(FCells.IterCellCol,FCells.IterCellRow);
        Cell := @C;
      end;
      Sz := FCells.GetFormulaPtgs(Cell,Ptgs);
      FOwner.FFormulas.AdjustCellRowsChanged(Ptgs,Sz,Index,ARow,ADeltaRows);
    end;
  end;

  FOwner.AdjustRowsFormulas1d(Index,ARow,ADeltaRows);

  for i := 0 to FOwner.Names.Count - 1 do begin
    Name := TXc12DefinedName(FOwner.Names[i]);
    FOwner.FFormulas.AdjustCellRowsChanged(Name.Ptgs,Name.PtgsSz,Index,ARow,ADeltaRows);
    TXLSName(Name).Update;
  end;
end;

procedure TXLSWorksheet.AfterRead;
begin
  FDrawing.AfterRead;
end;

procedure TXLSWorksheet.ApplyAutofilter;
var
  R,C: integer;
  FC: TXLSAutofilterColumn;
begin
  if FAutofilter.FilterColumns.Count < 1 then
    Exit;

  for R := FAutofilter.Row1 to FAutofilter.Row2 do begin
    if Rows[R].Assigned then
      Rows[R].Hidden := False;

    for C := FAutofilter.Col1 to FAutofilter.Col2 do begin
      FC := TXLSAutofilterColumn(FAutofilter.FilterColumns.Find(C));
      if FC <> Nil then begin
        case CellType[C,R] of
          xctBooleanFormula,
          xctBoolean       : Rows[R].Hidden := not FC.Apply(GetAsBoolean(C,R));
          xctStringFormula,
          xctString        : Rows[R].Hidden := not FC.Apply(GetAsString(C,R));
          xctFloatFormula,
          xctFloat         : Rows[R].Hidden := not FC.Apply(GetAsFloat(C,R));
        end;
      end;
    end;
  end;
end;

procedure TXLSWorksheet.Assign(ASource: TXLSWorksheet);
var
  i   : integer;
  S   : AxUCString;
  r,c : integer;
  iXF : integer;
  XF  : TXc12XF;
  Row : PXLSMMURowHeader;
begin
  Clear;

  ASource.CalcDimensions;

  if FOwner <> ASource.FOwner then begin
    FOwner.Names.Assign(ASource.FOwner.Names);
  end;

  r := -1;

  ASource.MMUCells.BeginIterate;
  while ASource.MMUCells.IterateNext do begin
    c := ASource.MMUCells.IterCellCol;
    if ASource.MMUCells.IterCellRow <> r then begin
      r := ASource.MMUCells.IterCellRow;

      Row := FCells.AddRow(r,0);
      Row^ := ASource.MMUCells.IterRow^;
    end;

    iXF := ASource.MMUCells.IterGetStyleIndex;
    if (iXF <> XLS_STYLE_DEFAULT_XF) and (ASource.FOwner <> FOwner) then begin
      XF := FManager.StyleSheet.XFs.Find(ASource.FOwner.FManager.StyleSheet.XFs[iXF]);

      if XF = Nil then
        XF := FManager.StyleSheet.XFs.AddExtern(ASource.FOwner.FManager.StyleSheet.XFs[iXF]);

      iXF := XF.Index;
    end;

    case ASource.MMUCells.IterCellType of
      xctNone          : ;
      xctBlank         : FCells.StoreBlank(c,r,iXF);
      xctBoolean       : FCells.StoreBoolean(c,r,iXF,ASource.MMUCells.IterGetBoolean);
      xctError         : FCells.StoreError(c,r,iXF,ASource.MMUCells.IterGetError);
      xctString        : begin
        i := ASource.MMUCells.IterGetStringIndex;
        FCells.StoreString(c,r,iXF,ASource.FManager.SST.ItemText[i]);
      end;
      xctFloat         : FCells.StoreFloat(c,r,iXF,ASource.MMUCells.IterGetFloat);
      xctFloatFormula,
      xctStringFormula,
      xctBooleanFormula,
      xctErrorFormula  : begin
        S := ASource.MMUCells.GetFormula(ASource.MMUCells.IterCell);
        case ASource.MMUCells.IterCellType of
          xctFloatFormula  : FCells.AddFormula(c,r,iXF,S,ASource.MMUCells.GetFormulaValFloat(ASource.MMUCells.IterCell));
          xctStringFormula : FCells.AddFormula(c,r,iXF,S,ASource.MMUCells.GetFormulaValString(ASource.MMUCells.IterCell));
          xctBooleanFormula: FCells.AddFormula(c,r,iXF,S,ASource.MMUCells.GetFormulaValBoolean(ASource.MMUCells.IterCell));
          xctErrorFormula  : FCells.AddFormula(c,r,iXF,S,ASource.MMUCells.GetFormulaValError(ASource.MMUCells.IterCell));
        end;
      end;
    end;
  end;

  FColumns.Assign(ASource.Columns);
  FMergedCells.Assign(ASource.MergedCells);
  FHyperlinks.Assign(ASource.Hyperlinks);
  FComments.Assign(ASource.FComments);
  FValidations.Assign(ASource.Validations);
  FCondFormats.Assign(ASource.FCondFormats);
  FAutofilter.Assign(ASource.FAutofilter);
end;

function TXLSWorksheet.AutoHeightRow(const ARow: integer; AMinHeight: integer = 0; AMaxHeight: integer = MAXINT): boolean;
{$ifdef BABOON}
begin
  Result := False;
end;
{$else}
var
  i        : integer;
  n        : integer;
  S        : AxUCString;
  H        : integer;
  MaxH     : integer;
  TM       : TEXTMETRIC;
  CurrFont : TXc12Font;
  Canvas   : TCanvas;
  XF       : TXc12XF;
  Runs     : PXc12FontRunArray;
  Cnt      : integer;
  RichPaint: TXBookRichPainter;
begin
  Result := False;
  Canvas := TCanvas.Create;
  try
    Canvas.Handle := GetDC(GetDesktopWindow());
    CurrFont := FManager.StyleSheet.Fonts.DefaultFont;
    CurrFont.CopyToTFont(Canvas.Font);
    GetTextMetrics(Canvas.Handle,TM);
    MaxH := 0;
    FCells.BeginIterate(ARow);
    while FCells.IterateNextCol do begin
      if FCells.IterCell.Data <> Nil then begin
        XF := FManager.StyleSheet.XFs[FCells.IterGetStyleIndex];
        H := 0;
        n := 1;

        if (FCells.CellType(FCells.IterCell) = xctString) then begin
          i := FCells.GetStringSST(FCells.IterCell);
          if SSTStringIsRichStr(FManager.SST[i]) then begin
            FCells.GetString(FCells.IterCell,Runs,Cnt);

            RichPaint := TXBookRichPainter.Create(Canvas);
            try
              RichPaint.AssignText(FManager.SST.ItemText[i],XF,Runs,Cnt);

              H := RichPaint.CalcHeight(FColumns[FCells.IterCellCol].PixelWidth);
            finally
              RichPaint.Free;
            end;
          end
          else if XF.Alignment.IsWrapText then begin
            S := FManager.SST.ItemText[i];
            for i := 1 to Length(S) do begin
              if S[i] = #13 then
                Inc(n);
            end;
          end;
        end;

        if H = 0 then begin
          if XF.Font <> CurrFont then begin
            CurrFont := XF.Font;
            CurrFont.CopyToTFont(Canvas.Font);
            GetTextMetrics(Canvas.Handle,TM);
          end;
          // Vertical
          if XF.Alignment.Rotation in [90,180,255] then
            H := Canvas.TextWidth(Self.GetAsFmtString(FCells.IterCellCol,ARow)) + (TM.tmInternalLeading * 2)
          else
            H := (TM.tmHeight + 2) * n;
        end;

        if H > MaxH then
          MaxH := H;
      end;
    end;
    if MaxH > 0 then begin
      H := Round((MaxH / (Canvas.Font.PixelsPerInch / 72)) * 20);
      if (FCells.IterRow = Nil) or (H <> FCells.IterRow.Height) then begin
        FManager.StyleSheet.Fonts.DefaultFont.CopyToTFont(Canvas.Font);
        FRows[ARow].Height := Fork(H,AMinHeight,AMaxHeight);
        Result := True;
      end;
    end;
  finally
    ReleaseDC(GetDesktopWindow(),Canvas.Handle);
    Canvas.Free;
  end;
end;
{$endif}

function TXLSWorksheet.AutoHeightRows(const Row1, Row2: integer; AMinHeight: integer = 0; AMaxHeight: integer = MAXINT): boolean;
var
  i: integer;
begin
  Result := False;
  for i := Row1 to Row2 do begin
    if AutoHeightRow(i,AMinHeight,AMaxHeight) then
      Result := True;
  end;
end;

procedure TXLSWorksheet.AutoSizeCell(const ACol, ARow: integer; const ASizeCol, ASizeRow: boolean);
{$ifdef BABOON}
begin

end;
{$else}
var
  C: TXLSCellItem;
  W,H: integer;
  TM: TEXTMETRIC;
  Canvas: TCanvas;
  XF: TXc12XF;
begin
  C := FCells.FindCell(ACol,ARow);
  if (C.Data = Nil) or (FCells.CellType(@C) = xctBlank) then
    Exit;
  Canvas := TCanvas.Create;
  try
    Canvas.Handle := GetDC(GetDesktopWindow());
    XF := FManager.StyleSheet.XFs[FCells.GetStyle(@C)];
    XF.Font.CopyToTFont(Canvas.Font);
    GetTextMetrics(Canvas.Handle,TM);
    if ASizeCol then begin
      if XF.Alignment.Rotation in [90,180,255] then
        W := Canvas.TextHeight(self.GetAsFmtString(ACol,ARow)) + (TM.tmInternalLeading * 2)
      else
        W := Canvas.TextWidth(GetAsFmtString(ACol,ARow)) + (TM.tmInternalLeading * 2);
      if XF.Alignment.HorizAlignment = chaRight then
        Inc(W,TM.tmInternalLeading);
      if W > 0 then
        FColumns[ACol].Width := Round(((W) / Canvas.TextWidth('0')) * 256);
    end;
    if ASizeRow then begin
      if XF.Alignment.Rotation in [90,180,255] then
        H := Canvas.TextWidth(self.GetAsFmtString(ACol,ARow)) + (TM.tmInternalLeading * 2)
      else
        H := TM.tmHeight + 1;
      if H > 0 then
        FRows[ARow].Height := Round((H / (Canvas.Font.PixelsPerInch / 72)) * 20);
    end;
  finally
    ReleaseDC(GetDesktopWindow(),Canvas.Handle);
    Canvas.Free;
  end;
end;
{$endif}

function TXLSWorksheet.AutoWidthCol(const ACol: integer; AMinWidth: integer = 0; AMaxWidth: integer = MAXINT): integer;
{$ifdef BABOON}
begin
  Result := 0;
end;
{$else}
var
  W: integer;
  tw: integer;
  r: TRect;
  S: AxUCString;
  TM: TEXTMETRIC;
  CurrFont: TXc12Font;
  Canvas: TCanvas;
  C: TXLSCellItem;
  XF: TXc12XF;
begin
  Canvas := TCanvas.Create;
  try
    Canvas.Handle := GetDC(GetDesktopWindow());
    CurrFont := FManager.StyleSheet.Fonts.DefaultFont;
    CurrFont.CopyToTFont(Canvas.Font);
    GetTextMetrics(Canvas.Handle,TM);
    Result := 0; // DefaultColWidth * Canvas.TextWidth('0');
    FCells.BeginIterateRow;
    while FCells.IterateNextRow do begin
      if FCells.FindCell(ACol,FCells.IterCellRow,C) and (FMergedCells.CellInAreas(ACol,FCells.IterCellRow) <= -1) then begin
        XF := FManager.StyleSheet.XFs[FCells.GetStyle(@C)];
        if Xf.Font <> CurrFont then begin
          CurrFont := XF.Font;
          CurrFont.CopyToTFont(Canvas.Font);
          GetTextMetrics(Canvas.Handle,TM);
        end;

        // *** User suggestion ***
        S := GetAsFmtString(ACol,FCells.IterCellRow);
        if XF.Alignment.IsWrapText then
        begin
          ZeroMemory(@r, SizeOf(r));
          DrawTextEx(Canvas.Handle, PChar(S), -1, r, DT_CALCRECT, nil);
          if XF.Alignment.Rotation in [90,180,255] then // Vertical
            tw := r.Bottom - r.Top
          else
            tw := r.Right - r.Left;
        end
        else
        begin
          if XF.Alignment.Rotation in [90,180,255] then // Vertical
            tw := Canvas.TextHeight(S)
          else
            tw := Canvas.TextWidth(S);
        end;
        W := tw + TM.tmAveCharWidth + 5;
        // *** End user suggestion ***

//        // Vertical
//        if XF.Alignment.Rotation in [90,180,255] then
//          W := Canvas.TextHeight(GetAsFmtString(ACol,FCells.IterCellRow)) + TM.tmAveCharWidth + 5
//        else begin
//          S := GetAsFmtString(ACol,FCells.IterCellRow);
//          W := Canvas.TextWidth(S) + TM.tmAveCharWidth + 5;
//        end;

        if XF.Alignment.Indent > 0 then
          Inc(W,TM.tmAveCharWidth * XF.Alignment.Indent);

        if XF.Alignment.HorizAlignment = chaRight then
          Inc(W,TM.tmAveCharWidth);
        if W > Result then
          Result := W;
      end;
    end;
    if Result > 0 then
      FColumns[ACol].PixelWidth := Fork(Result,AMinWidth,AMaxWidth); //Round(((Result) / FManager.StyleSheet.StdFontWidth) * 256);
  finally
    ReleaseDC(GetDesktopWindow(),Canvas.Handle);
    Canvas.Free;
  end;
end;
{$endif}

procedure TXLSWorksheet.AutoWidthCols(const Col1, Col2: integer; AMinWidth: integer = 0; AMaxWidth: integer = MAXINT);
var
  i: integer;
begin
  for i := Col1 to Col2 do
    AutoWidthCol(i,AMinWidth,AMaxWidth);
end;

procedure TXLSWorksheet.BeginFindText;
begin
  CalcDimensions;
  FCurrFindPos := -1;
  FCurrFindCol := 0;
  FCurrFindRow := 0;
end;

procedure TXLSWorksheet.BeginIterate;
begin
  FCells.BeginIterate;
end;

procedure TXLSWorksheet.CalcDimensions;
begin
  FCells.CalcDimensions;
  FXc12Sheet.Dimension := FCells.Dimension;
end;

procedure TXLSWorksheet.CalcDimensionsEx;
var
  FoundCell: boolean;
  Area: TXLSCellArea;
begin
  Area.Col1 := MAXINT;
  Area.Col2 := -(MAXINT - 1);
  Area.Row1 := MAXINT;
  Area.Row2 := -(MAXINT - 1);

  FCells.BeginIterateRow;
  while FCells.IterateNextRow do begin
    FoundCell := False;
    FCells.BeginIterateCol;
    while FCells.IterateNextCol do begin
      if FCells.IterCellType <> xctBlank then begin
        if FCells.IterCellCol < Area.Col1 then
          Area.Col1 := FCells.IterCellCol;
        if FCells.IterCellCol > Area.Col2 then
          Area.Col2 := FCells.IterCellCol;
        FoundCell := True;
      end;
    end;
    if FoundCell then begin
      if FCells.IterCellRow < Area.Row1 then
        Area.Row1 := FCells.IterCellRow;
      if FCells.IterCellRow > Area.Row2 then
        Area.Row2 := FCells.IterCellRow;
    end;
  end;
  Xc12Sheet.Dimension := Area;
end;

function TXLSWorksheet.Calculate(const ACol, ARow: integer): Variant;
begin
  Result := CalculateAsString(ACol, ARow);
end;

function TXLSWorksheet.CalculateAsString(const ACol, ARow: integer): AxUCString;
var
  Cell: TXLSCellItem;
  S: AxUCString;
  Ptgs: PXLSPtgs;
  Sz: integer;
begin
  if not FOwner.FCompiled then
    FOwner.CompileFormulas;

  if FCells.FindCell(ACol,ARow,Cell) and (FCells.CellType(@Cell) in XLSCellTypeFormulas) then begin
    if not FCells.IsFormulaCompiled(@Cell) then begin
      S := FCells.GetFormula(@Cell);
      Sz := FOwner.FFormulas.EncodeFormula(S,Ptgs,FIndex);
      if Sz > 0 then begin
        FCells.SetFormulaPtgs(ACol,ARow,@Cell,Ptgs,Sz);
        FreeMem(Ptgs);
      end;
    end;
    Sz := FCells.GetFormulaPtgs(@Cell,Ptgs);
    if (Sz > 0) and (Ptgs.Id = xptgArrayFmlaChild) then
      Result := CalculateAsString(PXLSPtgsArrayChildFmla(Ptgs).ParentCol,PXLSPtgsArrayChildFmla(Ptgs).ParentRow)
    else if (Sz > 0) and (Ptgs.Id = xptgDataTableFmlaChild) then
      Result := CalculateAsString(PXLSPtgsDataTableChildFmla(Ptgs).ParentCol,PXLSPtgsDataTableChildFmla(Ptgs).ParentRow)
    else
      Result := FOwner.FFormulas.EvaluateFormulaStr(FCells,FIndex,ACol,ARow,Ptgs,Sz);
  end;
end;

procedure TXLSWorksheet.Calculated(const IsCalculated: boolean);
begin
  // Not used anymore.
end;

function TXLSWorksheet.CalculateEx(Col, Row: integer; CalculateOptions: boolean): Variant;
begin
  Result := Calculate(Col,Row);
end;

procedure TXLSWorksheet.CalculateSheet;
begin
  FOwner.Calculate;
end;

procedure TXLSWorksheet.ClearCells(const ACol1, ARow1, ACol2, ARow2: integer);
begin
  DeleteCells(ACol1, ARow1, ACol2, ARow2,[xdoCellValues]);
end;

procedure TXLSWorksheet.Clear;
begin
  inherited;
  FTabColor := clDefault;
  FXc12Sheet.Clear;
  if FOwnsSheetView then
    FXc12SheetView.Clear
  else begin
    if FXc12SheetView <> Nil then
      FXc12SheetView.Free;
    FXc12SheetView := TXc12SheetView.Create;
  end;
  FColumns.Clear;
  FRows.Clear;
  FValidations.Clear;
  FPrintSettings.Clear;
  FPane.Clear;
  FProtection.Clear;
  FMergedCells.Clear;
  FHyperlinks.Clear;
  FCondFormats.Clear;
  FAutofilter.Clear;
  FComments.Clear;
  FSelectedAreas.Clear;
  FVisibleAreas.Clear;
  FDrawing.Clear;
  FXc12Sheet.Dimension := SetCellArea(0,0,0,0);
end;

procedure TXLSWorksheet.ClearCell(const ACol, ARow: integer);
begin
  DeleteCell(ACol,ARow,[xdoCellValues]);
end;

procedure TXLSWorksheet.ClearColumns(const ACol1, ACol2: integer);
begin
  if ACol1 > ACol2 then
    Exit;
  DeleteCells(ACol1,0,ACol2,XLS_MAXROW);
  FColumns.ClearColumns(ACol1,ACol2);
end;

procedure TXLSWorksheet.ClearData;
begin
  CalcDimensions;

  DeleteCells(FirstCol,FirstRow,LastCol,LastRow);

  FMergedCells.Clear;
  FValidations.Clear;
  FHyperlinks.Clear;
  FCondFormats.Clear;
  FAutofilter.Clear;
  FComments.Clear;

  FRows.Clear;
  FColumns.Clear;

  CalcDimensions;
end;

procedure TXLSWorksheet.ClearRows(const ARow1, ARow2: integer);
begin
  DeleteCells(0,ARow1,XLS_MAXCOL,ARow2);
  FCells.ClearRowHeaders(ARow1,ARow2);
end;

procedure TXLSWorksheet.ClearWorksheet;
begin
  Clear;
end;

procedure TXLSWorksheet.ColWidthChanged(ASender: TObject; const ACol1,ACol2: integer);
begin
  FComments.ColWidthChanged(ACol1,ACol2);
end;

function TXLSWorksheet.CompareCells(const ACell1, ACell2: PXLSCellItem; const ACaseSencitive: boolean): integer;
const
  CellPri: array[TXLSCellType] of integer = (
  5,    //xctNone
  4,    //xctBlank
  2,    //xctBoolean
  3,    //xctError
  1,    //xctString
  0,    //xctFloat
  0,    //xctFloatFormula
  1,    //xctStringFormula
  2,    //xctBooleanFormula
  3);   //xctErrorFormula
var
  T1,T2: TXLSCellType;
begin
  if ACell1.Data = Nil then T1 := xctNone else T1 := FCells.CellType(ACell1);
  if ACell2.Data = Nil then T2 := xctNone else T2 := FCells.CellType(ACell2);

  if T1 = T2 then begin
    case T1 of
      xctNone,
      xctBlank         : Result := 0;
      xctBoolean,
      xctBooleanFormula: Result := Integer(FCells.GetBoolean(ACell1)) - Integer(FCells.GetBoolean(ACell1));
      xctError,
      xctErrorFormula  : Result := 0;
      xctString,
      xctStringFormula : begin
        if ACaseSencitive then
          Result := CompareStr(FCells.GetString(ACell1),FCells.GetString(ACell2))
        else
          Result := CompareText(FCells.GetString(ACell1),FCells.GetString(ACell2));
      end;
      xctFloat,
      xctFloatFormula  : begin
        if FCells.GetFloat(ACell1) > FCells.GetFloat(ACell2) then
          Result := 1
        else if FCells.GetFloat(ACell1) < FCells.GetFloat(ACell2) then
          Result := -1
        else
          Result := 0;
      end;
      else Result := 0;
    end;
  end
  else
    Result := CellPri[T1] - CellPri[T2];
end;

function TXLSWorksheet.CompareRowCells(const ACell1, ACell2: PXLSCellItem; const ARow1, ARow2, ACol1, ACol2: integer; const AAscending,ACaseSencitive: boolean): integer;
var
  Col: integer;
  C1,C2: TXLSCellItem;
begin
  Result := CompareCells(ACell1,ACell2,ACaseSencitive);
  if Result = 0 then begin
    for Col := ACol1 to ACol2 do begin
      C1 := FCells.FindCell(Col,ARow1);
      C2 := FCells.FindCell(Col,ARow2);
      Result := CompareCells(@C1,@C2,ACaseSencitive);
      if Result <> 0 then
        Break;
    end;
  end;
  if not AAscending then
    Result := -Result;
end;

procedure TXLSWorksheet.CompileFormula(const ACol,ARow: integer; ACell: PXLSCellItem);
var
  S: AxUCString;
  Ptgs: PXLSPtgs;
  Sz: integer;
begin
  if (FCells.CellType(ACell) in XLSCellTypeFormulas) and not FCells.IsFormulaCompiled(ACell) then begin
    S := FCells.GetFormula(ACell);

    Sz := FOwner.FFormulas.EncodeFormula(S,Ptgs,FIndex);
    if Sz > 0 then begin
      FCells.SetFormulaPtgs(ACol,ARow,ACell,Ptgs,Sz);
      FreeMem(Ptgs);
    end;
  end;
end;

procedure TXLSWorksheet.CompileFormulas;
var
  S: AxUCString;
  Ptgs: PXLSPtgs;
  Sz: integer;
begin
  if FXc12Sheet.Tables.Count > 0 then begin
    FOwner.FFormulas.CurrTables := FXc12Sheet.Tables;
    FCells.BeginIterate;
    while FCells.IterateNext do begin
      if FCells.IterCellType in XLSCellTypeFormulas then begin
        if not FCells.IsFormulaCompiled(FCells.IterCell) then begin
          S := FCells.GetFormula(FCells.IterCell);
          Sz := FOwner.FFormulas.EncodeFormula(S,Ptgs,FIndex,FCells.IterCellCol,FCells.IterCellRow);
          if Sz > 0 then begin
            FCells.SetFormulaPtgs(FCells.IterCellCol,FCells.IterCellRow,FCells.IterCell,Ptgs,Sz);
            FreeMem(Ptgs);
          end;
        end;
      end;
    end;
  end
  else begin
    FOwner.FFormulas.CurrTables := Nil;
    FCells.BeginIterate;
    while FCells.IterateNext do begin
      if FCells.IterCellType in XLSCellTypeFormulas then begin
        if not FCells.IsFormulaCompiled(FCells.IterCell) then begin
          S := FCells.GetFormula(FCells.IterCell);

          Sz := FOwner.FFormulas.EncodeFormula(S,Ptgs,FIndex);
          if Sz > 0 then begin
            FCells.SetFormulaPtgs(FCells.IterCellCol,FCells.IterCellRow,FCells.IterCell,Ptgs,Sz);
            FreeMem(Ptgs);
          end;
        end;
      end;
    end;
  end;
{$ifdef _AXOLOT_DEBUG}
//  FCells.CheckIntegrity(Nil);
{$endif}
end;

procedure TXLSWorksheet.CopyCell(const ACol, ARow, ADestCol, ADestRow: integer; ALockStartRow: boolean; AClearFmla: boolean = False);
var
  Ptgs: PXLSPtgs;
  Sz: integer;
  Cell,NewCell: TXLSCellItem;
begin
  if InsideExtent(ADestCol,ADestRow) then begin
    Cell := FCells.CopyCell(ACol,ARow);
    if Cell.Data <> Nil then begin
      try
        FCells.InsertCell(ADestCol,ADestRow,@Cell);

        if AClearFmla then
          FCells.ClearFormulaKeepValue(ADestCol,ADestRow)
        else begin
          NewCell := FCells.FindCell(ADestCol,ADestRow);

          if FCells.CellType(@NewCell) in XLSCellTypeFormulas then begin
            if not FCells.IsFormulaCompiled(@NewCell) then
              CompileFormula(ADestCol,ADestRow,@NewCell);
            Sz := FCells.GetFormulaPtgs(@NewCell,Ptgs);
            FOwner.FFormulas.AdjustCell(Ptgs,Sz,ADestCol - ACol,ADestRow - ARow,ALockStartRow);
          end;
        end;
      finally
        FCells.FreeCell(@Cell);
      end;
    end;
  end;
end;

procedure TXLSWorksheet.CopyCells(const ACol1, ARow1, ACol2, ARow2, ADestCol, ADestRow: integer; const ACopyOptions: TCopyCellsOptions);
var
  C,R: integer;
  RowsBackward: boolean;
  ColsBackward: boolean;
begin
  if ccoCopyValues in ACopyOptions then begin
    ColsBackward := ADestCol > ACol1;
    RowsBackward := ADestRow > ARow1;

    if ColsBackward and RowsBackward then begin
      for R := ARow2 downto ARow1 do begin
        for C := ACol2 downto ACol1 do
          CopyCell(C,R,ADestCol + (C - ACol1),ADestRow + (R - ARow1),ccoLockStartRow in ACopyOptions,ccoCopyClearFormulas in ACopyOptions);
      end;
    end
    else if ColsBackward and not RowsBackward then begin
      for R := ARow1 to ARow2 do begin
        for C := ACol2 downto ACol1 do
          CopyCell(C,R,ADestCol + (C - ACol1),ADestRow + (R - ARow1),ccoLockStartRow in ACopyOptions,ccoCopyClearFormulas in ACopyOptions);
      end;
    end
    else if not ColsBackward and RowsBackward then begin
      for R := ARow2 downto ARow1 do begin
        for C := ACol1 to ACol2 do
          CopyCell(C,R,ADestCol + (C - ACol1),ADestRow + (R - ARow1),ccoLockStartRow in ACopyOptions,ccoCopyClearFormulas in ACopyOptions);
      end;
    end
    else begin
      for R := ARow1 to ARow2 do begin
        for C := ACol1 to ACol2 do
          CopyCell(C,R,ADestCol + (C - ACol1),ADestRow + (R - ARow1),ccoLockStartRow in ACopyOptions,ccoCopyClearFormulas in ACopyOptions);
      end;
    end;
  end;
  if ccoCopyValidations in ACopyOptions then
    FValidations.CopyLocal(ACol1, ARow1, ACol2, ARow2, ADestCol, ADestRow);
  if ccoCopyMerged in ACopyOptions then begin
    FMergedCells.Copy(ACol1, ARow1, ACol2, ARow2, ADestCol - ACol1, ADestRow - ARow1);
    if FMergedCells.CopyFailed then
      FManager.Errors.Warning('',XLSWARN_CAN_NOT_CHANGE_MERGED);
  end;
  FHyperlinks.Copy(ACol1, ARow1, ACol2, ARow2, ADestCol - ACol1, ADestRow - ARow1);
  if ccoCopyCondFmt in ACopyOptions then
    FCondFormats.CopyLocal(ACol1, ARow1, ACol2, ARow2, ADestCol, ADestRow);
  if ccoCopyShapes in ACopyOptions then
    FDrawing.Copy(ACol1, ARow1, ACol2, ARow2, ADestCol, ADestRow);
//  if ccoCopyNotes in ACopyOptions then
//    FComments.

end;

procedure TXLSWorksheet.CopyColumns(const ACol1, ACol2, ADestCol: integer);
begin
  if ACol1 <> ADestCol then begin
    CopyCells(ACol1,0,ACol2,XLS_MAXROW,ADestCol,0);
    FColumns.CopyColumns(ACol1,ACol2,ADestCol);
  end;
end;

procedure TXLSWorksheet.CopyRows(const ARow1, ARow2, ADestRow: integer);
begin
  if ARow1 <> ADestRow then begin
    FCells.CopyRowHeaders(ARow1,ARow2,ADestRow);
    CopyCells(0,ARow1,XLS_MAXCOL,ARow2,0,ADestRow);
  end;
end;

{$ifndef BABOON}
procedure TXLSWorksheet.CopyToRichEdit(const ACol, ARow: integer; var ARichEdit: TRichEdit);
type TRichText = record
   text: AxUCString;
   color: TColor;
   font: TFont;
 end;

var
  i,p: integer;
  FormatCount: integer;
  pXS: PXLSString;
  S,S2: AxUCString;
  Cell: TXLSCellItem;
  RT: TRichText;
  Fnt: TFont;
  XF: TXc12XF;
  Runs: PXc12FontRunArray;

procedure RichAdd(ASrc: TRichText; start: integer; out ADest: TRichEdit);
 begin
  ADest.SelStart := Start;
  ADest.SelLength := Length(ASrc.text);
  ADest.SelText := ASrc.Text;
  if(ASrc.Color <> TColor(nil))then begin
    ADest.SelStart := start;
    ADest.SelLength := Length(ASrc.text);
    ADest.SelAttributes.Color := ASrc.color;
  end;
  if(ASrc.Font <> TFont(nil))then begin
    ADest.SelStart := Start;
    ADest.SelLength := Length(ASrc.Text);
    ADest.SelAttributes.Assign(ASrc.Font);
  end;
  ADest.SelStart := 0;
 end;

begin
  ARichEdit.Text := '';
  if not FCells.FindCell(ACol,ARow,Cell) then
    Exit;
  p := 1;
  Fnt := TFont.Create;
  XF := FManager.StyleSheet.XFs[FCells.GetStyle(@Cell)];
  if FCells.CellType(@Cell) = xctString then begin
    if FManager.SST.IsFormatted[FCells.GetStringSST(@Cell)] then begin
      pXS := FManager.SST[FCells.GetStringSST(@Cell)];
      p := 1;
      S := FManager.SST.GetText(pXS);
      XF.Font.CopyToTFont(Fnt);
      Runs := FManager.SST.GetFontRuns(pXS);
      FormatCount := PXLSStringRich(pXS).FormatCount;
      for i := 0 to FormatCount - 1 do begin
        S2 := Copy(S,p,Runs[i].Index - p + 1);
        if S2 <> '' then begin
          RT.Text := S2;
          RT.Font := Fnt;
          RT.Color := Fnt.Color;
          RichAdd(RT,p - 1,ARichEdit);
          Runs[i].Font.CopyToTFont(Fnt);
        end;
        p := Runs[i].Index + 1;
      end;
      S2 := Copy(S,p,MAXINT);
      if S2 <> '' then begin
        RT.Text := S2;
        RT.Font := Fnt;
        RT.Color := Fnt.Color;
        RichAdd(RT,p - 1,ARichEdit);
      end;
    end
    else begin
      RT.Text := GetAsFmtString(ACol,ARow);
      XF.Font.CopyToTFont(Fnt);
      RT.Font := Fnt;
      RT.Color := Fnt.Color;
      RichAdd(RT, p - 1,ARichEdit);
     end;
  end
  else begin
    RT.Text := GetAsFmtString(ACol,ARow);
    XF.Font.CopyToTFont(Fnt);
    RT.Font := Fnt;
    RT.Color := Fnt.Color;
    RichAdd(RT,p - 1,ARichEdit);
  end;
  Fnt.Free;
end;

{$ifdef _AXOLOT_DEBUG}
function TXLSWorksheet.CountFormulas: integer;
begin
  Result := 0;
  FCells.BeginIterate;
  while FCells.IterateNext do begin
    if FCells.IterCellType in XLSCellTypeFormulas then
      Inc(Result);
  end;
end;

procedure TXLSWorksheet.DumpMem(const AList: TStrings);
begin
//  FCells.CheckIntegrity(AList);
end;

{$endif}

{$endif}

procedure TXLSWorksheet.EditPivotTables;
var
  i       : integer;
  PivTbl  : TXLSPivotTable;
  PTD     : TCT_pivotTableDefinition;
  SrcSheet: TXLSWorksheet;
begin
  FPivotTables.Clear;

  for i := 0 to FXc12Sheet.PivotTables.Count - 1 do begin
    PTD := FXc12Sheet.PivotTables[i];

    SrcSheet := FOwner.SheetByName(PTD.Cache.CacheSource.WorksheetSource.Sheet);
    if SrcSheet <> Nil then begin
      PivTbl := FPivotTables.EditTable(PTD);
      PivTbl.DestRef := CreateRelativeCells(PivTbl.TableDef.Location.Ref);

      PTD.Cache.CacheSource.WorksheetSource.RCells := SrcSheet.CreateRelativeCells(PTD.Cache.CacheSource.WorksheetSource.Ref);
    end;
  end;
end;

constructor TXLSWorksheet.Create(AClassFactory: TXLSClassFactory; AOwner: TXLSWorkbook; AManager: TXc12Manager; AXc12Sheet: TXc12DataWorksheet);
begin
  FClassFactory := AClassFactory;

  FOwner := AOwner;
  FManager := AManager;
  FXc12Sheet := AXc12Sheet;

  FTabColor := TColor(XLSCOLOR_AUTO);

  FDefDecimals := -1;

  FCells := FXc12Sheet.Cells;
  FRows := TXLSRows.Create(FCells,FManager.StyleSheet);
  FColumns := TXLSColumns.Create(FXc12Sheet.Columns,FManager.StyleSheet);
  FColumns.OnColWidthChange := ColWidthChanged;

  FMergedCells := TXLSMergedCells(FXc12Sheet.MergedCells);
  FValidations := TXLSDataValidations(FXc12Sheet.DataValidations);
  FHyperlinks := TXLSHyperlinks(FXc12Sheet.Hyperlinks);
  FCondFormats := TXLSConditionalFormats(FXc12Sheet.ConditionalFormatting);
  FAutofilter := TXLSAutofilter(FXc12Sheet.Autofilter);
  FComments := TXLSComments.Create(FManager,FXc12Sheet.Comments);

  FDrawing := TXLSDrawing.Create(FManager,FXc12Sheet.Drawing.WsDr,FManager.Errors);
  FDrawing.OnAddImage97 := DrawingAddImage97;
  FDrawing.SetProps(FColumns,FRows);

  FPivotTables := TXLSPivotTables.Create(FManager,TXLSRelCellsImpl.Create(Self));

//  FPrintSettings := TXLSPrintSettings.Create(FXc12Sheet.PageSetup,FXc12Sheet.HeaderFooter,FXc12Sheet.PageMargins,FXc12Sheet.ColBreaks,FXc12Sheet.RowBreaks,FXc12Sheet.PrintOptions);
  FPrintSettings := TXLSPrintSettings.Create(FXc12Sheet,FOwner.Names);
  if FPrintSettings.PaperSize = psNone then
    FPrintSettings.PaperSize := FManager.DefaultPaperSz;

  FOwnsSheetView := FXc12Sheet.SheetViews.Count <= 0;
  if FOwnsSheetView then
    FXc12SheetView := TXc12SheetView.Create
  else
    FXc12SheetView := FXc12Sheet.SheetViews[0];

  FPane := TXLSPane.Create(FXc12SheetView.Pane);
  FProtection := TXLSSheetProtection.Create(FXc12Sheet.SheetProtection);
  FSelectedAreas := TXLSSelectedAreas.Create;
  FVisibleAreas := TCellAreas.Create;

  FTheCell := TXLSCell.Create(FManager.StyleSheet,FCells);
  FTheCellFormat := TXLSFormattedCell.Create;

  FRange := TXLSRange.Create(FManager,Self);
end;

function TXLSWorksheet.CreateRelativeCells(ARefSTr: AxUCString): TXLSRelCells;
begin
  Result := CreateRelativeCells;
  Result.RefAbs := ARefStr;
end;

procedure TXLSWorksheet.DeleteCell(const ACol, ARow: integer; const AOptions: TXLSDeleteOptions = DefaultDeleteOptions);
begin
  if (xdoCellValues in AOptions) and (xdoFormats in AOptions) then begin
    FCells.DeleteCell(ACol,ARow);
    FHyperlinks.Delete(ACol, ARow, ACol, ARow);
    FMergedCells.Delete(ACol, ARow, ACol, ARow);
  end
  else if xdoCellValues in AOptions then begin
    FCells.ClearCell(ACol,ARow);
    FHyperlinks.Delete(ACol, ARow, ACol, ARow);
    FMergedCells.Delete(ACol, ARow, ACol, ARow);
  end;

  if xdoComments in AOptions then
    FComments.Delete(ACol, ARow, ACol, ARow);

  if xdoDataValidations in AOptions then
    FValidations.DeleteLocal(ACol, ARow, ACol, ARow);

  if xdoConditionalFormats in AOptions then
    FCondFormats.DeleteLocal(ACol, ARow, ACol, ARow);

  if xdoDrawing in AOptions then
    FDrawing.Delete(ACol, ARow, ACol, ARow);
end;

procedure TXLSWorksheet.DeleteCellObjects(ACol1, ARow1, ACol2, ARow2: integer);
begin
  FValidations.MoveLocal      (ACol1,ARow1,ACol2,ARow2,ACol2 - ACol1 + 1,ARow2 - ARow1 + 1);
  FMergedCells.DeleteAndAdjust(ACol1,ARow1,ACol2,ARow2);
  FHyperlinks.DeleteAndAdjust (ACol1,ARow1,ACol2,ARow2);
  FCondFormats.MoveLocal      (ACol1,ARow1,ACol2,ARow2,ACol2 - ACol1 + 1,ARow2 - ARow1 + 1);
  FAutofilter.Move            (ACol1,ARow1,ACol2,ARow2,ACol2 - ACol1 + 1,ARow2 - ARow1 + 1);
  FDrawing.Move               (ACol1,ARow1,ACol2,ARow2,ACol2 - ACol1 + 1,ARow2 - ARow1 + 1);
end;

procedure TXLSWorksheet.DeleteCells(const ACol1, ARow1, ACol2, ARow2: integer; const AOptions: TXLSDeleteOptions = DefaultDeleteOptions);
var
  Col,Row: integer;
  C1,R1,C2,R2: integer;
begin
  CalcDimensions;

  if (ACol1 = 0) and (ACol2 = XLS_MAXCOL) then
    FCells.DeleteRows(ARow1,ARow2)
  else begin
    C1 := Max(ACol1,FirstCol);
    R1 := Max(ARow1,FirstRow);
    C2 := Min(ACol2,LastCol);
    R2 := Min(ARow2,LastRow);

    for Col := C1 to C2 do begin
      for Row := R1 to R2 do
        if (xdoCellValues in AOptions) and (xdoFormats in AOptions) then
          FCells.DeleteCell(Col,Row)
        else if xdoCellValues in AOptions then
          FCells.ClearCell(Col,Row);
    end;
  end;

  if xdoCellValues in AOptions then begin
    FMergedCells.Delete(ACol1, ARow1, ACol2, ARow2);
    FHyperlinks.Delete(ACol1, ARow1, ACol2, ARow2);
  end;

  if xdoComments in AOptions then
    FComments.Delete(ACol1, ARow1, ACol2, ARow2);

  if xdoDataValidations in AOptions then
    FValidations.DeleteLocal(ACol1, ARow1, ACol2, ARow2);

  if xdoConditionalFormats in AOptions then
    FCondFormats.DeleteLocal(ACol1, ARow1, ACol2, ARow2);

  if xdoDrawing in AOptions then
    FDrawing.Delete(ACol1, ARow1, ACol2, ARow2);
end;

procedure TXLSWorksheet.DeleteColumns(const ACol1, ACol2: integer);
begin
  if ACol1 > ACol2 then
    Exit;
  FColumns.DeleteColumns(ACol1,ACol2);
  FCells.DeleteAndShiftColumns(ACol1,ACol2);

  DeleteCellObjects(ACol1,XLS_MAXROW,ACol2, XLS_MAXROW);

  AdjustColumnsFormulas(ACol1,-(ACol2 - ACol1 + 1));
end;

procedure TXLSWorksheet.DeleteRows(const ARow1, ARow2: integer);
begin
  if ARow1 > ARow2 then
    Exit;
  FCells.DeleteAndShiftRows(ARow1,ARow2);

  DeleteCellObjects(XLS_MAXCOL, ARow1, XLS_MAXCOL,ARow2);

  AdjustRowsFormulas(ARow1,-(ARow2 - ARow1 + 1));
end;

destructor TXLSWorksheet.Destroy;
begin
  FTheCellFormat.Free;
  FTheCell.Free;
  FRange.Free;
  FDrawing.Free;
  FPivotTables.Free;
  FRows.Free;
  FColumns.Free;

// These are owned by FXc12Sheet
//  FMergedCells.Free;
//  FValidations.Free;
//  FHyperlinks.Free;
//  FAutofilter.Free;
//  FCondFormats.Free;

  FComments.Free;

  FPrintSettings.Free;
  FPane.Free;
  FProtection.Free;
  FSelectedAreas.Free;
  FVisibleAreas.Free;

  if FOwnsSheetView then
    FXc12SheetView.Free;
  inherited;
end;

function TXLSWorksheet.DrawingAddImage97(AAnchor: TCT_TwoCellAnchor): boolean;
begin
  Result := FOwner.AddImage97(AAnchor,Index);
end;

procedure TXLSWorksheet.FillRandom(const AArea: AxUCString;const AValue: integer);
var
  C,R,C1,R1,C2,R2: integer;
  AC1,AR1,AC2,AR2: boolean;
begin
  AreaStrToColRow(AArea,C1,R1,C2,R2,AC1,AR1,AC2,AR2);
  for R := R1 to R2 do begin
    for C := C1 to C2 do
      SetAsFloat(C,R,Random(AValue));
  end;
end;

function TXLSWorksheet.FindCell(const ACol, ARow: integer): boolean;
var
  Cell: TXLSCellItem;
begin
  Result := FCells.FindCell(ACol,ARow,Cell);
end;

function TXLSWorksheet.FindText(const AText: AxUCString; CaseInsensitive: boolean): boolean;
var
  i: integer;
  S: AxUCString;
  Text: AxUCString;
begin
  if FCurrFindPos >= 1 then
    Inc(FCurrFindPos);
  if CaseInsensitive then
    Text := Uppercase(AText)
  else
    Text := AText;
  while FCurrFindRow <= FXc12Sheet.Dimension.Row2 do begin
    while FCurrFindCol <= FXc12Sheet.Dimension.Col2 do begin
      S := GetAsString(FCurrFindCol,FCurrFindRow);
      if S <> '' then begin
        if FCurrFindPos >= 1 then
          S := Copy(S,FCurrFindPos,MAXINT);
        if CaseInsensitive then
          S := Uppercase(S);
        i := FCurrFindPos;
        FCurrFindPos := Pos(Text,S);
        Result := FCurrFindPos >= 1;
        if Result then begin
          if i > 0 then
            Inc(FCurrFindPos,i);
          Exit;
        end;
      end;
      FCurrFindPos := -1;
      Inc(FCurrFindCol);
    end;
    FCurrFindCol := FXc12Sheet.Dimension.Col1;
    Inc(FCurrFindRow);
  end;
  Result := False;
end;

function TXLSWorksheet.FirstCol: integer;
begin
  Result := FXc12Sheet.Dimension.Col1;
end;

function TXLSWorksheet.FirstRow: integer;
begin
  Result := FXc12Sheet.Dimension.Row1;
end;

procedure TXLSWorksheet.FreezePanes(const AColumns, ARows: integer);
begin
  UnfreezePanes;

{$ifdef XLS_BIFF}
  FPane.FXc12Pane.Excel97 := FOwner.BIFF <> Nil;
{$endif}
  FPane.SplitColX := AColumns;
  FPane.SplitRowY := ARows;
  FPane.LeftCol := AColumns;
  FPane.TopRow := ARows;
  FPane.ActivePane := apBottomLeft;
  FPane.PaneType := ptFrozen;
end;

function TXLSWorksheet.GetAsArray(Col1, Row1, Col2, Row2, MaxLength: integer): AxUCString;
var
  C,R: integer;
begin
  Result := '{';
  for R := Row1 to Row2 do begin
    for C := Col1 to Col2 do begin
      Result := Result + '"' + GetAsString(C,R) + '"';
      if C < Col2 then
        Result := Result + FManager.ColSeparator;
      if Length(Result) > MaxLength then
        Break;
    end;
    if R < Row2 then
      Result := Result + FManager.RowSeparator;
    if Length(Result) > MaxLength then
      Break;
  end;
  Result := Result + '}';
end;

function TXLSWorksheet.GetAsBlank(ACol, ARow: integer): boolean;
begin
  Result := FCells.GetBlank(ACol,ARow);
end;

function TXLSWorksheet.GetAsBlankRef(const ARef: AxUCString): boolean;
var
  Col,Row: integer;
begin
  RefStrToColRow(ARef,Col,Row);
  Result := GetAsBlank(Col,Row);
end;

function TXLSWorksheet.GetAsBoolean(ACol, ARow: integer): boolean;
var
  Cell: TXLSCellItem;
  S   : AxUCString;
begin
  if FCells.FindCell(ACol,ARow,Cell) then begin
    case FCells.CellType(@Cell) of
      xctBoolean       : Result := FCells.GetBoolean(@Cell);
      xctBooleanFormula: Result := FCells.GetFormulaValBoolean(@Cell);
      xctString        : begin
        S := Uppercase(FCells.GetString(@Cell));
        if S = 'TRUE' then
          Result := True
        else if S = 'FALSE' then
          Result := False
        else
          raise XLSRWException.Create('Cell is missing/is of wrong type');
      end
      else               raise XLSRWException.Create('Cell is missing/is of wrong type');
    end;
  end
  else
    raise XLSRWException.Create('Cell is missing/is of wrong type');
end;

function TXLSWorksheet.GetAsBooleanRef(const ARef: AxUCString): boolean;
var
  Col,Row: integer;
begin
  RefStrToColRow(ARef,Col,Row);
  Result := GetAsBoolean(Col,Row);
end;

function TXLSWorksheet.GetAsBoolFormulaValue(ACol, ARow: integer): boolean;
begin
  Result := GetAsBoolean(ACol,ARow);
end;

function TXLSWorksheet.GetAsBoolFormulaValueRef(const ARef: AxUCString): boolean;
var
  Col,Row: integer;
begin
  RefStrToColRow(ARef,Col,Row);
  Result := GetAsBoolean(Col,Row);
end;

function TXLSWorksheet.GetAsDateTime(ACol, ARow: integer): TDateTime;
begin
  Result := GetAsFloat(ACol,ARow);
end;

function TXLSWorksheet.GetAsDateTimeRef(const ARef: AxUCString): TDateTime;
var
  Col,Row: integer;
begin
  RefStrToColRow(ARef,Col,Row);
  Result := GetAsFloat(Col,Row);
end;

function TXLSWorksheet.GetAsEmpty(ACol, ARow: integer): boolean;
begin
  Result := FCells.GetEmpty(ACol,ARow);
end;

function TXLSWorksheet.GetAsError(ACol, ARow: integer): TXc12CellError;
var
  Cell: TXLSCellItem;
begin
  if FCells.FindCell(ACol,ARow,Cell) then begin
    case FCells.CellType(@Cell) of
      xctError       : Result := FCells.GetError(@Cell);
      xctErrorFormula: Result := FCells.GetFormulaValError(@Cell);
      else               raise XLSRWException.Create('Cell is missing/is of wrong type');
    end;
  end
  else
    raise XLSRWException.Create('Cell is missing/is of wrong type');
end;

function TXLSWorksheet.GetAsErrorRef(const ARef: AxUCString): TXc12CellError;
var
  Col,Row: integer;
begin
  RefStrToColRow(ARef,Col,Row);
  Result := GetAsError(Col,Row);
end;

function TXLSWorksheet.GetAsFloat(ACol, ARow: integer): double;
var
  Cell: TXLSCellItem;
begin
  if FCells.FindCell(ACol,ARow,Cell) then begin
    case FCells.CellType(@Cell) of
//      xctCurrency,
      xctFloat        : Result := FCells.GetFloat(@Cell);
      xctFloatFormula : Result := FCells.GetFormulaValFloat(@Cell);
      xctString       : Result := StrToFloatDef(GetAsString(ACol,ARow),0);
      xctStringFormula: Result := StrToFloatDef(GetAsString(ACol,ARow),0);
      else              Result := 0;
    end;
  end
  else
    Result := 0;
//    raise XLSRWException.Create('Cell is missing/is of wrong type');
end;

function TXLSWorksheet.GetAsFloatRef(const ARef: AxUCString): double;
var
  Col,Row: integer;
begin
  RefStrToColRow(ARef,Col,Row);
  Result := GetAsFloat(Col,Row);
end;

function TXLSWorksheet.GetAsFormula(ACol, ARow: integer): AxUCString;
var
  Sz: integer;
  p: integer;
  S: AxUCString;
  i: integer;
  Ptgs: PXLSPtgs;
  Cell: TXLSCellItem;
begin
  if FCells.FindCell(ACol,ARow,Cell) and (FCells.CellType(@Cell) in XLSCellTypeFormulas) then begin
    if FCells.IsFormulaCompiled(@Cell) then begin
      Sz := FCells.GetFormulaPtgs(@Cell,Ptgs);
      Result := FOwner.Formulas.DecodeFormula(FCells,FIndex,Ptgs,Sz,False);
    end
    else begin
      Result := FCells.GetFormula(@Cell);
      if (Result <> '') and (Result[1] = '[') then begin
        p := CPos(']',Result);
        if p > 1 then begin
          S := Copy(Result,2,p - 2);
          i := StrToIntDef(S,-1);
          if i > 0 then
            Result := '[' + FManager.XLinks.Items[i - 1].ExternalBook.Filename + ']' + Copy(Result,p + 1,MAXINT);
        end;
      end;
    end;
  end
  else
    raise XLSRWException.Create('Cell is missing/is of wrong type');
end;

function TXLSWorksheet.GetAsFormulaRef(const ARef: AxUCString): AxUCString;
var
  Col,Row: integer;
begin
  RefStrToColRow(ARef,Col,Row);
  Result := GetAsFormula(Col,Row);
end;

function TXLSWorksheet.GetAsHTML(ACol, ARow: integer): AxUCString;
var
  FT: TXLSFormattedText;
begin
  if MakeFormattedText(ACol,ARow,FT) then begin
    try
      Result := FT.AsHTML;
    finally
      FT.Free;
    end;
  end
  else
    Result := '';
end;

function TXLSWorksheet.GetAsHTMLRef(const ARef: AxUCString): AxUCString;
var
  Col,Row: integer;
begin
  RefStrToColRow(ARef,Col,Row);
  Result := GetAsHTML(Col,Row);
end;

function TXLSWorksheet.GetAsHyperlink(ACol, ARow: integer): AxUCString;
var
  HLink: TXLSHyperlink;
begin
  HLink := FHyperlinks.Find(ACol,ARow);
  if HLink <> Nil then
    Result := HLink.Address
  else
    Result := '';
end;

function TXLSWorksheet.GetAsInteger(ACol, ARow: integer): integer;
begin
  Result := Round(GetAsFloat(ACol,ARow));
end;

function TXLSWorksheet.GetAsIntegerRef(const ARef: AxUCString): integer;
begin
  Result := Round(GetAsFloatRef(ARef));
end;

function TXLSWorksheet.GetAsNumFormulaValue(ACol, ARow: integer): double;
begin
  Result := GetAsFloat(ACol,ARow);
end;

function TXLSWorksheet.GetAsNumFormulaValueRef(const ARef: AxUCString): double;
begin
  Result := GetAsFloatRef(ARef);
end;

function TXLSWorksheet.GetAsRichText(ACol, ARow: integer): AnsiString;
var
  FT: TXLSFormattedText;
  RTF: TXLSRTFWriter;
begin
  if MakeFormattedText(ACol,ARow,FT) then begin
    try
      RTF := TXLSRTFWriter.Create(FT);
      try
        Result := AnsiString(RTF.RTF);
      finally
        RTF.Free;
      end;
    finally
      FT.Free;
    end;
  end
  else
    Result := '';
end;

function TXLSWorksheet.GetAsRichTextRef(const ARef: AxUCString): AxUCString;
var
  Col,Row: integer;
begin
  RefStrToColRow(ARef,Col,Row);
  Result := AxUCString(GetAsRichText(Col,Row));
end;

function TXLSWorksheet.GetAsSimpleTags(ACol, ARow: integer): AxUCString;
begin
  Result := '';
end;

function TXLSWorksheet.GetAsSimpleTagsRef(const ARef: AxUCString): AxUCString;
begin
  Result := '';
end;

function TXLSWorksheet.GetAsStrFormulaValue(ACol, ARow: integer): AxUCString;
begin
  Result := GetAsString(ACol,ARow);
end;

function TXLSWorksheet.GetAsStrFormulaValueRef(const ARef: AxUCString): AxUCString;
var
  Col,Row: integer;
begin
  RefStrToColRow(ARef,Col,Row);
  Result := GetAsString(Col,Row);
end;

function TXLSWorksheet.GetAsString(ACol, ARow: integer): AxUCString;
var
  Cell: TXLSCellItem;
begin
  if FCells.FindCell(ACol,ARow,Cell) then begin
    case FCells.CellType(@Cell) of
      xctNone,
      xctBlank:               Result := '';
//      xctCurrency,
      xctFloat:               Result := FloatToStr(FCells.GetFloat(@Cell));
      xctString:              Result := FCells.GetString(@Cell);
      xctBoolean:             if FCells.GetBoolean(@Cell) then Result := G_StrTRUE else Result := G_StrFALSE;
      xctError:               Result := Xc12CellErrorNames[FCells.GetError(@Cell)];
      xctFloatFormula:        if FOwner.ShowFormulas then
                                Result := GetAsFormula(ACol,ARow)
                              else
                                Result := FloatToStr(FCells.GetFormulaValFloat(@Cell));
      xctStringFormula:       if FOwner.ShowFormulas then
                                Result := GetAsFormula(ACol,ARow)
                              else
                                Result := FCells.GetFormulaValString(@Cell);
      xctBooleanFormula:      if FOwner.ShowFormulas then
                                Result := GetAsFormula(ACol,ARow)
                              else
                                if FCells.GetFormulaValBoolean(@Cell) then Result := G_StrTRUE else Result := G_StrFALSE;
      xctErrorFormula:        if FOwner.ShowFormulas then
                                Result := GetAsFormula(ACol,ARow)
                              else
                                Result := Xc12CellErrorNames[FCells.GetFormulaValError(@Cell)];
    end;
  end
  else
    Result := '';
end;

function TXLSWorksheet.GetAsStringRef(const ARef: AxUCString): AxUCString;
var
  Col,Row: integer;
begin
  RefStrToColRow(ARef,Col,Row);
  Result := GetAsString(Col,Row);
end;

function TXLSWorksheet.GetAsVariant(ACol, ARow: integer): Variant;
var
  Cell: TXLSCellItem;
begin
  if FCells.FindCell(ACol,ARow,Cell) then begin
    case FCells.CellType(@Cell) of
      xctNone,
      xctBlank:          Result := 0;
//      xctCurrency,
      xctFloat:          Result := FCells.GetFloat(@Cell);
      xctString:         Result := FCells.GetString(@Cell);
      xctBoolean:        Result := FCells.GetBoolean(@Cell);
      xctError:          Result := Xc12CellErrorNames[FCells.GetError(@Cell)];
      xctFloatFormula:   Result := FCells.GetFormulaValFloat(@Cell);
      xctStringFormula:  Result := FCells.GetFormulaValString(@Cell);
      xctBooleanFormula: Result := FCells.GetFormulaValBoolean(@Cell);
      xctErrorFormula:   Result := Xc12CellErrorNames[FCells.GetFormulaValError(@Cell)];
    end;
  end
  else
    Result := 0;
end;

function TXLSWorksheet.GetAsVariantRef(const ARef: AxUCString): Variant;
var
  Col,Row: integer;
begin
  RefStrToColRow(ARef,Col,Row);
  Result := GetAsVariant(Col,Row);
end;

function TXLSWorksheet.GetAutofilter: TXLSAutofilter;
begin
  Result := FAutofilter;
end;

function TXLSWorksheet.GetCell(ACol, ARow: integer): TXLSCell;
begin
  Result := FTheCell.Items[ACol,ARow];
//  if Result = Nil then
//    raise XLSRWException.Create('No cell at ' + ColRowToRefStr(ACol,ARow));
end;

function TXLSWorksheet.GetCellFormat(const ACol, ARow: integer): TXLSFormattedCell;
var
  Cell: TXLSCellItem;
  XF: TXc12XF;
begin
  Cell := FCells.FindCell(ACol,ARow);
  if Cell.Data <> Nil then
    XF := FManager.StyleSheet.XFs[FCells.GetStyle(@Cell)]
  else
    XF := FManager.StyleSheet.XFs.DefaultXF;
  FTheCellFormat.SetXF(XF);
  Result := FTheCellFormat;
end;

function TXLSWorksheet.GetCellType(ACol, ARow: integer): TXLSCellType;
var
  Cell: TXLSCellItem;
begin
  if FCells.FindCell(ACol,ARow,Cell) then
    Result := FCells.CellType(@Cell)
  else
    Result := xctNone;
end;

function TXLSWorksheet.GetColumns: TXLSColumns;
begin
  Result := FColumns;
end;

function TXLSWorksheet.GetConditionalFormats: TXLSConditionalFormats;
begin
  Result := FCondFormats;
end;

function TXLSWorksheet.GetDefaultColWidth: integer;
begin
  Result := Round(FXc12Sheet.SheetFormatPr.DefaultColWidth);
end;

function TXLSWorksheet.GetDefaultFormat(const ACol, ARow: integer): integer;
var
  R: PXLSMMURowHeader;
begin
  if not FColumns[ACol].Style.IsDefault then
    Result := FColumns[ACol].Style.Index
  else begin
    R := FCells.FindRow(ARow);
    if (R <> Nil) and (R.Style <> XLS_STYLE_DEFAULT_XF) then
      Result := R.Style
    else
      Result := XLS_STYLE_DEFAULT_XF;
  end;
end;

function TXLSWorksheet.GetDefaultRowHeight: integer;
begin
  Result := Round(FXc12Sheet.SheetFormatPr.DefaultRowHeight);
  if Result <= 0 then
    Result := XLS_DEFAULT_ROWHEIGHT;
end;

function TXLSWorksheet.GetDrawing: TXLSDrawing;
begin
  Result := FDrawing;
end;

procedure TXLSWorksheet.GetFindData(var ACol, ARow, ATextPos: integer; var AText: AxUCString);
var
  Cell: TXLSCellItem;
begin
  if FCurrFindPos < 1 then
    raise XLSRWException.Create('There is no text found');
  if not FCells.FindCell(FCurrFindCol,FCurrFindRow,Cell) then
    raise XLSRWException.Create('There is no text found');
  ACol := FCurrFindCol;
  ARow := FCurrFindRow;
  ATextPos := FCurrFindPos;
  AText := GetAsString(ACol,ARow);
end;

function TXLSWorksheet.GetFormatString(ACell: TXLSCellItem; AColor: PLongword): AxUCString;
var
  Sz: integer;
  Ptgs: PXLSPtgs;

function FormatCell: AxUCString;
var
  XF: TXc12XF;

function FormatNumber(V: double): AxUCString;

function DoUserFormat: AxUCString;
var
  S: AxUCString;
  Mask: TExcelMask;
begin
  Mask := TExcelMask.Create;
  try
    S := XF.NumFmt.Value;
    if (S <> '') then begin
      Mask.Mask := S;
      Result := Mask.FormatNumber(V);
      if AColor <> Nil then
        AColor^ := Mask.Color;
    end
    else
      Result := FloatToStr(V);
  finally
    Mask.Free;
  end;
end;

function FloatToFracStr(Value: double; MaxDenom: integer): AxUCString;
var
  an,ad: integer;
  bn,bd: integer;
  n,d: integer;
  sp,delta: double;
  sPrefix: AxUCString;
begin
  if Frac(Value) = 0 then begin
    Result := FloatToStr(Value);
    Exit;
  end;
  sPrefix := '';
  if Int(Value) > 0 then begin
    sPrefix := FloatToStr(Int(Value)) + ' ';
    Value := Frac(Value);
  end;
  an := 0;
  ad := 1;
  bn := 1;
  bd := 1;
  d := ad + bd;
  while (d <= MaxDenom) do begin
    sp := 1e-5 * d;
    n := an + bn;
    delta := Value * d - n;
    if delta > sp then begin
      an := n;
      ad := d;
    end
    else if delta < -sp then begin
      bn := n;
      bd := d;
    end
    else begin
      Result := sPrefix + SysUtils.Format('%d/%d',[n,d]);
      Exit;
    end;
    d := ad + bd;
  end;
  if (bd > MaxDenom) or (Abs(Value * ad - an) < Abs(Value * bd - bn)) then
    Result := sPrefix + SysUtils.Format('%d/%d',[an,ad])
  else
    Result := sPrefix + SysUtils.Format('%d/%d',[bn,bd]);
end;

begin
  XF := FManager.StyleSheet.XFs[FCells.GetStyle(@ACell)];
  case XF.NumFmt.Index of
    14: Result := DateToStr(V);
    20: Result := TimeToStr(V);
    22: Result := DateTimeToStr(V);
    else begin
      if XF.NumFmt.Value <> '' then
        Result := DoUserFormat
      else if FDefDecimals >= 0 then
        Result := FloatToStrF(V,ffNumber,15,FDefDecimals)
      else
        Result := FloatToStr(V);
    end;
  end;
end;

begin
  if ACell.Data <> Nil then begin
    case FCells.CellType(@ACell) of
      xctNone: ;
      xctBlank:              Result := '';
//      xctCurrency,
      xctFloat:              Result := FormatNumber(FCells.GetFloat(@ACell));
      xctString:             Result := FCells.GetString(@ACell);
      xctBoolean:            if FCells.GetBoolean(@ACell) then Result := G_StrTRUE else Result := G_StrFALSE;
      xctError:              Result := Xc12CellErrorNames[FCells.GetError(@ACell)];
      xctFloatFormula:       Result := FormatNumber(FCells.GetFormulaValFloat(@ACell));
      xctStringFormula:      Result := FCells.GetFormulaValString(@ACell);
      xctBooleanFormula:     if FCells.GetFormulaValBoolean(@ACell) then Result := G_StrTRUE else Result := G_StrFALSE;
      xctErrorFormula:       Result := Xc12CellErrorNames[FCells.GetFormulaValError(@ACell)];
    end;
  end
  else
    Result := '';
end;

begin
  if AColor <> Nil then
    AColor^ := $FF000000;
  if ACell.Data <> Nil then begin
    if FOwner.ShowFormulas then begin
      case FCells.CellType(@ACell) of
        xctNone,
        xctBlank: Result := '';
//        xctCurrency,
        xctFloat,
        xctString,
        xctBoolean,
        xctError: Result := FormatCell;
        xctFloatFormula,
        xctStringFormula,
        xctBooleanFormula,
        xctErrorFormula: begin
          Sz := FCells.GetFormulaPtgs(@ACell,Ptgs);
          Result := FOwner.FFormulas.DecodeFormula(FCells,FIndex,Ptgs,Sz,False);
        end;
      end;
    end
    else
      Result := FormatCell;
  end
  else
    Result := '';
end;

function TXLSWorksheet.GetFormulaType(ACol, ARow: integer): TXLSCellFormulaType;
var
  Cell: TXLSCellItem;
begin
  if FCells.FindCell(ACol,ARow,Cell) then
    Result := FCells.FormulaType(@Cell)
  else
    Result := xcftNormal;
end;

function TXLSWorksheet.GetAsFmtString(ACol, ARow: integer): AxUCString;
var
  Cell: TXLSCellItem;
begin
  if FCells.FindCell(ACol,ARow,Cell) then
    Result := GetFormatString(Cell,Nil)
  else
    Result := '';
end;

function TXLSWorksheet.GetAsFmtStringRef(const ARef: AxUCString): AxUCString;
var
  Col,Row: integer;
begin
  RefStrToColRow(ARef,Col,Row);
  Result := GetAsFmtString(Col,Row);
end;

function TXLSWorksheet.GetHyperlinks: TXLSHyperlinks;
begin
  Result := FHyperlinks;
end;

function TXLSWorksheet.GetIsDateTime(ACol, ARow: integer): boolean;
var
  S: AxUCString;
  XF: TXc12XF;
  Cell: TXLSCellItem;
  Mask: TExcelMask;
begin
  Result := False;

  if not FCells.FindCell(ACol,ARow,Cell) then
    Exit;

  if FCells.CellType(@Cell) = xctBlank then
    Exit;
  XF := FManager.StyleSheet.XFs[FCells.GetStyle(@Cell)];
  Result := XF.NumFmt.IsDateTime;
  if not Result then begin
    S := XF.NumFmt.Value;
    if S <> '' then begin
      Mask := TExcelMask.Create;
      try
        Mask.Mask := S;
        Result := Mask.IsDateTime;
      finally
        Mask.Free;
      end;
    end;
  end;
end;

function TXLSWorksheet.GetLeftCol: integer;
var
  SW: TXc12SheetView;
begin
  if FXc12Sheet.SheetViews.Count > 0 then begin
    SW := FXc12Sheet.SheetViews[0];
    Result := SW.TopLeftCell.Col1;
  end
  else
    Result := 0;
end;

function TXLSWorksheet.GetMergedCells: TXLSMergedCells;
begin
  Result := FMergedCells;
end;

function TXLSWorksheet.GetName: AxUCString;
begin
  Result := FXc12Sheet.Name;
end;

function TXLSWorksheet.GetOptions: TXLSSheetOptions;
begin
  Result :=[];

  if FXc12SheetView.ShowFormulas then
    Result := Result + [soShowFormulas];
  if FXc12SheetView.ShowGridLines then
    Result := Result + [soGridlines];
  if FXc12SheetView.ShowRowColHeaders then
    Result := Result + [soRowColHeadings];
  if FXc12SheetView.ShowZeros then
    Result := Result + [soShowZeros];
end;

function TXLSWorksheet.GetPane: TXLSPane;
begin
  Result := FPane;
end;

function TXLSWorksheet.GetPrintSettings: TXLSPrintSettings;
begin
  Result := FPrintSettings;
end;

function TXLSWorksheet.GetRecalcFormulas: boolean;
begin
  Result := FXc12Sheet.SheetCalcPr.FullCalcOnLoad;
end;

function TXLSWorksheet.GetRows: TXLSRows;
begin
  Result := FRows;
end;

function TXLSWorksheet.GetSelectedAreas: TXLSSelectedAreas;
begin
  Result := FSelectedAreas;
end;

procedure TXLSWorksheet.GetShapePixelSize(AShape: TXLSDrwTwoPosShape; out AWidth, AHeight: integer);
var
  i: integer;
  v: integer;
begin
  v := FColumns[AShape.Col1].PixelWidth;
  AWidth := v - Round(v * AShape.Col1Offs);
  for i := AShape.Col1 + 1 to AShape.Col2 - 1 do
    Inc(AWidth,FColumns[i].PixelWidth);
  Inc(AWidth,Round(FColumns[AShape.Col2].PixelWidth * AShape.Col2Offs));

  v := FRows[AShape.Col1].PixelHeight;
  AHeight := v - Round(v * AShape.Row1Offs);
  for i := AShape.Row1 + 1 to AShape.Row2 - 1 do
    Inc(AHeight,FRows[i].PixelHeight);
  Inc(AHeight,Round(FRows[AShape.Row2].PixelHeight * AShape.Row2Offs));
end;

function TXLSWorksheet.GetTabColor: longword;
begin
  Result := FTabColor;
end;

function TXLSWorksheet.GetTopRow: integer;
begin
  Result := Fxc12SheetView.TopLeftCell.Row1;
end;

function TXLSWorksheet.GetValidations: TXLSDataValidations;
begin
  Result := FValidations;
end;

function TXLSWorksheet.GetVisibility: TXc12Visibility;
begin
  Result := FXc12Sheet.State;
end;

function TXLSWorksheet.GetVisibleAreas: TCellAreas;
begin
  Result := FVisibleAreas;
end;

function TXLSWorksheet.GetWorkspaceOptions: TWorkspaceOptions;
begin
  Result := [];

  if FXc12Sheet.SheetPr.PageSetupPr.AutoPageBreaks then
    Result := Result + [woShowAutoBreaks];
  if FXc12Sheet.SheetPr.OutlinePr.ApplyStyles then
    Result := Result + [woApplyStyles];
  if FXc12Sheet.SheetPr.OutlinePr.SummaryBelow then
    Result := Result + [woRowSumsBelow];
  if FXc12Sheet.SheetPr.OutlinePr.SummaryRight then
    Result := Result + [woColSumsRight];
  if FXc12Sheet.SheetPr.PageSetupPr.FitToPage then
    Result := Result + [woFitToPage];
  if FXc12Sheet.SheetPr.OutlinePr.ShowOutlineSymbols then
    Result := Result + [woOutlineSymbols];
end;

function TXLSWorksheet.GetZoom: integer;
begin
  if FXc12Sheet.SheetViews.Count > 0 then
    Result := FXc12Sheet.SheetViews[0].ZoomScale
  else
    Result := 100;
end;

function TXLSWorksheet.GetZoomPreview: integer;
begin
  // Can't find property in Xc12
  Result := 100;
end;

procedure TXLSWorksheet.GroupColumns(const ACol1, ACol2: integer; ACollapsed: boolean);
var
  i,Level: integer;
  Col: TXLSColumn;
begin
  for i := ACol1 to ACol2 do begin
    Col := FColumns[i];
    Level := Col.OutlineLevel;
    Inc(Level);
    if Level > 7 then
      raise XLSRWException.Create('Column grouping level more than 7.');
    Col.OutlineLevel := Level;
    Col.CollapsedOutline := ACollapsed;
  end;
end;

procedure TXLSWorksheet.GroupRows(const ARow1, ARow2: integer; ACollapsed: boolean);
var
  i,Level: integer;
  Row: TXLSRow;
begin
  for i := ARow1 to ARow2 do begin
    Row := FRows[i];
    Level := Row.OutlineLevel;
    Inc(Level);
    if Level > 7 then
      raise XLSRWException.Create('Row grouping level more than 7.');
    Row.OutlineLevel := Level;
    Row.CollapsedOutline := ACollapsed;
  end;
end;

procedure TXLSWorksheet.ImportRichText(ACol, ARow: integer; AFilename: AxUCString);
var
  Stream: TFileStream;
begin
  Stream := TFileStream.Create(AFilename,fmOpenRead + fmShareDenyNone);
  try
    ImportRichText(ACol,ARow,Stream);
  finally
    Stream.Free;
  end;
end;

procedure TXLSWorksheet.ImportRichText(ACol, ARow: integer; AStream: TStream);
var
  R  : integer;
  RTF: TAXWReadRTF;
  Doc: TAXWLogDocEditor;
begin
  ClearCell(ACol,ARow);

  Doc := TAXWLogDocEditor.Create;
  try
    RTF := TAXWReadRTF.Create(Doc);
    try
      RTF.LoadFromStream(AStream);

      R := ARow;
      AXWDocToWorksheet(FManager,Doc,Self,FColumns,FRows,ACol,R);

//      AutoHeightRows(ARow,R);
    finally
      RTF.Free;
    end;
  finally
    Doc.Free;
  end;
end;

procedure TXLSWorksheet.InsertColumns(const ACol, AColCount: integer);
var
  i: integer;
  C,R: integer;
begin
  if AColCount > 0 then begin
    FCells.InsertColumns(ACol,AColCount);
    MoveCellObjects(ACol,0,XLS_MAXCOL,XLS_MAXROW,AColCount,0);
    FColumns.InsertColumnsKeepFmt(ACol,AColCount);

    AdjustColumnsFormulas(ACol,AColCount);

    if ACol > 0 then begin
      FCells.CalcDimensions;

      for R := FCells.Dimension.Row1 to FCells.Dimension.Row2 do begin
        i := FCells.GetStyle(ACol - 1,R);
        if i <> XLS_STYLE_DEFAULT_XF then begin
          for C := ACol to ACol + AColCount - 1 do
            FCells.StoreBlank(C,R,i);
        end;
      end;
    end;
  end;
end;

procedure TXLSWorksheet.InsertStringColValues(const ACol, ARow: integer; const AValues: array of AxUCString);
var
  R: integer;
begin
  for R := 0 to High(AValues) do
    SetAsString(ACol,ARow + R,AValues[R]);
end;

procedure TXLSWorksheet.InsertFloatColValues(const ACol, ARow: integer; const AValues: array of double);
var
  R: integer;
begin
  for R := 0 to High(AValues) do
    SetAsFloat(ACol,ARow + R,AValues[R]);
end;

procedure TXLSWorksheet.InsertRows(const ARow, ARowCount: integer);
var
  i: integer;
  R,C: integer;
  Row: PXLSMMURowHeader;
  RowStyle: integer;
begin
  if ARowCount > 0 then begin
    FCells.InsertRows(ARow,ARowCount);

    MoveCellObjects(0,ARow,XLS_MAXCOL,XLS_MAXROW,0,ARowCount);

    AdjustRowsFormulas(ARow,ARowCount);

    if ARow > 0 then begin
      FCells.CalcDimensions;

      Row := FCells.FindRow(ARow - 1);
      if Row <> Nil then begin
        RowStyle := Row.Style;
        for R := ARow to ARow + ARowCount - 1 do
          FCells.AddRow(R,RowStyle);
      end;

      for C := FCells.Dimension.Col1 to FCells.Dimension.Col2 do begin
        i := FCells.GetStyle(C,ARow - 1);
        if i <> XLS_STYLE_DEFAULT_XF then begin
          for R := ARow to ARow + ARowCount - 1 do
            FCells.StoreBlank(C,R,i);
        end;
      end;
    end;
  end;
end;

procedure TXLSWorksheet.InsertRowValues(const ACol, ARow: integer; const AValues: array of Variant);
var
  C: integer;
begin
  for C := 0 to High(AValues) do
    SetAsVariant(ACol + C,ARow,AValues[C]);
end;

procedure TXLSWorksheet.InsertStringRowValues(const ACol, ARow: integer; const AValues: array of AxUCString);
var
  C: integer;
begin
  for C := 0 to High(AValues) do
    SetAsString(ACol + C,ARow,AValues[C]);
end;

procedure TXLSWorksheet.InsertFloatRowValues(const ACol, ARow: integer; const AValues: array of double);
var
  C: integer;
begin
  for C := 0 to High(AValues) do
    SetAsFloat(ACol + C,ARow,AValues[C]);
end;

function TXLSWorksheet.MMUCells: TXLSCellMMU;
begin
  Result := FCells;
end;

function TXLSWorksheet.IsCellLocked(const ACol, ARow: integer): boolean;
var
  i: integer;
  Cell: TXLSCellItem;
  XF: TXc12XF;
begin
  Result := (FProtection.Password <> '') and FProtection.Sheet;
  if not Result then
    Exit;
//  XF := FManager.StyleSheet.XFs.DefaultXF;
  Cell := FCells.FindCell(ACol,ARow);
  if Cell.Data <> Nil then begin
    i := FCells.GetStyle(@Cell);
    XF := FManager.StyleSheet.XFs[i];
  end
  else begin
    XF := FColumns[ACol].Style;
    if XF.IsDefault then
      XF := FRows[ARow].Style;
  end;
  Result := cpLocked in XF.Protection;
end;

function TXLSWorksheet.IsEmpty: boolean;
begin
  Result := FCells.CalcDimensions;
end;

function TXLSWorksheet.IterateNext: boolean;
begin
  Result := FCells.IterateNext;
end;

procedure TXLSWorksheet.IteratePos(out ACol, ARow: integer);
begin
  ACol := FCells.IterCellCol;
  ARow := FCells.IterCellRow;
end;

function TXLSWorksheet.LastCol: integer;
begin
  Result := FXc12Sheet.Dimension.Col2;
end;

function TXLSWorksheet.LastRow: integer;
begin
  Result := FXc12Sheet.Dimension.Row2;
end;

function TXLSWorksheet.MakeFormattedText(const ACol,ARow: integer; var AFormattedText: TXLSFormattedText): boolean;
var
  Cell: TXLSCellItem;
begin
  Result := FCells.FindCell(ACol,ARow,Cell);
  if Result then
    AFormattedText := MakeFormattedText(Cell);
end;

function TXLSWorksheet.MakeFormattedText(const ACell: TXLSCellItem): TXLSFormattedText;
var
  i,j: integer;
  XStr: PXLSString;
  Runs: PXc12FontRunArray;
begin
  Result := TXLSFormattedText.Create;
  try
    i := FCells.GetStringSST(@ACell);
    j := FCells.GetStyle(@ACell);
    if (FCells.CellType(@ACell) = xctString) and FManager.SST.IsFormatted[i] then begin
      XStr := FManager.SST.Items[i];
      Runs := FManager.SST.GetFontRuns(XStr);
      Result.Add(FManager.SST.GetText(XStr),FManager.StyleSheet.XFs[j].Font,Runs,FManager.SST.GetFontRunsCount(XStr));
    end
    else
      Result.Add(GetFormatString(ACell,Nil),FManager.StyleSheet.XFs[j].Font);
  except
    Result.Free;
    Result := Nil;
  end;
end;

procedure TXLSWorksheet.MakeHyperlink(const ACol, ARow: integer; const AAddress, AText, ATooltip: AxUCString);
var
  HLink: TXLSHyperlink;
begin
  SetAsHyperlink(ACol,ARow,AAddress);
  HLink := FHyperlinks.Find(ACol,ARow);
  if HLink <> Nil then
    HLink.ToolTip := ATooltip;
  SetAsString(ACol,ARow,AText);
  Cell[ACol,ARow].FontColor := $0000FF;
  Cell[ACol,ARow].FontUnderline := xulSingle;
end;

procedure TXLSWorksheet.MakePivotTable(ACol1, ARow1, ACol2, ARow2: integer);
var
  Cache: TCT_pivotCacheDefinition;
  Ref  : TXLSRelCells;
begin
  Ref := CreateRelativeCells(ACol1, ARow1, ACol2, ARow2);

  Cache := FManager.Workbook.PivotCaches.Add;
  Cache.AddFields(Ref);
end;

procedure TXLSWorksheet.MergeCells;
var
  i: integer;
begin
  for i := 0 to FSelectedAreas.Count - 1 do begin
    if FSelectedAreas[i].CellCount > 1 then
      MergeCells(FSelectedAreas[i].Col1,FSelectedAreas[i].Row1,FSelectedAreas[i].Col2,FSelectedAreas[i].Row2);
  end;
end;

procedure TXLSWorksheet.MakeHyperlink(const ACol, ARow: integer; const AAddress, AText: AxUCString);
begin
  SetAsHyperlink(ACol,ARow,AAddress);
  SetAsString(ACol,ARow,AText);
  Cell[ACol,ARow].FontColor := $0000FF;
  Cell[ACol,ARow].FontUnderline := xulSingle;
end;

function TXLSWorksheet.MergeCells(Col1, Row1, Col2, Row2: integer): boolean;
var
  C,Cell: TXLSCellItem;
  Area: TCellArea;
  Col,Row: integer;
  CC,CR: integer;
begin
  NormalizeArea(Col1,Row1,Col2,Row2);
  Result := ClipAreaToSheet(Col1,Row1,Col2,Row2);
  if not Result then
    Exit;

  CC := -1;
  CR := -1;

  Area := FMergedCells.Include(Col1,Row1,Col2,Row2);

  // TODO Copy style from the source cells to all blank cells in the merged.
  // Any border around the source cell shall be applied as well.
  Cell.Data := Nil;
  for Row := Area.Row1 to Area.Row2 do begin
    for Col := Area.Col1 to Area.Col2 do begin
      if FCells.FindCell(Col,Row,C) and (Cell.Data = Nil) then begin
        Cell := C;
        CC := Col;
        CR := Row;
      end
      else
        SetAsBlank(Col,Row,True);
    end;
  end;
  if (Cell.Data <> Nil) and ((CC <> Area.Col1) or (CR <> Area.Row1)) then begin
    FCells.MoveCell(CC,CR,Area.Col1,Area.Row1);
    // It is cells that are dependent on the moved cell that shall be adjusted.
//    AdjustCell(Area.Col1,Area.Row1,Cell,CC - Area.Col1,CR - Area.Row1,False,False);
  end;
end;

procedure TXLSWorksheet.MoveCell(const ACol, ARow, ADestCol, ADestRow: integer; AAdjustCell: boolean = False; ALockStartRow: boolean = False);
begin
  if InsideExtent(ADestCol,ADestRow) then begin
    FCells.MoveCell(ACol, ARow, ADestCol, ADestRow);
    if AAdjustCell or ALockStartRow then
      AdjustCell(ACol,ARow,ADestCol,ADestRow,ALockStartRow);
  end;
end;

procedure TXLSWorksheet.MoveCells(const ACol1, ARow1, ACol2, ARow2, ADestCol, ADestRow: integer; const ACopyOptions: TCopyCellsOptions);
var
  C,R: integer;
  RowsBackward: boolean;
  ColsBackward: boolean;
begin
  if ccoCopyValues in ACopyOptions then begin
    ColsBackward := ADestCol > ACol1;
    RowsBackward := ADestRow > ARow1;

    if ColsBackward and RowsBackward then begin
      for R := ARow2 downto ARow1 do begin
        for C := ACol2 downto ACol1 do
          MoveCell(C,R,ADestCol + (C - ACol1),ADestRow + (R - ARow1),ccoAdjustCells in ACopyOptions,ccoLockStartRow in ACopyOptions);
      end;
    end
    else if ColsBackward and not RowsBackward then begin
      for R := ARow1 to ARow2 do begin
        for C := ACol2 downto ACol1 do
          MoveCell(C,R,ADestCol + (C - ACol1),ADestRow + (R - ARow1),ccoAdjustCells in ACopyOptions,ccoLockStartRow in ACopyOptions);
      end;
    end
    else if not ColsBackward and RowsBackward then begin
      for R := ARow2 downto ARow1 do begin
        for C := ACol1 to ACol2 do
          MoveCell(C,R,ADestCol + (C - ACol1),ADestRow + (R - ARow1),ccoAdjustCells in ACopyOptions,ccoLockStartRow in ACopyOptions);
      end;
    end
    else begin
      for R := ARow1 to ARow2 do begin
        for C := ACol1 to ACol2 do
          MoveCell(C,R,ADestCol + (C - ACol1),ADestRow + (R - ARow1),ccoAdjustCells in ACopyOptions,ccoLockStartRow in ACopyOptions);
      end;
    end;
  end;
  if ccoCopyValidations in ACopyOptions then
    FValidations.MoveLocal(ACol1, ARow1, ACol2, ARow2, ADestCol, ADestRow);
  if ccoCopyMerged in ACopyOptions then
    FMergedCells.Move(ACol1, ARow1, ACol2, ARow2, ADestCol - ACol1, ADestRow - ARow1);
  FHyperlinks.Move(ACol1, ARow1, ACol2, ARow2, ADestCol - ACol1, ADestRow - ARow1);
  if ccoCopyCondFmt in ACopyOptions then
    FCondFormats.MoveLocal(ACol1, ARow1, ACol2, ARow2, ADestCol, ADestRow);
  if ccoCopyShapes in ACopyOptions then
    FDrawing.Move(ACol1, ARow1, ACol2, ARow2, ADestCol, ADestRow);
end;

procedure TXLSWorksheet.MoveCellObjects(ACol1, ARow1, ACol2, ARow2, ADeltaCol, ADeltaRow: integer);
begin
  FValidations.MoveLocal     (ACol1,ARow1,ACol2,ARow2,ADeltaCol,ADeltaRow);
  FMergedCells.Move          (ACol1,ARow1,ACol2,ARow2,ADeltaCol,ADeltaRow);
  FHyperlinks.Move           (ACol1,ARow1,ACol2,ARow2,ADeltaCol,ADeltaRow);
  FCondFormats.MoveLocal     (ACol1,ARow1,ACol2,ARow2,ADeltaCol,ADeltaRow);
  FAutofilter.Move           (ACol1,ARow1,ACol2,ARow2,ADeltaCol,ADeltaRow);
  FDrawing.Move              (ACol1,ARow1,ACol2,ARow2,ADeltaCol,ADeltaRow);
end;

procedure TXLSWorksheet.MoveColumns(const ACol1, ACol2, ADestCol: integer);
begin
  if ACol1 <> ADestCol then begin
    MoveCells(ACol1,0,ACol2,XLS_MAXROW,ADestCol,0);
    FColumns.MoveColumns(ACol1,ACol2,ADestCol);
  end;
end;

procedure TXLSWorksheet.MoveRows(const ARow1, ARow2, ADestRow: integer);
begin
  if ARow1 <> ADestRow then begin
    FCells.CopyRowHeaders(ARow1,ARow2,ADestRow);
    FCells.ClearRowHeaders(ARow1,ARow2);
    MoveCells(0,ARow1,XLS_MAXCOL,ARow2,0,ADestRow);
  end;
end;

function TXLSWorksheet.RequestObject(AClass: TClass): TObject;
begin
  if AClass = TXc12DataWorksheet then
    Result := FXc12Sheet
  else if AClass = TCellAreas then
    Result := FSelectedAreas
  else if AClass = TXLSColumns then
    Result := FColumns
  else if AClass = TXLSRows then
    Result := FRows
  else
    Result := Nil;
end;

procedure TXLSWorksheet.RichTextCell(const ACol, ARow: integer; const AText: AxUCString; AFontRuns: TXc12DynFontRunArray);
var
  i: integer;
begin
  i := FManager.SST.AddFormattedString(AText,AFontRuns);
  FCells.StoreString(ACol,ARow,GetDefaultFormat(ACol,ARow),i);
end;

procedure TXLSWorksheet.RichTextLoadFromFile(const ACol, ARow: integer; const Filename: AxUCString; const AllText: boolean);
var
  Stream: TFileStream;
begin
  Stream := TFileStream.Create(Filename,fmOpenRead + fmShareDenyNone);;
  try
    RichTextLoadFromStream(ACol,ARow,Stream,AllText);
  finally
    Stream.Free;
  end;
end;

procedure TXLSWorksheet.RichTextLoadFromStream(const ACol, ARow: integer; Stream: TStream; const AllText: boolean);
var
  FR : TXc12DynFontRunArray;
  RTF: TAXWReadRTF;
  Doc: TAXWLogDocEditor;
begin
  Doc := TAXWLogDocEditor.Create;
  try
    RTF := TAXWReadRTF.Create(Doc);
    try
      RTF.LoadFromStream(Stream);
      if Length(Doc.PlainText) > MAX_EXCEL_STRSZ then
        raise XLSRWException.Create('Rich text string to long.');

      AXWDocToDynFontRuns(FManager,Doc,Self,FColumns,FRows,FR);

      RichTextCell(ACol,ARow,Doc.PlainText,FR);
    finally
      RTF.Free;
    end;
  finally
    Doc.Free;
  end;
end;

procedure TXLSWorksheet.RichTextSaveToFile(const ACol, ARow: integer; const Filename: AxUCString);
var
  Stream: TFileStream;
begin
  Stream := TFileStream.Create(Filename,fmCreate);
  try
    RichTextSaveToStream(ACol,ARow,Stream);
  finally
    Stream.Free;
  end;
end;

procedure TXLSWorksheet.RichTextSaveToStream(const ACol, ARow: integer; Stream: TStream);
var
  C   : TXLSCellItem;
  RTF : TAXWWriteRTF;
  Doc : TAXWLogDocEditor;
  Runs: PXc12FontRunArray;
  Cnt : integer;
begin
  C := FCells.FindCell(ACol,ARow);

  if C.Data = Nil then
    Exit;

  FCells.GetString(@C,Runs,Cnt);

  Doc := TAXWLogDocEditor.Create;
  try
    FontRunsToAXWDoc(FManager,Self,FColumns,FRows,GetAsString(ACol,ARow),Runs,Cnt,Doc);

    RTF := TAXWWriteRTF.Create(Doc);
    try
      RTF.SaveToStream(Stream);
    finally
      RTF.Free;
    end;
  finally
    Doc.Free;
  end;
end;

procedure TXLSWorksheet.RowHeightChanged(ASender: TObject; const ARow: integer);
begin

end;

procedure TXLSWorksheet.SetAsBlank(ACol, ARow: integer; const Value: boolean);
begin
  if FOwner.FDefaultFormat <> Nil then
    FCells.UpdateBlankOF(ACol,ARow,FOwner.FDefaultFormat.XF.Index)
  else
    FCells.UpdateBlankOF(ACol,ARow,GetDefaultFormat(ACol,ARow));
end;

procedure TXLSWorksheet.SetAsBlankRef(const ARef: AxUCString; const Value: boolean);
var
  C,R: integer;
begin
  if RefStrToColRow(ARef,C,R) then
    SetAsBlank(C,R,Value)
  else
    FManager.Errors.Warning('',XLSWARN_INVALIDCELLREF);
end;

procedure TXLSWorksheet.SetAsBoolean(ACol, ARow: integer; const Value: boolean);
begin
  if FOwner.FDefaultFormat <> Nil then
    FCells.UpdateBooleanOF(ACol,ARow,Value,FOwner.FDefaultFormat.XF.Index)
  else
    FCells.UpdateBooleanOF(ACol,ARow,Value,GetDefaultFormat(ACol,ARow));
end;

procedure TXLSWorksheet.SetAsBooleanRef(const ARef: AxUCString; const Value: boolean);
var
  C,R: integer;
begin
  if RefStrToColRow(ARef,C,R) then
    SetAsBoolean(C,R,Value)
  else
    FManager.Errors.Warning('',XLSWARN_INVALIDCELLREF);
end;

procedure TXLSWorksheet.SetAsBoolFormulaValue(ACol, ARow: integer; const Value: boolean);
begin
  FCells.AddFormulaVal(ACol,ARow,Value);
end;

procedure TXLSWorksheet.SetAsBoolFormulaValueRef(const ARef: AxUCString; const Value: boolean);
var
  C,R: integer;
begin
  if RefStrToColRow(ARef,C,R) then
    SetAsBoolFormulaValue(C,R,Value)
  else
    FManager.Errors.Warning('',XLSWARN_INVALIDCELLREF);
end;

procedure TXLSWorksheet.SetAsDateTime(ACol, ARow: integer; const Value: TDateTime);
var
  i: integer;
  XF: TXc12XF;
  Cell: TXLSCellItem;
begin
  if FOwner.FDefaultFormat <> Nil then
    FCells.UpdateFloatOF(ACol,ARow,Value,FOwner.FDefaultFormat.XF.Index)
  else begin
    if (Int(Value) <> 0) and (Frac(Value) <> 0) then
      i := XLS_NUMFMT_STD_DATETIME
    else if Int(Value) <> 0 then
      i := XLS_NUMFMT_STD_DATE
    else
      i := XLS_NUMFMT_STD_TIME;
//    else
//      i := XLS_NUMFMT_STD_DATE;

    Cell := FCells.FindCell(ACol,ARow);
    if Cell.Data <> Nil then
      XF := FManager.StyleSheet.XFs[FCells.GetStyle(@Cell)]
    else
      XF := FManager.StyleSheet.XFs[GetDefaultFormat(ACol,ARow)];
//      XF := FManager.StyleSheet.XFs.DefaultXF;

    if XF.NumFmt.Index <> i then begin
      FOwner.FCmdFormat.BeginEdit(Self);
      FOwner.FCmdFormat.Number.Format := ExcelStandardNumFormats[i];
      FOwner.FCmdFormat.Apply(ACol,ARow);

      Cell := FCells.FindCell(ACol,ARow);
      if Cell.Data <> Nil then
        XF := FManager.StyleSheet.XFs[FCells.GetStyle(@Cell)]
      else
//        XF := FManager.StyleSheet.XFs.DefaultXF;
        XF := FManager.StyleSheet.XFs[GetDefaultFormat(ACol,ARow)];
    end;

    FCells.UpdateFloatOF(ACol,ARow,Value,XF.Index);
  end;
end;

procedure TXLSWorksheet.SetAsDateTimeRef(const ARef: AxUCString; const Value: TDateTime);
var
  C,R: integer;
begin
  if RefStrToColRow(ARef,C,R) then
    SetAsDateTime(C,R,Value)
  else
    FManager.Errors.Warning('',XLSWARN_INVALIDCELLREF);
end;

procedure TXLSWorksheet.SetAsError(ACol, ARow: integer; const Value: TXc12CellError);
begin
  if FOwner.FDefaultFormat <> Nil then
    FCells.UpdateErrorOF(ACol,ARow,Value,FOwner.FDefaultFormat.XF.Index)
  else
    FCells.UpdateErrorOF(ACol,ARow,Value,GetDefaultFormat(ACol,ARow));
end;

procedure TXLSWorksheet.SetAsErrorRef(const ARef: AxUCString; const Value: TXc12CellError);
var
  C,R: integer;
begin
  if RefStrToColRow(ARef,C,R) then
    SetAsError(C,R,Value)
  else
    FManager.Errors.Warning('',XLSWARN_INVALIDCELLREF);
end;

procedure TXLSWorksheet.SetAsFloat(ACol, ARow: integer; AValue: double);
begin
  if FOwner.FDefaultFormat <> Nil then
    FCells.UpdateFloatOF(ACol,ARow,AValue,FOwner.FDefaultFormat.XF.Index)
  else
    FCells.UpdateFloatOF(ACol,ARow,AValue,GetDefaultFormat(ACol,ARow));
end;

procedure TXLSWorksheet.SetAsFloatRef(const ARef: AxUCString; const Value: double);
var
  C,R: integer;
begin
  if RefStrToColRow(ARef,C,R) then
    SetAsFloat(C,R,Value)
  else
    FManager.Errors.Warning('',XLSWARN_INVALIDCELLREF);
end;

procedure TXLSWorksheet.SetAsFormula(ACol, ARow: integer; const Value: AxUCString);
var
  Ptgs: PXLSPtgs;
  Ptgs97: PByteArray;
  Sz: integer;
  iStyle: integer;
begin
  if FOwner.FDefaultFormat <> Nil then
    iStyle := FOwner.FDefaultFormat.XF.Index
  else
    iStyle := GetDefaultFormat(ACol,ARow);
{$ifdef XLS_BIFF}
  if (FOwner.FBIFF <> Nil) and not FOwner.FBIFF.SaveFormulasAs2007 then begin
    GetMem(Ptgs97,1024);
    try
      Sz := FOwner.FBIFF.FormulaHandler.EncodeFormula(Value,0,Ptgs97,1024);
      FCells.AddFormula97(ACol,ARow,Ptgs97,Sz,iStyle);
    finally
      FreeMem(Ptgs97);
    end;
  end
  else begin
{$endif}
    Sz := FOwner.FFormulas.EncodeFormula(Value,Ptgs,FIndex,ACol,ARow,not FOwner.LocalizedFormulas);
    if Sz > 0 then begin
      FCells.AddFormula(ACol,ARow,iStyle,Ptgs,Sz,0);
      FreeMem(Ptgs);
    end;
{$ifdef XLS_BIFF}
  end;
{$endif}
end;

procedure TXLSWorksheet.SetAsFormulaRef(const ARef: AxUCString; const Value: AxUCString);
var
  C,R: integer;
begin
  if RefStrToColRow(ARef,C,R) then
    SetAsFormula(C,R,Value)
  else
    FManager.Errors.Warning('',XLSWARN_INVALIDCELLREF);
end;

procedure TXLSWorksheet.SetAsHyperlink(ACol, ARow: integer; const Value: AxUCString);
var
  HLink: TXLSHyperlink;
begin
  HLink := FHyperlinks.Find(ACol,ARow);
  if HLink = Nil then
    HLink := FHyperlinks.Add;
  if Copy(Lowercase(Value),1,3) = 'www' then
    // Hyperlinks in Excel 2007 must start with http://
    HLink.Address := 'http://' + Value
  else
    HLink.Address := Value;
  HLink.Col1 := ACol;
  HLink.Col2 := ACol;
  HLink.Row1 := ARow;
  HLink.Row2 := ARow;
end;

procedure TXLSWorksheet.SetAsInteger(ACol, ARow: integer; const Value: integer);
begin
  if FOwner.FDefaultFormat <> Nil then
    FCells.UpdateFloatOF(ACol,ARow,Value,FOwner.FDefaultFormat.XF.Index)
  else
    FCells.UpdateFloatOF(ACol,ARow,Value,GetDefaultFormat(ACol,ARow));
end;

procedure TXLSWorksheet.SetAsIntegerRef(const ARef: AxUCString; const Value: integer);
var
  C,R: integer;
begin
  if RefStrToColRow(ARef,C,R) then
    SetAsInteger(C,R,Value)
  else
   FManager.Errors.Warning('',XLSWARN_INVALIDCELLREF);
end;

procedure TXLSWorksheet.SetAsNumFormulaValue(ACol, ARow: integer; const Value: double);
begin
  FCells.AddFormulaVal(ACol,ARow,Value);
end;

procedure TXLSWorksheet.SetAsNumFormulaValueRef(const ARef: AxUCString; const Value: double);
var
  C,R: integer;
begin
  if RefStrToColRow(ARef,C,R) then
    SetAsNumFormulaValue(C,R,Value)
  else
   FManager.Errors.Warning('',XLSWARN_INVALIDCELLREF);
end;

procedure TXLSWorksheet.SetAsRichText(ACol, ARow: integer; const Value: AnsiString);
var
  Stream: TMemoryStream;
  S     : AnsiString;
begin
  Stream := TMemoryStream.Create;
  try
    S := AnsiString(Value);
    Stream.Write(Pointer(S)^,Length(S));
    Stream.Seek(0,soFromBeginning);
    RichTextLoadFromStream(ACol,ARow,Stream);
  finally
    Stream.Free;
  end;
end;

procedure TXLSWorksheet.SetAsRichTextRef(const ARef: AxUCString; const Value: AxUCString);
var
  C,R: integer;
begin
  if RefStrToColRow(ARef,C,R) then
    SetAsRichText(C,R,AnsiString(Value))
  else
   FManager.Errors.Warning('',XLSWARN_INVALIDCELLREF);
end;

procedure TXLSWorksheet.SetAsSimpleTags(ACol, ARow: integer; const Value: AxUCString);
var
  i: integer;
  Builder: TXLSFontRunBuilder;
begin
  Builder := TXLSFontRunBuilder.Create(FManager);
  try
    Builder.FromSimpleTags(Value);

    i := FManager.SST.AddFormattedString(Builder.Text,Builder.FontRuns);
    FCells.StoreString(ACol,ARow,GetDefaultFormat(ACol,ARow),i);
  finally
    Builder.Free;
  end;
end;

procedure TXLSWorksheet.SetAsSimpleTagsRef(const ARef, Value: AxUCString);
var
  C,R: integer;
begin
  if RefStrToColRow(ARef,C,R) then
    SetAsSimpleTags(C,R,Value)
  else
    FManager.Errors.Warning(Value,XLSWARN_INVALIDCELLREF);
end;

procedure TXLSWorksheet.SetAsStrFormulaValue(ACol, ARow: integer; const Value: AxUCString);
begin
  FCells.AddFormulaVal(ACol,ARow,Value);
end;

procedure TXLSWorksheet.SetAsStrFormulaValueRef(const ARef: AxUCString; const Value: AxUCString);
var
  C,R: integer;
begin
  if RefStrToColRow(ARef,C,R) then
    SetAsStrFormulaValue(C,R,Value)
  else
    FManager.Errors.Warning(Value,XLSWARN_INVALIDCELLREF);
end;

procedure TXLSWorksheet.SetAsString(ACol, ARow: integer; const Value: AxUCString);
begin
  if FOwner.FDefaultFormat <> Nil then
    FCells.UpdateStringOF(ACol,ARow,Value,FOwner.FDefaultFormat.XF.Index)
  else
    FCells.UpdateStringOF(ACol,ARow,Value,GetDefaultFormat(ACol,ARow));
end;

procedure TXLSWorksheet.SetAsStringRef(const ARef: AxUCString; const Value: AxUCString);
var
  C,R: integer;
begin
  if RefStrToColRow(ARef,C,R) then
    SetAsString(C,R,Value)
  else
   FManager.Errors.Warning(Value,XLSWARN_INVALIDCELLREF);
end;

procedure TXLSWorksheet.SetAsVariant(ACol, ARow: integer; const Value: Variant);
var
  C,R: integer;
begin
  if VarIsArray(Value) then begin
{$ifdef DELPHI_6_OR_LATER}
    case VarArrayDimCount(Value) of
      1: begin
        for R := VarArrayLowBound(Value,1) to VarArrayHighBound(Value,1) do
          SetAsVariant(ACol,ARow + R - VarArrayLowBound(Value,1),VarArrayGet(Value,[R]));
      end;
      2: begin
        for R := VarArrayLowBound(Value,2) to VarArrayHighBound(Value,2) do begin
          for C := VarArrayLowBound(Value,1) to VarArrayHighBound(Value,1) do
            SetAsVariant(ACol + C - VarArrayLowBound(Value,1),ARow + R - VarArrayLowBound(Value,2),VarArrayGet(Value,[C,R]));
        end;
      end;
      else
         FManager.Errors.Warning('',XLSWARN_INVALIDCELLVALUE);
    end;
{$else}
    raise XLSRWException.Create('AsVariant not supported by Delphi 5');
{$endif}
  end
  else begin
    case VarType(Value) of
      varSmallInt,
      varInteger : SetAsFloat(ACol,ARow,Value);
      varSingle,
      varDouble,
      varCurrency: SetAsFloat(ACol,ARow,Value);
      varDate    : SetAsDateTime(ACol,ARow,Value);
      varOleStr  : SetAsString(ACol,ARow,Value);
      varDispatch,
      varError   : SetAsString(ACol,ARow,Value);
      varBoolean : SetAsBoolean(ACol,ARow,Value);
      varUnknown : SetAsString(ACol,ARow,Value);
{$ifdef DELPHI_6_OR_LATER}
      varShortInt,
      varByte,
      varWord,
       varLongWord: SetAsInteger(ACol,ARow,Value);
      varInt64,
{$endif}
{$ifdef DELPHI_2007_OR_LATER}
      varUInt64  : SetAsFloat(ACol,ARow,Value);
{$endif}
      varString  : SetAsString(ACol,ARow,Value);
      varAny     : SetAsString(ACol,ARow,Value);
      varArray,
{$ifdef DELPHI_2007_OR_LATER}
      varUString,
{$endif}
      varByRef   : SetAsString(ACol,ARow,Value);
      else         SetAsString(ACol,ARow,Value);
    end;
  end;
end;

procedure TXLSWorksheet.SetAsVariantRef(const ARef: AxUCString; const Value: Variant);
var
  C,R: integer;
begin
  if RefStrToColRow(ARef,C,R) then
    SetAsVariant(C,R,Value)
  else
   FManager.Errors.Warning(Value,XLSWARN_INVALIDCELLREF);
end;

procedure TXLSWorksheet.SetDefaultColWidth(const Value: integer);
begin
  FXc12Sheet.SheetFormatPr.DefaultColWidth := Value;
end;

procedure TXLSWorksheet.SetDefaultRowHeight(const Value: integer);
begin
  FXc12Sheet.SheetFormatPr.DefaultRowHeight := Value;
end;

procedure TXLSWorksheet.SetIgnoreErrorNumbersAsText(const Value: boolean);
var
  IgErr: TXc12IgnoredError;
begin
  FXc12Sheet.IgnoredErrors.Clear;

  IgErr := FXc12Sheet.IgnoredErrors.Add;
  IgErr.Sqref.Add(0,0,XLS_MAXCOL,XLS_MAXROW);
  IgErr.NumberStoredAsText := True;
end;

procedure TXLSWorksheet.SetLeftCol(const Value: integer);
var
  Tmp: PXLSCellArea;
begin
  Tmp := @FXc12SheetView.TopLeftCell;
  Tmp.Col1 := Value;
end;

procedure TXLSWorksheet.SetName(const Value: AxUCString);
var
  S: AxUCString;
begin
  S := Copy(Value,1,31);
  if FManager.Worksheets.Find(S) >= 0 then
    FManager.Errors.Warning(S,XLSWARN_USEDSHEETNAME)
  else begin
    FXc12Sheet.Name := S;
{$ifdef XLS_BIFF}
    if FOwner.BIFF <> Nil then
      FOwner.BIFF[FXc12Sheet.Index].Name := S;
{$endif}
  end;
end;

procedure TXLSWorksheet.SetOptions(const Value: TXLSSheetOptions);
begin
  FXc12SheetView.ShowFormulas := soShowFormulas in Value;
  FXc12SheetView.ShowGridLines := soGridlines in Value;
  FXc12SheetView.ShowRowColHeaders := soRowColHeadings in Value;
  FXc12SheetView.ShowZeros := soShowZeros in Value;
end;

procedure TXLSWorksheet.SetRecalcFormulas(const Value: boolean);
begin
  FXc12Sheet.SheetCalcPr.FullCalcOnLoad := Value;
end;

procedure TXLSWorksheet.SetTabColor(const Value: longword);
begin
  FXc12Sheet.SheetPr.TabColor := RGBColorToXc12(Value);
end;

procedure TXLSWorksheet.SetTopRow(const Value: integer);
var
  Tmp: PXLSCellArea;
begin
  Tmp := @FXc12SheetView.TopLeftCell;
  Tmp.Row1 := Value;
end;

procedure TXLSWorksheet.SetVisibility(const Value: TXc12Visibility);
begin
  FXc12Sheet.State := Value;
end;

procedure TXLSWorksheet.SetWorkspaceOptions(const Value: TWorkspaceOptions);
begin
  FXc12Sheet.SheetPr.PageSetupPr.AutoPageBreaks := woShowAutoBreaks in Value;
  FXc12Sheet.SheetPr.OutlinePr.ApplyStyles := woApplyStyles in Value;
  FXc12Sheet.SheetPr.OutlinePr.SummaryBelow := woRowSumsBelow in Value;
  FXc12Sheet.SheetPr.OutlinePr.SummaryRight := woColSumsRight in Value;
  FXc12Sheet.SheetPr.PageSetupPr.FitToPage := woFitToPage in Value;
  FXc12Sheet.SheetPr.OutlinePr.ShowOutlineSymbols := woOutlineSymbols in Value;
end;

procedure TXLSWorksheet.SetZoom(const Value: integer);
begin
  FXc12SheetView.ZoomScale := Value;
end;

procedure TXLSWorksheet.SetZoomPreview(const Value: integer);
begin
  FXc12SheetView.ZoomScalePageLayoutView := Value;
end;

procedure TXLSWorksheet.Sort(const AAscending,ACaseSencitive: boolean);
begin
  if FSelectedAreas.Count = 1 then
    Sort(FSelectedAreas[0].Col1,FSelectedAreas[0].Row1,FSelectedAreas[0].Col2,FSelectedAreas[0].Row2,AAscending,ACaseSencitive)
  else
    FManager.Errors.Warning('Sort',XLSWARN_MULTIPLE_SEL_AREAS);
end;

procedure TXLSWorksheet.Sort(ACol1, ARow1, ACol2, ARow2: integer; const AAscending,ACaseSencitive: boolean);
var
  i,j,k: integer;
  C: integer;
  FirstRow,LastRow: integer;
  Temp: TXLSCellSortItem;
  SortCells: array of TXLSCellSortItem;
  ProgressCount: integer;
begin
  for FirstRow := ARow1 to ARow2 do begin
    if FCells.FindCell(ACol1,FirstRow,Temp.Cell) then
      Break;
  end;

  for LastRow := ARow2 downto ARow1 do begin
    if FCells.FindCell(ACol1,LastRow,Temp.Cell) then
      Break;
  end;

  if FirstRow >= LastRow then
    Exit;

  ProgressCount := (ARow2 - Arow1 + 1) * (ACol2 - ACol1 + 1);
  if ProgressCount > 10000 then
    FManager.BeginProgress(xptSortCells,ProgressCount div 2);

  SetLength(SortCells,LastRow - FirstRow + 1);
  for i := 0 to High(SortCells) do begin
    SortCells[i].Cell := FCells.CopyCell(ACol1,FirstRow + i);
    SortCells[i].Row := FirstRow + i;
  end;

  k := High(SortCells) shr 1;
  while k > 0 do begin
    for i := 0 to High(SortCells) - k do begin
      j := i;
      while (j >= 0) and (CompareRowCells(@SortCells[j],@SortCells[j + k],SortCells[j].Row,SortCells[j + k].Row,ACol1 + 1,ACol2,AAscending,ACaseSencitive) > 0) do begin
        Temp := SortCells[j];
        SortCells[j] := SortCells[j + k];
        SortCells[j + k] := Temp;
        if j > k then
          Dec(j, k)
        else
          j := 0
      end;
    end;
    FManager.WorkProgress(k);
    k := k shr 1;
  end;

  for i := 0 to High(SortCells) do begin
    if SortCells[i].Cell.Data <> Nil then begin
      FCells.InsertCell(ACol1,FirstRow + i,@SortCells[i]);
      System.FreeMem(SortCells[i].Cell.Data);
      SortCells[i].Cell.Data := Nil;
    end
    else
      FCells.DeleteCell(ACol1,FirstRow + i);
  end;

  for C := ACol1 + 1 to ACol2 do begin
    for i := 0 to High(SortCells) do
      SortCells[i].Cell := FCells.CopyCell(C,SortCells[i].Row);

    for i := 0 to High(SortCells) do begin
      if SortCells[i].Cell.Data <> Nil then begin
        FCells.InsertCell(C,FirstRow + i,@SortCells[i]);
        System.FreeMem(SortCells[i].Cell.Data);
        SortCells[i].Cell.Data := Nil;
      end
      else
        FCells.DeleteCell(C,FirstRow + i);
    end;
  end;

  FManager.EndProgress;
end;

procedure TXLSWorksheet.SplitPanes(const APointsWidth, APointsHeight: integer);
begin
  UnsplitPanes;
{$ifdef XLS_BIFF}
  FPane.FXc12Pane.Excel97 := FOwner.BIFF <> Nil;
{$endif}
  FPane.SplitColX := APointsWidth * 20;
  FPane.SplitRowY := APointsHeight * 20;
  FPane.TopRow := 0;
  FPane.LeftCol := 0;
  FPane.ActivePane := apTopLeft;
  FPane.PaneType := ptSplit;
end;

procedure TXLSWorksheet.UnfreezePanes;
begin
  FPane.Clear;
end;

function TXLSWorksheet.UngroupColumns(const ACol1, ACol2: integer): boolean;
var
  i    : integer;
  Level: integer;
begin
  Result := False;
  Level := -1;
  for i := ACol1 to ACol2 do begin
    if FColumns[i] = Nil then
      Exit;
    if Level < 0 then
      Level := FColumns[i].OutlineLevel
    else if FColumns[i].OutlineLevel <> Level then
      Exit;
  end;
  if Level > 0 then begin
    for i := ACol1 to ACol2 do
      FColumns[i].OutlineLevel := FColumns[i].OutlineLevel  - 1;

    Result := True;
  end;
end;

function TXLSWorksheet.UngroupRows(const ARow1, ARow2: integer): boolean;
var
  i    : integer;
  Level: integer;
begin
  Result := False;
  Level := -1;
  for i := ARow1 to ARow2 do begin
    if FRows[i] = Nil then
      Exit;
    if Level < 0 then
      Level := FRows[i].OutlineLevel
    else if FRows[i].OutlineLevel <> Level then
      Exit;
  end;
  if Level > 0 then begin
    for i := ARow1 to ARow2 do
      FRows[i].OutlineLevel := FRows[i].OutlineLevel  - 1;

    Result := True;
  end;
end;

procedure TXLSWorksheet.UnMergeCells;
var
  i: integer;
begin
  for i := 0 to FSelectedAreas.Count - 1 do begin
    if FSelectedAreas[i].CellCount > 1 then
      UnMergeCells(FSelectedAreas[i].Col1,FSelectedAreas[i].Row1,FSelectedAreas[i].Col2,FSelectedAreas[i].Row2);
  end;
end;

procedure TXLSWorksheet.UnMergeCells(Col1, Row1, Col2, Row2: integer);
begin
  FMergedCells.Delete(Col1, Row1, Col2, Row2);
end;

procedure TXLSWorksheet.UnsplitPanes;
begin
  FPane.Clear;
end;

procedure TXLSWorksheet.InsertColValues(const ACol, ARow: integer; const AValues: array of Variant);
var
  R: integer;
begin
  for R := 0 to High(AValues) do
    SetAsVariant(ACol,ARow + R,AValues[R]);
end;

function TXLSWorksheet.CreateRelativeCells(ACol1, ARow1, ACol2, ARow2: integer): TXLSRelCells;
begin
  Result := TXLSRelCellsImpl.Create(Self);
  Result.Col1 := ACol1;
  Result.Row1 := ARow1;
  Result.Col2 := ACol2;
  Result.Row2 := ARow2;
end;

function TXLSWorksheet.CreateRelativeCells(AArea: TXLSCellArea): TXLSRelCells;
begin
  Result := CreateRelativeCells(AArea.Col1,AArea.Row1,AArea.Col2,AArea.Row2);
end;

{ TXLSWorkbook }

function TXLSWorkbook.Add: TXLSWorksheet;
var
  Xc12Sheet: TXc12DataWorksheet;
begin
  Xc12Sheet := FManager.Worksheets.Add(-1);
  Xc12Sheet.Name := GetUniqueSheetname;;
  Result := TXLSWorksheet.Create(FClassFactory,Self,FManager,Xc12Sheet);
  FItems.Add(Result);
{$ifdef XLS_BIFF}
  if FBIFF <> Nil then begin
    FBIFF.Sheets.Add();
    FBIFF.FormulaHandler.InsertSheet(Count - 1);
  end;
{$endif}
end;

function TXLSWorkbook.AddImage97(AAnchor: TCT_TwoCellAnchor; const ASheetIndex: integer): boolean;
begin
  Result := False;
end;

procedure TXLSWorkbook.AdjustColumnsFormulas1d(const AActiveSheet, ACol, ADeltaCols: integer);
var
  i: integer;
  Cell: PXLSCellItem;
  Cells: TXLSCellMMU;
  C: TXLSCellItem;
  Ptgs: PXLSPtgs;
  Sz: integer;
begin
  if ADeltaCols <> 0 then begin
    for i := 0 to Count - 1 do begin
      if i = AActiveSheet then
        Continue;
      Cells := Items[i].MMUCells;
      Cells.BeginIterate;
      while Cells.IterateNext do begin
        Cell := Cells.IterCell;

        if Cells.CellType(Cell) in XLSCellTypeFormulas then begin
          if not Cells.IsFormulaCompiled(Cell) then begin
            Items[i].CompileFormula(Cells.IterCellCol,Cells.IterCellRow,Cell);
            // Cell is not valid after the cell is compiled.
            C := Cells.FindCell(Cells.IterCellCol,Cells.IterCellRow);
            Cell := @C;
          end;
          Sz := Cells.GetFormulaPtgs(Cell,Ptgs);
          FFormulas.AdjustCellColsChanged1d(Ptgs,Sz,AActiveSheet,ACol,ADeltaCols);
        end;
      end;
    end;
  end;
end;

procedure TXLSWorkbook.AdjustRowsFormulas1d(const AActiveSheet, ARow, ADeltaRows: integer);
var
  i: integer;
  Cell: PXLSCellItem;
  Cells: TXLSCellMMU;
  C: TXLSCellItem;
  Ptgs: PXLSPtgs;
  Sz: integer;
begin
  if ADeltaRows <> 0 then begin
    for i := 0 to Count - 1 do begin
      if i = AActiveSheet then
        Continue;
      Cells := Items[i].MMUCells;
      Cells.BeginIterate(Max(ARow + ADeltaRows,0));
      while Cells.IterateNext do begin
        Cell := Cells.IterCell;

        if Cells.CellType(Cell) in XLSCellTypeFormulas then begin
          if not Cells.IsFormulaCompiled(Cell) then begin
            Items[i].CompileFormula(Cells.IterCellCol,Cells.IterCellRow,Cell);
            // Cell is not valid after the cell is compiled.
            C := Cells.FindCell(Cells.IterCellCol,Cells.IterCellRow);
            Cell := @C;
          end;
          Sz := Cells.GetFormulaPtgs(Cell,Ptgs);
          FFormulas.AdjustCellRowsChanged1d(Ptgs,Sz,AActiveSheet,ARow,ADeltaRows);
        end;
      end;
    end;
  end;
end;

procedure TXLSWorkbook.AdjustSheetIndex(const ASheetIndex, ACount: integer);
var
  i: integer;
  Sheet: TXc12DataWorksheet;
  Ptgs: PXLSPtgs;
  PtgsSize: integer;
begin
  for i := 0 to FManager.Worksheets.Count - 1 do begin
    Sheet := FManager.Worksheets[i];

    Sheet.Cells.BeginIterate;
    while Sheet.Cells.IterateNext do begin
      if Sheet.Cells.IterCellType in XLSCellTypeFormulas then begin
        PtgsSize := Sheet.Cells.IterGetFormulaPtgs(Ptgs);
        FFormulas.AdjustSheetIndex(Ptgs,PtgsSize,ASheetIndex,ACount);
      end;
    end;
  end;

  FNames.AdjustSheetIndex(ASheetIndex,ACount);
end;

procedure TXLSWorkbook.AfterRead;
var
  i: integer;
begin
  FCompiled := False;
  for i := 0 to FManager.Worksheets.Count - 1 do
    AfterReadAdd(i);

  for i := 0 to FItems.Count - 1 do
    Items[i].AfterRead;

  FNames.AfterRead;
end;

procedure TXLSWorkbook.CalcDimensions;
var
  i: integer;
begin
  for i := 0 to FItems.Count - 1 do
    Items[i].CalcDimensions;
end;

procedure TXLSWorkbook.Calculate;
begin
  CalcDimensions;
  if not FCompiled then
    CompileFormulas;
  FFormulas.Calculate;
end;

{$ifdef _AXOLOT_DEBUG}
procedure TXLSWorkbook.CalculateAndVerify;
begin
  CalcDimensions;
  if not FCompiled then
    CompileFormulas;
  FFormulas.CalculateAndVerify;
end;

procedure TXLSWorkbook.CheckIntegrity(const AList: TStrings);
//var
//  i: integer;
begin
//  for i := 0 to FItems.Count - 1 do begin
//    if AList <> Nil then
//      AList.Add('======> Sheet: ' + Items[i].Name);
//    Items[i].FCells.CheckIntegrity(AList);
//  end;
end;

function TXLSWorkbook.CountFormulas: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to FItems.Count - 1 do
    Inc(Result,Items[i].CountFormulas);
end;
{$endif}

procedure TXLSWorkbook.Clear;
begin
//  FIsClearing := True;
//  inherited Clear;
//  FIsClearing := False;
  FItems.Clear;
  FOptionsDialog.Clear;
  FCompiled := False;
{$ifdef XLS_BIFF}
  if FBIFF <> Nil then
    FBIFF.Free;
  FBIFF := Nil;
{$endif}
  FCmdFormat.Clear;
end;

procedure TXLSWorkbook.ClearCells;
begin
  Clear;
end;

procedure TXLSWorkbook.ClearColumns(const ASheet, ACol1, ACol2: integer);
begin
  Items[ASheet].ClearColumns(ACol1,ACol2);
end;

procedure TXLSWorkbook.ClearRows(const ASheet, ARow1, ARow2: integer);
begin
  Items[ASheet].ClearRows(ARow1,ARow2);
end;

procedure TXLSWorkbook.CompileFormulas;
var
  i: integer;
begin
  FManager.BeginProgress(xptCompileFormulas,FItems.Count);
  CalcDimensions;
  for i := 0 to FItems.Count - 1 do begin
    FManager.WorkProgress(i + 1);
    Items[i].CompileFormulas;
  end;
  FManager.EndProgress;
  FCompiled := True;
end;

procedure TXLSWorkbook.CopyCells(SrcSheet, Col1, Row1, Col2, Row2, DestSheet, DestCol, DestRow: integer; const CopyOptions: TCopyCellsOptions; const ANoFormat: boolean);
var
  C,R: integer;
  Cell: TXLSCellItem;
  SrcCells, DestCells: TXLSCellMMU;
begin
  if (SrcSheet = DestSheet) and (Col1 = DestCol) and (Row1 = DestRow) then
    Exit;

  if SrcSheet = DestSheet then
    Items[SrcSheet].CopyCells(Col1,Row1,Col2,Row2,DestCol,DestRow,CopyOptions)
  else begin
    ClipAreaToExtent(Col1,Row1,Col2,Row2);
    if ccoCopyValues in CopyOptions then begin
      SrcCells := Items[SrcSheet].FCells;
      DestCells := Items[DestSheet].FCells;
      for R := Row1 to Row2 do begin
        for C := Col1 to Col2 do begin
           Cell := SrcCells.CopyCell(C,R);
           if Cell.Data <> Nil then begin
             DestCells.InsertCell(DestCol + (C - Col1),DestRow + (R - Row1),@Cell);
             if ccoCopyClearFormulas in CopyOptions then
               DestCells.ClearFormulaKeepValue(DestCol + (C - Col1),DestRow + (R - Row1));

             SrcCells.FreeCell(@Cell);
             if ccoAdjustCells in CopyOptions then
               Items[DestSheet].AdjustCell(C,R,DestCol + (C - Col1),DestRow + (R - Row1),False)
             else if ccoLockStartRow in CopyOptions then
               Items[DestSheet].AdjustCell(DestCol + (C - Col1),DestRow + (R - Row1),DestCol,DestRow,True);
           end
           else
             Items[DestSheet].FCells.DeleteCell(DestCol + (C - Col1),DestRow + (R - Row1));
        end;
      end;
    end;
    if ccoCopyMerged in CopyOptions then
      Items[DestSheet].MergedCells.FillAndClipHitList(DestCol, DestRow,DestCol + (Col2 - Col1), DestRow + (Row2 - Row1),Items[SrcSheet].MergedCells);
  end;
end;

function TXLSWorkbook.CopyCells(const ASrcSheet: integer; ASrcAreas: TBaseCellAreas; const ADestSheet,ADestCol,ADestRow: integer; const ANoFormat: boolean = False): boolean;
begin
  Result := ASrcAreas.Count <= 1;
  if not Result then
    FManager.Errors.Warning('Copy',XLSWARN_MULTIPLE_SEL_AREAS)
  else
    CopyCells(ASrcSheet,ASrcAreas[0].Col1,ASrcAreas[0].Row1,ASrcAreas[0].Col2,ASrcAreas[0].Row2,ADestSheet,ADestCol,ADestRow);
end;

procedure TXLSWorkbook.CopyCells(ASrcXLS: TXLSWorkbook; const ASrcSheet, ACol1, ARow1, ACol2, ARow2, ADestSheet, ADestCol, ADestRow: integer; const ANoFormat: boolean = False);
var
  i: integer;
  C,R: integer;
  DestC,DestR: integer;
  Sz: integer;
  Cell: TXLSCellItem;
  SrcCells: TXLSCellMMU;
  DestCells: TXLSCellMMU;
  SrcXF: TXc12XF;
  DestXF: TXc12XF;
  Ptgs: PXLSPtgs;
  SrcRowHdr,DestRowHdr: PXLSMMURowHeader;
  HitList: TCellAreas;
begin
  SrcCells := ASrcXLS[ASrcSheet].MMUCells;
  DestCells := Items[ADestSheet].MMUCells;
  for R := ARow1 to ARow2 do begin
    for C := ACol1 to ACol2 do begin
      if SrcCells.FindCell(C,R,Cell) then begin
        DestXF := Nil;

        DestC := ADestCol + (C - ACol1);
        DestR := ADestRow + (R - ARow1);

        i := SrcCells.GetStyle(@Cell);
        SrcXF := ASrcXLS.FManager.StyleSheet.XFs[i];
        if ANoFormat then
          FManager.StyleSheet.XFEditor.FreeStyle(SrcXF.Index)
        else begin
          DestXF := FManager.StyleSheet.XFs.Find(SrcXF);
          if DestXF = Nil then
            DestXF := FManager.StyleSheet.XFs.CopyAndAdd(SrcXF);
        end;

        if SrcCells.CellType(@Cell) in XLSCellTypeFormulas then begin
          if not SrcCells.IsFormulaCompiled(@Cell) then
            ASrcXLS[ASrcSheet].CompileFormula(C,R,@Cell);
          Sz := SrcCells.GetFormulaPtgs(@Cell,Ptgs);
          DestCells.FormulaHelper.Clear;
          DestCells.FormulaHelper.Col := DestC;
          DestCells.FormulaHelper.Row := DestR;
          DestCells.FormulaHelper.PtgsSize := Sz;
          DestCells.FormulaHelper.Ptgs := Ptgs;
          DestCells.FormulaHelper.FormulaType := xcftNormal;
          DestCells.FormulaHelper.Style := DestXF.Index;
          if ((DestC - C) <> 0) or ((DestR - R) <> 0) then
            FFormulas.AdjustCell(Ptgs,Sz,DestC - C,DestR - R);
        end;

        case SrcCells.CellType(@Cell) of
          xctBlank          : DestCells.StoreBlank(DestC,DestR,DestXF.Index);
          xctBoolean        : DestCells.StoreBoolean(DestC,DestR,DestXF.Index,SrcCells.GetBoolean(@Cell));
          xctError          : DestCells.StoreError(DestC,DestR,DestXF.Index,SrcCells.GetError(@Cell));
          xctString         : DestCells.StoreString(DestC,DestR,DestXF.Index,SrcCells.GetString(@Cell));
          xctFloat          : DestCells.StoreFloat(DestC,DestR,DestXF.Index,SrcCells.GetFloat(@Cell));
          xctFloatFormula   : DestCells.FormulaHelper.AsFloat := SrcCells.GetFloat(@Cell);
          xctStringFormula  : DestCells.FormulaHelper.AsString := SrcCells.GetString(@Cell);
          xctBooleanFormula : DestCells.FormulaHelper.AsBoolean := SrcCells.GetBoolean(@Cell);
          xctErrorFormula   : DestCells.FormulaHelper.AsError := SrcCells.GetError(@Cell);
        end;

        if SrcCells.CellType(@Cell) in XLSCellTypeFormulas then
          DestCells.StoreFormula;

        SrcRowHdr := SrcCells.FindRow(R);
        if SrcRowHdr <> Nil then begin
          DestRowHdr := DestCells.FindRow(DestR);
          if DestRowHdr <> Nil then
            DestRowHdr.Height := SrcRowHdr.Height;
        end;
      end;
    end;
  end;

  if ASrcXLS[ASrcSheet].MergedCells.Count > 0 then begin
    HitList := TCellAreas.Create;
    try
      ASrcXLS[ASrcSheet].MergedCells.FillHitList(ACol1, ARow1, ACol2, ARow2,HitList);
      if HitList.Count > 0 then begin
        HitList.Move(ADestCol - ACol1,ADestRow - ARow1);
        Items[ADestSheet].MergedCells.Assign(HitList);
      end;
    finally
      HitList.Free;
    end;
  end;

end;

procedure TXLSWorkbook.CopyColumnHeaders(const ASrcSheet, ACol1, ACol2, ADestSheet, ADestCol: integer);
var
  Cols: TXc12Columns;
begin
  Cols := Items[ASrcSheet].Columns.CopyHitList(ACol1,ACol2);
  try
    Items[ADestSheet].Columns.OffsetList(Cols,ADestCol - ACol1);
    Items[ADestSheet].Columns.InsertList(Cols);
  finally
    Cols.Free;
  end;
end;

procedure TXLSWorkbook.CopyColumns(const ASrcSheet, ACol1, ACol2, ADestSheet, ADestCol: integer);
var
  C1,R1,C2,R2: integer;
begin
  if ASrcSheet = ADestSheet then
    Items[ASrcSheet].CopyColumns(ACol1,ACol2,ADestCol)
  else begin
    GetSheetExtent(ASrcSheet,C1,R1,C2,R2);

    CopyCells(ASrcSheet,ACol1,0,ACol2,R2,ADestSheet,ADestCol,0);
    CopyColumnHeaders(ASrcSheet,ACol1,ACol2,ADestSheet,ADestCol);
  end;
end;

procedure TXLSWorkbook.CopyRows(const ASrcSheet, ARow1, ARow2, ADestSheet, ADestRow: integer);
var
  C1,R1,C2,R2: integer;
begin
  if ASrcSheet = ADestSheet then
    Items[ASrcSheet].CopyRows(ARow1,ARow2,ADestRow)
  else begin
    GetSheetExtent(ASrcSheet,C1,R1,C2,R2);

    CopyCells(ASrcSheet,0,ARow1,C2,ARow2,ADestSheet,0,ADestRow);
    Items[ADestSheet].FCells.CopyRowHeaders(Items[ASrcSheet].FCells,ARow1,ARow2,ADestRow);
  end;
end;

procedure TXLSWorkbook.CopySheet(ASrcXLS: TXLSWorkbook; const ASrcSheet, ADestSheet: integer);
var
  i: integer;
  SrcXF: TXc12XF;
  DestXF: TXc12XF;
  SrcXc12Sht,DestXc12Sht: TXc12DataWorksheet;
  Col: TXc12Column;
begin
  SrcXc12Sht := ASrcXLS.FManager.Worksheets[ASrcSheet];
  DestXc12Sht := FManager.Worksheets[ADestSheet];

  DestXc12Sht.SheetFormatPr.Assign(SrcXc12Sht.SheetFormatPr);

  ASrcXLS[ASrcSheet].CalcDimensions;
  CopyCells(ASrcXLS,ASrcSheet,0,0,ASrcXLS[ASrcSheet].LastCol,ASrcXLS[ASrcSheet].LastRow,ADestSheet,0,0);

  DestXc12Sht.Columns.Clear;
  for i := 0 to SrcXc12Sht.Columns.Count - 1 do begin
    SrcXF := SrcXc12Sht.Columns[i].Style;
    DestXF := FManager.StyleSheet.XFs.Find(SrcXF);
    if DestXF = Nil then
      DestXF := FManager.StyleSheet.XFs.CopyAndAdd(SrcXF);

    Col := DestXc12Sht.Columns.Add;
    Col.Assign(SrcXc12Sht.Columns[i]);
    Col.Style := DestXF;
  end;

  Items[ADestSheet].MMUCells.CopyRowHeaders(ASrcXLS[ASrcSheet].MMUCells,0,ASrcXLS[ASrcSheet].LastRow,0);
end;

procedure TXLSWorkbook.CopySheet(const ASrcSheet, ADestSheet: integer);
var
  C1,R1,C2,R2: integer;
  Opts: TCopyCellsOptions;
begin
  if ASrcSheet <> ADestSheet then begin
    Opts := CopyAllCells - [ccoCopyMerged];

    GetSheetExtent(ASrcSheet,C1,R1,C2,R2);

    CopyCells(ASrcSheet,C1,R1,C2,R2,ADestSheet,C1,R1,Opts);
    CopyColumnHeaders(ASrcSheet,C1,C2,ADestSheet,C1);
    Items[ADestSheet].FCells.CopyRowHeaders(Items[ASrcSheet].FCells,R1,R2,R1);

    Items[ADestSheet].MergedCells.Assign(Items[ASrcSheet].MergedCells);
    Items[ADestSheet].Validations.Assign(Items[ASrcSheet].Validations);
    Items[ADestSheet].ConditionalFormats.Assign(Items[ASrcSheet].ConditionalFormats);
  end;
end;

procedure TXLSWorkbook.CreateMembers;
{$ifndef BABOON}
var
  PB: PByteArray;
  Sz: integer;
{$endif}
begin
  FItems := TIndexObjectList.Create;

  FLocalizedFormulas := True;

  FDefaultName := 'Sheet';
  FShowFormulas := False;

  FNames := TXLSNames_Int(FManager.Workbook.DefinedNames);
  FOptionsDialog := TXLSOptionsDialog.Create(FManager.Workbook.CalcPr,FManager.Workbook.WorkbookPr);
  FWorkbookData := TXLSWorkbookData.Create(FManager.Workbook.BookViews);

  FCmdFormat := TXLSCmdFormat.Create(FManager);

  FDefaultPaperSize := psA4;
{$ifndef BABOON}
  PB := Nil;
  Sz := GetLocaleInfo(GetUserDefaultLCID,$100A {LOCALE_IPAPERSIZE},PChar(PB),0);
  GetMem(PB,Sz * 2);
  try
    if GetLocaleInfo(LOCALE_SYSTEM_DEFAULT,$100A {LOCALE_IPAPERSIZE},PChar(PB),Sz) <> 0 then begin
      case Char(PB[0]) of
        '1': FDefaultPaperSize := psLegal;
        '5': FDefaultPaperSize := psLetter;
        '8': FDefaultPaperSize := psA3;
        '9': FDefaultPaperSize := psA4;
      end;
    end;
  finally
     FreeMem(PB);
  end;
{$endif}

{$ifndef BABOON}
  FXLSClipboard := TXLSClipboard.Create(FManager);
{$endif}
end;

procedure TXLSWorkbook.Delete(const AIndex: integer; ACount: integer = 1);
var
  i: integer;
begin
  if (AIndex >= 0) and (AIndex < FManager.Worksheets.Count) and (ACount > 0) then begin
    if not FCompiled then
      CompileFormulas;

//    for i := 0 to ACount - 1 do begin
    for i := ACount - 1 downto 0 do begin //wenyue
      FManager.Worksheets.Delete(AIndex + i);
      FItems.Delete(AIndex + i);
{$ifdef XLS_BIFF}
      if FBIFF <> Nil then
        FBIFF.Sheets.Delete(AIndex + i,False);
{$endif}
    end;
    AdjustSheetIndex(AIndex,-ACount);
  end;
end;

procedure TXLSWorkbook.DeleteCells(const ASheet, ACol1, ARow1, ACol2, ARow2: integer; AOptions: TXLSDeleteOptions = DefaultDeleteOptions);
begin
  Items[ASheet].DeleteCells(ACol1,ARow1,ACol2,ARow2,AOptions);
end;

procedure TXLSWorkbook.DeleteCells(const ASheet: integer; AAreas: TBaseCellAreas; AOptions: TXLSDeleteOptions = DefaultDeleteOptions);
var
  i: integer;
begin
  for i := 0 to AAreas.Count - 1 do
    DeleteCells(ASheet,AAreas[i],AOptions);
end;

procedure TXLSWorkbook.DeleteCells(const ASheet: integer; AArea: TCellArea; AOptions: TXLSDeleteOptions = DefaultDeleteOptions);
begin
  Items[ASheet].DeleteCells(AArea.Col1,AArea.Row1,AArea.Col2,AArea.Row2,AOptions);
end;

procedure TXLSWorkbook.DeleteColumns(const ASheet, ACol1, ACol2: integer);
begin
  Items[ASheet].DeleteColumns(ACol1,ACol2);
end;

procedure TXLSWorkbook.DeleteRows(const ASheet, ARow1, ARow2: integer);
begin
  Items[ASheet].DeleteRows(ARow1,ARow2);
end;

destructor TXLSWorkbook.Destroy;
begin
//  FNames.Free;
  FOptionsDialog.Free;
  FWorkbookData.Free;
  FItems.Free;
  FCmdFormat.Free;
{$ifndef BABOON}
  FXLSClipboard.Free;
{$endif}
  inherited;
end;

function TXLSWorkbook.FindText(const Text: AxUCString; const CaseInsensitive: boolean): boolean;
begin
  while FCurrFindSheet < FItems.Count do begin
    Result := Items[FCurrFindSheet].FindText(Text,CaseInsensitive);
    if Result then
      Exit;
    Inc(FCurrFindSheet);
  end;
  Result := False;
end;

function TXLSWorkbook.GetBackup: boolean;
begin
  Result := FManager.Workbook.WorkbookPr.BackupFile;
end;

function TXLSWorkbook.GetDocProps: TXc12DocProps;
begin
  Result := FManager.DocProps;
end;

procedure TXLSWorkbook.GetFindData(var Sheet, Col, Row, TextPos: integer; var Text: AxUCString);
begin
  Items[FCurrFindSheet].GetFindData(Col, Row, TextPos,Text);
  Sheet := FCurrFindSheet;
end;

function TXLSWorkbook.GetFont: TXc12Font;
begin
  Result := FManager.StyleSheet.Fonts[0];
end;

function TXLSWorkbook.GetFormulasUncalced: boolean;
begin
  Result := FFormulas.FormulasUncalced;
end;

function TXLSWorkbook.GetInternalNames: TXLSNames;
begin
  Result := TXLSNames(FNames);
end;

function TXLSWorkbook.GetItems(Index: integer): TXLSWorksheet;
begin
  Result := TXLSWorksheet(FItems[Index]);
end;

function TXLSWorkbook.GetNameAsBoolean(const AName: AxUCString): boolean;
var
  Name: TXLSName;
begin
  Name := FNames.Find(AName);
  if (Name <> Nil) and (Name.SimpleName in [xsntRef,xsntArea]) then
    Result := Items[Name.Area.SheetIndex].AsBoolean[Name.Area.Col1 and not COL_ABSFLAG,Name.Area.Row1 and not ROW_ABSFLAG]
  else
    Result := False;
end;

function TXLSWorkbook.GetNameAsBoolFormulaValue(const AName: AxUCString): boolean;
begin
  Result := GetNameAsBoolean(AName);
end;

function TXLSWorkbook.GetNameAsError(const AName: AxUCString): TXc12CellError;
var
  Name: TXLSName;
begin
  Name := FNames.Find(AName);
  if (Name <> Nil) and (Name.SimpleName in [xsntRef,xsntArea]) then
    Result := Items[Name.Area.SheetIndex].AsError[Name.Area.Col1 and not COL_ABSFLAG,Name.Area.Row1 and not ROW_ABSFLAG]
  else
    Result := errUnknown;
end;

function TXLSWorkbook.GetNameAsFloat(const AName: AxUCString): double;
var
  Name: TXLSName;
begin
  Name := FNames.Find(AName);
  if (Name <> Nil) and (Name.SimpleName in [xsntRef,xsntArea]) then
    Result := Items[Name.Area.SheetIndex].AsFloat[Name.Area.Col1 and not COL_ABSFLAG,Name.Area.Row1 and not ROW_ABSFLAG]
  else
    Result := 0;
end;

function TXLSWorkbook.GetNameAsFmtString(const AName: AxUCString): AxUCString;
var
  Name: TXLSName;
begin
  Name := FNames.Find(AName);
  if (Name <> Nil) and (Name.SimpleName in [xsntRef,xsntArea]) then
    Result := Items[Name.Area.SheetIndex].AsFmtString[Name.Area.Col1 and not COL_ABSFLAG,Name.Area.Row1 and not ROW_ABSFLAG]
  else
    Result := '';
end;

function TXLSWorkbook.GetNameAsFormula(const AName: AxUCString): AxUCString;
var
  Name: TXLSName;
begin
  Name := FNames.Find(AName);
  if (Name <> Nil) and (Name.SimpleName in [xsntRef,xsntArea]) then
    Result := Items[Name.Area.SheetIndex].AsFormula[Name.Area.Col1 and not COL_ABSFLAG,Name.Area.Row1 and not ROW_ABSFLAG]
  else
    Result := '';
end;

function TXLSWorkbook.GetNameAsNumFormulaValue(const AName: AxUCString): double;
begin
  Result := GetNameAsFloat(AName);
end;

function TXLSWorkbook.GetNameAsStrFormulaValue(const AName: AxUCString): AxUCString;
begin
  Result := GetNameAsString(AName);
end;

function TXLSWorkbook.GetNameAsString(const AName: AxUCString): AxUCString;
var
  Name: TXLSName;
begin
  Name := FNames.Find(AName);
  if (Name <> Nil) and (Name.SimpleName in [xsntRef,xsntArea]) then
    Result := Items[Name.Area.SheetIndex].AsString[Name.Area.Col1 and not COL_ABSFLAG,Name.Area.Row1 and not ROW_ABSFLAG]
  else
    Result := '';
end;

function TXLSWorkbook.GetNames: TXLSNames;
begin
  Result := TXLSNames(FNames);
end;

function TXLSWorkbook.GetPalette(Index: integer): TColor;
begin
  Result := TColor(Xc12IndexColorPalette[Index]);
end;

function TXLSWorkbook.GetRefreshAll: boolean;
begin
  Result := FManager.Workbook.WorkbookPr.RefreshAllConnections;
end;

function TXLSWorkbook.GetSelectedTab: integer;
begin
  Result := FWorkbookData.SelectedTab;
end;

procedure TXLSWorkbook.GetSheetExtent(const ASheet: integer; out ACol1, Arow1, ACol2, ARow2: integer);
begin
  Items[ASheet].CalcDimensions;
  ACol1 := Items[ASheet].FirstCol;
  ARow1 := Items[ASheet].FirstRow;
  ACol2 := Items[ASheet].LastCol;
  ARow2 := Items[ASheet].LastRow;
end;

function TXLSWorkbook.GetUniqueSheetname: AxUCString;
var
  i: integer;
begin
  for i := 0 to 1000 do begin
    Result :=  DefaultName + IntToStr(FManager.Worksheets.Count + i);
    if SheetByName(Result) = Nil then
      Exit;
  end;
  raise XLSRWException.Create('Can not create unique worksheet name');
end;

function TXLSWorkbook.GetUserName: AxUCString;
begin
  Result := FManager.Workbook.UserName;
end;

{$ifdef XLS_BIFF}
function TXLSWorkbook.GetVBA: TXLSVBA;
begin
  Result := FBIFF.VBA;
end;
{$endif}

function TXLSWorkbook.GetWeakPassword: AxUCString;
begin
{$ifdef XLS_BIFF}
  if FBIFF <> Nil then
    Result := PasswordFromHash(FBIFF.Records.PASSWORD)
  else
{$endif}
    Result := PasswordFromHash(FManager.Workbook.WorkbookProtection.WorkbookPassword);
end;

procedure TXLSWorkbook.InsertColumns(const ASheet, ACol, AColCount: integer);
begin
  Items[ASheet].InsertColumns(ACol,AColCount);
end;

procedure TXLSWorkbook.InsertRows(const ASheet, ARow, ARowCount: integer);
begin
  Items[ASheet].InsertRows(ARow,ARowCount);
end;

procedure TXLSWorkbook.Insert(const AIndex: integer; ACount: integer = 1);
var
  i: integer;
  Xc12Sheet: TXc12DataWorksheet;
begin
  if (AIndex >= 0) and (ACount > 0) then begin
    if not FCompiled then
      CompileFormulas;

    if AIndex > FManager.Worksheets.Count then
      AdjustSheetIndex(AIndex,FManager.Worksheets.Count)
    else
      AdjustSheetIndex(AIndex,ACount);

    for i := 0 to ACount - 1 do begin
      Xc12Sheet := FManager.Worksheets.Insert(AIndex + i);
      Xc12Sheet.Name := GetUniqueSheetname;
      FItems.Insert(AIndex + i,TXLSWorksheet.Create(FClassFactory,Self,FManager,Xc12Sheet));
{$ifdef XLS_BIFF}
      if FBIFF <> Nil then
        FBIFF.Sheets.Insert(AIndex + i);
{$endif}
    end;
    FItems.ReIndex;
  end;
end;

function TXLSWorkbook.MaxColCount: integer;
begin
{$ifdef XLS_BIFF}
  if FBIFF <> Nil then
    Result := XLS_MAXCOL_97 + 1
  else
{$endif}
    Result := XLS_MAXCOLS;
end;

function TXLSWorkbook.MaxRowCount: integer;
begin
{$ifdef XLS_BIFF}
  if FBIFF <> Nil then
    Result := XLS_MAXROW_97 + 1
  else
{$endif}
    Result := XLS_MAXROWS;
end;

procedure TXLSWorkbook.MoveCells(const SrcSheet, Col1, Row1, Col2, Row2, DestSheet, DestCol, DestRow: integer; const CopyOptions: TCopyCellsOptions);
begin
  if SrcSheet = DestSheet then
    Items[SrcSheet].MoveCells(Col1, Row1, Col2, Row2, DestCol, DestRow,CopyOptions)
  else begin
    CopyCells(SrcSheet,Col1,Row1,Col2,Row2,DestSheet,DestCol,DestRow,CopyOptions);
    DeleteCells(SrcSheet,Col1,Row1,Col2,Row2);
  end;
end;

procedure TXLSWorkbook.MoveColumns(const ASrcSheet, ACol1, ACol2, ADestSheet, ADestCol: integer);
var
  C1,R1,C2,R2: integer;
  Cols: TXc12Columns;
begin
  if ASrcSheet = ADestSheet then
    Items[ASrcSheet].MoveColumns(ACol1,ACol2,ADestCol)
  else begin
    GetSheetExtent(ASrcSheet,C1,R1,C2,R2);

    MoveCells(ASrcSheet,ACol1,R1,ACol2,R2,ADestSheet,ADestCol,0);

    Cols := Items[ASrcSheet].Columns.CopyHitList(ACol1,ACol2);
    try
      Items[ADestSheet].Columns.OffsetList(Cols,ADestCol - ACol1);
      Items[ADestSheet].Columns.InsertList(Cols);
    finally
      Cols.Free;
    end;

    Items[ASrcSheet].Columns.ClearColumns(ACol1,ACol2);
  end;
end;

procedure TXLSWorkbook.MoveRows(const ASrcSheet, ARow1, ARow2, ADestSheet, ADestRow: integer);
var
  C1,R1,C2,R2: integer;
begin
  if ASrcSheet = ADestSheet then
    Items[ASrcSheet].MoveRows(ARow1,ARow2,ADestRow)
  else begin
    GetSheetExtent(ASrcSheet,C1,R1,C2,R2);

    MoveCells(ASrcSheet,C1,ARow1,C2,ARow2,ADestSheet,0,ADestRow);
    Items[ADestSheet].FCells.CopyRowHeaders(Items[ASrcSheet].FCells,ARow1,ARow2,ADestRow);
    Items[ASrcSheet].FCells.ClearRowHeaders(ARow1,ARow2);
  end;
end;

procedure TXLSWorkbook.SetBackup(const Value: boolean);
begin
  FManager.Workbook.WorkbookPr.BackupFile := Value;
end;

procedure TXLSWorkbook.SetDefaultName(const Value: AxUCString);
begin
  if Value <> '' then
    FDefaultName := Value
  else
    FDefaultName := 'Sheet';
end;

procedure TXLSWorkbook.SetFormulasUncalced(const Value: boolean);
begin
  FFormulas.FormulasUncalced := Value;
end;

procedure TXLSWorkbook.SetNameAsBoolean(const AName: AxUCString; const Value: boolean);
var
  Name: TXLSName;
begin
  Name := FNames.Find(AName);
  if (Name <> Nil) and (Name.SimpleName in [xsntRef,xsntArea]) then
    Items[Name.Area.SheetIndex].AsBoolean[Name.Area.Col1 and not COL_ABSFLAG,Name.Area.Row1 and not ROW_ABSFLAG] := Value;
end;

procedure TXLSWorkbook.SetNameAsBoolFormulaValue(const AName: AxUCString; const Value: boolean);
begin
  SetNameAsBoolean(AName,Value);
end;

procedure TXLSWorkbook.SetNameAsError(const AName: AxUCString; const Value: TXc12CellError);
var
  Name: TXLSName;
begin
  Name := FNames.Find(AName);
  if (Name <> Nil) and (Name.SimpleName in [xsntRef,xsntArea]) then
    Items[Name.Area.SheetIndex].AsError[Name.Area.Col1 and not COL_ABSFLAG,Name.Area.Row1 and not ROW_ABSFLAG] := Value;
end;

procedure TXLSWorkbook.SetNameAsFloat(const AName: AxUCString; const Value: double);
var
  Name: TXLSName;
begin
  Name := FNames.Find(AName);
  if (Name <> Nil) and (Name.SimpleName in [xsntRef,xsntArea]) then
    Items[Name.Area.SheetIndex].AsFloat[Name.Area.Col1 and not COL_ABSFLAG,Name.Area.Row1 and not ROW_ABSFLAG] := Value;
end;

procedure TXLSWorkbook.SetNameAsFormula(const AName: AxUCString; const Value: AxUCString);
var
  Name: TXLSName;
begin
  Name := FNames.Find(AName);
  if (Name <> Nil) and (Name.SimpleName in [xsntRef,xsntArea]) then
    Items[Name.Area.SheetIndex].AsFormula[Name.Area.Col1 and not COL_ABSFLAG,Name.Area.Row1 and not ROW_ABSFLAG] := Value;
end;

procedure TXLSWorkbook.SetNameAsNumFormulaValue(const AName: AxUCString; const Value: double);
begin
  SetNameAsFloat(AName,Value);
end;

procedure TXLSWorkbook.SetNameAsStrFormulaValue(const AName: AxUCString; const Value: AxUCString);
begin
  SetNameAsString(AName,Value);
end;

procedure TXLSWorkbook.SetNameAsString(const AName: AxUCString; const Value: AxUCString);
var
  Name: TXLSName;
begin
  Name := FNames.Find(AName);
  if (Name <> Nil) and (Name.SimpleName in [xsntRef,xsntArea]) then
    Items[Name.Area.SheetIndex].AsString[Name.Area.Col1,Name.Area.Row1] := Value;
end;

procedure TXLSWorkbook.SetPalette(Index: integer; const Value: TColor);
begin
  Xc12IndexColorPalette[Index] := Longword(Value);

  FManager.PaletteChanged := True;
end;

procedure TXLSWorkbook.SetRefreshAll(const Value: boolean);
begin
  FManager.Workbook.WorkbookPr.RefreshAllConnections := Value;
end;

procedure TXLSWorkbook.SetSelectedTab(const Value: integer);
begin
  FWorkbookData.SelectedTab := Value;
end;

procedure TXLSWorkbook.SetUserName(const Value: AxUCString);
begin
  FManager.Workbook.UserName := Value;
end;

function TXLSWorkbook.SheetByName(const Name: AxUCString): TXLSWorksheet;
var
  i: integer;
begin
  for i := 0 to FItems.Count - 1 do begin
    if SameText(name,Items[i].Name) then begin
      Result := Items[i];
      Exit;
    end;
  end;
  Result := Nil;
end;

procedure TXLSWorkbook.AfterReadAdd(AIndex: integer);
var
  Sheet: TXLSWorksheet;
begin
  Sheet := TXLSWorksheet.Create(FClassFactory,Self,FManager,FManager.Worksheets[AIndex]);
  FItems.Add(Sheet);
end;

procedure TXLSWorkbook.BeforeWrite;
var
  i,j: integer;
begin
  for i := 0 to FManager.Worksheets.Count - 1 do begin
    Items[i].CalcDimensions;
    if FOptionsDialog.FRightToLeftAssigned then begin
      for j := 0 to FManager.Worksheets[i].SheetViews.Count - 1 do
        FManager.Worksheets[i].SheetViews[j].RightToLeft := FOptionsDialog.RightToLeft;
    end;
  end;
end;

procedure TXLSWorkbook.BeginFindText;
var
  i: integer;
begin
  for i := 0 to FItems.Count - 1 do
    Items[i].BeginFindText;
  FCurrFindSheet := 0;
end;

function TXLSWorkbook.Count: integer;
begin
  Result := FItems.Count;
end;

function TXLSWorkbook.GetNameIsSimpeName(const AName: AxUCString): boolean;
var
  Name: TXLSName;
begin
  Name := FNames.Find(AName);
  Result := (Name <> Nil) and (Name.SimpleName in [xsntRef,xsntArea]);
end;

{ TXLSSelectedArea }

function TXLSSelectedArea.AsRect: TXLSCellArea;
begin
  Result.Col1 := Col1;
  Result.Row1 := Row1;
  Result.Col2 := Col2;
  Result.Row2 := Row2;
end;

function TXLSSelectedArea.GetCol1: integer;
begin
  Result := FCol1;
end;

function TXLSSelectedArea.GetCol2: integer;
begin
  Result := FCol2;
end;

function TXLSSelectedArea.GetRow1: integer;
begin
  Result := FRow1;
end;

function TXLSSelectedArea.GetRow2: integer;
begin
  Result := FRow2;
end;

function TXLSSelectedArea.Hit(ACol, ARow: integer): boolean;
begin
  Result := (ACol >= Col1) and (ARow >= Row1) and (ACol <= Col2) and (ARow <= Row2);
end;

procedure TXLSSelectedArea.Intersect(C1, R1, C2, R2: integer);
begin
  if C1 < Col1 then Col1 := C1;
  if R1 < Row1 then Row1 := R1;
  if C2 > Col2 then Col2 := C2;
  if R2 > Row2 then Row2 := R2;
end;

function TXLSSelectedArea.IsColumns: boolean;
begin
  Result := (FRow1 = 0) and (FRow2 = XLS_MAXROW);
end;

procedure TXLSSelectedArea.SetArea(C1, R1, C2, R2: integer);
begin
  Col1 := C1;
  Row1 := R1;
  Col2 := C2;
  Row2 := R2;
end;

procedure TXLSSelectedArea.SetCol1(const AValue: integer);
begin
  FCol1 := AValue;
end;

procedure TXLSSelectedArea.SetCol2(const AValue: integer);
begin
  FCol2 := AValue;
end;

procedure TXLSSelectedArea.SetRow1(const AValue: integer);
begin
  FRow1 := AValue;
end;

procedure TXLSSelectedArea.SetRow2(const AValue: integer);
begin
  FRow2 := AValue;
end;

{ TXLSSelectedAreas }

function TXLSSelectedAreas.Add: TXLSSelectedArea;
begin
  Result := TXLSSelectedArea.Create;
  inherited Add(Result);
end;

function TXLSSelectedAreas.Add(C1, R1, C2, R2: integer): TXLSSelectedArea;
begin
  Result := TXLSSelectedArea.Create;
  Result.Col1 := C1;
  Result.Row1 := R1;
  Result.Col2 := C2;
  Result.Row2 := R2;
  FActiveCol := C1;
  FActiveRow := R1;
  inherited Add(Result);
end;

function TXLSSelectedAreas.CellInAreas(ACol, ARow: integer; var EdgeHit: TSelectedEdgeHits; var AreaHit: TSelectedAreaHit): integer;
const
  E_LEFT   = $01;
  E_TOP    = $02;
  E_RIGHT  = $04;
  E_BOTTOM = $08;
var
  i: integer;
  Edges: byte;
begin
  Result := 0;
  AreaHit := sahNo;
  Edges := $0;
  if (ACol = FActiveCol) and (ARow = FActiveRow) then
    AreaHit := sahActiveCell;
  for i := 0 to Count - 1 do begin
    if (ACol >= Items[i].Col1) and (ACol <= Items[i].Col2) and (ARow >= Items[i].Row1) and (ARow <= Items[i].Row2) then begin
      if AreaHit <> sahActiveCell then
        AreaHit := sahEdge;
      if ACol <> Items[i].Col1 then Edges := Edges or E_LEFT;
      if ACol <> Items[i].Col2 then Edges := Edges or E_RIGHT;
      if ARow <> Items[i].Row1 then Edges := Edges or E_TOP;
      if ARow <> Items[i].Row2 then Edges := Edges or E_BOTTOM;
      if (AreaHit <> sahActiveCell) and (Edges = $0F) then begin
        Result := i;
        EdgeHit := [];
        AreaHit := sahInside;
        Exit;
      end;
    end;
    Result := i;
  end;
  EdgeHit := TSelectedEdgeHits((not Edges) and $0F);
end;

constructor TXLSSelectedAreas.Create;
begin
  inherited Create;
  Init;
end;

procedure TXLSSelectedAreas.CursorCell(ACol, ARow: integer);
begin
  Clear;

  Init(ACol,ARow);
end;

function TXLSSelectedAreas.CursorVisible: boolean;
begin
  Result := Count = 1;
end;

function TXLSSelectedAreas.First: TXLSSelectedArea;
begin
  Result := Items[0];
end;

function TXLSSelectedAreas.GetItems(Index: integer): TXLSSelectedArea;
begin
  Result := TXLSSelectedArea(inherited Items[Index]);
end;

procedure TXLSSelectedAreas.Init(C1, R1, C2, R2, ActC, ActR: integer);
begin
  with Add do begin
    Col1 := C1;
    Col2 := C2;
    Row1 := R1;
    Row2 := R2;
  end;
  FActiveCol := ActC;
  FActiveRow := ActR;
end;

procedure TXLSSelectedAreas.Init(ACol, ARow: integer);
begin
  with Add do begin
    Col1 := ACol;
    Col2 := ACol;
    Row1 := ARow;
    Row2 := ARow;
  end;
  FActiveCol := ACol;
  FActiveRow := ARow;
end;

function TXLSSelectedAreas.Last: TXLSSelectedArea;
begin
  Result := Items[Count - 1];
end;

procedure TXLSSelectedAreas.SetActiveCol(const Value: integer);
begin
  FActiveCol := Value;
end;

procedure TXLSSelectedAreas.SetActiveRow(const Value: integer);
begin
  FActiveRow := Value;
end;

{ TXLSHorizPagebreak }

constructor TXLSHorizPagebreak.Create(APageBreak: TXc12Break);
begin
  FXc12PageBreak := APageBreak;

  FXc12PageBreak.Man := True;
  Row := APageBreak.Id;
  Col1 := APageBreak.Min;
  Col2 := APageBreak.Max;
end;

function TXLSHorizPagebreak.GetCol1: integer;
begin
  Result := FXc12PageBreak.Min;
end;

function TXLSHorizPagebreak.GetCol2: integer;
begin
  Result := FXc12PageBreak.Max;
end;

function TXLSHorizPagebreak.GetRow: integer;
begin
  Result := FXc12PageBreak.Id;
end;

procedure TXLSHorizPagebreak.SetCol1(const Value: integer);
begin
  FXc12PageBreak.Min := Value;
end;

procedure TXLSHorizPagebreak.SetCol2(const Value: integer);
begin
  FXc12PageBreak.Max := Value;
end;

procedure TXLSHorizPagebreak.SetRow(const Value: integer);
begin
  FXc12PageBreak.Id := Value;
end;

{ TXLSHorizPagebreaks }

function TXLSHorizPagebreaks.Add: TXLSHorizPagebreak;
begin
  Result := TXLSHorizPagebreak.Create(FXc12PageBreaks.Add);
  Result.Col1 := 0;
  Result.Col2 := XLS_MAXCOL;
  inherited Add(Result);
end;

function TXLSHorizPagebreaks.Add(const ARow: integer): TXLSHorizPagebreak;
begin
  Result := Add;
  Result.Row := ARow;
end;

constructor TXLSHorizPagebreaks.Create(AXc12PageBreaks: TXc12PageBreaks);
var
  i: integer;
  HPB: TXLSHorizPagebreak;
begin
  inherited Create;

  FXc12PageBreaks := AXc12PageBreaks;

  for i := 0 to FXc12PageBreaks.Count - 1 do begin
    HPB := TXLSHorizPagebreak.Create(FXc12PageBreaks[i]);
    HPB.Row := FXc12PageBreaks[i].Id;
    HPB.Col1 := FXc12PageBreaks[i].Min;
    HPB.Col2 := FXc12PageBreaks[i].Max;

    inherited Add(HPB);
  end;
end;

function TXLSHorizPagebreaks.Find(const ARow: integer): TXLSHorizPagebreak;
var
  i: integer;
begin
  i := FindIndex(ARow);
  if i >= 0 then
    Result := Items[i]
  else
    Result := Nil;
end;

function TXLSHorizPagebreaks.FindIndex(const ARow: integer): integer;
begin
  for Result := 0 to Count - 1 do begin
    if Items[Result].Row = ARow then
      Exit;
  end;
  Result := -1;
end;

function TXLSHorizPagebreaks.GetItems(Index: integer): TXLSHorizPagebreak;
begin
  Result := TXLSHorizPagebreak(inherited Items[Index]);
end;

{ TXLSVertPagebreak }

constructor TXLSVertPagebreak.Create(APageBreak: TXc12Break);
begin
  FXc12PageBreak := APageBreak;

  FXc12PageBreak.Man := True;
  Col := APageBreak.Id;
  Row1 := APageBreak.Min;
  Row2 := APageBreak.Max;
end;

function TXLSVertPagebreak.GetCol: integer;
begin
  Result := FXc12PageBreak.Id;
end;

function TXLSVertPagebreak.GetRow1: integer;
begin
  Result := FXc12PageBreak.Min;
end;

function TXLSVertPagebreak.GetRow2: integer;
begin
  Result := FXc12PageBreak.Max;
end;

procedure TXLSVertPagebreak.SetCol(const Value: integer);
begin
  FXc12PageBreak.Id := Value;
end;

procedure TXLSVertPagebreak.SetRow1(const Value: integer);
begin
  FXc12PageBreak.Min := Value;
end;

procedure TXLSVertPagebreak.SetRow2(const Value: integer);
begin
  FXc12PageBreak.Max := Value;
end;

{ TXLSVertPagebreaksAbs }

function TXLSVertPagebreaks.Add: TXLSVertPagebreak;
begin
  Result := TXLSVertPagebreak.Create(FXc12PageBreaks.Add);
  Result.Row1 := 0;
  Result.Row2 := XLS_MAXROW;
  inherited Add(Result);
end;

function TXLSVertPagebreaks.Add(const ACol: integer): TXLSVertPagebreak;
begin
  Result := Add;
  Result.Col := ACol;
end;

constructor TXLSVertPagebreaks.Create(AXc12PageBreaks: TXc12PageBreaks);
var
  i: integer;
  VPB: TXLSVertPagebreak;
begin
  inherited Create;

  FXc12PageBreaks := AXc12PageBreaks;

  for i := 0 to FXc12PageBreaks.Count - 1 do begin
    VPB := TXLSVertPagebreak.Create(FXc12PageBreaks[i]);
    VPB.Col := FXc12PageBreaks[i].Id;
    VPB.Row1 := FXc12PageBreaks[i].Min;
    VPB.Row2 := FXc12PageBreaks[i].Max;

    inherited Add(VPB);
  end;
end;

function TXLSVertPagebreaks.Find(const ACol: integer): TXLSVertPagebreak;
var
  i: integer;
begin
  i := FindIndex(ACol);
  if i >= 0 then
    Result := Items[i]
  else
    Result := Nil;
end;

function TXLSVertPagebreaks.FindIndex(const ACol: integer): integer;
begin
  for Result := 0 to Count - 1 do begin
    if Items[Result].Col = ACol then
      Exit;
  end;
  Result := -1;
end;

function TXLSVertPagebreaks.GetItems(Index: integer): TXLSVertPagebreak;
begin
  Result := TXLSVertPagebreak(inherited Items[Index]);
end;

{ TXLSPrintSettings }

procedure TXLSPrintSettings.Clear;
begin
  FXc12Sheet.PageSetup.Clear;
  FXc12Sheet.HeaderFooter.Clear;
  FXc12Sheet.PageMargins.Clear;
  FXc12Sheet.ColBreaks.Clear;
  FXc12Sheet.RowBreaks.Clear;
  FXc12Sheet.PrintOptions.Clear;

  ClearPrintArea;
  ClearPrintTitles;
end;

procedure TXLSPrintSettings.ClearPrintArea;
begin
  FNames.Delete(bnPrintArea,FXc12Sheet.Index);
end;

procedure TXLSPrintSettings.ClearPrintTitles;
begin
  FNames.Delete(bnPrintTitles,FXc12Sheet.Index);
end;

constructor TXLSPrintSettings.Create(AXc12Sheet: TXc12DataWorksheet; ANames: TXLSNames);
begin
  FXc12Sheet := AXc12Sheet;
  FNames := ANames;
  FXLSHeaderFooter := TXLSHeaderFooter.Create(FXc12Sheet.HeaderFooter);
  FVertPagebreaks := TXLSVertPagebreaks.Create(FXc12Sheet.ColBreaks);
  FHorizPagebreaks := TXLSHorizPagebreaks.Create(FXc12Sheet.RowBreaks);
end;

destructor TXLSPrintSettings.Destroy;
begin
  FXLSHeaderFooter.Free;
  FHorizPagebreaks.Free;
  FVertPagebreaks.Free;
  inherited;
end;

function TXLSPrintSettings.GetCopies: integer;
begin
  Result := FXc12Sheet.PageSetup.Copies;
end;

function TXLSPrintSettings.GetFitHeight: integer;
begin
  Result := FXc12Sheet.PageSetup.FitToHeight;
end;

function TXLSPrintSettings.GetFitWidth: integer;
begin
  Result := FXc12Sheet.PageSetup.FitToWidth;
end;

function TXLSPrintSettings.GetFooterMargin: double;
begin
  Result := FXc12Sheet.PageMargins.Footer;
end;

function TXLSPrintSettings.GetFooterMarginCm: double;
begin
  Result := FXc12Sheet.PageMargins.Footer * 2.54
end;

function TXLSPrintSettings.GetHeaderMargin: double;
begin
  Result := FXc12Sheet.PageMargins.Header;
end;

function TXLSPrintSettings.GetHeaderMarginCm: double;
begin
  Result := FXc12Sheet.PageMargins.Header * 2.54;
end;

function TXLSPrintSettings.GetMarginBottom: double;
begin
  Result := FXc12Sheet.PageMargins.Bottom;
end;

function TXLSPrintSettings.GetMarginBottomCm: double;
begin
  Result := FXc12Sheet.PageMargins.Bottom * 2.54;
end;

function TXLSPrintSettings.GetMarginLeft: double;
begin
  Result := FXc12Sheet.PageMargins.Left;
end;

function TXLSPrintSettings.GetMarginLeftCm: double;
begin
  Result := FXc12Sheet.PageMargins.Left * 2.54;
end;

function TXLSPrintSettings.GetMarginRight: double;
begin
  Result := FXc12Sheet.PageMargins.Right;
end;

function TXLSPrintSettings.GetMarginRightCm: double;
begin
  Result := FXc12Sheet.PageMargins.Right * 2.54;
end;

function TXLSPrintSettings.GetMarginTop: double;
begin
  Result := FXc12Sheet.PageMargins.Top;
end;

function TXLSPrintSettings.GetMarginTopCm: double;
begin
  Result := FXc12Sheet.PageMargins.Top * 2.54;
end;

function TXLSPrintSettings.GetOptions: TPrintSetupOptions;
begin
  Result := [];
  if x12oPortrait in [FXc12Sheet.PageSetup.Orientation] then
    Result := Result + [psoPortrait];
  if FXc12Sheet.PageSetup.BlackAndWhite then
    Result := Result + [psoNoColor];
  if FXc12Sheet.PageSetup.Draft then
    Result := Result + [psoDraftQuality];
  if x12ccAsDisplayed in [FXc12Sheet.PageSetup.CellComments] then
    Result := Result + [psoNotes];
  if FXc12Sheet.PrintOptions.Headings then
    Result := Result + [psoRowColHeading];
  if FXc12Sheet.PrintOptions.GridLines then
    Result := Result + [psoGridlines];
  if FXc12Sheet.PrintOptions.HorizontalCentered then
    Result := Result + [psoHorizCenter];
  if FXc12Sheet.PrintOptions.VerticalCentered then
    Result := Result + [psoVertCenter];
end;

function TXLSPrintSettings.GetPaperSize: TXc12PaperSize;
begin
  Result := TXc12PaperSize(FXc12Sheet.PageSetup.PaperSize)
end;

function TXLSPrintSettings.GetPrintArea(out ACol1, ARow1, ACol2, ARow2: integer): boolean;
var
  Name: TXLSName;
begin
  Name := TXLSName(FNames.FindBuiltIn(bnPrintArea,FXc12Sheet.Index));
  Result := (Name <> Nil) and (Name.SimpleName = xsntArea);
  if Result then begin
    ACol1 := Name.SimpleArea.Col1 and not COL_ABSFLAG;
    ARow1 := Name.SimpleArea.Row1 and not ROW_ABSFLAG;
    ACol2 := Name.SimpleArea.Col2 and not COL_ABSFLAG;
    ARow2 := Name.SimpleArea.Row2 and not ROW_ABSFLAG;
  end;
end;

function TXLSPrintSettings.GetPrintTitles(out ACol1, ACol2, ARow1, ARow2: integer): boolean;
var
  Name: TXLSName;
begin
  Name := TXLSName(FNames.FindBuiltIn(bnPrintTitles,FXc12Sheet.Index));
  Result := (Name <> Nil) and (Name.SimpleName = xsntArea);
  if Result then begin
    ACol1 := Name.SimpleArea.Col1 and not COL_ABSFLAG;
    ARow1 := Name.SimpleArea.Row1 and not ROW_ABSFLAG;
    ACol2 := Name.SimpleArea.Col2 and not COL_ABSFLAG;
    ARow2 := Name.SimpleArea.Row2 and not ROW_ABSFLAG;
  end;
end;

function TXLSPrintSettings.GetPrintTitlesCols(out ACol1, ACol2: integer): boolean;
var
  R1,R2: integer;
begin
  Result := GetPrintTitles(ACol1,R1,ACol2,R2) and (ACol2 < XLS_MAXCOL);
end;

function TXLSPrintSettings.GetPrintTitlesRows(out ARow1, ARow2: integer): boolean;
var
  C1,C2: integer;
begin
  Result := GetPrintTitles(C1,ARow1,C2,ARow2) and (ARow2 < XLS_MAXROW);
end;

function TXLSPrintSettings.GetResolution: integer;
begin
  Result := FXc12Sheet.PageSetup.HorizontalDpi;
end;

function TXLSPrintSettings.GetScalingFactor: integer;
begin
  Result := FXc12Sheet.PageSetup.Scale;
end;

function TXLSPrintSettings.GetStartingPage: integer;
begin
  Result := FXc12Sheet.PageSetup.FirstPageNumber;
end;

function TXLSPrintSettings.PaperSizeDim: TXLSPointF;
begin
  case GetPaperSize of
    psNone:        begin Result.X := 0;   Result.Y := 0;   end;
    psLetter:      begin Result.X := 21.6; Result.Y := 27.9; end;
    psLetterSmall: begin Result.X := 21.6; Result.Y := 27.9; end;
    psTabloid:     begin Result.X := 27.9; Result.Y := 43.2; end;
    psLedger:      begin Result.X := 27.9; Result.Y := 43.2; end;
    psLegal:       begin Result.X := 21.6; Result.Y := 35.6; end;
    psStatement:   begin Result.X := 14.0; Result.Y := 21.6; end;
    psExecutive:   begin Result.X := 18.4; Result.Y := 26.7; end;
    psA3:          begin Result.X := 29.7; Result.Y := 42.0; end;
    psA4:          begin Result.X := 21.0; Result.Y := 29.7; end;
    psA4Small:     begin Result.X := 21.0; Result.Y := 29.7; end;
    psA5:          begin Result.X := 14.8; Result.Y := 21.0; end;
    psB4:          begin Result.X := 25.0; Result.Y := 35.4; end;
    psB5:          begin Result.X := 18.2; Result.Y := 25.7; end;
//    psFolio,                 //* Folio 8 12 x 13 in
    psQuarto:      begin Result.X := 215; Result.Y := 275; end;
//    ps10X14,                 //* 10x14 in
//    ps11X17,                 //* 11x17 in
//    psNote,                  //* Note 8.12 x 11 in
//    psEnv9,                  //* Envelope #9 3.78 x 8.78
//    psEnv10,                 //* Envelope #10 4.18 x 9.12
//    psEnv11,                 //* Envelope #11 4.12 x 10.38
//    psEnv12,                 //* Envelope #12 4 \276 x 11
//    psEnv14,                 //* Envelope #14 5 x 11.12
//    psCSheet,                //* C size sheet
//    psDSheet,                //* D size sheet
//    psESheet,                //* E size sheet
    psEnvDL:       begin Result.X := 11.0; Result.Y := 22.0; end;
    psEnvC5:       begin Result.X := 16.2; Result.Y := 22.9; end;
    psEnvC3:       begin Result.X := 32.4; Result.Y := 45.8; end;
    psEnvC4:       begin Result.X := 22.9; Result.Y := 32.4; end;
    psEnvC6:       begin Result.X := 11.4; Result.Y := 16.2; end;
    psEnvC65:      begin Result.X := 11.4; Result.Y := 22.9; end;
    psEnvB4:       begin Result.X := 25.0; Result.Y := 35.3; end;
    psEnvB5:       begin Result.X := 17.6; Result.Y := 25.0; end;
    psEnvB6:       begin Result.X := 17.6; Result.Y := 12.5; end;
    psEnvItaly:    begin Result.X := 11.0; Result.Y := 23.0; end;
//    psEnvMonarch,            //* Envelope Monarch 3.875 x 7.5 in
//    psEnvPersonal,           //* 6 34 Envelope 3 58 x 6 12 in
//    psFanfoldUS,             //* US Std Fanfold 14 78 x 11 in
//    psFanfoldStdGerman,      //* German Std Fanfold 8 12 x 12 in
//    psFanfoldLglGerman,      //* German Legal Fanfold 8 12 x 13 in
    psISO_B4:      begin Result.X := 25.0; Result.Y := 35.3; end;
    psJapanesePostcard: begin Result.X := 10.0; Result.Y := 14.8; end;
//    ps9X11,                  //* 9 x 11 in
//    ps10X11,                 //* 10 x 11 in
//    ps15X11,                 //* 15 x 11 in
    psEnvInvite:   begin Result.X := 22.0; Result.Y := 22.0; end;
//    psReserved48,            //* RESERVED--DO NOT USE
//    psReserved49,            //* RESERVED--DO NOT USE
//    psLetterExtra,           //* Letter Extra 9 \275 x 12 in
//    psLegalExtra,            //* Legal Extra 9 \275 x 15 in
//    psTabloidExtra,          //* Tabloid Extra 11.69 x 18 in
//    psA4Extra,               //* A4 Extra 9.27 x 12.69 in
//    psLetterTransverse,      //* Letter Transverse 8 \275 x 11 in
    psA4Transverse: begin Result.X := 21.0; Result.Y := 29.7; end;
//    psLetterExtraTransverse, //* Letter Extra Transverse 9\275 x 12 in
    psAPlus:        begin Result.X := 22.7; Result.Y := 35.6; end;
    psBPlus:        begin Result.X := 30.5; Result.Y := 48.7; end;
//    psLetterPlus,            //* Letter Plus 8.5 x 12.69 in
    psA4Plus:       begin Result.X := 21.0; Result.Y := 33.0; end;
    psA5Transverse: begin Result.X := 14.8; Result.Y := 21.0; end;
    psB5transverse: begin Result.X := 18.2; Result.Y := 25.7; end;
    psA3Extra:      begin Result.X := 32.2; Result.Y := 44.5; end;
    psA5Extra:      begin Result.X := 17.4; Result.Y := 23.5; end;
    psB5Extra:      begin Result.X := 20.1; Result.Y := 27.6; end;
    psA2:           begin Result.X := 42.0; Result.Y := 59.4; end;
    psA3Transverse: begin Result.X := 29.7; Result.Y := 42.0; end;
    psA3ExtraTransverse: begin Result.X := 32.2; Result.Y := 44.5; end;
  end;
end;

procedure TXLSPrintSettings.PrintArea(const ACol1, ARow1, ACol2, ARow2: integer);
var
  Name: TXLSName;
begin
  Name := TXLSName(FNames.FindBuiltIn(bnPrintArea,FXc12Sheet.Index));
  if Name = Nil then
    Name := TXLSName(FNames.Add(bnPrintArea,FXc12Sheet.Index));

  Name.Definition := '''' + FXc12Sheet.Name + '''!' +  AreaToRefStr(ACol1,ARow1,ACol2,ARow2,True,True,True,True);
end;

procedure TXLSPrintSettings.PrintTitles(const ACol1,ACol2,ARow1,ARow2: integer);
var
  Name: TXLSName;
  A1,A2: AxUCString;
begin
  if ((ACol1 = -1) or (ACol2 = -1)) and ((ARow1 = -1) or (ARow2 = -1)) then
    ClearPrintTitles
  else begin
    Name := TXLSName(FNames.FindBuiltIn(bnPrintTitles,FXc12Sheet.Index));
    if Name = Nil then
      Name := TXLSName(FNames.Add(bnPrintTitles,FXc12Sheet.Index));

    A1 := '';
    A2 := '';
    if (ACol1 >= 0) and (ACol2 >= 0) then
      A1 := '''' + FXc12Sheet.Name + '''!' +  ColsToRefStr(ACol1,ACol2,True);
    if (ARow1 >= 0) and (ARow2 >= 0) then
      A2 := '''' + FXc12Sheet.Name + '''!' +  RowsToRefStr(ARow1,ARow2,True);

    if (A1 <> '') and (A2 <> '') then
      Name.Definition := A1 + ',' + A2
    else if A1 <> '' then
      Name.Definition := A1
    else
      Name.Definition := A2;
  end;
end;

procedure TXLSPrintSettings.SetCopies(const Value: integer);
begin
  FXc12Sheet.PageSetup.Copies := Value;
end;

procedure TXLSPrintSettings.SetFitHeight(const Value: integer);
begin
  FXc12Sheet.PageSetup.FitToHeight := Value;
end;

procedure TXLSPrintSettings.SetFitWidth(const Value: integer);
begin
  FXc12Sheet.PageSetup.FitToWidth := Value;
end;

procedure TXLSPrintSettings.SetFooterMargin(const Value: double);
begin
  FXc12Sheet.PageMargins.Footer := Value;
end;

procedure TXLSPrintSettings.SetFooterMarginCm(const Value: double);
begin
  FXc12Sheet.PageMargins.Footer := Value / 2.54;
end;

procedure TXLSPrintSettings.SetHeaderMargin(const Value: double);
begin
  FXc12Sheet.PageMargins.Header := Value;
end;

procedure TXLSPrintSettings.SetHeaderMarginCm(const Value: double);
begin
  FXc12Sheet.PageMargins.Header := Value / 2.54;
end;

procedure TXLSPrintSettings.SetMarginBottom(const Value: double);
begin
  FXc12Sheet.PageMargins.Bottom := Value;
end;

procedure TXLSPrintSettings.SetMarginBottomCm(const Value: double);
begin
  FXc12Sheet.PageMargins.Bottom := Value / 2.54;
end;

procedure TXLSPrintSettings.SetMarginLeft(const Value: double);
begin
  FXc12Sheet.PageMargins.Left := Value;
end;

procedure TXLSPrintSettings.SetMarginLeftCm(const Value: double);
begin
  FXc12Sheet.PageMargins.Left := Value / 2.54;
end;

procedure TXLSPrintSettings.SetMarginRight(const Value: double);
begin
  FXc12Sheet.PageMargins.Right := Value;
end;

procedure TXLSPrintSettings.SetMarginRightCm(const Value: double);
begin
  FXc12Sheet.PageMargins.Right := Value / 2.54;
end;

procedure TXLSPrintSettings.SetMarginTop(const Value: double);
begin
  FXc12Sheet.PageMargins.Top := Value;
end;

procedure TXLSPrintSettings.SetMarginTopCm(const Value: double);
begin
  FXc12Sheet.PageMargins.Top := Value / 2.54;
end;

procedure TXLSPrintSettings.SetOptions(const Value: TPrintSetupOptions);
begin
  if psoPortrait in Value then
    FXc12Sheet.PageSetup.Orientation := x12oPortrait
  else
    FXc12Sheet.PageSetup.Orientation := x12oLandscape;

  FXc12Sheet.PageSetup.BlackAndWhite := psoNoColor in Value;

  FXc12Sheet.PageSetup.Draft := psoDraftQuality in Value;

  if psoNotes in Value then
    FXc12Sheet.PageSetup.CellComments := x12ccAsDisplayed
  else
    FXc12Sheet.PageSetup.CellComments := x12ccNone;

  FXc12Sheet.PrintOptions.Headings := psoRowColHeading in Value;

  FXc12Sheet.PrintOptions.GridLines := psoGridlines in Value;

  FXc12Sheet.PrintOptions.HorizontalCentered := psoHorizCenter in Value;

  FXc12Sheet.PrintOptions.VerticalCentered := psoVertCenter in Value;
end;

procedure TXLSPrintSettings.SetPaperSize(const Value: TXc12PaperSize);
begin
  FXc12Sheet.PageSetup.PaperSize := Integer(Value);
end;

procedure TXLSPrintSettings.SetResolution(const Value: integer);
begin
  FXc12Sheet.PageSetup.HorizontalDpi := Value;
end;

procedure TXLSPrintSettings.SetScalingFactor(const Value: integer);
begin
  FXc12Sheet.PageSetup.Scale := Value;
end;

procedure TXLSPrintSettings.SetStartingPage(const Value: integer);
begin
  FXc12Sheet.PageSetup.FirstPageNumber := Value;
end;

{ TXLSPane }

procedure TXLSPane.Clear;
begin
  FXc12Pane.Clear;
end;

constructor TXLSPane.Create(const AXc12Pane: TXc12Pane);
begin
  FXc12Pane := AXc12Pane;
end;

function TXLSPane.GetActivePane: TActivePane;
begin
  case FXc12Pane.ActivePane of
    x12pBottomRight: Result := apBottomRight;
    x12pTopRight   : Result := apTopRight;
    x12pBottomLeft : Result := apBottomLeft;
    x12pTopLeft    : Result := apTopLeft;
    else             Result := apBottomLeft;
  end;
end;

function TXLSPane.GetLeftCol: integer;
begin
  Result := FXc12Pane.TopLeftCell.Col1;
end;

function TXLSPane.GetPaneType: TPaneType;
begin
  case FXc12Pane.State of
    x12psSplit      : Result := ptSplit;
    x12psFrozen     : Result := ptFrozen;
    x12psFrozenSplit: Result := ptFrozenSplit;
    else              Result := ptNone;
  end;
end;

function TXLSPane.GetSplitColX: double;
begin
  Result := FXc12Pane.XSplit;
end;

function TXLSPane.GetSplitRowY: double;
begin
  Result := FXc12Pane.YSplit;
end;

function TXLSPane.GetTopRow: integer;
begin
  Result := FXc12Pane.TopLeftCell.Row1;
end;

procedure TXLSPane.SetActivePane(const Value: TActivePane);
begin
  case Value of
    apBottomRight: FXc12Pane.ActivePane := x12pBottomRight;
    apTopRight   : FXc12Pane.ActivePane := x12pTopRight;
    apBottomLeft : FXc12Pane.ActivePane := x12pBottomLeft;
    apTopLeft    : FXc12Pane.ActivePane := x12pTopLeft;
  end;
end;

procedure TXLSPane.SetLeftCol(const Value: integer);
var
  Tmp: PXLSCellArea;
begin
  Tmp := @FXc12Pane.TopLeftCell;
  Tmp.Col1 := Value;
  if Tmp.Row1 < 0 then
    Tmp.Row1 := 0;

  FXc12Pane.XSplit := Value;
end;

procedure TXLSPane.SetPaneType(const Value: TPaneType);
begin
  case Value of
    ptSplit      : FXc12Pane.State := x12psSplit;
    ptFrozen     : FXc12Pane.State := x12psFrozen;
    ptFrozenSplit: FXc12Pane.State := x12psFrozenSplit;
  end;
end;

procedure TXLSPane.SetSplitColX(const Value: double);
begin
  FXc12Pane.XSplit := Value;
end;

procedure TXLSPane.SetSplitRowY(const Value: double);
begin
  FXc12Pane.YSplit := Value;
end;

procedure TXLSPane.SetTopRow(const Value: integer);
var
  Tmp: PXLSCellArea;
begin
  Tmp := @FXc12Pane.TopLeftCell;
  Tmp.Row1 := Value;
  if Tmp.Col1 < 0 then
    Tmp.Col1 := 0;

  FXc12Pane.YSplit := Value;
end;

{ TXLSHeaderFooter }

procedure TXLSHeaderFooter.Clear;
begin
  FXc12HeaderFooter.Clear;
end;

constructor TXLSHeaderFooter.Create(AXc12HeaderFooter: TXc12HeaderFooter);
begin
  FXc12HeaderFooter := AXc12HeaderFooter;
end;

function TXLSHeaderFooter.GetAlignWithMargins: boolean;
begin
  Result := FXc12HeaderFooter.AlignWithMargins;
end;

function TXLSHeaderFooter.GetDifferentFirst: boolean;
begin
  Result := FXc12HeaderFooter.DifferentFirst;
end;

function TXLSHeaderFooter.GetDifferentOddEven: boolean;
begin
  Result := FXc12HeaderFooter.DifferentOddEven;
end;

function TXLSHeaderFooter.GetEvenFooter: AxUCString;
begin
  Result := FXc12HeaderFooter.EvenFooter;
end;

function TXLSHeaderFooter.GetEvenHeader: AxUCString;
begin
  Result := FXc12HeaderFooter.EvenHeader;
end;

function TXLSHeaderFooter.GetFirstFooter: AxUCString;
begin
  Result := FXc12HeaderFooter.FirstFooter;
end;

function TXLSHeaderFooter.GetFirstHeader: AxUCString;
begin
  Result := FXc12HeaderFooter.FirstHeader;
end;

function TXLSHeaderFooter.GetOddFooter: AxUCString;
begin
  Result := FXc12HeaderFooter.OddFooter;
end;

function TXLSHeaderFooter.GetOddHeader: AxUCString;
begin
  Result := FXc12HeaderFooter.OddHeader;
end;

function TXLSHeaderFooter.GetScaleWithDoc: boolean;
begin
  Result := FXc12HeaderFooter.ScaleWithDoc;
end;

procedure TXLSHeaderFooter.SetAlignWithMargins(const Value: boolean);
begin
  FXc12HeaderFooter.AlignWithMargins := Value;
end;

procedure TXLSHeaderFooter.SetDifferentFirst(const Value: boolean);
begin
  FXc12HeaderFooter.DifferentFirst := Value;
end;

procedure TXLSHeaderFooter.SetDifferentOddEven(const Value: boolean);
begin
  FXc12HeaderFooter.DifferentOddEven := Value;
end;

procedure TXLSHeaderFooter.SetEvenFooter(const Value: AxUCString);
begin
  FXc12HeaderFooter.EvenFooter := Value;
end;

procedure TXLSHeaderFooter.SetEvenHeader(const Value: AxUCString);
begin
  FXc12HeaderFooter.EvenHeader := Value;
end;

procedure TXLSHeaderFooter.SetFirstFooter(const Value: AxUCString);
begin
  FXc12HeaderFooter.FirstFooter := Value;
end;

procedure TXLSHeaderFooter.SetFirstHeader(const Value: AxUCString);
begin
  FXc12HeaderFooter.FirstHeader := Value;
end;

procedure TXLSHeaderFooter.SetOddFooter(const Value: AxUCString);
begin
  FXc12HeaderFooter.OddFooter := Value;
end;

procedure TXLSHeaderFooter.SetOddHeader(const Value: AxUCString);
begin
  FXc12HeaderFooter.OddHeader := Value;
end;

procedure TXLSHeaderFooter.SetScaleWithDoc(const Value: boolean);
begin
  FXc12HeaderFooter.ScaleWithDoc := Value;
end;

{ TXLSOptionsDialog }

procedure TXLSOptionsDialog.Clear;
begin
  FRightToLeftAssigned := False;
  FRightToLeft := False;
  FSaveRecalc := False;
end;

constructor TXLSOptionsDialog.Create(ACalcPr: TXc12CalcPr; AWorkbookPr: TXc12WorkbookPr);
begin
  FCalcPr := ACalcPr;
  FWorkbookPr := AWorkbookPr;
end;

function TXLSOptionsDialog.GetMaxIterations: integer;
begin
  Result := FCalcPr.IterateCount;
end;

function TXLSOptionsDialog.GetCalcMode: TXc12CalcMode;
begin
  Result := FCalcPr.CalcMode;
end;

function TXLSOptionsDialog.GetDate1904: boolean;
begin
  Result := FWorkbookPr.Date1904;
end;

function TXLSOptionsDialog.GetDelta: double;
begin
  Result := FCalcPr.IterateDelta;
end;

function TXLSOptionsDialog.GetIteration: boolean;
begin
  Result := FCalcPr.Iterate;
end;

function TXLSOptionsDialog.GetPrecisionAsDisplayed: boolean;
begin
  Result := not FCalcPr.FullPrecision;
end;

function TXLSOptionsDialog.GetR1C1Mode: boolean;
begin
  Result := FCalcPr.RefMode = x12rmR1C1;
end;

function TXLSOptionsDialog.GetRecalcBeforeSave: boolean;
begin
  Result := FCalcPr.CalcOnSave;
end;

function TXLSOptionsDialog.GetRightToLeft: boolean;
begin
  Result := False;
end;

function TXLSOptionsDialog.GetSaveExtLinkVal: boolean;
begin
  Result := FWorkbookPr.SaveExternalLinkValues;
end;

function TXLSOptionsDialog.GetSaveRecalc: boolean;
begin
  Result := FSaveRecalc;
end;

function TXLSOptionsDialog.GetShowObjects: TXc12Objects;
begin
  Result := FWorkbookPr.ShowObjects;
end;

function TXLSOptionsDialog.GetUncalced: boolean;
begin
  Result := not FCalcPr.CalcCompleted;
end;

procedure TXLSOptionsDialog.SetMaxIterations(const Value: integer);
begin
  FCalcPr.IterateCount := Value;
end;

procedure TXLSOptionsDialog.SetCalcMode(const Value: TXc12CalcMode);
begin
  FCalcPr.CalcMode := Value;
end;

procedure TXLSOptionsDialog.SetDate1904(const Value: boolean);
begin
  FWorkbookPr.Date1904 := Value;
end;

procedure TXLSOptionsDialog.SetDelta(const Value: double);
begin
  FCalcPr.IterateDelta := Value;
end;

procedure TXLSOptionsDialog.SetIteration(const Value: boolean);
begin
  FCalcPr.Iterate := Value;
end;

procedure TXLSOptionsDialog.SetPrecisionAsDisplayed(const Value: boolean);
begin
  FCalcPr.FullPrecision := not Value;
end;

procedure TXLSOptionsDialog.SetR1C1Mode(const Value: boolean);
begin
  if Value then
    FCalcPr.RefMode := x12rmR1C1
  else
    FCalcPr.RefMode := x12rmA1;
end;

procedure TXLSOptionsDialog.SetRecalcBeforeSave(const Value: boolean);
begin
  FCalcPr.CalcOnSave := Value;
end;

procedure TXLSOptionsDialog.SetRightToLeft(const Value: boolean);
begin
  FRightToLeftAssigned := True;
end;

procedure TXLSOptionsDialog.SetSaveExtLinkVal(const Value: boolean);
begin
  FWorkbookPr.SaveExternalLinkValues := Value;
end;

procedure TXLSOptionsDialog.SetSaveRecalc(const Value: boolean);
begin
  FSaveRecalc := Value;
end;

procedure TXLSOptionsDialog.SetShowObjects(const Value: TXc12Objects);
begin
  FWorkbookPr.ShowObjects := Value;
end;

procedure TXLSOptionsDialog.SetUncalced(const Value: boolean);
begin
  FCalcPr.CalcCompleted := not Value;
end;

{ TXLSWorkbookData }

constructor TXLSWorkbookData.Create(ABookViews: TXc12BookViews);
begin
  FBookViews := ABookViews;
end;

function TXLSWorkbookData.ReadHeight: integer;
begin
  Result := FBookViews[0].WindowHeight;
end;

function TXLSWorkbookData.ReadLeft: integer;
begin
  Result := FBookViews[0].XWindow;
end;

function TXLSWorkbookData.ReadOptions: TWorkbookOptions;
begin
  Result := [];
  if FBookViews[0].Visibility in [x12vHidden,x12vVeryHidden] then
    Result := Result + [woHidden];
  if FBookViews[0].Minimized then
    Result := Result + [woIconized];
  if FBookViews[0].ShowVerticalScroll then
    Result := Result + [woVScroll];
  if FBookViews[0].ShowHorizontalScroll then
    Result := Result + [woHScroll];
  if FBookViews[0].ShowSheetTabs then
    Result := Result + [woTabs];
end;

function TXLSWorkbookData.ReadSelectedTab: integer;
begin
  Result := FBookViews[0].ActiveTab;
end;

function TXLSWorkbookData.ReadTop: integer;
begin
  Result := FBookViews[0].YWindow;
end;

function TXLSWorkbookData.ReadWidth: integer;
begin
  Result := FBookViews[0].WindowWidth;
end;

procedure TXLSWorkbookData.WriteHeight(const Value: integer);
begin
  FBookViews[0].WindowHeight := Value;
end;

procedure TXLSWorkbookData.WriteLeft(const Value: integer);
begin
  FBookViews[0].XWindow := Value;
end;

procedure TXLSWorkbookData.WriteOptions(const Value: TWorkbookOptions);
begin
  if woHidden in Value then
    FBookViews[0].Visibility := x12vHidden
  else
    FBookViews[0].Visibility := x12vVisible;
  FBookViews[0].Minimized := woIconized in Value;
  FBookViews[0].ShowVerticalScroll := woVScroll in Value;
  FBookViews[0].ShowHorizontalScroll := woHScroll in Value;
  FBookViews[0].ShowSheetTabs := woTabs in Value;
end;

procedure TXLSWorkbookData.WriteSelectedTab(const Value: integer);
begin
  FBookViews[0].ActiveTab := Value;
end;

procedure TXLSWorkbookData.WriteTop(const Value: integer);
begin
  FBookViews[0].YWindow := Value;
end;

procedure TXLSWorkbookData.WriteWidth(const Value: integer);
begin
  FBookViews[0].WindowWidth := Value;
end;

{ TXLSSheetProtection }

procedure TXLSSheetProtection.Clear;
begin
  FProtection.Clear;
end;

constructor TXLSSheetProtection.Create(AProtection: TXc12SheetProtection);
begin
  FProtection := AProtection;
end;

function TXLSSheetProtection.GetAutoFilter: boolean;
begin
  Result := FProtection.AutoFilter;
end;

function TXLSSheetProtection.GetDeleteColumns: boolean;
begin
  Result := FProtection.DeleteColumns;
end;

function TXLSSheetProtection.GetDeleteRows: boolean;
begin
  Result := FProtection.DeleteRows;
end;

function TXLSSheetProtection.GetFormatCells: boolean;
begin
  Result := FProtection.FormatCells;
end;

function TXLSSheetProtection.GetFormatColumns: boolean;
begin
  Result := FProtection.FormatColumns;
end;

function TXLSSheetProtection.GetFormatRows: boolean;
begin
  Result := FProtection.FormatRows;
end;

function TXLSSheetProtection.GetInsertColumns: boolean;
begin
  Result := FProtection.InsertColumns;
end;

function TXLSSheetProtection.GetInsertHyperlinks: boolean;
begin
  Result := FProtection.InsertHyperlinks;
end;

function TXLSSheetProtection.GetInsertRows: boolean;
begin
  Result := FProtection.InsertRows;
end;

function TXLSSheetProtection.GetObjects: boolean;
begin
  Result := FProtection.Objects;
end;

function TXLSSheetProtection.GetPasswordString: AxUCString;
begin
  Result := FProtection.PasswordAsString;
end;

function TXLSSheetProtection.GetPivotTables: boolean;
begin
  Result := FProtection.PivotTables;
end;

function TXLSSheetProtection.GetScenarios: boolean;
begin
  Result := FProtection.Scenarios;
end;

function TXLSSheetProtection.GetSelectLockedCells: boolean;
begin
  Result := FProtection.SelectLockedCells;
end;

function TXLSSheetProtection.GetSelectUnlockedCells: boolean;
begin
  Result := FProtection.SelectUnlockedCells;
end;

function TXLSSheetProtection.GetSheet: boolean;
begin
  Result := FProtection.Sheet;
end;

function TXLSSheetProtection.GetSort: boolean;
begin
  Result := FProtection.Sort;
end;

procedure TXLSSheetProtection.SetAutoFilter(const Value: boolean);
begin
  FProtection.AutoFilter := Value;
end;

procedure TXLSSheetProtection.SetDeleteColumns(const Value: boolean);
begin
  FProtection.DeleteColumns := Value;
end;

procedure TXLSSheetProtection.SetDeleteRows(const Value: boolean);
begin
  FProtection.DeleteRows := Value;
end;

procedure TXLSSheetProtection.SetFormatCells(const Value: boolean);
begin
  FProtection.FormatCells := Value;
end;

procedure TXLSSheetProtection.SetFormatColumns(const Value: boolean);
begin
  FProtection.FormatColumns := Value;
end;

procedure TXLSSheetProtection.SetFormatRows(const Value: boolean);
begin
  FProtection.FormatRows := Value;
end;

procedure TXLSSheetProtection.SetInsertColumns(const Value: boolean);
begin
  FProtection.InsertColumns := Value;
end;

procedure TXLSSheetProtection.SetInsertHyperlinks(const Value: boolean);
begin
  FProtection.InsertHyperlinks := Value;
end;

procedure TXLSSheetProtection.SetInsertRows(const Value: boolean);
begin
  FProtection.InsertRows := Value;
end;

procedure TXLSSheetProtection.SetObjects(const Value: boolean);
begin
  FProtection.Objects := Value;
end;

procedure TXLSSheetProtection.SetPasswordString(const Value: AxUCString);
begin
  FProtection.PasswordAsString := Value;

  if (Value <> '') and not FAssigned then begin
    FProtection.Sheet := True;
    FProtection.Objects := True;
    FProtection.Scenarios := True;
  end;
end;

procedure TXLSSheetProtection.SetPivotTables(const Value: boolean);
begin
  FProtection.PivotTables := Value;
end;

procedure TXLSSheetProtection.SetScenarios(const Value: boolean);
begin
  FProtection.Scenarios := Value;
end;

procedure TXLSSheetProtection.SetSelectLockedCells(const Value: boolean);
begin
  FProtection.SelectLockedCells := Value;
end;

procedure TXLSSheetProtection.SetSelectUnlockedCells(const Value: boolean);
begin
  FProtection.SelectUnlockedCells := Value;
end;

procedure TXLSSheetProtection.SetSheet(const Value: boolean);
begin
  FProtection.Sheet := Value;
end;

procedure TXLSSheetProtection.SetSort(const Value: boolean);
begin
  FProtection.Sort := Value;
end;

{ TXLSVirtualCellsImpl }

procedure TXLSRelCellsImpl.ApplyCommand;
begin
  FSheet.FOwner.CmdFormat.Apply(FCol1,FRow1,FCol2,FRow2);
end;

procedure TXLSRelCellsImpl.ApplyCommand(const ACol, ARow: integer);
begin
  FSheet.FOwner.CmdFormat.Apply(ACol,ARow);

  FSheet.FOwner.CmdFormat.Clear;
end;

procedure TXLSRelCellsImpl.ApplyCommand(ACol1, ARow1, ACol2, ARow2: integer);
begin
  FSheet.FOwner.CmdFormat.Apply(ACol1,ARow1,ACol2,ARow2);

  FSheet.FOwner.CmdFormat.Clear;
end;

procedure TXLSRelCellsImpl.AutoWidthCol(ACol: integer);
begin
  FSheet.AutoWidthCol(ACol)
end;

procedure TXLSRelCellsImpl.BeginCommand;
begin
  inherited;

  FSheet.FOwner.CmdFormat.BeginEdit(FSheet);
end;

procedure TXLSRelCellsImpl.ClearAll(ACol1, ARow1, ACol2, ARow2: integer);
begin
  FSheet.DeleteCells(ACol1, ARow1, ACol2, ARow2);
end;

function TXLSRelCellsImpl.Clone: TXLSRelCells;
begin
  Result := TXLSRelCellsImpl.Create(FSheet);
end;

constructor TXLSRelCellsImpl.Create(ASheet: TXLSWorksheet);
begin
  inherited Create;

  FSheet := ASheet;
  if ASheet <> Nil then
    FCommand := FSheet.FOwner.CmdFormat.Commands;
end;

function TXLSRelCellsImpl.GetAsBlank(ACol, ARow: integer): boolean;
begin
  if FSheet <> Nil then
    Result := FSheet.GetAsBlank(ACol,ARow)
  else
    Result := False;
end;

function TXLSRelCellsImpl.GetAsBoolean(ACol, ARow: integer): boolean;
begin
  if FSheet <> Nil then
    Result := FSheet.GetAsBoolean(ACol,ARow)
  else
    Result := False;
end;

function TXLSRelCellsImpl.GetAsEmpty(ACol, ARow: integer): boolean;
begin
  if FSheet <> Nil then
    Result := FSheet.GetAsEmpty(ACol,ARow)
  else
    Result := False;
end;

function TXLSRelCellsImpl.GetAsError(ACol, ARow: integer): TXc12CellError;
begin
  if FSheet <> Nil then
    Result := FSheet.GetAsError(ACol,ARow)
  else
    Result := errUnknown;
end;

function TXLSRelCellsImpl.GetAsFloat(ACol, ARow: integer): double;
begin
  if FSheet <> Nil then
    Result := FSheet.GetAsFloat(ACol,ARow)
  else
    Result := 0;
end;

function TXLSRelCellsImpl.GetAsFmtString(ACol, ARow: integer): AxUCString;
begin
  if FSheet <> Nil then
    Result := FSheet.GetAsFmtString(ACol,ARow)
  else
    Result := '';
end;

function TXLSRelCellsImpl.GetAsString(ACol, ARow: integer): AxUCString;
begin
  if FSheet <> Nil then
    Result := FSheet.GetAsString(ACol,ARow)
  else
    Result := '';
end;

function TXLSRelCellsImpl.GetCellType(ACol, ARow: integer): TXLSCellType;
var
  C: TXLSCellItem;
begin
  Result := xctNone;

  if FSheet <> Nil then begin
    C := FSheet.Xc12Sheet.Cells.FindCell(ACol,ARow);
    if C.Data <> Nil then
      Result := FSheet.Xc12Sheet.Cells.CellType(@C);
  end;
end;

function TXLSRelCellsImpl.GetIsDateTime(Col, Row: integer): boolean;
begin
  if FSheet <> Nil then
    Result := FSheet.GetIsDateTime(Col,Row)
  else
    Result := False;
end;

function TXLSRelCellsImpl.GetAsSharedItem(ACol, ARow: integer): TXLSSharedItemsValue;
var
  Cell: TXLSCellItem;
begin
  if (FSheet <> Nil) and FSheet.FCells.FindCell(ACol,ARow,Cell) then begin
    case FSheet.FCells.CellType(@Cell) of
      xctBlank: Result := TXLSSharedItemsValueBlank.Create;
      xctBoolean,
      xctBooleanFormula: begin
        Result := TXLSSharedItemsValueBoolean.Create;
        Result.AsBoolean := FSheet.FCells.GetBoolean(@Cell);
      end;
      xctError,
      xctErrorFormula: begin
        Result := TXLSSharedItemsValueError.Create;
        Result.AsError := FSheet.FCells.GetError(@Cell);
      end;
      xctString,
      xctStringFormula: begin
        Result := TXLSSharedItemsValueString.Create;
        Result.AsString := FSheet.FCells.GetString(@Cell);
      end;
      xctFloat,
      xctFloatFormula: begin
        if FSheet.IsDateTime[ACol,ARow] then begin
          Result := TXLSSharedItemsValueDate.Create;
          Result.AsDate := FSheet.FCells.GetFloat(@Cell);
        end
        else begin
          Result := TXLSSharedItemsValueNumeric.Create;
          Result.AsNumeric := FSheet.FCells.GetFloat(@Cell);
        end;
      end
      else Result := Nil;
    end;
  end
  else
    Result := Nil;
end;

procedure TXLSRelCellsImpl.SetAsBlank(ACol, ARow: integer; const Value: boolean);
begin
  if FAutoExtend then
    DoAutoExtend(ACol,ARow);

  FSheet.SetAsBlank(ACol,ARow,Value);
end;

procedure TXLSRelCellsImpl.SetAsBoolean(ACol, ARow: integer; const AValue: boolean);
begin
  if FAutoExtend then
    DoAutoExtend(ACol,ARow);

  FSheet.SetAsBoolean(ACol,ARow,AValue);
end;

procedure TXLSRelCellsImpl.SetAsError(ACol, ARow: integer; const AValue: TXc12CellError);
begin
  if FAutoExtend then
    DoAutoExtend(ACol,ARow);

  FSheet.SetAsError(ACol,ARow,AValue);
end;

procedure TXLSRelCellsImpl.SetAsFloat(ACol, ARow: integer; const AValue: double);
begin
  if FAutoExtend then
    DoAutoExtend(ACol,ARow);

  FSheet.SetAsFloat(ACol,ARow,AValue);
end;

procedure TXLSRelCellsImpl.SetAsSharedItem(ACol, ARow: integer; Value: TXLSSharedItemsValue);
begin
  if FAutoExtend then
    DoAutoExtend(ACol,ARow);

  case Value.Type_ of
    xsivtBoolean: FSheet.SetAsBoolean(ACol,ARow,Value.AsBoolean);
    xsivtDate   : FSheet.SetAsDateTime(ACol,ARow,Value.AsDate);
    xsivtError  : FSheet.SetAsError(ACol,ARow,Value.AsError);
    xsivtBlank  : ;
    xsivtNumeric: FSheet.SetAsFloat(ACol,ARow,Value.AsNumeric);
    xsivtString : FSheet.SetAsString(ACol,ARow,Value.AsString);
  end;
end;

procedure TXLSRelCellsImpl.SetAsString(ACol, ARow: integer;  const AValue: AxUCString);
begin
  if FAutoExtend then
    DoAutoExtend(ACol,ARow);

  FSheet.SetAsString(ACol,ARow,AValue);
end;

function TXLSRelCellsImpl.SheetIndex: integer;
begin
  Result := FSheet.Index;
end;

function TXLSRelCellsImpl.SheetName: AxUCString;
begin
  if FSheet <> Nil then
    Result := FSheet.Name
  else
    Result := '???';
end;

end.
