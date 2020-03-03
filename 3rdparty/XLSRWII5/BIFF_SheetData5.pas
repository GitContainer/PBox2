unit BIFF_SheetData5;

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

uses SysUtils, Classes,
{$ifdef BABOON}
{$else}
     Windows, vcl.Graphics,
{$endif}
     Contnrs, Math,
{$ifdef DELPHI_XE2_OR_LATER}
     System.UITypes,
{$endif}
     BIFF_Utils5,BIFF_RecsII5, BIFF_Stream5, BIFF_EncodeFormulaII5, BIFF_Validate5,
     BIFF_RecordStorage5,BIFF_Escher5, BIFF_EscherTypes5,
     BIFF_DrawingObj5, BIFF_ControlObj5, BIFF_FormulaHandler5, BIFF_DrawingObjChart5,
     BIFF_MergedCells5, BIFF_CondFmt5, BIFF_CellAreas5, BIFF_DecodeFormula5,
     BIFF_Autofilter5, BIFF_EscherCtrlObj5, BIFF_WideStrList5,
     Xc12Utils5, Xc12Manager5,
     XLSUtils5;

//* Pane type when a worksheet has splitted panes.
type TPaneType = (ptNone,    //* No panes.
                  ptFrozen,  //* Frozen panes,
                  ptSplit    //* Splitted panes.
                  );

//* Hide state for worksheets.
type THiddenState = (hsVisible,   //* The worksheet is visible.
                     hsHidden,    //* Worksheet is hidden.
                     hsVeryHidden //* Worksheet is hidden, and it's data can only be accessed from VBA (macro) procedures.
                     );
//* Type of worksheet.
type TWorksheetType = (wtSheet,       //* Normal worksheet.
                       wtVBModule,    //* Visual Basic module.
                       wtExcel4Macro  //* Excel 4 macro sheet.
                       );

//* Specifies which actions that can be performed on a protected worksheet.
type TSheetProtection = (spEditObjects,          //* Edit objects
                         spEditScenarios,        //* Edit scenarios
                         spEditCellFormatting,   //* Change cell formattingm
                         spEditColumnFormatting, //* Change column formatting
                         spEditRowFormatting,    //* Change row formatting
                         spInsertColumns,        //* Insert columns
                         spInsertRows,           //* Insert rows
                         spInsertHyperlinks,     //* Insert hyperlinks
                         spDeleteColumns,        //* Delete columns
                         spDeleteRows,           //* Delete rows
                         spSelectLockedCells,    //* Select locked cells
                         spSortCellRange,        //* Sort cell range
                         spEditAutoFileters,     //* Edit auto fileters
                         spEditPivotTables,      //* Edit pivot tables
                         spSelectUnlockedCells   //* Select unlocked cell
                         );

type TSheetProtections = set of TSheetProtection;
//* Default worksheet protections.
const DefaultSheetProtections = [spEditObjects,spEditScenarios,spEditCellFormatting,
                         spEditColumnFormatting,spEditRowFormatting,spInsertColumns,
                         spInsertRows,spInsertHyperlinks,spDeleteColumns,spDeleteRows,
                         spSelectLockedCells,spSortCellRange,spEditAutoFileters,
                         spEditPivotTables,spSelectUnlockedCells];

type TSheet = class;

     //* ~exclude
     TSelectedAreaHit = (sahNo,sahInside,sahEdge,sahActiveCell);
     //* ~exclude
     TSelectedEdgeHit = (sehLeft,sehTop,sehRight,sehBottom);
     //* ~exclude
     TSelectedEdgeHits = set of TSelectedEdgeHit;

//* A selected area is an area that is a part of the current selection.
//* The current selection may contain a single cell (the cell with the cursor),
//* or multiple areas.
     TSelectedArea = class(TCellArea97)
private
public
     destructor Destroy; override;
     //* Normalizes the selection, that is, checks that Col1 <= Col2 and Row1 <= Row2
     procedure Normalize;
     //* Tests if Col and Row is inside the area.
     //* ~param Col Column of the test cell.
     //* ~param Row Row of the test cell.
     //* ~result True if the cell is within the area.
     function  Hit(Col,Row: integer): boolean;
     //* ~exclude
     function  AsRect: TRecCellAreaI;
     //* Sets the dimension of the area.
     //* ~param C1 Left column.
     //* ~param R1 Top row.
     //* ~param C2 Right column.
     //* ~param R2 Bottom row.
     procedure SetArea(C1,R1,C2,R2: integer);
     //* Sets the current area to the intersection of the current area and C1,R1,C2,R2.
     //* ~param C1 Left column.
     //* ~param R1 Top row.
     //* ~param C2 Right column.
     //* ~param R2 Bottom row.
     procedure Intersect(C1,R1,C2,R2: integer);
     end;

//* List of selected areas. The first element in the list (index zero)
//* is created automatically and is always available.
     TSelectedAreas = class(TCellAreas97)
private
     FActiveCol,FActiveRow: integer;
     FActiveArea: integer;

     function  GetItems(Index: integer): TSelectedArea;
     procedure SetActiveCol(const Value: integer);
     procedure SetActiveRow(const Value: integer);
public
     //* ~exclude
     constructor Create;
     //* ~exclude
     destructor Destroy; override;
     //* ~exclude
     procedure Clear; override;
     //* ~exclude
     procedure Init(Col: integer = 0; Row: integer = 0); overload;
     //* ~exclude
     procedure Init(C1,R1,C2,R2,ActC,ActR: integer); overload;
     //* ~exclude
     procedure Move(Col,Row: integer); override;
     //* Adds a new area.
     //* ~result The new TSelectedArea object.
     function  Add: TSelectedArea; overload;
     //* Adds a new area and initiates it to C1,R1,C2,R2.
     //* ~param C1 Left column.
     //* ~param R1 Top row.
     //* ~param C2 Right column.
     //* ~param R2 Bottom row.
     //* ~result The new TSelectedArea object.
     function  Add(C1,R1,C2,R2: integer): TSelectedArea; overload;
     //* ~exclude
     function  CellInAreas(Col,Row: integer; var EdgeHit: TSelectedEdgeHits; var AreaHit: TSelectedAreaHit): integer;
     //* ~exclude
     function  CursorVisible: boolean;
     //* ~exclude
     function  First: TSelectedArea;
     //* ~exclude
     function  Last: TSelectedArea;

     //* ActiveCol is the column where the cursor is. See also ~[link ActiveRow]
     property ActiveCol: integer read FActiveCol write SetActiveCol;
     //* ActiveRow is the row where the cursor is. See also ~[link ActiveCol]
     property ActiveRow: integer read FActiveRow write SetActiveRow;
     //* Returns an index to the area where the cursor is in.
     property ActiveArea: integer read FActiveArea write FActiveArea;
     //* The areas in the list.
     property Items[Index: integer]: TSelectedArea read GetItems; default;
     end;

//* TPane stores settings for window panes.
     TPane = class(TPersistent)
private
     FPaneType: TPaneType;
     FSplitColX: integer;
     FSplitRowY: integer;
     FLeftCol: integer;
     FTopRow: integer;
     FActivePane: byte;
     FSelections: TBaseRecordStorage;

     procedure SetActivePane(const Value: byte);
public
     //* ~exclude
     constructor Create;
     //* ~exclude
     destructor Destroy; override;
     //* ~exclude
     procedure Clear;
     //* ~exclude
     property Selections: TBaseRecordStorage read FSelections;
     //* Which pane that is active, i.e the pane with the cursor.
     property ActivePane: byte read FActivePane write SetActivePane;
published
     //* Pane style.
     property PaneType: TPaneType read FPaneType write FPaneType;
     //* Vertical split.
     //* Split pane (PaneType = ptSplit): width of left pane in units of 1/20
     //* of a point.~[br]
     //* Frozen pane (PaneType = ptFrozen): Number of visible columns in left pane.~[br]
     //* If there not shall be a vertical split, set SplitColX to zero.
     property SplitColX: integer read FSplitColX write FSplitColX;
     //* Horizontal split.
     //* Split pane (PaneType = ptSplit): height of top pane in units of 1/20
     //* of a point.~[br]
     //* Frozen pane (PaneType = ptFrozen): Number of visible rows in top pane.
     //* If there not shall be a horizontal split, set SplitColY to zero.
     property SplitRowY: integer read FSplitRowY write FSplitRowY;
     //* First visible column in right pane.
     property LeftCol: integer read FLeftCol write FLeftCol;
     //* First visible row in bottom pane.
     property TopRow: integer read FTopRow write FTopRow;
     end;

     //* Options for the worksheet.
     TSheetOption = (soGridlines,      //* Show gridlines.
                     soRowColHeadings, //* Show row and column headings.
                     soProtected,      //* Worksheet is protected.
                     soR1C1Mode,       //* Show cells in R1C1 mode.
                     soIteration,
                     soShowFormulas,   //* Show formulas instead of formula result.
                     soFrozenPanes,    //* Worksheet has frozen panes.
                     soShowZeros      //* Show zero values.
                     );
     TSheetOptions = set of TSheetOption;

     TWorkspaceOption = (woShowAutoBreaks, //* Automatic page breaks are visible.
                         woApplyStyles,    //* Automatic styles are applied to an outline
                         woRowSumsBelow,   //* Summary rows appear below detail in an outline
                         woColSumsRight,   //* Summary columns appear to the right of detail in an outline
                         woFitToPage,      //* Page fit option is on.
                         woOutlineSymbols  //* Outline symbols are displayed.
                         );
     TWorkspaceOptions = set of TWorkspaceOption;

//* Base class for worksheets.
     TBasicSheet = class(TCollectionItem)
private
     function GetIndex: integer;
protected
     FName: AxUCString;
     FSheetIndex: integer;

     procedure SetName(Value: AxUCString);
     procedure SetNameNoCheck(Value: AxUCString);
     function  GetName: AxUCString;
public
     //* ~exclude
     constructor Create(Collection: TCollection); override;
     //* ~exclude
     destructor Destroy; override;

     //* Index of the worksheet.
     property Index: integer read GetIndex;
published
     //* The name of the worksheet.
     property Name: AxUCString read GetName write SetName;
     end;

     TSheets = class;

//* Normal worksheet.
     TSheet = class(TBasicSheet)
private
     FRecords: TRecordStorageSheet;
     FHasDefaultRecords: boolean;

     FRecalcFormulas: boolean;
     FMergedCells: TMergedCells;
     FValidations: TDataValidations;
     FPane: TPane;
     FDrawingObjects: TDrawingObjects;
     FControlObjects: TControlObjects;
     FCharts: TDrwCharts;
     FEscherDrawing: TEscherDrawing;
     FConditionalFormats: TConditionalFormats;
     FAutofilters: TAutofilters;
     FSelectedAreas: TSelectedAreas;
     FHidden: THiddenState;
     FSheetProtections: TSheetProtections;
     FWorksheetType: TWorksheetType;
     FTabColor: TXc12IndexColor;
     FDimension: TRec32CellArea;
     FVisibleAreas: TCellAreas97;
     FCachedMSORecs: TBaseRecordStorage;

     function  GetWorkspaceOptions: TWorkspaceOptions;
     procedure SetWorkspaceOptions(const Value: TWorkspaceOptions);
     function  GetDefaultColWidth: word;
     procedure SetDefaultColWidth(const Value: word);
     function  GetFirstCol: longword;
     function  GetFirstRow: longword;
     function  GetLastCol: longword;
     function  GetLastRow: longword;
     procedure SetFirstCol(const Value: longword);
     procedure SetFirstRow(const Value: longword);
     procedure SetLastCol(const Value: longword);
     procedure SetLastRow(const Value: longword);
     function  GetOptions: TSheetOptions;
     function  GetZoom: word;
     function  GetZoomPreview: word;
     procedure SetOptions(const Value: TSheetOptions);
     procedure SetZoom(const Value: word);
     procedure SetZoomPreview(const Value: word);
     function  GetLeftCol: integer;
     function  GetTopRow: integer;
     procedure SetLeftCol(const Value: integer);
     procedure SetTopRow(const Value: integer);
protected
     function  GetDisplayName: string; override;
     procedure WriteBuf(Stream: TXLSStream; RecId,Size: word; P: Pointer);
     procedure CheckFirstLast(ACol,ARow: integer);
     function  GetDefaultWriteFormat(Version: TExcelVersion; FormatIndex: integer): word;

     procedure OnEscherReadShape(Sender: TObject; Shape: TShape);
public
     //* ~exclude
     constructor Create(Collection: TCollection); override;
     //* ~exclude
     destructor Destroy; override;
     //* Clears (empties) the worksheet of all data.
     procedure ClearData;
     // **********************************************
     // *********** For internal use only. ***********
     // **********************************************

     //* ~exclude
     procedure NameIndexChanged(Delta: integer);
     //* ~exclude
     procedure AfterFileRead;

     procedure SetRightToLeft(AValue: boolean);

     property  _Int_CachedMSORecs: TBaseRecordStorage read FCachedMSORecs;
     property  _Int_Records: TRecordStorageSheet read FRecords;
     //* ~exclude
     property  _Int_EscherDrawing: TEscherDrawing read FEscherDrawing;
     //* ~exclude
     property _Int_HasDefaultRecords: boolean read FHasDefaultRecords write FHasDefaultRecords;
     //* ~exclude
     property _Int_SheetIndex: integer read FSheetIndex write FSheetIndex;
     // **********************************************
     // *********** End internal use only. ***********
     // **********************************************
     //* Clears the worksheet of all cell data.
     procedure ClearWorksheet;
     procedure Assign(Source: TPersistent); override;
     function  AddChart: TDrwChart;

     //* The first column of the worksheet with a cell value.
     property FirstCol: longword read GetFirstCol write SetFirstCol;
     //* The last column of the worksheet with a cell value.
     property LastCol: longword read GetLastCol write SetLastCol;
     //* The first row of the worksheet with a cell value.
     property FirstRow: longword read GetFirstRow write SetFirstRow;
     //* The last row of the worksheet with a cell value.
     property LastRow: longword read GetLastRow write SetLastRow;
     //* The first column in the upper left corner.
     property LeftCol: integer read GetLeftCol write SetLeftCol;
     //* The first row in the upper left corner.
     property TopRow: integer read GetTopRow write SetTopRow;
     //* List of selected areas.
     property SelectedAreas: TSelectedAreas read FSelectedAreas;
     //* List of charts in the worksheet.
     property Charts: TDrwCharts read FCharts write FCharts;
     //* Type of the worksheet. The normal is wtSheet. When files are read,
     //* the type can be also be wtVBModule,wtExcel4Macro. Do not make any
     //* changes to these sheets.
     property WorksheetType: TWorksheetType read FWorksheetType;
//     property Count: integer read FCellCount;
     //* Autofilters on the worksheet.
     property Autofilters: TAutofilters read FAutofilters write FAutofilters;
     //* Color of the workseet's tab.
     property TabColor: TXc12IndexColor read FTabColor write FTabColor;
     //* Used by TXLSSpreadSheet.
     //* Defining areas in VisibleAreas limits the visibility of the cells on
     //* the worksheet to this areas. When VisibleAreas is empty, the entire
     //* worksheet is visible.
     property VisibleAreas: TCellAreas97 read FVisibleAreas write FVisibleAreas;
published
     //* Default column width.
     //* The width is in units of 1/256s of a character width.
     property DefaultColWidth: word read GetDefaultColWidth write SetDefaultColWidth;
     //* List of merged cells.
     property MergedCells: TMergedCells read FMergedCells write FMergedCells;
     //* Worksheet options.
     property Options: TSheetOptions read GetOptions write SetOptions;
     //* Workspace options.
     property WorkspaceOptions: TWorkspaceOptions read GetWorkspaceOptions write SetWorkspaceOptions;
     //* Protection for the worksheet.
     property SheetProtection: TSheetProtections read FSheetProtections write FSheetProtections;
     //* Zoom magnification.
     property Zoom: word read GetZoom write SetZoom;
     //* Zoom magnification in page break preview.
     property ZoomPreview: word read GetZoomPreview write SetZoomPreview;
     //* If formulas shall be recalced when the file is opened by Excel.
     //* Set this property to true if formulas shall be marked for recalculation.
     //* As TXLSReadWriteII2 don't calculate formulas, this is advised. Default is True.
     property RecalcFormulas: boolean read FRecalcFormulas write FRecalcFormulas;
     //* True if the worksheet is hidden.
     property Hidden: THiddenState read FHidden write FHidden;
     //* List with data validations.
     property Validations: TDataValidations read FValidations write FValidations;
     //* List with drawing objects.
     property DrawingObjects: TDrawingObjects read FDrawingObjects write FDrawingObjects;
     //* List with control objects.
     property ControlsObjects: TControlObjects read FControlObjects write FControlObjects;
     //* Pane settings for the worksheet.
     property Pane: TPane read FPane;
     //* List with conditional formats.
     property ConditionalFormats: TConditionalFormats read FConditionalFormats write FConditionalFormats;
     end;

//* List with TSheet objects.
     TSheets = class(TCollection)
private
     FFormulaHandler: TFormulaHandler;

     function  GetSheet(Index: integer): TSheet;
protected
     FOwner: TPersistent;
     FManager: TXc12Manager;

     function  GetOwner: TPersistent; override;
     procedure Reindex;
public
     //* ~exclude
     constructor Create(AOwner: TPersistent; FormulaHandler: TFormulaHandler; AManager: TXc12Manager);
     //* ~exclude
     destructor Destroy; override;
     //* Inserts a new TSheet at position Index.
     //* ~param Index The position to insert the sheet.
     //* ~result The new TSheet object.
     function Insert(Index: Integer): TSheet;
     //* Delete the sheet at Index.
     //* ~param Index Index of the sheet to be deleted.
     procedure Delete(Index: integer; const AAddSheet: boolean = True);
     //* Deletes all sheets.
     //* When deleting all sheets, a new, empty sheet is added automatically,
     //* as there has to be at least one sheet.
     procedure Clear;
     //* Add a new sheet.
     //* ~param WorksheetType Is always wtSheet.
     //* ~result The new TSheet object.
     function  Add(WorksheetType: TWorksheetType = wtSheet): TSheet;
     //* Searches for a worksheet by it's name.
     //* ~param Name The name of the worksheet.
     //* ~result The TSheet object. If not found, Nil is returned.
     function  SheetByName(Name: AxUCString): TSheet;
     function  SplitSheetCellRef(Ref: AxUCString; out SheetIndex,Col1,Row1,Col2,Row2: integer): boolean;
     // **********************************************
     // *********** For internal use only. ***********
     // **********************************************
     procedure ClearAll;
     //* ~exclude
     property FormulaHandler: TFormulaHandler read FFormulaHandler;
     //* ~exclude
//     property  SST: TSST2 read FSST;
     // **********************************************
     // *********** End internal use only. ***********
     // **********************************************

     //* The sheets in the list.
     property  Items[Index: integer]: TSheet read GetSheet; default;
     end;

implementation

uses BIFF5,  BIFF_ReadII5;

{ TSheets }

constructor TSheets.Create(AOwner: TPersistent; FormulaHandler: TFormulaHandler; AManager: TXc12Manager);
begin
  inherited Create(TSheet);
  FOwner := AOwner;
  FFormulaHandler := FormulaHandler;
  FManager := AManager;
end;

destructor TSheets.Destroy;
begin
  inherited;
end;

function TSheets.Insert(Index: Integer): TSheet;
begin
  Result := TSheet(inherited Insert(Index));
  TBIFF5(FOwner).FormulaHandler.InsertSheet(Index);
  Reindex;
end;

procedure TSheets.Reindex;
var
  i: integer;
begin
  for i := 0 to Count - 1 do
    Items[i].FSheetIndex := i;
end;

procedure TSheets.Delete(Index: integer; const AAddSheet: boolean = True);
begin
{$ifndef ver120}
  inherited Delete(Index);
{$endif}
  if (Count < 1) and AAddSheet then
    Add
  else
    TBIFF5(FOwner).FormulaHandler.DeleteSheet(Index);
  Reindex;
end;
                           
procedure TSheets.ClearAll;
var
  i: integer;
begin
  for i := 0 to Count - 1 do begin
    Items[i].ClearData;
  end;
  inherited Clear;
end;

procedure TSheets.Clear;
begin
  ClearAll;
  Add;
end;

function TSheets.GetSheet(Index: integer): TSheet;
begin
  Result := TSheet(inherited Items[Index]);
end;

function TSheets.GetOwner: TPersistent;
begin
  Result := FOwner;
end;

function TSheets.SheetByName(Name: AxUCString): TSheet;
var
  i: integer;
begin
  for i := 0 to Count - 1 do begin
{$ifdef ver130}
    if AnsiCompareText(Items[i].Name,Name) = 0 then begin
{$else}
    if WideCompareText(Items[i].Name,Name) = 0 then begin
{$endif}
      Result := Items[i];
      Exit;
    end;
  end;
  Result := Nil;
end;

function TSheets.SplitSheetCellRef(Ref: AxUCString; out SheetIndex, Col1, Row1, Col2, Row2: integer): boolean;
var
  p: integer;
  S: AxUCString;
  Sht: TSheet;
  Tmp: boolean;
begin
  Result := False;
  p := CPos('!',Ref);
  if p >= 1 then begin
    S := Copy(Ref,1,p - 1);
    Sht := SheetByName(S);
    if Sht = Nil then
      Exit;
    SheetIndex := Sht.Index;
    S := Copy(Ref,p + 1,MAXINT);
  end
  else begin
    SheetIndex := -1;
    S := Ref;
  end;
  if CPos(':',S) > 1 then
    Result := AreaStrToColRow(S,Col1,Row1,Col2,Row2)
  else begin
    Result := RefStrToColRow(S,Col1,Row1,Tmp,Tmp);
    Col2 := Col1;
    Row2 := Row1;
  end;
end;

function TSheets.Add(WorksheetType: TWorksheetType = wtSheet): TSheet;
begin
//  FFormulaHandler.Cells.Add;
  Result := TSheet(inherited Add);
  Result.FWorksheetType := WorksheetType;
end;

{ TSheet }

constructor TSheet.Create(Collection: TCollection);
begin
  inherited Create(Collection);
  FTabColor := xcAutomatic;
  FSheetProtections := DefaultSheetProtections;
  FRecords := TRecordStorageSheet.Create;
  FRecords.SetDefaultData;
  FCachedMSORecs := TBaseRecordStorage.Create;
//  FCells := TSheets(Collection).FormulaHandler.Cells[Index];
  FRecalcFormulas := True;
  FMergedCells := TMergedCells.Create;
  FValidations := TDataValidations.Create(Self,TBIFF5(TSheets(Collection).FOwner).FormulaHandler);
  with TBIFF5(TSheets(Collection).FOwner) do
    FEscherDrawing := TEscherDrawing.Create(MSOPictures,Fonts,InternalNames);
  FEscherDrawing.OnReadShape := OnEscherReadShape;
  with TBIFF5(TSheets(Collection).FOwner) do
    FDrawingObjects := TDrawingObjects.Create(Self,FEscherDrawing,FormulaHandler,TSheets(Collection).FManager.StyleSheet.Fonts);
  FControlObjects := TControlObjects.Create(Self,FEscherDrawing,TBIFF5(TSheets(Collection).FOwner).FormulaHandler);
  FCharts := TDrwCharts.Create(Self,FEscherDrawing,TBIFF5(TSheets(Collection).FOwner).FormulaHandler,TBIFF5(TSheets(Collection).FOwner).Fonts);
  FPane := TPane.Create;
  FConditionalFormats := TConditionalFormats.Create(Self,TBIFF5(TSheets(Collection).FOwner).FormulaHandler);
  FAutofilters := TAutofilters.Create(Self,FControlObjects.ComboBoxes,TBIFF5(TSheets(Collection).FOwner).InternalNames);
  FSelectedAreas := TSelectedAreas.Create;
  FVisibleAreas := TCellAreas97.Create;
  SetNameNoCheck('Sheet' + IntToStr(ID + 1));
end;

destructor TSheet.Destroy;
begin
  FVisibleAreas.Free;
  FValidations.Free;
  FMergedCells.Free;
  FRecords.Free;
  FCachedMSORecs.Free;
  FDrawingObjects.Free;
  FControlObjects.Free;
  FCharts.Free;
  FPane.Free;
  FConditionalFormats.Free;
  FAutofilters.Free;
  FEscherDrawing.Free;
  FSelectedAreas.Free;

//  Deleting FormulaHandler.Cells is done in TSheets.Delete
//  if TSheets(Collection).FormulaHandler.Cells.Count > 0 then
//    TSheets(Collection).FormulaHandler.Cells.Delete(Index);
  inherited Destroy;
end;

procedure TSheet.ClearData;
begin
  FTabColor := xcAutomatic;
  FSheetProtections := DefaultSheetProtections;
  FRecords.Clear;
  FRecords.SetDefaultData;
  FCachedMSORecs.Clear;
  FHasDefaultRecords := False;
  FValidations.Clear;
  FMergedCells.Clear;
  FDrawingObjects.Clear;
  FControlObjects.Clear;
  FEscherDrawing.Clear;
  FPane.Clear;
  FConditionalFormats.Clear;
  FAutofilters.Clear;
  FSelectedAreas.Clear;
  FVisibleAreas.Clear;
  FDimension.Col1 := 0;
  FDimension.Row1 := 0;
  FDimension.Col2 := 0;
  FDimension.Row2 := 0;
//  FCells.Clear;
end;

procedure TSheet.CheckFirstLast(ACol,ARow: integer);
begin
  if Longword(ACol) < FirstCol then FirstCol := Longword(ACol);
  if Longword(ACol) > LastCol then LastCol := Longword(ACol);
  if Longword(ARow) < FirstRow then FirstRow := Longword(ARow);
  if Longword(ARow) > LastRow then LastRow := Longword(ARow);
end;

function TSheet.GetDisplayName: string;
begin
  inherited GetDisplayName;
  Result := GetName;
end;

procedure TSheet.WriteBuf(Stream: TXLSStream; RecId,Size: word; P: Pointer);
begin
  Stream.WriteHeader(RecID,Size);
  if Size > 0 then
    Stream.Write(P^,Size);
end;

// Check this...
function TSheet.GetDefaultWriteFormat(Version: TExcelVersion; FormatIndex: integer): word;
begin
  if FormatIndex < 0 then
    Result := XLS_STYLE_DEFAULT_XF_97
  else if TBIFF5(TSheets(Collection).FOwner).WriteDefaultData then begin
    if Version < xvExcel50 then
      Result := FormatIndex + DEFAULT_FORMAT40 + 1
    else if Version < xvExcel97 then
      Result := FormatIndex + DEFAULT_FORMATS_COUNT_50
    else
      Result := FormatIndex;
  end
  else
    Result := FormatIndex;
end;

procedure TSheet.OnEscherReadShape(Sender: TObject; Shape: TShape);
begin
  if Shape is TShapeClientAnchor then begin
    case Shape.ShapeType of
      msosptTextBox: begin
        if Shape.ExShape is TShapeOutsideMsoNote then
          FDrawingObjects.Notes.AddFromFile(TShapeClientAnchor(Shape))
        else
          FDrawingObjects.Texts.AddFromFile(TShapeClientAnchor(Shape));
      end;
      msosptHostControl: begin
        if Shape.ExShape <> Nil then begin
          if Shape.ExShape is TShapeControlComboBox then
            FControlObjects.ComboBoxes.AddFromFile(TShapeClientAnchor(Shape))
          else if Shape.ExShape is TShapeControlListBox then
            FControlObjects.ListBoxes.AddFromFile(TShapeClientAnchor(Shape))
          else if Shape.ExShape is TShapeControlButton then
            FControlObjects.Buttons.AddFromFile(TShapeClientAnchor(Shape))
          else if Shape.ExShape is TShapeControlCheckbox then
            FControlObjects.Checkboxes.AddFromFile(TShapeClientAnchor(Shape))
          else if Shape.ExShape is TShapeOutsideMsoChart then
            FCharts.AddFromFile(TShapeClientAnchor(Shape));
        end;
      end;
      msosptPictureFrame: begin
        FDrawingObjects.Pictures.AddFromFile(TShapeClientAnchor(Shape));
      end;
    end;
  end;
end;

function TSheet.AddChart: TDrwChart;
begin
  Result := FCharts.Add(Index);
end;

procedure TSheet.ClearWorksheet;
begin

end;

procedure TSheet.Assign(Source: TPersistent);
begin
  if not (Source is TSheet) then
    raise XLSRWException.Create('Object is not TSheet');
  DefaultColWidth := TSheet(Source).DefaultColWidth;
  Options := TSheet(Source).Options;
  WorkspaceOptions := TSheet(Source).WorkspaceOptions;
  SheetProtection := TSheet(Source).SheetProtection;
  Zoom := TSheet(Source).Zoom;
  ZoomPreview := TSheet(Source).ZoomPreview;
  RecalcFormulas := TSheet(Source).RecalcFormulas;
  Hidden := TSheet(Source).Hidden;

  FMergedCells.Assign(TSheet(Source).FMergedCells);
end;

function TSheet.GetLeftCol: integer;
begin
  Result := FRecords.WINDOW2.LeftCol;
end;

function TSheet.GetTopRow: integer;
begin
  Result := FRecords.WINDOW2.TopRow;
end;

procedure TSheet.SetLeftCol(const Value: integer);
begin
  if (Value >= 0) and (Value <= MAXCOL) then
    FRecords.WINDOW2.LeftCol := Value;
end;

procedure TSheet.SetTopRow(const Value: integer);
begin
  if (Value >= 0) and (Value <= MAXROW) then
    FRecords.WINDOW2.TopRow := Value;
end;

procedure TSheet.NameIndexChanged(Delta: integer);
var
  i: integer;
begin
  for i := 0 to FEscherDrawing.ShapeCount - 1 do begin
    if FEscherDrawing.Group[i].ExShape <> Nil then begin
      if FEscherDrawing.Group[i].ExShape is TShapeControlButton then
        TShapeControlButton(FEscherDrawing.Group[i].ExShape).NameIndexChanged(Delta)
      else if FEscherDrawing.Group[i].ExShape is TShapeControlListBox then
        TShapeControlListBox(FEscherDrawing.Group[i].ExShape).NameIndexChanged(Delta);
    end;
  end;
end;

{ TBasicSheet }

constructor TBasicSheet.Create(Collection: TCollection);
begin
  inherited;
end;

destructor TBasicSheet.Destroy;
begin
  inherited;
end;

function TBasicSheet.GetIndex: integer;
begin
  Result := inherited Index;
end;

function TBasicSheet.GetName: AxUCString;
begin
  Result := FName;
end;

procedure TBasicSheet.SetName(Value: AxUCString);
//var
//  i: integer;
begin
  if MyWideUppercase(FName) = MyWideUppercase(Value) then
    Exit;
// Skip name check.
//  for i := 0 to Collection.Count - 1 do begin
//    if (Index <> i ) and (MyWideUppercase(TSheets(Collection).Items[i].Name) = MyWideUppercase(Value)) then
//      raise XLSRWException.Create('Sheet name "' + Value + '" allready exists');
//  end;
  FName := Value;
end;

procedure TSheet.AfterFileRead;
begin
  if TBIFF5(TSheets(Collection).FOwner).Version <= xvExcel97 then begin
    FDimension.Col1 := FRecords.DIMENSIONS.FirstCol;
    FDimension.Row1 := FRecords.DIMENSIONS.FirstRow;
    if FRecords.DIMENSIONS.LastCol > 0 then
      FDimension.Col2 := FRecords.DIMENSIONS.LastCol - 1
    else
      FDimension.Col2 := 0;
    if FRecords.DIMENSIONS.LastRow > 0 then
      FDimension.Row2 := FRecords.DIMENSIONS.LastRow - 1
    else
      FDimension.Row2 := 0;
  end;
end;

function TSheet.GetWorkspaceOptions: TWorkspaceOptions;
var
  V: word;
begin
  V := FRecords.WSBOOL;
  Result := [];
  if (V and $0001) = $0001 then
    Result := Result + [woShowAutoBreaks];
  if (V and $0020) = $0020 then
    Result := Result + [woApplyStyles];
  if (V and $0040) = $0040 then
    Result := Result + [woRowSumsBelow];
  if (V and $0080) = $0080 then
    Result := Result + [woColSumsRight];
  if (V and $0100) = $0100 then
    Result := Result + [woFitToPage];
  if (V and $0400) = $0400 then
    Result := Result + [woOutlineSymbols];
end;

procedure TSheet.SetWorkspaceOptions(const Value: TWorkspaceOptions);
var
  V: word;
begin
  V := $0000;
  if woShowAutoBreaks in Value then
    V := V or $0001;
  if woApplyStyles in Value then
    V := V or $0020;
  if woRowSumsBelow in Value then
    V := V or $0040;
  if woColSumsRight in Value then
    V := V or $0080;
  if woFitToPage in Value then
    V := V or $0100;
  if woOutlineSymbols in Value then
    V := V or $0400;
  FRecords.WSBOOL := V;
end;

function TSheet.GetDefaultColWidth: word;
begin
  Result := FRecords.DEFCOLWIDTH;
end;

procedure TSheet.SetDefaultColWidth(const Value: word);
begin
  FRecords.DEFCOLWIDTH := Value;
end;

function TSheet.GetFirstCol: longword;
begin
  Result := FDimension.Col1;
end;

function TSheet.GetFirstRow: longword;
begin
  Result := FDimension.Row1;
end;

function TSheet.GetLastCol: longword;
begin
  Result := FDimension.Col2;
end;

function TSheet.GetLastRow: longword;
begin
  Result := FDimension.Row2;
end;

procedure TSheet.SetFirstCol(const Value: longword);
begin
  FDimension.Col1 := Value;
  if FDimension.Col1 <= MAXCOL_97 then
    FRecords.DIMENSIONS.FirstCol := Value
  else
    FRecords.DIMENSIONS.FirstCol := MAXCOL_97;
end;

procedure TSheet.SetFirstRow(const Value: longword);
begin
  FDimension.Row1 := Value;
  if FDimension.Row1 <= MAXROW_97 then
    FRecords.DIMENSIONS.FirstRow := Value
  else
    FRecords.DIMENSIONS.FirstRow := MAXROW_97;
end;

procedure TSheet.SetLastCol(const Value: longword);
begin
  FDimension.Col2 := Value;
  if FDimension.Col1 <= (MAXCOL_97 - 1) then
    FRecords.DIMENSIONS.LastCol := Value + 1
  else
    FRecords.DIMENSIONS.LastCol := MAXCOL_97;
end;

procedure TSheet.SetLastRow(const Value: longword);
begin
  FDimension.Row2 := Value;
  if FDimension.Row1 <= (MAXROW_97 - 1) then
    FRecords.DIMENSIONS.LastRow := Value + 1
  else
    FRecords.DIMENSIONS.LastRow := MAXROW_97;
end;

function TSheet.GetOptions: TSheetOptions;
begin
  Result := [];
  if (FRecords.WINDOW2.Options and $0001) = $0001 then
    Result := Result + [soShowFormulas];
  if (FRecords.WINDOW2.Options and $0002) = $0002 then
    Result := Result + [soGridlines];
  if (FRecords.WINDOW2.Options and $0004) = $0004 then
    Result := Result + [soRowColHeadings];
  if (FRecords.WINDOW2.Options and $0008) = $0008 then
    Result := Result + [soFrozenPanes];
  if (FRecords.WINDOW2.Options and $0010) = $0010 then
    Result := Result + [soShowZeros];
end;

function TSheet.GetZoom: word;
begin
  Result := FRecords.WINDOW2.Zoom;
end;

function TSheet.GetZoomPreview: word;
begin
  Result := FRecords.WINDOW2.ZoomPreview;
end;

procedure TSheet.SetOptions(const Value: TSheetOptions);
begin
  FRecords.WINDOW2.Options := $06B6 and not ($0001 or $0002 or $0004 or $0008 or $0010 or $0200);

//  FRecords.WINDOW2.Options := $0000;
  if soShowFormulas in Value then
    FRecords.WINDOW2.Options := FRecords.WINDOW2.Options or $0001;
  if soGridlines in Value then
    FRecords.WINDOW2.Options := FRecords.WINDOW2.Options or $0002;
  if soRowColHeadings in Value then
    FRecords.WINDOW2.Options := FRecords.WINDOW2.Options or $0004;
  if soFrozenPanes in Value then
    FRecords.WINDOW2.Options := FRecords.WINDOW2.Options or $0008;
  if soShowZeros in Value then
    FRecords.WINDOW2.Options := FRecords.WINDOW2.Options or $0010;

end;

procedure TSheet.SetRightToLeft(AValue: boolean);
begin
  if AValue then
    FRecords.WINDOW2.Options := FRecords.WINDOW2.Options or $0040
  else
    FRecords.WINDOW2.Options := FRecords.WINDOW2.Options and not $0040
end;

procedure TSheet.SetZoom(const Value: word);
begin
  FRecords.WINDOW2.Zoom := Value;
end;

procedure TSheet.SetZoomPreview(const Value: word);
begin
  FRecords.WINDOW2.ZoomPreview := Value;
end;

{ TBasicSheet }

procedure TBasicSheet.SetNameNoCheck(Value: AxUCString);
begin
  FName := Value;
end;

{ TPane }

procedure TPane.Clear;
begin
  FPaneType := ptNone;
  FSplitColX := 0;
  FSplitRowY := 0;
  FLeftCol := 0;
  FTopRow := 0;
  FActivePane := 3;
  FSelections.Clear;
end;

constructor TPane.Create;
begin
  FSelections := TBaseRecordStorage.Create;
end;

destructor TPane.Destroy;
begin
  FSelections.Free;
  inherited;
end;

procedure TPane.SetActivePane(const Value: byte);
begin
  if Value > 3 then
    raise XLSRWException.Create('Value out of range');
  FActivePane := Value;
end;

{ TSelectedAreas }

function TSelectedAreas.Add: TSelectedArea;
begin
  Result := TSelectedArea.Create;
  inherited Add(Result);
end;

function TSelectedAreas.Add(C1, R1, C2, R2: integer): TSelectedArea;
begin
  Result := TSelectedArea.Create;
  Result.FCol1 := C1;
  Result.FRow1 := R1;
  Result.FCol2 := C2;
  Result.FRow2 := R2;
  FActiveCol := C1;
  FActiveRow := R1;
  inherited Add(Result);
end;

function TSelectedAreas.CellInAreas(Col, Row: integer; var EdgeHit: TSelectedEdgeHits; var AreaHit: TSelectedAreaHit): integer;
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
  if (Col = FActiveCol) and (Row = FActiveRow) then
    AreaHit := sahActiveCell;
  for i := 0 to Count - 1 do begin
    if (Col >= Items[i].FCol1) and (Col <= Items[i].FCol2) and (Row >= Items[i].FRow1) and (Row <= Items[i].FRow2) then begin
      if AreaHit <> sahActiveCell then
        AreaHit := sahEdge;
      if Col <> Items[i].FCol1 then Edges := Edges or E_LEFT;
      if Col <> Items[i].FCol2 then Edges := Edges or E_RIGHT;
      if Row <> Items[i].FRow1 then Edges := Edges or E_TOP;
      if Row <> Items[i].FRow2 then Edges := Edges or E_BOTTOM;
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

procedure TSelectedAreas.Clear;
begin
  inherited Clear;
end;

constructor TSelectedAreas.Create;
begin
  inherited Create;
  Init;
end;

function TSelectedAreas.CursorVisible: boolean;
begin
  Result := Count = 1;
end;

destructor TSelectedAreas.Destroy;
begin
  inherited;
end;

function TSelectedAreas.First: TSelectedArea;
begin
  Result := Items[0];
end;

function TSelectedAreas.GetItems(Index: integer): TSelectedArea;
begin
  Result := TSelectedArea(inherited Items[Index]);
end;

procedure TSelectedAreas.Init(Col: integer = 0; Row: integer = 0);
begin
  with Add do begin
    Col1 := Col;
    Col2 := Col;
    Row1 := Row;
    Row2 := Row;
  end;
  FActiveCol := Col;
  FActiveRow := Row;
end;

procedure TSelectedAreas.Init(C1,R1,C2,R2,ActC,ActR: integer);
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

function TSelectedAreas.Last: TSelectedArea;
begin
  Result := Items[Count - 1];
end;

procedure TSelectedAreas.Move(Col, Row: integer);
begin
  while Count > 1 do
    Delete(Count - 1);
  with Items[0] do begin
    Col1 := Col;
    Col2 := Col;
    Row1 := Row;
    Row2 := Row;
  end;
  FActiveCol := Col;
  FActiveRow := Row;
end;

procedure TSelectedAreas.SetActiveCol(const Value: integer);
begin
//  Clear;
  FActiveCol := Value;
//  Init(FActiveCol,FActiveRow);
end;

procedure TSelectedAreas.SetActiveRow(const Value: integer);
begin
//  Clear;
  FActiveRow := Value;
//  Init(FActiveCol,FActiveRow);
end;

{ TSelectedArea }

function TSelectedArea.AsRect: TRecCellAreaI;
begin
  Result.Col1 := FCol1;
  Result.Row1 := FRow1;
  Result.Col2 := FCol2;
  Result.Row2 := FRow2;
end;

destructor TSelectedArea.Destroy;
begin

  inherited;
end;

function TSelectedArea.Hit(Col, Row: integer): boolean;
begin
  Result := (Col >= FCol1) and (Row >= FRow1) and (Col <= FCol2) and (Row <= FRow2);
end;

procedure TSelectedArea.Intersect(C1, R1, C2, R2: integer);
begin
  if C1 < FCol1 then FCol1 := C1;
  if R1 < FRow1 then FRow1 := R1;
  if C2 > FCol2 then FCol2 := C2;
  if R2 > FRow2 then FRow2 := R2;
end;

procedure TSelectedArea.Normalize;

procedure Swap(var W1,W2: integer);
var
  T: Word;
begin
  T := W1;
  W1 := W2;
  W2 := T;
end;

begin
  if FCol1 > FCol2 then
    Swap(FCol1,FCol2);
  if FRow1 > FRow2 then
    Swap(FRow1,FRow2);
end;

procedure TSelectedArea.SetArea(C1, R1, C2, R2: integer);
begin
  FCol1 := C1;
  FRow1 := R1;
  FCol2 := C2;
  FRow2 := R2;
end;

end.
