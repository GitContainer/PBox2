unit XLSBook2;

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

uses {* Delphi  *} Classes, SysUtils, vcl.Controls, Messages, Windows, Math,
                   vcl.Graphics, Contnrs, vcl.Forms, vcl.ActnList, UxTheme,

     {* XLSRWII *} Xc12DataStyleSheet5, Xc12Manager5, Xc12Utils5, Xc12DataComments5,
                   XLSUtils5, XLSCellAreas5, XLSSheetData5, XLSReadWriteII5,
                   XLSMMU5, XLSCellMMU5, XLSCmdFormat5, XLSEvaluateFmla5,

     (* AXW     *) XSSIEKeys,

     {* XLSBook *} XBookSkin2, XBookSheet2, XBookPaint2, XBookOptions2,
                   XBookUtils2, XBookHintWindow2, XBookPaintLayers2, XBookComponent2,
                   XBookPaintGDI2, XBookSysVar2, XBookWindows2, XBookInplaceEdit2,
                   XBookUI2;

//* Quick start
//*
//* The TXLSSpreadSheet is a viewer and user interface of TXLSReadWriteII. This
//* means that all access to data, such as cells, formatting or anything else
//* stored in the Excel file is trough the XLS or XLSSheet property.
//  For example, if you want to set a cell vale, use a code like this:
//*
//*  // Use the XLSSheet property. This is the visible worksheet.
//*  XLSSpreadSheet.XLSSheet.AsString[2,2] := 'Hello';
//*  // Use the XLS property, sheet 0.
//*  XLSSpreadSheet.XLS[0].AsFloat[2,3] := 125.5;
//*  // Repaint the worksheet.
//*  XLSSpreadSheet.InvalidateSheet;
//*
//* The exception from this is the Command method that is a easy way to do some
//* usefull formatting.
//*
//* TXLSSpreadSheet Methods
//* ----------------------------------------------------------------------------
//* Execute a formatting command. See below for possible values.
//* If the inplace editor is active, the selected word is formatted. If not the
//* selected cell(s) are formatted. When formatting cells, Command always
//* operates on the selected cells. XLSSpreadSheet.XLSSheet.SelectedAreas
//* property.
//* Example:
//*   // Set the cell or the selected word in the inplace editor to bold text.
//*   XLSSpreadSheet(xsscFmtFontBold,True);
//* Command(...);
//* ----------------------------------------------------------------------------
//* Clear;
//* Clears the worksheet. The same as starting a new, empty worksheet.
//* ----------------------------------------------------------------------------
//* Selects the worksheet to be visible.
//* SetSheet(const ASheetIndex: integer);
//* ----------------------------------------------------------------------------
//* Reads an Excel file. The name of the file is in the Filename property.
//* Read;
//* ----------------------------------------------------------------------------
//* Writes the file. The name of the file is in the Filename property.
//* Write;
//* ----------------------------------------------------------------------------
//* Clears the worksheet. The same as starting a new, empty worksheet.
//* CLear;
//* ----------------------------------------------------------------------------
//* Repaints the canvas.
//* InvalidateSheet;
//* ----------------------------------------------------------------------------
//* Reloads the sheet data and repaints the canvas. Use this method if something
//* else than cell values has changed, such as row height or column width.
//* InvalidateAndReloadSheet;
//* ----------------------------------------------------------------------------
//*
//* TXLSSpreadSheet properties.
//*
//* ----------------------------------------------------------------------------
//* XLSSheet         The visible worksheet (TXLSWorksheet)
//* ----------------------------------------------------------------------------
//* XSSSheet         The user interface of TXLSWorksheet (TXLSBookSheet)
//* ----------------------------------------------------------------------------
//* UI               Object that handles the logic for dialogs an controls that
//*                  sets spreadsheet values. See FrmFormatCells.pas for more info.
//* ----------------------------------------------------------------------------
//* ComponentVersion The version of the TXLSSpreadSheet component.
//* ----------------------------------------------------------------------------
//* ReadOnly         Spreadsheet is read only.
//* ----------------------------------------------------------------------------
//* SkinStyle        The look of the spreadsheet. can be xssExcelXP for Excel 97 style
//*                  of xssExcel2007 for Excel 2007 style. The xssExcel2007 style
//*                  requires that themed controls are enabled.
//* ----------------------------------------------------------------------------
//*
//* TXLSSpreadSheet events.
//*
//* ----------------------------------------------------------------------------
//*     Event fired after a worksheet becomes visible.
//*     property OnAfterSheetChange: TIntegerEvent read FAfterSheetChangedEvent write FAfterSheetChangedEvent;
//* ----------------------------------------------------------------------------
//*     Event fired when the position of the cell cursor is changed.
//*     property OnCellChanged: TXColRowEvent read GetOnCellChanged write SetOnCellChanged;
//* ----------------------------------------------------------------------------
//*     Event fired when the selected cell or cells changes.
//*     property OnSelectionChanged: TNotifyEvent read GetSelectionChanged write SetSelectionChanged;
//* ----------------------------------------------------------------------------
//*     Event fired before the inplace editor becomes visible. If Value is set to False, the editor will not be activated.
//*     property OnEditorOpen: TXBooleanEvent read FEditorOpenEvent write SetEditorOpenEvent;
//* ----------------------------------------------------------------------------
//*     Event fired before the inplace editor is closed. If Value is set to False, the text in the editor is not saved to the cell.
//*     If you want to examine the text in the editor, use this a code like this:
//*        if XLSSpreadSheet.XSSSheet.InplaceEditor.Text = 'oink' then
//*          Value := False;
//*     property OnEditorClose: TXBooleanEvent read FEditorCloseEvent write SetEditorCloseEvent;
//* ----------------------------------------------------------------------------
//*     Event fired if the value in the inplace editor not permits formatting, such as formulas.
//*     Use this event to disable formatting controls.
//*     property OnEditorDisableFmt: TNotifyEvent read FEditorDisableFmtEvent write SetEditorDisableFmt;
//* ----------------------------------------------------------------------------
//*     Event fired when a hyperlink is clicked.
//*     property OnHyperlinkClick: TXStringEvent read FHyperlinkClickEvent write SetHyperlinkClickEvent;
//* ----------------------------------------------------------------------------
//*     Event fired when the format of the text at the caret is changed.
//*     property OnIECharFormatChanged: TNotifyEvent read FIECharFmtEvent write FIECharFmtEvent;
//* ----------------------------------------------------------------------------
//*     property OnNotification: TXNotifyEvent read GetNotificationEvent write SetNotificationEvent;
//* ----------------------------------------------------------------------------


type TXSSCommand = (xsscFmtFontBold,             // Boolean. Bold typeface
                    xsscFmtFontItalic,           // Boolean. Italic typeface
                    xsscFmtFontUnderline,        // Boolean. Underline text
                    xsscFmtFontName,             // String. Typeface name.
                    xsscFmtFontSize,             // Float. Typeface size.
                    xsscFmtFontColor,            // Integer (RGB value) or TXc12Color.
                                                 // Typeface color.
                    xsscFmtAlignHorizLeft,       // Boolean. Align text horizontal left
                    xsscFmtAlignHorizCenter,     // Boolean. Align text horizontal center
                    xsscFmtAlignHorizRight,      // Boolean. Align text horizontal right

                    xsscFmtAlignVertTop,         // Boolean. Align text vertical top
                    xsscFmtAlignVertMiddle,      // Boolean. Align text vertical middle
                    xsscFmtAlignVertBottom,      // Boolean. Align text vertical bottom

                    xsscFmtRotate,               // Integer

                    xsscFmtIndent,               // Integer
                    xsscFmtWrapText,             // Boolean

                    xsscFmtCellColor,            // Integer (RGB value) or TXc12Color.
                                                 // Cell background color.
                    xsscFmtBorderThinLeft,       // No value. Set left cell border to thin
                    xsscFmtBorderThinRight,      // No value. Set right cell border to thin
                    xsscFmtBorderThinTop,        // No value. Set top cell border to thin
                    xsscFmtBorderThinBottom,     // No value. Set bottom cell border to thin
                    xsscFmtBorderThinOutline,    // No value. Set the border around the cell(s) to thin.
                    xsscFmtBorderThinInside,     // No value. Set the border inside the cell(s) to thin.
                    xsscFmtBorderNoBorder,       // No value. Remove cell borders.

                    xsscMergeCells,              // No value. Merge and center cells.
                    xsscUnMergeCells,            // No value. Delets all merged cells in selection.

                    xsscEditCopy,
                    xsscEditCut,
                    xsscEditPaste,
                    xsscEditPasteSpecial
                    );

const XSSCommandBorder = [xsscFmtBorderThinLeft,
                          xsscFmtBorderThinLeft,
                          xsscFmtBorderThinRight,
                          xsscFmtBorderThinTop,
                          xsscFmtBorderThinBottom,
                          xsscFmtBorderThinOutline,
                          xsscFmtBorderThinInside,
                          xsscFmtBorderNoBorder];

// Conditional formats icons bitmap. Fix alpha channel.
// Conditional formats are not scrolled correct.

// Known issues.
// - Conditional formats:
//   a. Number formats don't works:
//   b. Icons is not visible when printing.

//* Event fired when a hyperlink is clicked.
//* ~param Sender The TXLSSpreadSheet object.
//* ~param HyperlinkType The type of hyperlink clicked.
//* ~param SheetIndex The index of the worksheet where the hyperlink is.
//* ~param Col Cell column for the hyperlink.
//* ~param Row Cell row for the hyperlink.
//* ~param DefaultAction Set this parameter to True, in order to let TXLSSpreadSheet handle the default action. The defined default action is when a cell link is clicked. If DefaultAction is True, the cursor is moved to the target cell.
// type TClickHyperlinkEvent = procedure(Sender: TObject; HyperlinkType: THyperlinkType; SheetIndex,Col,Row: integer; DefaultAction: boolean) of object;

type TXLSSpreadSheet = class;

//             R
// L           i
// e     C     g
// f     o     h
// t     l     t
// +-----+-----+ Top
// |     |     |
// +-----+-----+ Row
// |     |     |
// +-----+-----+ Bottom

//* This is the TXLSSpreadSheet component.
//* Operations on the file data is done trough the TXLSReadWriteII object
//* (~[link XLS] property). The exception is when reading a file, as that
//* shall be with the Read method. Otherwise is not the canvas updated correctly.
    TXLSSpreadSheet = class(TXSSComponent)
private
     FManager         : TXc12Manager;
     FXLS             : TXLSReadWriteII5;

     FUI              : TXLSBookUI;

     FShiftState      : TXSSShiftKeys;
     FLastKey         : TXSSKey;

     FDoRepaint       : boolean;
     FComponentIsValid: boolean;
     FDelayedRead     : boolean;
     FX1,FY1,FX2,FY2  : integer;
     FCX1,FCY1,
     FCX2,FCY2        : integer;
     FOptions         : TXLSBookOptions;
     FXSSSheet        : TXLSBookSheet;
     FPanesXSplit,
     FPanesYSplit     : integer;
     FMultilayer      : boolean;
     FXLSSheet        : TXLSWorkSheet;
     FPendingSkinStyle: TXBookSkinStyle;

     FAfterSheetChangedEvent: TIntegerEvent;
     FCellChangeEvent       : TXColRowEvent;
     FSelectionChangedEvent : TNotifyEvent;
     FHyperlinkClickEvent   : TXStringEvent;
     FNotificationEvent     : TXNotifyEvent;
//     FShapeClickEvent: TShapeEvent;

     FEditorOpenEvent       : TXBooleanEvent;
     FEditorCloseEvent      : TXBooleanEvent;
     FEditorDisableFmtEvent : TNotifyEvent;
     FIECharFmtEvent        : TNotifyEvent;

     function  GetFilename: AxUCString;
     procedure SetFilename(const Value: AxUCString);
     procedure CreateTheComponent;
     function  GetComponentVersion: AxUCString;
     procedure SetComponentVersion(const Value: AxUCString);
     function  GetXLSSheet: TXLSWorkSheet;
     function  GetReadOnly: boolean;
     procedure SetReadOnly(const Value: boolean);
     function  GetSkinStyle: TXBookSkinStyle;
     procedure SetSkinStyle(const Value: TXBookSkinStyle);
     function  GetXSSSheet: TXLSBookSheet;
     function  GetOnCellChanged: TXColRowEvent;
     procedure SetOnCellChanged(const Value: TXColRowEvent);
     procedure SetEditorCloseEvent(const Value: TXBooleanEvent);
     procedure SetEditorOpenEvent(const Value: TXBooleanEvent);
     procedure SetHyperlinkClickEvent(const Value: TXStringEvent);
//     function  GetIECharFmtEvent: TNotifyEvent;
//     procedure SetIECharFmtEvent(const Value: TNotifyEvent);
     procedure SetEditorDisableFmt(const Value: TNotifyEvent);
     function  GetSelectionChanged: TNotifyEvent;
     procedure SetSelectionChanged(const Value: TNotifyEvent);
     function  GetNotificationEvent: TXNotifyEvent;
     procedure SetNotificationEvent(const Value: TXNotifyEvent);
     function  GetUserFunctionEvent: TUserFunctionEvent;
     procedure SetUserFunctionEvent(const Value: TUserFunctionEvent);
protected
     procedure CalcMetrics;
     procedure CreateHandle; override;
     procedure TabClickEvent(Sender: TObject; TabIndex: integer);
     procedure XLSSheetChanged(const ASheetIndex: integer);
     procedure Int_Clear;
     procedure XLSAfterRead(ASender: TObject);
     procedure XLSChanged(SheetIndex: integer);

     procedure WMKeyDown(var Message: TWMKeyDown); message WM_KEYDOWN;
     procedure WMKeyUp(var Message: TWMKeyUp); message WM_KEYUP;
//     procedure WndProc(var Message: TMessage); override;
public
     //* ~exclude
     constructor Create(AOwner: TComponent); override;
     //* ~exclude
     destructor Destroy; override;
     //* ~exclude
     procedure SetBounds(ALeft, ATop, AWidth, AHeight: Integer); override;
     procedure KeyDown(var Key: Word; Shift: TShiftState); override;

     procedure Command(const ACommand: TXSSCommand); overload;
     procedure Command(const ACommand: TXSSCommand; const AValue: boolean); overload;
     procedure Command(const ACommand: TXSSCommand; const AValue: integer); overload;
     procedure Command(const ACommand: TXSSCommand; const AValue: double); overload;
     procedure Command(const ACommand: TXSSCommand; const AValue: AxUCString); overload;
     procedure Command(const ACommand: TXSSCommand; const AValue: TXc12Color); overload;

     //* Repaints the component's canvas.
     procedure Paint; override;
      //* Reads the Excel file in the ~[link Filename] property. Do not call the
      //* Read method of the TXLSReadWriteII object (~[link XLS] property),
      //* as that will not update the canvas correctly.
     procedure Read;
     //* Writes the file. The name of the file is in the ~[link Filename] property.
     procedure Write;
     //* Clears the worksheet. The same as starting a new, empty worksheet.
     procedure Clear;
     procedure LoadFromStream(AStream: TStream);
     //* Sets the visible worksheet. SheetIndex is the index of the worksheet.
     procedure SetSheet(const ASheetIndex: integer);
     //* Sets the position of the cursor. SheetIndex is the index of the worksheet.
     //* Col and Row is the new cursor position. If the cursor position is outside
     //* the visible area, the worksheet will be scrolled so that it's visible.
     procedure SetCursorPos(SheetIndex,Col,Row: integer);

     procedure MergeCellsSelectedAreas;

     //* Repaints the canvas.
     procedure InvalidateSheet;
     //* Recalcs the sheet layout, rows ,columns and repaints.
     procedure InvalidateAndReloadSheet;
     //* Repaints an area of the canvas. SheetIndex is the worksheet for which the
     //* canvas shall be repainted, Col1,Row1,Col2,Row2 is the area.
     //* If the worksheet not is visible, this method has no effect.
     procedure InvalidateArea(SheetIndex,Col1,Row1,Col2,Row2: integer); overload;
     procedure InvalidateArea(Area: TCellArea); overload;
     procedure InvalidateRows;
     //* Repaints a cell on the canvas. SheetIndex is the worksheet for which the
     //* canvas shall be repainted, Col and Row is the cell.
     //* If the worksheet not is visible, this method has no effect.
     procedure InvalidateCell(SheetIndex,Col,Row: integer);

     //* Active worksheet. See the documentaion for XLSReadWriteII
     //* for more details.
     property XLSSheet: TXLSWorkSheet read GetXLSSheet;

     property XSSSheet: TXLSBookSheet read GetXSSSheet;

     property UI: TXLSBookUI read FUI;
published
     //* Version of the component.
     property ComponentVersion: AxUCString read GetComponentVersion write SetComponentVersion;

     //* The TXLSReadWriteII5 object. See the documentaion for XLSReadWriteII
     //* for more details.
     property XLS: TXLSReadWriteII5 read FXLS;

     //* The name of the Excel file.
     property Filename: AxUCString read GetFilename write SetFilename;
//     property Options: TXLSBookOptions read FOptions write FOptions;
     //* Set this to True if the file shall be opened in read only mode.
     property ReadOnly: boolean read GetReadOnly write SetReadOnly;

     property Options: TXLSBookOptions read FOptions;

     property SkinStyle: TXBookSkinStyle read GetSkinStyle write SetSkinStyle;

     //* Event fired after a worksheet becomes visible.
     property OnAfterSheetChange: TIntegerEvent read FAfterSheetChangedEvent write FAfterSheetChangedEvent;
     //* Event fired when the position of the cell cursor is changed.
     property OnCellChanged: TXColRowEvent read GetOnCellChanged write SetOnCellChanged;
     //* Event fired when the selected cell or cells changes.
     property OnSelectionChanged: TNotifyEvent read GetSelectionChanged write SetSelectionChanged;
     //* Event fired before the inplace editor becomes visible. If Value is set to False, the editor will not be activated.
     property OnEditorOpen: TXBooleanEvent read FEditorOpenEvent write SetEditorOpenEvent;
     //* Event fired before the inplace editor is closed. If Value is set to False, the text in the editor is not saved to the cell.
     //* If you want to examine the text in the editor, use this a code like this:
     //*   if XLSSpreadSheet.XSSSheet.InplaceEditor.Text = 'oink' then
     //*     Value := False;
     property OnEditorClose: TXBooleanEvent read FEditorCloseEvent write SetEditorCloseEvent;
     //* Event fired if the value in the inplace editor not permits formatting, such as formulas.
     //* Use this event to disable formatting controls.
     property OnEditorDisableFmt: TNotifyEvent read FEditorDisableFmtEvent write SetEditorDisableFmt;
     //* Event fired when a hyperlink is clicked.
     property OnHyperlinkClick: TXStringEvent read FHyperlinkClickEvent write SetHyperlinkClickEvent;
     //* Event fired when a shape (drawing object or a picture) is clicked.
//     property OnShapeClick: TShapeEvent read FShapeClickEvent write FShapeClickEvent;
     //*
     property OnIECharFormatChanged: TNotifyEvent read FIECharFmtEvent write FIECharFmtEvent;
     //* Notification event.
     //* The Notification value can be:
     //*   xbnCellLocked The user is trying to edit a locked cell.
     property OnNotification: TXNotifyEvent read GetNotificationEvent write SetNotificationEvent;
     //* Event fired when a user function is evaluated.
     property OnUserFunction: TUserFunctionEvent read GetUserFunctionEvent write SetUserFunctionEvent;

     property Align;
     property Anchors;
     property Constraints;
     property UseDockManager default True;
     property DockSite;
     property DragCursor;
     property DragKind;
     property DragMode;
     property Enabled;
     property ParentShowHint;
     property PopupMenu;
     property ShowHint;
     property TabOrder;
     property TabStop;
     property Visible;

     property OnCanResize;
     property OnClick;
     property OnConstrainedResize;
     property OnContextPopup;
     property OnDockDrop;
     property OnDockOver;
     property OnDblClick;
     property OnDragDrop;
     property OnDragOver;
     property OnEndDock;
     property OnEndDrag;
     property OnEnter;
     property OnExit;
     property OnGetSiteInfo;
     property OnMouseDown;
{$ifdef DELPHI_2009_OR_LATER}
     property OnMouseEnter;
     property OnMouseLeave;
{$endif}
     property OnMouseMove;
     property OnMouseUp;
     property OnResize;
     property OnStartDock;
     property OnStartDrag;
     property OnUnDock;

     property OnKeyDown;
     property OnKeyPress;
     property OnKeyUp;
     end;

implementation

{$R XLSBook.res}

{ TXLSSpreadSheet }

procedure TXLSSpreadSheet.CalcMetrics;
begin
  FOptions.General.FontSize := FXLS.Font.Size;
  FOptions.General.FontName := FXLS.Font.Name;

  FSkin.AssignSystemFont(FXLS.Font.Name,Round(FXLS.Font.Size * FSkin.GDI.Zoom),Integer(FXLS.Font.Charset));
end;

procedure TXLSSpreadSheet.Clear;
begin
  FXLS.Clear;
  FXLS.Filename := '';
  XLSChanged(0);
end;

procedure TXLSSpreadSheet.Command(const ACommand: TXSSCommand; const AValue: AxUCString);
begin
  if FXSSSheet.InplaceEditor <> Nil then begin
    case ACommand of
      xsscFmtFontName     : FXSSSheet.InplaceEditor.CHP.FontName := AValue;
    end;
    FXSSSheet.InplaceEditor.Command(axcFormatCHP);
  end
  else begin
    FXLS.CmdFormat.BeginEdit(FXLSSheet);
    case ACommand of
      xsscFmtFontName     : FXLS.CmdFormat.Font.Name := AValue;
    end;
    FXLS.CmdFormat.Apply;

    InvalidateSheet;
  end;
end;

procedure TXLSSpreadSheet.Command(const ACommand: TXSSCommand);
begin
  if ACommand in [xsscEditCopy,xsscEditCut,xsscEditPaste,xsscEditPasteSpecial,xsscMergeCells,xsscUnMergeCells] then begin
    case ACommand of
      xsscEditCopy            : FXSSSheet.HandleKey(kyCopy,[]);
      xsscEditCut             : FXSSSheet.HandleKey(kyCut,[]);
      xsscEditPaste           : FXSSSheet.HandleKey(kyPaste,[]);
      xsscEditPasteSpecial    : FXSSSheet.HandleKey(kyPasteSpecial,[]);

      xsscMergeCells          : FXLSSheet.MergeCells;
      xsscUnMergeCells        : FXLSSheet.UnMergeCells;
    end;
  end
  else begin
    FXLS.CmdFormat.BeginEdit(FXLSSheet);
    FXLS.CmdFormat.Border.Color.Color := MakeXc12ColorAuto;
    FXLS.CmdFormat.Border.Style := cbsThin;

    case ACommand of
      xsscFmtBorderThinLeft   : FXLS.CmdFormat.Border.Side[cbsLeft] := True;
      xsscFmtBorderThinRight  : FXLS.CmdFormat.Border.Side[cbsRight] := True;
      xsscFmtBorderThinTop    : FXLS.CmdFormat.Border.Side[cbsTop] := True;
      xsscFmtBorderThinBottom : FXLS.CmdFormat.Border.Side[cbsBottom] := True;
      xsscFmtBorderThinOutline: FXLS.CmdFormat.Border.Preset(cbspOutline);
      xsscFmtBorderThinInside : FXLS.CmdFormat.Border.Preset(cbspInside);
      xsscFmtBorderNoBorder   : FXLS.CmdFormat.Border.Preset(cbspNone);

      xsscFmtAlignHorizLeft   : FXLS.CmdFormat.Alignment.Horizontal := chaLeft;
      xsscFmtAlignHorizCenter : FXLS.CmdFormat.Alignment.Horizontal := chaCenter;
      xsscFmtAlignHorizRight  : FXLS.CmdFormat.Alignment.Horizontal := chaRight;

      xsscFmtAlignVertTop     : FXLS.CmdFormat.Alignment.Vertical := cvaTop;
      xsscFmtAlignVertMiddle  : FXLS.CmdFormat.Alignment.Vertical := cvaCenter;
      xsscFmtAlignVertBottom  : FXLS.CmdFormat.Alignment.Vertical := cvaBottom;

      xsscEditCopy            : FXSSSheet.HandleKey(kyCopy,[]);
      xsscEditCut             : FXSSSheet.HandleKey(kyCut,[]);
      xsscEditPaste           : FXSSSheet.HandleKey(kyPaste,[]);
      xsscEditPasteSpecial    : FXSSSheet.HandleKey(kyPasteSpecial,[]);
    end;

    FXLS.CmdFormat.Apply;
  end;

  InvalidateSheet;
end;

procedure TXLSSpreadSheet.Command(const ACommand: TXSSCommand; const AValue: TXc12Color);
begin
  if FXSSSheet.InplaceEditor <> Nil then begin
    case ACommand of
      xsscFmtFontColor    : FXSSSheet.InplaceEditor.CHP.Color := AValue.ARGB;
    end;
    FXSSSheet.InplaceEditor.Command(axcFormatCHP);
  end
  else begin
    FXLS.CmdFormat.BeginEdit(FXLSSheet);
    case ACommand of
      xsscFmtFontColor     : FXLS.CmdFormat.Font.Color.Color := AValue;
      xsscFmtCellColor     : FXLS.CmdFormat.Fill.BackgroundColor.Color := AValue;
    end;
    FXLS.CmdFormat.Apply;

    InvalidateSheet;
  end;
end;

procedure TXLSSpreadSheet.Command(const ACommand: TXSSCommand; const AValue: boolean);
begin
  if FXSSSheet.InplaceEditor <> Nil then begin
    case ACommand of
      xsscFmtFontBold     : FXSSSheet.InplaceEditor.Command(axcFormatBold);
      xsscFmtFontItalic   : FXSSSheet.InplaceEditor.Command(axcFormatItalic);
      xsscFmtFontUnderline: FXSSSheet.InplaceEditor.Command(axcFormatUnderline);
    end;
  end
  else begin
    FXLS.CmdFormat.BeginEdit(FXLSSheet);
    case ACommand of
      xsscFmtFontBold        : FXLS.CmdFormat.Font.Bold := AValue;
      xsscFmtFontItalic      : FXLS.CmdFormat.Font.Italic := AValue;
      xsscFmtFontUnderline   : begin
        if AValue then
          FXLS.CmdFormat.Font.Underline := xulSingle
        else
          FXLS.CmdFormat.Font.Underline := xulNone;
      end;

      xsscFmtAlignHorizLeft  : FXLS.CmdFormat.Alignment.Horizontal := chaLeft;
      xsscFmtAlignHorizCenter: FXLS.CmdFormat.Alignment.Horizontal := chaCenter;
      xsscFmtAlignHorizRight : FXLS.CmdFormat.Alignment.Horizontal := chaRight;

      xsscFmtAlignVertTop    : FXLS.CmdFormat.Alignment.Vertical := cvaTop;
      xsscFmtAlignVertMiddle : FXLS.CmdFormat.Alignment.Vertical := cvaCenter;
      xsscFmtAlignVertBottom : FXLS.CmdFormat.Alignment.Vertical := cvaBottom;

      xsscFmtWrapText        : FXLS.CmdFormat.Alignment.WrapText := AValue;
    end;
    FXLS.CmdFormat.Apply;

    InvalidateSheet;
  end;
end;

procedure TXLSSpreadSheet.Command(const ACommand: TXSSCommand; const AValue: double);
begin
  if FXSSSheet.InplaceEditor <> Nil then begin
    case ACommand of
      xsscFmtFontSize     : FXSSSheet.InplaceEditor.CHP.Size := AValue;
    end;
    FXSSSheet.InplaceEditor.Command(axcFormatCHP);
  end
  else begin
    FXLS.CmdFormat.BeginEdit(FXLSSheet);
    case ACommand of
      xsscFmtFontSize     : FXLS.CmdFormat.Font.Size := AValue;
    end;
    FXLS.CmdFormat.Apply;

    InvalidateAndReloadSheet;
  end;
end;

procedure TXLSSpreadSheet.Command(const ACommand: TXSSCommand; const AValue: integer);
begin
  if FXSSSheet.InplaceEditor <> Nil then begin
  end
  else begin
    FXLS.CmdFormat.BeginEdit(FXLSSheet);
    case ACommand of
      xsscFmtFontColor     : FXLS.CmdFormat.Font.Color.RGB := AValue;
      xsscFmtCellColor     : FXLS.CmdFormat.Fill.BackgroundColor.RGB := AValue;
      xsscFmtRotate        : FXLS.CmdFormat.Alignment.Rotation := Fork(AValue,-90,90);
      xsscFmtIndent        : FXLS.CmdFormat.Alignment.Rotation := Fork(AValue,0,32);
    end;
    FXLS.CmdFormat.Apply;

    InvalidateSheet;
  end;
end;

procedure TXLSSpreadSheet.Int_Clear;
begin
  FRootWin.DeleteChilds;
  FSkin.Clear;
end;

constructor TXLSSpreadSheet.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);

  FPendingSkinStyle := xssExcelNone;

  InitThemeLibrary;

  FMultilayer := True;
  FDoRepaint := True;
  ControlStyle := ControlStyle + [csCaptureMouse,csOpaque];

  FXLS := TXLSReadWriteII5.Create(Self);
  FXLS.CmdFormat.SetRowHeight := True;
//  FXLS.OnFunction := XLSFuncEvent;
  FManager := FXLS.Manager;
  FManager.OnXSSAfterRead := XLSAfterRead;

  FOptions := TXLSBookOptions.Create(FXLS);

  FUI := TXLSBookUI.Create(FXLS);

  Height := 200;
  Width := 400;
  FY2 := 200;
  FX2 := 400;
  TabStop := True;
end;

procedure TXLSSpreadSheet.CreateHandle;
begin
  inherited;
  if not FComponentIsValid and (Parent <> Nil) then begin
    CreateTheComponent;
    FComponentIsValid := True;
    SetBounds(Left,Top,Width,Height);
    if FDelayedRead then begin
      FDelayedRead := False;
      Read;
    end;
  end;
end;

procedure TXLSSpreadSheet.CreateTheComponent;
begin
  CalcMetrics;
  XLSChanged(FXLS.SelectedTab);
  if FPendingSkinStyle <> xssExcelNone then begin
    SetSkinStyle(FPendingSkinStyle);
    FPendingSkinStyle := xssExcelNone;
  end;

  if FRootWin.FocusedWin = Nil then
    FXSSSheet.Focus;
end;

destructor TXLSSpreadSheet.Destroy;
begin
//  Int_Clear;
  FOptions.Free;
  FXLS.Free;
  FUI.Free;
  inherited;
end;

procedure TXLSSpreadSheet.XLSAfterRead;
var
  i,j: integer;
  C: TXc12Comment;
begin
  for i := 0 to FManager.Worksheets.Count - 1 do begin
    for j := 0 to FManager.Worksheets[i].Comments.Count - 1 do begin
      C := FManager.Worksheets[i].Comments[j];
      C.Col1Offs := FSkin.GDI.PixelsToEMU(C.Col1Offs);
      C.Row1Offs := FSkin.GDI.PixelsToEMU(C.Row1Offs);
      C.Col2Offs := FSkin.GDI.PixelsToEMU(C.Col2Offs);
      C.Row2Offs := FSkin.GDI.PixelsToEMU(C.Row2Offs);
    end;
  end;
end;

function TXLSSpreadSheet.GetFilename: AxUCString;
begin
  Result := FXLS.Filename;
end;

function TXLSSpreadSheet.GetNotificationEvent: TXNotifyEvent;
begin
  Result := FNotificationEvent;
end;

//function TXLSSpreadSheet.GetIECharFmtEvent: TNotifyEvent;
//begin
//  Result := FIECharFmtEvent;
//end;

function TXLSSpreadSheet.GetOnCellChanged: TXColRowEvent;
begin
  Result := FCellChangeEvent;
end;

function TXLSSpreadSheet.GetReadOnly: boolean;
begin
  Result := FOptions.ReadOnly;
end;

function TXLSSpreadSheet.GetSelectionChanged: TNotifyEvent;
begin
  Result := FSelectionChangedEvent;
end;

function TXLSSpreadSheet.GetSkinStyle: TXBookSkinStyle;
begin
  Result := FSkin.Style;
end;

function TXLSSpreadSheet.GetUserFunctionEvent: TUserFunctionEvent;
begin
  Result := FXLS.OnUserFunction;
end;

function TXLSSpreadSheet.GetXLSSheet: TXLSWorkSheet;
begin
  Result := FXLSSheet;
end;

function TXLSSpreadSheet.GetXSSSheet: TXLSBookSheet;
begin
  Result := FXSSSheet;
end;

procedure TXLSSpreadSheet.MergeCellsSelectedAreas;
var
  i: integer;
begin
  for i := 0 to FXLSSheet.SelectedAreas.Count - 1 do begin
    with FXLSSheet.SelectedAreas[i] do begin
      FXLSSheet.MergeCells(Col1,Row1,Col2,Row2);
      InvalidateArea(-1,Col1,Row1,Col2,Row2);
    end;
  end;
end;

procedure TXLSSpreadSheet.SetCursorPos(SheetIndex, Col, Row: integer);
begin
//  XLSChanged(TabIndex);
  FXLS.SelectedTab := SheetIndex;
  FXLSSheet := FXLS[SheetIndex];
  if Col > FXSSSheet.Columns.RightCol then
    FXLSSheet.LeftCol := Col;
  if Row > FXSSSheet.Rows.BottomRow then
    FXLSSheet.TopRow := Row;
  FXLSSheet.SelectedAreas.Clear;
  FXLSSheet.SelectedAreas.Add(Col,Row,Col,Row);
  FXSSSheet.SheetChanged(FXLS[SheetIndex]);
  FXSSSheet.Tabs.SelectedTab := SheetIndex;
//  Paint;
  Invalidate;
end;

procedure TXLSSpreadSheet.SetEditorCloseEvent(const Value: TXBooleanEvent);
begin
  FEditorCloseEvent := Value;
  if FXSSSheet <> Nil then
    FXSSSheet.OnEditorClose := FEditorCloseEvent;
end;

procedure TXLSSpreadSheet.SetEditorDisableFmt(const Value: TNotifyEvent);
begin
  FEditorDisableFmtEvent := Value;
  if FXSSSheet <> Nil then
    FXSSSheet.OnEditorDisableFmt := FEditorDisableFmtEvent;
end;

procedure TXLSSpreadSheet.SetEditorOpenEvent(const Value: TXBooleanEvent);
begin
  FEditorOpenEvent := Value;
  if FXSSSheet <> Nil then
    FXSSSheet.OnEditorOpen := FEditorOpenEvent;
end;

procedure TXLSSpreadSheet.Paint;
begin
  inherited;

  if FXLSSheet.Zoom > 0 then
    FSkin.GDI.Zoom := FXLSSheet.Zoom / 100;

  if not FComponentIsValid or not FDoRepaint then begin
    FDoRepaint := True;
    Exit;
  end;

end;

procedure TXLSSpreadSheet.Read;
begin
  if not FComponentIsValid then begin
    FDelayedRead := True;
    Exit;
  end;

  FXLS.Read;

  XLSChanged(FXLS.SelectedTab);
end;

procedure TXLSSpreadSheet.SetBounds(ALeft, ATop, AWidth, AHeight: Integer);
var
  X,Y: integer;
begin
  inherited;
  if {FPrinting or} not FComponentIsValid then
    Exit;

  if ((AWidth - 1) = FX2) and ((AHeight - 1) = FY2) then
    Exit;

  FX1 := 0;
  FY1 := 0;
  FX2 := AWidth - 1;
  FY2 := AHeight - 1;
  FCX1 := FX1 + BOOK_BORDERWIDTH;
  FCY1 := FY1 + BOOK_BORDERWIDTH;
  FCX2 := FX2 - BOOK_BORDERWIDTH;
  FCY2 := FY2 - BOOK_BORDERWIDTH;

//  FDoRepaint := ((AWidth - 1) > FX2) or ((AHeight - 1) > FY2);
  // Subtract scroll bar
  X := FCX2;
  Y := FCY2;
  if FOptions.View.HorizontalScroll then
    Dec(X,GetSystemMetrics(SM_CXHSCROLL));
  if FOptions.View.VerticalScroll then
    Dec(Y,GetSystemMetrics(SM_CYHSCROLL	));
  FSkin.SetSize(AWidth,AHeight);
  FSkin.SetClientSize(FCX1,FCY1,X,Y);
  FXSSSheet.SetSize(FCX1, FCY1, FCX2, FCY2);
//  SetPanesRects;
end;

procedure TXLSSpreadSheet.SetFilename(const Value: AxUCString);
begin
  FXLS.Filename := Value;
end;

procedure TXLSSpreadSheet.SetHyperlinkClickEvent(const Value: TXStringEvent);
begin
  FHyperlinkClickEvent := Value;
  if FXSSSheet <> Nil then
    FXSSSheet.OnHyperlinkClick := FHyperlinkClickEvent;
end;

procedure TXLSSpreadSheet.SetNotificationEvent(const Value: TXNotifyEvent);
begin
  FNotificationEvent := Value;
  if FXSSSheet <> Nil then
    FXSSSheet.OnNotification := FNotificationEvent;
end;

//procedure TXLSSpreadSheet.SetIECharFmtEvent(const Value: TNotifyEvent);
//begin
//  FIECharFmtEvent := Value;
//  if FXSSSheet <> Nil then
//    FXSSSheet.OnIECharFmt := FIECharFmtEvent;
//end;

procedure TXLSSpreadSheet.SetOnCellChanged(const Value: TXColRowEvent);
begin
  FCellChangeEvent := Value;
  if FXSSSheet <> Nil then
    FXSSSheet.OnCellChanged := FCellChangeEvent;
end;

procedure TXLSSpreadSheet.WMKeyDown(var Message: TWMKeyDown);
begin
  if (FRootWin.FocusedWin <> Nil) and (FRootWin.FocusedWin is TXBookInplaceEditor) then
    inherited
  else begin
    if FLastKey = kyEnd then begin
      case Message.CharCode of
        VK_UP    : FRootWin.HandleKey(kyFirstRow,FShiftState);
        VK_DOWN  : FRootWin.HandleKey(kyLastRow,FShiftState);
        VK_LEFT  : FRootWin.HandleKey(kyFirstCol,FShiftState);
        VK_RIGHT : FRootWin.HandleKey(kyLastCol,FShiftState);
      end;
    end
    else begin
//      kyCopy,kyCut,kyPast,kyDelete
      case Message.CharCode of
        $43,$63 {c,C} : if FShiftState = [skCtrl] then FRootWin.HandleKey(kyCopy,FShiftState);
        $56,$76 {v,V} : if FShiftState = [skCtrl] then FRootWin.HandleKey(kyPaste,FShiftState);
        $58,$78 {x,X} : if FShiftState = [skCtrl] then FRootWin.HandleKey(kyCut,FShiftState);
        VK_DELETE     : if FShiftState = []       then FRootWin.HandleKey(kyDelete,FShiftState);

        VK_ESCAPE     : FRootWin.HandleKey(kyEscape,FShiftState);

        VK_TAB        : FRootWin.HandleKey(kyTab,FShiftState);

        VK_UP         : FRootWin.HandleKey(kyUp,FShiftState);
        VK_DOWN       : FRootWin.HandleKey(kyDown,FShiftState);
        VK_LEFT       : FRootWin.HandleKey(kyLeft,FShiftState);
        VK_RIGHT      : FRootWin.HandleKey(kyRight,FShiftState);
        VK_PRIOR      : FRootWin.HandleKey(kyPgUp,FShiftState);
        VK_NEXT       : FRootWin.HandleKey(kyPgDown,FShiftState);

        VK_F2         : begin
          if FShiftState = [] then
            FRootWin.HandleKey(kyInplaceEdit,FShiftState);
        end;

        VK_CONTROL: FShiftState := FShiftState + [skCtrl];
        VK_SHIFT  : FShiftState := FShiftState + [skShift];
      end;
    end;
    FLastKey := kyNone;
  end;
end;

procedure TXLSSpreadSheet.WMKeyUp(var Message: TWMKeyUp);
begin
  case Message.CharCode of
    VK_CONTROL: FShiftState := FShiftState - [skCtrl];
    VK_SHIFT  : FShiftState := FShiftState - [skShift];
    VK_END    : FLastKey := kyEnd;
  end;
end;

procedure TXLSSpreadSheet.SetReadOnly(const Value: boolean);
begin
  FOptions.ReadOnly := Value;
end;

procedure TXLSSpreadSheet.SetSelectionChanged(const Value: TNotifyEvent);
begin
  FSelectionChangedEvent := Value;
  if FXSSSheet <> Nil then
    FXSSSheet.OnSelectionChanged := FSelectionChangedEvent;
end;

procedure TXLSSpreadSheet.SetSheet(const ASheetIndex: integer);
begin
//  XLSChanged(TabIndex);
  FXLS.SelectedTab := ASheetIndex;
  FXLSSheet := FXLS[ASheetIndex];
  XLSSheetChanged(ASheetIndex);
  Invalidate;
end;

procedure TXLSSpreadSheet.SetSkinStyle(const Value: TXBookSkinStyle);
begin
  if FSkin <> Nil then begin
    FSkin.Style := Value;
    FXSSSheet.Drawing.Skin.Style := Value;
    XLSChanged(FXLS.SelectedTab);
  end
  else
    FPendingSkinStyle := Value;
end;

procedure TXLSSpreadSheet.SetUserFunctionEvent(const Value: TUserFunctionEvent);
begin
  FXLS.OnUserFunction := Value;
end;

function TXLSSpreadSheet.GetComponentVersion: AxUCString;
begin
  Result := CurrentVersionNumber;
end;

procedure TXLSSpreadSheet.SetComponentVersion(const Value: AxUCString);
begin

end;

procedure TXLSSpreadSheet.InvalidateArea(SheetIndex, Col1, Row1, Col2, Row2: integer);
begin
  FXSSSheet.HideCursor;
  FXSSSheet.InvalidateArea(Col1,Row1,Col2,Row2);
  FXSSSheet.ShowCursor;
  FLayers.ReleaseHandle;
end;

procedure TXLSSpreadSheet.InvalidateAndReloadSheet;
begin
  FXSSSheet.InvalidateAndReload;
  XLSChanged(FXLS.SelectedTab);
  InvalidateSheet;
end;

procedure TXLSSpreadSheet.InvalidateArea(Area: TCellArea);
begin
  InvalidateArea(-1,Max(Area.Col1 - 1,0),Max(Area.Row1 - 1,0),Area.Col2,Area.Row2);
end;

procedure TXLSSpreadSheet.InvalidateCell(SheetIndex, Col, Row: integer);
begin
  InvalidateArea(SheetIndex,Col,Row,Col,Row);
end;

procedure TXLSSpreadSheet.InvalidateRows;
begin
  FXSSSheet.Rows.CalcHeaders;
end;

procedure TXLSSpreadSheet.KeyDown(var Key: Word; Shift: TShiftState);
begin
  inherited;
  if FXSSSheet.InplaceEditor <> Nil then begin
    if Key in [VK_ESCAPE,VK_RETURN] then begin
      FXSSSheet.HideInplaceEdit(Key = VK_RETURN);
      FXSSSheet.Focus;
    end
    else if Key = VK_TAB then begin
      if Shift = [] then
        FXSSSheet.HideInplaceEdit(True,kyRight)
      else if Shift = [ssShift] then
        FXSSSheet.HideInplaceEdit(True,kyLeft);
      FXSSSheet.Focus;
    end
    else
      FXSSSheet.InplaceEditor.Command(KeyToCommand(Key,Shift));
  end;
end;

procedure TXLSSpreadSheet.LoadFromStream(AStream: TStream);
begin
  FXLS.LoadFromStream(AStream);
  XLSChanged(FXLS.SelectedTab);
end;

procedure TXLSSpreadSheet.InvalidateSheet;
begin
  FXSSSheet.HideCursor;
  FXSSSheet.Drawing.LoadObjects(FXLSSheet.Xc12Sheet,FXLSSheet.Drawing);
  FXSSSheet.Drawing.CalcObjects;
  FXSSSheet.InvalidateAndReload;
//  FXSSSheet.InvalidateArea(0,0,MAXINT,MAXINT);
  FXSSSheet.ShowCursor;
  Repaint;
  FSystem.ProcessRequests;
end;

procedure TXLSSpreadSheet.TabClickEvent(Sender: TObject; TabIndex: integer);
begin
  SetSheet(TabIndex);
end;

procedure TXLSSpreadSheet.XLSChanged(SheetIndex: integer);
var
  i: integer;
//  Pict: TXLSBookPictureBMP;
//  Meta: TXLSBookPictureMetafile;
begin
  Int_Clear;
  FXLSSheet := FXLS[SheetIndex];
  FPanesXSplit := 0;
  FPanesYSplit := 0;

  FXSSSheet := TXLSBookSheet.Create(Self,FRootWin,FOptions,FLayers,FManager,FXLS,FXLS[0]);
  FXSSSheet.Enabled := False;
  try
    FXSSSheet.Visible := True;
    FRootWin.Clear;
    FRootWin.Add(FXSSSheet);

    CalcMetrics;
    XLSSheetChanged(SheetIndex);
    FXSSSheet.SetSize(FX1 + BOOK_BORDERWIDTH, FY1 + BOOK_BORDERWIDTH, FX2 - BOOK_BORDERWIDTH, FY2 - BOOK_BORDERWIDTH);
    FXSSSheet.Tabs.DeleteChilds;
    for i := 0 to FXLS.Count - 1 do
      FXSSSheet.Tabs.Add(FXLS[i].Name,FXLS[i].TabColor);
    FXSSSheet.Tabs.SelectedTab := SheetIndex;
    FXSSSheet.Tabs.OnClick := TabClickEvent;
    FXSSSheet.OnCellChanged := FCellChangeEvent;
    FXSSSheet.OnSelectionChanged := FSelectionChangedEvent;
    FXSSSheet.OnEditorOpen := FEditorOpenEvent;
    FXSSSheet.OnEditorClose := FEditorCloseEvent;
    FXSSSheet.OnEditorDisableFmt := FEditorDisableFmtEvent;
    FXSSSheet.OnHyperlinkClick := FHyperlinkClickEvent;
    FXSSSheet.OnIECharFmt := FIECharFmtEvent;
    FXSSSheet.OnNotification := FNotificationEvent;

    Invalidate;
  finally
    FXSSSheet.Enabled := True;
  end;
end;

procedure TXLSSpreadSheet.XLSSheetChanged(const ASheetIndex: integer);
begin
  FXSSSheet.SheetChanged(FXLS[ASheetIndex]);
  if Assigned(FAfterSheetChangedEvent) then
    FAfterSheetChangedEvent(Self,ASheetIndex);
  if Assigned(FCellChangeEvent) then
    FCellChangeEvent(Self,FXLSSheet.SelectedAreas.ActiveCol,FXLSSheet.SelectedAreas.ActiveRow);
  FUI.Sheet := FXLS[ASheetIndex];
end;

procedure TXLSSpreadSheet.Write;
begin
  FXLS.Write;
end;

end.
