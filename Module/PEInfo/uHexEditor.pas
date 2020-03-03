unit uHexEditor;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, grids;

const
  WM_STATECHANGED = WM_USER + 2;

  ctMacToISO: array [128 .. 255] of AnsiChar = (                                                    //
    #$C4, #$C5, #$C7, #$C9, #$D1, #$D6, #$DC, #$E1, #$E0, #$E2, #$E4, #$E3, #$E5, #$E7, #$E9, #$E8, //
    #$EA, #$EB, #$ED, #$EC, #$EE, #$EF, #$F1, #$F3, #$F2, #$F4, #$F6, #$F5, #$FA, #$F9, #$FB, #$FC, //
    #$DD, #$B0, #$A2, #$A3, #$A7, #$80, #$B6, #$DF, #$AE, #$A9, #$81, #$B4, #$A8, #$82, #$C6, #$D8, //
    #$83, #$B1, #$BE, #$84, #$A5, #$B5, #$8F, #$85, #$BD, #$BC, #$86, #$AA, #$BA, #$87, #$E6, #$F8, //
    #$BF, #$A1, #$AC, #$88, #$9F, #$89, #$90, #$AB, #$BB, #$8A, #$A0, #$C0, #$C3, #$D5, #$91, #$A6, //
    #$AD, #$8B, #$B3, #$B2, #$8C, #$B9, #$F7, #$D7, #$FF, #$8D, #$8E, #$A4, #$D0, #$F0, #$DE, #$FE, //
    #$FD, #$B7, #$92, #$93, #$94, #$C2, #$CA, #$C1, #$CB, #$C8, #$CD, #$CE, #$CF, #$CC, #$D3, #$D4, //
    #$95, #$D2, #$DA, #$DB, #$D9, #$9E, #$96, #$97, #$AF, #$98, #$99, #$9A, #$B8, #$9B, #$9C, #$9D  //
    );

  ctISOToMac: array [128 .. 255] of AnsiChar = (                                                    //
    #$A5, #$AA, #$AD, #$B0, #$B3, #$B7, #$BA, #$BD, #$C3, #$C5, #$C9, #$D1, #$D4, #$D9, #$DA, #$B6, //
    #$C6, #$CE, #$E2, #$E3, #$E4, #$F0, #$F6, #$F7, #$F9, #$FA, #$FB, #$FD, #$FE, #$FF, #$F5, #$C4, //
    #$CA, #$C1, #$A2, #$A3, #$DB, #$B4, #$CF, #$A4, #$AC, #$A9, #$BB, #$C7, #$C2, #$D0, #$A8, #$F8, //
    #$A1, #$B1, #$D3, #$D2, #$AB, #$B5, #$A6, #$E1, #$FC, #$D5, #$BC, #$C8, #$B9, #$B8, #$B2, #$C0, //
    #$CB, #$E7, #$E5, #$CC, #$80, #$81, #$AE, #$82, #$E9, #$83, #$E6, #$E8, #$ED, #$EA, #$EB, #$EC, //
    #$DC, #$84, #$F1, #$EE, #$EF, #$CD, #$85, #$D7, #$AF, #$F4, #$F2, #$F3, #$86, #$A0, #$DE, #$A7, //
    #$88, #$87, #$89, #$8B, #$8A, #$8C, #$BE, #$8D, #$8F, #$8E, #$90, #$91, #$93, #$92, #$94, #$95, //
    #$DD, #$96, #$98, #$97, #$99, #$9B, #$9A, #$D6, #$BF, #$9D, #$9C, #$9E, #$9F, #$E0, #$DF, #$D8  //
    );

  ctEBCDICToISO: array [0 .. 255] of AnsiChar = (                            //
    #0, #0, #0, #0, #0, #0, #0, #0, #0, #0, #0, #0, #0, #0, #0, #0,          //
    #0, #0, #0, #0, #0, #0, #0, #0, #0, #0, #0, #0, #0, #0, #0, #0,          //
    #0, #0, #0, #0, #0, #0, #0, #0, #0, #0, #0, #0, #0, #0, #0, #0,          //
    #0, #0, #0, #0, #0, #0, #0, #0, #0, #0, #0, #0, #0, #0, #0, #0,          //
    ' ', #0, #0, #0, #0, #0, #0, #0, #0, #0, #0, '.', '<', '(', '+', '?',    //
    '&', '&', #0, #0, #0, #0, #0, #0, #0, #0, '!', '$', '*', ')', ';', #0,   //
    '-', '/', #0, #0, #0, #0, #0, #0, #0, #0, '|', ',', '%', '_', '>', '?',  //
    #0, #0, #0, #0, #0, #0, #0, #0, #0, '`', ':', '#', '@', '''', '=', '"',  //
    #0, 'a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', #0, #0, #0, #0, #0, #0, //
    #0, 'j', 'k', 'l', 'm', 'n', 'o', 'p', 'q', 'r', #0, #0, #0, #0, #0, #0, //
    #0, '~', 's', 't', 'u', 'v', 'w', 'x', 'y', 'z', #0, #0, #0, #0, #0, #0, //
    #0, #0, #0, #0, #0, #0, #0, #0, #0, #0, #0, #0, #0, #0, #0, #0,          //
    #0, 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', #0, #0, #0, #0, #0, #0, //
    #0, 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', #0, #0, #0, #0, #0, #0, //
    '\', #0, 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', #0, #0, #0, #0, #0, #0, //
    '0', '1', '2', '3', '4', '5', '6', '7', '8', '9', #0, #0, #0, #0, #0, #0 //
    );

  ctISOToEBCDIC: array [0 .. 255] of AnsiChar = (                                   //
    #00, #00, #00, #00, #00, #00, #00, #00, #00, #00, #00, #00, #00, #00, #00, #00, //
    #00, #00, #00, #00, #00, #00, #00, #00, #00, #00, #00, #00, #00, #00, #00, #00, //
    '@', 'Z', '', '{', '[', 'l', 'P', '}', 'M', ']', '\', 'N', 'k', '`', 'K', 'a', //
    '?', '?', '?', '?', '?', '?', '?', '?', '?', '?', 'z', '^', 'L', '~', 'n', 'o', //
    '|', '?', '?', '?', '?', '?', '?', '?', '?', '?', '?', '?', '?', '?', '?', '?', //
    '?', '?', '?', '?', '?', '?', '?', '?', '?', '?', '?', #00, '?', #00, #00, 'm', //
    'y', '?', '?', '?', '?', '?', '?', '?', '?', '?', '?', '?', '?', '?', '?', '?', //
    '?', '?', '?', '?', '?', '?', '?', '?', '?', '?', '?', #00, 'j', #00, '?', #00, //
    #00, #00, #00, #00, #00, #00, #00, #00, #00, #00, #00, #00, #00, #00, #00, #00, //
    #00, #00, #00, #00, #00, #00, #00, #00, #00, #00, #00, #00, #00, #00, #00, #00, //
    #00, #00, #00, #00, #00, #00, #00, #00, #00, #00, #00, #00, #00, #00, #00, #00, //
    #00, #00, #00, #00, #00, #00, #00, #00, #00, #00, #00, #00, #00, #00, #00, #00, //
    #00, #00, #00, #00, #00, #00, #00, #00, #00, #00, #00, #00, #00, #00, #00, #00, //
    #00, #00, #00, #00, #00, #00, #00, #00, #00, #00, #00, #00, #00, #00, #00, #00, //
    #00, #00, #00, #00, #00, #00, #00, #00, #00, #00, #00, #00, #00, #00, #00, #00, //
    #00, #00, #00, #00, #00, #00, #00, #00, #00, #00, #00, #00, #00, #00, 'O', #00  //
    );

type
  TLongPoint = record
    x, y: LongInt;
  end;

  TCaretStyle         = (csFull, csLeftLine, csBottomLine);
  TOffsetDisplayStyle = (odHex, odDec, odOctal, odNone);
  TTranslationType    = (ttAnsi, ttDos8, ttASCII, ttMac, ttEBCDIC);
  PUndoRec            = ^TUndoRec;

  TUndoRec = packed record
    Typ: Byte;
    Changed: Boolean;
    Modified: Boolean;
    CurPos: Integer;
    C1st: Byte;
    CharField: Boolean;
    SelS, SelE, SelP, Pos, Count, ReplCount: DWORD;
    Buffer: Byte;
  end;

  TColors = class(TPersistent)
  private
    FOffset            : TColor;
    FOddColumn         : TColor;
    FOddInverted       : TColor;
    FEvenColumn        : TColor;
    FEvenInverted      : TColor;
    FParent            : TControl;
    FPositionBackground: TColor;
    FCursorFrame       : TColor;
    FBackground        : TColor;
    FChangedText       : TColor;
    FPositionText      : TColor;
    FChangedBackground : TColor;
  protected
    procedure SetBackground(const Value: TColor);
    procedure SetChangedBackground(const Value: TColor);
    procedure SetChangedText(const Value: TColor);
    procedure SetCursorFrame(const Value: TColor);
    procedure SetPositionBackground(const Value: TColor);
    procedure SetPositionText(const Value: TColor);
    procedure SetEvenColumn(const Value: TColor);
    procedure SetOddColumn(const Value: TColor);
    procedure SetOffset(const Value: TColor);
  public
    constructor Create(Parent: TControl);
  published
    property Background        : TColor read FBackground write SetBackground;
    property PositionBackground: TColor read FPositionBackground write SetPositionBackground;
    property PositionText      : TColor read FPositionText write SetPositionText;
    property ChangedBackground : TColor read FChangedBackground write SetChangedBackground;
    property ChangedText       : TColor read FChangedText write SetChangedText;
    property CursorFrame       : TColor read FCursorFrame write SetCursorFrame;
    property Offset            : TColor read FOffset write SetOffset;
    property OddColumn         : TColor read FOddColumn write SetOddColumn;
    property EvenColumn        : TColor read FEvenColumn write SetEvenColumn;
  end;

  THexEditor = class(TCustomGrid)
  private
    fCharWidth, fCharHeight: Integer;
    fInsertOn              : Boolean;
    fCaretBitmap           : TBitmap;
    fColors                : TColors;
    fBytesPerLine          : Integer;
    fOffSetDisplayWidth    : Integer;
    fBPL2                  : Integer;
    fDataSize              : Integer;
    fIntFile               : TFileStream;
    fSwapNibbles           : Integer;
    fFocusFrame            : Boolean;
    fUndoMem               : TMemoryStream;
    fReadOnly              : Boolean;
    fBytesPerColumn        : Integer;
    fPosInChars            : Boolean;
    fIntBuffer             : PByteArray;
    fIntBufferPos          : Integer;
    fFileName              : string;
    fInternalName          : string;
    fChangedBytes          : TBits;
    fMarker                : array [0 .. 9] of Integer;
    fSelST, fSelPO, fSelEN : Integer;
    fIsSelecting           : Boolean;
    fCanUndo               : Boolean;
    fUndoDesc              : string;
    fUndoCount             : Integer;
    fStateChanged          : TNotifyEvent;
    fTranslation           : TTranslationType;
    fModified              : Boolean;
    fCreateBackup          : Boolean;
    fBackupExt             : string;
    FOffsetDisplay         : TOffsetDisplayStyle;
    FOffsetChar            : Char;
    fCaretStyle            : TCaretStyle;
    fShowMarkerCol         : Boolean;
    fLastKeyWasALT         : Boolean;
    fMaskWhiteSpaces       : Boolean;
    fMaskChar              : Char;
    fNoSizeChange          : Boolean;
    fVariableLineLength    : Boolean;
    fOffsets               : TList;
    fAllowInsertMode       : Boolean;
    fAutoCaretMode         : Boolean;
    property Color;
    procedure InternalErase(const BackSp: Boolean);
  protected
    procedure SetCaretStyle(const Value: TCaretStyle);
    procedure SetFocusFrame(const Value: Boolean);
    procedure SetBytesPerColumn(const Value: Integer);
    procedure SetSwapNibbles(const Value: Boolean);
    function GetSwapNibbles: Boolean;
    function GetBytesPerColumn: Integer;
    procedure SetShowMarkerColumn(const Value: Boolean);
    procedure SetOffsetDisplayWidth;
    procedure SetOffsetChar(const Value: Char);
    procedure SetOffsetDisplay(const Value: TOffsetDisplayStyle);
    procedure CreateColoredCaret;
    procedure SetColors(const Value: TColors);
    function GetMemory(aIndex: Integer): Char;
    procedure SetMemory(aIndex: Integer; aChar: Char);
    procedure SetReadOnly(aValue: Boolean);
    procedure TestStream;
    procedure AdjustMetrics;
    function GetDataSize: Integer;
    procedure CalcSizes;
    procedure DrawCell(ACol, ARow: LongInt; ARect: TRect; AState: TGridDrawState); override;
    function SelectCell(ACol, ARow: LongInt): Boolean; override;
    function GetPosAtCursor(const ACol, ARow: Integer): Integer;
    function GetCursorAtPos(aPos: Integer; aChars: Boolean): TLongPoint;
    function GetOtherFieldCol(ACol: Integer; var Chars: Boolean): Integer;
    function CheckSelectCell(ACol, ARow: Integer): Boolean;
    procedure WMChar(var Msg: TWMChar); message WM_CHAR;
    procedure WMSTATECHANGED(var Msg: TMessage); message WM_STATECHANGED;
    function GetByteAtPos(aPos: Integer): Byte;
    procedure SetByteAtPos(aPos: Integer; aByte: Byte);
    procedure GetMemAtPos(aBuffer: PByteArray; aPos, aCount: Integer);
    procedure SetMemAtPos(aBuffer: PByteArray; aPos, aCount: Integer);
    procedure ChangeByte(aOldByte, aNewByte: Byte; aPos, ACol, ARow: Integer);
    procedure KeyDown(var Key: Word; Shift: TShiftState); override;
    procedure KeyUp(var Key: Word; Shift: TShiftState); override;
    function HasChanged(aPos: Integer): Boolean;
    function IsMarkerPos(aPos: Integer): Integer;
    function GetMarkerRow(aWhich: Byte): Integer;
    function ParseKeyDown(aShift: TShiftState; aChar: Char): Boolean;
    function IsSelected(aPos: Integer): Boolean;
    procedure RedrawPos(aFrom, aTo: Integer);
    procedure ResetSelection(aDraw: Boolean);
    procedure ResetUndo;
{$IFDEF _debug}
    procedure ShowSelState;
{$ENDIF}
    procedure Select(aCurCol, aCurRow, aNewCol, aNewRow: Integer);
    procedure MouseDown(Button: TMouseButton; Shift: TShiftState; x, y: Integer); override;
    function CreateUndo(aType: Integer; aPos, aCount, aReplCount: Integer): Boolean;
    procedure DoCreateUndo(aType: Integer; aPos, aCount, aReplCount: Integer);
    function GetSelStart: Integer;
    function GetSelEnd: Integer;
    function GetSelCount: Integer;
    procedure SetSelStart(aValue: Integer);
    procedure SetSelEnd(aValue: Integer);
    procedure SetInCharField(aValue: Boolean);
    function GetInCharField: Boolean;
    procedure Loaded; override;
    procedure CreateWnd; override;
    procedure WMSetFocus(var Msg: TWMSetFocus); message WM_SETFOCUS;
    procedure WMKillFocus(var Msg: TWMKillFocus); message WM_KILLFOCUS;
    procedure SetTranslation(aValue: TTranslationType);
    function TranslateToAnsiChar(aByte: Byte): Char;
    function TranslateFromAnsiChar(aByte: Byte): Char;
    procedure SetModified(aValue: Boolean);
    procedure SetBytesPerLine(aValue: Integer);
    procedure SetChanged(aPos: Integer; aValue: Boolean);
    procedure InternalInsertBuffer(aBuffer: PChar; aSize, aPos: Integer);
    procedure InternalAppendBuffer(aBuffer: PChar; aSize: Integer);
    procedure MoveFileMem(aFrom, aTo, aCount: Integer);
    procedure CheckInternalBuffer(aPos: Integer);
    procedure SetInternalBufferByte(aPos: Integer; aByte: Byte);
    procedure InternalGetCurSel(var aSP, aEP, ACol, ARow: Integer);
    procedure InternalDeleteSelection(aSP, aEP, aNCol, aNRow: Integer);
    procedure WMVScroll(var Msg: TWMVScroll); message WM_VSCROLL;
    procedure WMHScroll(var Msg: TWMHScroll); message WM_HSCROLL;
    function InternalDeleteNibble(const aPos: Integer; const HighNibble: Boolean): Boolean;
    function InternalInsertNibble(const aPos: Integer; const HighNibble: Boolean): Boolean;
    function CreateShift4BitStream(const aStart: Integer; var vName: TFileName): TFileStream;
    procedure InternalConvertRange(const aFrom, aTo: Integer; const aTransFrom, aTransTo: TTranslationType);
    function GetMarker(aIndex: Byte): Integer;
    procedure SetMarker(aIndex: Byte; const aValue: Integer);
    procedure SetMaskWhiteSpaces(const aValue: Boolean);
    procedure SetMaskChar(const aValue: Char);
    procedure SetAsText(const aValue: string);
    procedure SetAsHex(const aValue: string);
    function GetAsText: string;
    function GetAsHex: string;
    procedure FreeFile;
    procedure SetVariableLineLength(const aValue: Boolean);
    procedure AdjustLineLengthsCount;
    function GetLineLength(aLine: Integer): Integer;
    procedure SetLineLength(aLine, aLength: Integer);
    function GetLineOffset(aLine: Integer): Integer;
    function OutOfBounds(const ACol, ARow: Integer): Boolean;
    function GetRow(const aPos: Integer): Integer;
    procedure SetNoSizeChange(const aValue: Boolean);
    procedure SetAllowInsertMode(const aValue: Boolean);
    function GetIsInsertMode: Boolean;
    procedure SetAutoCaretMode(const aValue: Boolean);
    procedure StateNotification;
  public
    constructor Create(aOwner: TComponent); override;
    destructor Destroy; override;
{$IFDEF _debug}
    procedure SaveUndo(aFileName: TFileName);
{$ENDIF}
    function Seek(const aOffset, aOrigin: Integer; const FailIfOutOfRange: Boolean): Boolean;
    function Find(aBuffer: PChar; const aCount, aStart, aEnd: Integer; const IgnoreCase, SearchText: Boolean): Integer;
    procedure DeleteSelection;
    function LoadFromStream(Strm: TStream): Boolean;
    function LoadFromFile(const Filename: string): Boolean;
    function SaveToStream(Strm: TStream): Boolean;
    function SaveToFile(const Filename: string): Boolean;
    function Undo: Boolean;
    procedure CreateEmptyFile(const TempName: string);
    function BufferFromFile(aPos: Integer; var aCount: Integer): PChar;
    procedure InsertBuffer(aBuffer: PChar; aSize, aPos: Integer);
    procedure AppendBuffer(aBuffer: PChar; aSize: Integer);
    procedure ReplaceSelection(aBuffer: PChar; aSize: Integer);
    function GetCursorPos: Integer;
    function DeleteNibble(const aPos: Integer; const HighNibble: Boolean): Boolean;
    function InsertNibble(const aPos: Integer; const HighNibble: Boolean): Boolean;
    procedure ConvertRange(const aFrom, aTo: Integer; const aTransFrom, aTransTo: TTranslationType);
    procedure ClearOffsets;
    procedure SetLineLengths(aLengths: TList);
    property SelStart: Integer read GetSelStart write SetSelStart;
    property SelEnd: Integer read GetSelEnd write SetSelEnd;
    property SelCount: Integer read GetSelCount;
    property CanUndo: Boolean read fCanUndo;
    property InCharField: Boolean read GetInCharField write SetInCharField;
    property UndoDescription: string read fUndoDesc;
    property ReadOnly: Boolean read fReadOnly write SetReadOnly;
    property Modified: Boolean read fModified write SetModified;
    property DataSize: Integer read GetDataSize;
    property Data[Index: Integer]: Char read GetMemory write SetMemory;
    property AsText: string read GetAsText write SetAsText;
    property AsHex: string read GetAsHex write SetAsHex;
    property Canvas;
    property Col;
    property LeftCol;
    property Row;
    property TabStops;
    property TopRow;
    property Filename: string read fFileName;
    property Marker[Index: Byte]: Integer read GetMarker write SetMarker;
    property VariableLineLength: Boolean read fVariableLineLength write SetVariableLineLength;
    property LineLength[Index: Integer]: Integer read GetLineLength write SetLineLength;
    property LineOffset[Index: Integer]: Integer read GetLineOffset;
    property IsInsertMode: Boolean read GetIsInsertMode;
  published
    property ShowMarkerColumn: Boolean read fShowMarkerCol write SetShowMarkerColumn default True;
    property BytesPerColumn  : Integer read GetBytesPerColumn write SetBytesPerColumn default 4;
    property OnStateChanged  : TNotifyEvent read fStateChanged write fStateChanged;
    property Translation     : TTranslationType read fTranslation write SetTranslation;
    property CreateBackup    : Boolean read fCreateBackup write fCreateBackup default True;
    property BackupExtension : string read fBackupExt write fBackupExt;
    property Align;
    property BorderStyle;
    property OffsetDisplay: TOffsetDisplayStyle read FOffsetDisplay write SetOffsetDisplay;
    property BytesPerLine : Integer read fBytesPerLine write SetBytesPerLine;
    property CaretStyle   : TCaretStyle read fCaretStyle write SetCaretStyle default csFull;
    property Colors       : TColors read fColors write SetColors;
    property Ctl3D;
    property DragCursor;
    property DragMode;
    property Enabled;
    property FocusFrame: Boolean read fFocusFrame write SetFocusFrame;
    property Font;
    property GridLineWidth default 0;
    property OffsetSeparator: Char read FOffsetChar write SetOffsetChar;
    property SwapNibbles    : Boolean read GetSwapNibbles write SetSwapNibbles default False;
    property MaskWhiteSpaces: Boolean read fMaskWhiteSpaces write SetMaskWhiteSpaces default True;
    property MaskChar       : Char read fMaskChar write SetMaskChar default '.';
    property NoSizeChange   : Boolean read fNoSizeChange write SetNoSizeChange default False;
    property AllowInsertMode: Boolean read fAllowInsertMode write SetAllowInsertMode default False;
    property AutoCaretMode  : Boolean read fAutoCaretMode write SetAutoCaretMode default True;
    property ParentCtl3D;
    property ParentShowHint;
    property PopupMenu;
    property ScrollBars;
    property ShowHint;
    property TabOrder;
    property TabStop;
    property Visible;
    property OnClick;
    property OnDblClick;
    property OnDragDrop;
    property OnDragOver;
    property OnEndDrag;
    property OnEnter;
    property OnExit;
    property OnKeyDown;
    property OnKeyPress;
    property OnKeyUp;
    property OnMouseDown;
    property OnMouseMove;
    property OnMouseUp;
    property OnStartDrag;
{$IFDEF ver120}
    property Anchors;
    property BiDiMode;
    property Constraints;
    property DragKind;
    property ParentBiDiMode;
    property OnEndDock;
    property OnMouseWheel;
    property OnMouseWheelDown;
    property OnMouseWheelUp;
    property OnStartDock;
{$ENDIF}
  end;

  THexToCanvas = class(TComponent)
  private
    fHexEditor                      : THexEditor;
    fFont                           : TFont;
    fTopM, fLeftM, fBottomM, fRightM: Integer;
    fLpP, fBpL, fBpC                : Integer;
    fOffsDy, fMemDy                 : TOffsetDisplayStyle;
    fCharDy                         : Boolean;
    fOffsCr, fMemCr, fCharCr        : Char;
    fShrink, fStretch               : Boolean;
    fSwapNibbles                    : Boolean;
    procedure SetFont(Value: TFont);
    procedure SetHexEditor(Value: THexEditor);
  protected
    procedure Notification(aComponent: TComponent; aOperation: TOperation); Override;
  public
    Constructor Create(aOwner: TComponent); override;
    Destructor Destroy; override;
    function Draw(aCanvas: TCanvas; const aStart, aEnd: Integer; const TopLine, BottomLine: string): Integer;
    procedure GetLayout;
    property TopMargin: Integer read fTopM write fTopM;
    property LeftMargin: Integer read fLeftM write fLeftM;
    property RightMargin: Integer read fRightM write fRightM;
    property BottomMargin: Integer read fBottomM write fBottomM;
    property LinesPerPage: Integer read fLpP;
  published
    property hexeditor         : THexEditor read fHexEditor write SetHexEditor;
    property Font              : TFont read fFont write SetFont;
    property BytesPerLine      : Integer read fBpL write fBpL default 16;
    property OffsetDisplay     : TOffsetDisplayStyle read fOffsDy write fOffsDy default odHex;
    property OffsetSeparator   : Char read fOffsCr write fOffsCr default ':';
    property MemFieldDisplay   : TOffsetDisplayStyle read fMemDy write fMemDy default odHex;
    property MemFieldSeparator : Char read fMemCr write fMemCr default ';';
    property DisplayCharField  : Boolean read fCharDy write fCharDy default True;
    property CharFieldSeparator: Char read fCharCr write fCharCr default #0;
    property ShrinkToFit       : Boolean read fShrink write fShrink default True;
    property StretchToFit      : Boolean read fStretch write fStretch default True;
    property BytesPerColumn    : Integer read fBpC write fBpC default 2;
    property SwapNibbles       : Boolean read fSwapNibbles write fSwapNibbles default False;
  end;

function Min(a1, a2: Integer): Integer;
function Max(a1, a2: Integer): Integer;
function LongPoint(aX, aY: LongInt): TLongPoint;
function IsKeyDown(aKey: Integer): Boolean;
procedure TranslateBufferFromAnsi(const TType: TTranslationType; aBuffer, bBuffer: PChar; const aCount: Integer);
procedure TranslateBufferToAnsi(const TType: TTranslationType; aBuffer, bBuffer: PChar; const aCount: Integer);
function ConvertHexToBin(aFrom, aTo: PChar; const aCount: Integer; const SwapNibbles: Boolean; var BytesTranslated: Integer): PChar;
function ConvertBinToHex(aFrom, aTo: PChar; const aCount: Integer; const SwapNibbles: Boolean): PChar;
function IntToOctal(const aValue: Integer): string;
{$IFDEF VER120}
function Int64ToOctal(const aValue: Int64): string;
{$ENDIF}
procedure Register;

implementation

const
  U_Byte_changed      = 0;
  U_Byte_removed      = 1;
  U_Insert_buffer     = 2;
  U_Replace_Selection = 3;
  U_Append_Buffer     = 4;
  U_Nibble_Insert     = 5;
  U_Nibble_Delete     = 6;
  U_Convert           = 7;

  UndoSTR: array [U_Byte_changed .. U_Convert] of string = ( //
    'Byte changed',                                          //
    'Byte(s) removed',                                       //
    'Insert buffer',                                         //
    'Replace selection',                                     //
    'Append buffer',                                         //
    'Insert nibble',                                         //
    'Delete nibble',                                         //
    'Convert'                                                //
    );

  cMax_Undo = 100;
  cBuf_Size = 65536;
  HexCHL    = '0123456789abcdef';
  HexCHU    = '0123456789ABCDEF';
  HexCHA    = HexCHL + HexCHU;

function Invert(Color: TColor): TColor;
begin
  Result := RGB(255 - GetRValue(Color), 255 - GetGValue(Color), 255 - GetBValue(Color));
end;

procedure TranslateBufferFromAnsi(const TType: TTranslationType; aBuffer, bBuffer: PChar; const aCount: Integer);
var
  pct: Integer;
  pch: AnsiChar;
begin
  case TType of
    ttAnsi:
      Move(aBuffer^, bBuffer^, aCount);
    ttDos8, ttASCII:
      CharToOEMBuff(aBuffer, PAnsiChar(bBuffer), aCount);
    ttMac:
      if aCount > 0 then
        for pct := 0 to Pred(aCount) do
        begin
          pch := AnsiChar(aBuffer[pct]);
          if pch < #128 then
            bBuffer[pct] := Char(pch)
          else
            bBuffer[pct] := Char(ctISOToMac[Ord(pch)]);
        end;
    ttEBCDIC:
      if aCount > 0 then
        for pct        := 0 to Pred(aCount) do
          bBuffer[pct] := Char(ctISOToEBCDIC[Ord(aBuffer[pct])]);
  end;
end;

procedure TranslateBufferToAnsi(const TType: TTranslationType; aBuffer, bBuffer: PChar; const aCount: Integer);
var
  pct: Integer;
  pch: AnsiChar;
begin
  case TType of
    ttAnsi:
      Move(aBuffer^, bBuffer^, aCount);
    ttDos8, ttASCII:
      OEMToCharBuff(PAnsiChar(aBuffer), bBuffer, aCount);
    ttMac:
      if aCount > 0 then
        for pct := 0 to Pred(aCount) do
        begin
          pch := AnsiChar(aBuffer[pct]);
          if pch < #128 then
            bBuffer[pct] := Char(pch)
          else
            bBuffer[pct] := Char(ctMacToISO[Ord(pch)]);
        end;
    ttEBCDIC:
      if aCount > 0 then
        for pct        := 0 to Pred(aCount) do
          bBuffer[pct] := Char(ctEBCDICToISO[Ord(aBuffer[pct])]);
  end;
end;

function FillLeft(const FillChar: Char; const IntStr: string; const MaxLen: Integer): string;
begin
  Result := IntStr;
  while Length(Result) < MaxLen do
    Result := FillChar + Result;
end;

function OEMChar(aByte: Byte): Char;
var
  psr: string;
begin
  psr := Char(aByte) + #0;
  OEMToChar(@psr[1], @psr[1]);
  Result := psr[1];
end;

function CharOEM(aByte: Byte): Char;
var
  psr: string;
begin
  psr := Char(aByte) + #0;
  CharToOEM(@psr[1], @psr[1]);
  Result := psr[1];
end;

procedure Register;
begin
  RegisterComponents('Merkes'' Pages', [THexEditor, THexToCanvas]);
end;

function GetTempName: string;
var
  pPT: string;
begin
  SetLength(pPT, MAX_PATH + 1);
  SetLength(pPT, GetTempPath(MAX_PATH, @pPT[1]));
  pPT := Trim(pPT);
  if pPT[Length(pPT)] <> '\' then
    pPT := pPT + '\';
  repeat
    Result := pPT + IntToHex(GetTickCount, 8) + '.MPHT';
  until not FileExists(Result);
end;

function CanOpenFile(const aName: TFileName; var Readonly: Boolean): Boolean;
var
  fHandle: THandle;
begin
  Result   := False;
  ReadOnly := True;
  if FileExists(aName) then
  begin
    fHandle := FileOpen(aName, fmOpenRead or fmShareDenyNone);
    if fHandle <> INVALID_HANDLE_VALUE then
    begin
      FileClose(fHandle);
      Result  := True;
      fHandle := FileOpen(aName, fmOpenReadWrite);
      if fHandle <> INVALID_HANDLE_VALUE then
      begin
        FileClose(fHandle);
        ReadOnly := False;
      end;
    end;
  end;
end;

function IsKeyDown(aKey: Integer): Boolean;
begin
  Result := (GetKeyState(aKey) and (not 1)) <> 0;
end;

function Min(a1, a2: Integer): Integer;
begin
  if a1 < a2 then
    Result := a1
  else
    Result := a2;
end;

function Max(a1, a2: Integer): Integer;
begin
  if a1 > a2 then
    Result := a1
  else
    Result := a2;
end;

function LongPoint(aX, aY: LongInt): TLongPoint;
begin
  Result.x := aX;
  Result.y := aY;
end;

function ConvertHexToBin(aFrom, aTo: PChar; const aCount: Integer; const SwapNibbles: Boolean; var BytesTranslated: Integer): PChar;
var
  lHi: Boolean;
  lCT: Integer;
  lBy: Byte;
  lNb: Char;
begin
  Result          := aTo;
  BytesTranslated := 0;
  lHi             := True;
  lBy             := 0;
  for lCT         := 0 to Pred(aCount) do
    if Pos(aFrom[lCT], HexCHA) <> 0 then
    begin
      lNb := UpCase(aFrom[lCT]);
      if lHi then
        lBy := ((Pos(lNb, HexCHU) - 1) * 16)
      else
        lBy := lBy or ((Pos(lNb, HexCHU) - 1));
      lHi   := not lHi;
      if lHi then
      begin
        if SwapNibbles then
          aTo[BytesTranslated] := Char(((lBy and 15) * 16) or ((lBy and $F0) shr 4))
        else
          aTo[BytesTranslated] := Char(lBy);
        Inc(BytesTranslated);
      end;
    end;
end;

function ConvertBinToHex(aFrom, aTo: PChar; const aCount: Integer; const SwapNibbles: Boolean): PChar;
var
  lCT: Integer;
  lBy: Byte;
  lCX: Integer;
begin
  Result  := aTo;
  lCX     := 0;
  for lCT := 0 to Pred(aCount) do
  begin
    lBy := Ord(aFrom[lCT]);
    if SwapNibbles then
    begin
      aTo[lCX]     := UpCase(HexCHU[(lBy and 15) + 1]);
      aTo[lCX + 1] := UpCase(HexCHU[(lBy shr 4) + 1])
    end
    else
    begin
      aTo[lCX + 1] := UpCase(HexCHU[(lBy and 15) + 1]);
      aTo[lCX]     := UpCase(HexCHU[(lBy shr 4) + 1])
    end;
    Inc(lCX, 2);
  end;
  aTo[lCX] := #0;
end;

const
  gOctalChars = '01234567';

function IntToOctal(const aValue: Integer): string;
var
  lVal: Integer;
begin
  Result := '';
  lVal   := aValue;
  repeat
    Result := gOctalChars[(lVal mod 8) + 1] + Result;
    lVal   := lVal shr 3;
  until lVal = 0;
  Result := '0' + Result;
end;
{$IFDEF VER120}

function Int64ToOctal(const aValue: Int64): string;
var
  lVal: Int64;
begin
  Result := '';
  lVal   := aValue;
  repeat
    Result := gOctalChars[(lVal mod 8) + 1] + Result;
    lVal   := lVal shr 3;
  until lVal = 0;
  Result := '0' + Result;
end;
{$ENDIF}

constructor THexEditor.Create(aOwner: TComponent);
begin
  inherited Create(aOwner);
  fColors             := TColors.Create(Self);
  ParentColor         := False;
  fIntFile            := nil;
  fUndoMem            := nil;
  Color               := fColors.Background;
  fCharWidth          := -1;
  fShowMarkerCol      := True;
  fOffSetDisplayWidth := -1;
  fBytesPerLine       := 16;
  FOffsetChar         := ':';
  FOffsetDisplay      := odHex;
  fCaretStyle         := csFull;
  fFocusFrame         := True;
  fSwapNibbles        := 0;
  fFileName           := '---';
  Font.Name           := 'Courier';
  Font.Size           := 12;
  BorderStyle         := bsSingle;
  DefaultDrawing      := False;
  Options             := [goVertLine, goHorzLine, goTabs, gOThumbTracking];
  GridLineWidth       := 0;
  fBytesPerColumn     := 4;
  Ctl3D               := False;
  Cursor              := crIBeam;
  fChangedBytes       := TBits.Create;
  FillChar(fMarker[0], SizeOf(fMarker), $FF);
  fSelST       := -1;
  fSelPO       := -1;
  fSelEN       := -1;
  fIsSelecting := False;
  ResetUndo;
  DefaultColWidth  := 0;
  DefaultRowHeight := 0;
  ColCount         := fBytesPerLine * 3 + 3;
  RowCount         := 1;
  fTranslation     := ttAnsi;
  fCanUndo         := False;
  fModified        := False;
  fReadOnly        := True;
  fCreateBackup    := True;
  fBackupExt       := '.bak';
  fInternalName    := GetTempName;
  fIntBufferPos    := -1;
  GetMem(fIntBuffer, cBuf_Size);
  fDataSize           := -1;
  fBPL2               := 2 * fBytesPerLine;
  fLastKeyWasALT      := False;
  fMaskWhiteSpaces    := True;
  fMaskChar           := '.';
  fCaretBitmap        := TBitmap.Create;
  fNoSizeChange       := False;
  fVariableLineLength := False;
  fOffsets            := TList.Create;
  fAllowInsertMode    := False;
  fInsertOn           := False;
  fAutoCaretMode      := True;
end;

procedure THexEditor.FreeFile;
begin
  if fIntFile <> nil then
  begin
    fIntFile.Size := 0;
    fIntFile.Free;
    fIntFile := nil;
  end;
end;

destructor THexEditor.Destroy;
begin
  FreeFile;
  if fUndoMem <> nil then
  begin
    fUndoMem.Size := 0;
    fUndoMem.Free;
    fUndoMem := nil;
  end;
  fChangedBytes.Free;
  if FileExists(fInternalName) then
    DeleteFile(fInternalName);
  FreeMem(fIntBuffer, cBuf_Size);
  fColors.Free;
  fCaretBitmap.Free;
  fOffsets.Clear;
  fOffsets.Free;
  inherited Destroy;
end;

procedure THexEditor.AdjustMetrics;
var
  pct: Integer;
begin
  Canvas.Font.Assign(Font);
  fCharWidth := Canvas.TextWidth('w');
  SetOffsetDisplayWidth;
  ColWidths[1] := fCharWidth * Integer(fShowMarkerCol);
  for pct      := 0 to Pred(fBytesPerLine * 2) do
    if (((pct + 2) mod fBytesPerColumn) = 1) then
      ColWidths[pct + 2] := fCharWidth * 2
    else
      ColWidths[pct + 2] := fCharWidth;
  for pct                := fBytesPerLine * 2 to (fBytesPerLine * 3) do
    ColWidths[pct + 2]   := fCharWidth;
  fCharHeight            := Canvas.TextHeight('yY') + 2;
  DefaultRowHeight       := fCharHeight;
end;

function THexEditor.GetDataSize: Integer;
begin
  Result := fDataSize;
  if (fDataSize = -1) or (fIntBufferPos = -1) then
  begin
    if fIntFile = nil then
      Result := 0
    else
      Result  := fIntFile.Size;
    fDataSize := Result
  end
end;

procedure THexEditor.CreateEmptyFile;
begin
  FreeFile;
  fFileName := TempName;
  ResetUndo;
  ResetSelection(False);
  fChangedBytes.Size := 0;
  CalcSizes;
  fModified := False;
  fReadOnly := True;
  MoveColRow(2, 0, True, True);
end;

function THexEditor.SaveToStream(Strm: TStream): Boolean;
var
  MemStrm: TMemoryStream;
  pCr    : TCursor;
begin
  Result  := True;
  pCr     := Cursor;
  Cursor  := crHourGlass;
  MemStrm := TMemoryStream.Create;
  try
    try
      fIntFile.Position := 0;
      MemStrm.LoadFromStream(fIntFile);
      MemStrm.SaveToStream(Strm);
    except
      Result := False;
    end;
  finally
    MemStrm.Free;
    Cursor := pCr;
  end;
end;

function THexEditor.SaveToFile(const Filename: string): Boolean;
var
  pCr: TCursor;
begin
  Result := True;
  pCr    := Cursor;
  Cursor := crHourGlass;
  try
    if fCreateBackup and fModified and (fFileName = Filename) then
      if not CopyFile(PChar(Filename), PChar(ChangeFileExt(Filename, fBackupExt)), False) then
        Exit;
    try
      fIntFile.Free;
      Result := CopyFile(PChar(fInternalName), PChar(Filename), False);
    except
      Result := False;
    end;
    fIntFile := TFileStream.Create(fInternalName, fmOpenReadWrite);
    if Result then
    begin
      fChangedBytes.Size := 0;
      fModified          := False;
      fReadOnly          := False;
      fFileName          := Filename;
      ResetUndo;
    end;
  finally
    Cursor := pCr;
    Invalidate;
  end;
end;

function THexEditor.LoadFromStream(Strm: TStream): Boolean;
var
  pCr    : TCursor;
  MemStrm: TMemoryStream;
begin
  FreeFile;
  pCr     := Cursor;
  Cursor  := crHourGlass;
  MemStrm := TMemoryStream.Create;
  try
    Strm.Position := 0;
    MemStrm.LoadFromStream(Strm);
    MemStrm.SaveToFile(fInternalName);
  finally
    MemStrm.Free;
  end;
  SetFileAttributes(PChar(fInternalName), 0);
  fIntFile := TFileStream.Create(fInternalName, fmOpenReadWrite);
  try
    fIntFile.Position := 0;
    Result            := True;
  finally
    Cursor := pCr;
    ResetUndo;
    fChangedBytes.Size := 0;
    CalcSizes;
    fModified    := False;
    fIsSelecting := False;
    MoveColRow(2, 0, True, True);
  end;
end;

function THexEditor.LoadFromFile(const Filename: string): Boolean;
var
  pCr: TCursor;
begin
  Result := True;
  if CanOpenFile(Filename, fReadOnly) then
  begin
    FreeFile;
    pCr    := Cursor;
    Cursor := crHourGlass;
    CopyFile(PChar(Filename), PChar(fInternalName), False);
    SetFileAttributes(PChar(fInternalName), 0);
    fIntFile := TFileStream.Create(fInternalName, fmOpenReadWrite);
    try
      fIntFile.Position := 0;
      fFileName         := Filename;
      Result            := True;
    finally
      Cursor := pCr;
      ResetUndo;
      fChangedBytes.Size := 0;
      CalcSizes;
      fModified    := False;
      fIsSelecting := False;
      MoveColRow(2, 0, True, True);
    end;
  end
end;

procedure THexEditor.CalcSizes;
  function CalcVarRowCount: Integer;
  var
    pct, pPos: Integer;
  begin
    pct  := DataSize div fBytesPerLine;
    pPos := 0;
    while pPos < DataSize do
    begin
      pPos := LineOffset[pct];
      Inc(pct);
    end;
    Result := Max(0, pct - 1);
  end;

begin
  fDataSize := -1;
  if fChangedBytes.Size > DataSize then
    fChangedBytes.Size := DataSize;
  if DataSize < 1 then
  begin
    FixedCols := 2;
    RowCount  := 1;
    ColCount  := fBytesPerLine * 3 + 3;
    if fOffsets.Count = 0 then
      LineLength[0] := fBytesPerLine;
  end
  else
  begin
    if not fVariableLineLength then
      RowCount := (DataSize + (fBytesPerLine - 1)) div fBytesPerLine
    else
      RowCount := CalcVarRowCount;
    ColCount   := fBytesPerLine * 3 + 3;
    FixedCols  := 2;
  end;
  FixedRows     := 0;
  fIntBufferPos := -1;
  AdjustMetrics;
end;

function THexEditor.TranslateFromAnsiChar(aByte: Byte): Char;
begin
  case fTranslation of
    ttAnsi:
      begin
        if aByte < 32 then
          Result := #0
        else
          Result := Char(aByte);
      end;
    ttDos8, ttASCII:
      begin
        if ((fTranslation = ttDos8) or (aByte < 128)) and (aByte > 31) then
          Result := CharOEM(aByte)
        else
          Result := #0;
      end;
    ttMac:
      begin
        if aByte < 32 then
          Result := #0
        else if aByte < 128 then
          Result := Char(aByte)
        else
          Result := Char(ctISOToMac[aByte]);
      end;
    ttEBCDIC:
      begin
        Result := Char(ctISOToEBCDIC[aByte]);
      end;
  else
    Result := #0;
  end;
end;

function THexEditor.TranslateToAnsiChar(aByte: Byte): Char;
begin
  case fTranslation of
    ttAnsi:
      begin
        Result := Char(aByte);
      end;
    ttDos8, ttASCII:
      begin
        Result := OEMChar(aByte);
        if ((fTranslation = ttASCII) and (aByte > 127)) then
          Result := fMaskChar;
      end;
    ttMac:
      begin
        if aByte < 128 then
          Result := Char(aByte)
        else
          Result := Char(ctMacToISO[aByte]);
      end;
    ttEBCDIC:
      begin
        Result := Char(ctEBCDICToISO[aByte]);
        if Result = #0 then
          Result := fMaskChar;
      end;
  else
    Result := fMaskChar;
  end;
  if fMaskWhiteSpaces and (Result < #32) then
    Result := fMaskChar;
end;

function THexEditor.OutOfBounds(const ACol, ARow: Integer): Boolean;
var
  pInCH  : Boolean;
  pMaxCol: Integer;
begin
  Result := False;
  if not fVariableLineLength then
    Exit;

  pInCH := ACol > (2 + fBPL2);
  if pInCH then
    pMaxCol := (fBytesPerLine * 2) + 2 + LineLength[ARow]
  else
    pMaxCol := (LineLength[ARow] * 2) + 1;
  Result    := (ACol > pMaxCol);
end;

procedure THexEditor.DrawCell(ACol, ARow: LongInt; ARect: TRect; AState: TGridDrawState);
var
  pTMP               : Boolean;
  pOddCol            : Boolean;
  pChan              : Boolean;
  pSZ                : Integer;
  pAP                : Integer;
  pCO                : string;
  pSFR               : string;
  pCanText, pCanBrush: TColor;
  procedure _TextOut;
  begin
    with Canvas do
    begin
      SetTextColor(Handle, ColorToRGB(pCanText));
      SetBKColor(Handle, ColorToRGB(pCanBrush));
      ExtTextOut(Handle, ARect.Left, ARect.Top, ETO_CLIPPED or ETO_OPAQUE, @ARect, PChar(pCO), Length(pCO), nil);
    end;
  end;

begin
  if (ARow = 0) and (DataSize < 1) then
  begin
    pCO := '   ';
    if ACol = 0 then
      case FOffsetDisplay of
        odHex:
          pCO := '0x0' + FOffsetChar;
        odDec:
          pCO := '0' + FOffsetChar;
        odOctal:
          pCO := 'o 0' + FOffsetChar;
      end;
    pCanBrush := fColors.Background;
    pCanText  := Colors.Offset;
    _TextOut;
    if ACol = 2 then
    begin
      SetCaretPos(ARect.Left, ARect.Top);
    end;
    Exit;
  end;

  pAP := LineOffset[ARow];
  if ACol = 0 then
  begin
    case FOffsetDisplay of
      odNone:
        pCO := ' ';
      odHex:
        pCO := '0x' + IntToHex(pAP, fOffSetDisplayWidth - 3) + FOffsetChar;
      odDec:
        pCO := FillLeft(' ', IntToStr(pAP), fOffSetDisplayWidth - 1) + FOffsetChar;
      odOctal:
        pCO := 'o ' + FillLeft('0', IntToOctal(pAP), fOffSetDisplayWidth - 3) + FOffsetChar;
    end;
    pCanBrush := fColors.Background;
    pCanText  := Colors.Offset;
    _TextOut;
    Exit;
  end;

  if (ACol = 1) then
  begin
    if (IsMarkerPos(ARow) > -1) then
    begin
      pCanText         := fColors.PositionText;
      pCanBrush        := fColors.PositionBackground;
      pSZ              := Canvas.Font.Size;
      pSFR             := Canvas.Font.Name;
      Canvas.Font.Name := 'Arial';
      Canvas.Font.Size := Round(pSZ * 0.75);
      pCO              := IntToStr(IsMarkerPos(ARow));
      _TextOut;
      Canvas.Font.Size := pSZ;
      Canvas.Font.Name := pSFR;
    end
    else
    begin
      pCanBrush := fColors.Background;
      pCanText  := Font.Color;
      pCO       := ' ';
      _TextOut;
    end;
    Exit;
  end;

  if (ACol = fBPL2 + 2) then
  begin
    pCanBrush := fColors.Background;
    pCanText  := Font.Color;
    pCO       := ' ';
    _TextOut;
    Exit;
  end;

  CheckInternalBuffer(pAP);
  pAP := GetPosAtCursor(ACol, ARow);
  if (pAP >= DataSize) or (fVariableLineLength and OutOfBounds(ACol, ARow)) then
  begin
    pCanBrush := fColors.Background;
    pCanText  := Font.Color;
    pCO       := ' ';
    _TextOut;
    Exit;
  end;

  if not fPosInChars then
  begin
    if ((ACol - 2) mod 2) = fSwapNibbles then
      pCO := HexCHU[fIntBuffer[pAP - fIntBufferPos] shr 4 + 1]
    else
      pCO := HexCHU[fIntBuffer[pAP - fIntBufferPos] and 15 + 1]
  end
  else
    pCO := TranslateToAnsiChar(fIntBuffer[pAP - fIntBufferPos]);

  pChan   := (HasChanged(pAP));
  pOddCol := (((ACol - 2) div fBytesPerColumn) mod 2) = 0;
  if pChan then
  begin
    pCanText  := fColors.ChangedText;
    pCanBrush := fColors.ChangedBackground;
  end
  else
  begin
    pCanBrush := fColors.Background;
    pCanText  := Font.Color;
    if not fPosInChars then
      if pOddCol then
        pCanText := Colors.OddColumn
      else
        pCanText := Colors.EvenColumn;
  end;

  if (fSelPO <> -1) and IsSelected(pAP) then
  begin
    pSZ       := pCanBrush;
    pCanBrush := pCanText;
    pCanText  := pSZ;
    if not(pChan or fPosInChars) then
      if pOddCol then
        pCanText := Colors.FOddInverted
      else
        pCanText := Colors.FEvenInverted;
  end;

  _TextOut;
  if ARow = Row then
  begin
    if (ACol = Col) then
    begin
      if Focused then
      begin
        SetCaretPos(ARect.Left, ARect.Top);
      end
    end
    else if (GetOtherFieldCol(Col, pTMP) = ACol) and Focused then
    begin
      if fFocusFrame then
        Canvas.DrawFocusRect(Rect(ARect.Left, ARect.Top, ARect.Left + fCharWidth, ARect.Bottom))
      else
      begin
        Canvas.Pen.Color   := fColors.CursorFrame;
        Canvas.Brush.Style := bsClear;
        Canvas.Rectangle(ARect.Left, ARect.Top, ARect.Left + fCharWidth, ARect.Bottom);
      end;
    end
  end;
end;
{$IFDEF _debug}

procedure THexEditor.ShowSelState;
begin
  if fIsSelecting then
  begin
    TForm(Owner).Caption := 'a'
  end
  else
    TForm(Owner).Caption := '-';
end;
{$ENDIF}

function THexEditor.SelectCell(ACol, ARow: LongInt): Boolean;
var
  pRow : Integer;
  pRect: TRect;
  pTMP : Boolean;
  pOC  : Integer;
begin
  pRow := Row;
{$IFDEF _debug}
  ShowSelState;
{$ENDIF}
  if DataSize > 0 then
    Result := CheckSelectCell(ACol, ARow)
  else
  begin
    if not((ACol = 2) and (ARow = 0)) then
      Result := False
    else
    begin
      Result := True;
      Exit;
    end;
  end;
  if Result then
  begin
    pOC   := GetOtherFieldCol(Col, pTMP);
    pRect := CellRect(pOC, pRow);
    InvalidateRect(Handle, @pRect, False);
    pOC   := GetOtherFieldCol(ACol, pTMP);
    pRect := CellRect(pOC, ARow);
    InvalidateRect(Handle, @pRect, False);
    if fIsSelecting then
      Select(Col, Row, ACol, ARow)
    else
      ResetSelection(True);
    pRect := CellRect(ACol, ARow);
    SetCaretPos(pRect.Left, pRect.Top);
  end;
end;

function THexEditor.GetPosAtCursor(const ACol, ARow: Integer): Integer;
begin
  fPosInChars := ACol > (2 + fBPL2);
  Result      := LineOffset[ARow];
  if fPosInChars then
    Result := Result + (ACol - (3 + fBPL2))
  else
    Result := Result + ((ACol - 2) div 2);
  if Result < 0 then
    Result := 0;
end;

function THexEditor.GetRow(const aPos: Integer): Integer;
var
  pct: Integer;
begin
  if not fVariableLineLength then
    Result := aPos div fBytesPerLine
  else
  begin
    Result  := 0;
    for pct := 0 to RowCount - 1 do
    begin
      if LineOffset[pct] > aPos then
      begin
        Result := pct - 1;
        Break;
      end;
      Result := RowCount - 1;
    end;
  end;
end;

function THexEditor.GetCursorAtPos(aPos: Integer; aChars: Boolean): TLongPoint;
var
  pct: Integer;
begin
  if aPos < 0 then
  begin
    Result.y := 0;
    Result.x := 2;
    Exit;
  end;
  Result.y := GetRow(aPos);
  if not fVariableLineLength then
    pct := aPos mod fBytesPerLine
  else
    pct := aPos - LineOffset[Result.y];
  if aChars then
    Result.x := pct + (3 + fBPL2)
  else
    Result.x := (pct * 2) + 2;
end;

function THexEditor.GetOtherFieldCol(ACol: Integer; var Chars: Boolean): Integer;
var
  pct: Integer;
begin
  Chars := ACol > (2 + fBPL2);
  if Chars then
  begin
    pct    := (ACol - (3 + fBPL2));
    Result := (pct * 2) + 2;
  end
  else
  begin
    pct    := ((ACol - 2) div 2);
    Result := pct + (3 + fBPL2);
  end;
end;

function THexEditor.CheckSelectCell(ACol, ARow: Integer): Boolean;
var
  pTP: TLongPoint;
{$J+}
const
  pCan: Boolean     = True;
  pClicked: Boolean = False;
{$J-}
begin
  Result := Inherited SelectCell(ACol, ARow);
  if (Result and fVariableLineLength and OutOfBounds(ACol, ARow)) then
    Result := False;
  if not pCan then
    Exit;
  try
    pCan := False;
    if Result then
    begin
      if not(IsKeyDown(VK_SHIFT) or IsKeyDown(VK_LBUTTON)) then
        ResetSelection(True);
      if GetPosAtCursor(ACol, ARow) >= DataSize then
      begin
        GetPosAtCursor(Col, Row);
        pTP := GetCursorAtPos(DataSize - 1, fPosInChars);
        MoveColRow(pTP.x, pTP.y, True, True);
        Result := False;
      end
      else if ACol = (2 + fBPL2) then
      begin
        Result := False;
        if IsKeyDown(VK_LBUTTON) then
        begin
          ACol := ACol - 1;
          ACol := Max(2, ACol);
          MoveColRow(ACol, ARow, True, True);
          Exit;
        end;
      end;
    end;
  finally
    pCan := True;
  end;
end;

procedure THexEditor.WMChar(var Msg: TWMChar);
var
  pPos          : Integer;
  pch           : Char;
  pOldBT, pNewBT: Byte;
  pTP           : TLongPoint;
begin
  pPos := GetPosAtCursor(Col, Row);
  if (pPos >= DataSize) and not IsInsertMode then
    Exit;
  if not fPosInChars then
  begin
    pch := Char(Msg.CharCode);
    if Pos(pch, HexCHA) <> 0 then
    begin
      pch := UpCase(pch);
      if not IsInsertMode then
        ResetSelection(True);
      pTP := GetCursorAtPos(pPos, fPosInChars);
      if DataSize > pPos then
        pOldBT := GetByteAtPos(pPos)
      else
        pOldBT := 0;
      if (pTP.x = (Col - fSwapNibbles)) or (SelCount <> 0) then
        pNewBT := pOldBT and 15 + ((Pos(pch, HexCHU) - 1) * 16)
      else
        pNewBT := (pOldBT and $F0) + (Pos(pch, HexCHU) - 1);
      if IsInsertMode and ((pTP.x = Col) or (SelCount > 0)) then
      begin
        if fSwapNibbles = 0 then
          pNewBT := pNewBT and $F0
        else
          pNewBT := pNewBT and $0F;
        if DataSize = 0 then
          AppendBuffer(@pNewBT, 1)
        else if SelCount = 0 then
          InsertBuffer(@pNewBT, 1, pPos)
        else
          ReplaceSelection(@pNewBT, 1);
      end
      else
      begin
        ChangeByte(pOldBT, pNewBT, pPos, Col, Row);
        if IsInsertMode and (pTP.x <> Col) and (pPos + 1 = DataSize) then
        begin
          pNewBT := 0;
          AppendBuffer(@pNewBT, 1);
          Exit;
        end;
      end;
      ParseKeyDown([], Char(VK_RIGHT));
    end;
  end
  else
  begin
    if not fLastKeyWasALT then
      pch := TranslateFromAnsiChar(Msg.CharCode)
    else
      pch := CharOEM(Msg.CharCode);
    if (pch < #32) and (not fLastKeyWasALT) then
      Exit;
    fLastKeyWasALT := False;
    if not IsInsertMode then
      ResetSelection(True);
    pTP    := GetCursorAtPos(pPos, fPosInChars);
    pOldBT := GetByteAtPos(pPos);
    if IsInsertMode then
    begin
      if SelCount > 0 then
        ReplaceSelection(@pch, 1)
      else
        InsertBuffer(@pch, 1, pPos)
    end
    else
      ChangeByte(pOldBT, Ord(pch), pPos, Col, Row);
    ParseKeyDown([], Char(VK_RIGHT));
  end;
end;

procedure THexEditor.SetByteAtPos(aPos: Integer; aByte: Byte);
begin
  fIntFile.Position := aPos;
  fIntFile.Write(aByte, SizeOf(Byte));
end;

function THexEditor.GetByteAtPos(aPos: Integer): Byte;
begin
  fIntFile.Position := aPos;
  fIntFile.Read(Result, SizeOf(Byte));
end;

procedure THexEditor.GetMemAtPos(aBuffer: PByteArray; aPos, aCount: Integer);
begin
  fIntFile.Position := aPos;
  fIntFile.Read(aBuffer^, aCount);
end;

procedure THexEditor.SetMemAtPos(aBuffer: PByteArray; aPos, aCount: Integer);
begin
  fIntFile.Position := aPos;
  fIntFile.Write(aBuffer^, aCount);
end;

procedure THexEditor.ChangeByte(aOldByte, aNewByte: Byte; aPos, ACol, ARow: Integer);
var
  pRect: TRect;
  pTMP : Boolean;
  pCol : Integer;
  pTP  : TLongPoint;
begin
  if aOldByte = aNewByte then
    Exit;
  if not CreateUndo(U_Byte_changed, aPos, 1, 0) then
    Exit;
  SetByteAtPos(aPos, aNewByte);
  SetInternalBufferByte(aPos, aNewByte);
  if not IsInsertMode then
    fChangedBytes.Bits[aPos] := True;
  pTP                        := GetCursorAtPos(aPos, False);
  ACol                       := pTP.x;
  pCol                       := GetOtherFieldCol(ACol, pTMP);
  pRect                      := BoxRect(ACol, ARow, ACol + 1, ARow);
  InvalidateRect(Handle, @pRect, False);
  pRect := BoxRect(pCol, ARow, pCol, ARow);
  InvalidateRect(Handle, @pRect, False);
end;

function THexEditor.ParseKeyDown(aShift: TShiftState; aChar: Char): Boolean;
  function CheckIfLastCol(const ACol, ARow: Integer): Boolean;
  begin
    Result := (not OutOfBounds(ACol, ARow)) and OutOfBounds(ACol + 1, ARow);
  end;
  function GetLastCol(const ACol, ARow: Integer): Integer;
  begin
    if ACol > (2 + fBPL2) then
      Result := 3 + fBPL2
    else
      Result := 2;
    while not((not OutOfBounds(Result, ARow)) and OutOfBounds(Result + 1, ARow)) do
      Inc(Result);
  end;

var
  pct     : Integer;
  pTP     : TLongPoint;
  pRow    : Integer;
  pLastCol: Boolean;
begin
  Result   := False;
  pLastCol := False;
  if not((aShift <> []) or (aChar = #16)) then
    if not IsInsertMode then
      ResetSelection(True);
  if aChar = Char(VK_PRIOR) then
  begin
    if fVariableLineLength then
      pLastCol := CheckIfLastCol(Col, Row);
    if ssCtrl in aShift then
    begin
      pRow := TopRow;
      pct  := Col;
      if pRow > -1 then
      begin
        if fVariableLineLength and pLastCol then
          pct := GetLastCol(pct, pRow)
        else
          while OutOfBounds(pct, pRow) do
            Dec(pct);
        MoveColRow(pct, pRow, True, True);
      end;
    end
    else
    begin
      pRow   := Max(0, Row - VisibleRowCount + 1);
      TopRow := Max(0, TopRow - VisibleRowCount + 1);
      pct    := Col;
      if pRow > -1 then
      begin
        if fVariableLineLength and pLastCol then
          pct := GetLastCol(pct, pRow)
        else
          while OutOfBounds(pct, pRow) do
            Dec(pct);
        MoveColRow(pct, pRow, True, True);
      end;
    end;
    Result := True;
  end;
  if aChar = Char(VK_NEXT) then
  begin
    if fVariableLineLength then
      pLastCol := CheckIfLastCol(Col, Row);
    if ssCtrl in aShift then
    begin
      pRow := Min(RowCount - 1, TopRow + VisibleRowCount - 1);
      pct  := Col;
      if pRow > -1 then
      begin
        if fVariableLineLength and pLastCol then
          pct := GetLastCol(pct, pRow)
        else
          while OutOfBounds(pct, pRow) do
            Dec(pct);
        MoveColRow(pct, pRow, True, True);
      end;
    end
    else
    begin
      pRow   := Min(RowCount - 1, Row + VisibleRowCount - 1);
      TopRow := Min(Max(0, RowCount - VisibleRowCount), TopRow + VisibleRowCount - 1);
      pct    := Col;
      if pRow > -1 then
      begin
        if fVariableLineLength and pLastCol then
          pct := GetLastCol(pct, pRow)
        else
          while OutOfBounds(pct, pRow) do
            Dec(pct);
        MoveColRow(pct, pRow, True, True);
      end;
    end;
    Result := True;
  end;
  if aChar = Char(VK_HOME) then
  begin
    GetPosAtCursor(Col, Row);
    if (ssCtrl in aShift) then
    begin
      if not fPosInChars then
        MoveColRow(2, 0, True, True)
      else
        MoveColRow(GetOtherFieldCol(2, fPosInChars), 0, True, True);
    end
    else
    begin
      if not fPosInChars then
        MoveColRow(2, Row, True, True)
      else
        MoveColRow(GetOtherFieldCol(2, fPosInChars), Row, True, True);
    end;
    Result := True;
  end;
  if aChar = Char(VK_END) then
  begin
    GetPosAtCursor(Col, Row);
    if (ssCtrl in aShift) then
    begin
      pTP := GetCursorAtPos(DataSize - 1, fPosInChars);
      MoveColRow(pTP.x, pTP.y, True, True)
    end
    else
    begin
      if not fPosInChars then
      begin
        pct := GetPosAtCursor(2, Row + 1) - 1;
        if pct >= DataSize then
          pct := DataSize - 1;
        pTP   := GetCursorAtPos(pct, fPosInChars);
        MoveColRow(pTP.x, pTP.y, True, True)
      end
      else
      begin
        pct := GetPosAtCursor(2, Row + 1) - 1;
        if pct >= DataSize then
          pct := DataSize - 1;
        pTP   := GetCursorAtPos(pct, True);
        MoveColRow(pTP.x, pTP.y, True, True)
      end
    end;
    Result := True;
  end;
  if (aChar = Char(VK_LEFT)) and (not(ssCtrl in aShift)) then
  begin
    pct := GetPosAtCursor(Col, Row) - 1;
    if fPosInChars then
    begin
      if pct < 0 then
        pct := 0;
      pTP   := GetCursorAtPos(pct, fPosInChars);
      MoveColRow(pTP.x, pTP.y, True, True);
    end
    else
    begin
      pct := pct + 1;
      pTP := GetCursorAtPos(pct, False);
      if pTP.x < Col then
        MoveColRow(Col - 1, Row, True, True)
      else
      begin
        pct := pct - 1;
        if pct >= 0 then
        begin
          pTP := GetCursorAtPos(pct, fPosInChars);
          MoveColRow(pTP.x + 1, pTP.y, True, True);
        end;
      end
    end;
    Result := True;
  end;
  if (aChar = Char(VK_RIGHT)) and (not(ssCtrl in aShift)) then
  begin
    pct := GetPosAtCursor(Col, Row) + 1;
    if fPosInChars then
    begin
      if pct >= DataSize then
        pct := DataSize - 1;
      pTP   := GetCursorAtPos(pct, fPosInChars);
      MoveColRow(pTP.x, pTP.y, True, True);
    end
    else
    begin
      pct := pct - 1;
      pTP := GetCursorAtPos(pct, False);
      if pTP.x = Col then
        MoveColRow(Col + 1, Row, True, True)
      else
      begin
        pct := pct + 1;
        if pct < DataSize then
        begin
          pTP := GetCursorAtPos(pct, fPosInChars);
          MoveColRow(pTP.x, pTP.y, True, True);
        end;
      end
    end;
    Result := True;
  end;
  if (aChar = Char(VK_RIGHT)) and (ssCtrl in aShift) then
  begin
    pct := ColCount - 1;
    while OutOfBounds(pct, Row) do
      Dec(pct);
    MoveColRow(pct, Row, True, True);
    Result := True;
  end;
  if (aChar = Char(VK_DOWN)) and (not(ssCtrl in aShift)) then
  begin
    if fVariableLineLength then
      pLastCol := CheckIfLastCol(Col, Row);
    pRow       := Row + 1;
    pct        := Col;
    if pRow < RowCount then
    begin
      if fVariableLineLength and pLastCol then
        pct := GetLastCol(pct, pRow)
      else
        while OutOfBounds(pct, pRow) do
          Dec(pct);
      MoveColRow(pct, pRow, True, True);
    end;
    Result := True;
  end;
  if (aChar = Char(VK_UP)) and (not(ssCtrl in aShift)) then
  begin
    if fVariableLineLength then
      pLastCol := CheckIfLastCol(Col, Row);
    pRow       := Row - 1;
    pct        := Col;
    if pRow > -1 then
    begin
      if fVariableLineLength and pLastCol then
        pct := GetLastCol(pct, pRow)
      else
        while OutOfBounds(pct, pRow) do
          Dec(pct);
      MoveColRow(pct, pRow, True, True);
    end;
    Result := True;
  end;
  if (ssCtrl in aShift) and (aChar = 'T') then
  begin
    if DataSize > 0 then
      Col  := GetOtherFieldCol(Col, fPosInChars);
    Result := True;
  end;
  if ((aShift = []) or (aShift = [ssShift])) and (aChar = Char(VK_TAB)) then
  begin
    if DataSize > 0 then
      Col  := GetOtherFieldCol(Col, fPosInChars);
    Result := True;
  end
  else if (aShift = [ssCtrl, ssShift]) and ((aChar >= '0') and (aChar <= '9')) then
  begin
    SetMarker(Ord(aChar) - Ord('0'), Row);
    Result := True;
  end
  else if (aShift = [ssCtrl]) and ((aChar >= '0') and (aChar <= '9')) then
  begin
    ResetSelection(True);
    pRow := GetMarkerRow(Ord(aChar) - Ord('0'));
    if pRow < RowCount then
      MoveColRow(2, pRow, True, True)
    else
      SetMarker(Ord(aChar) - Ord('0'), 1);
    Result := True;
  end
  else if (aShift = [ssShift]) and (aChar = #16) then
  begin
    if not fIsSelecting then
      ResetSelection(True);
    fIsSelecting := True;
    Result       := True;
  end;
end;

procedure THexEditor.KeyUp(var Key: Word; Shift: TShiftState);
begin
  fLastKeyWasALT := Key = VK_MENU;
{$IFDEF _debug}
  TForm(Owner).Caption := IntToStr(Key);
{$ENDIF}
  inherited KeyUp(Key, Shift);
end;

procedure THexEditor.KeyDown(var Key: Word; Shift: TShiftState);
var
  PChar: Char;
begin
  if Key = VK_INSERT then
  begin
    Key       := 0;
    fInsertOn := not fInsertOn;
    if fAutoCaretMode then
      SetAutoCaretMode(fAutoCaretMode);
    fChangedBytes.Size := 0;
    Invalidate;
    StateNotification;
    Exit;
  end;
  PChar := Char(Key);
  if Key = 8 then
  begin
    if IsInsertMode then
    begin
      Key := 0;
      if SelCount > 0 then
        DeleteSelection;
      InternalErase(True)
    end
    else
      Key := VK_LEFT;
  end;
  if (Key = VK_DELETE) then
  begin
    Key := 0;
    if (Shift = [ssCtrl]) or ((SelCount > 0) and IsInsertMode) then
      DeleteSelection
    else if IsInsertMode then
      InternalErase(False);
  end;
  if ParseKeyDown(Shift, PChar) then
    Key := 0
  else
    inherited KeyDown(Key, Shift);
{$IFDEF _debug}
  ShowSelState;
{$ENDIF}
end;

function THexEditor.HasChanged(aPos: Integer): Boolean;
begin
  Result := False;
  if IsInsertMode then
    Exit;
  if fChangedBytes.Size > aPos then
    Result := fChangedBytes.Bits[aPos];
end;

function THexEditor.IsMarkerPos(aPos: Integer): Integer;
var
  pct: Integer;
begin
  Result  := -1;
  for pct := 0 to 9 do
    if aPos = fMarker[pct] then
    begin
      Result := pct;
      Exit;
    end;
end;

function THexEditor.GetMarker(aIndex: Byte): Integer;
begin
  if aIndex > 9 then
    Raise Exception.Create('SetMarker : Invalid marker index');
  Result := fMarker[aIndex];
end;

procedure THexEditor.SetMarker(aIndex: Byte; const aValue: Integer);
begin
  if aIndex > 9 then
    Raise Exception.Create('SetMarker : Invalid marker index');
  if fMarker[aIndex] <> aValue then
  begin
    fMarker[aIndex] := aValue;
    Invalidate;
  end;
end;

function THexEditor.GetMarkerRow(aWhich: Byte): Integer;
begin
  Result := Row;
  if fMarker[aWhich] <> -1 then
    Result := fMarker[aWhich];
end;

function THexEditor.IsSelected(aPos: Integer): Boolean;
begin
  Result := False;
  if (fSelPO <> -1) and (aPos >= fSelST) and (aPos <= fSelEN) then
    Result := True;
end;

procedure THexEditor.Select(aCurCol, aCurRow, aNewCol, aNewRow: Integer);
var
  pOST, pOEN, pNAP: Integer;
begin
  pOEN := fSelEN;
  pOST := fSelST;
  pNAP := GetPosAtCursor(aNewCol, aNewRow);
  if fSelPO = -1 then
  begin
    fSelPO := GetPosAtCursor(aCurCol, aCurRow);
    if IsInsertMode then
    begin
      if fSelPO > (pNAP) then
        fSelPO := fSelPO - 1
      else if fSelPO < (pNAP) then
        pNAP := pNAP - 1;
    end;
    pOST   := pNAP;
    pOEN   := pNAP;
    fSelST := Min(pOST, fSelPO);
    fSelEN := Max(fSelPO, pOEN);
    RedrawPos(fSelST, fSelEN);
  end
  else
  begin
    if pNAP < fSelPO then
    begin
      fSelST := pNAP;
      fSelEN := fSelPO;
      RedrawPos(Min(fSelST, pOST), Max(fSelST, pOST));
      RedrawPos(Min(fSelEN, pOEN), Max(fSelEN, pOEN));
    end
    else
    begin
      if IsInsertMode then
        pNAP := pNAP - 1;
      fSelEN := pNAP;
      fSelST := fSelPO;
      RedrawPos(Min(fSelST, pOST), Max(fSelST, pOST));
      RedrawPos(Min(fSelEN, pOEN), Max(fSelEN, pOEN));
    end;
  end;
  StateNotification;
end;

procedure THexEditor.RedrawPos(aFrom, aTo: Integer);
var
  pR: TRect;
begin
  aFrom := GetRow(aFrom);
  aTo   := GetRow(aTo);
  pR    := BoxRect(2, aFrom, ColCount - 1, aTo);
  InvalidateRect(Handle, @pR, False);
end;

procedure THexEditor.ResetSelection(aDraw: Boolean);
var
  pOldFrom, pOldTo: Integer;
begin
  fIsSelecting := False;
  pOldFrom     := fSelST;
  pOldTo       := fSelEN;
  fSelST       := -1;
  fSelPO       := -1;
  fSelEN       := -1;
  if aDraw then
    RedrawPos(pOldFrom, pOldTo);
  StateNotification;
end;

procedure THexEditor.MouseDown(Button: TMouseButton; Shift: TShiftState; x, y: Integer);
begin
  inherited;
  if Button = mbLeft then
  begin
    ResetSelection(True);
    if not(ssDouble in Shift) then
      fIsSelecting := True;
  end;
{$IFDEF _debug}
  ShowSelState;
{$ENDIF}
end;

procedure THexEditor.InternalGetCurSel(var aSP, aEP, ACol, ARow: Integer);
var
  pTP: TLongPoint;
begin
  if fSelPO = -1 then
  begin
    aSP  := GetPosAtCursor(Col, Row);
    aEP  := aSP + 1;
    ACol := Col;
    ARow := Row;
  end
  else
  begin
    aSP := fSelST;
    aEP := fSelEN + 1;
    GetPosAtCursor(Col, Row);
    pTP  := GetCursorAtPos(fSelST, fPosInChars);
    ACol := pTP.x;
    ARow := pTP.y;
  end;
  if fChangedBytes.Size > aSP then
    fChangedBytes.Size := aSP;
end;

function THexEditor.CreateShift4BitStream(const aStart: Integer; var vName: TFileName): TFileStream;
var
  pbt1, pBt2: Byte;
  par       : array [0 .. 511] of Byte;
  pct       : Integer;
begin
  Result := nil;
  if aStart >= DataSize then
    Exit;
  vName             := GetTempName;
  Result            := TFileStream.Create(vName, fmCreate);
  Result.Position   := 0;
  fIntFile.Position := aStart;
  pbt1              := 0;
  while fIntFile.Position < DataSize do
  begin
    FillChar(par[0], 512, 0);
    fIntFile.Read(par[0], 512);
    for pct := 0 to 511 do
    begin
      pBt2     := par[pct] and 15;
      par[pct] := (par[pct] shr 4) or (pbt1 shl 4);
      pbt1     := pBt2;
    end;
    Result.Write(par[0], 512);
  end;
  Result.Position := 0;
end;

function THexEditor.InternalInsertNibble(const aPos: Integer; const HighNibble: Boolean): Boolean;
var
  pOldCur : TCursor;
  fST     : TFileStream;
  pName   : TFileName;
  pOldSize: Integer;
  pBT     : Byte;
begin
  Result := False;
  TestStream;
  if DataSize = 0 then
    Exit;
  pOldCur  := Cursor;
  pOldSize := fIntFile.Size;
  Cursor   := crHourGlass;
  try
    fIntFile.Position := aPos;
    fIntFile.Read(pBT, 1);
    fST := CreateShift4BitStream(aPos, pName);
    with fST do
      try
        fIntFile.Position := aPos;
        fIntFile.CopyFrom(fST, fST.Size);
      finally
        Free;
        DeleteFile(pName);
      end;
    fIntFile.Position := aPos;
    if HighNibble then
      pBT := pBT shr 4
    else
      pBT := pBT and 240;
    fIntFile.Write(pBT, 1);
    Result        := True;
    fIntFile.Size := pOldSize + 1;
  finally
    Cursor := pOldCur;
  end;
end;

function THexEditor.InsertNibble(const aPos: Integer; const HighNibble: Boolean): Boolean;
const
  Byt: Byte = 0;
begin
  Result := False;
  if DataSize < 1 then
  begin
    ResetSelection(False);
    AppendBuffer(@Byt, 1);
    Result := True;
    Exit;
  end;
  if (aPos >= DataSize) or (aPos < 0) then
    Exit;
  if not CreateUndo(U_Nibble_Insert, aPos, 0, 0) then
    Exit;
  ResetSelection(False);
  Result := InternalInsertNibble(aPos, HighNibble);
  if Result and (fChangedBytes.Size >= (aPos)) then
    fChangedBytes.Size := aPos;
  fIntBufferPos        := -1;
  CalcSizes;
end;

function THexEditor.InternalDeleteNibble(const aPos: Integer; const HighNibble: Boolean): Boolean;
var
  pOldCur   : TCursor;
  fST       : TFileStream;
  pName     : TFileName;
  pOldSize  : Integer;
  pbt1, pBt2: Byte;
begin
  Result := False;
  TestStream;
  if DataSize = 0 then
    Exit;
  pOldCur  := Cursor;
  pOldSize := fIntFile.Size;
  Cursor   := crHourGlass;
  try
    fIntFile.Position := aPos;
    fIntFile.Read(pbt1, 1);
    fST := CreateShift4BitStream(aPos, pName);
    with fST do
      try
        fIntFile.Position := aPos;
        Position          := 1;
        fIntFile.CopyFrom(fST, fST.Size - 1);
      finally
        Free;
        DeleteFile(pName);
      end;
    fIntFile.Position := aPos;
    if not HighNibble then
    begin
      fIntFile.Read(pBt2, 1);
      fIntFile.Seek(-1, soFromCurrent);
      pbt1 := (pbt1 and 240) or (pBt2 and 15);
      fIntFile.Write(pbt1, 1);
    end;
    Result        := True;
    fIntFile.Size := pOldSize;
  finally
    Cursor := pOldCur;
  end;
end;

function THexEditor.DeleteNibble(const aPos: Integer; const HighNibble: Boolean): Boolean;
begin
  Result := False;
  if (aPos >= DataSize) or (aPos < 0) then
    Exit;
  if not CreateUndo(U_Nibble_Delete, aPos, 0, 0) then
    Exit;
  ResetSelection(False);
  Result := InternalDeleteNibble(aPos, HighNibble);
  if Result and (fChangedBytes.Size >= (aPos)) then
    fChangedBytes.Size := aPos;
  fIntBufferPos        := -1;
  CalcSizes;
end;

procedure THexEditor.InternalConvertRange(const aFrom, aTo: Integer; const aTransFrom, aTransTo: TTranslationType);
var
  pSize: Integer;
  pBUF : PChar;
  pOCR : TCursor;
begin
  pSize  := (aTo - aFrom) + 1;
  pOCR   := Cursor;
  Cursor := crHourGlass;
  GetMem(pBUF, pSize);
  try
    fIntFile.Position := aFrom;
    fIntFile.Read(pBUF^, pSize);
    TranslateBufferToAnsi(aTransFrom, pBUF, pBUF, pSize);
    TranslateBufferFromAnsi(aTransTo, pBUF, pBUF, pSize);
    fIntFile.Position := aFrom;
    fIntFile.Write(pBUF^, pSize);
  finally
    FreeMem(pBUF, pSize);
    Cursor := pOCR;
  end;
end;

procedure THexEditor.ConvertRange(const aFrom, aTo: Integer; const aTransFrom, aTransTo: TTranslationType);
begin
  if aFrom > aTo then
    Exit;
  if aTransFrom = aTransTo then
    Exit;
  if (aTo >= DataSize) or (aFrom < 0) then
    Exit;
  if not CreateUndo(U_Convert, aFrom, (aTo - aFrom) + 1, 0) then
    Exit;
  InternalConvertRange(aFrom, aTo, aTransFrom, aTransTo);
  fIntBufferPos := -1;
  Invalidate;
end;

procedure THexEditor.InternalDeleteSelection(aSP, aEP, aNCol, aNRow: Integer);
var
  pTP: TLongPoint;
begin
  if aEP <= (DataSize - 1) then
    MoveFileMem(aEP, aSP, DataSize - aEP);
  fIntFile.Size := DataSize - (aEP - aSP);
  aEP           := GetPosAtCursor(aNCol, aNRow);
  if aEP >= DataSize then
  begin
    pTP := GetCursorAtPos(DataSize - 1, fPosInChars);
    MoveColRow(pTP.x, pTP.y, True, True);
  end
  else
    MoveColRow(aNCol, aNRow, True, True);
  CalcSizes;
  ResetSelection(False);
  Invalidate;
end;

procedure THexEditor.DeleteSelection;
var
  pSP, pEP    : Integer;
  pNCol, pNROW: Integer;
begin
  InternalGetCurSel(pSP, pEP, pNCol, pNROW);
  if not CreateUndo(U_Byte_removed, pSP, pEP - pSP, 0) then
    Exit;
  InternalDeleteSelection(pSP, pEP, pNCol, pNROW);
end;

function THexEditor.CreateUndo(aType: Integer; aPos, aCount, aReplCount: Integer): Boolean;
begin
  Result := False;
  if DataSize > 0 then
    Result := True;
  if not Result then
    if (aType = U_Insert_buffer) or (aType = U_Append_Buffer) then
      Result := True;
  if fNoSizeChange and Result then
    if (aType = U_Byte_removed) or (aType = U_Insert_buffer) or (aType = U_Append_Buffer) or (aType = U_Nibble_Insert) or (aType = U_Nibble_Delete) or ((aType = U_Replace_Selection) and (aCount <> aReplCount)) then
      Result := False;
  if Result then
    fCanUndo := Result;
  if Result then
  begin
    DoCreateUndo(aType, aPos, aCount, aReplCount);
    fModified := True;
  end;
  StateNotification;
end;

procedure THexEditor.ResetUndo;
begin
  fCanUndo   := False;
  fUndoDesc  := '无需撤销';
  fUndoCount := 0;
  if fUndoMem <> nil then
  begin
    fUndoMem.Size := 0;
    fUndoMem.Free;
    fUndoMem := nil;
  end;
  if Assigned(fStateChanged) then
    fStateChanged(Self);
end;

function THexEditor.GetSelStart: Integer;
begin
  if fSelPO = -1 then
  begin
    Result := GetPosAtCursor(Col, Row);
  end
  else
    Result := fSelPO;
end;

function THexEditor.GetSelEnd: Integer;
begin
  if fSelPO = -1 then
    Result := GetPosAtCursor(Col, Row)
  else
  begin
    Result := fSelEN;
    if fSelPO = fSelEN then
      Result := fSelST;
  end;
end;

procedure THexEditor.SetSelStart(aValue: Integer);
var
  pTP: TLongPoint;
begin
  if (aValue < 0) or (aValue >= DataSize) then
    raise Exception.Create('Invalid SelStart')
  else
  begin
    ResetSelection(True);
    GetPosAtCursor(Col, Row);
    pTP := GetCursorAtPos(aValue, fPosInChars);
    MoveColRow(pTP.x, pTP.y, True, True);
  end;
end;

procedure THexEditor.SetSelEnd(aValue: Integer);
var
  pTP: TLongPoint;
begin
  if (aValue < 0) or (aValue >= DataSize) then
    raise Exception.Create('Invalid SelEnd')
  else
  begin
    ResetSelection(True);
    GetPosAtCursor(Col, Row);
    pTP := GetCursorAtPos(aValue, fPosInChars);
    Select(Col, Row, pTP.x, pTP.y);
  end;
end;

procedure THexEditor.SetInCharField(aValue: Boolean);
begin
  if DataSize < 1 then
    Exit;
  GetPosAtCursor(Col, Row);
  if fPosInChars <> aValue then
    MoveColRow(GetOtherFieldCol(Col, fPosInChars), Row, True, True);
end;

function THexEditor.GetInCharField: Boolean;
begin
  Result := False;
  if DataSize < 1 then
    Exit;
  GetPosAtCursor(Col, Row);
  Result := fPosInChars;
end;

procedure THexEditor.Loaded;
begin
  inherited;
  CreateEmptyFile('Untitled');
end;

procedure THexEditor.CreateWnd;
begin
  inherited;
  if (csDesigning in ComponentState) or (fFileName = '---') then
    CreateEmptyFile('未命名');
end;

procedure THexEditor.WMSetFocus(var Msg: TWMSetFocus);
begin
  inherited;
  CreateColoredCaret;
  SetCaretPos(-50, -50);
  ShowCaret(Handle);
  Invalidate;
end;

procedure THexEditor.WMKillFocus(var Msg: TWMKillFocus);
begin
  inherited;
  HideCaret(Handle);
  DestroyCaret();
  fIsSelecting := False;
  Invalidate;
end;

procedure THexEditor.WMSTATECHANGED(var Msg: TMessage);
begin
  if Msg.WParam = 7 then
    if Assigned(fStateChanged) then
      fStateChanged(Self);
end;

procedure THexEditor.SetTranslation(aValue: TTranslationType);
begin
  if fTranslation <> aValue then
  begin
    fTranslation := aValue;
    Invalidate;
  end;
end;

procedure THexEditor.SetModified(aValue: Boolean);
begin
  fModified := aValue;
  if not aValue then
  begin
    fCanUndo           := False;
    fChangedBytes.Size := 0;
    Invalidate;
  end;
end;

procedure THexEditor.SetBytesPerLine(aValue: Integer);
var
  pPS, pSP, pSS, pSE: Integer;
  pTP               : TLongPoint;
begin
  if (aValue < 1) or (aValue > 256) then
    raise Exception.Create('Invalid BytesPerLine argument')
  else if fBytesPerLine <> aValue then
  begin
    LockWindowUpdate(Handle);
    ClearOffsets;
    fVariableLineLength := False;
    pSP                 := fSelPO;
    pSS                 := fSelST;
    pSE                 := fSelEN;
    pPS                 := GetPosAtCursor(Col, Row);
    fBytesPerLine       := aValue;
    fBPL2               := aValue * 2;
    CalcSizes;
    pTP := GetCursorAtPos(pPS, fPosInChars);
    MoveColRow(pTP.x, pTP.y, True, True);
    fSelPO := pSP;
    fSelST := pSS;
    fSelEN := pSE;
    LockWindowUpdate(0);
  end;
end;

procedure THexEditor.InternalAppendBuffer(aBuffer: PChar; aSize: Integer);
var
  pct: Integer;
begin
  TestStream;
  if DataSize = 0 then
  begin
    fIntFile.Position  := 0;
    fChangedBytes.Size := 0;
  end;
  pct           := DataSize;
  fIntFile.Size := pct + aSize;
  SetMemAtPos(PByteArray(aBuffer), pct, aSize);
  CalcSizes;
end;

procedure THexEditor.InternalInsertBuffer(aBuffer: PChar; aSize, aPos: Integer);
var
  pct: Integer;
begin
  TestStream;
  if DataSize = 0 then
  begin
    fIntFile.Position  := 0;
    fChangedBytes.Size := 0;
  end;
  pct           := DataSize;
  fIntFile.Size := pct + aSize;
  if aPos < pct then
    MoveFileMem(aPos, aPos + aSize, DataSize - aPos - aSize);
  SetMemAtPos(PByteArray(aBuffer), aPos, aSize);
  CalcSizes;
end;

procedure THexEditor.InsertBuffer(aBuffer: PChar; aSize, aPos: Integer);
var
  pct: Integer;
begin
  if not CreateUndo(U_Insert_buffer, aPos, aSize, 0) then
    Exit;
  InternalInsertBuffer(aBuffer, aSize, aPos);
  if fChangedBytes.Size >= (aPos) then
    fChangedBytes.Size := aPos;
  pct                  := GetPosAtCursor(Col, Row);
  if pct = aPos then
  begin
    fSelPO := aPos;
    fSelST := aPos;
    fSelEN := aPos + aSize - 1;
    StateNotification;
  end;
  Invalidate;
end;

procedure THexEditor.AppendBuffer(aBuffer: PChar; aSize: Integer);
var
  pct: Integer;
  pTP: TLongPoint;
begin
  if not CreateUndo(U_Append_Buffer, DataSize, aSize, 0) then
    Exit;
  if fChangedBytes.Size >= (DataSize) then
    fChangedBytes.Size := DataSize;
  pct                  := DataSize;
  InternalAppendBuffer(aBuffer, aSize);
  GetPosAtCursor(Col, Row);
  pTP := GetCursorAtPos(pct, fPosInChars);
  MoveColRow(pTP.x, pTP.y, True, True);
  fSelPO := pct;
  fSelST := pct;
  fSelEN := pct + aSize - 1;
  StateNotification;
  Invalidate;
end;

procedure THexEditor.ReplaceSelection(aBuffer: PChar; aSize: Integer);
var
  pSP, pEP, pCol, pRow: Integer;
  pct                 : Integer;
  pOldCol, pOldRow    : Integer;
begin
  if fSelPO = -1 then
    InsertBuffer(aBuffer, aSize, GetSelStart)
  else
  begin
    if fNoSizeChange then
    begin
      if aSize > SelCount then
        aSize := SelCount
      else if SelCount > aSize then
      begin
        SelStart := Min(SelStart, SelEnd);
        SelEnd   := SelStart + aSize - 1;
      end;
    end;
    if not CreateUndo(U_Replace_Selection, fSelST, aSize, SelCount) then
      Exit;
    pOldCol := Col;
    pOldRow := Row;
    InternalGetCurSel(pSP, pEP, pCol, pRow);
    InternalDeleteSelection(pSP, pEP, pCol, pRow);
    InternalInsertBuffer(aBuffer, aSize, pSP);
    if fChangedBytes.Size >= pSP then
      fChangedBytes.Size := Max(0, pSP);
    pct                  := GetPosAtCursor(pOldCol, pOldRow);
    if (pct = pSP) and (DataSize > pct) then
    begin
      MoveColRow(pOldCol, pOldRow, True, True);
      fSelPO := pSP;
      fSelST := pSP;
      fSelEN := pSP + aSize - 1;
      StateNotification;
    end;
  end;
end;

procedure THexEditor.DoCreateUndo(aType: Integer; aPos, aCount, aReplCount: Integer);
  procedure FillBuffer(var aBuffer: TUndoRec; aSize: Integer);
  var
    pTP: TLongPoint;
  begin
    FillChar(aBuffer, aSize, 0);
    aBuffer.Typ    := aType;
    aBuffer.CurPos := GetPosAtCursor(Col, Row);
    if not fPosInChars then
    begin
      pTP          := GetCursorAtPos(aBuffer.CurPos, fPosInChars);
      aBuffer.C1st := Col - pTP.x;
    end;
    aBuffer.CharField := fPosInChars;
    aBuffer.SelS      := fSelST;
    aBuffer.SelE      := fSelEN;
    aBuffer.SelP      := fSelPO;
    aBuffer.Pos       := aPos;
    aBuffer.Count     := aCount;
    aBuffer.ReplCount := aReplCount;
    aBuffer.Modified  := fModified;
  end;
  procedure DeleteFirstUndo;
  var
    pSK: Integer;
    pct: Integer;
    pPT: Pointer;
  begin
    fUndoMem.Position := fUndoMem.Size;
    pct               := fUndoMem.Position;
    while fUndoMem.Position <> 0 do
    begin
      fUndoMem.Seek(-4, soFromCurrent);
      fUndoMem.Read(pSK, 4);
      pct := fUndoMem.Position;
      fUndoMem.Seek(-pSK, soFromCurrent);
    end;
{$IFDEF Win32}
    Integer(pPT) := Integer(fUndoMem.Memory) + pct;
{$ELSE}
    pPT           := nil;
    Integer(pPT^) := Integer(fUndoMem.Memory) + pct;
{$ENDIF}
    Move(pPT^, fUndoMem.Memory^, fUndoMem.Size - pct);
    fUndoMem.Size     := fUndoMem.Size - pct;
    fUndoMem.Position := fUndoMem.Size;
    Dec(fUndoCount);
  end;

var
  pBUF: PUndoRec;
  par : PByteArray;
begin
  fUndoDesc := UndoSTR[aType];
  if fUndoMem = nil then
    fUndoMem        := TMemoryStream.Create;
  fUndoMem.Position := fUndoMem.Size;
  Inc(fUndoCount);
  if fUndoCount > cMax_Undo then
    DeleteFirstUndo;
  case aType of
    U_Byte_changed:
      begin
        GetMem(pBUF, SizeOf(TUndoRec));
        try
          FillBuffer(pBUF^, SizeOf(TUndoRec));
          pBUF.Buffer  := GetByteAtPos(aPos);
          pBUF.Changed := HasChanged(aPos);
          fUndoMem.Write(pBUF^, SizeOf(TUndoRec));
          aPos := SizeOf(TUndoRec) + 4;
          fUndoMem.Write(aPos, 4);
        finally
          FreeMem(pBUF, SizeOf(TUndoRec));
        end;
      end;
    U_Byte_removed:
      begin
        GetMem(pBUF, SizeOf(TUndoRec) + aCount - 1);
        try
          FillBuffer(pBUF^, SizeOf(TUndoRec));
          par := @pBUF.Buffer;
          GetMemAtPos(par, aPos, aCount);
          fUndoMem.Write(pBUF^, SizeOf(TUndoRec) + aCount - 1);
          aPos := SizeOf(TUndoRec) + 4 + aCount - 1;
          fUndoMem.Write(aPos, 4);
        finally
          FreeMem(pBUF, SizeOf(TUndoRec) + aCount - 1);
          fIntBufferPos := -1;
        end;
      end;
    U_Insert_buffer:
      begin
        GetMem(pBUF, SizeOf(TUndoRec));
        try
          FillBuffer(pBUF^, SizeOf(TUndoRec));
          fUndoMem.Write(pBUF^, SizeOf(TUndoRec));
          aPos := SizeOf(TUndoRec) + 4;
          fUndoMem.Write(aPos, 4);
        finally
          FreeMem(pBUF, SizeOf(TUndoRec) + aCount - 1);
          fIntBufferPos := -1;
        end;
      end;
    U_Replace_Selection:
      begin
        GetMem(pBUF, SizeOf(TUndoRec) + aReplCount - 1);
        try
          FillBuffer(pBUF^, SizeOf(TUndoRec));
          par := @pBUF.Buffer;
          GetMemAtPos(par, aPos, aReplCount);
          fUndoMem.Write(pBUF^, SizeOf(TUndoRec) + aReplCount - 1);
          aPos := SizeOf(TUndoRec) + 4 + aReplCount - 1;
          fUndoMem.Write(aPos, 4);
        finally
          FreeMem(pBUF, SizeOf(TUndoRec) + aReplCount - 1);
          fIntBufferPos := -1;
        end;
      end;
    U_Append_Buffer:
      begin
        GetMem(pBUF, SizeOf(TUndoRec));
        try
          FillBuffer(pBUF^, SizeOf(TUndoRec));
          fUndoMem.Write(pBUF^, SizeOf(TUndoRec));
          aPos := SizeOf(TUndoRec) + 4;
          fUndoMem.Write(aPos, 4);
        finally
          FreeMem(pBUF, SizeOf(TUndoRec));
          fIntBufferPos := -1;
        end;
      end;
    U_Nibble_Insert:
      begin
        GetMem(pBUF, SizeOf(TUndoRec));
        try
          FillBuffer(pBUF^, SizeOf(TUndoRec));
          pBUF.Buffer  := GetByteAtPos(aPos);
          pBUF.Changed := HasChanged(aPos);
          fUndoMem.Write(pBUF^, SizeOf(TUndoRec));
          aPos := SizeOf(TUndoRec) + 4;
          fUndoMem.Write(aPos, 4);
        finally
          FreeMem(pBUF, SizeOf(TUndoRec));
        end;
      end;
    U_Nibble_Delete:
      begin
        GetMem(pBUF, SizeOf(TUndoRec));
        try
          FillBuffer(pBUF^, SizeOf(TUndoRec));
          pBUF.Buffer  := GetByteAtPos(aPos);
          pBUF.Changed := HasChanged(aPos);
          fUndoMem.Write(pBUF^, SizeOf(TUndoRec));
          aPos := SizeOf(TUndoRec) + 4;
          fUndoMem.Write(aPos, 4);
        finally
          FreeMem(pBUF, SizeOf(TUndoRec));
        end;
      end;
    U_Convert:
      begin
        GetMem(pBUF, SizeOf(TUndoRec) + aCount - 1);
        try
          FillBuffer(pBUF^, SizeOf(TUndoRec));
          par := @pBUF.Buffer;
          GetMemAtPos(par, aPos, aCount);
          fUndoMem.Write(pBUF^, SizeOf(TUndoRec) + aCount - 1);
          aPos := SizeOf(TUndoRec) + 4 + aCount - 1;
          fUndoMem.Write(aPos, 4);
        finally
          FreeMem(pBUF, SizeOf(TUndoRec) + aCount - 1);
          fIntBufferPos := -1;
        end;
      end;
  end;
end;

function THexEditor.Undo: Boolean;
  procedure SetBuffer(aBuffer: TUndoRec);
  var
    pTP: TLongPoint;
  begin
    pTP := GetCursorAtPos(aBuffer.CurPos, aBuffer.CharField);
    if not aBuffer.CharField then
      if DataSize > 0 then
        pTP.x := pTP.x + aBuffer.C1st;
    MoveColRow(pTP.x, pTP.y, True, True);
    fSelST    := aBuffer.SelS;
    fSelEN    := aBuffer.SelE;
    fSelPO    := aBuffer.SelP;
    fModified := aBuffer.Modified;
  end;
  function SetUndoPointer(var aUR: TUndoRec): Byte;
  var
    pSK: Integer;
  begin
    fUndoMem.Position := fUndoMem.Size - 4;
    fUndoMem.Read(pSK, 4);
    fUndoMem.Seek(-pSK, soFromCurrent);
    fUndoMem.Read(aUR, SizeOf(TUndoRec));
    Result := aUR.Typ;
  end;
  procedure NextUndo(aCount: Integer);
  var
    aUR: TUndoRec;
  begin
    fDataSize := -1;
    fUndoMem.SetSize(Max(0, fUndoMem.Size - aCount));
    Dec(fUndoCount);
    if fUndoMem.Size < 5 then
    begin
      ResetUndo;
    end
    else
    begin
      fUndoDesc := UndoSTR[(SetUndoPointer(aUR))];
      StateNotification;
    end;
  end;

var
  pTY: Byte;
  pUR: TUndoRec;
begin
  Result := False;
  if not fCanUndo then
  begin
    ResetUndo;
    Exit;
  end;
  if (fUndoMem <> nil) and (fUndoMem.Size > 4) then
  begin
    pTY := SetUndoPointer(pUR);
    case pTY of
      U_Byte_changed:
        begin
          SetByteAtPos(pUR.Pos, pUR.Buffer);
          SetChanged(pUR.Pos, pUR.Changed);
          SetBuffer(pUR);
          RedrawPos(pUR.Pos, pUR.Pos);
          SetInternalBufferByte(pUR.Pos, pUR.Buffer);
          NextUndo(SizeOf(TUndoRec) + 4);
        end;
      U_Byte_removed:
        begin
          fIntBufferPos := -1;
          InternalInsertBuffer(Pointer(Integer(fUndoMem.Memory) + fUndoMem.Position - 1), pUR.Count, pUR.Pos);
          SetBuffer(pUR);
          if DWORD(fChangedBytes.Size) >= (pUR.Pos) then
            fChangedBytes.Size := pUR.Pos;
          Invalidate;
          NextUndo(SizeOf(TUndoRec) + 4 + pUR.Count - 1);
        end;
      U_Insert_buffer:
        begin
          fIntBufferPos := -1;
          InternalDeleteSelection(pUR.Pos, pUR.Pos + pUR.Count, 10, 0);
          SetBuffer(pUR);
          if DWORD(fChangedBytes.Size) >= (pUR.Pos) then
            fChangedBytes.Size := pUR.Pos;
          Invalidate;
          NextUndo(SizeOf(TUndoRec) + 4);
        end;
      U_Replace_Selection:
        begin
          fIntBufferPos := -1;
          InternalDeleteSelection(pUR.Pos, pUR.Pos + pUR.Count, 10, 0);
          InternalInsertBuffer(Pointer(Integer(fUndoMem.Memory) + fUndoMem.Position - 1), pUR.ReplCount, pUR.Pos);
          SetBuffer(pUR);
          if DWORD(fChangedBytes.Size) >= (pUR.Pos) then
            fChangedBytes.Size := Max(0, pUR.Pos - 1);
          Invalidate;
          NextUndo(SizeOf(TUndoRec) + 4 + pUR.ReplCount - 1);
        end;
      U_Append_Buffer:
        begin
          fIntBufferPos := -1;
          Col           := 2;
          fIntFile.Size := pUR.Pos;
          CalcSizes;
          if DWORD(fChangedBytes.Size) >= (pUR.Pos) then
            fChangedBytes.Size := pUR.Pos;
          SetBuffer(pUR);
          Invalidate;
          NextUndo(SizeOf(TUndoRec) + 4);
        end;
      U_Nibble_Insert:
        begin
          fIntBufferPos := -1;
          InternalDeleteNibble(pUR.Pos, False);
          SetByteAtPos(pUR.Pos, pUR.Buffer);
          SetChanged(pUR.Pos, pUR.Changed);
          SetBuffer(pUR);
          if DWORD(fChangedBytes.Size) >= (pUR.Pos) then
            fChangedBytes.Size := pUR.Pos;
          fIntFile.Size        := fIntFile.Size - 1;
          CalcSizes;
          Invalidate;
          NextUndo(SizeOf(TUndoRec) + 4);
        end;
      U_Nibble_Delete:
        begin
          fIntBufferPos := -1;
          InternalInsertNibble(pUR.Pos, False);
          SetByteAtPos(pUR.Pos, pUR.Buffer);
          SetChanged(pUR.Pos, pUR.Changed);
          SetBuffer(pUR);
          if DWORD(fChangedBytes.Size) >= (pUR.Pos) then
            fChangedBytes.Size := pUR.Pos;
          fIntFile.Size        := fIntFile.Size - 1;
          CalcSizes;
          Invalidate;
          NextUndo(SizeOf(TUndoRec) + 4);
        end;
      U_Convert:
        begin
          fIntBufferPos := -1;
          SetMemAtPos(Pointer(Integer(fUndoMem.Memory) + fUndoMem.Position - 1), pUR.Pos, pUR.Count);
          SetBuffer(pUR);
          if DWORD(fChangedBytes.Size) >= (pUR.Pos) then
            fChangedBytes.Size := pUR.Pos;
          Invalidate;
          NextUndo(SizeOf(TUndoRec) + 4 + pUR.Count - 1);
        end;
    end;
  end
  else
    ResetUndo;
end;

procedure THexEditor.SetChanged(aPos: Integer; aValue: Boolean);
begin
  if IsInsertMode then
    fChangedBytes.Size := 0;
  if not aValue then
    if fChangedBytes.Size <= aPos then
      Exit;
  fChangedBytes[aPos] := aValue;
end;
{$IFDEF _debug}

procedure THexEditor.SaveUndo(aFileName: TFileName);
begin
  if fUndoMem <> nil then
    fUndoMem.SaveToFile(aFileName);
end;
{$ENDIF}

procedure THexEditor.MoveFileMem(aFrom, aTo, aCount: Integer);
var
  pBU: PChar;
begin
  GetMem(pBU, aCount);
  try
    fIntFile.Position := aFrom;
    fIntFile.Read(pBU^, aCount);
    fIntFile.Position := aTo;
    fIntFile.Write(pBU^, aCount);
  finally
    FreeMem(pBU, aCount);
  end;
end;

procedure THexEditor.CheckInternalBuffer(aPos: Integer);
var
  pFR: Integer;
begin
  if (fIntBufferPos = -1) or (aPos < fIntBufferPos) or ((aPos + fBytesPerLine) > (fIntBufferPos + cBuf_Size)) then
  begin
    fDataSize := -1;
    pFR       := aPos - (cBuf_Size div 2);
    if pFR < 0 then
      pFR             := 0;
    fIntFile.Position := pFR;
    fIntFile.Read(fIntBuffer^, cBuf_Size);
    fIntBufferPos := pFR;
  end;
end;

procedure THexEditor.SetInternalBufferByte(aPos: Integer; aByte: Byte);
begin
  if (aPos < fIntBufferPos) or ((aPos + fBytesPerLine) > (fIntBufferPos + cBuf_Size)) or (fIntBufferPos = -1) then
    Exit;
  fIntBuffer[aPos - fIntBufferPos] := aByte;
end;

function THexEditor.GetCursorPos: Integer;
begin
  Result := GetPosAtCursor(Col, Row);
  if Result < 0 then
    Result := 0;
  if Result > Max(0, DataSize - 1) then
    Result := Max(0, DataSize - 1)
end;

function THexEditor.GetSelCount: Integer;
begin
  if fSelPO = -1 then
    Result := 0
  else
    Result := Max(fSelST, fSelEN) - Min(fSelST, fSelEN) + 1;
end;

procedure THexEditor.TestStream;
begin
  if fIntFile = nil then
  begin
    DeleteFile(fInternalName);
    fIntFile           := TFileStream.Create(fInternalName, fmCreate);
    fIntFile.Position  := 0;
    fChangedBytes.Size := 0;
  end;
end;

function THexEditor.GetMemory(aIndex: Integer): Char;
begin
  if (aIndex < 0) or (aIndex >= DataSize) then
    Raise Exception.Create('Invalid GetMemory index')
  else
  begin
    fIntFile.Position := aIndex;
    fIntFile.Read(Result, 1);
  end;
end;

procedure THexEditor.SetMemory(aIndex: Integer; aChar: Char);
begin
  if (aIndex < 0) or (aIndex >= DataSize) then
    Raise Exception.Create('Invalid SetMemory index')
  else
  begin
    fIntFile.Position := aIndex;
    fIntFile.Write(aChar, 1);
    fIntBufferPos := -1;
  end;
end;

procedure THexEditor.SetReadOnly(aValue: Boolean);
begin
  if aValue and (not fReadOnly) then
  begin
    fReadOnly := True;
    StateNotification;
  end;
end;

function THexEditor.BufferFromFile(aPos: Integer; var aCount: Integer): PChar;
begin
  if (aPos < 0) or (aPos >= DataSize) then
    raise Exception.Create('Invalid BufferFromFile argument')
  else
  begin
    if (aPos + aCount) > DataSize then
      aCount := (DataSize - aPos) + 1;
    GetMem(Result, aCount);
    try
      fIntFile.Position := aPos;
      fIntFile.Read(Result^, aCount);
    except
      FreeMem(Result, aCount);
      Result := nil;
      aCount := 0;
    end;
  end;
end;

procedure THexEditor.WMVScroll(var Msg: TWMVScroll);
var
  pRC: TRect;
begin
  inherited;
  pRC := CellRect(Col, Row);
  if pRC.Left + pRC.Bottom = 0 then
    SetCaretPos(-50, -50)
  else
    SetCaretPos(pRC.Left, pRC.Top);
end;

procedure THexEditor.WMHScroll(var Msg: TWMHScroll);
var
  pRC: TRect;
begin
  inherited;
  pRC := CellRect(Col, Row);
  if pRC.Left + pRC.Bottom = 0 then
    SetCaretPos(-50, -50)
  else
    SetCaretPos(pRC.Left, pRC.Top);
end;

procedure THexEditor.CreateColoredCaret;
begin
  DestroyCaret();
  fCaretBitmap.Width              := fCharWidth;
  fCaretBitmap.Height             := fCharHeight - 2;
  fCaretBitmap.Canvas.Brush.Color := clBlack;
  fCaretBitmap.Canvas.FillRect(Rect(0, 0, fCharWidth, fCharHeight - 2));
  fCaretBitmap.Canvas.Brush.Color := fColors.CursorFrame xor $00FFFFFF;
  case fCaretStyle of
    csFull:
      fCaretBitmap.Canvas.FillRect(Rect(0, 0, fCharWidth, fCharHeight - 2));
    csLeftLine:
      fCaretBitmap.Canvas.FillRect(Rect(0, 0, 2, fCharHeight - 2));
    csBottomLine:
      fCaretBitmap.Canvas.FillRect(Rect(0, fCharHeight - 4, fCharWidth, fCharHeight - 2));
  end;
  CreateCaret(Handle, fCaretBitmap.Handle, 0, 0);
end;

procedure THexEditor.SetBytesPerColumn(const Value: Integer);
begin
  if fBytesPerColumn <> (Value * 2) then
  begin
    fBytesPerColumn := Value * 2;
    AdjustMetrics;
    Invalidate;
  end;
end;

function THexEditor.GetBytesPerColumn: Integer;
begin
  Result := fBytesPerColumn div 2;
end;

function THexEditor.Find(aBuffer: PChar; const aCount, aStart, aEnd: Integer; const IgnoreCase, SearchText: Boolean): Integer;
var
  pCr                       : TCursor;
  pChAct                    : Char;
  pCMem, pCFind, pCHit, pEnd: Integer;
begin
  Result := -1;
  pEnd   := aEnd;
  if pEnd >= DataSize then
    pEnd := DataSize - 1;
  if aCount < 1 then
    Exit;
  if aStart + aCount > (pEnd + 1) then
    Exit;
  pCr    := Cursor;
  Cursor := crHourGlass;
  if SearchText and (fTranslation <> ttAnsi) then
    TranslateBufferFromAnsi(fTranslation, aBuffer, aBuffer, aCount);
  try
    if IgnoreCase then
      CharLowerBuff(aBuffer, aCount);
    pCMem  := aStart;
    pCFind := 0;
    pCHit  := pCMem + 1;
    repeat
{$IFDEF _debug}
      if (pCMem mod 100000) = 0 then
        TForm(Owner).Caption := IntToStr(pCMem);
{$ENDIF}
      if pCMem > pEnd then
        Exit;
      CheckInternalBuffer(pCMem);
      pChAct := Char(fIntBuffer[pCMem - fIntBufferPos]);
      if IgnoreCase then
        CharLowerBuff(@pChAct, 1);
      if (pChAct = aBuffer[pCFind]) then
      begin
        if pCFind = (aCount - 1) then
        begin
          Result := pCMem - aCount + 1;
          Exit;
        end
        else
        begin
          if pCFind = 0 then
            pCHit := pCMem + 1;
          Inc(pCMem);
          Inc(pCFind);
        end;
      end
      else
      begin
        pCMem  := pCHit;
        pCFind := 0;
        pCHit  := pCMem + 1;
      end;
    until False;
  finally
    Cursor := pCr;
  end;
end;

procedure THexEditor.SetOffsetDisplayWidth;
begin
  if FOffsetDisplay = odNone then
    fOffSetDisplayWidth := 0
  else
  begin
    if FOffsetDisplay = odHex then
      fOffSetDisplayWidth := Length(IntToHex(LineOffset[RowCount - 1], 1)) + 3
    else if FOffsetDisplay = odDec then
      fOffSetDisplayWidth := Length(IntToStr(LineOffset[RowCount - 1])) + 1
    else
      fOffSetDisplayWidth := Length(IntToOctal(LineOffset[RowCount - 1])) + 3;
  end;
  ColWidths[0] := fOffSetDisplayWidth * fCharWidth;
end;

procedure THexEditor.SetShowMarkerColumn(const Value: Boolean);
begin
  if Value <> fShowMarkerCol then
  begin
    fShowMarkerCol := Value;
    AdjustMetrics;
  end;
end;

function THexEditor.Seek(const aOffset, aOrigin: Integer; const FailIfOutOfRange: Boolean): Boolean;
var
  pNP: Integer;
begin
  Result := False;
  pNP    := GetCursorPos;
  case aOrigin of
    soFromBeginning:
      pNP := aOffset;
    soFromCurrent:
      pNP := GetCursorPos + aOffset;
    soFromEnd:
      pNP := DataSize + aOffset - 1;
  end;
  if DataSize < 1 then
    Exit;
  if pNP < 0 then
  begin
    pNP := 0;
    if FailIfOutOfRange then
      Exit;
  end;
  if pNP >= DataSize then
  begin
    pNP := DataSize - 1;
    if FailIfOutOfRange then
      Exit;
  end;
  SelStart := pNP;
  Result   := True;
end;

procedure THexEditor.SetSwapNibbles(const Value: Boolean);
begin
  if Integer(Value) <> fSwapNibbles then
  begin
    fSwapNibbles := Integer(Value);
    Invalidate;
  end;
end;

function THexEditor.GetSwapNibbles: Boolean;
begin
  Result := Boolean(fSwapNibbles);
end;

procedure THexEditor.SetColors(const Value: TColors);
begin
  fColors := Value;
end;

procedure THexEditor.SetOffsetChar(const Value: Char);
begin
  if (FOffsetChar <> Value) then
  begin
    FOffsetChar := Value;
    Invalidate;
  end;
end;

procedure THexEditor.SetOffsetDisplay(const Value: TOffsetDisplayStyle);
begin
  if FOffsetDisplay <> Value then
  begin
    FOffsetDisplay := Value;
    SetOffsetDisplayWidth;
    Invalidate;
  end;
end;

procedure THexEditor.SetCaretStyle(const Value: TCaretStyle);
begin
  if fCaretStyle <> Value then
  begin
    fCaretStyle := Value;
    if Focused then
    begin
      CreateColoredCaret;
      SetCaretPos(-50, -50);
      ShowCaret(Handle);
      Invalidate;
    end;
  end;
end;

procedure THexEditor.SetFocusFrame(const Value: Boolean);
begin
  if fFocusFrame <> Value then
  begin
    fFocusFrame := Value;
    Invalidate;
  end;
end;

procedure THexEditor.SetMaskWhiteSpaces(const aValue: Boolean);
begin
  if fMaskWhiteSpaces <> aValue then
  begin
    fMaskWhiteSpaces := aValue;
    Invalidate;
  end;
end;

procedure THexEditor.SetMaskChar(const aValue: Char);
begin
  if fMaskChar <> aValue then
  begin
    fMaskChar := aValue;
    Invalidate;
  end;
end;

procedure THexEditor.SetAsText(const aValue: string);
var
  lPC: PChar;
begin
  if DataSize > 0 then
  begin
    SelStart := 0;
    SelEnd   := DataSize - 1;
  end;
  GetMem(lPC, Length(aValue));
  try
    Move(aValue[1], lPC^, Length(aValue));
    TranslateBufferFromAnsi(fTranslation, @aValue[1], lPC, Length(aValue));
    ReplaceSelection(lPC, Length(aValue));
  finally
    FreeMem(lPC);
  end;
end;

procedure THexEditor.SetAsHex(const aValue: string);
var
  buf: PChar;
  lBD: Integer;
begin
  if DataSize > 0 then
  begin
    SelStart := 0;
    SelEnd   := DataSize - 1;
  end;

  GetMem(buf, Length(aValue));

  try
    ConvertHexToBin(@aValue[1], buf, Length(aValue), SwapNibbles, lBD);
    ReplaceSelection(buf, lBD);
  finally
    FreeMem(buf);
  end;
end;

function THexEditor.GetAsText: string;
begin
  if DataSize < 1 then
    Result := ''
  else
  begin
    SetLength(Result, DataSize);
    GetMemAtPos(@Result[1], 0, DataSize);
  end;
end;

function THexEditor.GetAsHex: string;
var
  buf: PChar;
  lSZ: Integer;
begin
  if DataSize < 1 then
    Result := ''
  else
  begin
    lSZ := DataSize;
    GetMem(buf, DataSize);
    try
      buf := BufferFromFile(0, lSZ);
      SetLength(Result, DataSize * 2);
      ConvertBinToHex(buf, @Result[1], DataSize, SwapNibbles);
    finally
      FreeMem(buf, DataSize);
    end;
  end;
end;

procedure THexEditor.SetVariableLineLength(const aValue: Boolean);
var
  pPos         : Integer;
  pt           : TLongPoint;
  pSS, pSE, pSP: Integer;
begin
  if aValue <> fVariableLineLength then
  begin
    pSP                 := fSelPO;
    pSS                 := fSelST;
    pSE                 := fSelEN;
    pPos                := GetPosAtCursor(Col, Row);
    fVariableLineLength := aValue;
    CalcSizes;
    pt := GetCursorAtPos(pPos, fPosInChars);
    MoveColRow(pt.x, pt.y, True, True);
    Application.ProcessMessages;
    fSelST := pSS;
    fSelEN := pSE;
    fSelPO := pSP;
    Invalidate;
  end;
end;

procedure THexEditor.AdjustLineLengthsCount;
begin
  if fOffsets.Count = 0 then
    fOffsets.Add(Pointer(0));
  while fOffsets.Count < (RowCount + 1) do
    fOffsets.Add(Pointer(fBytesPerLine + Integer(fOffsets[fOffsets.Count - 1])));
end;

procedure THexEditor.SetLineLength(aLine, aLength: Integer);
var
  pct: Integer;
  pdf: Integer;
begin
  AdjustLineLengthsCount;
  if (aLength < 1) or (aLength > fBytesPerLine) then
  begin
    Raise Exception.Create('Invalid Line Length argument');
    Exit;
  end;
  while fOffsets.Count < (aLine + 2) do
    fOffsets.Add(Pointer(fBytesPerLine + Integer(fOffsets[fOffsets.Count - 1])));
  pdf := LineLength[aLine] - aLength;
  if pdf <> 0 then
  begin
    for pct         := fOffsets.Count - 1 downto aLine + 1 do
      fOffsets[pct] := Pointer(Integer(fOffsets[pct]) - pdf);
    if fVariableLineLength then
    begin
      CalcSizes;
      Invalidate;
    end;
  end;
end;

function THexEditor.GetLineLength(aLine: Integer): Integer;
begin
  if not fVariableLineLength then
    Result := fBytesPerLine
  else
  begin
    AdjustLineLengthsCount;
    while fOffsets.Count < (aLine + 2) do
      fOffsets.Add(Pointer(fBytesPerLine + Integer(fOffsets[fOffsets.Count - 1])));
    Result := Integer(fOffsets[aLine + 1]) - Integer(fOffsets[aLine]);
  end;
end;

function THexEditor.GetLineOffset(aLine: Integer): Integer;
begin
  if not fVariableLineLength then
    Result := aLine * fBytesPerLine
  else
  begin
    AdjustLineLengthsCount;
    while fOffsets.Count < (aLine + 2) do
      fOffsets.Add(Pointer(fBytesPerLine + Integer(fOffsets[fOffsets.Count - 1])));
    Result := Integer(fOffsets[aLine]);
  end;
end;

procedure THexEditor.ClearOffsets;
begin
  fOffsets.Clear;
end;

procedure THexEditor.SetLineLengths(aLengths: TList);
var
  pct               : Integer;
  pPos              : Integer;
  pSP, pSS, pSE, pPS: Integer;
  pTP               : TLongPoint;
  pInCH             : Boolean;
begin
  pSP   := fSelPO;
  pSS   := fSelST;
  pSE   := fSelEN;
  pPS   := GetPosAtCursor(Col, Row);
  pInCH := fPosInChars;
  fOffsets.Clear;
  if aLengths.Count > 0 then
  begin
    pPos    := 0;
    for pct := 0 to aLengths.Count - 1 do
    begin
      fOffsets.Add(Pointer(pPos));
      pPos := pPos + Integer(aLengths[pct]);
    end;
  end;
  CalcSizes;
  pTP := GetCursorAtPos(pPS, pInCH);
  MoveColRow(pTP.x, pTP.y, True, True);
  fSelPO := pSP;
  fSelST := pSS;
  fSelEN := pSE;
  Invalidate;
end;

function THexEditor.GetIsInsertMode: Boolean;
begin
  Result := fInsertOn and (not fNoSizeChange) and fAllowInsertMode;
end;

procedure THexEditor.SetAllowInsertMode(const aValue: Boolean);
begin
  if fNoSizeChange then
    fAllowInsertMode := False
  else
    fAllowInsertMode := aValue;
  StateNotification;
end;

procedure THexEditor.SetNoSizeChange(const aValue: Boolean);
begin
  fNoSizeChange   := aValue;
  AllowInsertMode := fAllowInsertMode;
end;

procedure THexEditor.StateNotification;
begin
  if HandleAllocated then
    PostMessage(Handle, WM_STATECHANGED, 7, 7);
end;

procedure THexEditor.InternalErase(const BackSp: Boolean);
var
  nPos: Integer;
begin
  nPos := GetCursorPos;
  if BackSp then
  begin
    if nPos = 0 then
      Exit;
    if not CreateUndo(U_Byte_removed, nPos - 1, 1, 0) then
      Exit;
    InternalDeleteSelection(nPos - 1, nPos, Col, Row);
    Seek(nPos - 1, soFromBeginning, True);
  end
  else
  begin
    if nPos = DataSize then
      Exit;
    if CreateUndo(U_Byte_removed, nPos, 1, 0) then
      InternalDeleteSelection(nPos, nPos + 1, Col, Row);
  end;
end;

procedure THexEditor.SetAutoCaretMode(const aValue: Boolean);
begin
  fAutoCaretMode := aValue;
  if aValue then
  begin
    if IsInsertMode then
      CaretStyle := csLeftLine
    else
      CaretStyle := csFull;
  end;
end;

constructor TColors.Create(Parent: TControl);
begin
  inherited Create;
  FBackground         := clWindow;
  FPositionText       := clWhite;
  FChangedText        := clMaroon;
  FCursorFrame        := clNavy;
  FOffset             := clBlack;
  FOddColumn          := clBlue;
  FEvenColumn         := clNavy;
  FOddInverted        := Invert(FOddColumn);
  FEvenInverted       := Invert(FEvenColumn);
  FChangedBackground  := $00A8FFFF;
  FPositionBackground := clMaroon;
  FParent             := Parent;
end;

procedure TColors.SetBackground(const Value: TColor);
begin
  if FBackground <> Value then
  begin
    FBackground               := Value;
    THexEditor(FParent).Color := Value;
    FParent.Repaint;
  end;
end;

procedure TColors.SetChangedBackground(const Value: TColor);
begin
  if FChangedBackground <> Value then
  begin
    FChangedBackground := Value;
    FParent.Invalidate;
  end;
end;

procedure TColors.SetChangedText(const Value: TColor);
begin
  if FChangedText <> Value then
  begin
    FChangedText := Value;
    FParent.Invalidate;
  end;
end;

procedure TColors.SetCursorFrame(const Value: TColor);
begin
  if FCursorFrame <> Value then
  begin
    FCursorFrame := Value;
    FParent.Invalidate;
  end;
end;

procedure TColors.SetEvenColumn(const Value: TColor);
begin
  if FEvenColumn <> Value then
  begin
    FEvenColumn   := Value;
    FEvenInverted := Invert(FEvenColumn);
    FParent.Invalidate;
  end;
end;

procedure TColors.SetOddColumn(const Value: TColor);
begin
  if FOddColumn <> Value then
  begin
    FOddColumn   := Value;
    FOddInverted := Invert(FOddColumn);
    FParent.Invalidate;
  end;
end;

procedure TColors.SetOffset(const Value: TColor);
begin
  if FOffset <> Value then
  begin
    FOffset := Value;
    FParent.Invalidate;
  end;
end;

procedure TColors.SetPositionBackground(const Value: TColor);
begin
  if FPositionBackground <> Value then
  begin
    FPositionBackground := Value;
    FParent.Invalidate;
  end;
end;

procedure TColors.SetPositionText(const Value: TColor);
begin
  if FPositionText <> Value then
  begin
    FPositionText := Value;
    FParent.Invalidate;
  end;
end;

Constructor THexToCanvas.Create(aOwner: TComponent);
begin
  Inherited Create(aOwner);
  fHexEditor   := nil;
  fFont        := TFont.Create;
  Font.Name    := 'Courier';
  Font.Size    := 12;
  fBpL         := 16;
  fOffsDy      := odHex;
  fOffsCr      := ':';
  fMemDy       := odHex;
  fMemCr       := ';';
  fCharDy      := True;
  fCharCr      := #0;
  fShrink      := True;
  fStretch     := True;
  fBpC         := 2;
  fSwapNibbles := False;
end;

Destructor THexToCanvas.Destroy;
begin
  fFont.Free;
  inherited;
end;

procedure THexToCanvas.SetFont(Value: TFont);
begin
  fFont.Assign(Value);
end;

procedure THexToCanvas.SetHexEditor(Value: THexEditor);
begin
  fHexEditor := Value;
  if Value <> nil then
    Value.FreeNotification(Self);
end;

procedure THexToCanvas.Notification(aComponent: TComponent; aOperation: TOperation);
begin
  if fHexEditor <> nil then
    if aOperation = opRemove then
      if aComponent = fHexEditor then
        fHexEditor := nil;
end;

procedure THexToCanvas.GetLayout;
begin
  if fHexEditor <> nil then
  begin
    fFont.Assign(fHexEditor.Font);
    fBpC         := fHexEditor.BytesPerColumn;
    fOffsCr      := fHexEditor.OffsetSeparator;
    fOffsDy      := fHexEditor.OffsetDisplay;
    fBpL         := fHexEditor.BytesPerLine;
    fMemDy       := odHex;
    fMemCr       := ' ';
    fCharDy      := True;
    fCharCr      := #0;
    fSwapNibbles := Boolean(fHexEditor.SwapNibbles);
  end;
end;

function THexToCanvas.Draw(aCanvas: TCanvas; const aStart, aEnd: Integer; const TopLine, BottomLine: string): Integer;
  function GetOneLine(aPos, aEnd: Integer): string;
    function GetByteHex(aPos, aEnd: Integer): string;
    begin
      if aPos > aEnd then
        Result := '  '
      else
      begin
        Result := IntToHex(fHexEditor.GetByteAtPos(aPos), 2);
        if fSwapNibbles and (Length(Result) = 2) then
          Result := Result[2] + Result[1];
      end;
    end;
    function GetByteDec(aPos, aEnd: Integer): string;
    begin
      if aPos > aEnd then
        Result := '   '
      else
        Result := FillLeft(' ', IntToStr(fHexEditor.GetByteAtPos(aPos)), 3);
    end;
    function GetByteOctal(aPos, aEnd: Integer): string;
    begin
      if aPos > aEnd then
        Result := '    '
      else
        Result := FillLeft('0', IntToOctal(fHexEditor.GetByteAtPos(aPos)), 4);
    end;

  var
    pct: Integer;
  begin
    case fOffsDy of
      odNone:
        Result := '';
      odHex:
        Result := '0x' + IntToHex(aPos, fHexEditor.fOffSetDisplayWidth - 3);
      odDec:
        Result := FillLeft(' ', IntToStr(aPos), fHexEditor.fOffSetDisplayWidth - 1);
      odOctal:
        Result := 'o ' + FillLeft('0', IntToOctal(aPos), fHexEditor.fOffSetDisplayWidth - 3);
    end;
    if fOffsCr <> #0 then
      Result := Result + fOffsCr;
    if fMemDy = odHex then
    begin
      for pct := 1 to fBpL do
      begin
        Result := Result + GetByteHex(aPos - 1 + pct, aEnd);
        if (pct mod fBpC) = 0 then
          Result := Result + ' ';
      end;
    end
    else if fMemDy = odDec then
    begin
      for pct := 1 to fBpL do
      begin
        Result := Result + GetByteDec(aPos - 1 + pct, aEnd);
        if (pct mod fBpC) = 0 then
          Result := Result + ' ';
      end;
    end
    else if fMemDy = odOctal then
    begin
      for pct := 1 to fBpL do
      begin
        Result := Result + GetByteOctal(aPos - 1 + pct, aEnd);
        if (pct mod fBpC) = 0 then
          Result := Result + ' ';
      end;
    end;
    if fMemCr <> #0 then
      Result := Result + fMemCr;
    if fCharDy then
      for pct := 1 to fBpL do
        if (aPos + pct - 1) > aEnd then
          Result := Result + ' '
        else
          Result := Result + fHexEditor.TranslateToAnsiChar(fHexEditor.GetByteAtPos(aPos + pct - 1));
    if fCharCr <> #0 then
      Result := Result + fCharCr;
  end;

var
  tmpFont                   : TFont;
  OneLine                   : string;
  lLen, lHe, lPos, lup, pEnd: Integer;
begin
  Result := -1;
  if fBpL < 1 then
    Exit;

  if fHexEditor = nil then
  begin
    Result := MaxInt;
    Exit;
  end;

  pEnd := aEnd;
  if pEnd >= fHexEditor.DataSize then
    pEnd := fHexEditor.DataSize - 1;
  if aStart > pEnd then
    Exit;

  OneLine := GetOneLine(aStart, pEnd);
  tmpFont := TFont.Create;
  try
    tmpFont.Assign(aCanvas.Font);
    aCanvas.Font.Assign(fFont);
    if fStretch then
    begin
      lLen := aCanvas.TextWidth(OneLine);
      while lLen < (fRightM - fLeftM) do
      begin
        aCanvas.Font.Size := aCanvas.Font.Size + 1;
        lLen              := aCanvas.TextWidth(OneLine);
      end;
      while lLen > (fRightM - fLeftM) do
      begin
        aCanvas.Font.Size := aCanvas.Font.Size - 1;
        lLen              := aCanvas.TextWidth(OneLine);
      end;
    end;
    if fShrink then
    begin
      lLen := aCanvas.TextWidth(OneLine);
      while lLen > (fRightM - fLeftM) do
      begin
        aCanvas.Font.Size := aCanvas.Font.Size - 1;
        lLen              := aCanvas.TextWidth(OneLine);
      end;
    end;
    lHe := Round(aCanvas.TextHeight(OneLine) * 1.2);
    if lHe = aCanvas.TextHeight(OneLine) then
      Inc(lHe);
    lPos := aStart;
    lup  := fTopM;
    if TopLine <> '' then
    begin
      aCanvas.TextOut(fLeftM, lup, TopLine);
      lup := lup + lHe;
    end;
    if BottomLine <> '' then
      fBottomM := fBottomM - lHe;
    while (lHe + lup) <= fBottomM do
    begin
      aCanvas.TextOut(fLeftM, lup, OneLine);
      lPos := lPos + fBpL;
      if lPos > pEnd then
      begin
        lPos := pEnd + 1;
        Break;
      end;
      OneLine := GetOneLine(lPos, pEnd);
      lup     := lup + lHe;
    end;
    Result := lPos;
    if BottomLine <> '' then
      aCanvas.TextOut(fLeftM, fBottomM, BottomLine);
  finally
    aCanvas.Font.Assign(tmpFont);
    tmpFont.Free;
  end;
end;

end.
