unit XLSReadRTF;

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

uses Classes, SysUtils, Contnrs, StrUtils, IniFiles, Math,
     XLSUtils5,
     XLSAXWEditor;

type TRtfJust = (justL, justR, justC, justF);
type TRtfSBK  = (sbkNon, sbkCol, sbkEvn, sbkOdd, sbkPg);
type TRtfPGN  = (pgDec, pgURom, pgLRom, pgULtr, pgLLtr);
// Underline

type TRtfULType   = (
{ axcuNone            } rultNone,
{ axcuWords           } rultW,
{ axcuDouble          } rultDB,
{ axcuThick           } rultTH,
{ axcuDotted          } rultD,
{ axcuDottedHeavy     } rultThD,
{ axcuDash            } rultDash,
{ axcuDashedHeavy     } rultThDash,
{ axcuDashLong        } rultLDash,
{ axcuDashLongHeavy   } rultThLDash,
{ axcuDotDash         } rultDashD,
{ axcuDashDotHeavy    } rultThDashD,
{ axcuDotDotDash      } rultDashDD,
{ axcuDashDotDotHeavy } rultThDashDD,
{ axcuWave            } rultWave,
{ axcuWavyHeavy       } rultHWave,
{ axcuWavyDouble      } rultUlDbWave);

// What types of properties are there?
type TRtfIPROP = (
ipropNil,
// CHP =========================================================================
ipropBold, ipropItalic, ipropUnderline, ipropUnderlineNone, ipropUnderlineColor, ipropULType, ipropFontSize,
ipropChBrdr, ipropChCbPat, ipropSub, ipropSuper, ipropStrikeTrough, ipropNoSuperSub,
ipropFontColor,
// PAP =========================================================================
ipropLeftInd, ipropRightInd, ipropFirstInd, ipropJust,
ipropTabPos, ipropTabC, ipropTabR, ipropTabDec, ipropTabLDot, ipropTabLHyph, ipropTabLULine,
ipropTabLThick, ipropListLevel,ipropBorderStyle, ipropSpaceBefore, ipropSpaceAfter,
ipropLineSpacing, ipropRTL, ipropParaColor,

ipropBrdrt,ipropBrdrb,ipropBrdrl,ipropBrdrr,ipropBrdrbtw,ipropBrdrbar,ipropBox,ipropBrdrs,
ipropBrdrw,ipropBrdrcf,ipropBrsp,ipropBrdrnil,ipropBrdrtbl,

// SEP =========================================================================
 ipropCols, ipropPgnX, ipropPgnY, ipropSbk, ipropPgnFormat,
// DOP =========================================================================
ipropXaPage, ipropYaPage, ipropXaLeft, ipropXaRight, ipropYaTop, ipropYaBottom,
ipropPgnStart, ipropFacingp, ipropLandscape,

ipropPar, ipropPard, ipropPlain, ipropSectd,

// Table =======================================================================
ipropIntbl, ipropITap, ipropTrowD, ipropCell, ipropNestCell, ipropRow, ipropNestRow,
ipropLastRow, ipropCellX,
ipropTblBorderL,ipropTblBorderT,ipropTblBorderR,ipropTblBorderB,ipropTblBorderIV,ipropTblBorderIH,
ipropCellBorderL,ipropCellBorderT,ipropCellBorderR,ipropCellBorderB,
ipropRowHeight,ipropCellColor,ipropCellNoShading,
ipropTblCellMargL,ipropTblCellMargT,ipropTblCellMargR,ipropTblCellMargB,
ipropCellVAlign,

// Colors ======================================================================
ipropRed, ipropGreen, ipropBlue,
// Char ========================================================================
ipropUnicode, ipropUnicodeGroup, ipropTab, ipropPageBreak, ipropLineBreak,
// Picture =====================================================================
ipropPicture, ipropPictWidth, ipropPictHeigth, ipropPictGoalWidth, ipropPictGoalHeight,
ipropPictScaleX, ipropPictScaleY, ipropPictCropT, ipropPictCropB, ipropPictCropL,
ipropPictCropR, ipropBlipTag,
// Font table ==================================================================
ipropFont, ipropCharSet,
// Bookmark ====================================================================
ipropBkmkColFirst,ipropBkmkColLast,
// Bullets and numberings ======================================================
ipropLS,
// Styles ======================================================================
ipropParaStyle, ipropCharRunStyle, ipropTableStyle, ipropSectStyle,
// Other =======================================================================
ipropDataField,
// End =========================================================================
ipropMax);

type TRtfSEP = record
     cCols: integer;                  // number of columns
     SBK: TRtfSBK;                    // section break type
     xaPgn: integer;                  // x position of page number in twips
     yaPgn: integer;                  // y position of page number in twips
     pgnFormat: TRtfPGN;              // how the page number is formatted
     end;                             // SEction Properties

type TRtfDOP = record
     xaPage: integer;                 // page width in twips
     yaPage: integer;                 // page height in twips
     xaLeft: integer;                 // left margin in twips
     yaTop: integer;                  // top margin in twips
     xaRight: integer;                // right margin in twips
     yaBottom: integer;               // bottom margin in twips
     pgnStart: integer;               // starting page number in twips
     fFacingp: byte;                  // facing pages enabled?
     fLandscape: byte;                // landscape or portrait??
     end;                             // DOcument Properties

               // Rtf Destination State
type TRtfRDS = (rdsNorm, rdsSkip);

type TRtfRIS = (risNorm, risBin, risHex);               // Rtf Internal State

type TRtfIPFN = (ipfnBin, ipfnHex, ipfnSkipDest);
type TRtfIDEST = (idestNil,idestSkip, idestFontTbl, idestColorTbl,
                  idestPict,idestShpPict,idestNonShpPict,idestBlipUid,
                  idestNestTableProps, idestNoNestTables,
//                  idestBkmkStart, idestBkmkEnd,
                  idestField, idestFieldInstr, idestFieldResult,
                  idestStylesheet,
                  idestAnnotation,idestAnnStart,idestAnnEnd,idestAnnRef,idestAnnTime,
                  idestAnnDate,idestAnnAuthor,idestAnnInitials,idestAnnParent,
                  idestAnnIcon,
                  idestGenerator,
                  idestBullet);

type TRtfKWD = (kwdChar, kwdDest, kwdProp, kwdSpec);

type TRTFStack = class;
     TRtfDestColorTable = class;
     TRtfDestFontTable = class;
     TAXWReadRTF = class;

     TRtfPAP = class(TAXWPAPX)
protected
     FInTable   : boolean;
     FTblNesting: integer;
     FCurrBorder: TAXWDocPropBorder;
     FColorTable: TRtfDestColorTable;
public
     constructor Create(APAP: TAXWPAP; AColorTable: TRtfDestColorTable);

     procedure Clear; override;

     procedure Assign(APAP: TRtfPAP); overload;

     function  ApplyKeyword(iprop: TRtfIPROP; val: integer): boolean;

     function  DoBorders(ABorderDest: TRtfIPROP): TAXWDocPropBorder;

     property InTable   : boolean read FInTable write FInTable;
     property TblNesting: integer read FTblNesting write FTblNesting;
     end;

     TRtfCHP = class(TAXWCHPX)
protected
     FCurrBorder: TAXWDocPropBorder;
     FColorTable: TRtfDestColorTable;
     FFontTable : TRtfDestFontTable;
public
     constructor Create(ACHP: TAXWCHP; AColorTable: TRtfDestColorTable; AFontTable: TRtfDestFontTable);
     destructor Destroy; override;

     procedure Clear; override;

     procedure Assign(ACHP: TRtfCHP); overload;

     function  AddCurrBorder: TAXWDocPropBorder;
     function  ApplyKeyword(iprop: TRtfIPROP; val: integer): boolean;

     property CurrBorder: TAXWDocPropBorder read FCurrBorder;
     end;


     TRtfDestination = class(TObject)
protected
     FType         : TRtfIDEST;
     FReader       : TAXWReadRTF;
     FParentDest   : TRtfDestination;
     FGroupLevel   : integer;
     FText         : AxUCString;
public
     constructor Create(AReader: TAXWReadRTF; AType: TRtfIDEST); overload;
     constructor Create(AReader: TAXWReadRTF; AType: TRtfIDEST; AParentDest: TRtfDestination); overload;
     destructor Destroy; override;

     function  Persistent: boolean; virtual;
     function  HasDocText: boolean; virtual;

     procedure ChangeType(ANewType: TRtfIDEST); virtual;

     procedure BeginDest; virtual;
     procedure EndDest; virtual;
     procedure ChildEndDest(AChild: TRtfDestination); virtual;
     procedure ParseChar(var ARIS: TRtfRIS; AChar: AnsiChar); virtual;
     procedure BeginGroup; virtual;
     procedure EndGroup; virtual;
     procedure ApplyKeyword(iprop: TRtfIPROP; val: integer); virtual;
     procedure EndComposedValue; virtual;

     property Type_: TRtfIDEST read FType;
     end;

     TRtfDestStream = class(TRtfDestination)
protected
     FCurrByte     : byte;
     FCurrByteUpper: boolean;
     FStream       : TMemoryStream;
public
     constructor Create(AReader: TAXWReadRTF; AType: TRtfIDEST; AParentDest: TRtfDestination); overload;
     constructor Create(AReader: TAXWReadRTF; AType: TRtfIDEST); overload;
     destructor Destroy; override;

     procedure ParseChar(var ARIS: TRtfRIS; AChar: AnsiChar); override;
     procedure EndDest; override;
     end;

     TRtfDestColorTable = class(TRtfDestination)
protected
     FColorTable: array of longword;
     FR: byte;
     FG: byte;
     FB: byte;
public
     function  Persistent: boolean; override;

     procedure ApplyKeyword(iprop: TRtfIPROP; val: integer); override;
     procedure EndComposedValue; override;

     function  Color(AClIndex: integer; ADefault: longword): longword;
     end;

     TRtfFont = class(TObject)
protected
     FName: AxUCString;
     FId  : integer;
public
     property Name: AxUCString read FName write FName;
     property Id: integer read FId write FId;
     end;

     TRtfDestFontTable = class(TRtfDestination)
private
     function GetFont(AId: integer): TRtfFont;
protected
     FFonts      : THashedStringList;
     FDefaultFont: TRtfFont;
     FCurrFont   : TRtfFont;
public
     constructor Create(AReader: TAXWReadRTF; AType: TRtfIDEST);
     destructor Destroy; override;

     function  Persistent: boolean; override;

     procedure ParseChar(var ARIS: TRtfRIS; AChar: AnsiChar); override;
     procedure ApplyKeyword(iprop: TRtfIPROP; val: integer); override;
     procedure EndComposedValue; override;
     procedure EndGroup; override;

     property Font[AId: integer]: TRtfFont read GetFont; default;
     end;

//     TRtfDestBookmark = class(TRtfDestination)
//protected
//     FBookmarks: THashedStringList;
//     FColFirst : integer;
//     FColLast  : integer;
//     FSel      : TAXWFieldSelectionImpl;
//public
//     constructor Create(AReader: TAXWReadRTF; AType: TRtfIDEST; ABookmarks: THashedStringList);
//     destructor Destroy; override;
//
//     function  Persistent: boolean; override;
//
//     procedure ApplyKeyword(iprop: TRtfIPROP; val: integer); override;
//     procedure EndGroup; override;
//     end;

     TRftDestListText = class(TRtfDestination)
protected
public
     constructor Create(AReader: TAXWReadRTF; AType: TRtfIDEST);

     procedure ApplyKeyword(iprop: TRtfIPROP; val: integer); override;
     procedure EndGroup; override;

     procedure ParseChar(var ARIS: TRtfRIS; AChar: AnsiChar); override;
     end;

     TRtfStyle = class(TObject)
protected
     FReader: TAXWReadRTF;
     FId    : integer;

     procedure AddStyle(AStyle: TAXWStyle);
public
     constructor Create(AReader: TAXWReadRTF; AId: integer);

     procedure ApplyKeyword(iprop: TRtfIPROP; val: integer); virtual; abstract;
     procedure Save(AName: AxUCString); virtual; abstract;
     end;

     TRtfCharStyle = class(TRtfStyle)
protected
     FCHPX: TRtfCHP;
public
     constructor Create(AReader: TAXWReadRTF; AId: integer);
     destructor Destroy; override;

     procedure ApplyKeyword(iprop: TRtfIPROP; val: integer); override;
     procedure Save(AName: AxUCString); override;
     end;

     TRtfParaStyle = class(TRtfStyle)
protected
     FCHPX: TRtfCHP;
     FPAPX: TRtfPAP;
public
     constructor Create(AReader: TAXWReadRTF; AId: integer);
     destructor Destroy; override;

     procedure ApplyKeyword(iprop: TRtfIPROP; val: integer); override;
     procedure Save(AName: AxUCString); override;
     end;

     TRtfDestStylesheet = class(TRtfDestination)
protected
     FCurrStyle: TRtfStyle;
public
     procedure ApplyKeyword(iprop: TRtfIPROP; val: integer); override;
     procedure BeginGroup; override;
     procedure EndGroup; override;
     end;

     TRtfTableStyle = class(TRtfStyle)
protected
public
     procedure ApplyKeyword(iprop: TRtfIPROP; val: integer); override;
     procedure Save(AName: AxUCString); override;
     end;

     TRtfSectStyle = class(TRtfStyle)
protected
public
     procedure ApplyKeyword(iprop: TRtfIPROP; val: integer); override;
     procedure Save(AName: AxUCString); override;
     end;

     TRtfDestField = class(TRtfDestination)
protected
     FInstrDone: boolean;
public
     function  HasDocText: boolean; override;
     procedure ChangeType(ANewType: TRtfIDEST); override;

     procedure EndDest; override;
     procedure ParseChar(var ARIS: TRtfRIS; AChar: AnsiChar); override;
     procedure ApplyKeyword(iprop: TRtfIPROP; val: integer); override;
     end;

     TRtfDestComment = class(TRtfDestination)
protected
     FComments: TStrings;
     FRef     : AxUCString;
public
     constructor Create(AReader: TAXWReadRTF; AType: TRtfIDEST; AComments: TStrings);

     procedure ApplyKeyword(iprop: TRtfIPROP; val: integer); override;

     procedure EndDest; override;
     procedure ChildEndDest(AChild: TRtfdestination); override;
     end;

     TRtfDestStringParam = class(TRtfDestination)
protected
public
     procedure EndDest; override;
     end;

     TRtfDestPict = class(TRtfDestStream)
protected
     FPictType  : TAXWPictureType;
//     FWrap      : TAXWGrObjWrap;
     FExData    : integer; // N value for metafiles and bitmaps.
     FWidth     : integer;
     FHeight    : integer;
     FGoalWidth : double;
     FGoalHeight: double;
     FScaleX    : double;
     FScaleY    : double;
     FCropLeft  : integer;
     FCropTop   : integer;
     FCropRight : integer;
     FCropBottom: integer;
     FTag       : integer;
public
     constructor Create(AReader: TAXWReadRTF; AType: TRtfIDEST);
     destructor Destroy; override;

     procedure EndGroup; override;
     procedure ApplyKeyword(iprop: TRtfIPROP; val: integer); override;

     procedure EndDest; override;
     procedure Clear;

     property PictType  : TAXWPictureType read FPictType write FPictType;
     property ExData    : integer read FExData write FExData;
     property Width     : integer read FWidth write FWidth;
     property Height    : integer read FHeight write FHeight;
     property GoalWidth : double read FGoalWidth write FGoalWidth;
     property GoalHeight: double read FGoalHeight write FGoalHeight;
     property ScaleX    : double read FScaleX write FScaleX;
     property ScaleY    : double read FScaleY write FScaleY;
     property CropLeft  : integer read FCropLeft write FCropLeft;
     property CropTop   : integer read FCropTop write FCropTop;
     property CropRight : integer read FCropRight write FCropRight;
     property CropBottom: integer read FCropBottom write FCropBottom;
     property Tag       : integer read FTag write FTag;
     end;

     TRTFStackItem = class(TObject)
private
     FOwner: TRTFStack;
     FRDS: TRtfRDS;
     FRIS: TRtfRIS;
     FCHP: TRtfCHP;
     FPAP: TRtfPAP;
     FSEP: TRtfSEP;
     FDOP: TRtfDOP;
     FDest: TRtfDestination;
public
     constructor Create(AOwner: TRTFStack; RDS: TRtfRDS; RIS: TRtfRIS; CHP: TRtfCHP; PAP: TRtfPAP; SEP: TRtfSEP; DOP: TRtfDOP);
     destructor Destroy; override;

     property RDS: TRtfRDS read FRDS write FRDS;
     property RIS: TRtfRIS read FRIS write FRIS;
     property CHP: TRtfCHP read FCHP write FCHP;
     property PAP: TRtfPAP read FPAP write FPAP;
     property SEP: TRtfSEP read FSEP write FSEP;
     property DOP: TRtfDOP read FDOP write FDOP;
     property Dest: TRtfDestination read FDest write FDest;
     end;

     TRTFStack = class(TObjectList)
private
     FOwner: TAXWReadRTF;
     FCurrDest: TRtfDestination;

     function GetItems(Index: integer): TRTFStackItem;
     function FindDest: TRtfDestination;
public
     constructor Create(AOwner: TAXWReadRTF);

     procedure Push(RDS: TRtfRDS; RIS: TRtfRIS; CHP: TRtfCHP; PAP: TRtfPAP; SEP: TRtfSEP; DOP: TRtfDOP);
     procedure Pop(var RDS: TRtfRDS; var RIS: TRtfRIS; CHP: TAXWCHPX; PAP: TRtfPAP; var SEP: TRtfSEP; var DOP: TRtfDOP; var ADest: TRtfDestination);
     procedure SetTOSDest(ADest: TRtfDestination);

     function  Peek: TRTFStackItem;
     function  PeekL2: TRTFStackItem;
     function  PeekDestTypeL2: TRtfIDEST;
     function  CurrDest: TRtfDestination;

     property Items[Index: integer]: TRTFStackItem read GetItems; default;
     end;

     TRtfTableStackItem = class(TObject)
protected
     FCell     : TAXWTableCell;
     FCellX    : TAXWIntegerList;
     FCellColor: TAXWIntegerList;
     FBorders  : TAXWDocPropBordersList;
     FRowHeight: double;
public
     constructor Create;
     destructor Destroy; override;

     procedure Clear;
     procedure Assign(AItem: TRtfTableStackItem);

     property Cell: TAXWTableCell read FCell write FCell;
     property CellX: TAXWIntegerList read FCellX;
     property Borders: TAXWDocPropBordersList read FBorders;
     property CellColor: TAXWIntegerList read FCellColor;
     property RowHeight: double read FRowHeight write FRowHeight;
     end;

     TRtfTableStack = class(TObjectStack)
protected
     FColorTable: TRtfDestColorTable;
     FCurrItem  : TRtfTableStackItem;
public
     constructor Create(AColorTable: TRtfDestColorTable);
     destructor Destroy; override;

     procedure Clear;
     procedure ClearRow;

     procedure ApplyOnRow(ARow: TAXWTableRow);

     procedure Push(ACell: TAXWTableCell);
     procedure Pop(var ATable: TAXWTable; var ARow: TAXWTableRow; var ACell: TAXWTableCell);
     function  Peek: TRtfTableStackItem;

     property CurrItem: TRtfTableStackItem read FCurrItem;
     end;

     TAXWReadRTF = class(TObject)
private
     FStack         : TRTFStack;
     FTableStack    : TRtfTableStack;
     cGroup         : integer;
     fSkipDestIfUnk : boolean;
     cbBin          : integer;
     lParam         : integer;
     FRDS           : TRtfRDS;
     FRIS           : TRtfRIS;
     FCHP           : TRtfCHP;
     FPAP           : TRtfPAP;
     FSEP           : TRtfSEP;
     FDOP           : TRtfDOP;
     FText          : AxUCString;
     FExText        : AxUCString;

     FCharRunType   : TAXWCharRunType;

     FFontTable     : TRtfDestFontTable;
     FColorTable    : TRtfDestColorTable;

     FSymbols       : THashedStringList;

     FDoc           : TAXWLogDocEditor;

     FTabLeader     : TAXWTabStopLeader;
     FTabAlignment  : TAXWTabStopAlignment;
     FBookmarks     : THashedStringList;
     FComments      : THashedStringList;
     FStyles        : THashedStringList;

     FCurrPara      : TAXWLogPara;
     FCurrStream    : TMemoryStream;
     FCurrBorder    : TAXWDocPropBorder;
     FCurrCell      : TAXWTableCell;
     FCurrRow       : TAXWTableRow;
     FCurrTable     : TAXWTable;
//     FCurrSimpField : TAXWSimpleField;
     FCurrHyperlink : TAXWHyperlink;
     FInheritPara   : boolean;

     FLastUnicodeChar: WideChar;

     FGenerator     : AxUCString;

     function  ecRtfParse(fp: TStream): integer;
     procedure ecPushRtfState;
     function  ecPopRtfState: integer;
     function  ecParseRtfKeyword(fp: TStream): integer;
     function  ecParseChar(ch: AnsiChar): integer;
     function  ecTranslateKeyword(szKeyword: AnsiString; param: integer; fParam: boolean; fp: TStream): integer;
     procedure ecPrintChar(ch: AxUCChar);
     function  ecEndGroupAction: integer;
     function  ecApplyPropChange(IPROP: TRtfIPROP; val: integer): integer;
     function  ecChangeDest(IDEST: TRtfIDEST; Param: integer): integer;
     function  ecParseSpecialKeyword(IPFN: TRtfIPFN): integer;
     function  GetPara: TAXWLogPara;
     procedure AddText;
     procedure AddTable(AParas: TAXWLogParas);
     procedure CheckCell;
     procedure CheckTable;
//     function ecParseHexByte: integer;
public
     constructor Create(AEditor: TAXWLogDocEditor);
     destructor Destroy; override;

     procedure LoadFromFile(const Filename: AxUCString); overload;
     procedure LoadFromStream(Stream: TStream); overload;

     property Generator: AxUCString read FGenerator;
     end;

// RTF parser error codes

const ecOK                = 0;       // Everything's fine!
const ecStackUnderflow    = 1;       // Unmatched '}'
const ecStackOverflow     = 2;       // Too many '{' -- memory exhausted
const ecUnmatchedBrace    = 3;       // RTF ended during an open group.
const ecInvalidHex        = 4;       // invalid hex character found in data
const ecBadTable          = 5;       // RTF table (sym or prop) invalid
const ecAssertion         = 6;       // Assertion failure
const ecEndOfFile         = 7;       // End of file reached while reading RTF

implementation

type TSYM = record
    Keyword  : AnsiString;      // RTF keyword
    dflt     : integer;         // default value to use
    fPassDflt: boolean;         // true to use default value from this table
    KWD      : TRtfKWD;         // base action to take
    idx      : integer;         // index into property table if kwd == kwdProp
                                // index into destination table if kwd == kwdDest
                                // character to print if kwd == kwdChar
    end;

const rgsymRtf: array[0..244] of TSYM = (
(    Keyword: 'rtf';         dflt:                      0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropNil)),
(    Keyword: 'rtlch';       dflt:                      0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropNil)),
(    Keyword: 'ltrch';       dflt:                      0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropNil)),


(    Keyword: 'plain';       dflt:                      1; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropPlain)),
(    Keyword: 'b';           dflt:                      1; fPassDflt: True ; KWD: kwdProp; idx: Integer(ipropBold)),
(    Keyword: 'ul';          dflt:                      1; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropUnderline)),
(    Keyword: 'ulnone';      dflt:                      1; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropUnderlineNone)),
(    Keyword: 'ulc';         dflt:                      0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropUnderlineColor)),
(    Keyword: 'strike';      dflt:                      1; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropStrikeTrough)),

(    Keyword: 'uldash';      dflt:          Ord(rultDash); fPassDflt: False; KWD: kwdProp; idx: Integer(ipropULType)),
(    Keyword: 'uldashd';     dflt:         Ord(rultDashD); fPassDflt: False; KWD: kwdProp; idx: Integer(ipropULType)),
(    Keyword: 'uldashdd';    dflt:        Ord(rultDashDD); fPassDflt: False; KWD: kwdProp; idx: Integer(ipropULType)),
(    Keyword: 'uldb';        dflt:            Ord(rultDB); fPassDflt: False; KWD: kwdProp; idx: Integer(ipropULType)),
(    Keyword: 'ulhwave';     dflt:         Ord(rultHWave); fPassDflt: False; KWD: kwdProp; idx: Integer(ipropULType)),
(    Keyword: 'ulldash';     dflt:         Ord(rultLDash); fPassDflt: False; KWD: kwdProp; idx: Integer(ipropULType)),
(    Keyword: 'uld';         dflt:             Ord(rultD); fPassDflt: False; KWD: kwdProp; idx: Integer(ipropULType)),
(    Keyword: 'ulth';        dflt:            Ord(rultTH); fPassDflt: False; KWD: kwdProp; idx: Integer(ipropULType)),
(    Keyword: 'ulthd';       dflt:           Ord(rultTHD); fPassDflt: False; KWD: kwdProp; idx: Integer(ipropULType)),
(    Keyword: 'ulthdash';    dflt:        Ord(rultTHDash); fPassDflt: False; KWD: kwdProp; idx: Integer(ipropULType)),
(    Keyword: 'ulthdashd';   dflt:       Ord(rultTHDashD); fPassDflt: False; KWD: kwdProp; idx: Integer(ipropULType)),
(    Keyword: 'ulthdashdd';  dflt:      Ord(rultTHDashDD); fPassDflt: False; KWD: kwdProp; idx: Integer(ipropULType)),
(    Keyword: 'ulthldash';   dflt:       Ord(rultTHLDash); fPassDflt: False; KWD: kwdProp; idx: Integer(ipropULType)),
(    Keyword: 'ululdbwave';  dflt:      Ord(rultUlDbWave); fPassDflt: False; KWD: kwdProp; idx: Integer(ipropULType)),
(    Keyword: 'ulw';         dflt:             Ord(rultW); fPassDflt: False; KWD: kwdProp; idx: Integer(ipropULType)),
(    Keyword: 'ulwave';      dflt:          Ord(rultWave); fPassDflt: False; KWD: kwdProp; idx: Integer(ipropULType)),

(    Keyword: 'i';           dflt:                      1; fPassDflt: True ; KWD: kwdProp; idx: Integer(ipropItalic)),
(    Keyword: 'fs';          dflt:                     24; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropFontSize)),
(    Keyword: 'cf';          dflt:                      0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropFontColor)),
(    Keyword: 'li';          dflt:                      0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropLeftInd)),
(    Keyword: 'ri';          dflt:                      0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropRightInd)),
(    Keyword: 'fi';          dflt:                      0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropFirstInd)),
(    Keyword: 'ilvl';        dflt:                      0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropListLevel)),
(    Keyword: 'chbrdr';      dflt:                      0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropChBrdr)),
(    Keyword: 'chcbpat';     dflt:                      0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropChcbpat)),

(    Keyword: 'sub';         dflt:                      1; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropSub)),
(    Keyword: 'super';       dflt:                      1; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropSuper)),
(    Keyword: 'nosupersub';  dflt:                      1; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropNoSuperSub)),

(    Keyword: 'brdrt';       dflt:                      0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropBrdrt)),
(    Keyword: 'brdrb';       dflt:                      0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropBrdrb)),
(    Keyword: 'brdrl';       dflt:                      0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropBrdrl)),
(    Keyword: 'brdrr';       dflt:                      0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropBrdrr)),

(    Keyword: 'brdrbtw';     dflt:                      0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropBrdrbtw)),
(    Keyword: 'brdrbar';     dflt:                      0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropBrdrbar)),
(    Keyword: 'box';         dflt:                      0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropBox)),

(    Keyword: 'brdrw';       dflt:                      0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropBrdrw)),
(    Keyword: 'brdrcf';      dflt:                      0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropBrdrcf)),
(    Keyword: 'brsp';        dflt:                      0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropBrsp)),
(    Keyword: 'brdrnil';     dflt:                      0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropBrdrnil)),
(    Keyword: 'brdrtbl';     dflt:                      0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropBrdrtbl)),

(    Keyword: 'brdrs';       dflt:         Ord(stbSingle); fPassDflt: False; KWD: kwdProp; idx: Integer(ipropBorderStyle)),
(    Keyword: 'brdrth';      dflt:          Ord(stbThick); fPassDflt: False; KWD: kwdProp; idx: Integer(ipropBorderStyle)),
//(    Keyword: 'brdrsh'; dflt: Ord(); fPassDflt: False; KWD: kwdProp; idx: Integer(ipropBorderStyle)),
(    Keyword: 'brdrdb';      dflt:         Ord(stbDouble); fPassDflt: False; KWD: kwdProp; idx: Integer(ipropBorderStyle)),
(    Keyword: 'brdrdot';     dflt:         Ord(stbDotted); fPassDflt: False; KWD: kwdProp; idx: Integer(ipropBorderStyle)),
(    Keyword: 'brdrdash';    dflt:         Ord(stbDashed); fPassDflt: False; KWD: kwdProp; idx: Integer(ipropBorderStyle)),
(    Keyword: 'brdrhair';    dflt:         Ord(stbSingle); fPassDflt: False; KWD: kwdProp; idx: Integer(ipropBorderStyle)),
(    Keyword: 'brdrinset';   dflt:          Ord(stbInset); fPassDflt: False; KWD: kwdProp; idx: Integer(ipropBorderStyle)),
(    Keyword: 'brdrdashsm';  dflt:   Ord(stbDashSmallGap); fPassDflt: False; KWD: kwdProp; idx: Integer(ipropBorderStyle)),
(    Keyword: 'brdrdashd';   dflt:        Ord(stbDotDash); fPassDflt: False; KWD: kwdProp; idx: Integer(ipropBorderStyle)),
(    Keyword: 'brdrdashdd';  dflt:     Ord(stbDotDotDash); fPassDflt: False; KWD: kwdProp; idx: Integer(ipropBorderStyle)),
(    Keyword: 'brdroutset';  dflt:         Ord(stbOutset); fPassDflt: False; KWD: kwdProp; idx: Integer(ipropBorderStyle)),
(    Keyword: 'brdrtriple';  dflt:         Ord(stbTriple); fPassDflt: False; KWD: kwdProp; idx: Integer(ipropBorderStyle)),
(    Keyword: 'brdrtnthsg';  dflt: Ord(stbThickThinSmallGap); fPassDflt: False; KWD: kwdProp; idx: Integer(ipropBorderStyle)),
(    Keyword: 'brdrthtnsg';  dflt: Ord(stbThinThickSmallGap); fPassDflt: False; KWD: kwdProp; idx: Integer(ipropBorderStyle)),
(    Keyword: 'brdrtnthtnsg';dflt: Ord(stbThinThickThinSmallGap); fPassDflt: False; KWD: kwdProp; idx: Integer(ipropBorderStyle)),
(    Keyword: 'brdrtnthmg';  dflt: Ord(stbThickThinMediumGap); fPassDflt: False; KWD: kwdProp; idx: Integer(ipropBorderStyle)),
(    Keyword: 'brdrthtnmg';  dflt: Ord(stbThinThickMediumGap); fPassDflt: False; KWD: kwdProp; idx: Integer(ipropBorderStyle)),
(    Keyword: 'brdrtnthtnmg';dflt: Ord(stbThinThickThinMediumGap); fPassDflt: False; KWD: kwdProp; idx: Integer(ipropBorderStyle)),
(    Keyword: 'brdrtnthlg';  dflt: Ord(stbThickThinLargeGap); fPassDflt: False; KWD: kwdProp; idx: Integer(ipropBorderStyle)),
(    Keyword: 'brdrthtnlg';  dflt: Ord(stbThinThickLargeGap); fPassDflt: False; KWD: kwdProp; idx: Integer(ipropBorderStyle)),
(    Keyword: 'brdrtnthtnlg';dflt: Ord(stbThinThickThinLargeGap); fPassDflt: False; KWD: kwdProp; idx: Integer(ipropBorderStyle)),
(    Keyword: 'brdrwavy';    dflt:           Ord(stbWave); fPassDflt: False; KWD: kwdProp; idx: Integer(ipropBorderStyle)),
(    Keyword: 'brdrwavydb';  dflt:     Ord(stbDoubleWave); fPassDflt: False; KWD: kwdProp; idx: Integer(ipropBorderStyle)),
//(    Keyword: 'brdrdashdotstr'; dflt: Ord(); fPassDflt: False; KWD: kwdProp; idx: Integer(ipropBorderStyle)),
(    Keyword: 'brdremboss';  dflt:   Ord(stbThreeDEmboss); fPassDflt: False; KWD: kwdProp; idx: Integer(ipropBorderStyle)),
(    Keyword: 'brdrengrave'; dflt:  Ord(stbThreeDEngrave); fPassDflt: False; KWD: kwdProp; idx: Integer(ipropBorderStyle)),
//(    Keyword: 'brdrframe'; dflt: Ord(); fPassDflt: False; KWD: kwdProp; idx: Integer(ipropBorderStyle)),

// Bullets and numberings ======================================================
(    Keyword: 'pn';          dflt:                  0; fPassDflt: False; KWD: kwdDest; idx: Integer(idestSkip)),
(    Keyword: 'pntext';      dflt:                  0; fPassDflt: False; KWD: kwdDest; idx: Integer(idestSkip)),
(    Keyword: 'pntxtb';      dflt:                  0; fPassDflt: False; KWD: kwdDest; idx: Integer(idestSkip)),
(    Keyword: 'listtext';    dflt:                  0; fPassDflt: False; KWD: kwdDest; idx: Integer(idestSkip)),
(    Keyword: 'ls';          dflt:                  0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropLS)),

(    Keyword: 'cols';        dflt:                  1; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropCols)),
(    Keyword: 'sbknone';     dflt:    Integer(sbkNon); fPassDflt:  True; KWD: kwdProp; idx: Integer(ipropSbk)),
(    Keyword: 'sbkcol';      dflt:    Integer(sbkCol); fPassDflt:  True; KWD: kwdProp; idx: Integer(ipropSbk)),
(    Keyword: 'sbkeven';     dflt:    Integer(sbkEvn); fPassDflt:  True; KWD: kwdProp; idx: Integer(ipropSbk)),
(    Keyword: 'sbkodd';      dflt:    Integer(sbkOdd); fPassDflt:  True; KWD: kwdProp; idx: Integer(ipropSbk)),
(    Keyword: 'sbkpage';     dflt:     Integer(sbkPg); fPassDflt:  True; KWD: kwdProp; idx: Integer(ipropSbk)),
(    Keyword: 'pgnx';        dflt:                  0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropPgnX)),
(    Keyword: 'pgny';        dflt:                  0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropPgnY)),
(    Keyword: 'pgndec';      dflt:     Integer(pgDec); fPassDflt:  True; KWD: kwdProp; idx: Integer(ipropPgnFormat)),
(    Keyword: 'pgnucrm';     dflt:    Integer(pgURom); fPassDflt:  True; KWD: kwdProp; idx: Integer(ipropPgnFormat)),
(    Keyword: 'pgnlcrm';     dflt:    Integer(pgLRom); fPassDflt:  True; KWD: kwdProp; idx: Integer(ipropPgnFormat)),
(    Keyword: 'pgnucltr';    dflt:    Integer(pgULtr); fPassDflt:  True; KWD: kwdProp; idx: Integer(ipropPgnFormat)),
(    Keyword: 'pgnlcltr';    dflt:    Integer(pgLLtr); fPassDflt:  True; KWD: kwdProp; idx: Integer(ipropPgnFormat)),
(    Keyword: 'qc';          dflt:  Byte(axptaCenter); fPassDflt:  True; KWD: kwdProp; idx: Integer(ipropJust)),
(    Keyword: 'ql';          dflt:    Byte(axptaLeft); fPassDflt:  True; KWD: kwdProp; idx: Integer(ipropJust)),
(    Keyword: 'qr';          dflt:   Byte(axptaRight); fPassDflt:  True; KWD: kwdProp; idx: Integer(ipropJust)),
(    Keyword: 'qj';          dflt: Byte(axptaJustify); fPassDflt:  True; KWD: kwdProp; idx: Integer(ipropJust)),
(    Keyword: 'sb';          dflt:                  0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropSpaceBefore)),
(    Keyword: 'sa';          dflt:                  0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropSpaceAfter)),
(    Keyword: 'sl';          dflt:                  0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropLineSpacing)),
(    Keyword: 'rtlpar';      dflt:                  1; fPassDflt:  True; KWD: kwdProp; idx: Integer(ipropRTL)),
(    Keyword: 'ltrpar';      dflt:                  0; fPassDflt:  True; KWD: kwdProp; idx: Integer(ipropRTL)),
(    Keyword: 'cbpat';       dflt:                  0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropParaColor)),

(    Keyword: 'paperw';      dflt:              12240; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropXaPage)),
(    Keyword: 'paperh';      dflt:              15480; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropYaPage)),
(    Keyword: 'margl';       dflt:               1800; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropXaLeft)),
(    Keyword: 'margr';       dflt:               1800; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropXaRight)),
(    Keyword: 'margt';       dflt:               1440; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropYaTop)),
(    Keyword: 'margb';       dflt:               1440; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropYaBottom)),
(    Keyword: 'pgnstart';    dflt:                  1; fPassDflt:  True; KWD: kwdProp; idx: Integer(ipropPgnStart)),
(    Keyword: 'facingp';     dflt:                  1; fPassDflt:  True; KWD: kwdProp; idx: Integer(ipropFacingp)),
(    Keyword: 'landscape';   dflt:                  1; fPassDflt:  True; KWD: kwdProp; idx: Integer(ipropLandscape)),
(    Keyword: 'pard';        dflt:                  1; fPassDflt:  True; KWD: kwdProp; idx: Integer(ipropPard)),
(    Keyword: 'par';         dflt:                  1; fPassDflt:  True; KWD: kwdProp; idx: Integer(ipropPar)),
(    Keyword: '\0x0a';       dflt:                  0; fPassDflt: False; KWD: kwdChar; idx: $0A),
(    Keyword: '\0x0d';       dflt:                  0; fPassDflt: False; KWD: kwdChar; idx: $0A),
(    Keyword: 'tab';         dflt:                  0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropTab)),
(    Keyword: 'ldblquote';   dflt:                  0; fPassDflt: False; KWD: kwdChar; idx: Ord('"')),
(    Keyword: 'rdblquote';   dflt:                  0; fPassDflt: False; KWD: kwdChar; idx: Ord('"')),
(    Keyword: 'bin';         dflt:                  0; fPassDflt: False; KWD: kwdSpec; idx: Integer(ipfnBin)),
(    Keyword: '''';          dflt:                  0; fPassDflt: False; KWD: kwdSpec; idx: Integer(ipfnHex)),
(    Keyword: 'u';           dflt:                  0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropUnicode)),
(    Keyword: 'uc';          dflt:                  0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropUnicodeGroup)),
(    Keyword: 'author';      dflt:                  0; fPassDflt: False; KWD: kwdDest; idx: Integer(idestSkip)),
(    Keyword: 'buptim';      dflt:                  0; fPassDflt: False; KWD: kwdDest; idx: Integer(idestSkip)),
(    Keyword: 'page';        dflt:                  0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropPageBreak)),
(    Keyword: 'line';        dflt:                  0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropLineBreak)),

(    Keyword: 'colortbl';    dflt:                  0; fPassDflt: False; KWD: kwdDest; idx: Integer(idestColorTbl)),
(    Keyword: 'red';         dflt:                  0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropRed)),
(    Keyword: 'green';       dflt:                  0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropGreen)),
(    Keyword: 'blue';        dflt:                  0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropBlue)),

(    Keyword: 'comment';     dflt:                  0; fPassDflt: False; KWD: kwdDest; idx: Integer(idestSkip)),
(    Keyword: 'creatim';     dflt:                  0; fPassDflt: False; KWD: kwdDest; idx: Integer(idestSkip)),
(    Keyword: 'doccomm';     dflt:                  0; fPassDflt: False; KWD: kwdDest; idx: Integer(idestSkip)),
(    Keyword: 'f';           dflt:                 -1; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropFont)),
(    Keyword: 'fonttbl';     dflt:                  0; fPassDflt: False; KWD: kwdDest; idx: Integer(idestFontTbl)),
(    Keyword: 'fcharset';    dflt:                  0; fPassDflt:  True; KWD: kwdProp; idx: Integer(ipropCharSet)),
(    Keyword: 'footer';      dflt:                  0; fPassDflt: False; KWD: kwdDest; idx: Integer(idestSkip)),
(    Keyword: 'footerf';     dflt:                  0; fPassDflt: False; KWD: kwdDest; idx: Integer(idestSkip)),
(    Keyword: 'footerl';     dflt:                  0; fPassDflt: False; KWD: kwdDest; idx: Integer(idestSkip)),
(    Keyword: 'footerr';     dflt:                  0; fPassDflt: False; KWD: kwdDest; idx: Integer(idestSkip)),
(    Keyword: 'footnote';    dflt:                  0; fPassDflt: False; KWD: kwdDest; idx: Integer(idestSkip)),
(    Keyword: 'ftncn';       dflt:                  0; fPassDflt: False; KWD: kwdDest; idx: Integer(idestSkip)),
(    Keyword: 'ftnsep';      dflt:                  0; fPassDflt: False; KWD: kwdDest; idx: Integer(idestSkip)),
(    Keyword: 'ftnsepc';     dflt:                  0; fPassDflt: False; KWD: kwdDest; idx: Integer(idestSkip)),
(    Keyword: 'header';      dflt:                  0; fPassDflt: False; KWD: kwdDest; idx: Integer(idestSkip)),
(    Keyword: 'headerf';     dflt:                  0; fPassDflt: False; KWD: kwdDest; idx: Integer(idestSkip)),
(    Keyword: 'headerl';     dflt:                  0; fPassDflt: False; KWD: kwdDest; idx: Integer(idestSkip)),
(    Keyword: 'headerr';     dflt:                  0; fPassDflt: False; KWD: kwdDest; idx: Integer(idestSkip)),
(    Keyword: 'info';        dflt:                  0; fPassDflt: False; KWD: kwdDest; idx: Integer(idestSkip)),
(    Keyword: 'keywords';    dflt:                  0; fPassDflt: False; KWD: kwdDest; idx: Integer(idestSkip)),
(    Keyword: 'operator';    dflt:                  0; fPassDflt: False; KWD: kwdDest; idx: Integer(idestSkip)),
(    Keyword: 'printim';     dflt:                  0; fPassDflt: False; KWD: kwdDest; idx: Integer(idestSkip)),
(    Keyword: 'private1';    dflt:                  0; fPassDflt: False; KWD: kwdDest; idx: Integer(idestSkip)),
(    Keyword: 'revtim';      dflt:                  0; fPassDflt: False; KWD: kwdDest; idx: Integer(idestSkip)),
(    Keyword: 'rxe';         dflt:                  0; fPassDflt: False; KWD: kwdDest; idx: Integer(idestSkip)),
(    Keyword: 'subject';     dflt:                  0; fPassDflt: False; KWD: kwdDest; idx: Integer(idestSkip)),
(    Keyword: 'tc';          dflt:                  0; fPassDflt: False; KWD: kwdDest; idx: Integer(idestSkip)),
(    Keyword: 'title';       dflt:                  0; fPassDflt: False; KWD: kwdDest; idx: Integer(idestSkip)),
(    Keyword: 'txe';         dflt:                  0; fPassDflt: False; KWD: kwdDest; idx: Integer(idestSkip)),
(    Keyword: 'xe';          dflt:                  0; fPassDflt: False; KWD: kwdDest; idx: Integer(idestSkip)),
// Table =======================================================================
(    Keyword: 'intbl';       dflt:                  0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropIntbl)),
(    Keyword: 'itap';        dflt:                  1; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropITap)),
(    Keyword: 'trowd';       dflt:                  0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropTrowD)),
(    Keyword: 'cell';        dflt:                  0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropCell)),
(    Keyword: 'nestcell';    dflt:                  0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropNestCell)),
(    Keyword: 'row';         dflt:                  0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropRow)),
(    Keyword: 'nestrow';     dflt:                  0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropNestRow)),
(    Keyword: 'lastrow';     dflt:                  0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropLastRow)),
(    Keyword: 'nesttableprops'; dflt:               0; fPassDflt: False; KWD: kwdDest; idx: Integer(idestNestTableProps)),
(    Keyword: 'nonesttables'; dflt:                 0; fPassDflt: False; KWD: kwdDest; idx: Integer(idestNoNestTables)),
(    Keyword: 'cellx';       dflt:                  0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropCellX)),
(    Keyword: 'trbrdrh';     dflt:                  0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropTblBorderIH)),
(    Keyword: 'trbrdrv';     dflt:                  0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropTblBorderIV)),
(    Keyword: 'trbrdrl';     dflt:                  0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropTblBorderL)),
(    Keyword: 'trbrdrt';     dflt:                  0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropTblBorderT)),
(    Keyword: 'trbrdrr';     dflt:                  0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropTblBorderR)),
(    Keyword: 'trbrdrb';     dflt:                  0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropTblBorderB)),

(    Keyword: 'trpaddl';     dflt:                  0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropTblCellMargL)),
(    Keyword: 'trpaddt';     dflt:                  0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropTblCellMargT)),
(    Keyword: 'trpaddr';     dflt:                  0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropTblCellMargR)),
(    Keyword: 'trpaddb';     dflt:                  0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropTblCellMargB)),

(    Keyword: 'clbrdrl';     dflt:                  0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropCellBorderL)),
(    Keyword: 'clbrdrt';     dflt:                  0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropCellBorderT)),
(    Keyword: 'clbrdrr';     dflt:                  0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropCellBorderR)),
(    Keyword: 'clbrdrb';     dflt:                  0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropCellBorderB)),
(    Keyword: 'clcbpat';     dflt:                  0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropCellColor)),
(    Keyword: 'clshdrawnil'; dflt:                  0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropCellNoShading)),

(    Keyword: 'clvertalt';   dflt:      Ord(atcavTop); fPassDflt: True;  KWD: kwdProp; idx: Integer(ipropCellVAlign)),
(    Keyword: 'clvertalc';   dflt:   Ord(atcavCenter); fPassDflt: True;  KWD: kwdProp; idx: Integer(ipropCellVAlign)),
(    Keyword: 'clvertalb';   dflt:   Ord(atcavBottom); fPassDflt: True;  KWD: kwdProp; idx: Integer(ipropCellVAlign)),

(    Keyword: 'trrh';        dflt:                  0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropRowHeight)),

// Tab =========================================================================
(    Keyword: 'tx';          dflt:                  0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropTabPos)),
(    Keyword: 'tqr';         dflt:                  0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropTabR)),
(    Keyword: 'tqc';         dflt:                  0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropTabC)),
(    Keyword: 'tqdec';       dflt:                  0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropTabDec)),
(    Keyword: 'tldot';       dflt:                  0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropTabLDot)),
(    Keyword: 'tlhyph';      dflt:                  0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropTabLHyph)),
(    Keyword: 'tlul';        dflt:                  0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropTabLULine)),
(    Keyword: 'tlth';        dflt:                  0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropTabLThick)),

// Picture =====================================================================
(    Keyword: 'pict';        dflt:                  0; fPassDflt: False; KWD: kwdDest; idx: Integer(idestPict)),
(    Keyword: 'shppict';     dflt:                  0; fPassDflt: False; KWD: kwdDest; idx: Integer(idestShpPict)),
(    Keyword: 'nonshppict';  dflt:                  0; fPassDflt: False; KWD: kwdDest; idx: Integer(idestSkip)),

(    Keyword: 'emfblip';     dflt:    Integer(aptEMF); fPassDflt:  True; KWD: kwdProp; idx: Integer(ipropPicture)),
(    Keyword: 'pngblip';     dflt:    Integer(aptPNG); fPassDflt:  True; KWD: kwdProp; idx: Integer(ipropPicture)),
(    Keyword: 'jpegblip';    dflt:    Integer(aptJPG); fPassDflt:  True; KWD: kwdProp; idx: Integer(ipropPicture)),
(    Keyword: 'macpict';     dflt:    Integer(aptPIC); fPassDflt:  True; KWD: kwdProp; idx: Integer(ipropPicture)),
(    Keyword: 'pmmetafile';  dflt:    Integer(aptWMF); fPassDflt:  True; KWD: kwdProp; idx: Integer(ipropPicture)),
(    Keyword: 'wmetafile';   dflt:    Integer(aptWMF); fPassDflt:  True; KWD: kwdProp; idx: Integer(ipropPicture)),
(    Keyword: 'dibitmap';    dflt:    Integer(aptBMP); fPassDflt:  True; KWD: kwdProp; idx: Integer(ipropPicture)),
(    Keyword: 'wbitmap';     dflt:    Integer(aptRAW); fPassDflt:  True; KWD: kwdProp; idx: Integer(ipropPicture)),

(    Keyword: 'picw';        dflt:                  0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropPictWidth)),
(    Keyword: 'pich';        dflt:                  0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropPictHeigth)),
(    Keyword: 'picwgoal';    dflt:                  0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropPictGoalWidth)),
(    Keyword: 'pichgoal';    dflt:                  0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropPictGoalHeight)),
(    Keyword: 'picscalex';   dflt:                  1; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropPictScaleX)),
(    Keyword: 'picscaley';   dflt:                  1; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropPictScaleY)),
(    Keyword: 'piccropt';    dflt:                  0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropPictCropT)),
(    Keyword: 'piccropb';    dflt:                  0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropPictCropB)),
(    Keyword: 'piccropl';    dflt:                  0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropPictCropL)),
(    Keyword: 'piccropr';    dflt:                  0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropPictCropR)),
//(    Keyword: 'blipuid';     dflt:                  0; fPassDflt: False; KWD: kwdDest; idx: Integer(idestBlipUid)),
(    Keyword: 'bliptag';     dflt:                  0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropBlipTag)),

(    Keyword: 'bkmkcolf';    dflt:                 -1; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropBkmkColFirst)),
(    Keyword: 'bkmkcoll';    dflt:                 -1; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropBkmkColLast)),

(    Keyword: 's';           dflt:                  0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropParaStyle)),
(    Keyword: 'cs';          dflt:                  0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropCharRunStyle)),
(    Keyword: 'ts';          dflt:                  0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropTableStyle)),
(    Keyword: 'ds';          dflt:                  0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropSectStyle)),

// Other =======================================================================
(    Keyword: 'datafield';   dflt:                  0; fPassDflt: False; KWD: kwdProp; idx: Integer(ipropDataField)),

//// Destination =================================================================
//(    Keyword: 'bkmkstart';   dflt:                  0; fPassDflt: False; KWD: kwdDest; idx: Integer(idestBkmkStart)),
//(    Keyword: 'bkmkend';     dflt:                  0; fPassDflt: False; KWD: kwdDest; idx: Integer(idestBkmkEnd)),

(    Keyword: 'field';       dflt:                  0; fPassDflt: False; KWD: kwdDest; idx: Integer(idestField)),
(    Keyword: 'fldinst';     dflt:                  0; fPassDflt: False; KWD: kwdDest; idx: Integer(idestFieldInstr)),
(    Keyword: 'fldrslt';     dflt:                  0; fPassDflt: False; KWD: kwdDest; idx: Integer(idestFieldResult)),

(    Keyword: 'stylesheet';  dflt:                  0; fPassDflt: False; KWD: kwdDest; idx: Integer(idestStylesheet)),

(    Keyword: 'generator';   dflt:                  0; fPassDflt: False; KWD: kwdDest; idx: Integer(idestGenerator)),

// Comments ====================================================================
(    Keyword: 'annotation';  dflt:                  0; fPassDflt: False; KWD: kwdDest; idx: Integer(idestAnnotation)),
(    Keyword: 'atnid';       dflt:                  0; fPassDflt: False; KWD: kwdDest; idx: Integer(idestAnnInitials)),
(    Keyword: 'atrfstart';   dflt:                  0; fPassDflt: False; KWD: kwdDest; idx: Integer(idestAnnStart)),
(    Keyword: 'atrfend';     dflt:                  0; fPassDflt: False; KWD: kwdDest; idx: Integer(idestAnnEnd)),
(    Keyword: 'atnref';      dflt:                  0; fPassDflt: False; KWD: kwdDest; idx: Integer(idestAnnRef)),
(    Keyword: 'atntime';     dflt:                  0; fPassDflt: False; KWD: kwdDest; idx: Integer(idestAnnTime)),
(    Keyword: 'atndate';     dflt:                  0; fPassDflt: False; KWD: kwdDest; idx: Integer(idestAnnDate)),
(    Keyword: 'atnauthor';   dflt:                  0; fPassDflt: False; KWD: kwdDest; idx: Integer(idestAnnAuthor)),
//(    Keyword: 'atnparent';   dflt:                  0; fPassDflt: False; KWD: kwdDest; idx: Integer(idestAnnParent)),
//(    Keyword: 'atnicn';      dflt:                  0; fPassDflt: False; KWD: kwdDest; idx: Integer(idestAnnIcon)),

// Character ===================================================================
(    Keyword: '*';           dflt:                  0; fPassDflt: False; KWD: kwdSpec; idx: Integer(ipfnSkipDest)),
(    Keyword: '{';           dflt:                  0; fPassDflt: False; KWD: kwdChar; idx: Ord('{')),
(    Keyword: '}';           dflt:                  0; fPassDflt: False; KWD: kwdChar; idx: Ord('}')),
(    Keyword: '\';           dflt:                  0; fPassDflt: False; KWD: kwdChar; idx: Ord('\')));

const isymMax = sizeof(rgsymRtf) / sizeof(TSYM);

{ TRTFReader }

procedure TAXWReadRTF.AddTable(AParas: TAXWLogParas);
var
  Paras: TAXWLogParas;
begin
  CheckCell;

  FTableStack.Push(FCurrCell);

  if FCurrTable <> Nil then
    Paras := FCurrTable.LastCell.Paras
  else
    Paras := AParas;


  FCurrTable := Paras.AddTable;
  FCurrTable.TAPX.SetDefaultBorders;
  FCurrRow := FCurrTable.Add;
  FCurrCell := FCurrRow.Add;
end;

procedure TAXWReadRTF.AddText;
var
  CR: TAXWCharRun;
begin
  if FText <> '' then begin
    case FCharRunType of
      acrtText     : CR := GetPara.Runs.Add;
      acrtHyperlink: begin
        if FCurrHyperlink <> Nil then
          CR := GetPara.Runs.AddHyperlink(Nil,FCurrHyperlink)
        else
          CR := GetPara.Runs.Add;
      end;
//      acrtSimpleField: begin
//        if FCurrSimpField <> Nil then
//          CR := GetPara.Runs.AddSimpleField(Nil,FCurrSimpField,True)
//        else
//          CR := GetPara.Runs.Add;
//      end
      else           CR := GetPara.Runs.Add;
    end;

    if FCHP.CurrBorder <> Nil then begin
      FCHP.CurrBorder.Style := stbSingle;
      FCHP.AddBorder.Assign(FCHP.CurrBorder);
    end;

    if FExText <> '' then begin
      FText := FExText + FText;
      FExText := '';
    end;

    CR.Text := CR.Text + FText;
    FText := '';
    if FCHP.Count > 0 then begin
      CR.AddCHPX;
      CR.CHPX.Assign(FCHP);
    end;
  end;
end;

procedure TAXWReadRTF.CheckCell;
begin
  if FCurrTable <> Nil then begin
    if FCurrRow = Nil then begin
      FCurrRow := FCurrTable.Add;
      FCurrCell := Nil;
    end;
    if FCurrCell = Nil then
      FCurrCell := FCurrRow.Add;
  end;
end;

procedure TAXWReadRTF.CheckTable;
var
  Paras: TAXWLogParas;
begin
  if FCurrTable <> Nil then
    Paras := FCurrTable.LastCell.Paras
  else
    Paras := FDoc.Paras;

  while FPAP.TblNesting > FTableStack.Count do begin
    AddTable(Paras);
    Paras := FCurrCell.Paras;
  end;
  while FPAP.TblNesting < FTableStack.Count do
    FTableStack.Pop(FCurrTable,FCurrRow,FCurrCell);

  if FCurrTable = Nil then
    AddTable(Paras);

  CheckCell;
end;

constructor TAXWReadRTF.Create(AEditor: TAXWLogDocEditor);
var
  i: integer;
begin
  FDoc := AEditor;

  FColorTable := TRtfDestColorTable.Create(Self,idestColorTbl);
  FFontTable := TRtfDestFontTable.Create(Self,idestFontTbl);

  FCHP := TRtfCHP.Create(FDoc.MasterCHP,FColorTable,FFontTable);
  FPAP := TRtfPAP.Create(FDoc.MasterPAP,FColorTable);

  FStack := TRTFStack.Create(Self);
  FTableStack := TRtfTableStack.Create(FColorTable);

  FBookmarks := THashedStringList.Create;
  FComments := THashedStringList.Create;
  FStyles := THashedStringList.Create;

  FSymbols := THashedStringList.Create;
  for i := 0 to High(rgsymRtf) do
    FSymbols.Add(AxUCString(rgsymRtf[i].Keyword));

  FCurrStream := TMemoryStream.Create;

  FTabAlignment := atsaLeft;

  FInheritPara := True;
end;

function TAXWReadRTF.GetPara: TAXWLogPara;
var
  Paras: TAXWLogParas;
begin
  if FPAP.InTable then begin
    CheckTable;
    Paras := FCurrCell.Paras;
  end
  else begin
    Paras := FDoc.Paras;
    FTableStack.Clear;
    FCurrTable := Nil;
    FCurrRow := Nil;
    FCurrCell := Nil;
  end;

  if (FCurrPara = Nil) or (Paras.Count <= 0) or (Paras._Last.Type_ <> alptPara) then begin

    if FInheritPara then
      Result := Paras.AppendPara
    else
      Result := Paras._Add;
    FInheritPara := True;

    if FPAP.Count > 0 then begin
      Result.AddPAPX;
      Result.PAPX.Assign(FPAP);
    end;
  end
  else
    Result := Paras.Last;

  FCurrPara := Result;
end;

destructor TAXWReadRTF.Destroy;
begin
  FBookmarks.Free;
  FComments.Free;
  FStyles.Free;
  FSymbols.Free;

  FStack.Free;
  FTableStack.Free;

  FCHP.Free;
  FPAP.Free;

  FColorTable.Free;
  FFontTable.Free;

  FCurrStream.Free;

  inherited Destroy;
end;

procedure TAXWReadRTF.LoadFromFile(const Filename: AxUCString);
var
  Stream: TFileStream;
begin
  Stream := TFileStream.Create(Filename,fmOpenRead or fmShareDenyNone);
  try
    LoadFromStream(Stream);
  finally
    Stream.Free;
  end;
end;

procedure TAXWReadRTF.LoadFromStream(Stream: TStream);
var
  Res: integer;
begin
  FDoc.Clear;

  FText := '';
  Res := ecRtfParse(Stream);
  if Res <> ecOk then
    raise XLSRWException.CreateFmt('Error while reading RTF document (%d)',[Res]);
end;

function TAXWReadRTF.ecRtfParse(fp: TStream): integer;
var
  val    : byte;
  ch     : AnsiChar;
  cNibble: integer;
  b      : integer;
begin
  cNibble := 2;
  b := 0;
  while fp.Read(val,1) = 1 do begin
    ch := AnsiChar(val);
    if cGroup < 0 then begin
      Result := ecStackUnderflow;
      Exit;
    end;
    if FRIS = risBin then begin                      // if we're parsing binary data, handle it directly
      Result := ecParseChar(ch);
      if Result <> Integer(ecOK) then
          Exit;
    end
    else begin
      case ch of
        '{': begin
           AddText;

           ecPushRtfState;

           if FStack.CurrDest <> Nil then
             FStack.CurrDest.BeginGroup;
        end;
        '}': begin
           AddText;

           Result := ecPopRtfState;

           if Result <> ecOK then
             Exit;
          end;
        '\': begin
//           AddText;

           Result := ecParseRtfKeyword(fp);
           if Result <> ecOK then
             Exit;
          end;
          ';': begin
            if FStack.CurrDest <> Nil then
              FStack.CurrDest.EndComposedValue;
          end;
        Char($0D),Char($0A): ;          // cr and lf are noise characters...
        else begin
          if FRIS = risNorm then begin

            Result := ecParseChar(ch);
            if Result <> ecOK then
              Exit;
          end
          else begin
                         // parsing hex data
            if FRIS <> risHex then begin
              Result := ecAssertion;
              Exit;
            end;
            b := b shl 4;
            if CharInSet(AXUCChar(ch),['0'..'9']) then
              Inc(b,Integer(ch) - Ord('0'))
            else begin
              if CharInSet(AXUCChar(ch),['a'..'f']) then
                Inc(b,Integer(ch) - Ord('a') + 10)
              else if CharInSet(AXUCChar(ch),['A'..'F']) then
                Inc(b,Integer(ch) - Ord('A') + 10)
              else begin
                Result := ecInvalidHex;
                Exit;
              end;
            end;
            Dec(cNibble);
            if cNibble <= 0 then begin
              if FLastUnicodeChar <> WideChar(b) then
                Result := ecParseChar(AnsiChar(b))
              else
                Result := ecOk;
              if Result <> ecOK then
                  Exit;
              cNibble := 2;
              b := 0;
              FRIS := risNorm;
              FLastUnicodeChar := #0;
            end;
          end;                   // end else (ris != risNorm)
        end;
      end;
    end;           // else (ris != risBin)
  end;               // while
  if cGroup < 0 then
      Result := ecStackUnderflow
  else if cGroup > 0 then
      Result := ecUnmatchedBrace
  else
  Result := ecOK;
end;

procedure TAXWReadRTF.ecPushRtfState;
begin
  FStack.Push(FRDS,FRIS,FCHP,FPAP,FSEP,FDOP);
  FRIS := risNorm;
  Inc(cGroup);
end;

function TAXWReadRTF.ecPopRtfState: integer;
var
  Dest: TRtfDestination;
begin
  if FStack.Count <= 0 then begin
    Result := ecStackUnderflow;
    Exit;
  end;

  if FRDS <> FStack[FStack.Count - 1].RDS then begin
    Result := ecEndGroupAction;
    if Result <> ecOK then
      Exit;
  end;

  if FStack.CurrDest <> Nil then
    FStack.CurrDest.EndGroup;

  FStack.Pop(FRDS,FRIS,FCHP,FPAP,FSEP,FDOP,Dest);
  if Dest <> Nil then begin
    Dest.EndDest;
    if not Dest.Persistent then
      Dest.Free;
  end;

  Dec(cGroup);
  Result := ecOK;
end;

function TAXWReadRTF.ecParseRtfKeyword(fp: TStream): integer;
var
  ch: AnsiChar;
  fParam: boolean;
  fNeg: boolean;
  param: integer;
  pch: AnsiString;
  szKeyword: AnsiString;
  szParameter: AnsiString;
begin
  fParam := False;
  fNeg := False;
  param := 0;

  if fp.Read(ch,1) <> 1 then begin
    Result := ecEndOfFile;
    Exit;
  end;
  // !isalpha(ch)
  if not CharInSet(AXUCChar(ch),['a'..'z','A'..'Z']) then begin
    szKeyword := ch;
    Result := ecTranslateKeyword(szKeyword, 0, fParam, fp);
    Exit;
  end;
  pch := ch;
  // isalpha(ch)
  while (fp.Read(ch,1) = 1) and (ch in ['a'..'z','A'..'Z']) do
    pch := pch + ch;
  szKeyword := pch;
  if ch = '-' then begin
    fNeg := True;
    if fp.Read(ch,1) <> 1 then begin
      Result := ecEndOfFile;
      Exit;
    end;
  end;
  if AnsiChar(ch) in ['0'..'9'] then begin
    fParam := True;         // a digit after the control means we have a parameter
    pch := ch;
    while (fp.Read(ch,1) = 1) and (AnsiChar(ch) in ['0'..'9']) do
      pch := pch + ch;
    szParameter := pch;
    param := StrToInt(AxUCString(szParameter));
    if fNeg then
      param := -param;
    lParam := StrToInt(AxUCString(szParameter));
    if fNeg then
      lParam := -lParam;
  end;
  if ch <> ' ' then
    fp.Seek(-1,soFromCurrent);

  Result := ecTranslateKeyword(szKeyword, param, fParam, fp);
end;

function TAXWReadRTF.ecParseChar(ch: AnsiChar): integer;
begin
  Result := ecOK;

  case FRDS of
    rdsSkip: ;        // Toss this character.
    rdsNorm: begin
      if (FStack.CurrDest <> Nil) and not FStack.CurrDest.HasDocText then
        FStack.CurrDest.ParseChar(FRIS,ch)
      else
        ecPrintChar(AxUCChar(ch));
    end;
    else begin
    // handle other destinations....
    end;
  end;
end;

procedure TAXWReadRTF.ecPrintChar(ch: AxUCChar);
begin
  FText := FText + ch;
end;

{ TRTFStack }

constructor TRTFStack.Create(AOwner: TAXWReadRTF);
begin
  inherited Create;

  FOwner := AOwner;
end;

function TRTFStack.CurrDest: TRtfDestination;
begin
  Result := FCurrDest;
//  Result := FindDest;
end;

function TRTFStack.FindDest: TRtfDestination;
var
  i: integer;
begin
  for i := Count - 1 downto 0 do begin
    if Items[i].Dest <> Nil then begin
      Result := Items[i].Dest;
      Exit;
    end;
  end;
  Result := Nil;
end;

function TRTFStack.GetItems(Index: integer): TRTFStackItem;
begin
  Result := TRTFStackItem(inherited Items[Index]);
end;

function TRTFStack.Peek: TRTFStackItem;
begin
  Result := TRTFStackItem(Last);
end;

function TRTFStack.PeekDestTypeL2: TRtfIDEST;
begin
  if (Count > 1) and (PeekL2.Dest <> Nil) then
    Result := PeekL2.Dest.Type_
  else
    Result := idestNil;
end;

function TRTFStack.PeekL2: TRTFStackItem;
begin
  Result := Items[Count - 2];
end;

procedure TRTFStack.Pop(var RDS: TRtfRDS; var RIS: TRtfRIS; CHP: TAXWCHPX; PAP: TRtfPAP; var SEP: TRtfSEP; var DOP: TRtfDOP; var ADest: TRtfDestination);
var
  i: integer;
begin
  i := Count - 1;
  RDS := Items[i].RDS;
  RIS := Items[i].RIS;
  CHP.Assign(Items[i].CHP);
  PAP.Assign(Items[i].PAP);
  SEP := Items[i].SEP;
  DOP := Items[i].DOP;
  ADest := Items[i].Dest;
  Delete(i);

  FCurrDest := FindDest;
end;

procedure TRTFStack.Push(RDS: TRtfRDS; RIS: TRtfRIS; CHP: TRtfCHP; PAP: TRtfPAP; SEP: TRtfSEP; DOP: TRtfDOP);
begin
  inherited Add(TRTFStackItem.Create(Self,RDS,RIS,CHP,PAP,SEP,DOP));
end;

procedure TRTFStack.SetTOSDest(ADest: TRtfDestination);
begin
  Peek.Dest := ADest;
  FCurrDest := ADest;
end;

{ TRTFStackItem }

constructor TRTFStackItem.Create(AOwner: TRTFStack; RDS: TRtfRDS; RIS: TRtfRIS; CHP: TRtfCHP; PAP: TRtfPAP; SEP: TRtfSEP; DOP: TRtfDOP);
begin
  FOwner := AOwner;

  FRDS := RDS;
  FRIS := RIS;
  FCHP := TRtfCHP.Create(FOwner.FOwner.FDoc.MasterCHP,CHP.FColorTable,CHP.FFontTable);
  FCHP.Assign(CHP);
  FPAP := TRtfPAP.Create(FOwner.FOwner.FDoc.MasterPAP,PAP.FColorTable);
  FPAP.Assign(PAP);
  FSEP := SEP;
  FDOP := DOP;
  FDest := Nil;
end;

destructor TRTFStackItem.Destroy;
begin
  FCHP.Free;
  FPAP.Free;

  inherited;
end;

function TAXWReadRTF.ecApplyPropChange(iprop: TRtfIPROP; val: integer): integer;
var
  i    : integer;
//  Style: TAXWStyle;
  Para: TAXWLogPara;
begin
  Result := ecOK;

  if (FStack.CurrDest <> Nil) and not FStack.CurrDest.HasDocText then begin
    FStack.CurrDest.ApplyKeyword(iprop,val);
    Exit;
  end;

  AddText;

  if Frds <> rdsSkip then begin
    if FPAP.ApplyKeyword(iprop,val) then
      Exit;
    if FCHP.ApplyKeyword(iprop,val) then
      Exit;
    case iprop of
// CHP =========================================================================
      ipropCharRunStyle: begin
        i := FStyles.IndexOf(IntToStr(Val));
        if (i >= 0) and (TAXWStyle(FStyles.Objects[i]).Type_ = astCharRun) then
          FCHP.Style := TAXWCharStyle(FStyles.Objects[i]);
      end;
      ipropChBrdr     : FCurrBorder := FCHP.AddCurrBorder;

// Tabs ========================================================================
      ipropTabPos   : begin
        Para := GetPara;
        Para.AddTabStops;
        Para.TabStops.Add(Integer(Val) / 20);
        if FTabLeader <> atslNone then
          Para.TabStops.Last.Leader := FTabLeader;
        if FTabAlignment <> atsaLeft then
          Para.TabStops.Last.Alignment := FTabAlignment;

        FTabAlignment := atsaLeft;
        FTabLeader := atslNone;
      end;
      ipropTabC      : FTabAlignment := atsaCenter;
      ipropTabR      : FTabAlignment := atsaRight;
      ipropTabDec    : FTabAlignment := atsaDecimal;
      ipropTabLDot   : FTabLeader := atslDot;
      ipropTabLHyph  : FTabLeader := atslHyphen;
      ipropTabLULine : FTabLeader := atslUnderscore;
      ipropTabLThick : FTabLeader := atslHeavy;

      ipropBrdrt,
      ipropBrdrl,
      ipropBrdrr,
      ipropBrdrb      : FCurrBorder := FPAP.DoBorders(iprop);

      ipropCellBorderL: FCurrBorder := FTableStack.CurrItem.Borders.AddLeft;
      ipropCellBorderT: FCurrBorder := FTableStack.CurrItem.Borders.AddTop;
      ipropCellBorderR: FCurrBorder := FTableStack.CurrItem.Borders.AddRight;
      ipropCellBorderB: FCurrBorder := FTableStack.CurrItem.Borders.AddBottom;

      ipropBrdrw      : if FCurrBorder <> Nil then FCurrBorder.Width := Val / 20;
      ipropBrdrcf     : if FCurrBorder <> Nil then FCurrBorder.Color := FColorTable.Color(Val,FCurrBorder.Color);
      ipropBrsp       : if FCurrBorder <> Nil then FCurrBorder.Space := Val / 20;
      ipropBorderStyle: if FCurrBorder <> Nil then FCurrBorder.Style := TST_Border(Val);

// SEP =========================================================================
// Correct SEP ?
      ipropXaLeft     : FDoc.SEP.MargLeft := Val / 20;
      ipropXaRight    : FDoc.SEP.MargRight := Val / 20;

// Special =====================================================================
      ipropPar      : begin
//        if FCurrPara = Nil then
//          // Add empty para.
//          GetPara;
//
//        FCurrPara := Nil;
        // Add line break to cell.
        FText := FText + #13;
      end;
      ipropPard     : begin
        FPAP.Clear;
        FDoc.Paras.AutoNumbering := anstNone;
        FInheritPara := False;
//        FCurrPara := Nil;
      end;
      ipropPlain    : FCHP.Clear;
      ipropSectd    : FillChar(FSEP,SizeOf(TRtfSEP),#0);

// Char  =======================================================================
      ipropTab      : GetPara.Runs.AddTab;
      ipropLineBreak: GetPara.Runs.AddBreak(acrbtLineBreak);
      ipropPageBreak: GetPara.Runs.AddBreak(acrbtPageBreak);
      ipropUnicode  : begin
        if FRDS = rdsNorm then begin
          FLastUnicodeChar := AXUCChar(Val);
          ecPrintChar(FLastUnicodeChar);
        end;
      end;
      ipropUnicodeGroup : ;

// Table  ======================================================================
      ipropIntbl    : FPAP.InTable := True;
      ipropNestCell,
      ipropCell     : begin
        CheckTable;

//        AddText;

        CheckCell;
        FCurrCell := Nil;
      end;
      ipropCellX    : FTableStack.CurrItem.CellX.Add(Val);
      ipropCellColor: FTableStack.CurrItem.CellColor.Add(Val);
      ipropCellNoShading: FTableStack.CurrItem.CellColor.Add(-1);
      ipropRow,
      ipropNestRow  : begin
        if FCurrRow <> Nil then
          FTableStack.ApplyOnRow(FCurrRow);

//        AddText;

        FCurrCell := Nil;
        FCurrRow := Nil;
      end;
      ipropTRowD    : begin
        FTableStack.ClearRow;
        FCurrBorder := Nil;
      end;
      ipropITap     : begin
        FPAP.TblNesting := Val;
        FPAP.InTable := Val > 0;
      end;
      ipropRowHeight: FTableStack.CurrItem.RowHeight := Val / 20;

      ipropTblCellMargL: if FCurrTable <> Nil then FCurrTable.TAPX.CellMargLeft := Val / 20;
      ipropTblCellMargT: if FCurrTable <> Nil then FCurrTable.TAPX.CellMargTop := Val / 20;
      ipropTblCellMargR: if FCurrTable <> Nil then FCurrTable.TAPX.CellMargRight := Val / 20;
      ipropTblCellMargB: if FCurrTable <> Nil then FCurrTable.TAPX.CellMargBottom := Val / 20;

      ipropCellVAlign  : if FCurrCell <> Nil then FCurrCell.AlignVert := TAXWTableCellAlignVert(Val);
// Bullets and numbering =======================================================
      ipropLS          : FDoc.Paras.AutoNumbering := anstBullet;
    end;
  end;
end;

function TAXWReadRTF.ecTranslateKeyword(szKeyword: AnsiString; param: integer; fParam: boolean; fp: TStream): integer;
var
  isym: integer;
begin
  Result := ecBadTable;

  isym := FSymbols.IndexOf(AnsiLowercase(AxUCString(szKeyword)));

  if isym < 0 then begin           // control word not found
    if fSkipDestIfUnk then         // if this is a new destination
      FRDS := rdsSkip;          // skip the destination                                    // else just discard it
    fSkipDestIfUnk := False;
    Result := ecOK;
    Exit;
  end;

  // found it!  use kwd and idx to determine what to do with it.

  fSkipDestIfUnk := False;
  case rgsymRtf[isym].kwd of
    kwdProp: begin
      if rgsymRtf[isym].fPassDflt and not fParam then
        param := rgsymRtf[isym].dflt;
      Result := ecApplyPropChange(TRtfIPROP(rgsymRtf[isym].idx), param);
    end;
    kwdChar: begin
      Result := ecParseChar(AnsiChar(rgsymRtf[isym].idx));
    end;
    kwdDest: begin
      Result := ecChangeDest(TRtfIDEST(rgsymRtf[isym].idx),Param);
    end;
    kwdSpec:
      Result := ecParseSpecialKeyword(TRtfIPFN(rgsymRtf[isym].idx));
  end;
end;

function TAXWReadRTF.ecChangeDest(idest: TRtfIDEST; Param: integer): integer;
var
  Dest: TRtfDestination;
begin
  if FRDS = rdsSkip then
    Result := ecOK                // don't do anything
  else begin
    Dest := Nil;

    case idest of
      idestNil         : ;
      idestShpPict     : ;
      idestPict        : Dest := TRtfDestPict.Create(Self,idest);
      idestFontTbl     : Dest := FFontTable;
      idestColorTbl    : Dest := FColorTable;
//      idestBkmkStart   : Dest := TRtfDestBookmark.Create(Self,idest,FBookmarks);
//      idestBkmkEnd     : Dest := TRtfDestBookmark.Create(Self,idest,FBookmarks);
      idestStylesheet  : Dest := TRtfDestStylesheet.Create(Self,idest);

      idestField       : Dest := TRtfDestField.Create(Self,idest);
      idestFieldInstr  : begin
        if FStack.PeekDestTypeL2 = idestField then
          FStack.PeekL2.Dest.ChangeType(idest)
        else
          FRDS := rdsSkip;
      end;

      idestFieldResult : begin
        if FStack.PeekDestTypeL2 = idestFieldInstr then
          FStack.PeekL2.Dest.ChangeType(idest)
        else
          FRDS := rdsSkip;
      end;

      idestGenerator   : Dest := TRtfDestStringParam.Create(Self,idestGenerator);

      idestBullet      : Dest := TRftDestListText.Create(Self,idest);

      idestAnnotation  : Dest := TRtfDestComment.Create(Self,idest,FComments);
      idestAnnStart    : Dest := TRtfDestComment.Create(Self,idest,FComments);
      idestAnnEnd      : Dest := TRtfDestComment.Create(Self,idest,FComments);
      idestAnnRef      : begin
       if FStack.PeekDestTypeL2 = idestAnnotation then
         Dest := TRtfDestStringParam.Create(Self,idest,FStack.PeekL2.Dest);
      end;
      idestAnnTime     : Dest := TRtfDestStringParam.Create(Self,idest);
      idestAnnDate     : Dest := TRtfDestStringParam.Create(Self,idest);
      idestAnnAuthor   : Dest := TRtfDestStringParam.Create(Self,idest);
      idestAnnInitials : Dest := TRtfDestStringParam.Create(Self,idest);
//      idestAnnParent   : Dest := TRtfDestStringParam.Create(Self,idest);
//      idestAnnIcon     : Dest := TRtfDestStringParam.Create(Self,idest);

      idestNestTableProps: ;

      else
        FRDS := rdsSkip;              // when in doubt, skip it...
    end;

    if Dest <> Nil then begin
      FStack.SetTOSDest(Dest);
      Dest.BeginDest;
    end;

    Result := ecOK;
  end;
end;

function TAXWReadRTF.ecEndGroupAction: integer;
begin
  case FRDS of
    rdsNorm     : ;
    rdsSkip     : ;
  end;
  Result := ecOK;
end;

function TAXWReadRTF.ecParseSpecialKeyword(ipfn: TRtfIPFN): integer;
begin
  if (FRDS = rdsSkip) and (ipfn <> ipfnBin) then  // if we're skipping, and it's not
    Result := ecOK                        // the \bin keyword, ignore it.
  else begin
    case ipfn of
      ipfnBin: begin
        FRIS := risBin;
        cbBin := lParam;
      end;
      ipfnSkipDest:
        fSkipDestIfUnk := True;
      ipfnHex:
        FRIS := risHex;
      else begin
        Result := ecBadTable;
        Exit;
      end;
    end;
    Result := ecOK;
  end;
end;

{ TRtfPAP }

function TRtfPAP.ApplyKeyword(iprop: TRtfIPROP; val: integer): boolean;
begin
  Result := True;
  case iprop of
    ipropJust     : Alignment := TAXWPapTextAlign(Val);

    ipropLeftInd  : IndentLeft := Integer(Val) / 20;
    ipropRightInd : IndentRight := Integer(Val) / 20;
    ipropFirstInd : begin
      if Val > 0 then
        IndentFirstLine := Integer(Val) / 20
      else
        IndentHanging := -Integer(Val) / 20;
    end;

    ipropParaColor  : Color := FColorTable.Color(Val,Color);

    ipropSpaceBefore: SpaceBefore := Val / 20;
    ipropSpaceAfter : SpaceAfter := Val / 20;
    ipropLineSpacing: LineSpacing :=  Val / 240;
//      ipropRTL        : ;
    else              Result := False;
  end;
end;

procedure TRtfPAP.Assign(APAP: TRtfPAP);
begin
  inherited Assign(APAP);

  FCurrBorder := APAP.FCurrBorder;
  FInTable := APAP.FInTable;
  FTblNesting := APAP.FTblNesting;
end;

procedure TRtfPAP.Clear;
begin
  inherited;

  FCurrBorder := Nil;
  FTblNesting := 1;
  FInTable := False;
end;

constructor TRtfPAP.Create(APAP: TAXWPAP; AColorTable: TRtfDestColorTable);
begin
  inherited Create(APAP);

  FColorTable := AColorTable;
  FTblNesting := 1;
end;

function  TRtfPAP.DoBorders(ABorderDest: TRtfIPROP): TAXWDocPropBorder;
begin
  case ABorderDest of
    ipropBrdrt: FCurrBorder := AddBorderTop;
    ipropBrdrl: FCurrBorder := AddBorderLeft;
    ipropBrdrr: FCurrBorder := AddBorderRight;
    ipropBrdrb: FCurrBorder := AddBorderBottom;
  end;
  Result := FCurrBorder;
end;

{ TRtfCHP }

function TRtfCHP.AddCurrBorder: TAXWDocPropBorder;
begin
  if FCurrBorder = Nil then
    FCurrBorder := TAXWDocPropBorder.Create;
  Result := FCurrBorder;
end;

function TRtfCHP.ApplyKeyword(iprop: TRtfIPROP; val: integer): boolean;
begin
  Result := True;

  case iprop of
    ipropFont          : FontName := FFontTable[val].Name;
    ipropBold          : Bold := Val = 1;
    ipropItalic        : Italic := Val = 1;
    ipropFontSize      : Size := Val / 2;
    ipropFontColor     : Color := FColorTable.Color(Val,Master.Color);
    ipropUnderline     : Underline := axcuSingle;
    ipropUnderlineNone : Underline := axcuNone;
    ipropUnderlineColor: UnderlineColor := FColorTable.Color(Val,Master.UnderlineColor);
    ipropStrikeTrough  : StrikeTrough := True;

    ipropULType   : begin
      case TRtfULType(Val) of
        rultDash    : Underline := axcuDash;
        rultDashD   : Underline := axcuDotDash;
        rultDashDD  : Underline := axcuDotDotDash;
        rultDB      : Underline := axcuDouble;
        rultHWave   : Underline := axcuWavyHeavy;
        rultLDash   : Underline := axcuDashLong;
        rultD       : Underline := axcuDotted;
        rultTH      : Underline := axcuThick;
        rultTHD     : Underline := axcuDottedHeavy;
        rultTHDash  : Underline := axcuDashedHeavy;
        rultTHDashD : Underline := axcuDashDotHeavy;
        rultTHDashDD: Underline := axcuDashDotDotHeavy;
        rultTHLDash : Underline := axcuDashLongHeavy;
        rultUlDbWave: Underline := axcuWavyDouble;
        rultW       : Underline := axcuWords;
        rultWave    : Underline := axcuWave;
      end;
    end;

    ipropSub        : SubSuperscript := axcssSubscript;
    ipropSuper      : SubSuperscript := axcssSuperscript;
    ipropNoSuperSub : SubSuperscript := axcssNone;

    ipropChcbpat    : FillColor := FColorTable.Color(Val,Master.FillColor);

    else              Result := False;
  end;
end;

procedure TRtfCHP.Assign(ACHP: TRtfCHP);
begin
  inherited Assign(ACHP);
end;

procedure TRtfCHP.Clear;
begin
  inherited;

  if FCurrBorder <> Nil then begin
    FCurrBorder.Free;
    FCurrBorder := Nil;
  end;
end;

constructor TRtfCHP.Create(ACHP: TAXWCHP; AColorTable: TRtfDestColorTable; AFontTable: TRtfDestFontTable);
begin
  inherited Create(ACHP);

  FColortable := ACOlorTable;
  FFontTable := AFontTable;
end;

destructor TRtfCHP.Destroy;
begin
  if FCurrBorder <> Nil then
    FCurrBorder.Free;

  inherited;
end;

{ TRtfPicture }

procedure TRtfDestPict.ApplyKeyword(iprop: TRtfIPROP; val: integer);
begin
  case iprop of
    ipropPicture       : FPictType := TAXWPictureType(Val);
    ipropPictWidth     : FWidth := Val;
    ipropPictHeigth    : FHeight := Val;
    ipropPictGoalWidth : FGoalWidth := Val / 20;
    ipropPictGoalHeight: FGoalHeight := Val / 20;
    ipropPictScaleX    : FScaleX := Val / 100;
    ipropPictScaleY    : FScaleY := Val / 100;
    ipropPictCropT     : FCropTop := Val;
    ipropPictCropB     : FCropBottom := Val;
    ipropPictCropL     : FCropLeft := Val;
    ipropPictCropR     : FCropRight := Val;
    ipropBlipTag       : FTag := Val;
  end;
end;

procedure TRtfDestPict.Clear;
begin
  FPictType  := aptUnknown;
  FExData    := -1;
  FWidth     := 0;
  FHeight    := 0;
  FGoalWidth := 0;
  FGoalHeight:= 0;
  FScaleX    := 1;
  FScaleY    := 1;
  FCropLeft  := 0;
  FCropTop   := 0;
  FCropRight := 0;
  FCropBottom:= 0;
//  FWrap      := agowInLine;
end;

constructor TRtfDestPict.Create(AReader: TAXWReadRTF; AType: TRtfIDEST);
begin
  inherited Create(AReader,AType);

  Clear;
end;

destructor TRtfDestPict.Destroy;
begin
  inherited;
end;

procedure TRtfDestPict.EndGroup;
begin
  inherited EndGroup;
end;

procedure TRtfDestPict.EndDest;
//var
//  Para: TAXWLogPara;
//  Pict: TAXWPicture;
//  CRGr: TAXWCharRunGraphic;
//  W,H : double;
begin
//  if FStream.Size > 0 then begin
////    FStream.SaveToFile('d:\temp\t1.jpg');
////    FStream.Seek(0,soFromBeginning);
//
//    Para := FReader.GetPara;
//
//    Pict := FReader.FDoc.Pictures.AddPicture(FStream,AXWPictureTypeName[FPictType]);
//    CRGr := Para.Runs.AddGraphic(Pict);
//
//    W := FGoalWidth * FScaleX;
//    if W <= 0 then
//      W := FWidth;
//    H := FGoalHeight * FScaleY;
//    if H <= 0 then
//      H := FHeight;
//    CRGr.Graphic.Width := W;
//    CRGr.Graphic.Height := H;
//  end;

  inherited EndDest;
end;

{ TRtfDestination }

procedure TRtfDestination.ApplyKeyword(iprop: TRtfIPROP; val: integer);
begin

end;

procedure TRtfDestination.BeginDest;
begin
  FGroupLevel := 0;
end;

procedure TRtfDestination.BeginGroup;
begin
  Inc(FGroupLevel);
end;

procedure TRtfDestination.ChangeType(ANewType: TRtfIDEST);
begin
  FType := ANewType;
end;

procedure TRtfDestination.ChildEndDest(AChild: TRtfDestination);
begin

end;

constructor TRtfDestination.Create(AReader: TAXWReadRTF; AType: TRtfIDEST; AParentDest: TRtfDestination);
begin
  Create(AReader,AType);

  FParentDest := AParentDest;
end;

constructor TRtfDestination.Create(AReader: TAXWReadRTF; AType: TRtfIDEST);
begin
  FReader := AReader;
  FType := AType;
end;

destructor TRtfDestination.Destroy;
begin

  inherited;
end;

procedure TRtfDestination.EndComposedValue;
begin

end;

procedure TRtfDestination.EndDest;
begin
  if FParentDest <> Nil then
    FParentDest.ChildEndDest(Self);

  FGroupLevel := 0;
end;

procedure TRtfDestination.EndGroup;
begin
  Dec(FGroupLevel);
end;

function TRtfDestination.HasDocText: boolean;
begin
  Result := False;
end;

procedure TRtfDestination.ParseChar(var ARIS: TRtfRIS; AChar: AnsiChar);
begin
  FText := FText + AxUCChar(AChar);
end;

function TRtfDestination.Persistent: boolean;
begin
  Result := False;
end;

{ TRtfDestColorTable }

procedure TRtfDestColorTable.ApplyKeyword(iprop: TRtfIPROP; val: integer);
begin
  case iprop of
    ipropRed  : FR := Byte(val);
    ipropGreen: FG := Byte(val);
    ipropBlue : FB := Byte(val);
  end;
end;

procedure TRtfDestColorTable.EndComposedValue;
begin
  SetLength(FColorTable,Length(FColorTable) + 1);
  FColorTable[High(FColorTable)] := (FR shl 16) or (FG shl 8) or FB;
end;

function TRtfDestColorTable.Color(AClIndex: integer; ADefault: longword): longword;
begin
  if AClIndex <= High(FColorTable) then
    Result := FColorTable[AClIndex]
  else
    Result := ADefault;
end;

function TRtfDestColorTable.Persistent: boolean;
begin
  Result := True;
end;

{ TRtfDestFontTable }

procedure TRtfDestFontTable.ApplyKeyword(iprop: TRtfIPROP; val: integer);
begin
  case iprop of
    ipropFont: begin
      FCurrFont := TRtfFont.Create;
      FCurrFont.Id := Val;
      FFonts.AddObject(IntToStr(FCurrFont.Id),FCurrFont);
    end;
  end;
end;

constructor TRtfDestFontTable.Create(AReader: TAXWReadRTF; AType: TRtfIDEST);
begin
  inherited Create(AReader,AType);

  FFonts := THashedStringList.Create;
  FDefaultFont := TRtfFont.Create;
  FDefaultFont.Name := AXW_CHP_DEFAULT_FONT;
end;

destructor TRtfDestFontTable.Destroy;
var
  i: integer;
begin
  for i := 0 to FFonts.Count - 1 do
    FFonts.Objects[i].Free;

  FFonts.Free;

  FDefaultFont.Free;

  inherited;
end;

procedure TRtfDestFontTable.EndComposedValue;
begin

end;

procedure TRtfDestFontTable.EndGroup;
begin
  inherited;

end;

function TRtfDestFontTable.GetFont(AId: integer): TRtfFont;
var
  i: integer;
begin
  i := FFonts.IndexOf(IntToStr(AId));

  if i >= 0 then
    Result := TRtfFont(FFonts.Objects[i])
  else
    Result := FDefaultFont;
end;

procedure TRtfDestFontTable.ParseChar(var ARIS: TRtfRIS; AChar: AnsiChar);
begin
  if FGroupLevel = 1 then
    FCurrFont.Name := FCurrFont.Name + AxUCChar(AChar);
end;

function TRtfDestFontTable.Persistent: boolean;
begin
  Result := True;
end;

{ TRtfDestBookmark }

//procedure TRtfDestBookmark.ApplyKeyword(iprop: TRtfIPROP; val: integer);
//begin
//  case iprop of
//    ipropBkmkColFirst: FColFirst := val;
//    ipropBkmkColLast : FColLast := val;
//  end;
//end;
//
//constructor TRtfDestBookmark.Create(AReader: TAXWReadRTF; AType: TRtfIDEST; ABookmarks: THashedStringList);
//begin
//  inherited Create(AReader,AType);
//
//  FBookmarks := ABookmarks;
//end;
//
//destructor TRtfDestBookmark.Destroy;
//begin
//
//  inherited;
//end;
//
//procedure TRtfDestBookmark.EndGroup;
//var
//  i    : integer;
//  Bmk  : TAXWBookmark;
//  CR   : TAXWCharRunLinked;
//  Start: TRtfDestBookmark;
//  Para : TAXWLogPara;
//begin
//  inherited;
//
//  Para := FReader.GetPara;
//  if FType = idestBkmkEnd then begin
//    i := FBookmarks.IndexOf(FText);
//    if i >= 0 then begin
//      Start := TRtfDestBookmark(FBookmarks.Objects[i]);
//      CR := TAXWCharRunLinked(Start.FSel.FirstRun);
//      // Not single point bookmark.
//      if ((Para.Count - 1) > CR.Index) or (Para.Runs <> CR.Parent) then begin
//        FSel := TAXWFieldSelectionImpl(CR.AsBookmark.Selection);
//        FSel.LastPara := Para;
//        CR.AddNext(Nil,Para.Runs);
//        FSel.LastRun := CR.LinkedNext;
//      end;
//
//      FBookmarks.Objects[i].Free;
//      FBookmarks.Delete(i);
//    end;
//  end
//  else begin
//    FSel := TAXWFieldSelectionImpl.Create;
//    FSel.FirstPara := Para;
//    Bmk := FReader.FDoc.Bookmarks.Add(FText,FSel);
//    Bmk.FirstCol := FColFirst;
//    Bmk.LastCol := FColLast;
//    FSel<.FirstRun := Para.Runs.AddLinked(Nil,Bmk);
//    FBookmarks.AddObject(FText,Self);
//  end;
//end;
//
//function TRtfDestBookmark.Persistent: boolean;
//begin
//  Result := FType = idestBkmkStart;
//end;

{ TRtfDestStyle }

procedure TRtfStyle.AddStyle(AStyle: TAXWStyle);
begin
  FReader.FStyles.AddObject(IntToStr(FId),AStyle);
end;

constructor TRtfStyle.Create(AReader: TAXWReadRTF; AId: integer);
begin
  FReader := AReader;

  FId := AId;
end;

{ TRtfDestParaStyle }

procedure TRtfParaStyle.ApplyKeyword(iprop: TRtfIPROP; val: integer);
begin
  inherited ApplyKeyword(iprop,val);

  case iprop of
    ipropLeftInd  : FPAPX.IndentLeft := Val / 20;
    ipropRightInd : FPAPX.IndentRight := Val / 20;
    ipropFirstInd : begin
      if Val > 0 then
        FPAPX.IndentFirstLine := Val / 20
      else
        FPAPX.IndentHanging := -Val / 20;
    end;
  end;
end;

constructor TRtfParaStyle.Create(AReader: TAXWReadRTF; AId: integer);
begin
  inherited Create(AReader,AId);

  FCHPX := TRtfCHP.Create(AReader.FDoc.MasterCHP,AReader.FColorTable,AReader.FFontTable);
  FPAPX := TRtfPAP.Create(AReader.FDoc.MasterPAP,AReader.FColorTable);
end;

destructor TRtfParaStyle.Destroy;
begin
  FPAPX.Free;
  FCHPX.Free;

  inherited;
end;

procedure TRtfParaStyle.Save(AName: AxUCString);
var
  S    : TAXWStyle;
  Style: TAXWParaStyle;
begin
  if (AName <> '') then begin

    S := FReader.FDoc.Styles.FindByName(AName,astPara);
    if S <> Nil then
      Style := TAXWParaStyle(S)
    else
      Style := FReader.FDoc.Styles.AddPara(AName);

    Style.PAPX.Assign(FPAPX);

    AddStyle(Style);
  end;

  inherited;
end;

{ TRtfDestStylesheet }

procedure TRtfDestStylesheet.ApplyKeyword(iprop: TRtfIPROP; val: integer);
begin
  inherited;

  case iprop of
    ipropCharRunStyle: FCurrStyle := TRtfCharStyle.Create(FReader,Val);
    ipropParaStyle   : FCurrStyle := TRtfParaStyle.Create(FReader,Val);
    ipropTableStyle  : FCurrStyle := TRtfTableStyle.Create(FReader,Val);
    ipropSectStyle   : FCurrStyle := TRtfSectStyle.Create(FReader,Val);
  end;
end;

procedure TRtfDestStylesheet.BeginGroup;
begin
  inherited;

  if FGroupLevel = 1 then begin
    if FCurrStyle <> Nil then
      FCurrStyle.Free;
    FCurrStyle := Nil;
  end;
end;

procedure TRtfDestStylesheet.EndGroup;
begin
  if FGroupLevel = 1 then begin
    if FCurrStyle <> Nil then begin
      FCurrStyle.Save(FText);
      FCurrStyle.Free;
    end;
    FCurrStyle := Nil;
  end;

  FText := '';

  inherited;
end;

{ TRtfDestCharStyle }

procedure TRtfCharStyle.ApplyKeyword(iprop: TRtfIPROP; val: integer);
begin
  inherited ApplyKeyword(iprop,val);

  FCHPX.ApplyKeyword(iprop,val);
end;

constructor TRtfCharStyle.Create(AReader: TAXWReadRTF; AId: integer);
begin
  inherited Create(AReader,AId);

  FCHPX := TRtfCHP.Create(AReader.FDoc.MasterCHP,AReader.FColorTable,AReader.FFontTable);
end;

destructor TRtfCharStyle.Destroy;
begin
  FCHPX.Free;

  inherited;
end;

procedure TRtfCharStyle.Save(AName: AxUCString);
var
  S    : TAXWStyle;
  Style: TAXWCharStyle;
begin
  if (AName <> '') then begin

    S := FReader.FDoc.Styles.FindByName(AName,astCharRun);
    if S <> Nil then
      Style := TAXWCharStyle(S)
    else
      Style := FReader.FDoc.Styles.AddChar(AName);

    Style.CHPX.Assign(FCHPX);

    AddStyle(Style);
  end;

  inherited;
end;

{ TRtfDestTableStyle }

procedure TRtfTableStyle.ApplyKeyword(iprop: TRtfIPROP; val: integer);
begin
  inherited;

end;

{ TRtfDestSectStyle }

procedure TRtfSectStyle.ApplyKeyword(iprop: TRtfIPROP; val: integer);
begin
  inherited;

end;

procedure TRtfTableStyle.Save(AName: AxUCString);
begin

end;

procedure TRtfSectStyle.Save(AName: AxUCString);
begin

end;

{ TRtfDestField }

procedure TRtfDestField.ApplyKeyword(iprop: TRtfIPROP; val: integer);
begin
  case iprop of
    ipropDataField: FInstrDone := True;
  end;
end;

procedure TRtfDestField.ChangeType(ANewType: TRtfIDEST);
var
  S: AxUCString;
  i: integer;
begin
  if FType = idestFieldInstr then begin
    FText := Trim(FText);
    S := Uppercase(SplitAtChar(' ',FText));
    // TODO Hyperlinks to bookmarks.
    if (S = 'HYPERLINK') and (FText <> '') then begin
      // Strip any parameters
      i := RCPos(' ',FText);
      if i > 1 then
        FText := Copy(FText,i + 1,MAXINT);

      StripQuotes(FText);
      FReader.FCurrHyperlink := FReader.FDoc.Hyperlinks.Add(FText);
      FReader.FCharRunType := acrtHyperlink;
    end
    else begin
//      i := RCPos(' ',FText);
//      if i > 1 then begin
//        S := Copy(FText,1,i - 1);
//        FText := Copy(FText,i + 1,MAXINT);
//      end;
//      StripQuotes(FText);
//
//      FReader.FCurrSimpField := CreateSimpleField(S + ' ' + FText);
//      FReader.FCharRunType := acrtSimpleField;
    end;
  end;
  inherited ChangeType(ANewType);
end;

procedure TRtfDestField.EndDest;
begin
  inherited EndDest;

  FReader.FCharRunType := acrtText;
  FReader.FCurrHyperlink := Nil;
//  FReader.FCurrSimpField := Nil;
end;

function TRtfDestField.HasDocText: boolean;
begin
  Result := FType = idestFieldResult;
end;

procedure TRtfDestField.ParseChar(var ARIS: TRtfRIS; AChar: AnsiChar);
begin
  if not FInstrDone then
    inherited;
end;

{ TRtfDestStringParam }

procedure TRtfDestStringParam.EndDest;
//var
//  Cmt: TAXWComment;
begin
  case FType of
    idestGenerator: FReader.FGenerator := FText;
    idestAnnAuthor: begin
//      if FReader.FComments.Count > 0 then begin
//        Cmt := TAXWComment(FReader.FComments.Objects[FReader.FComments.Count - 1]);
//        Cmt.Author := FText;
//      end;
    end;
    idestAnnInitials: begin
//      if FReader.FComments.Count > 0 then begin
//        Cmt := TAXWComment(FReader.FComments.Objects[FReader.FComments.Count - 1]);
//        Cmt.Initials := FText;
//      end;
    end;
  end;

  inherited EndDest;
end;

{ TRtfDestComment }

procedure TRtfDestComment.ApplyKeyword(iprop: TRtfIPROP; val: integer);
begin
  inherited ApplyKeyword(iprop,val);

  case iprop of
    ipropPar: FText := FText + #13;
  end;
end;

procedure TRtfDestComment.ChildEndDest(AChild: TRtfdestination);
begin
  inherited;

  if (FType = idestAnnotation) and (AChild.FType = idestAnnRef) then
    FRef := AChild.FText;
end;

constructor TRtfDestComment.Create(AReader: TAXWReadRTF; AType: TRtfIDEST; AComments: TStrings);
begin
  inherited Create(AReader,AType);

  FComments := AComments;
end;

procedure TRtfDestComment.EndDest;
//var
//  i  : integer;
//  Sel: TAXWFieldSelectionImpl;
//  Cmt: TAXWComment;
begin
  inherited EndDest;

//  case FType of
//    idestAnnStart: begin
//      Sel := TAXWFieldSelectionImpl.Create;
//      Sel.FirstPara := FReader.GetPara;
//      Sel.LastPara := Sel.FirstPara;
//
//      Cmt := TAXWComment.Create(Sel);
//
//      Sel.FirstRun := Sel.FirstPara.Runs.AddLinked(FReader.FCHP,Cmt);
//      Sel.LastRun := Sel.LastPara.Runs.Last;
//
//      FComments.AddObject(FText,Cmt);
//    end;
//    idestAnnEnd: begin
//      i := FComments.IndexOf(FText);
//      if i < 0 then
//        Exit;
//
//      Cmt := TAXWComment(FComments.Objects[i]);
//
//      Sel := TAXWFieldSelectionImpl(Cmt.Selection);
//
//      Sel.LastPara := FReader.GetPara;
//      Sel.LastRun := TAXWCharRunLinked(Sel.FirstRun).AddNext(Nil,Sel.LastPara.Runs);
//    end;
//    idestAnnotation: begin
//      i := FComments.IndexOf(FRef);
//      if i < 0 then
//        Exit;
//
//      Cmt := TAXWComment(FComments.Objects[i]);
//      Cmt.Text.Text := FText;
//    end;
//  end;
end;

{ TRtfDestStream }

constructor TRtfDestStream.Create(AReader: TAXWReadRTF; AType: TRtfIDEST; AParentDest: TRtfDestination);
begin
  inherited Create(AReader,AType,AParentDest);

  FStream := TMemoryStream.Create;
  FCurrByteUpper := True;
end;

constructor TRtfDestStream.Create(AReader: TAXWReadRTF; AType: TRtfIDEST);
begin
  inherited Create(AReader,AType);

  FStream := TMemoryStream.Create;
  FCurrByteUpper := True;
end;

destructor TRtfDestStream.Destroy;
begin
  FStream.Free;

  inherited;
end;

procedure TRtfDestStream.EndDest;
begin
  FStream.Size := 0;
end;

procedure TRtfDestStream.ParseChar(var ARIS: TRtfRIS; AChar: AnsiChar);
var
  b: byte;
begin
  case ARIS of
    risNorm: begin
      b := HexCharTable[Integer(AChar)];
      if FCurrByteUpper then
        FCurrByte := b shl 4
      else begin
        FCurrByte := FCurrByte + b;
        FStream.Write(FCurrByte,1);
      end;
      FCurrByteUpper := not FCurrByteUpper;
    end;
    risBin : ;
    risHex : ;
  end;
end;

{ TRtfTableStack }

procedure TRtfTableStack.ApplyOnRow(ARow: TAXWTableRow);
var
  i    : integer;
  X1,X2: double;
begin
  X1 := 0;
  for i := 0 to Min(FCurrItem.CellX.Count - 1,ARow.Count - 1) do begin
    X2 := FCurrItem.CellX[i] / 20;
    ARow[i].Width := X2 - X1;
    X1 := X2;
  end;
  for i := 0 to Min(FCurrItem.CellColor.Count - 1,ARow.Count - 1) do begin
    if FCurrItem.CellColor[i] >= 0 then begin
      ARow[i].AddProps;
      ARow[i].Props.FillColor := FColorTable.Color(FCurrItem.CellColor[i],AXW_COLOR_AUTOMATIC);
    end;
  end;
  for i := 0 to Min(FCurrItem.Borders.Count - 1,ARow.Count - 1) do begin
    if FCurrItem.Borders[i].Assigned then begin
      ARow[i].AddProps;
      ARow[i].Props.Borders.Assign(FCurrItem.Borders[i]);
    end;
  end;

  if FCurrItem.RowHeight > 0 then
    ARow.AHeight := FCurrItem.RowHeight;
end;

procedure TRtfTableStack.Clear;
var
  Item: TRtfTableStackItem;
begin
  while Count > 0 do begin
    Item := TRtfTableStackItem(inherited Pop);
    Item.Free;
  end;
end;

procedure TRtfTableStack.ClearRow;
begin
  FCurrItem.Clear;
end;

constructor TRtfTableStack.Create(AColorTable: TRtfDestColorTable);
begin
  inherited Create;

  FColorTable := AColorTable;
  FCurrItem := TRtfTableStackItem.Create;
end;

destructor TRtfTableStack.Destroy;
begin
  Clear;

  FCurrItem.Free;

  inherited;
end;

function TRtfTableStack.Peek: TRtfTableStackItem;
begin
  Result := TRtfTableStackItem(inherited Peek);
end;

procedure TRtfTableStack.Pop(var ATable: TAXWTable; var ARow: TAXWTableRow; var ACell: TAXWTableCell);
var
  Item: TRtfTableStackItem;
begin
  Item := TRtfTableStackItem(inherited Pop);
  FCurrItem.Assign(Item);
  ACell := Item.Cell;
  if ACell <> Nil then begin
    ARow := ACell.Row;
    ATable := ARow.Table;
  end
  else begin
    ARow := Nil;
    ATable := Nil;
  end;
  Item.Free;
end;

procedure TRtfTableStack.Push(ACell: TAXWTableCell);
var
  Item: TRtfTableStackItem;
begin
  Item := TRtfTableStackItem.Create;
  Item.Cell := ACell;
  Item.Assign(FCurrItem);

  inherited Push(Item);

  ClearRow;
end;

{ TRtfTableStackItem }

procedure TRtfTableStackItem.Assign(AItem: TRtfTableStackItem);
begin
  FCellX.Assign(AItem.FCellX);
  FCellColor.Assign(AItem.FCellColor);
  FBorders.Assign(AItem.FBorders);
  FRowHeight := AItem.FRowHeight;
end;

procedure TRtfTableStackItem.Clear;
begin
  FCell := Nil;
  FCellX.Clear;
  FCellColor.Clear;
  FBorders.Clear;
  FRowHeight := 0;
end;

constructor TRtfTableStackItem.Create;
begin
  FCellX := TAXWIntegerList.Create;
  FBorders := TAXWDocPropBordersList.Create;
  FCellColor := TAXWIntegerList.Create;
end;

destructor TRtfTableStackItem.Destroy;
begin
  FCellX.Free;
  FBorders.Free;
  FCellColor.Free;

  inherited;
end;

{ TRftDestListText }

procedure TRftDestListText.ApplyKeyword(iprop: TRtfIPROP; val: integer);
begin
  inherited ApplyKeyword(iprop,val);

end;

constructor TRftDestListText.Create(AReader: TAXWReadRTF; AType: TRtfIDEST);
begin
  inherited Create(AReader,AType);

  FReader.FDoc.Paras.AutoNumbering := anstBullet;
end;

procedure TRftDestListText.EndGroup;
begin
  inherited EndGroup;

  FReader.FExText := FText;
  FText := '';
end;

procedure TRftDestListText.ParseChar(var ARIS: TRtfRIS; AChar: AnsiChar);
begin
  FText := AXUCChar(AChar);
end;

end.
