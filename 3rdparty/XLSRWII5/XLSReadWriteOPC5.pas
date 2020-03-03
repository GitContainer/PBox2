unit XLSReadWriteOPC5;

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

// This unit implements Acts II, III, VII and IX of the
// Open Packaging Convetion, signed in Wienna 1934.

                                           
uses Classes, SysUtils, IniFiles, Contnrs,
     XLSUtils5, XLSReadWriteZIP5,
     xpgPUtils, xpgParseOPC, xpgParseContentType;

const OPC_DIR_CONTENTTYPES   = '[Content_Types].xml';

const OOXML_URI_MARKUP_COMPATIBILITY               = 'http://schemas.openxmlformats.org/markup-compatibility/2006';

const OOXML_URI_PACKAGE_CONTENTTYPES               = 'http://schemas.openxmlformats.org/package/2006/content-types';
const OOXML_URI_PACKAGE_RELATIONSHIPS              = 'http://schemas.openxmlformats.org/package/2006/relationships';

const OOXML_URI_URN_SCHEMAS_MICROSOFT_OFFICE       = 'urn:schemas-microsoft-com:office:office';
const OOXML_URI_URN_SCHEMAS_MICROSOFT_OFFICE_WORD  = 'urn:schemas-microsoft-com:office:word';
const OOXML_URI_URN_SCHEMAS_MICROSOFT_OFFICE_WORDML= 'http://schemas.microsoft.com/office/word/2006/wordml';
const OOXML_URI_URN_SCHEMAS_MICROSOFT_VML          = 'urn:schemas-microsoft-com:vml';

const OOXML_URI_OFFICEDOC_EXTENDED_PROPERIES       = 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties';
const OOXML_URI_OFFICEDOC_EXTENDED_DOCPROPSVTYPES  = 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes';
const OOXML_URI_OFFICEDOC_RELATIONSHIPS            = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const OOXML_URI_OFFICEDOC_DRAWING                  = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const OOXML_URI_OFFICEDOC_CHART                    = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const OOXML_URI_OFFICEDOC_CHARTDRAWING             = 'http://schemas.openxmlformats.org/drawingml/2006/chartDrawing';
const OOXML_URI_OFFICEDOC_MATH                     = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
const OOXML_URI_OFFICEDOC_DRAWINGML_WPDRAWING      = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';

const OOXML_URI_SPREADSHEETML_MAIN                 = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

const OOXML_URI_WORDPROCESSINGML_MAIN              = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

const OPC_ROOT                  = '__ROOT__';
const OPC_XLSX_DOCPROPS_APP     = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties';
const OPC_XLSX_DOCPROPS_CORE    = 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties';
const OPC_XLSX_WORKBOOK         = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';
const OPC_XLSX_WORKSHEET        = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet';
const OPC_XLSX_CHARTSHEET       = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet';
const OPC_XLSX_CALCCHAIN        = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain';
const OPC_XLSX_SST              = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings';
const OPC_XLSX_CONNECTIONS      = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections';
const OPC_XLSX_STYLES           = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles';
const OPC_XLSX_THEME            = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme';
const OPC_XLSX_HYPERLINK        = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';
const OPC_XLSX_COMMENTS         = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments';
const OPC_XLSX_DRAWING          = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing';
const OPC_XLSX_VMLDRAWING       = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing';
const OPC_XLSX_TABLE            = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table';
const OPC_XLSX_QUERYTABLE       = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable';
const OPC_XLSX_XLINK            = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink';
const OPC_XLSX_XLINK_PATH       = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath';
const OPC_XLSX_PRINTER_SETTINGS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings';
const OPC_XLSX_IMAGE            = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
const OPC_XLSX_CHART            = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart';
const OPC_XLSX_CHARTSTYLE       = 'http://schemas.microsoft.com/office/2011/relationships/chartStyle';
const OPC_XLSX_CHARTCOLORS      = 'http://schemas.microsoft.com/office/2011/relationships/chartColorStyle';
const OPC_XLSX_CHARTUSERSHAPES  = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartUserShapes';
const OPC_XLSX_PIVOTCACHEDEF    = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition';
const OPC_XLSX_PIVOTCACHERECS   = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords';
const OPC_XLSX_PIVOTTABLE       = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable';

const OPC_XLSX_DOCPROPS_APP_TARGET     = 'docProps/app.xml';
const OPC_XLSX_DOCPROPS_CORE_TARGET    = 'docProps/core.xml';
const OPC_XLSX_WORKBOOK_TARGET         = 'xl/workbook.xml';
const OPC_XLSX_WORKSHEET_TARGET        = 'worksheets/sheet%d.xml';
const OPC_XLSX_WORKSHEET_TARGET_BIN    = 'worksheets/sheet%d.bin';
const OPC_XLSX_CHARTSHEET_TARGET       = 'chartsheets/sheet%d.xml';
const OPC_XLSX_CALCCHAIN_TARGET        = 'calcChain.xml';
const OPC_XLSX_SST_TARGET              = 'sharedStrings.xml';
const OPC_XLSX_CONNECTIONS_TARGET      = 'connections.xml';
const OPC_XLSX_STYLES_TARGET           = 'styles.xml';
const OPC_XLSX_THEME_TARGET            = 'theme/theme%d.xml';
const OPC_XLSX_HYPERLINK_TARGET        = '';
const OPC_XLSX_COMMENTS_TARGET         = '../comments%d.xml';
const OPC_XLSX_TABLE_TARGET            = '../tables/table%d.xml';
const OPC_XLSX_QUERYTABLE_TARGET       = '../queryTables/queryTable%d.xml';
const OPC_XLSX_VMLDRAWING_TARGET       = '../drawings/vmlDrawing%d.vml';
const OPC_XLSX_DRAWING_TARGET          = '../drawings/drawing%d.xml';
const OPC_XLSX_XLINK_TARGET            = 'externalLinks/externalLink%d.xml';
const OPC_XLSX_PRINTER_SETTINGS_TARGET = '../printerSettings/printerSettings%d.bin';
const OPC_XLSX_IMAGE_TARGET            = '../media/image%d.%s';
const OPC_XLSX_CHART_TARGET            = '../charts/chart%d.xml';
const OPC_XLSX_CHARTSTYLE_TARGET       = 'style%d.xml';
const OPC_XLSX_CHARTCOLORS_TARGET      = 'colors%d.xml';
const OPC_XLSX_PIVOTCACHEDEF_TARGET    = 'pivotCache/pivotCacheDefinition%d.xml';
const OPC_XLSX_PIVOTCACHERECS_TARGET   = 'pivotCacheRecords%d.xml';
const OPC_XLSX_PIVOTTABLE_TARGET       = '../pivotTables/pivotTable%d.xml';

const OPC_DOCX_DOCPROPS_APP            = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties';
const OPC_DOCX_DOCPROPS_CORE           = 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties';
const OPC_DOCX_DOCUMENT                = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';
const OPC_DOCX_SETTINGS                = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings';
const OPC_DOCX_STYLES                  = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles';
const OPC_DOCX_THEME                   = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme';
const OPC_DOCX_NUMBERING               = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering';
const OPC_DOCX_IMAGE                   = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
const OPC_DOCX_HEADER                  = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header';
const OPC_DOCX_FOOTER                  = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer';
const OPC_DOCX_FOOTNOTES               = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes';
const OPC_DOCX_ENDNOTES                = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes';

const OPC_DOCX_DOCPROPS_APP_TARGET     = 'docProps/app.xml';
const OPC_DOCX_DOCPROPS_CORE_TARGET    = 'docProps/core.xml';
const OPC_DOCX_DOCUMENT_TARGET         = 'word/document.xml';
const OPC_DOCX_SETTINGS_TARGET         = 'settings.xml';
const OPC_DOCX_STYLES_TARGET           = 'styles.xml';
const OPC_DOCX_THEME_TARGET            = 'theme/theme%d.xml';
const OPC_DOCX_NUMBERING_TARGET        = 'numbering.xml';
const OPC_DOCX_IMAGE_TARGET            = '../media/image%d.%s';
const OPC_DOCX_HEADER_TARGET           = 'header%d.xml';
const OPC_DOCX_FOOTER_TARGET           = 'footer%d.xml';
const OPC_DOCX_FOOTNOTES_TARGET        = 'footnotes.xml';
const OPC_DOCX_ENDNOTES_TARGET         = 'endnotes.xml';

type TOPCTargetMode = (otmExternal,otmInternal);

type TOPC_ContentType = class(TObject)
protected
     FContentType: AxUCString;
public
     property ContentType: AxUCString read FContentType write FContentType;
     end;

type TOPC_ContentTypeOverride = class(TOPC_ContentType)
protected
     FPartName: AxUCString;
public
     property PartName: AxUCString read FPartName write FPartName;
     end;

type TOPC_ContentTypeDefault = class(TOPC_ContentType)
protected
     FExtension: AxUCString;
public
     property Extension: AxUCString read FExtension write FExtension;
     end;

type TOPC_ContentTypes = class(TObjectList)
private
     function GetItems(Index: integer): TOPC_ContentType;
protected
     FXLSXMode: boolean;
public
     constructor Create;

     procedure Clear; reintroduce;

     procedure SetDefault;

     function  FindOverride(APartName: AxUCString): integer;
     function  FindDefault(AExtension: AxUCString): integer;

     procedure AddOverride(APartName, AContentType: AxUCString);
     procedure AddDefault(AExtension, AContentType: AxUCString);

     property XLSXMode: boolean read FXLSXMode write FXLSXMode;
     property Items[Index: integer]: TOPC_ContentType read GetItems; default;
     end;

type TOPCState = (opcsClosed,opcsRead,opcsWrite);

type TOPCContainer = class;
     TOPCItemList = class;

     TOPCItem = class(TObject)
private
     function  GetItems(Index: integer): TOPCItem;
     function  GetTarget: AxUCString;
     procedure SetTarget(const Value: AxUCString);
     function  GetId: AxUCString;
protected
     FOwner     : TOPCContainer;
     FParent    : TOPCItem;
     FIndex     : integer;
{$ifdef DELPHI_5}
     FChilds    : TStringList;
{$else}
     FChilds    : THashedStringList;
{$endif}
     FState     : TOPCState;

     FPath      : AxUCString;
     FName      : AxUCString;

     FType      : AxUCString;
     FTarget    : AxUCString;
     FTargetMode: TOPCTargetMode;
     FText      : AxUCString;

     FCheckedOut: boolean;
     FData      : TStream;
     FContent   : AxUCString;

     FDebugId   : integer;

     procedure ReadChilds;
     procedure WriteChilds;
     function  Add(ARelationship: TCT_Relationship): TOPCItem;
     function  MakeFilename_rels: AxUCString;
     function  MakeId: AxUCString;
public
     constructor Create(AOwner: TOPCContainer; AParent: TOPCItem);
     destructor Destroy; override;

     procedure Clear;

     function  IsRoot: boolean;
     function  Count: integer;
     function  Find(AId: AxUCString): TOPCItem;
     function  FindByType(ATypeName: AxUCString; AList: TOPCItemList; AClearList: boolean = True): boolean; overload;
     function  FindByType(ATypeName: AxUCString): TOPCItem; overload;
     function  FindByTarget(ATarget: AxUCString): TOPCItem;
     function  FindTarget(AId: AxUCString): AxUCString;
     procedure Checkout;
     procedure CheckoutAll;

     function  AbsolutePath: AxUCString;
     function  AbsoluteTarget: AxUCString;

     function  AddNew(AType,ATarget: AxUCString): TOPCItem;
     function  AddNewExternal(AType,ATarget: AxUCString): TOPCItem;
     procedure Close;

     property Parent    : TOPCItem read FParent;
     property State     : TOPCState read FState;
     // Paths can be relative (..)
     property Path      : AxUCString read FPath write FPath;
     property Name      : AxUCString read FName write FName;
     property Type_     : AxUCString read FType write FType;
     property Target    : AxUCString read GetTarget write SetTarget;
     property TargetMode: TOPCTargetMode read FTargetMode write FTargetMode;
     property Id        : AxUCString read GetId;
     property Text      : AxUCString read FText write FText;
     property CheckedOut: boolean read FCheckedOut;
     property Data      : TStream read FData write FData;
     property Content   : AxUCString read FContent write FContent;

     property Childs[Index: integer]: TOPCItem read GetItems; default;
     end;

     TOPCItemList = class(TObjectList)
private
     function GetItems(Index: integer): TOPCItem;
protected
public
     constructor Create;

     property Items[Index: integer]: TOPCItem read GetItems; default;
     end;

     TOPCContainer = class(TObject)
protected
     FState: TOPCState;
     FZIP: TXLSZipArchiveIntf;
     FAutoCheckout: boolean;

     FRoot: TOPCItem;
     FXmlns: AxUCString;
     FContentTypes: TOPC_ContentTypes;

     procedure DebugDump(AFileName: AxUCString);
     procedure ReadContentTypes;
     procedure WriteContentTypes;
public
     constructor Create;
     destructor Destroy; override;
     procedure Clear;

     procedure OpenRead(AStream: TStream; const AUseAlternateZip: boolean); virtual;
     procedure OpenWrite(AStream: TStream; const AUseAlternateZip: boolean); virtual;
     procedure Close; virtual;

     function  ReadContentType: TStream;

     procedure SaveItemData(AItem: TOPCItem);
     procedure CollectUnchecked(AList: TOPCItemList);
     procedure SaveUnchecked(AList: TOPCItemList);

     function  ItemCreate(AOwner: TOPCItem; AType,ATarget,AContentType: AxUCString): TOPCItem;
     function  ItemCreateExternal(AOwner: TOPCItem; AType,ATarget: AxUCString): TOPCItem;
     function  ItemOpenRead(AOwner: TOPCItem; AType: AxUCString): TOPCItem; overload;
     function  ItemOpenRead(AOwner: TOPCItem; AType, AId: AxUCString): TOPCItem; overload;
     function  ItemOpenStream(AItem: TOPCItem): TStream;
     procedure ItemWrite(AItem: TOPCItem; AStream: TStream);
     procedure ItemClose(AItem: TOPCItem);
     function  ItemCreateStream(AItem: TOPCItem): TStream;
     procedure ItemCloseStream(AItem: TOPCItem; AStream: TStream);

     property State: TOPCState read FState;
     property AutoCheckout: boolean read FAutoCheckout write FAutoCheckout;
     property Root: TOPCItem read FRoot;
     property Xmlns: AxUCString read FXmlns write FXmlns;
     end;

type TOPC_XLSX = class(TOPCContainer)
protected
     FResultList          : TOPCItemList;
     FWorkbook            : TOPCItem;
     FCurrSheet           : TOPCItem;
     FCurrChart           : TOPCItem;
     FCurrDrawing         : TOPCItem;
     FCurrTable           : TOPCItem;
     FPrinterSettingsCount: integer;
     FChartCount          : integer;
     FTableCount          : integer;
     FQueryTableCount     : integer;
     FPivotTableCount     : integer;

     procedure CheckWorkbookOpen;
     function  ReadItemData(AItem: TOPCItem): TStream;
public
     constructor Create;
     destructor Destroy; override;
     procedure Clear;

     procedure OpenRead(AStream: TStream; const AUseAlternateZip: boolean); override;
     procedure OpenWrite(AStream: TStream; const AUseAlternateZip: boolean); override;
     procedure Close; override;

     function  FindSheet(AIndex: integer): TOPCItem;
     function  FindSheetBin(AIndex: integer): TOPCItem;
     function  FindDrawingId(AOPCSheet: TOPCItem): AxUCString;
     function  FindVmlDrawingId(AOPCSheet: TOPCItem): AxUCString;
     function  FindPrinterSettingsId(AOPCSheet: TOPCItem): AxUCString;
     function  FindPictureId(AOPCSheet: TOPCItem; AIndex: integer): AxUCString;
     function  FindTableId(AOPCSheet: TOPCItem; ATarget: AxUCString): AxUCString;
     function  FindAndOpenChart(AId: AxUCString): TOPCItem;

     function  ReadDocPropsApp: TStream;
     function  ReadDocPropsCore: TStream;
     function  ReadStyles: TStream;
     function  ReadSST: TStream;
     function  ReadConnections: TStream;
     function  ReadWorkbook: TStream;
     function  ReadComments(ASheetId: AxUCString): TStream;

     function  ReadTable(AId: AxUCString): TStream;
     function  ReadQueryTable: TStream;
     procedure CloseTable;

     function  ReadXLink(AId: AxUCString): TStream;
     function  ReadPrinterSettings(AId: AxUCString): TStream;
     function  ReadChart(AId: AxUCString): TStream;
     function  ReadChartStyle(AChart: TOPCItem; AId: AxUCString): TStream;
     function  ReadChartColors(AChart: TOPCItem; AId: AxUCString): TStream;
     function  ReadAny(AItem: TOPCItem): TStream;
     function  ReadVmlDrawing(AId: AxUCString; ARels: TStrings): TStream;
     function  ReadPivotCacheDefinition(AId: AxUCString): TStream;
     function  ReadPivotCacheRecords(const AOwnerId,AId: AxUCString): TStream;
     function  ReadPivotTable(AId: AxUCString): TStream;
     function  ReadTheme: TStream;

     function  OpenAndReadDrawing(AId: AxUCString): TStream;
     function  ReadDrawingMedia(AId: AxUCString; out ATarget: AxUCString): TStream;
     function  GetDrawingImages(AList: TStrings): boolean;
     procedure CloseDrawing;

     function  SheetIsChartSheet(AId: AxUCString): boolean;
     function  OpenAndReadSheet(AId: AxUCString): TStream;
     function  OpenAndReadChartSheet(AId: AxUCString): TStream;
     function  GetSheetPivotTables(AList: TStrings): boolean;
     procedure CloseSheet;

     function  AddDocPropsCore(AStream: TStream): TOPCItem;
     function  AddDocPropsApp(AStream: TStream): TOPCItem;
     function  AddTheme(AStream: TStream; AIndex: integer): TOPCItem;

     function  CreateSST: TOPCItem;
     function  CreateConnections: TOPCItem;
     function  CreateStyles: TOPCItem;
     function  CreateWorkbook(AMacroEnabled: boolean): TOPCItem;
     function  CreateSheet(AIndex: integer): TOPCItem;
     function  CreateChartSheet(AIndex: integer): TOPCItem;
     function  CreateComment(AOwner: TOPCItem; AIndex: integer): TOPCItem;
     function  CreateTable(AOwner: TOPCItem): TOPCItem;
     function  CreateQueryTable(AOwner: TOPCItem): TOPCItem;
     function  CreateDrawing(AOwner: TOPCItem; AIndex: integer): TOPCItem;
     function  CreateImage(const AIndex: integer; const AImageExt: AxUCString): TOPCItem;
     function  CreateChart: TOPCItem;
     procedure CloseChart(AItem: TOPCItem);
     function  CreateChartStyle(AOwner: TOPCItem): TOPCItem;
     function  CreateChartColors(AOwner: TOPCItem): TOPCItem;
     function  CreateChartUserShapes(AOwner: TOPCItem; AIndex: integer): TOPCItem;
     function  CreateVmlDrawing(AOwner: TOPCItem; AIndex: integer): TOPCItem;
     function  CreateXLink(AIndex: integer): TOPCItem;
     function  CreateXLinkPath(AOwner: TOPCItem; AFilename: AxUCString): TOPCItem;
     function  CreateHyperlink(AOwner: TOPCItem; ATarget: AxUCString): TOPCItem;
     function  CreatePivotCacheDefinition(AOwner: TOPCItem; AIndex: integer): TOPCItem;
     function  CreatePivotCacheRecords(AOwner: TOPCItem; AIndex: integer): TOPCItem;
     function  CreatePivotTable: TOPCItem;

     function  WritePrinterSettings(AOwner: TOPCItem; AStream: TStream): AxUCString;

     function  FindSheetItemTarget(ASheetId,AItemId: AxUCString): AxUCString;

     property Workbook: TOPCItem read FWorkbook;
     property CurrSheet: TOPCItem read FCurrSheet;
     property CurrDrawing: TOPCItem read FCurrDrawing;
     end;

type TOPC_DOCX = class(TOPCContainer)
protected
     FResultList          : TOPCItemList;

     FDocument            : TOPCItem;
     FNumbering           : TOPCItem;
     FStyles              : TOPCItem;
public
     constructor Create;
     destructor Destroy; override;

     procedure OpenRead(AStream: TStream; const AUseAlternateZip: boolean); override;
     procedure OpenWrite(AStream: TStream; const AUseAlternateZip: boolean); override;
     procedure Close; override;

     function  FindHeaders: boolean;
     function  FindFooters: boolean;

     function  ReadDocument: TStream;
     function  ReadNumbering: TStream;
     function  ReadSettings: TStream;
     function  ReadTheme: TStream;
     function  ReadStyles: TStream;
     function  ReadDocImages: boolean;
     function  ReadNumberingImages: boolean;
     function  ReadHeaderFooterImages: boolean;
     function  ReadHyperlinks(const AList: TStrings): boolean;
     function  ReadImage(const ArId: AxUCString; var AName: AxUCString): TStream;
     function  ReadHeader(const ArId: AxUCString): TStream;
     function  ReadFooter(const ArId: AxUCString): TStream;
     function  ReadFootnotes: TStream;
     function  ReadEndnotes: TStream;

     function  AddDocPropsCore(AStream: TStream): TOPCItem;
     function  AddDocPropsApp(AStream: TStream): TOPCItem;
     function  AddSettings(AStream: TStream): TOPCItem;
     function  AddTheme(AStream: TStream): TOPCItem;
     function  AddImage(AIndex: integer; AExt: AxUCString; AStream: TStream): TOPCItem;

     function  CreateDocument: TOPCItem;
     function  CreateStyles: TOPCItem;
     function  CreateNumbering: TOPCItem;
     function  CreateHeader(AIndex: integer): TOPCItem;
     function  CreateFooter(AIndex: integer): TOPCItem;
     function  CreateFootnotes: TOPCItem;
     function  CreateEndnotes: TOPCItem;

     property Document: TOPCItem read FDocument;
     property ResultList: TOPCItemList read FResultList;
     end;

implementation

procedure DebugSave(AFilename: string; AStream: TStream);
var
  FStream: TFileStream;
begin
  FStream := TFileStream.Create(AFilename,fmCreate);
  try
    AStream.Seek(0,soFromBeginning);
    FStream.CopyFrom(AStream,AStream.Size);
  finally
    FStream.Free;
  end;
end;

function XMLExpandPath(const APath: AxUCString): AxUCString;
var
  S,S1,S2: AxUCString;
  p: integer;
begin
  S := APath;
  p := Pos('../',S);
  while p > 1 do begin
    S1 := Copy(S,1,p - 2);
    S2 := Copy(S,p + 2,MAXINT);
    p := RCPos('/',S1);
    if p > 1 then
      S1 := Copy(S1,1,p - 1)
    else
      S1 := '';
    S := S1 + S2;

    p := Pos('../',S);
  end;

  if (S <> '') and (S[1] = '/') then
    S := Copy(S,2,MAXINT);

  Result := S;
end;

{ TOPCItem }

function TOPCItem.AbsolutePath: AxUCString;
var
  p: integer;
begin
  Result := AbsoluteTarget;
  p := RCPos('/',Result);
  if p > 1 then
    Result := Copy(Result,1,p - 1);
end;

function TOPCItem.AbsoluteTarget: AxUCString;
var
  S: AxUCString;
  OItem: TOPCItem;

function NormalizePath(const APath: AxUCString): AxUCString;
begin
  if APath <> '' then begin
    Result := APath;
    if Result[Length(Result)] = '/' then
      Result := Copy(Result,1,Length(Result) - 1);
    if Result[1] = '/' then
      Result := Copy(Result,2,MAXINT);
  end
  else
    Result := '';
end;

begin
  // Some none Excel files has absolute paths.
  if Copy(FPath,1,1) = '/' then
    Result := Copy(FPath,2,MAXINT) + '/' + FName
  else begin
    S := NormalizePath(FName);
    OItem := Self;
    while (OItem <> Nil) and (OItem.FName <> '') do begin
      if OItem.FPath <> '' then
        S := NormalizePath(OItem.FPath) + '/' + S;
      OItem := OItem.FParent;
    end;

    Result := XMLExpandPath(S);
  end;
end;

function TOPCItem.Add(ARelationship: TCT_Relationship): TOPCItem;
begin
{$ifdef _AXOLOT_DEBUG}
  if FChilds.IndexOf(ARelationship.Id) >= 0 then
    raise XLSRWException.Create('Duplicate id in OPC');
{$endif}
  Result := TOPCItem.Create(FOwner,Self);
  if Integer(ARelationship.TargetMode) <> XPG_UNKNOWN_ENUM then
    Result.TargetMode := TOPCTargetMode(ARelationship.TargetMode);
  Result.Type_ := ARelationship.Type_;
  Result.Target := ARelationship.Target;
  Result.Text := ARelationship.Content;

  FChilds.AddObject(ARelationship.Id,Result);
  Result.ReadChilds;
end;

function TOPCItem.AddNew(AType, ATarget: AxUCString): TOPCItem;
begin
  Result := FindByTarget(ATarget);
  if Result <> Nil then
    Exit;

  Result := TOPCItem.Create(FOwner,Self);
  Result.Type_ := AType;
  Result.Target := ATarget;
  FChilds.AddObject(MakeId,Result);

  Result.FState := opcsWrite;
end;

function TOPCItem.AddNewExternal(AType, ATarget: AxUCString): TOPCItem;
begin
  Result := TOPCItem.Create(FOwner,Self);
  Result.TargetMode := otmExternal;
  Result.Type_ := AType;
  Result.Target := ATarget;
  FChilds.AddObject(MakeId,Result);

  Result.FState := opcsWrite;
end;

procedure TOPCItem.Checkout;
begin
  FCheckedOut := True;
end;

procedure TOPCItem.CheckoutAll;

procedure CheckOutChilds(AOPC: TOPCItem);
var
  i: integer;
begin
  AOPC.FCheckedOut := True;
  for i := 0 to AOPC.Count - 1 do
    CheckOutChilds(AOPC[i]);
end;

begin
  CheckOutChilds(Self);
end;

procedure TOPCItem.Clear;
var
  i: integer;
begin
  if FData <> Nil then
    FData.Free;

  for i := 0 to FChilds.Count - 1 do begin
    if FChilds.Objects[i] <> Nil then
      FChilds.Objects[i].Free;
  end;

  FChilds.Clear;
end;

procedure TOPCItem.Close;
begin
  FState := opcsClosed;
end;

function TOPCItem.Count: integer;
begin
  Result := FChilds.Count;
end;

constructor TOPCItem.Create(AOwner: TOPCContainer; AParent: TOPCItem);
begin
  FOwner := AOwner;
  FParent := AParent;
  FTargetMode := otmInternal;
  if AParent <> Nil then
    FIndex := AParent.FChilds.Count
  else
    FIndex := -1;
{$ifdef DELPHI_5}
  FChilds := TStringList.Create;
{$else}
  FChilds := THashedStringList.Create;
{$endif}
end;

destructor TOPCItem.Destroy;
begin
  Clear;

  FChilds.Free;
  inherited;
end;

function TOPCItem.FindByType(ATypeName: AxUCString; AList: TOPCItemList; AClearList: boolean = True): boolean;
var
  i: integer;
begin
  if AClearList then
    AList.Clear;

  for i := 0 to FChilds.Count - 1 do begin
    if Childs[i].Type_ = ATypeName then
      AList.Add(Childs[i]);
  end;
  Result := AList.Count > 0;
end;

function TOPCItem.Find(AId: AxUCString): TOPCItem;
var
  i: integer;
begin
  i := FChilds.IndexOf(AId);
  if i >= 0 then
    Result := TOPCItem(FChilds.Objects[i])
  else
    Result := Nil;
end;

function TOPCItem.FindByTarget(ATarget: AxUCString): TOPCItem;
var
  i: integer;
begin
  for i := 0 to FChilds.Count - 1 do begin
    if TOPCItem(FChilds.Objects[i]).Target = ATarget then begin
      Result := TOPCItem(FChilds.Objects[i]);
      Exit;
    end;
  end;
  Result := Nil;
end;

function TOPCItem.FindByType(ATypeName: AxUCString): TOPCItem;
var
  List: TOPCItemList;
begin
  Result := Nil;
  List := TOPCItemList.Create;
  try
    if FindByType(ATypeName,List) then begin
      if List.Count > 1 then
        raise XLSRWException.CreateFmt('OPC: only one item of type %s excpected',[ATypeName]);
      Result := List[0];
    end;
  finally
    List.Free;
  end;
end;

function TOPCItem.FindTarget(AId: AxUCString): AxUCString;
var
  i: integer;
  Item: TOPCItem;
begin
  i := FChilds.IndexOf(AId);
  if i >= 0 then begin
    Item := TOPCItem(FChilds.Objects[i]);
    Result := Item.Target;
  end
  else
    Result := '';
end;

function TOPCItem.GetId: AxUCString;
begin
  Result := FParent.FChilds[FIndex];
end;

function TOPCItem.GetItems(Index: integer): TOPCItem;
begin
  Result := TOPCItem(FChilds.Objects[Index]);
end;

function TOPCItem.GetTarget: AxUCString;
begin
  Result := FPath;
  if Result <> '' then
    Result := Result + '/' + FName
  else
    Result := FName;
end;

function TOPCItem.IsRoot: boolean;
begin
  Result := FType = OPC_ROOT;
end;

function TOPCItem.MakeFilename_rels: AxUCString;
begin
  Result := Format('%s/_rels/%s.rels',[AbsolutePath,Name]);

  if (Length(Result) > 1) and (Result[1] = '/') then
    Result := Copy(Result,2,MAXINT);
end;

function TOPCItem.MakeId: AxUCString;
var
  n: integer;
begin
  n := FChilds.Count + 1;
  Result := 'rId' + IntToStr(n);
{$ifdef _AXOLOT_DEBUG}
  while FChilds.IndexOf(Result) >= 0 do begin
    Inc(n);
    Result := 'rId' + IntToStr(n);
  end;
{$endif}
end;

procedure TOPCItem.ReadChilds;
var
  i: integer;
  FileName: AxUCString;
  XML: TXPGDocOPC;
  Stream: TStream;
begin
  Filename := MakeFilename_rels;

  i := FOwner.FZIP.Find(Filename);
  if i < 0 then
    Exit;
//    raise XLSRWException.CreateFmt('Can not find OPC item "%s"',[Filename]);

  Stream := FOwner.FZIP.OpenStream(FileName);
  try
    XML := TXPGDocOPC.Create;
    try
      XML.LoadFromStream(Stream);

      for i := 0 to XML.Relationships.RelationshipXpgList.Count - 1 do
        Add(XML.Root.Relationships.RelationshipXpgList[i]);

    finally
      XML.Free;
    end;
  finally
    Stream.Free;
  end;
end;

procedure TOPCItem.SetTarget(const Value: AxUCString);
var
  p: integer;
begin
  if FTargetMode = otmExternal then
    FName := Value
  else begin
    p := RCPos('/',Value);
    if p > 0 then begin
      FPath := Copy(Value,1,p - 1);
      FName := Copy(Value,p + 1,MAXINT);
    end
    else begin
      FPath := '';
      FName := Value;
    end;
  end;
end;

procedure TOPCItem.WriteChilds;
var
  i: integer;
  XML: TXPGDocOPC;
  Rel: TCT_Relationship;
  Stream: TMemoryStream;
begin
  if Count < 1 then
    Exit;

  XML := TXPGDocOPC.Create;
  if FOwner.FXmlns <> '' then
    XML.Root.RootAttributes.AddNameValue('xmlns',FOwner.FXmlns);
  try
    for i := 0 to Count - 1 do begin
      Rel := XML.Relationships.RelationshipXpgList.Add;
      Rel.Id := Childs[i].Id;
      Rel.Type_ := Childs[i].Type_;
      if Childs[i].TargetMode <> otmInternal then
        Rel.TargetMode := TST_TargetMode(Childs[i].TargetMode);
      Rel.Target := Childs[i].Target;
    end;

    Stream := TMemoryStream.Create;
    try
      XML.SaveToStream(Stream);
      FOwner.FZIP.Write(MakeFilename_rels,Stream);
    finally
      Stream.Free;
    end;

    for i := 0 to Count - 1 do
      Childs[i].WriteChilds;
  finally
    XML.Free;
  end;
end;

{ TOPCContainer }

procedure TOPCContainer.Clear;
begin
  FRoot.Clear;
  FContentTypes.Clear;
end;

procedure TOPCContainer.ItemClose(AItem: TOPCItem);
begin
  if not (AItem.FState in [opcsRead,opcsWrite]) then
    raise XLSRWException.Create('Close on not open OPC Item.');
  AItem.FState := opcsClosed;
end;

procedure TOPCContainer.ItemCloseStream(AItem: TOPCItem; AStream: TStream);
begin
  FZIP.CloseStream(AStream);
end;

procedure TOPCContainer.Close;
begin
  case FState of
    opcsClosed: ;
    opcsRead: ;
    opcsWrite: begin
      WriteContentTypes;
      FRoot.WriteChilds;
      FZIP.Close;
    end;
  end;
  FState := opcsClosed;
end;

procedure TOPCContainer.CollectUnchecked(AList: TOPCItemList);

procedure DoChilds(AItem: TOPCItem);
var
  i: integer;
begin
  if not AItem.CheckedOut and (AItem.Name <> '') then
    AList.Add(AItem);
  for i := 0 to AItem.Count - 1 do
    DoChilds(AItem[i]);
end;

begin
  DoChilds(FRoot);
end;

constructor TOPCContainer.Create;
begin
  FRoot := TOPCItem.Create(Self,Nil);
  FRoot.Type_ := OPC_ROOT;
  FContentTypes := TOPC_ContentTypes.Create;
end;

procedure TOPCContainer.DebugDump(AFileName: AxUCString);
var
  Lines: TStringList;

procedure Dump(AItem: TOPCItem; AIndent: integer);
var
  S: string;
  i: integer;
begin
  SetLength(S,AIndent);
  for i := 1 to AIndent do
    S[i] := ' ';
  for i := 0 to AItem.Count - 1  do begin
    Lines.Add(S + AItem[i].Target + ' (' + AItem[i].AbsoluteTarget + ')');
    if AItem[i].Count > 0 then
      Dump(AItem[i],AIndent + 2);
  end;
end;

begin
  Lines := TStringList.Create;
  try
    Dump(FRoot,0);

    Lines.SaveToFile(AFileName);
  finally
    Lines.Free;
  end;
end;

destructor TOPCContainer.Destroy;
begin
  if FZIP <> Nil then
    FZIP.Free;
  FRoot.Free;
  FContentTypes.Free;
  inherited;
end;

function TOPCContainer.ItemCreate(AOwner: TOPCItem; AType, ATarget, AContentType: AxUCString): TOPCItem;
begin
  Result := AOwner.AddNew(AType,ATarget);
  Result.FState := opcsWrite;
  if AContentType <> '' then
    FContentTypes.AddOverride(Result.AbsoluteTarget,AContentType);
end;

function TOPCContainer.ItemCreateExternal(AOwner: TOPCItem; AType, ATarget: AxUCString): TOPCItem;
begin
  Result := AOwner.AddNewExternal(AType,ATarget);
  Result.FState := opcsWrite;
end;

function TOPCContainer.ItemCreateStream(AItem: TOPCItem): TStream;
begin
  if AItem.State <> opcsWrite then
    raise XLSRWException.Create('OPC Item is not opened for writing.');
  Result := FZIP.CreateStream(AItem.AbsolutePath,AItem.Name);
end;

function TOPCContainer.ItemOpenRead(AOwner: TOPCItem; AType, AId: AxUCString): TOPCItem;
var
  i: integer;
begin
  i := AOwner.FChilds.IndexOf(AId);
  if i >= 0 then begin
    Result := TOPCItem(AOwner.FChilds.Objects[i]);
    if Result.Type_ <> AType then
      raise XLSRWException.CreateFmt('OPC Item is not of excpected type "%s"',[AType]);
    Result.FState := opcsRead;
  end
  else
    Result := Nil;
end;

function TOPCContainer.ItemOpenStream(AItem: TOPCItem): TStream;
begin
  Result := FZIP.OpenStream(AItem.AbsoluteTarget);
  if FAutoCheckout then
    AItem.Checkout;
end;

function TOPCContainer.ItemOpenRead(AOwner: TOPCItem; AType: AxUCString): TOPCItem;
begin
  Result := AOwner.FindByType(AType);
  if Result <> Nil then
    Result.FState := opcsRead;
end;

procedure TOPCContainer.ItemWrite(AItem: TOPCItem; AStream: TStream);
begin
  if AItem.State <> opcsWrite then
    raise XLSRWException.Create('OPC Item is not opened for writing.');
  FZIP.Write(AItem.AbsolutePath,AItem.Name,AStream);
end;

procedure TOPCContainer.OpenRead(AStream: TStream; const AUseAlternateZip: boolean);
begin
  if FState <> opcsClosed then
    raise XLSRWException.Create('OPC is allready open');

  if FZIP <> Nil then
    FZIP.Free;
  if AUseAlternateZip then
{$ifdef DELPHI_XE2_OR_LATER}
    FZIP := TXLSZipArchiveIntfDelphi.Create
{$else}
    raise XLSRWException.Create('No alternate Zip defined.')
{$endif}
  else
    FZIP := TXLSZipArchiveIntfXLS.Create;

  Clear;

  FState := opcsRead;

  FZIP.LoadFromStream(AStream);

  ReadContentTypes;

  FRoot.ReadChilds;
end;

procedure TOPCContainer.OpenWrite(AStream: TStream; const AUseAlternateZip: boolean);
begin
  if FZIP <> Nil then
    FZIP.Free;
  if AUseAlternateZip then
{$ifdef DELPHI_XE2_OR_LATER}
    FZIP := TXLSZipArchiveIntfDelphi.Create
{$else}
    raise XLSRWException.Create('No alternate Zip defined.')
{$endif}
  else
    FZIP := TXLSZipArchiveIntfXLS.Create;

  FState := opcsWrite;
  FZIP.SaveToStream(AStream);
end;

function TOPCContainer.ReadContentType: TStream;
var
  i: integer;
begin
  i := FZIP.Find(OPC_DIR_CONTENTTYPES);
  if i >= 0 then
    Result := FZIP.OpenStream(OPC_DIR_CONTENTTYPES)
  else
    Result := Nil;
end;

procedure TOPCContainer.ReadContentTypes;
var
  i: integer;
  XML: TXPGDocContentType;
  Stream: TStream;
begin
  XML := TXPGDocContentType.Create;
  try
    Stream := FZIP.OpenStream(OPC_DIR_CONTENTTYPES);
    try
      XML.LoadFromStream(Stream);

      for i := 0 to XML.Root.Types.DefaultXpgList.Count - 1 do
        FContentTypes.AddDefault(XML.Root.Types.DefaultXpgList[i].Extension,XML.Root.Types.DefaultXpgList[i].ContentType);
      for i := 0 to XML.Root.Types.OverrideXpgList.Count - 1 do
        FContentTypes.AddOverride(XML.Root.Types.OverrideXpgList[i].PartName,XML.Root.Types.OverrideXpgList[i].ContentType);
    finally
      Stream.Free;
    end;
  finally
    XML.Free;
  end;
end;

procedure TOPCContainer.SaveItemData(AItem: TOPCItem);
var
  i: integer;
  S: AxUCString;
  Stream: TStream;
begin
  S := AItem.AbsoluteTarget;
  i := FZIP.Find(S);
  if i < 0 then
    Exit;
//    raise XLSRWException.CreateFmt('Can not find "%s" in OPC',[AItem.AbsoluteTarget]);

  try
    Stream := FZIP.OpenStream(S);
    try
      // This can be optimized by saving to a compression stream insted, but
      // there shall not be large ammounts of data here so it's probably
      // screaming over very little wool.
      AItem.Data := TMemoryStream.Create;
      AItem.Data.CopyFrom(Stream,Stream.Size);
    finally
      Stream.Free;
    end;

    S := '/' + S;
    i := FContentTypes.FindOverride(S);
    if i >= 0 then
      AItem.FContent := TOPC_ContentTypeOverride(FContentTypes[i]).ContentType;
  except
    AItem.Data.Free;
    AItem.Data := Nil;
  end;
end;

procedure TOPCContainer.SaveUnchecked(AList: TOPCItemList);
var
  i: integer;
begin
  CollectUnchecked(AList);
  for i := 0 to AList.Count - 1 do begin
    if AList[i].TargetMode = otmInternal then
      SaveItemData(AList[i]);
  end;
end;

procedure TOPCContainer.WriteContentTypes;
var
  i: integer;
  XML: TXPGDocContentType;
  CtOverride: TCT_Override;
  CtDefault: TCT_Default;
  Stream: TMemoryStream;
begin
  XML := TXPGDocContentType.Create;
  XML.Root.RootAttributes.AddNameValue('xmlns',OOXML_URI_PACKAGE_CONTENTTYPES);
  try
    for i := 0 to FContentTypes.Count - 1 do begin
      if FContentTypes[i] is TOPC_ContentTypeOverride then begin
        CtOverride := XML.Types.OverrideXpgList.Add;
        CtOverride.ContentType := FContentTypes[i].ContentType;
        CtOverride.PartName := TOPC_ContentTypeOverride(FContentTypes[i]).PartName;
      end
      else begin
        CtDefault := XML.Types.DefaultXpgList.Add;
        CtDefault.ContentType := FContentTypes[i].ContentType;
        CtDefault.Extension := TOPC_ContentTypeDefault(FContentTypes[i]).Extension;
      end;
    end;

    Stream := TMemoryStream.Create;
    try
      XML.SaveToStream(Stream);
      FZIP.Write('',OPC_DIR_CONTENTTYPES,Stream);
    finally
      Stream.Free;
    end;
  finally
    XML.Free;
  end;
end;

{ TOPC_XLSX }

function TOPC_XLSX.AddDocPropsApp(AStream: TStream): TOPCItem;
begin
  Result := FRoot.AddNew(OPC_XLSX_DOCPROPS_APP,OPC_XLSX_DOCPROPS_APP_TARGET);
  FZIP.Write(Result.AbsolutePath,Result.Name,AStream);
  FContentTypes.AddOverride(Result.AbsoluteTarget,'application/vnd.openxmlformats-officedocument.extended-properties+xml');
end;

function TOPC_XLSX.AddDocPropsCore(AStream: TStream): TOPCItem;
begin
  Result := FRoot.AddNew(OPC_XLSX_DOCPROPS_CORE,OPC_XLSX_DOCPROPS_CORE_TARGET);
  FZIP.Write(Result.AbsolutePath,Result.Name,AStream);
  FContentTypes.AddOverride(Result.AbsoluteTarget,'application/vnd.openxmlformats-package.core-properties+xml');
end;

function TOPC_XLSX.AddTheme(AStream: TStream; AIndex: integer): TOPCItem;
var
  S: string;
begin
  S := Format(OPC_XLSX_THEME_TARGET,[AIndex]);
  Result := FWorkbook.AddNew(OPC_XLSX_THEME,S);
  FZIP.Write(Result.AbsolutePath,Result.Name,AStream);
  FContentTypes.AddOverride(Result.AbsoluteTarget,'application/vnd.openxmlformats-officedocument.theme+xml');
end;

procedure TOPC_XLSX.CheckWorkbookOpen;
begin
  if FWorkbook = Nil then
    raise XLSRWException.Create('OPC Workbook is not open.');
//  if FWorkbook.State <> opcsWrite then
//    raise XLSRWException.Create('OPC Workbook is not opened for writing');
end;

procedure TOPC_XLSX.Clear;
begin
  inherited Clear;
  FResultList.Clear;
  FWorkbook := Nil;
end;

procedure TOPC_XLSX.Close;
begin
  inherited Close;

  case FState of
    opcsClosed: ;
    opcsRead: ;
    opcsWrite: begin
      CheckWorkbookOpen;
      ItemClose(FWorkbook);
    end;
  end;

  FWorkbook := Nil;
end;

procedure TOPC_XLSX.CloseChart(AItem: TOPCItem);
begin
  ItemClose(AItem);

  FCurrChart := Nil;
end;

procedure TOPC_XLSX.CloseDrawing;
begin
  ItemClose(FCurrDrawing);
  FCurrDrawing := Nil;
end;

procedure TOPC_XLSX.CloseSheet;
begin
  ItemClose(FCurrSheet);
  FCurrSheet := Nil;
end;

procedure TOPC_XLSX.CloseTable;
begin
  ItemClose(FCurrTable);
  FCurrTable := Nil;
end;

constructor TOPC_XLSX.Create;
begin
  inherited Create;

  FContentTypes.XLSXMode := True;
  FContentTypes.SetDefault;

  FAutoCheckout := True;
  FXmlns := OOXML_URI_PACKAGE_RELATIONSHIPS;
  FResultList := TOPCItemList.Create;
end;

function TOPC_XLSX.CreateChart: TOPCItem;
var
  S: string;
begin
  Inc(FChartCount);
  S := Format(OPC_XLSX_CHART_TARGET,[FChartCount]);
  Result := ItemCreate(FCurrDrawing,OPC_XLSX_CHART,S,'application/vnd.openxmlformats-officedocument.drawingml.chart+xml');
  FCurrChart := Result;
end;

function TOPC_XLSX.CreateChartColors(AOwner: TOPCItem): TOPCItem;
var
  S: string;
begin
  S := Format(OPC_XLSX_CHARTCOLORS_TARGET,[FChartCount]);
  Result := ItemCreate(AOwner,OPC_XLSX_CHARTCOLORS,S,'application/vnd.ms-office.chartcolorstyle+xml');
end;

function TOPC_XLSX.CreateChartSheet(AIndex: integer): TOPCItem;
var
  S: string;
begin
  S := Format(OPC_XLSX_CHARTSHEET_TARGET,[AIndex]);
  Result := ItemCreate(FWorkbook,OPC_XLSX_CHARTSHEET,S,'application/vnd.openxmlformats-officedocument.spreadsheetml.Chartsheet+xml');
  FCurrSheet := Result;
end;

function TOPC_XLSX.CreateChartStyle(AOwner: TOPCItem): TOPCItem;
var
  S: string;
begin
  S := Format(OPC_XLSX_CHARTSTYLE_TARGET,[FChartCount]);
  Result := ItemCreate(AOwner,OPC_XLSX_CHARTSTYLE,S,'application/vnd.ms-office.chartstyle+xml');
end;

function TOPC_XLSX.CreateChartUserShapes(AOwner: TOPCItem; AIndex: integer): TOPCItem;
var
  S: string;
begin
  Inc(FChartCount);
  S := Format(OPC_XLSX_DRAWING_TARGET,[AIndex]);
  Result := ItemCreate(AOwner,OPC_XLSX_CHARTUSERSHAPES,S,'application/vnd.openxmlformats-officedocument.drawingml.chartshapes+xml');
  FCurrDrawing := Result;
end;

function TOPC_XLSX.CreateComment(AOwner: TOPCItem; AIndex: integer): TOPCItem;
var
  S: string;
begin
  S := Format(OPC_XLSX_COMMENTS_TARGET,[AIndex]);
  Result := ItemCreate(AOwner,OPC_XLSX_COMMENTS,S,'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml');
end;

function TOPC_XLSX.CreateConnections: TOPCItem;
begin
  CheckWorkbookOpen;
  Result := ItemCreate(FWorkbook,OPC_XLSX_CONNECTIONS,OPC_XLSX_CONNECTIONS_TARGET,'application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml');
end;

function TOPC_XLSX.CreateDrawing(AOwner: TOPCItem; AIndex: integer): TOPCItem;
var
  S: string;
begin
  S := Format(OPC_XLSX_DRAWING_TARGET,[AIndex]);
  Result := ItemCreate(AOwner,OPC_XLSX_DRAWING,S,'application/vnd.openxmlformats-officedocument.drawing+xml');
  FCurrDrawing := Result;
end;

function TOPC_XLSX.CreateHyperlink(AOwner: TOPCItem; ATarget: AxUCString): TOPCItem;
begin
  Result := ItemCreateExternal(AOwner,OPC_XLSX_HYPERLINK,ATarget);
end;

function TOPC_XLSX.CreateImage(const AIndex: integer; const AImageExt: AxUCString): TOPCItem;
var
  S: string;
begin
  S := Format(OPC_XLSX_IMAGE_TARGET,[AIndex,AImageExt]);
  if FCurrChart <> Nil then
    Result := ItemCreate(FCurrChart,OPC_XLSX_IMAGE,S,'')
  else
    Result := ItemCreate(FCurrDrawing,OPC_XLSX_IMAGE,S,'');
end;

function TOPC_XLSX.CreatePivotCacheDefinition(AOwner: TOPCItem; AIndex: integer): TOPCItem;
var
  S: string;
begin
  if AOwner = Nil then
    AOwner := FWorkbook;
  S := Format(OPC_XLSX_PIVOTCACHEDEF_TARGET,[AIndex]);
  Result := ItemCreate(AOwner,OPC_XLSX_PIVOTCACHEDEF,S,'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml');
end;

function TOPC_XLSX.CreatePivotCacheRecords(AOwner: TOPCItem; AIndex: integer): TOPCItem;
var
  S: string;
begin
  S := Format(OPC_XLSX_PIVOTCACHERECS_TARGET,[AIndex]);
  Result := ItemCreate(AOwner,OPC_XLSX_PIVOTCACHERECS,S,'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml');
end;

function TOPC_XLSX.CreatePivotTable: TOPCItem;
var
  S: string;
begin
  Inc(FPivotTableCount);
  S := Format(OPC_XLSX_PIVOTTABLE_TARGET,[FPivotTableCount]);
  Result := ItemCreate(FCurrSheet,OPC_XLSX_PIVOTTABLE,S,'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml');
end;

function TOPC_XLSX.CreateQueryTable(AOwner: TOPCItem): TOPCItem;
var
  S: string;
begin
  Inc(FQueryTableCount);
  S := Format(OPC_XLSX_QUERYTABLE_TARGET,[FQueryTableCount]);
  Result := ItemCreate(AOwner,OPC_XLSX_QUERYTABLE,S,'application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml');
end;

destructor TOPC_XLSX.Destroy;
begin
  FResultList.Free;
  inherited;
end;

function TOPC_XLSX.FindAndOpenChart(AId: AxUCString): TOPCItem;
begin
  Result := ItemOpenRead(FCurrDrawing,OPC_XLSX_CHART,AId);
  if Result = Nil then
    raise XLSRWException.Create('Can not find Table in OPC');
end;

function TOPC_XLSX.FindDrawingId(AOPCSheet: TOPCItem): AxUCString;
var
  Item: TOPCItem;
begin
  Item := AOPCSheet.FindByType(OPC_XLSX_DRAWING);
  if Item <> Nil then
    Result := Item.Id
  else
    Result := '';
end;

function TOPC_XLSX.FindPrinterSettingsId(AOPCSheet: TOPCItem): AxUCString;
var
  Item: TOPCItem;
begin
  Item := AOPCSheet.FindByType(OPC_XLSX_PRINTER_SETTINGS);
  if Item <> Nil then
    Result := Item.Id
  else
    Result := '';
end;

function TOPC_XLSX.FindPictureId(AOPCSheet: TOPCItem; AIndex: integer): AxUCString;
var
  Item: TOPCItem;
begin
  Item := AOPCSheet.FindByTarget(Format(OPC_XLSX_VMLDRAWING_TARGET,[AIndex]));
  if Item <> Nil then
    Result := Item.Id
  else
    Result := '';
end;

function TOPC_XLSX.FindSheet(AIndex: integer): TOPCItem;
begin
  Result := FWorkbook.FindByTarget(Format(OPC_XLSX_WORKSHEET_TARGET,[AIndex]));
end;

function TOPC_XLSX.FindSheetBin(AIndex: integer): TOPCItem;
begin
  Result := FWorkbook.FindByTarget(Format(OPC_XLSX_WORKSHEET_TARGET_BIN,[AIndex]));
end;

function TOPC_XLSX.FindSheetItemTarget(ASheetId,AItemId: AxUCString): AxUCString;
var
  i: integer;
  Item: TOPCItem;
begin
  i := FWorkbook.FChilds.IndexOf(ASheetId);
  if i < 0 then
    raise XLSRWException.Create('Can not find worksheet in OPC');

  Item := TOPCItem(FWorkbook.Childs[i]);
  Result := Item.FindTarget(AItemId);
end;

function TOPC_XLSX.FindTableId(AOPCSheet: TOPCItem; ATarget: AxUCString): AxUCString;
var
  Item: TOPCItem;
begin
  Item := AOPCSheet.FindByTarget(ATarget);
  if Item <> Nil then
    Result := Item.Id
  else
    Result := '';
end;

function TOPC_XLSX.FindVmlDrawingId(AOPCSheet: TOPCItem): AxUCString;
var
  Item: TOPCItem;
begin
  Item := AOPCSheet.FindByType(OPC_XLSX_VMLDRAWING);
  if Item <> Nil then
    Result := Item.Id
  else
    Result := '';
end;

function TOPC_XLSX.GetDrawingImages(AList: TStrings): boolean;
var
  i: integer;
begin
  for i := 0 to FCurrDrawing.Count - 1 do begin
    if FCurrDrawing[i].Type_ = OPC_XLSX_IMAGE then
      AList.Add(FCurrDrawing[i].Id);
  end;
  Result := AList.Count > 0;
end;

function TOPC_XLSX.GetSheetPivotTables(AList: TStrings): boolean;
var
  i: integer;
begin
  for i := 0 to FCurrSheet.Count - 1 do begin
    if FCurrSheet[i].Type_ = OPC_XLSX_PIVOTTABLE then
      AList.Add(FCurrSheet[i].Id);
  end;
  Result := AList.Count > 0;
end;

procedure TOPC_XLSX.OpenRead(AStream: TStream; const AUseAlternateZip: boolean);
begin
  inherited OpenRead(AStream,AUseAlternateZip);

  FResultList.Clear;
  if not FRoot.FindByType(OPC_XLSX_WORKBOOK,FResultList) then
    raise XLSRWException.Create('Can not find workbook in OPC');
  if FResultList.Count > 1 then
    raise XLSRWException.Create('More than one workbook in OPC');
  FWorkbook := FResultList[0];
  FResultList.Clear;

  if FRoot.FindByType(OPC_XLSX_CALCCHAIN,FResultList) then
    FResultList[0].Checkout;
end;

procedure TOPC_XLSX.OpenWrite(AStream: TStream; const AUseAlternateZip: boolean);
begin
  inherited OpenWrite(AStream,AUseAlternateZip);
  FPrinterSettingsCount := 0;
end;

function TOPC_XLSX.ReadAny(AItem: TOPCItem): TStream;
begin
  AItem.FState := opcsRead;
  Result := ReadItemData(AItem);
  ItemClose(AItem);
end;

function TOPC_XLSX.ReadChart(AId: AxUCString): TStream;
var
  Item: TOPCItem;
begin
  Item := ItemOpenRead(FCurrDrawing,OPC_XLSX_CHART,AId);
  if Item = Nil then
    raise XLSRWException.Create('Can not find Chart in OPC');
  Result := ReadItemData(Item);
  ItemClose(Item);
end;

function TOPC_XLSX.ReadChartColors(AChart: TOPCItem; AId: AxUCString): TStream;
var
  Item: TOPCItem;
begin
  Item := ItemOpenRead(AChart,OPC_XLSX_CHARTCOLORS,AId);
  if Item = Nil then
    raise XLSRWException.Create('Can not find Table in OPC');
  Result := ReadItemData(Item);
  ItemClose(Item);
end;

function TOPC_XLSX.ReadChartStyle(AChart: TOPCItem; AId: AxUCString): TStream;
var
  Item: TOPCItem;
begin
  Item := ItemOpenRead(AChart,OPC_XLSX_CHARTSTYLE,AId);
  if Item = Nil then
    raise XLSRWException.Create('Can not find Table in OPC');
  Result := ReadItemData(Item);
  ItemClose(Item);
end;

function TOPC_XLSX.ReadComments(ASheetId: AxUCString): TStream;
var
  Item: TOPCItem;
begin
  Result := Nil;
  Item := ItemOpenRead(FWorkbook,OPC_XLSX_WORKSHEET,ASheetId);
  if Item <> Nil then begin
    ItemClose(Item);
    Item := ItemOpenRead(Item,OPC_XLSX_COMMENTS);
    if Item <> Nil then begin
      Result := ReadItemData(Item);
      ItemClose(Item);
    end;
  end;
end;

function TOPC_XLSX.ReadConnections: TStream;
var
  Item: TOPCItem;
begin
  Item := ItemOpenRead(FWorkbook,OPC_XLSX_CONNECTIONS);
  if Item <> Nil then begin
    Result := ReadItemData(Item);
    ItemClose(Item);
  end
  else
    Result := Nil;
end;

function TOPC_XLSX.ReadDocPropsApp: TStream;
var
  Item: TOPCItem;
begin
  Item := ItemOpenRead(FRoot,OPC_XLSX_DOCPROPS_APP);
  if Item <> Nil then begin
    Result := ReadItemData(Item);
    ItemClose(Item);
  end
  else
    Result := Nil;
end;

function TOPC_XLSX.ReadDocPropsCore: TStream;
var
  Item: TOPCItem;
begin
  Item := ItemOpenRead(FRoot,OPC_XLSX_DOCPROPS_CORE);
  if Item <> Nil then begin
    Result := ReadItemData(Item);
    ItemClose(Item);
  end
  else
    Result := Nil;
end;

function TOPC_XLSX.ReadDrawingMedia(AId: AxUCString; out ATarget: AxUCString): TStream;
var
  Item: TOPCItem;
begin
  Item := ItemOpenRead(FCurrDrawing,OPC_XLSX_IMAGE,AId);
  if Item.TargetMode = otmInternal then begin
    if Item = Nil then
      raise XLSRWException.Create('Can not find Image in OPC');
    Result := ReadItemData(Item);
    ATarget := Item.Target;
  end
  else begin
    ATarget := '';
    Result := Nil;
  end;
  ItemClose(Item);
end;

function TOPC_XLSX.ReadItemData(AItem: TOPCItem): TStream;
begin
  Result := ItemOpenStream(AItem);
end;

function TOPC_XLSX.ReadPivotCacheDefinition(AId: AxUCString): TStream;
var
  OPC: TOPCItem;
begin
  OPC := ItemOpenRead(FWorkbook,OPC_XLSX_PIVOTCACHEDEF,AId);
  if OPC = Nil then
    raise XLSRWException.CreateFmt('Can not find pivotCacheDefintion %s in OPC',[AId]);
  Result := ReadItemData(OPC);
  ItemClose(OPC);
end;

function TOPC_XLSX.ReadPivotCacheRecords(const AOwnerId, AId: AxUCString): TStream;
var
  OPC: TOPCItem;
  OPCOwner: TOPCItem;
begin
  OPCOwner := ItemOpenRead(FWorkbook,OPC_XLSX_PIVOTCACHEDEF,AOwnerId);
  if OPCOwner <> Nil then begin
    OPC := ItemOpenRead(OPCOwner,OPC_XLSX_PIVOTCACHERECS,AId);
    if OPC = Nil then
      raise XLSRWException.CreateFmt('Can not find pivotCacheRecords %s in OPC',[AId]);
    Result := ReadItemData(OPC);
    ItemClose(OPC);
  end
  else
    raise XLSRWException.CreateFmt('Can not find pivotCacheDefintion %s in OPC',[AOwnerId]);
end;

function TOPC_XLSX.ReadPivotTable(AId: AxUCString): TStream;
var
  OPC: TOPCItem;
begin
  OPC := ItemOpenRead(FCurrSheet,OPC_XLSX_PIVOTTABLE,AId);
  if OPC = Nil then
    raise XLSRWException.CreateFmt('Can not find pivotTable %s in OPC',[AId]);
  Result := ReadItemData(OPC);
  // The PivotTable childs, RecordDefinition and RecordCache is owned by the workbook.
  OPC.CheckoutAll;
  ItemClose(OPC);
end;

function TOPC_XLSX.ReadPrinterSettings(AId: AxUCString): TStream;
var
  Item: TOPCItem;
begin
  Result := Nil;
  Item := ItemOpenRead(FCurrSheet,OPC_XLSX_PRINTER_SETTINGS,AId);
  if Item <> Nil then begin
    Result := ReadItemData(Item);
    ItemClose(Item);
  end;
end;

function TOPC_XLSX.ReadQueryTable: TStream;
var
  Item: TOPCItem;
begin
  // Don't search by rId as the rId not is stored in the Table.
  Item := ItemOpenRead(FCurrTable,OPC_XLSX_QUERYTABLE);
  if Item <> Nil then begin
    Result := ReadItemData(Item);
    ItemClose(Item);
  end
  else
    Result := Nil;
end;

function TOPC_XLSX.OpenAndReadChartSheet(AId: AxUCString): TStream;
begin
  FCurrSheet := ItemOpenRead(FWorkbook,OPC_XLSX_CHARTSHEET,AId);
  if FCurrSheet = Nil then
    raise XLSRWException.Create('Can not find worksheet in OPC');

  Result := ReadItemData(FCurrSheet);
end;

function TOPC_XLSX.OpenAndReadDrawing(AId: AxUCString): TStream;
begin
  FCurrDrawing := ItemOpenRead(FCurrSheet,OPC_XLSX_DRAWING,AId);
  if FCurrDrawing = Nil then
    raise XLSRWException.Create('Can not find Drawing in OPC');
  Result := ReadItemData(FCurrDrawing);
end;

function TOPC_XLSX.OpenAndReadSheet(AId: AxUCString): TStream;
begin
  FCurrSheet := ItemOpenRead(FWorkbook,OPC_XLSX_WORKSHEET,AId);
  if FCurrSheet = Nil then
    raise XLSRWException.Create('Can not find worksheet in OPC');

  Result := ReadItemData(FCurrSheet);
end;

function TOPC_XLSX.ReadVmlDrawing(AId: AxUCString; ARels: TStrings): TStream;
var
  i  : integer;
  OPC: TOPCItem;
begin
  OPC := ItemOpenRead(FCurrSheet,OPC_XLSX_VMLDRAWING,AId);
  if OPC = Nil then
    raise XLSRWException.CreateFmt('Can not find VmlDrawing %s in OPC',[AId]);
  Result := ReadItemData(OPC);

  for i := 0 to OPC.Count - 1 do
    ARels.Add(OPC[i].Type_ + '&' + OPC[i].Target);

  ItemClose(OPC);
end;

function TOPC_XLSX.ReadSST: TStream;
var
  Item: TOPCItem;
begin
  Item := ItemOpenRead(FWorkbook,OPC_XLSX_SST);
  if Item <> Nil then begin
    Result := ReadItemData(Item);
    ItemClose(Item);
  end
  else
    Result := Nil;
end;

function TOPC_XLSX.ReadStyles: TStream;
var
  Item: TOPCItem;
begin
  Item := ItemOpenRead(FWorkbook,OPC_XLSX_STYLES);
  if Item <> Nil then begin
    Result := ReadItemData(Item);
    ItemClose(Item);
  end
  else
    Result := Nil;
end;

function TOPC_XLSX.ReadTable(AId: AxUCString): TStream;
begin
  FCurrTable := ItemOpenRead(FCurrSheet,OPC_XLSX_TABLE,AId);
  if FCurrTable = Nil then
    raise XLSRWException.Create('Can not find Table in OPC');
  Result := ReadItemData(FCurrTable);
end;

function TOPC_XLSX.ReadTheme: TStream;
var
  Item: TOPCItem;
begin
  Item := ItemOpenRead(FWorkbook,OPC_XLSX_THEME);
  if Item <> Nil then begin
    Result := ItemOpenStream(Item);
    ItemClose(Item);
  end
  else
    Result := Nil;
end;

function TOPC_XLSX.ReadWorkbook: TStream;
begin
  FWorkbook := ItemOpenRead(FRoot,OPC_XLSX_WORKBOOK);
  if FWorkbook <> Nil then begin
    Result := ReadItemData(FWorkbook);
    ItemClose(FWorkbook);
  end
  else
    Result := Nil;
end;

function TOPC_XLSX.ReadXLink(AId: AxUCString): TStream;
var
  Item: TOPCItem;
begin
  Item := ItemOpenRead(FWorkbook,OPC_XLSX_XLINK,AId);
  if Item = Nil then
    raise XLSRWException.Create('Can not find ExternalLink in OPC');

  Result := ReadItemData(Item);
  ItemClose(Item);
end;

function TOPC_XLSX.SheetIsChartSheet(AId: AxUCString): boolean;
var
  OPC: TOPCItem;
begin
  OPC := FWorkbook.Find(AId);
  if OPC <> Nil then
    Result := OPC.Type_ = OPC_XLSX_CHARTSHEET
  else
    Result := False;
end;

function TOPC_XLSX.WritePrinterSettings(AOwner: TOPCItem; AStream: TStream): AxUCString;
var
  S: string;
  Item: TOPCItem;
begin
  Inc(FPrinterSettingsCount);
  S := Format(OPC_XLSX_PRINTER_SETTINGS_TARGET,[FPrinterSettingsCount]);
  Item := ItemCreate(AOwner,OPC_XLSX_PRINTER_SETTINGS,S,'');
  ItemWrite(Item,AStream);
  Result := Item.Id;
  ItemClose(Item);
end;

function TOPC_XLSX.CreateSheet(AIndex: integer): TOPCItem;
var
  S: string;
begin
  S := Format(OPC_XLSX_WORKSHEET_TARGET,[AIndex]);
  Result := ItemCreate(FWorkbook,OPC_XLSX_WORKSHEET,S,'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml');
  FCurrSheet := Result;
end;

function TOPC_XLSX.CreateSST: TOPCItem;
begin
  CheckWorkbookOpen;
  Result := ItemCreate(FWorkbook,OPC_XLSX_SST,OPC_XLSX_SST_TARGET,'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml');
end;

function TOPC_XLSX.CreateStyles: TOPCItem;
begin
  CheckWorkbookOpen;
  Result := ItemCreate(FWorkbook,OPC_XLSX_STYLES,OPC_XLSX_STYLES_TARGET,'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml');
end;

function TOPC_XLSX.CreateTable(AOwner: TOPCItem): TOPCItem;
var
  S: string;
begin
  Inc(FTableCount);
  S := Format(OPC_XLSX_TABLE_TARGET,[FTableCount]);
  Result := ItemCreate(AOwner,OPC_XLSX_TABLE,S,'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml');
end;

function TOPC_XLSX.CreateVmlDrawing(AOwner: TOPCItem; AIndex: integer): TOPCItem;
var
  S: string;
begin
  S := Format(OPC_XLSX_VMLDRAWING_TARGET,[AIndex]);
  Result := ItemCreate(AOwner,OPC_XLSX_VMLDRAWING,S,'');
end;

function TOPC_XLSX.CreateWorkbook(AMacroEnabled: boolean): TOPCItem;
begin
  if FWorkbook <> Nil then
    raise XLSRWException.Create('OPC workbook is open');
  if AMacroEnabled then
    Result := ItemCreate(FRoot,OPC_XLSX_WORKBOOK,OPC_XLSX_WORKBOOK_TARGET,'application/vnd.ms-excel.sheet.macroEnabled.main+xml')
  else
    Result := ItemCreate(FRoot,OPC_XLSX_WORKBOOK,OPC_XLSX_WORKBOOK_TARGET,'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml');
  FWorkbook := Result;
end;

function TOPC_XLSX.CreateXLink(AIndex: integer): TOPCItem;
var
  S: string;
begin
  S := Format(OPC_XLSX_XLINK_TARGET,[AIndex]);
  Result := ItemCreate(FWorkbook,OPC_XLSX_XLINK,S,'application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml');
end;

function TOPC_XLSX.CreateXLinkPath(AOwner: TOPCItem; AFilename: AxUCString): TOPCItem;
begin
  Result := ItemCreateExternal(AOwner,OPC_XLSX_XLINK_PATH,AFilename);
end;

{ TOPCItemList }

constructor TOPCItemList.Create;
begin
  inherited Create(False);
end;

function TOPCItemList.GetItems(Index: integer): TOPCItem;
begin
  Result := TOPCItem(inherited Items[Index]);
end;

{ TOPC_ContentTypes }

procedure TOPC_ContentTypes.AddDefault(AExtension, AContentType: AxUCString);
var
  CTD: TOPC_ContentTypeDefault;
begin
  if FindDefault(AExtension) >= 0 then
    Exit;
  CTD := TOPC_ContentTypeDefault.Create;
  CTD.Extension := AExtension;
  CTD.ContentType := AContentType;
  inherited Add(CTD);
end;

procedure TOPC_ContentTypes.AddOverride(APartName, AContentType: AxUCString);
var
  CTD: TOPC_ContentTypeOverride;
begin
  if (APartName <> '') and (APartName[1] <> '/') then
    APartName := '/' + APartName;

  if FindOverride(APartName) >= 0 then
    Exit;
  CTD := TOPC_ContentTypeOverride.Create;
  CTD.PartName := APartName;
  CTD.ContentType := AContentType;
  inherited Add(CTD);
end;

procedure TOPC_ContentTypes.Clear;
begin
  inherited Clear;

  SetDefault;
end;

constructor TOPC_ContentTypes.Create;
begin
  inherited Create;
end;

function TOPC_ContentTypes.FindDefault(AExtension: AxUCString): integer;
begin
  for Result := 0 to Count - 1  do begin
    if (Items[Result] is TOPC_ContentTypeDefault) and (TOPC_ContentTypeDefault(Items[Result]).Extension = AExtension) then
      Exit;
  end;
  Result := -1;
end;

function TOPC_ContentTypes.FindOverride(APartName: AxUCString): integer;
begin
  for Result := 0 to Count - 1  do begin
    if (Items[Result] is TOPC_ContentTypeOverride) and (TOPC_ContentTypeOverride(Items[Result]).PartName = APartName) then
      Exit;
  end;
  Result := -1;
end;

function TOPC_ContentTypes.GetItems(Index: integer): TOPC_ContentType;
begin
  Result := TOPC_ContentType(inherited Items[Index]);
end;

procedure TOPC_ContentTypes.SetDefault;
begin
  inherited Clear;

  if FXLSXMode then begin
    AddDefault('bin','application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings');
    AddDefault('rels','application/vnd.openxmlformats-package.relationships+xml');
    AddDefault('xml','application/xml');
    AddDefault('vml','application/vnd.openxmlformats-officedocument.vmlDrawing');
    AddDefault('jpeg','image/jpeg');
    AddDefault('png','image/png');
    AddDefault('gif','image/gif');
    AddDefault('wmf','image/x-wmf');
    AddDefault('emf','image/x-emf');
  end
  else begin
    AddDefault('rels','application/vnd.openxmlformats-package.relationships+xml');
    AddDefault('xml','application/xml');
    AddDefault('jpeg','image/jpeg');
    AddDefault('png','image/png');
    AddDefault('gif','image/gif');
    AddDefault('bmp','image/bmp');
    AddDefault('wmf','image/x-wmf');
    AddDefault('emf','image/x-emf');
  end;
end;

{ TOPC_DOCX }

function TOPC_DOCX.AddDocPropsApp(AStream: TStream): TOPCItem;
begin
  Result := FRoot.AddNew(OPC_DOCX_DOCPROPS_APP,OPC_DOCX_DOCPROPS_APP_TARGET);
  FZIP.Write(Result.AbsolutePath,Result.Name,AStream);
  FContentTypes.AddOverride(Result.AbsoluteTarget,'application/vnd.openxmlformats-officedocument.extended-properties+xml');
end;

function TOPC_DOCX.AddDocPropsCore(AStream: TStream): TOPCItem;
begin
  Result := FRoot.AddNew(OPC_DOCX_DOCPROPS_CORE,OPC_DOCX_DOCPROPS_CORE_TARGET);
  FZIP.Write(Result.AbsolutePath,Result.Name,AStream);
  FContentTypes.AddOverride(Result.AbsoluteTarget,'application/vnd.openxmlformats-package.core-properties+xml');
end;

function TOPC_DOCX.AddImage(AIndex: integer; AExt: AxUCString; AStream: TStream): TOPCItem;
var
  S: string;
begin
  S := Format(OPC_DOCX_IMAGE_TARGET,[AIndex,AExt]);
  Result := ItemCreate(FDocument,OPC_DOCX_IMAGE,S,'');
  FZIP.Write(Result.AbsolutePath,Result.Name,AStream);
end;

function TOPC_DOCX.AddSettings(AStream: TStream): TOPCItem;
begin
  Result := ItemCreate(FDocument,OPC_DOCX_SETTINGS,OPC_DOCX_SETTINGS_TARGET,'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml');
  FZIP.Write(Result.AbsolutePath,Result.Name,AStream);
end;

function TOPC_DOCX.AddTheme(AStream: TStream): TOPCItem;
begin
  Result := ItemCreate(FDocument,OPC_DOCX_THEME,Format(OPC_DOCX_THEME_TARGET,[1]),'application/vnd.openxmlformats-officedocument.theme+xml');
  FZIP.Write(Result.AbsolutePath,Result.Name,AStream);
end;

procedure TOPC_DOCX.Close;
begin
  inherited Close;

  case FState of
    opcsClosed: ;
    opcsRead: ;
    opcsWrite: begin
      ItemClose(FDocument);
    end;
  end;

  FDocument := Nil;
  FNumbering := Nil;
end;

constructor TOPC_DOCX.Create;
begin
  inherited Create;

  FContentTypes.SetDefault;

  FAutoCheckout := False;
  FXmlns := OOXML_URI_PACKAGE_RELATIONSHIPS;
  FResultList := TOPCItemList.Create;
end;

function TOPC_DOCX.CreateDocument: TOPCItem;
begin
  if FDocument <> Nil then
    raise XLSRWException.Create('OPC workbook is open');
  Result := ItemCreate(FRoot,OPC_DOCX_DOCUMENT,OPC_DOCX_DOCUMENT_TARGET,'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml');
  FDocument := Result;
end;

function TOPC_DOCX.CreateEndnotes: TOPCItem;
begin
  Result := ItemCreate(FDocument,OPC_DOCX_ENDNOTES,OPC_DOCX_ENDNOTES_TARGET,'application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml');
end;

function TOPC_DOCX.CreateFooter(AIndex: integer): TOPCItem;
begin
  Result := ItemCreate(FDocument,OPC_DOCX_FOOTER,Format(OPC_DOCX_FOOTER_TARGET,[AIndex]),'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml');
end;

function TOPC_DOCX.CreateFootnotes: TOPCItem;
begin
  Result := ItemCreate(FDocument,OPC_DOCX_FOOTNOTES,OPC_DOCX_FOOTNOTES_TARGET,'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml');
end;

function TOPC_DOCX.CreateHeader(AIndex: integer): TOPCItem;
begin
  Result := ItemCreate(FDocument,OPC_DOCX_HEADER,Format(OPC_DOCX_HEADER_TARGET,[AIndex]),'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml');
end;

function TOPC_DOCX.CreateNumbering: TOPCItem;
begin
  Result := ItemCreate(FDocument,OPC_DOCX_NUMBERING,OPC_DOCX_NUMBERING_TARGET,'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml');
end;

function TOPC_DOCX.CreateStyles: TOPCItem;
begin
  Result := ItemCreate(FDocument,OPC_DOCX_STYLES,OPC_DOCX_STYLES_TARGET,'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml');
end;

destructor TOPC_DOCX.Destroy;
begin
  FResultList.Free;

  inherited;
end;

function TOPC_DOCX.FindFooters: boolean;
begin
  Result := FDocument.FindByType(OPC_DOCX_FOOTER,FResultList);
end;

function TOPC_DOCX.FindHeaders: boolean;
begin
  Result := FDocument.FindByType(OPC_DOCX_HEADER,FResultList);
end;

function TOPC_DOCX.ReadHeader(const ArId: AxUCString): TStream;
var
  Item: TOPCItem;
begin
  Item := ItemOpenRead(FDocument,OPC_DOCX_HEADER,ArId);
  if Item <> Nil then begin
    Result := ItemOpenStream(Item);
    ItemClose(Item);
  end
  else
    Result := Nil;
end;

function TOPC_DOCX.ReadHeaderFooterImages: boolean;
var
  i: integer;
  HdrFtr: TOPCItemList;
begin
  HdrFtr := TOPCItemList.Create;
  try
    FResultList.Clear;

    FDocument.FindByType(OPC_DOCX_HEADER,HdrFtr);
    for i := 0 to HdrFtr.Count - 1 do
      HdrFtr[i].FindByType(OPC_XLSX_IMAGE,FResultList,False);

    FDocument.FindByType(OPC_DOCX_FOOTER,HdrFtr);
    for i := 0 to HdrFtr.Count - 1 do
      HdrFtr[i].FindByType(OPC_XLSX_IMAGE,FResultList,False);
  finally
    HdrFtr.Free;
  end;

  Result := FResultList.Count > 0;
end;

function TOPC_DOCX.ReadHyperlinks(const AList: TStrings): boolean;
var
  i: integer;
begin
  for i := 0 to FDocument.Count - 1 do begin
    if FDocument[i].Type_ = OPC_XLSX_HYPERLINK then
      AList.Add(FDocument[i].Id + ';' + FDocument[i].Target);
  end;
  Result := AList.Count > 0;
end;

function TOPC_DOCX.ReadImage(const ArId: AxUCString; var AName: AxUCString): TStream;
var
  i: integer;
begin
  for i := 0 to FDocument.Count - 1 do begin
    if (FDocument[i].Type_ = OPC_DOCX_IMAGE) and (FDocument[i].Id = ArId) then begin

    end;
  end;
  Result := Nil;
end;

procedure TOPC_DOCX.OpenRead(AStream: TStream; const AUseAlternateZip: boolean);
begin
  inherited OpenRead(AStream,AUseAlternateZip);

  FResultList.Clear;
  if not FRoot.FindByType(OPC_DOCX_DOCUMENT,FResultList) then
    raise XLSRWException.Create('Can not find document in OPC');
  if FResultList.Count > 1 then
    raise XLSRWException.Create('More than one document in OPC');
  FDocument := FResultList[0];
  FResultList.Clear;
end;

procedure TOPC_DOCX.OpenWrite(AStream: TStream; const AUseAlternateZip: boolean);
begin
  inherited OpenWrite(AStream,AUseAlternateZip);
end;

function TOPC_DOCX.ReadDocImages: boolean;
begin
  FResultList.Clear;
  Result := FDocument.FindByType(OPC_DOCX_IMAGE,FResultList);
end;

function TOPC_DOCX.ReadDocument: TStream;
begin
  FDocument := ItemOpenRead(FRoot,OPC_DOCX_DOCUMENT);
  if FDocument <> Nil then begin
    Result := ItemOpenStream(FDocument);
    ItemClose(FDocument);
  end
  else
    Result := Nil;
end;

function TOPC_DOCX.ReadEndnotes: TStream;
var
  Item: TOPCItem;
begin
  Item := ItemOpenRead(FDocument,OPC_DOCX_ENDNOTES);
  if Item <> Nil then begin
    Result := ItemOpenStream(Item);
    ItemClose(Item);
  end
  else
    Result := Nil;
end;

function TOPC_DOCX.ReadFooter(const ArId: AxUCString): TStream;
var
  Item: TOPCItem;
begin
  Item := ItemOpenRead(FDocument,OPC_DOCX_FOOTER,ArId);
  if Item <> Nil then begin
    Result := ItemOpenStream(Item);
    ItemClose(Item);
  end
  else
    Result := Nil;
end;

function TOPC_DOCX.ReadFootnotes: TStream;
var
  Item: TOPCItem;
begin
  Item := ItemOpenRead(FDocument,OPC_DOCX_FOOTNOTES);
  if Item <> Nil then begin
    Result := ItemOpenStream(Item);
    ItemClose(Item);
  end
  else
    Result := Nil;
end;

function TOPC_DOCX.ReadNumbering: TStream;
begin
  FNumbering := ItemOpenRead(FDocument,OPC_DOCX_NUMBERING);
  if FNumbering <> Nil then begin
    Result := ItemOpenStream(FNumbering);
    ItemClose(FNumbering);
  end
  else
    Result := Nil;
end;

function TOPC_DOCX.ReadNumberingImages: boolean;
begin
  FResultList.Clear;
  Result := FNumbering.FindByType(OPC_DOCX_IMAGE,FResultList);
end;

function TOPC_DOCX.ReadSettings: TStream;
var
  Item: TOPCItem;
begin
  Item := ItemOpenRead(FDocument,OPC_DOCX_SETTINGS);
  if Item <> Nil then begin
    Result := ItemOpenStream(Item);
    ItemClose(Item);
  end
  else
    Result := Nil;
end;

function TOPC_DOCX.ReadStyles: TStream;
begin
  FStyles := ItemOpenRead(FDocument,OPC_DOCX_STYLES);
  if FStyles <> Nil then begin
    Result := ItemOpenStream(FStyles);
    ItemClose(FStyles);
  end
  else
    Result := Nil;
end;

function TOPC_DOCX.ReadTheme: TStream;
var
  Item: TOPCItem;
begin
  Item := ItemOpenRead(FDocument,OPC_XLSX_THEME);
  if Item <> Nil then begin
    Result := ItemOpenStream(Item);
    ItemClose(Item);
  end
  else
    Result := Nil;
end;

end.
