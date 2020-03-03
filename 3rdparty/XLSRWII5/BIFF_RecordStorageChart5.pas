unit BIFF_RecordStorageChart5;

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

uses Classes, SysUtils, Contnrs,
     BIFF_RecordStorage5, BIFF_RecsII5, BIFF_Stream5, BIFF_EscherTypes5,
     Xc12DataStylesheet5,
     XLSUtils5;

const ID_CHARTRECORDROOT = $FFFA;
const ID_CHARTRECORD_DEFAULTLEGEND = $F001;

type TDefaultRecordData = (drdAll,drdLegend,drdSerie,drdDataformat,drdStyleArea,
                           drdStyleBarColumn,drdStyleLine,drdStylePie,
                           drdStyleRadar,drdStyleScatter,drdStyleSurface);

type TFontIdListData = class(TObject)
private
     FFontIndex,FFBIIndex: integer;
public
     property FontIndex: integer read FFontIndex write FFontIndex;
     property FBIIndex: integer read FFBIIndex write FFBIIndex;
     end;

type TFontIdList = class(TObjectList)
private
     function GetItems(Index: integer): TFontIdListData;
protected
     FCurrent: integer;
public
     procedure AddFont(Index: integer);
     procedure Reset;
     function  GetNext(FBI: integer): integer;
     function  FindFBI(Id: integer): integer;

     property Items[Index: integer]: TFontIdListData read GetItems;
     end;

type TDefaultRecord = record
     Id: word;
     Data: AnsiString;
     end;

type TChartRecord = class(TObjectList)
private
     FRecId: word;
     FLength: word;
     FData: PByteArray;
     FParent: TChartRecord;

     function  GetItems(Index: integer): TChartRecord;
     procedure ReadDefault(DefRecs: array of TDefaultRecord; var Index: integer; Fonts: TXc12Fonts);
     function  FindFBIFont(Id: integer): TXc12Font;
     procedure FixupFBI;
     procedure ReadFromStream(Stream: TXLSStream; Fonts: TXc12Fonts);
public
     constructor Create(Parent: TChartRecord; Header: TBIFFHeader; D: PByteArray); overload;
     constructor Create(Parent: TChartRecord; RecId,Length: word); overload;
     constructor Create(Parent: TChartRecord; RecId: word; Obj: TObject); overload;
     destructor Destroy; override;
     procedure Write(Stream: TXLSStream);
     procedure WriteToFile(Filename: AxUCString; Fonts: TXc12Fonts; AsText: boolean);
     procedure ReadFromFile(Filename: AxUCString; Fonts: TXc12Fonts);
     procedure ReadFromBuffer(Buf: PByteArray; BufSize: integer; Fonts: TXc12Fonts);
     procedure Resize(Delta: integer);
     procedure Read(Stream: TXLSStream; PBuf: PByteArray; Fonts: TXc12Fonts);
     procedure ReadDefaultRecords(RecData: TDefaultRecordData);
     function  LastRec: TChartRecord;
     function  InsertRecord(Index: integer; RecId,Length: word; IsUpdate: boolean = False): TChartRecord;
     function  FindRecord(Id: integer): integer;
     function  FindRecordChilds(Id: integer): boolean;
     function  RemoveRecord(Id: integer): boolean;
     function  Root: TChartRecord;

     property Items[Index: integer]: TChartRecord read GetItems; default;
     property Parent: TChartRecord read FParent;
     property RecId: word read FRecId;
     property Length: word read FLength;
     property Data: PByteArray read FData;
     end;

type TChartRecordUpdate = class(TChartRecord)
private
     FUpdateEvent: TNotifyEvent;
public
     property OnUpdate: TNotifyEvent read FUpdateEvent write FUpdateEvent;
     end;

implementation

uses BIFF_Escher5, BIFF_DrawingObjChart5;

var
  FFontIdList: TFontIdList;

const DefaultRecordsSerieIndex = 22;

const DefaultRecordsLegend: array[0..7] of TDefaultRecord = (
(Id: $1033; Data: ''), //BEGIN:
(Id: $104F; Data: #$05#$00#$02#$00#$DB#$0D#$00#$00#$AB#$06#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00), // POS
(Id: $1025; Data: #$02#$02#$01#$00#$00#$00#$00#$00#$AF#$FF#$FF#$FF#$86#$FF#$FF#$FF#$00#$00#$00#$00#$00#$00#$00#$00#$B1#$00#$4D#$00#$A0#$15#$00#$00), // TEXT
(Id: $1033; Data: ''), //BEGIN:
(Id: $104F; Data: #$02#$00#$02#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00), // POS
(Id: $1051; Data: #$00#$01#$00#$00#$00#$00#$00#$00), // AI
(Id: $1034; Data: ''), //END:
(Id: $1034; Data: '')); //END:

const DefaultRecordsAll: array[0..82] of TDefaultRecord = (
(Id: $0809; Data: #$00#$06#$20#$00#$46#$18#$CD#$07#$C1#$80#$00#$00#$06#$02#$00#$00), // BOF
(Id: $0014; Data: ''), // HEADER
(Id: $0015; Data: ''), // FOOTER
(Id: $0083; Data: #$00#$00), // HCENTER
(Id: $0084; Data: #$00#$00), //VCENTER
(Id: $00A1; Data: #$09#$00#$64#$00#$01#$00#$01#$00#$01#$00#$00#$00#$58#$02#$58#$02#$00#$00#$00#$00#$00#$00#$E0#$3F#$00#$00#$00#$00#$00#$00#$E0#$3F#$01#$00), // SETUP
(Id: $0033; Data: #$03#$00),// ???
                                                           // Four extra bytes for pointer to the font.
(Id: $1060; Data: #$08#$16#$57#$12#$A0#$00#$00#$00#$F0#$00#$00#$00#$00#$00), // FBI:
(Id: $1060; Data: #$08#$16#$57#$12#$A0#$00#$01#$00#$F1#$00#$00#$00#$00#$00), // FBI:

(Id: $0012; Data: #$00#$00), // PROTECT
(Id: $1001; Data: #$00#$00), // UNITS:
(Id: $1002; Data: #$00#$00#$00#$00#$00#$00#$00#$00#$E8#$7F#$21#$01#$00#$40#$BF#$00), // CHART:
(Id: $1033; Data: ''), // BEGIN:
(Id: $00A0; Data: #$01#$00#$01#$00), // SCL
(Id: $1064; Data: #$00#$00#$01#$00#$00#$00#$01#$00), // PLOTGROWTH:
(Id: $1032; Data: #$00#$00#$02#$00), // FRAME:
(Id: $1033; Data: ''), // BEGIN:
(Id: $1007; Data: #$00#$00#$00#$00#$00#$00#$00#$00#$09#$00#$4D#$00), // LINEFORMAT:
(Id: $100A; Data: #$FF#$FF#$FF#$00#$00#$00#$00#$00#$01#$00#$01#$00#$4E#$00#$4D#$00), // AREAFORMAT:
(Id: $1034; Data: ''), // END:

(Id: $1003; Data: #$01#$00#$01#$00#$0A#$00#$0A#$00#$01#$00#$00#$00), // SERIES:
(Id: $1033; Data: ''), // BEGIN:
(Id: $1051; Data: #$00#$01#$00#$00#$00#$00#$00#$00), // AI:
(Id: $1051; Data: #$01#$02#$00#$00#$00#$00#$0B#$00#$3B#$00#$00#$00#$00#$09#$00#$00#$00#$00#$00), // AI:
(Id: $1051; Data: #$02#$00#$00#$00#$00#$00#$00#$00), // AI:
(Id: $1051; Data: #$03#$01#$00#$00#$00#$00#$00#$00), // AI:
(Id: $1006; Data: #$FF#$FF#$00#$00#$00#$00#$00#$00), // DATAFORMAT:
(Id: $1033; Data: ''), // BEGIN:
(Id: $105F; Data: #$00#$00),// ???
(Id: $1034; Data: ''), // END:
(Id: $1045; Data: #$00#$00), // SERTOCRT:
(Id: $1034; Data: ''), // END:
(Id: $1044; Data: #$0A#$00#$00#$00), // SHTPROPS:
(Id: $1024; Data: #$02#$00), // DEFAULTTEXT:
(Id: $1025; Data: #$02#$02#$01#$00#$00#$00#$00#$00#$CB#$FF#$FF#$FF#$AE#$FF#$FF#$FF#$00#$00#$00#$00#$00#$00#$00#$00#$B1#$00#$4D#$00#$A0#$09#$00#$00), // TEXT:
(Id: $1033; Data: ''), // BEGIN:
(Id: $104F; Data: #$02#$00#$02#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00), // POS:
(Id: $1026; Data: #$F0#$00), // FONTX:
(Id: $1051; Data: #$00#$01#$00#$00#$00#$00#$00#$00), // AI:
(Id: $1034; Data: ''), // END:
(Id: $1024; Data: #$03#$00), // DEFAULTTEXT:
(Id: $1025; Data: #$02#$02#$01#$00#$00#$00#$00#$00#$CB#$FF#$FF#$FF#$AE#$FF#$FF#$FF#$00#$00#$00#$00#$00#$00#$00#$00#$B1#$00#$4D#$00#$A0#$09#$00#$00), // TEXT:
(Id: $1033; Data: ''), // BEGIN:
(Id: $104F; Data: #$02#$00#$02#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00), // POS:
(Id: $1026; Data: #$F1#$00), // FONTX:
(Id: $1051; Data: #$00#$01#$00#$00#$00#$00#$00#$00), // AI:
(Id: $1034; Data: ''), // END:
(Id: $1046; Data: #$01#$00), // AXESUSED:
(Id: $1041; Data: #$00#$00#$94#$01#$00#$00#$47#$01#$00#$00#$23#$0B#$00#$00#$FD#$0B#$00#$00), // AXISPARENT:
(Id: $1033; Data: ''), // BEGIN:
(Id: $104F; Data: #$02#$00#$02#$00#$6A#$00#$00#$00#$A3#$00#$00#$00#$4D#$0C#$00#$00#$59#$0E#$00#$00), // POS:

(Id: $101D; Data: #$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00), // AXIS:
(Id: $1033; Data: ''), // BEGIN:
(Id: $1020; Data: #$01#$00#$01#$00#$01#$00#$01#$00), // CATSERRANGE:
(Id: $1062; Data: #$00#$00#$00#$00#$01#$00#$00#$00#$01#$00#$00#$00#$00#$00#$00#$00#$EF#$00), // AXCEXT:
(Id: $101E; Data: #$02#$00#$03#$01#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$23#$00#$4D#$00#$00#$00), // TICK:
(Id: $1034; Data: ''), // END:

(Id: $101D; Data: #$01#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00), // AXIS:
(Id: $1033; Data: ''), // BEGIN:
(Id: $101F; Data: #$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$1F#$01), // VALUERANGE:
(Id: $101E; Data: #$02#$00#$03#$01#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$23#$00#$4D#$00#$00#$00), // TICK:
(Id: $1021; Data: #$01#$00), // AXISLINEFORMAT:
(Id: $1007; Data: #$00#$00#$00#$00#$00#$00#$FF#$FF#$09#$00#$4D#$00), // LINEFORMAT:
(Id: $1034; Data: ''), // END:

(Id: $1035; Data: ''), // PLOTAREA:
(Id: $1032; Data: #$00#$00#$03#$00), // FRAME:
(Id: $1033; Data: ''), // BEGIN:
(Id: $1007; Data: #$80#$80#$80#$00#$00#$00#$00#$00#$00#$00#$17#$00), // LINEFORMAT:
(Id: $100A; Data: #$C0#$C0#$C0#$00#$00#$00#$00#$00#$01#$00#$00#$00#$16#$00#$4F#$00), // AREAFORMAT:
(Id: $1034; Data: ''), // END:
(Id: $1014; Data: #$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00), // CHARTFORMAT:
(Id: $1033; Data: ''), // BEGIN:
(Id: $1017; Data: #$00#$00#$96#$00#$00#$00), // BAR:
(Id: $1022; Data: #$00#$00#$00#$00#$00#$00#$00#$00#$0F#$00), // CHARTFORMATLINK:

(Id: $1015; Data: #$DB#$0D#$00#$00#$AB#$06#$00#$00#$85#$01#$00#$00#$F4#$00#$00#$00#$03#$01#$1F#$00), // LEGEND
(Id: ID_CHARTRECORD_DEFAULTLEGEND; Data: ''), // DefaultRecordsLegend

(Id: $1034; Data: ''), // END:
(Id: $1034; Data: ''), // END:
(Id: $1034; Data: ''), // END:
//(Id: $0200; Data: #$00#$00#$00#$00#$0A#$00#$00#$00#$00#$00#$01#$00#$00#$00), // DIMENSIONS
(Id: $1065; Data: #$02#$00), // SIINDEX:
(Id: $1065; Data: #$01#$00), // SIINDEX:
(Id: $1065; Data: #$03#$00), // SIINDEX:
(Id: $000A; Data: '')); // EOF

const DefaultRecordsStyleArea: array[0..11] of TDefaultRecord = (
(Id: $101A; Data: #$00#$00), //AREA
(Id: $1022; Data: #$00#$00#$00#$00#$00#$00#$00#$00#$0F#$00), //CHARTFORMATLINK

(Id: $1015; Data: #$DB#$0D#$00#$00#$AB#$06#$00#$00#$85#$01#$00#$00#$F4#$00#$00#$00#$03#$01#$1F#$00), // LEGEND
(Id: ID_CHARTRECORD_DEFAULTLEGEND; Data: ''), // DefaultRecordsLegend

(Id: $1006; Data: #$00#$00#$00#$00#$FD#$FF#$00#$00), //DATAFORMAT
(Id: $1033; Data: ''), //BEGIN
(Id: $105F; Data: #$00#$00), // GEOMETRY
(Id: $1007; Data: #$00#$00#$00#$00#$00#$00#$FF#$FF#$09#$00#$4D#$00), //LINEFORMAT
(Id: $100A; Data: #$FF#$FF#$FF#$00#$00#$00#$00#$00#$01#$00#$01#$00#$4E#$00#$4D#$00), //AREAFORMAT
(Id: $100B; Data: #$00#$00), //PIEFORMAT
(Id: $1009; Data: #$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$4D#$00#$4D#$00#$3C#$00#$00#$00), //MARKERFORMAT
(Id: $1034; Data: '')); //END

const DefaultRecordsStyleBarColumn: array[0..3] of TDefaultRecord = (
(Id: $1017; Data: #$00#$00#$96#$00#$01#$00), // BAR
(Id: $1022; Data: #$00#$00#$00#$00#$00#$00#$00#$00#$0F#$00), // CHARTFORMATLINK

(Id: $1015; Data: #$DB#$0D#$00#$00#$AB#$06#$00#$00#$85#$01#$00#$00#$F4#$00#$00#$00#$03#$01#$1F#$00), // LEGEND
(Id: ID_CHARTRECORD_DEFAULTLEGEND; Data: '')); // DefaultRecordsLegend

const DefaultRecordsStyleScatter: array[0..11] of TDefaultRecord = (
(Id: $101B; Data: #$64#$00#$01#$00#$00#$00), // SCATTER
(Id: $1022; Data: #$00#$00#$00#$00#$00#$00#$00#$00#$0F#$00), // CHARTFORMATLINK

(Id: $1015; Data: #$DB#$0D#$00#$00#$AB#$06#$00#$00#$85#$01#$00#$00#$F4#$00#$00#$00#$03#$01#$1F#$00), // LEGEND
(Id: ID_CHARTRECORD_DEFAULTLEGEND; Data: ''), // DefaultRecordsLegend

(Id: $1006; Data: #$00#$00#$00#$00#$FD#$FF#$00#$00), // DATAFORMAT
(Id: $1033; Data: ''), // BEGIN
(Id: $105F; Data: #$00#$00), // ???
(Id: $1007; Data: #$00#$00#$00#$00#$00#$00#$FF#$FF#$09#$00#$4D#$00), // LINEFORMAT
(Id: $100A; Data: #$FF#$FF#$FF#$00#$00#$00#$00#$00#$01#$00#$01#$00#$4E#$00#$4D#$00), // AREAFORMAT
(Id: $100B; Data: #$00#$00), // PIEFORMAT
(Id: $1009; Data: #$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$4D#$00#$4D#$00#$3C#$00#$00#$00), // MARKERFORMAT
(Id: $1034; Data: ''));// END

const DefaultRecordsStylePie: array[0..3] of TDefaultRecord = (
(Id: $1019; Data: #$00#$00#$00#$00#$02#$00), //PIE
(Id: $1022; Data: #$00#$00#$00#$00#$00#$00#$00#$00#$0F#$00), //CHARTFORMATLINK
(Id: $1015; Data: #$DB#$0D#$00#$00#$AB#$06#$00#$00#$85#$01#$00#$00#$F4#$00#$00#$00#$03#$01#$1F#$00), // LEGEND
(Id: ID_CHARTRECORD_DEFAULTLEGEND; Data: '')); // DefaultRecordsLegend

const DefaultRecordsStyleLine: array[0..3] of TDefaultRecord = (
(Id: $1018; Data: #$00#$00), // Line
(Id: $1022; Data: #$00#$00#$00#$00#$00#$00#$00#$00#$0F#$00), //CHARTFORMATLINK
(Id: $1015; Data: #$DB#$0D#$00#$00#$AB#$06#$00#$00#$85#$01#$00#$00#$F4#$00#$00#$00#$03#$01#$1F#$00), // LEGEND
(Id: ID_CHARTRECORD_DEFAULTLEGEND; Data: '')); // DefaultRecordsLegend

const DefaultRecordsStyleRadar: array[0..11] of TDefaultRecord = (
(Id: $103E; Data: #$01#$00#$12#$00), // RADAR
(Id: $1022; Data: #$00#$00#$00#$00#$00#$00#$00#$00#$0F#$00), // CHARTFORMATLINK

(Id: $1015; Data: #$DB#$0D#$00#$00#$AB#$06#$00#$00#$85#$01#$00#$00#$F4#$00#$00#$00#$03#$01#$1F#$00), // LEGEND
(Id: ID_CHARTRECORD_DEFAULTLEGEND; Data: ''), // DefaultRecordsLegend

(Id: $1006; Data: #$00#$00#$00#$00#$FD#$FF#$00#$00), // DATAFORMAT
(Id: $1033; Data: ''), // BEGIN
(Id: $105F; Data: #$00#$00), // _GEOMETRY_
(Id: $1007; Data: #$00#$00#$00#$00#$00#$00#$FF#$FF#$09#$00#$4D#$00), // LINEFORMAT
(Id: $100A; Data: #$FF#$FF#$FF#$00#$00#$00#$00#$00#$01#$00#$01#$00#$4E#$00#$4D#$00), // AREAFORMAT
(Id: $100B; Data: #$00#$00), // PIEFORMAT
(Id: $1009; Data: #$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$4D#$00#$4D#$00#$3C#$00#$00#$00), // MARKERFORMAT
(Id: $1034; Data: '')); // END

const DefaultRecordsStyleSurface: array[0..11] of TDefaultRecord = (
(Id: $103F; Data: #$01#$00), // SURFACE
(Id: $1022; Data: #$00#$00#$00#$00#$00#$00#$00#$00#$0F#$00), // CHARTFORMATLINK

(Id: $1015; Data: #$DB#$0D#$00#$00#$AB#$06#$00#$00#$85#$01#$00#$00#$F4#$00#$00#$00#$03#$01#$1F#$00), // LEGEND
(Id: ID_CHARTRECORD_DEFAULTLEGEND; Data: ''), // DefaultRecordsLegend

(Id: $1006; Data: #$00#$00#$00#$00#$FD#$FF#$00#$00), // DATAFORMAT
(Id: $1033; Data: ''), // BEGIN
(Id: $105F; Data: #$00#$00), // _GEOMETRY_
(Id: $1007; Data: #$00#$00#$00#$00#$00#$00#$FF#$FF#$09#$00#$4D#$00), // LINEFORMAT
(Id: $100A; Data: #$FF#$FF#$FF#$00#$00#$00#$00#$00#$01#$00#$01#$00#$4E#$00#$4D#$00), // AREAFORMAT
(Id: $100B; Data: #$00#$00), // PIEFORMAT
(Id: $1009; Data: #$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$00#$4D#$00#$4D#$00#$3C#$00#$00#$00), // MARKERFORMAT
(Id: $1034; Data: '')); // END

const DefaultRecordsDataformat: array[0..4] of TDefaultRecord = (
// (Id: $1033; Data: ''), // BEGIN
(Id: $105F; Data:  #$00#$00), // _GEOMETRY_
(Id: $1007; Data:  #$00#$00#$00#$00#$00#$00#$00#$00#$01#$00#$4F#$00), // LINEFORMAT
(Id: $100A; Data:  #$CC#$FF#$CC#$00#$FF#$FF#$FF#$00#$01#$00#$00#$00#$2A#$00#$09#$00), // AREAFORMAT
(Id: $100B; Data:  #$00#$00), // PIEFORMAT
(Id: $1009; Data:  #$00#$00#$00#$00#$00#$00#$00#$00#$02#$00#$01#$00#$4D#$00#$4D#$00#$3C#$00#$00#$00)); // MARKERFORMAT
// (Id: $1034; Data: '')); // END


{ TChartRecord }

function TChartRecord.FindFBIFont(Id: integer): TXc12Font;
var
  i: integer;
  Rec: TChartRecord;
  HasFBI: boolean;
begin
  HasFBI := False;
  Rec := Root;
  for i := 0 to Rec.Count - 1 do begin
    HasFBI := Rec[i].FRecId = CHARTRECID_FBI;
    if HasFBI and (PCRecFBI_Font(Rec[i].FData).Index = Id) then begin
      Result := TXc12Font(PCRecFBI_Font(Rec[i].FData).Xc12Font);
      Exit;
    end;
  end;
  Result := Nil;
  if HasFBI then
    raise XLSRWException.Create('Can not find FBI');
end;

function TChartRecord.FindRecord(Id: integer): integer;
begin
  for Result := 0 to Count - 1 do begin
    if Items[Result].RecId = Id then
      Exit;
  end;
  Result := -1;
end;

function TChartRecord.FindRecordChilds(Id: integer): boolean;
var
  i: integer;
begin
  for i := 0 to Count - 1 do begin
    if Items[i].Count > 0 then begin
      Result := Items[i].FindRecordChilds(Id);
      if Result then
        Exit;
    end;
    if Items[i].RecId = Id then begin
      Result := True;
      Exit;
    end;
  end;
  Result := False;
end;

function TChartRecord.GetItems(Index: integer): TChartRecord;
begin
  Result := TChartRecord(Inherited Items[Index]);
end;

function TChartRecord.InsertRecord(Index: integer; RecId, Length: word; IsUpdate: boolean = False): TChartRecord;
begin
  if Index < 0 then
    Index := Count;
  if IsUpdate then
    Insert(Index,TChartRecordUpdate.Create(Self,RecId,Length))
  else
    Insert(Index,TChartRecord.Create(Self,RecId,Length));
  Result := Items[Index];
end;

function TChartRecord.LastRec: TChartRecord;
begin
  if Count <= 0 then
    raise XLSRWException.Create('Parent record is missing in chart');
  Result := Items[Count - 1];
end;

procedure TChartRecord.Read(Stream: TXLSStream; PBuf: PByteArray; Fonts: TXc12Fonts);
var
  i,Inx: integer;
  H,Header: TBIFFHeader;
  RecId: word;
  P,Buf: PByteArray;
  Shape: TShapeOutsideMsoChart;
  SheetCharts: TSheetCharts;
  XFont: TXc12Font;
  O: TObject;
begin
  while Stream.ReadHeader(Header) = SizeOf(TBIFFHeader) do begin
    case Header.RecId of
      BIFFRECID_EOF: begin
        Stream.Read(PBuf^,Header.Length);
        Add(TChartRecord.Create(Self,Header,PBuf));
        Exit;
      end;
      BIFFRECID_MSODRAWING: begin
        // Empty drawing. Reading empty drawing causes problems as these
        // don't have an entry in the MSODRAWINGGROUP.
        // A better detection ought to be used, but this is
        // complicated and empty drawings are rare. The correct method is
        // to check the number of shapes. Is more than one for a non-empty
        // drawing.
        if Header.Length <= 80 then
          Stream.Read(PBuf^,Header.Length)
        else begin
          i := Add(TChartRecord.Create(Self,Header,Nil));
          O := Items[i].Root.Parent;
          if O.ClassName = 'TShapeOutsideMsoChart' then begin
            Shape := TShapeOutsideMsoChart(Items[i].Root.Parent);
            Shape.Drawing := TEscherDrawing.Create(Shape.DrawingGroup,Nil,Nil);
            Shape.Drawing.LoadFromStream(Stream,PBuf);
          end
          else if O.ClassName = 'TSheetCharts' then begin
            SheetCharts := TSheetCharts(Items[i].Root.Parent);
            SheetCharts[SheetCharts.Count - 1].AddDrawing(TEscherDrawing.Create(SheetCharts.DrawingGroup,Fonts,Nil));
            SheetCharts[SheetCharts.Count - 1].Drawing.LoadFromStream(Stream,PBuf);
          end
          else
            raise XLSRWException.Create('Chart record parent is not of excpected type');
        end;
      end;
      CHARTRECID_BEGIN: begin
        Items[Count - 1].Read(Stream,PBuf,Fonts);
      end;
      CHARTRECID_GELFRAME: begin
        H.Length := Header.Length;
        H.RecID := Header.RecID;
        Buf := AllocMem(H.Length);
        try
          Stream.Read(Buf^,Header.Length);
          repeat
            RecId := Stream.PeekHeader;
            if ((RecId = CHARTRECID_GELFRAME) or (RecId = BIFFRECID_CONTINUE)) then begin
              if Stream.ReadHeader(Header) <> SizeOf(TBIFFHeader) then
                Break;
              Stream.Read(PBuf^,Header.Length);

              ReAllocMem(Buf,H.Length + Header.Length);
              P := PByteArray(NativeInt(Buf) + H.Length);
              System.Move(PBuf^,P^,Header.Length);
              Inc(H.Length,Header.Length);
            end;
          until ((RecId <> CHARTRECID_GELFRAME) and (RecId <> BIFFRECID_CONTINUE));
          Add(TChartRecordUpdate.Create(Self,H,Buf));
        finally
          FreeMem(Buf);
        end;
      end;
      CHARTRECID_FBI: begin
        Stream.Read(PBuf^,Header.Length);
        Header.Length := SizeOf(TCRecFBI_Font);
        if FFontIdList <> Nil then begin
          Inx := FFontIdList.GetNext(PCRecFBI(PBuf).Index);
          PCRecFBI_Font(PBuf).Xc12Font := Fonts[Inx];
        end
        else
          PCRecFBI_Font(PBuf).Xc12Font := Fonts[PCRecFBI(PBuf).Index];
        TXc12Font(PCRecFBI_Font(PBuf).Xc12Font).Locked := True;
        Add(TChartRecord.Create(Self,Header,PBuf));
      end;
      CHARTRECID_FONTX: begin
        Stream.Read(PBuf^,Header.Length);
        XFont := FindFBIFont(PCRecFONTX(PBuf).FontIndex);
        if XFont <> Nil then
          Add(TChartRecord.Create(Self,Header.RecId,XFont))
        else
          Add(TChartRecord.Create(Self,Header.RecId,Fonts[PCRecFONTX(PBuf).FontIndex]));
      end;
      CHARTRECID_END: begin
        Stream.Read(PBuf^,Header.Length);
        Exit;
      end;
      else begin
        if (Header.RecID and $00FF) = BIFFRECID_BOF then
          Stream.ReadUnencryptedSync(PBuf^,Header.Length)
        else
          Stream.Read(PBuf^,Header.Length);
        Add(TChartRecord.Create(Self,Header,PBuf));
      end;
    end;
  end;
end;

procedure TChartRecord.ReadDefault(DefRecs: array of TDefaultRecord; var Index: integer; Fonts: TXc12Fonts);
var
  i: integer;
  Rec: TChartRecord;
  Header: TBIFFHeader;
begin
  while Index <= High(DefRecs) do begin
    Header.RecID := DefRecs[Index].Id;
    Header.Length := System.Length(DefRecs[Index].Data);
    case Header.RecId of
      BIFFRECID_EOF: begin
        Add(TChartRecord.Create(Self,Header,Nil));
        Exit;
      end;
      CHARTRECID_BEGIN: begin
        Inc(Index);
        if Count > 0 then
          Items[Count - 1].ReadDefault(DefRecs,Index,Fonts);
//        FChilds.ReadDefault(Index);
      end;
      CHARTRECID_END:
        Exit;
      CHARTRECID_FBI: begin
        Add(TChartRecord.Create(Self,Header,PByteArray(@DefRecs[Index].Data[1])));
        Rec := Items[Count - 1];
        PCRecFBI_Font(Rec.FData).Xc12Font := Fonts.Add;
        TXc12Font(PCRecFBI_Font(Rec.FData).Xc12Font).Locked := True;
      end;
      CHARTRECID_FONTX: begin
        Add(TChartRecord.Create(Self,Header.RecId,FindFBIFont(PCRecFONTX(@DefRecs[Index].Data[1]).FontIndex)));
      end;
      ID_CHARTRECORD_DEFAULTLEGEND: begin
        i := 0;
        ReadDefault(DefaultRecordsLegend,i,Fonts);
      end;
      else begin
        if Header.Length > 0 then
          Add(TChartRecord.Create(Self,Header,PByteArray(@DefRecs[Index].Data[1])))
        else
          Add(TChartRecord.Create(Self,Header,Nil));
      end;
    end;
    Inc(Index);
  end;
end;

procedure TChartRecord.Write(Stream: TXLSStream);
var
  i,j: integer;
  Buf: PByteArray;
  Shape: TShapeOutsideMsoChart;
  SheetCharts: TSheetCharts;
  FontTbl: array of integer;
begin
  for i := 0 to Count - 1 do begin
    case Items[i].FRecId of
      CHARTRECID_GELFRAME: begin
        if (Items[i] is TChartRecordUpdate) and Assigned(TChartRecordUpdate(Items[i]).OnUpdate) then
          TChartRecordUpdate(Items[i]).OnUpdate(Self);
        Stream.WriteCONTINUE(CHARTRECID_GELFRAME,Items[i].FData^,Items[i].Length);
      end;
      BIFFRECID_MSODRAWING: begin
        GetMem(Buf,MAXRECSZ_97);
        try
          if Items[i].Root.Parent.ClassName = 'TShapeOutsideMsoChart' then begin
            Shape := TShapeOutsideMsoChart(Items[i].Root.Parent);
            Shape.Drawing.SaveToStream(Stream,Buf);
          end
          else if Items[i].Root.Parent.ClassName = 'TSheetCharts' then begin
            SheetCharts := TSheetCharts(Items[i].Root.Parent);
            SheetCharts[SheetCharts.CurrIndex].Drawing.SaveToStream(Stream,Buf);
          end
          else
            raise XLSRWException.Create('Chart record parent is not of excpected type');
        finally
          FreeMem(Buf);
        end;
      end;
      CHARTRECID_FBI: begin
        Stream.WWord(Items[i].FRecId);
        PCRecFBI(Items[i].FData).Index := TXc12Font(PCRecFBI_Font(Items[i].FData).Xc12Font).Index;
        Stream.WWord(SizeOf(TCRecFBI));
        Stream.Write(Items[i].FData^,SizeOf(TCRecFBI));

        SetLength(FontTbl,System.Length(FontTbl) + 1);
        FontTbl[High(FontTbl)] := PCRecFBI(Items[i].FData).Index;
      end;
      CHARTRECID_FONTX: begin
        for j := 0 to High(FontTbl) do begin
          if TXc12Font(Items[i].Data).Index = FontTbl[j] then begin
            Stream.WWord(Items[i].FRecId);
            Stream.WWord(2);
            Stream.WWord(0);
            Stream.WWord(j);
          end;
        end;
      end;
      else begin
        Stream.WWord(Items[i].FRecId);
        Stream.WWord(Items[i].FLength);
        if Items[i].FLength > 0 then
          Stream.Write(Items[i].FData^,Items[i].FLength);
      end;
    end;
    if Items[i].Count > 0 then begin
      Stream.WriteHeader(CHARTRECID_BEGIN,0);
      Items[i].Write(Stream);
      Stream.WriteHeader(CHARTRECID_END,0);
    end;
  end;
end;

procedure TChartRecord.WriteToFile(Filename: AxUCString; Fonts: TXc12Fonts; AsText: boolean);
var
  b: byte;
  S: string;
  i: integer;
//  Sz: integer;
  Stream: TXLSStream;
  Buf: PByteArray;
  Lines: TStringList;
begin
  Stream := TXLSStream.Create(Nil);
  try
    if AsText then
      Stream.OpenRawMemStreamWrite
    else
      Stream.OpenRawFileStreamWrite(Filename);
    for i := 0 to Count - 1 do begin
      if Items[i].FRecId = CHARTRECID_FBI then begin
        GetMem(Buf,MAXRECSZ_97);
        try
          raise XLSRWException.Create('TODO Write font');
//          Sz := TXc12Font(PCRecFBI_Font(Items[i].FData).Xc12Font).CopyToBuf(Buf,xvExcel97);
//          Stream.WriteHeader(BIFFRECID_FONT,Sz);
//          Stream.Write(Buf^,Sz);
        finally
          FreeMem(Buf);
        end;
      end;
    end;
    Write(Stream);
    if AsText then begin
      Lines := TStringList.Create;
      Lines.Add(' ');
      try
        i := 0;
        S := '';
        Stream.Seek(0,soFromBeginning);
        while Stream.Read(b,1) = 1 do begin
          S := S + Format('$%.2X,',[b]);
          if System.Length(S) > 80 then begin
            Lines.Add(S);
            S := '';
          end;
          Inc(i);
        end;
        if S <> '' then
          S := Copy(S,1,System.Length(S) - 1);
        Lines.Add(S + ');');
        Lines[0] := Format('ChartBinData: array[0..%d] = (',[i - 1]);
        Lines.SaveToFile(Filename);
      finally
        Lines.Free;
      end;
    end;
  finally
    Stream.Free;
  end;
end;

constructor TChartRecord.Create(Parent: TChartRecord; Header: TBIFFHeader; D: PByteArray);
begin
  inherited Create;
  GetMem(FData,Header.Length);
  if D <> Nil then
    System.Move(D^,FData^,Header.Length);
  FParent := Parent;
  FRecId := Header.RecID;
  FLength := Header.Length;
end;

constructor TChartRecord.Create(Parent: TChartRecord; RecId, Length: word);
begin
  inherited Create;
  GetMem(FData,Length);
  FParent := Parent;
  FRecId := RecID;
  FLength := Length;
end;

constructor TChartRecord.Create(Parent: TChartRecord; RecId: word; Obj: TObject);
begin
  inherited Create;
  FData := PByteArray(Obj);
  FParent := Parent;
  FRecId := RecID;
  FLength := 0;
end;

destructor TChartRecord.Destroy;
//var
//  Fonts: TXc12Fonts;
begin
  if (RecId = CHARTRECID_FONTX) and (Root.Parent.ClassName = 'TShapeOutsideMsoChart') then begin
//    Fonts := TShapeOutsideMsoChart(Root.Parent).Fonts;

//    if not TXc12Font(FData).Locked and (TXFont(FData).UsageCount <= 0) and (TXFont(FData).Index < Fonts.Count) then
//      Fonts.Delete(TXFont(FData).Index);

  end
  else if FLength > 0 then begin
    FreeMem(FData);
    FLength := 0;
  end;
  FData := Nil;
  inherited;
end;

procedure TChartRecord.Resize(Delta: integer);
begin
  Inc(FLength,Delta);
  ReAllocMem(FData,FLength);
end;

procedure TChartRecord.ReadDefaultRecords(RecData: TDefaultRecordData);
var
  i: integer;
  Fonts: TXc12Fonts;
begin
  Fonts := TShapeOutsideMsoChart(Root.Parent).Fonts;
  i := 0;
  case RecData of
    drdAll:            ReadDefault(DefaultRecordsAll,i,Fonts);
    drdLegend:         ReadDefault(DefaultRecordsLegend,i,Fonts);
    drdSerie: begin
      i := DefaultRecordsSerieIndex;
      ReadDefault(DefaultRecordsAll,i,Fonts);
    end;
    drdDataformat:     ReadDefault(DefaultRecordsDataformat,i,Fonts);
    drdStyleArea:      ReadDefault(DefaultRecordsStyleArea,i,Fonts);
    drdStyleBarColumn: ReadDefault(DefaultRecordsStyleBarColumn,i,Fonts);
    drdStyleLine:      ReadDefault(DefaultRecordsStyleLine,i,Fonts);
    drdStylePie:       ReadDefault(DefaultRecordsStylePie,i,Fonts);
    drdStyleRadar:     ReadDefault(DefaultRecordsStyleRadar,i,Fonts);
    drdStyleScatter:   ReadDefault(DefaultRecordsStyleScatter,i,Fonts);
    drdStyleSurface:   ReadDefault(DefaultRecordsStyleSurface,i,Fonts);
  end;
end;

procedure TChartRecord.FixupFBI;
var
  i: integer;
  Rec: TChartRecord;
begin
  FFontIdList.Reset;
  Rec := Root;
  for i := 0 to Rec.Count - 1 do begin
    if Rec[i].FRecId = CHARTRECID_FBI then
      PCRecFBI(Rec[i].FData).Index := FFontIdList.GetNext(PCRecFBI(Rec[i].FData).Index);
  end;
end;

procedure TChartRecord.ReadFromBuffer(Buf: PByteArray; BufSize: integer; Fonts: TXc12Fonts);
var
  Stream: TXLSStream;
begin
  Stream := TXLSStream.Create(Nil);
  try
    Stream.OpenRawMemStreamRead;
    Stream.Write(Buf^,BufSize);
    Stream.Seek(0,soFromBeginning);
    ReadFromStream(Stream,Fonts);
  finally
    Stream.Free;
  end;
end;

procedure TChartRecord.ReadFromFile(Filename: AxUCString; Fonts: TXc12Fonts);
var
  Stream: TXLSStream;
begin
  Stream := TXLSStream.Create(Nil);
  try
    Stream.OpenRawFileStreamRead(Filename);
    ReadFromStream(Stream,Fonts);
  finally
    Stream.Free;
  end;
end;

procedure TChartRecord.ReadFromStream(Stream: TXLSStream; Fonts: TXc12Fonts);
var
  Buf: PByteArray;
  Header: TBIFFHeader;
begin
  FFontIdList := TFontIdList.Create;
  try
    GetMem(Buf,MAXRECSZ_97);
    try
      while Stream.PeekHeader = BIFFRECID_FONT do begin
        Stream.ReadHeader(Header);
        Stream.Read(Buf^,Header.Length);
        raise XLSRWException.Create('TODO Read font');
//        FFontIdList.AddFont(Fonts.Add.CopyFromBuf(Buf));
      end;
      Read(Stream,Buf,Fonts);
    finally
      FreeMem(Buf);
    end;
    FixupFBI;
  finally
    FFontIdList.Free;
    FFontIdList := Nil;
  end;
end;

function TChartRecord.RemoveRecord(Id: integer): boolean;
var
  i: integer;
begin
  i := FindRecord(Id);
  Result := i >= 0;
  if Result then
    Delete(i);
end;

function TChartRecord.Root: TChartRecord;
begin
  Result := Self;
  while (Result <> Nil) and (Result.FRecId <> ID_CHARTRECORDROOT) do
    Result := Result.FParent;
  if Result.FRecId <> ID_CHARTRECORDROOT then
    Result := Nil;
end;

{ TFontIdList }

procedure TFontIdList.AddFont(Index: integer);
var
  Data: TFontIdListData;
begin
  Data := TFontIdListData.Create;
  Data.FFontIndex := Index;
  inherited Add(Data);
end;

function TFontIdList.FindFBI(Id: integer): integer;
var
  i: integer;
begin
  for i := 0 to Count - 1 do begin
    if Items[i].FFBIIndex = Id then begin
      Result := Items[i].FFontIndex;
      Exit;
    end;
  end;
  Result := -1;
end;

function TFontIdList.GetItems(Index: integer): TFontIdListData;
begin
  Result := TFontIdListData(inherited Items[Index]);
end;

function TFontIdList.GetNext(FBI: integer): integer;
begin
  Result := Items[FCurrent].FFontIndex;
  Items[FCurrent].FFBIIndex := FBI;
  Inc(FCurrent);
end;

procedure TFontIdList.Reset;
begin
  FCurrent := 0;
end;

end.
