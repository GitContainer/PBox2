unit BIFF12_Recs5;

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

//type TRecNames = record
//     Id: integer;
//     Name: string;
//     end;
//
//const RecNames: array[0..130] of TRecNames = (
//(Id: 0 ; Name: 'RowHdr'),
//(Id: 1 ; Name: 'CellBlank'),
//(Id: 2 ; Name: 'CellRk'),
//(Id: 3 ; Name: 'CellError'),
//(Id: 4 ; Name: 'CellBool'),
//(Id: 5 ; Name: 'CellReal'),
//(Id: 6 ; Name: 'CellSt'),
//(Id: 7 ; Name: 'CellIsst'),
//(Id: 8 ; Name: 'FmlaString'),
//(Id: 9 ; Name: 'FmlaNum'),
//(Id: 10; Name: 'FmlaBool'),
//(Id: 11; Name: 'FmlaError'),
//(Id: 12; Name: 'Cell0Blank'),
//(Id: 13; Name: 'Cell0Rk'),
//(Id: 15; Name: 'Cell0Bool'),
//(Id: 16; Name: 'Cell0Real'),
//(Id: 17; Name: 'Cell0St'),
//(Id: 18; Name: 'Cell0Isst'),
//(Id: 49; Name: 'CellMeta'),
//(Id: 50; Name: 'ValueMeta'),
//(Id: 60; Name: 'ColInfo'),
//(Id: 61; Name: 'CellRString'),
//(Id: 62; Name: 'Cell0RString'),
//(Id: 64; Name: 'DVal'),
//(Id: 129 ; Name: 'BeginSheet'),
//(Id: 130 ; Name: 'EndSheet'),
//(Id: 133 ; Name: 'BeginWsViews'),
//(Id: 134 ; Name: 'EndWsViews'),
//(Id: 137 ; Name: 'BeginWsView'),
//(Id: 138 ; Name: 'EndWsView'),
//(Id: 145 ; Name: 'BeginSheetData'),
//(Id: 146 ; Name: 'EndSheetData'),
//(Id: 147 ; Name: 'WsProp'),
//(Id: 148 ; Name: 'WsDim'),
//(Id: 151 ; Name: 'Pane'),
//(Id: 151 ; Name: 'Pane'),
//(Id: 152 ; Name: 'Sel'),
//(Id: 152 ; Name: 'Sel'),
//(Id: 176 ; Name: 'MergeCell'),
//(Id: 177 ; Name: 'BeginMergeCells'),
//(Id: 178 ; Name: 'EndMergeCells'),
//(Id: 307 ; Name: 'BeginSXSELECT'),
//(Id: 308 ; Name: 'EndSXSELECT'),
//(Id: 390 ; Name: 'BeginColInfos'),
//(Id: 391 ; Name: 'EndColInfos'),
//(Id: 392 ; Name: 'BeginRwBrk'),
//(Id: 392 ; Name: 'BeginRwBrk'),
//(Id: 393 ; Name: 'EndRwBrk'),
//(Id: 393 ; Name: 'EndRwBrk'),
//(Id: 394 ; Name: 'BeginColBrk'),
//(Id: 394 ; Name: 'BeginColBrk'),
//(Id: 395 ; Name: 'EndColBrk'),
//(Id: 395 ; Name: 'EndColBrk'),
//(Id: 396 ; Name: 'Brk'),
//(Id: 396 ; Name: 'Brk'),
//(Id: 396 ; Name: 'Brk'),
//(Id: 396 ; Name: 'Brk'),
//(Id: 422 ; Name: 'BeginUserShViews'),
//(Id: 423 ; Name: 'BeginUserShView'),
//(Id: 424 ; Name: 'EndUserShView'),
//(Id: 425 ; Name: 'EndUserShViews'),
//(Id: 426 ; Name: 'ShrFmla'),
//(Id: 427 ; Name: 'ArrFmla'),
//(Id: 428 ; Name: 'Table'),
//(Id: 461 ; Name: 'BeginConditionalFormatting'),
//(Id: 462 ; Name: 'EndConditionalFormatting'),
//(Id: 463 ; Name: 'BeginCFRule'),
//(Id: 464 ; Name: 'EndCFRule'),
//(Id: 465 ; Name: 'BeginIconSet'),
//(Id: 466 ; Name: 'EndIconSet'),
//(Id: 467 ; Name: 'BeginDataBar'),
//(Id: 468 ; Name: 'EndDataBar'),
//(Id: 469 ; Name: 'BeginColorScale'),
//(Id: 470 ; Name: 'EndColorScale'),
//(Id: 471 ; Name: 'CFVO'),
//(Id: 476 ; Name: 'Margins'),
//(Id: 477 ; Name: 'PrintOptions'),
//(Id: 477 ; Name: 'PrintOptions'),
//(Id: 478 ; Name: 'PageSetup'),
//(Id: 478 ; Name: 'PageSetup'),
//(Id: 485 ; Name: 'WsFmtInfo'),
//(Id: 494 ; Name: 'Hlink'),
//(Id: 495 ; Name: 'BeginDcon'),
//(Id: 496 ; Name: 'EndDcon'),
//(Id: 497 ; Name: 'BeginDrefs'),
//(Id: 498 ; Name: 'EndDrefs'),
//(Id: 499 ; Name: 'Dref'),
//(Id: 500 ; Name: 'BeginScenMan'),
//(Id: 501 ; Name: 'EndScenMan'),
//(Id: 502 ; Name: 'BeginSct'),
//(Id: 503 ; Name: 'EndSct'),
//(Id: 504 ; Name: 'Slc'),
//(Id: 535 ; Name: 'SheetProtection'),
//(Id: 536 ; Name: 'RangeProtection'),
//(Id: 537 ; Name: 'PhoneticInfo'),
//(Id: 550 ; Name: 'Drawing'),
//(Id: 551 ; Name: 'LegacyDrawing'),
//(Id: 552 ; Name: 'LegacyDrawingHF'),
//(Id: 554 ; Name: 'BeginWebPubItems'),
//(Id: 555 ; Name: 'EndWebPubItems'),
//(Id: 556 ; Name: 'BeginWebPubItem'),
//(Id: 557 ; Name: 'EndWebPubItem'),
//(Id: 562 ; Name: 'Bkhim'),
//(Id: 564 ; Name: 'Color'),
//(Id: 564 ; Name: 'Color'),
//(Id: 573 ; Name: 'BeginDVals'),
//(Id: 574 ; Name: 'EndDVals'),
//(Id: 589 ; Name: 'CellSmartTagProperty'),
//(Id: 590 ; Name: 'BeginCellSmartTag'),
//(Id: 591 ; Name: 'EndCellSmartTag'),
//(Id: 592 ; Name: 'BeginCellSmartTags'),
//(Id: 593 ; Name: 'EndCellSmartTags'),
//(Id: 594 ; Name: 'BeginSmartTags'),
//(Id: 595 ; Name: 'EndSmartTags'),
//(Id: 605 ; Name: 'BeginCellWatches'),
//(Id: 606 ; Name: 'EndCellWatches'),
//(Id: 607 ; Name: 'CellWatch'),
//(Id: 625 ; Name: 'BigName'),
//(Id: 638 ; Name: 'BeginOleObjects'),
//(Id: 639 ; Name: 'OleObject'),
//(Id: 640 ; Name: 'EndOleObjects'),
//(Id: 643 ; Name: 'BeginActiveXControls'),
//(Id: 644 ; Name: 'ActiveX'),
//(Id: 645 ; Name: 'EndActiveXControls'),
//(Id: 648 ; Name: 'BeginCellIgnoreECs'),
//(Id: 649 ; Name: 'CellIgnoreEC'),
//(Id: 650 ; Name: 'EndCellIgnoreECs'),
//(Id: 660 ; Name: 'BeginTableParts'),
//(Id: 661 ; Name: 'TablePart'),
//(Id: 662 ; Name: 'EndTableParts'),
//(Id: 663 ; Name: 'SheetCalcProp'));


const BIFFRecNames: array[0..663] of string = (
'RowHdr',
'CellBlank',
'CellRk',
'CellError',
'CellBool',
'CellReal',
'CellSt',
'CellIsst',
'FmlaString',
'FmlaNum',
'FmlaBool',
'FmlaError',
'Cell0Blank',
'Cell0Rk',
'',
'Cell0Bool',
'Cell0Real',
'Cell0St',
'Cell0Isst',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'CellMeta',
'ValueMeta',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'ColInfo',
'Cell0RString',
'CellRString',
'',
'DVal',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'BeginSheet',
'EndSheet',
'',
'',
'BeginWsViews',
'EndWsViews',
'',
'',
'BeginWsView',
'EndWsView',
'',
'',
'',
'',
'',
'',
'BeginSheetData',
'EndSheetData',
'WsProp',
'WsDim',
'',
'',
'Pane',
'Sel',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'MergeCell',
'BeginMergeCells',
'EndMergeCells',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'BeginSXSELECT',
'EndSXSELECT',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'BeginColInfos',
'EndColInfos',
'BeginRwBrk',
'EndRwBrk',
'BeginColBrk',
'EndColBrk',
'Brk',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'BeginUserShViews',
'BeginUserShView',
'EndUserShView',
'EndUserShViews',
'ShrFmla',
'ArrFmla',
'Table',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'BeginConditionalFormatting',
'EndConditionalFormatting',
'BeginCFRule',
'EndCFRule',
'BeginIconSet',
'EndIconSet',
'BeginDataBar',
'EndDataBar',
'BeginColorScale',
'EndColorScale',
'CFVO',
'',
'',
'',
'',
'Margins',
'PrintOptions',
'PageSetup',
'',
'',
'',
'',
'',
'',
'WsFmtInfo',
'',
'',
'',
'',
'',
'',
'',
'',
'Hlink',
'BeginDcon',
'EndDcon',
'BeginDrefs',
'EndDrefs',
'Dref',
'BeginScenMan',
'EndScenMan',
'BeginSct',
'EndSct',
'Slc',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'SheetProtection',
'RangeProtection',
'PhoneticInfo',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'Drawing',
'LegacyDrawing',
'LegacyDrawingHF',
'',
'BeginWebPubItems',
'EndWebPubItems',
'BeginWebPubItem',
'EndWebPubItem',
'',
'',
'',
'',
'Bkhim',
'',
'Color',
'',
'',
'',
'',
'',
'',
'',
'',
'BeginDVals',
'EndDVals',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'CellSmartTagProperty',
'BeginCellSmartTag',
'EndCellSmartTag',
'BeginCellSmartTags',
'EndCellSmartTags',
'BeginSmartTags',
'EndSmartTags',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'BeginCellWatches',
'EndCellWatches',
'CellWatch',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'BigName',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'BeginOleObjects',
'OleObject',
'EndOleObjects',
'',
'',
'BeginActiveXControls',
'ActiveX',
'EndActiveXControls',
'',
'',
'BeginCellIgnoreECs',
'CellIgnoreEC',
'EndCellIgnoreECs',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'BeginTableParts',
'TablePart',
'EndTableParts',
'SheetCalcProp');

const BIFF12_REC_ROWHDR           = 0;
const BIFF12_REC_CELLRK           = 2;
const BIFF12_REC_CELLERROR        = 3;
const BIFF12_REC_CELLBOOL         = 4;
const BIFF12_REC_CELLREAL         = 5;
const BIFF12_REC_CELLST           = 6;
const BIFF12_REC_FMLASTRING       = 8;
const BIFF12_REC_FMLANUM          = 9;
const BIFF12_REC_FMLABOOL         = 10;
const BIFF12_REC_FMLAERROR        = 11;
const BIFF12_REC_CELLRSTRING      = 62;
const BIFF12_REC_BEGINSHEET       = 129;
const BIFF12_REC_BEGINSHEETDATA   = 145;
const BIFF12_REC_WSDIM            = 148;

type TRK = packed record
     case integer of
       0: (V: double);
       1: (DW: array [0..1] of longint);
     end;

type PBIFF12_RowHdr = ^TBIFF12_RowHdr;
     TBIFF12_RowHdr = packed record
     Row   : integer;
     Style : integer;
     Height: word;
     end;

type PBIFF12_CellRk = ^TBIFF12_CellRk;
     TBIFF12_CellRk = packed record
     Col: integer;
     Options: longword;
     RK: longword;
     end;

type PBIFF12_CellError = ^TBIFF12_CellError;
     TBIFF12_CellError = packed record
     Col: integer;
     Options: longword;
     Error: byte;
     end;

type PBIFF12_CellBool = ^TBIFF12_CellBool;
     TBIFF12_CellBool = packed record
     Col: integer;
     Options: longword;
     Bool: byte;
     end;

type PBIFF12_CellReal = ^TBIFF12_CellReal;
     TBIFF12_CellReal = packed record
     Col: integer;
     Options: longword;
     Num: double;
     end;

type PBIFF12_FmlaNum = ^TBIFF12_FmlaNum;
     TBIFF12_FmlaNum = packed record
     Col: integer;
     Options: longword;
     Num: double;
     FmlOptions: word;
     FMLA: array[0..0] of byte;
     end;

type PBIFF12_FmlaString = ^TBIFF12_FmlaString;
     TBIFF12_FmlaString = packed record
     Col: integer;
     Options: longword;
     Length: integer;
     Str: array[0..0] of Char;
     end;

type PBIFF12_FmlaBool = ^TBIFF12_FmlaBool;
     TBIFF12_FmlaBool = packed record
     Col: integer;
     Options: longword;
     Bool: byte;
     FmlOptions: word;
     FMLA: array[0..0] of byte;
     end;

type PBIFF12_FmlaError = ^TBIFF12_FmlaError;
     TBIFF12_FmlaError = packed record
     Col: integer;
     Options: longword;
     Error: byte;
     FmlOptions: word;
     FMLA: array[0..0] of byte;
     end;

type PBIFF12_CellRString = ^TBIFF12_CellRString;
     TBIFF12_CellRString = packed record
     Col: integer;
     Options: longword;
     StrOptions: byte;
     Length: integer;
     Str: array[0..0] of Char;
     end;

type PBIFF12_WsDim = ^TBIFF12_WsDim;
     TBIFF12_WsDim = packed record
     Row1: longword;
     Row2: longword;
     Col1: longword;
     Col2: longword;
     end;

implementation

end.
