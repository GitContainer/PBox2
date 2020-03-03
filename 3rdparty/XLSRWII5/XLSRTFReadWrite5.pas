unit XLSRTFReadWrite5;

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
{$ifdef BABOON}
  {$ifdef DELPHI_XE5_OR_LATER}
     FMX.Graphics,
  {$else}
     FMX.Types,
  {$endif}
{$else}
     vcl.Graphics,
{$endif}
{$ifdef DELPHI_6_OR_LATER}
     StrUtils,
{$endif}
{$ifdef DELPHI_XE3_OR_LATER}
     System.UITypes,
{$endif}
     XLSUtils5, XLSTools5;

type TIndexFont = class(TFont)
private
     FIndex: integer;
public
     function Equal(F: TIndexFont): boolean;
     function AsRTF: AxUCString;

     property Index: integer read FIndex write FIndex;
     end;

type TIndexFontList = class(TObjectList)
private
     function GetItems(Index: integer): TIndexFont;
public
     function AddFont: TIndexFont;

     property Items[Index: integer]: TIndexFont read GetItems; default;
     end;

type TCHP = record
    fBold: byte;
    fUnderline: byte;
    fItalic: byte;
    FontSize: integer;
    FontColor: TColor;
    FontName: AxUCString;
    end;                  // Character Properties

type TJust =  (justL, justR, justC, justF);

type TPAP = record
    xaLeft: integer;                 // left indent in twips
    xaRight: integer;                // right indent in twips
    xaFirst: integer;                // first line indent in twips
    JUST: TJust;                     // justification
    end;                             // PAragraph Properties

type TSBK = (sbkNon, sbkCol, sbkEvn, sbkOdd, sbkPg);
type TPGN = (pgDec, pgURom, pgLRom, pgULtr, pgLLtr);

type TSEP = record
    cCols: integer;                  // number of columns
    SBK: TSBK;                       // section break type
    xaPgn: integer;                  // x position of page number in twips
    yaPgn: integer;                  // y position of page number in twips
    pgnFormat: TPGN;                       // how the page number is formatted
    end;                             // SEction Properties

type TDOP = record
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

type TRDS = (rdsNorm, rdsSkip, rdsFontTbl, rdsFont);              // Rtf Destination State
type TRIS = (risNorm, risBin, risHex, risColorTbl);        // Rtf Internal State

type PSAVE = ^TSAVE;
    TSAVE = record
    pNext: PSAVE;         // next save
    CHP: TCHP;
    PAP: TPAP;
    SEP: TSEP;
    DOP: TDOP;
    RDS: TRDS;
    RIS: TRIS;
    end;

// What types of properties are there? (Index into FrgProp).
type TIPROP = (ipropBold, ipropItalic, ipropUnderline, ipropFontSize,
              ipropLeftInd,
              ipropRightInd, ipropFirstInd, ipropCols, ipropPgnX,
              ipropPgnY, ipropXaPage, ipropYaPage, ipropXaLeft,
              ipropXaRight, ipropYaTop, ipropYaBottom, ipropPgnStart,
              ipropSbk, ipropPgnFormat, ipropFacingp, ipropLandscape,
              ipropJust, ipropPard, ipropPlain, ipropSectd,
              ipropMax);

type TACTN = (actnSpec, actnByte, actnWord);
type TPROPTYPE = (propChp, propPap, propSep, propDop);

type TPROP = record
    ACTN: TACTN;              // size of value
    prop: TPROPTYPE;
    PData: PByteArray;
    end;

type TIPFN = (ipfnBin, ipfnHex, ipfnSkipDest);
type TIDEST = (idestPict, idestSkip, idestFont, idestFontTbl, idestColorTbl,idestFontColor, idestCharSet);
type TKWD = (kwdChar, kwdUnicode, kwdUnicodeGroup, kwdDest, kwdProp, kwdSpec);

type TRtfFont = class(TObject)
private
     FId: integer;
     FCharSet: integer;
     FName: AxUCString;
public
     constructor Create(Id: integer);

     property Id: integer read FId write FId;
     property CharSet: integer read FCharSet write FCharSet;
     property Name: AxUCString read FName write FName;
     end;

type TRtfFonts = class(TObjectList)
private
     FSortDirty: boolean;

     procedure DoSort;
     function GetItems(Id: integer): TRtfFont;
     function GetItemByIndex(Index: integer): TRtfFont;
public
     procedure Add(Id: integer);
     function Last: TRtfFont;

     property Items[Id: integer]: TRtfFont read GetItems; default;
     property ItemByIndex[Index: integer]: TRtfFont read GetItemByIndex;
     end;

type TRTFStackItem = class(TObject)
private
     FRDS: TRDS;
     FRIS: TRIS;
     FCHP: TCHP;
     FPAP: TPAP;
     FSEP: TSEP;
     FDOP: TDOP;
public
     constructor Create(RDS: TRDS; RIS: TRIS; CHP: TCHP; PAP: TPAP; SEP: TSEP; DOP: TDOP);

     property RDS: TRDS read FRDS write FRDS;
     property RIS: TRIS read FRIS write FRIS;
     property CHP: TCHP read FCHP write FCHP;
     property PAP: TPAP read FPAP write FPAP;
     property SEP: TSEP read FSEP write FSEP;
     property DOP: TDOP read FDOP write FDOP;
     end;

type TRTFStack = class(TObjectList)
private
     function GetItems(Index: integer): TRTFStackItem;
public
     procedure Push(RDS: TRDS; RIS: TRIS; CHP: TCHP; PAP: TPAP; SEP: TSEP; DOP: TDOP);
     procedure Pop(var RDS: TRDS; var RIS: TRIS; var CHP: TCHP; var PAP: TPAP; var SEP: TSEP; var DOP: TDOP);

     property Items[Index: integer]: TRTFStackItem read GetItems; default;
     end;

type TRTFReader = class(TObject)
private
     FrgProp: array of TPROP;
     FStack: TRTFStack;
     cGroup: integer;
     fSkipDestIfUnk: boolean;
     cbBin: integer;
     lParam: integer;
     FRDS: TRDS;
     FRIS: TRIS;
     FCHP: TCHP;
     FPAP: TPAP;
     FSEP: TSEP;
     FDOP: TDOP;
     FInUnicodeGroup: boolean;
     FOutputText: boolean;
     FIncludeAllText: boolean;
     FText: AxUCString;
     FCurrTextPos,FLastTextPos: integer;
     // Shall not be AxUCString. Only 7-bit ascii is permited in RTF files.
     FStringParam: AxUCString;
     FFonts: TRtfFonts;
     FCurrFont: TIndexFont;
     FFontRuns: TIndexFontList;
     FColorTable: array of TColor;

     function  ecRtfParse(fp: TStream): integer;
     procedure ecPushRtfState;
     function  ecPopRtfState: integer;
     function  ecParseRtfKeyword(fp: TStream): integer;
     function  ecParseChar(ch: AxUCChar): integer;
     function  ecTranslateKeyword(szKeyword: AnsiString; param: integer; fParam: boolean): integer;
     procedure ecPrintChar(ch: AxUCChar);
     function  ecEndGroupAction(RDS: TRDS): integer;
     function  ecApplyPropChange(IPROP: TIPROP; val: integer): integer;
     function  ecChangeDest(IDEST: TIDEST): integer;
     function  ecParseSpecialKeyword(IPFN: TIPFN): integer;
     function  ecParseSpecialProperty(IPROP: TIPROP; val: integer): integer;
     function  ecReadColorTbl(fp: TStream): integer;
     procedure FontChanged;
//     function ecParseHexByte: integer;
public
     constructor Create;
     destructor Destroy; override;
     procedure LoadFromFile(Filename: AxUCString);
     procedure LoadFromStream(Stream: TStream);

     property Text: AxUCString read FText;
     property FontRuns: TIndexFontList read FFontRuns;
     property IncludeAllText: boolean read FIncludeAllText write FIncludeAllText;
     end;

type TXLSRTFWriter = class(TObject)
private
     FRTF: AxUCString;
     FText: TXLSFormattedText;

     procedure BeginHeader;
     procedure WriteFontTable;
     procedure WriteColorTable;
     procedure EndHeader;
     procedure BeginInfo;
     procedure WriteDocText;
     procedure EndInfo;
     function  EncodeText(S: AxUCString): AxUCString;
     function  GetRTF: AxUCString;
     procedure MakeRTF;
public
     constructor Create(AText: TXLSFormattedText);
     destructor Destroy; override;
     procedure SaveToFile(Filename: AxUCString);
     procedure SaveToStream(Stream: TStream);

     property RTF: AxUCString read GetRTF;
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

function CPos(C: AxUCChar; S: AxUCString): integer;
begin
  for Result := 1 to Length(S) do begin
    if S[Result] = C then
      Exit;
  end;
  Result := -1;
end;

function GetFirstWord(C: AxUCChar; var S: AxUCString): AxUCString;
var
  p: integer;
begin
  p := CPos(C,S);
  if p < 1 then begin
    Result := S;
    S := '';
  end
  else begin
    Result := Copy(S,1,p - 1);
    S := Copy(S,p + 1,MAXINT);
  end;
end;

{ TRTFReader }

constructor TRTFReader.Create;

procedure AddProp(ACTN: TACTN; prop: TPROPTYPE; PData: PByteArray);
begin
  SetLength(FrgProp,Length(FrgProp) + 1);
  FrgProp[High(FrgProp)].ACTN := ACTN;
  FrgProp[High(FrgProp)].prop := prop;
  FrgProp[High(FrgProp)].PData := PData;
end;


begin
  FIncludeAllText := True;
  FStack := TRTFStack.Create;
  FFonts := TRtfFonts.Create;
  FCurrFont := TIndexFont.Create;
  FFontRuns := TIndexFontList.Create;
  AddProp(actnByte,propChp,@FCHP.fBold);
  AddProp(actnByte,propChp,@FCHP.fItalic);
  AddProp(actnByte,propChp,@FCHP.fUnderline);
  AddProp(actnWord,propChp,@FCHP.FontSize);
  AddProp(actnWord,propPap,@FPAP.xaLeft);
  AddProp(actnWord,propPap,@FPAP.xaRight);
  AddProp(actnWord,propPap,@FPAP.xaFirst);
  AddProp(actnWord,propSep,@FSEP.cCols);
  AddProp(actnWord,propSep,@FSEP.xaPgn);
  AddProp(actnWord,propSep,@FSEP.yaPgn);
  AddProp(actnWord,propDop,@FDOP.xaPage);
  AddProp(actnWord,propDop,@FDOP.yaPage);
  AddProp(actnWord,propDop,@FDOP.xaLeft);
  AddProp(actnWord,propDop,@FDOP.xaRight);
  AddProp(actnWord,propDop,@FDOP.yaTop);
  AddProp(actnWord,propDop,@FDOP.yaBottom);
  AddProp(actnWord,propDop,@FDOP.pgnStart);
  AddProp(actnByte,propSep,@FSEP.sbk);
  AddProp(actnByte,propSep,@FSEP.pgnFormat);
  AddProp(actnByte,propDop,@FDOP.fFacingp);
  AddProp(actnByte,propDop,@FDOP.fLandscape);
  AddProp(actnByte,propPap,@FPAP.just);
  AddProp(actnSpec,propPap,Nil);
  AddProp(actnSpec,propChp,Nil);
  AddProp(actnSpec,propSep,Nil);
end;

destructor TRTFReader.Destroy;
begin
  FStack.Free;
  FFonts.Free;
  FCurrFont.Free;
  FFontRuns.Free;
  inherited Destroy;
end;

procedure TRTFReader.LoadFromFile(Filename: AxUCString);
var
  Stream: TFileStream;
begin
  Stream := TFileStream.Create(Filename,fmOpenRead);
  try
    LoadFromStream(Stream);
  finally
    Stream.Free;
  end;
end;

procedure TRTFReader.LoadFromStream(Stream: TStream);
var
  Res: integer;
begin
  FText := '';
  FCurrTextPos := 0;
  FLastTextPos := 0;
  FOutputText := False;
  FFonts.Clear;
  SetLength(FColorTable,0);
  Res := ecRtfParse(Stream);
  if Res <> ecOk then
    raise XLSRWException.CreateFmt('Error while reading RTF document (%d)',[Res]);
  FontChanged;
end;

//
// %%Function: ecRtfParse
//
// Step 1:
// Isolate RTF keywords and send them to ecParseRtfKeyword;
// Push and pop state at the start and end of RTF groups;
// Send text to ecParseChar for further processing.
//

function TRTFReader.ecRtfParse(fp: TStream): integer;
var
  val: byte;
  ch: AxUCChar;
  cNibble: integer;
  b: integer;
begin
  cNibble := 2;
  b := 0;
  while fp.Read(val,1) = 1 do begin
    ch := AxUCChar(val);
    if cGroup < 0 then begin
      Result := ecStackUnderflow;
      Exit;
    end;
    if FRIS = risBin then begin                      // if we're parsing binary data, handle it directly
      Result := ecParseChar(AxUCChar(ch));
      if Result <> Integer(ecOK) then
          Exit;
    end
    else begin
      case ch of
        '{': begin
           FOutputText := False;
           FInUnicodeGroup := False;
           ecPushRtfState;
        end;
        '}': begin
           FOutputText := False;
           FInUnicodeGroup := False;
           Result := ecPopRtfState;
           if Result <> ecOK then
             Exit;
          end;
        '\': begin
           Result := ecParseRtfKeyword(fp);
           if (Result = ecOk) and (FRIS = risColorTbl) then
             Result := ecReadColorTbl(fp);
           if Result <> ecOK then
             Exit
          end;
        Char($0D),Char($0A): ;          // cr and lf are noise characters...
        else begin
          if FRIS = risNorm then begin

            Result := ecParseChar(AxUCChar(ch));
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
            if CharInSet(ch,['0'..'9']) then
              Inc(b,Integer(ch) - Ord('0'))
            else begin
              if CharInSet(ch,['a'..'f']) then
                Inc(b,Integer(ch) - Ord('a') + 10)
              else if CharInSet(ch,['A'..'F']) then
                Inc(b,Integer(ch) - Ord('A') + 10)
              else begin
                Result := ecInvalidHex;
                Exit;
              end;
            end;
            Dec(cNibble);
            if cNibble <= 0 then begin
              Result := ecParseChar(AxUCChar(b));
              if Result <> ecOK then
                  Exit;
              cNibble := 2;
              b := 0;
              FRIS := risNorm;
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

//
// %%Function: ecPushRtfState
//
// Save relevant info on a linked list of SAVE structures.
//

procedure TRTFReader.ecPushRtfState;
begin
  FStack.Push(FRDS,FRIS,FCHP,FPAP,FSEP,FDOP);
  FRIS := risNorm;
  Inc(cGroup);
end;

//
// %%Function: ecPopRtfState
//
// If we're ending a destination (that is, the destination is changing),
// call ecEndGroupAction.
// Always restore relevant info from the top of the SAVE list.
//

function TRTFReader.ecPopRtfState: integer;
begin
  if FStack.Count <= 0 then begin
    Result := ecStackUnderflow;
    Exit;
  end;

  if FRDS <> FStack[FStack.Count - 1].RDS then begin
    Result := ecEndGroupAction(FRDS);
    if Result <> ecOK then
      Exit;
  end;
  FStack.Pop(FRDS,FRIS,FCHP,FPAP,FSEP,FDOP);

  Dec(cGroup);
  Result := ecOK;
end;

//
// %%Function: ecParseRtfKeyword
//
// Step 2:
// get a control word (and its associated value) and
// call ecTranslateKeyword to dispatch the control.
//

function TRTFReader.ecParseRtfKeyword(fp: TStream): integer;
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
  if not CharInSet(AxUCChar(ch),['a'..'z','A'..'Z']) then begin
    szKeyword := ch;
    Result := ecTranslateKeyword(szKeyword, 0, fParam);
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
  if (ch in ['0'..'9']) then begin
    fParam := True;         // a digit after the control means we have a parameter
    pch := ch;
    while (fp.Read(ch,1) = 1) and (ch in ['0'..'9']) do
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
  Result := ecTranslateKeyword(szKeyword, param, fParam);
end;

//
// %%Function: ecParseChar
//
// Route the character to the appropriate destination stream.
//

function TRTFReader.ecParseChar(ch: AxUCChar): integer;
begin
  Result := ecOK;
                         // Unicode parts contains some mysterious hex characthers...
  if FOutputText and not (FInUnicodeGroup and (FRIS = risHex)) then
    ecPrintChar(ch);

  if (FRIS = risBin) and ((cbBin - 1) <= 0) then
    FRIS := risNorm;
  case FRDS of
    rdsSkip: ;        // Toss this character.
    rdsNorm: begin
        // Output a character. Properties are valid at this point.
//      ecPrintChar(ch);
    end;
    rdsFont:
      FStringParam := FStringParam + ch;
    else begin
    // handle other destinations....
    end;
  end;
end;

//
// %%Function: ecPrintChar
//
// Send a character to the output file.
//

procedure TRTFReader.ecPrintChar(ch: AxUCChar);
begin
  // Excel don't like $0A
  if ch = #$000D then
    ch := #$000A;
  FText := FText + ch;
  Inc(FCurrTextPos);
end;

{ TRTFStack }

function TRTFStack.GetItems(Index: integer): TRTFStackItem;
begin
  Result := TRTFStackItem(inherited Items[Index]);
end;

procedure TRTFStack.Pop(var RDS: TRDS; var RIS: TRIS; var CHP: TCHP; var PAP: TPAP; var SEP: TSEP; var DOP: TDOP);
var
  i: integer;
begin
  i := Count - 1;
  RDS := Items[i].RDS;
  RIS := Items[i].RIS;
  CHP := Items[i].CHP;
  PAP := Items[i].PAP;
  SEP := Items[i].SEP;
  DOP := Items[i].DOP;
  Delete(i);
end;

procedure TRTFStack.Push(RDS: TRDS; RIS: TRIS; CHP: TCHP; PAP: TPAP; SEP: TSEP; DOP: TDOP);
begin
  inherited Add(TRTFStackItem.Create(RDS,RIS,CHP,PAP,SEP,DOP));
end;

{ TRTFStackItem }

constructor TRTFStackItem.Create(RDS: TRDS; RIS: TRIS; CHP: TCHP; PAP: TPAP; SEP: TSEP; DOP: TDOP);
begin
  FRDS := RDS;
  FRIS := RIS;
  FCHP := CHP;
  FPAP := PAP;
  FSEP := SEP;
  FDOP := DOP;
end;


type TSYM = record
    Keyword: AnsiString;        // RTF keyword
    dflt: integer;              // default value to use
    fPassDflt: boolean;         // true to use default value from this table
    KWD: TKWD;               // base action to take
    idx: integer;               // index into property table if kwd == kwdProp
                            // index into destination table if kwd == kwdDest
                            // character to print if kwd == kwdChar
    end;

// Keyword descriptions
const rgsymRtf: array[0..84] of TSYM = (
//  keyword     dflt    fPassDflt   kwd         idx
(    Keyword: 'plain'; dflt:       1; fPassDflt:      False; KWD:     kwdProp; idx:    Integer(ipropPlain)),
(    Keyword: 'b'; dflt:       1; fPassDflt:      False; KWD:     kwdProp; idx:    Integer(ipropBold)),
(    Keyword: 'ul'; dflt:       1; fPassDflt:      False; KWD:     kwdProp; idx:    Integer(ipropUnderline)),
(    Keyword: 'i'; dflt:       1; fPassDflt:      False; KWD:     kwdProp; idx:    Integer(ipropItalic)),
(    Keyword: 'fs'; dflt:      24; fPassDflt:      False; KWD:     kwdProp; idx:    Integer(ipropFontSize)),
(    Keyword: 'cf'; dflt:      0; fPassDflt:      False; KWD:     kwdDest; idx:    Integer(idestFontColor)),
(    Keyword: 'li'; dflt:      0; fPassDflt:      False; KWD:     kwdProp; idx:    Integer(ipropLeftInd)),
(    Keyword: 'ri'; dflt:      0; fPassDflt:      False; KWD:     kwdProp; idx:    Integer(ipropRightInd)),
(    Keyword: 'fi'; dflt:      0; fPassDflt:      False; KWD:     kwdProp; idx:    Integer(ipropFirstInd)),
(    Keyword: 'cols'; dflt:    1; fPassDflt:      False; KWD:     kwdProp; idx:    Integer(ipropCols)),
(    Keyword: 'sbknone'; dflt: Integer(sbkNon); fPassDflt: True; KWD:      kwdProp; idx:    Integer(ipropSbk)),
(    Keyword: 'sbkcol'; dflt:  Integer(sbkCol); fPassDflt: True; KWD:      kwdProp; idx:    Integer(ipropSbk)),
(    Keyword: 'sbkeven'; dflt: Integer(sbkEvn); fPassDflt: True; KWD:      kwdProp; idx:    Integer(ipropSbk)),
(    Keyword: 'sbkodd'; dflt:  Integer(sbkOdd); fPassDflt: True; KWD:      kwdProp; idx:    Integer(ipropSbk)),
(    Keyword: 'sbkpage'; dflt: Integer(sbkPg); fPassDflt:  True; KWD:      kwdProp; idx:    Integer(ipropSbk)),
(    Keyword: 'pgnx'; dflt:    0; fPassDflt:      False; KWD:     kwdProp; idx:    Integer(ipropPgnX)),
(    Keyword: 'pgny'; dflt:    0; fPassDflt:      False; KWD:     kwdProp; idx:    Integer(ipropPgnY)),
(    Keyword: 'pgndec'; dflt:  Integer(pgDec); fPassDflt:  True; KWD:      kwdProp; idx:    Integer(ipropPgnFormat)),
(    Keyword: 'pgnucrm'; dflt: Integer(pgURom); fPassDflt: True; KWD:      kwdProp; idx:    Integer(ipropPgnFormat)),
(    Keyword: 'pgnlcrm'; dflt: Integer(pgLRom); fPassDflt: True; KWD:      kwdProp; idx:    Integer(ipropPgnFormat)),
(    Keyword: 'pgnucltr'; dflt:Integer(pgULtr); fPassDflt: True; KWD:      kwdProp; idx:    Integer(ipropPgnFormat)),
(    Keyword: 'pgnlcltr'; dflt:Integer(pgLLtr); fPassDflt: True; KWD:      kwdProp; idx:    Integer(ipropPgnFormat)),
(    Keyword: 'qc'; dflt:      Integer(justC); fPassDflt:  True; KWD:      kwdProp; idx:    Integer(ipropJust)),
(    Keyword: 'ql'; dflt:      Integer(justL); fPassDflt:  True; KWD:      kwdProp; idx:    Integer(ipropJust)),
(    Keyword: 'qr'; dflt:      Integer(justR); fPassDflt:  True; KWD:      kwdProp; idx:    Integer(ipropJust)),
(    Keyword: 'qj'; dflt:      Integer(justF); fPassDflt:  True; KWD:      kwdProp; idx:    Integer(ipropJust)),
(    Keyword: 'paperw'; dflt:  12240; fPassDflt:  False; KWD:     kwdProp; idx:    Integer(ipropXaPage)),
(    Keyword: 'paperh'; dflt:  15480; fPassDflt:  False; KWD:     kwdProp; idx:    Integer(ipropYaPage)),
(    Keyword: 'margl'; dflt:   1800; fPassDflt:   False; KWD:     kwdProp; idx:    Integer(ipropXaLeft)),
(    Keyword: 'margr'; dflt:   1800; fPassDflt:   False; KWD:     kwdProp; idx:    Integer(ipropXaRight)),
(    Keyword: 'margt'; dflt:   1440; fPassDflt:   False; KWD:     kwdProp; idx:    Integer(ipropYaTop)),
(    Keyword: 'margb'; dflt:   1440; fPassDflt:   False; KWD:     kwdProp; idx:    Integer(ipropYaBottom)),
(    Keyword: 'pgnstart'; dflt:1; fPassDflt:      True; KWD:      kwdProp; idx:    Integer(ipropPgnStart)),
(    Keyword: 'facingp'; dflt: 1; fPassDflt:      True; KWD:      kwdProp; idx:    Integer(ipropFacingp)),
(    Keyword: 'landscape'; dflt: 1; fPassDflt:      True; KWD:      kwdProp; idx:    Integer(ipropLandscape)),
(    Keyword: 'pard'; dflt: 1; fPassDflt:      True; KWD:      kwdProp; idx:    Integer(ipropPard)),
(    Keyword: 'par'; dflt:     0; fPassDflt:      False; KWD:     kwdChar; idx:    $0A),
(    Keyword: '\0x0a'; dflt:    0; fPassDflt:     False; KWD:     kwdChar; idx:   $0A),
(    Keyword: '\0x0d'; dflt:    0; fPassDflt:     False; KWD:     kwdChar; idx:   $0A),
(    Keyword: 'tab'; dflt:     0; fPassDflt:      False; KWD:     kwdChar; idx:    $09),
(    Keyword: 'ldblquote'; dflt: 0; fPassDflt:      False; KWD:     kwdChar; idx:    Ord('"')),
(    Keyword: 'rdblquote'; dflt: 0; fPassDflt:      False; KWD:     kwdChar; idx:    Ord('"')),
(    Keyword: 'bin'; dflt:     0; fPassDflt:      False; KWD:     kwdSpec; idx:    Integer(ipfnBin)),
(    Keyword: '*'; dflt:        0; fPassDflt:     False; KWD:     kwdSpec; idx:    Integer(ipfnSkipDest)),
(    Keyword: ''''; dflt:        0; fPassDflt:    False; KWD:     kwdSpec; idx:    Integer(ipfnHex)),
(    Keyword: 'u'; dflt:        0; fPassDflt:     False; KWD:     kwdUnicode; idx:    0),
(    Keyword: 'uc'; dflt:        0; fPassDflt:     False; KWD:     kwdUnicodeGroup; idx:    0),
(    Keyword: 'author'; dflt:  0; fPassDflt:      False; KWD:     kwdDest; idx:    Integer(idestSkip)),
(    Keyword: 'buptim'; dflt:  0; fPassDflt:      False; KWD:     kwdDest; idx:    Integer(idestSkip)),
(    Keyword: 'colortbl'; dflt:0; fPassDflt:      False; KWD:     kwdDest; idx:    Integer(idestColorTbl)),
(    Keyword: 'comment'; dflt: 0; fPassDflt:      False; KWD:     kwdDest; idx:    Integer(idestSkip)),
(    Keyword: 'creatim'; dflt: 0; fPassDflt:      False; KWD:     kwdDest; idx:    Integer(idestSkip)),
(    Keyword: 'doccomm'; dflt: 0; fPassDflt:      False; KWD:     kwdDest; idx:    Integer(idestSkip)),
(    Keyword: 'f';       dflt: -1; fPassDflt:     False; KWD:     kwdDest; idx:    Integer(idestFont)),
(    Keyword: 'fonttbl'; dflt: 0; fPassDflt:      False; KWD:     kwdDest; idx:    Integer(idestFontTbl)),
(    Keyword: 'fcharset'; dflt: 0; fPassDflt:      True; KWD:     kwdDest; idx:    Integer(idestCharSet)),
(    Keyword: 'footer'; dflt:  0; fPassDflt:      False; KWD:     kwdDest; idx:    Integer(idestSkip)),
(    Keyword: 'footerf'; dflt: 0; fPassDflt:      False; KWD:     kwdDest; idx:    Integer(idestSkip)),
(    Keyword: 'footerl'; dflt: 0; fPassDflt:      False; KWD:     kwdDest; idx:    Integer(idestSkip)),
(    Keyword: 'footerr'; dflt: 0; fPassDflt:      False; KWD:     kwdDest; idx:    Integer(idestSkip)),
(    Keyword: 'footnote'; dflt:0; fPassDflt:      False; KWD:     kwdDest; idx:    Integer(idestSkip)),
(    Keyword: 'ftncn'; dflt:   0; fPassDflt:      False; KWD:     kwdDest; idx:    Integer(idestSkip)),
(    Keyword: 'ftnsep'; dflt:  0; fPassDflt:      False; KWD:     kwdDest; idx:    Integer(idestSkip)),
(    Keyword: 'ftnsepc'; dflt: 0; fPassDflt:      False; KWD:     kwdDest; idx:    Integer(idestSkip)),
(    Keyword: 'header'; dflt:  0; fPassDflt:      False; KWD:     kwdDest; idx:    Integer(idestSkip)),
(    Keyword: 'headerf'; dflt: 0; fPassDflt:      False; KWD:     kwdDest; idx:    Integer(idestSkip)),
(    Keyword: 'headerl'; dflt: 0; fPassDflt:      False; KWD:     kwdDest; idx:    Integer(idestSkip)),
(    Keyword: 'headerr'; dflt: 0; fPassDflt:      False; KWD:     kwdDest; idx:    Integer(idestSkip)),
(    Keyword: 'info'; dflt:    0; fPassDflt:      False; KWD:     kwdDest; idx:    Integer(idestSkip)),
(    Keyword: 'keywords'; dflt:0; fPassDflt:      False; KWD:     kwdDest; idx:    Integer(idestSkip)),
(    Keyword: 'operator'; dflt:0; fPassDflt:      False; KWD:     kwdDest; idx:    Integer(idestSkip)),
(    Keyword: 'pict'; dflt:    0; fPassDflt:      False; KWD:     kwdDest; idx:    Integer(idestSkip)),
(    Keyword: 'printim'; dflt: 0; fPassDflt:      False; KWD:     kwdDest; idx:    Integer(idestSkip)),
(    Keyword: 'private1'; dflt:0; fPassDflt:      False; KWD:     kwdDest; idx:    Integer(idestSkip)),
(    Keyword: 'revtim'; dflt:  0; fPassDflt:      False; KWD:     kwdDest; idx:    Integer(idestSkip)),
(    Keyword: 'rxe'; dflt:     0; fPassDflt:      False; KWD:     kwdDest; idx:    Integer(idestSkip)),
(    Keyword: 'stylesheet'; dflt:  0; fPassDflt:      False; KWD:     kwdDest; idx:    Integer(idestSkip)),
(    Keyword: 'subject'; dflt: 0; fPassDflt:      False; KWD:     kwdDest; idx:    Integer(idestSkip)),
(    Keyword: 'tc'; dflt:      0; fPassDflt:      False; KWD:     kwdDest; idx:    Integer(idestSkip)),
(    Keyword: 'title'; dflt:   0; fPassDflt:      False; KWD:     kwdDest; idx:    Integer(idestSkip)),
(    Keyword: 'txe'; dflt:     0; fPassDflt:      False; KWD:     kwdDest; idx:    Integer(idestSkip)),
(    Keyword: 'xe'; dflt:      0; fPassDflt:      False; KWD:     kwdDest; idx:    Integer(idestSkip)),
(    Keyword: '{'; dflt:       0; fPassDflt:      False; KWD:     kwdChar; idx:    Ord('{')),
(    Keyword: '}'; dflt:       0; fPassDflt:      False; KWD:     kwdChar; idx:    Ord('}')),
(    Keyword: '\'; dflt:       0; fPassDflt:      False; KWD:     kwdChar; idx:    Ord('\')));

const isymMax = sizeof(rgsymRtf) / sizeof(TSYM);

//
// %%Function: ecApplyPropChange
//
// Set the property identified by _iprop_ to the value _val_.
//
//

function TRTFReader.ecApplyPropChange(iprop: TIPROP; val: integer): integer;
begin
  Result := ecOK;
{
  if iprop = ipropPlain then begin
    FOutputText := cGroup = 1;
    if FOutputText then
      SetDefault;
  end
}
  if Frds = rdsSkip then
    Result := ecOK
  else begin
    case FrgProp[Integer(iprop)].actn of
      actnByte: FrgProp[Integer(iprop)].PData[0] := Byte(Val);
      actnWord: PWordArray(FrgProp[Integer(iprop)].PData)[0] := Word(Val);
      actnSpec: Result := ecParseSpecialProperty(iprop, val);
      else
        Result := ecBadTable;
    end;
    if FrgProp[Integer(iprop)].prop = propChp then begin
      FontChanged;
    end;
  end;
end;

//
// %%Function: ecParseSpecialProperty
//
// Set a property that requires code to evaluate.
//

function TRTFReader.ecParseSpecialProperty(iprop: TIPROP; val: integer): integer;
begin
  Result := ecOK;
  case iprop of
    ipropPard: begin
      FillChar(FPAP,SizeOf(TPAP),#0);
      if FIncludeAllText then
        FOutputText := True;
    end;
    ipropPlain:
      FillChar(FCHP,SizeOf(TCHP),#0);
    ipropSectd:
      FillChar(FSEP,SizeOf(TSEP),#0);
    else
      Result := ecBadTable;
  end;
end;

//
// %%Function: ecTranslateKeyword.
//
// Step 3.
// Search rgsymRtf for szKeyword and evaluate it appropriately.
//
// Inputs:
// szKeyword:   The RTF control to evaluate.
// param:       The parameter of the RTF control.
// fParam:      fTrue if the control had a parameter; (that is, if param is valid)
//              fFalse if it did not.
//

function TRTFReader.ecTranslateKeyword(szKeyword: AnsiString; param: integer; fParam: boolean): integer;
var
  i,isym: integer;
begin
  Result := ecBadTable;
  // search for szKeyword in rgsymRtf
  szKeyword := AnsiString(Lowercase(AxUCString(szKeyword)));
  isym := -1;
  for i := 0 to High(rgsymRtf) do begin
    if rgsymRtf[i].Keyword = szKeyword then begin
      isym := i;
      Break;
    end;
  end;

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
      if rgsymRtf[isym].fPassDflt or not fParam then
        param := rgsymRtf[isym].dflt;
      Result := ecApplyPropChange(TIPROP
      (rgsymRtf[isym].idx), param);
    end;
    kwdChar: begin
      FOutputText := True;
      Result := ecParseChar(AxUCChar(rgsymRtf[isym].idx));
    end;
    kwdUnicode:
      Result := ecParseChar(AxUCChar(lParam));
    kwdUnicodeGroup: begin
      FInUnicodeGroup := lParam = 2;
      Result := ecOk;
    end;
    kwdDest: begin
      Result := ecChangeDest(TIDEST(rgsymRtf[isym].idx));
    end;
    kwdSpec:
      Result := ecParseSpecialKeyword(TIPFN(rgsymRtf[isym].idx));
  end;
end;

//
// %%Function: ecChangeDest
//
// Change to the destination specified by idest.
// There's usually more to do here than this...
//

function TRTFReader.ecChangeDest(idest: TIDEST): integer;
begin
  if FRDS = rdsSkip then
    Result := ecOK                // don't do anything
  else begin
    case idest of
      idestPict: begin
        FRDS := rdsSkip;
      end;
      idestFontTbl: begin
        FRDS := rdsFontTbl;
      end;
      idestFont: begin
        if FRDS = rdsFontTbl then begin
          FRDS := rdsFont;
          FStringParam := '';
          FFonts.Add(lParam);
        end
        else begin
          // Not sure if this is correct, but it solves problems with missing fonts.
          if FFonts.Items[lParam] = Nil then
            FFonts.Add(lParam);

          FCHP.FontName := FFonts[lParam].Name;
          FontChanged;
          // Not sure if this is the correct way to detect document text,
          // but all text seems to be receded with a font.
          FOutputText := True;
        end;
      end;
      idestColorTbl:
        FRIS := risColorTbl;
      idestFontColor: begin
        if lParam <= High(FColorTable) then begin
          FCHP.FontColor := FColorTable[lParam];
          FontChanged;
        end
        else begin
          Result := ecAssertion;
          Exit;
        end;
      end;
      idestCharSet: begin
        if FRDS = rdsFont then begin
          FFonts.Last.CharSet := lParam;
        end;
      end;
      else
        FRDS := rdsSkip;              // when in doubt, skip it...
    end;
    Result := ecOK;
  end;
end;

//
// %%Function: ecEndGroupAction
//
// The destination specified by rds is coming to a close.
// If there's any cleanup that needs to be done, do it now.
//

function TRTFReader.ecEndGroupAction(rds: TRDS): integer;
begin
  if RDS = rdsFont then begin
    if RightStr(FStringParam,1) = ';' then
      FStringParam := Copy(FStringParam,1,Length(FStringParam) -1);
    FFonts.Last.Name := FStringParam;
  end;
  Result := ecOK;
end;

//
// %%Function: ecParseSpecialKeyword
//
// Evaluate an RTF control that needs special processing.
//

function TRTFReader.ecParseSpecialKeyword(ipfn: TIPFN): integer;
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

{ TRtfFont }

constructor TRtfFont.Create(Id: integer);
begin
  FId := Id;
end;

{ TRtfFonts }

procedure TRtfFonts.Add(Id: integer);
begin
  inherited Add(TRtfFont.Create(Id));
  FSortDirty := True;
end;

function RtfFontsComparer(AItem1, AItem2: Pointer): integer;
begin
  Result := TRtfFont(AItem1).FId - TRtfFont(AItem2).FId;
end;

procedure TRtfFonts.DoSort;
begin
  Sort(RtfFontsComparer);
  FSortDirty := False;
end;

function TRtfFonts.GetItemByIndex(Index: integer): TRtfFont;
begin
  Result := TRtfFont(inherited Items[Index]);
end;

function TRtfFonts.GetItems(Id: integer): TRtfFont;
var
  i,First,Last: Integer;
begin
  if FSortDirty then
    DoSort;

  First := 0;
  Last := Count - 1;

  while First <= Last do begin
    i := (First + Last) div 2;
    if ItemByIndex[i].FId = id then begin
      Result := ItemByIndex[i];
      Exit;
    end;
    if ItemByIndex[i].FId > Id then
      Last := i - 1
    else
      First := i + 1;
  end;
  Result := Nil;
end;

function TRtfFonts.Last: TRtfFont;
begin
  Result := TRtfFont(inherited Items[Count - 1]);
end;

procedure TRTFReader.FontChanged;
var
  F: TIndexFont;
  Changed: boolean;
begin
  if FCurrTextPos > FLastTextPos then begin
    if (FFontRuns.Count <= 0) or not FFontRuns[FFontRuns.Count - 1].Equal(FCurrFont) then begin
      F := TIndexFont.Create;
      F.Assign(FCurrFont);
      F.Index := FLastTextPos;
      FFontRuns.Add(F);
    end;
    FLastTextPos := FCurrTextPos;
  end;
  Changed := (FCHP.fBold = 1) <> ({$ifdef BABOON}TFontStyle.{$endif}fsBold in FCurrFont.Style);
  if not Changed then
    Changed := (FCHP.fUnderline = 1) <> ({$ifdef BABOON}TFontStyle.{$endif}fsUnderline in FCurrFont.Style);
  if not Changed then
    Changed := (FCHP.fItalic = 1) <> ({$ifdef BABOON}TFontStyle.{$endif}fsItalic in FCurrFont.Style);
  if not Changed then
    Changed := FCHP.FontSize <> (FCurrFont.Size / 2);
{$ifndef BABOON}
  if not Changed then
    Changed := Longword(FCHP.FontColor) <> Longword(FCurrFont.Color);
{$endif}
  if not Changed then
{$ifdef BABOON}
    Changed := AnsiLowercase(FCHP.FontName) <> AnsiLowercase(FCurrFont.Family);
{$else}
    Changed := AnsiLowercase(FCHP.FontName) <> AnsiLowercase(FCurrFont.Name);
{$endif}
  if Changed then begin
    FCurrFont.Style := [];
    if FCHP.fBold = 1 then
      FCurrFont.Style := FCurrFont.Style + [{$ifdef BABOON}TFontStyle.{$endif}fsBold];
    if FCHP.fUnderline = 1 then
      FCurrFont.Style := FCurrFont.Style + [{$ifdef BABOON}TFontStyle.{$endif}fsUnderline];
    if FCHP.fItalic = 1 then
      FCurrFont.Style := FCurrFont.Style + [{$ifdef BABOON}TFontStyle.{$endif}fsItalic];
    FCurrFont.Size := FCHP.FontSize div 2;
{$ifndef BABOON}
    FCurrFont.Color := FCHP.FontColor;
{$endif}
{$ifdef BABOON}
    FCurrFont.Family := FCHP.FontName;
{$else}
    FCurrFont.Name := FCHP.FontName;
{$endif}
  end;
end;

{ TIndexFontList }

function TIndexFontList.AddFont: TIndexFont;
begin
  Result := TIndexFont.Create;
  Add(Result);
end;

function TIndexFontList.GetItems(Index: integer): TIndexFont;
begin
  Result := TIndexFont(inherited Items[Index]);
end;

{ TIndexFont }

function TIndexFont.AsRTF: AxUCString;
begin
  Result := Format('\fs%d',[Size * 2]);
  if {$ifdef BABOON}TFontStyle.{$endif}fsBold in Style then
    Result := Result + '\b1'
  else
    Result := Result + '\b0';
  if {$ifdef BABOON}TFontStyle.{$endif}fsItalic in Style then
    Result := Result + '\i1'
  else
    Result := Result + '\i0' ;
  if {$ifdef BABOON}TFontStyle.{$endif}fsUnderline in Style then
    Result := Result + '\ul1'
  else
    Result := Result + '\ul0';

end;

function TIndexFont.Equal(F: TIndexFont): boolean;
begin
  Result := (Style = F.Style) and (Size = F.Size) and
{$ifdef BABOON}
  (Family = F.Family);
{$else}
  (Color = F.Color) and (Name = F.Name);
{$endif}
end;

function TRTFReader.ecReadColorTbl(fp: TStream): integer;
var
  S,sColor: AxUCString;
  Red,Green,Blue: byte;
  ch: char;

procedure SetColor(S: string);
begin
  if Copy(S,1,3) = 'red' then
    Red := StrToIntDef(Copy(S,4,3),0)
  else if Copy(S,1,5) = 'green' then
    Green := StrToIntDef(Copy(S,6,3),0)
  else if Copy(S,1,4) = 'blue' then
    Blue := StrToIntDef(Copy(S,5,3),0)
end;

begin
  Result := ecOk;
  S := '';
  while fp.Read(ch,1) = 1 do begin
    if ch = '}' then begin
      fp.Seek(-1,soFromCurrent);
      while S <> '' do begin
        Red := 0;
        Green := 0;
        Blue := 0;
        sColor := Copy(GetFirstWord(';',S),2,MAXINT);
        if sColor <> '' then
          SetColor(GetFirstWord('\',sColor));
        if sColor <> '' then
          SetColor(GetFirstWord('\',sColor));
        if sColor <> '' then
          SetColor(sColor);
        SetLength(FColorTable,Length(FColorTable) + 1);
        FColorTable[High(FColorTable)] := (Blue shl 16) + (Green shl 8) + Red;
      end;
      Exit;
    end;
    S := S + ch;
  end;
  Result := ecEndOfFile;
end;

{ TRTFWriter }

procedure TXLSRTFWriter.BeginHeader;
begin
  FRTF := FRTF + '{\rtf1\ansi \ansicpg1252';
end;

procedure TXLSRTFWriter.BeginInfo;
begin
  FRTF := FRTF + '{\info}\plain ';
end;

constructor TXLSRTFWriter.Create(AText: TXLSFormattedText);
begin
  FText := AText;
end;

destructor TXLSRTFWriter.Destroy;
begin
  inherited;
end;

function TXLSRTFWriter.EncodeText(S: AxUCString): AxUCString;
var
  W: word;
  i: integer;
begin
  Result := '';
  for i := 1 to Length(S) do begin
    W := Word(S[i]);
    if W <= $7F then begin
      case W of
        $09:
          Result := Result + '\tab ';
        $0A:
          Result := Result + '\par ';
        else
          Result := Result + AxUCChar(W);
      end;
    end
    else if W <= $FF then
      Result := Result + Format('\''%.2x',[W])
    // The characther after the unicode one, is for old rtf reader. Don't
    // know a meaningsfull value for it.
    else if W <= $7FFF then
      Result := Result + Format('\u%d_',[W])
    else
      Result := Result + Format('\u-%d_',[$FFFF - W + 1]);
  end;
end;

procedure TXLSRTFWriter.EndHeader;
begin

end;

procedure TXLSRTFWriter.EndInfo;
begin
  FRTF := FRTF + '}';
end;

function TXLSRTFWriter.GetRTF: AxUCString;
begin
  MakeRTF;
  Result := FRTF;
end;

procedure TXLSRTFWriter.MakeRTF;
begin
  FRTF := '';
  BeginHeader;
  WriteFontTable;
  WriteColorTable;
  EndHeader;
  BeginInfo;
  WriteDocText;
  EndInfo;
end;

procedure TXLSRTFWriter.SaveToFile(Filename: AxUCString);
var
  Stream: TFileStream;
begin
  Stream := TFileStream.Create(Filename,fmCreate);
  try
    SaveToStream(Stream);
  finally
    Stream.Free;
  end;
end;

procedure TXLSRTFWriter.SaveToStream(Stream: TStream);
begin
  MakeRTF;
  Stream.Write(Pointer(FRTF)^,Length(FRTF));
end;

procedure TXLSRTFWriter.WriteColorTable;
var
  i: integer;
  V: longword;
begin
  // The first color is the default (black).
  FRTF := FRTF + '{\colortbl\red0\green0\blue0;';
  for i := 0 to FText.Count - 1 do begin
    V := FText[i].Font.Color.ARGB;
    FRTF := FRTF + Format('\red%d\green%d\blue%d;',[V and $0000FF,(V shr 8) and $0000FF,(V shr 16) and $0000FF]);
  end;
  FRTF := FRTF + '}';
end;

procedure TXLSRTFWriter.WriteDocText;
var
  i: integer;
begin
  FRTF := FRTF + '\f0' + FText[0].FontAsRTF;
  if FText[0].Font.Color.ARGB <> Longword(clBlack) then
    FRTF := FRTF + '\cf1';
  for i := 1 to FText.Count - 1 do begin
    FRTF := FRTF + ' ' + EncodeText(FText[i].Text);
    FRTF := FRTF + Format('\f%d',[i]) + FText[i].FontAsRTF;
    if FText[0].Font.Color.ARGB <> Longword(clBlack) then
      FRTF := FRTF + Format('\cf%d',[i + 1])
    else
      FRTF := FRTF + '\cf';
  end;
end;

procedure TXLSRTFWriter.WriteFontTable;
var
  i: integer;
begin
  FRTF := FRTF + '{\fonttbl';
  for i := 0 to FText.Count - 1 do begin
    FRTF := FRTF + Format('{\f%d\fnil %s}',[i,FText[i].Font.Name]);
  end;
  FRTF := FRTF + '}';
end;

end.


