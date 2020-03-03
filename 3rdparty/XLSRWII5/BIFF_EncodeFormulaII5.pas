unit BIFF_EncodeFormulaII5;

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

uses SysUtils, Classes, Contnrs,
     BIFF_RecsII5, BIFF_Utils5, BIFF_ExcelFuncII5,
     Xc12Utils5,
     XLSUtils5, XLSFormulaTypes5;

type TUnknownNameEvent = procedure (Name: AxUCString; var ID: integer) of object;
type TUnknownExternNameEvent = procedure (Path,Filename,SheetName,Name: AxUCString; IsRef: boolean; var ExtIndex,NameIndex: integer) of object;
type TTokenErrorEvent = procedure(ErrorId: integer; ErrorStr: AxUCString) of object;

type TTokenType = (ttNone,ttAdd,ttConcat,ttSub,ttMul,ttDiv,ttPercent,ttPower,
                   ttGT,ttGE,ttLT,ttLE,ttEQ,ttNE,
                   ttParenL,ttParenR,ttSpace,ttListSep,
                   ttVectL,ttVectR,
                   ttFunc,ttRef,ttArea,ttRef3d,ttArea3d,ttRefX,ttAreaX,
                   ttNumber,ttString,
                   ttPath,ttFile,
                   ttName,ttNameX,ttSheetName,ttSheetNameX,
                   ttError
                   );

const TTokenName: array[0..Integer(High(TTokenType))] of string =
                  ('ttNone','ttAdd','ttConcat','ttSub','ttMul','ttDiv','ttPercent','ttPower',
                   'ttGT','ttGE','ttLT','ttLE','ttEQ','ttNE',
                   'ttParenL','ttParenR','ttSpace','ttListSep',
                   'ttVectL','ttVectR',
                   'ttFunc','ttRef','ttArea','ttRef3d','ttArea3d','ttRefX','ttAreaX',
                   'ttNumber','ttString',
                   'ttPath','ttFile',
                   'ttName','ttNameX','ttSheetName','ttSheetNameX',
                   'ttError'
                   );

type PXLSToken = ^TXLSToken;
     TXLSToken = record
     Token: AxUCString;
     TokenType: TTokenType;
     end;

type TXLSTokenizer = class(TList)
private
     FValue: AxUCString;
     FOrigValue: AxUCString;
     FLen: integer;
     FError: boolean;
     FTokenErrorEvent: TTokenErrorEvent;
     FFormulaVersion: TFormulaVersion;

     function  GetItems(Index: integer): PXLSToken;
     function  DoName(i: integer): integer;
     function  DoQuotedName(i: integer): integer;
     function  IsRef(i: integer): integer;
     function  AddOp(Op: string; TType: TTokenType): integer;
     function  DoString(i: integer): integer;
     function  DoFilename(i: integer): integer;
     function  DoNumber(i: integer): integer;
     function  DoError(i: integer): integer;
     procedure ChangeLast(TType: TTokenType);
     function  GetLast: TTokenType;
     function  GetLast2: TTokenType;
     function  GetName(Index: integer): AxUCString;
public
     procedure Tokenize(Value: AxUCString);
     procedure Add(Token: AxUCString; TType: TTokenType);
     procedure Clear; override;
     function  NextToken(Index: integer): TTokenType;

     property Items[Index: integer]: PXLSToken read GetItems; default;
     // For debug
     property Name[Index: integer]: AxUCString read GetName;
     property LastError: boolean read FError;
     property Value: AxUCString read FOrigValue;
     property FormulaVersion: TFormulaVersion write FFormulaVersion;

     property OnTokenError: TTokenErrorEvent read FTokenErrorEvent write FTokenErrorEvent;
     end;

type TEncodeFormula = class(TObject)
private
    FTokenizer: TXLSTokenizer;
    FExcelVersion: TExcelVersion;
    FStrTRUE: AxUCString;
    FStrFALSE: AxUCString;
    FBuffer: PByteArray;
    FBufSize,FCurrBufSize: integer;
    FLastSheetIndex: integer;
    FFuncArgClass: byte;
    FLastPtgPos: integer;
    FWSStack: TStack;
    FFormulaVersion: TFormulaVersion;

    FUnknownFuncEvent: TUnknownNameEvent;
    FUnknownNameEvent: TUnknownNameEvent;
    FExternNameEvent: TUnknownExternNameEvent;
    FFormulaErrorEvent: TFormulaErrorEvent;

    procedure AddPtg(Ptg: byte); overload;
    procedure AddPtg(Ptg: byte; const Data; Size: integer); overload;
    procedure AddBuf(const Data; Size: integer);
    procedure AddWS;
    procedure AddConstant(i1: integer);
    procedure AddVectorConst(i1: integer);
    function  AddCellRef(i1: integer): boolean;
    function  Add3dCellRef(i1: integer): boolean;
    function  AddBoolConst(i1: integer): boolean;
    function  AddFunction(var Index: integer): boolean;
    function  AddVariable(i1: integer): boolean;
    function  TokenToPtg(Token: TTokenType): byte;
    function  GetFuncId(S: AxUCString): integer;
    procedure ValueError(Id: integer; S: AxUCString);
    function  GetNextFuncArg(Index: integer): integer;
    procedure AddValue(var Index: integer);
    procedure CheckWS(var Index: integer);
    procedure Expression(var Index: integer; FuncArgClass: byte);
    procedure DoOpCompare(var Index: integer);
    procedure DoOpConcat(var Index: integer);
    procedure DoOpAddSub(var Index: integer);
    procedure DoOpMulDiv(var Index: integer);
    procedure DoOpExp(var Index: integer);
    procedure DoOpPercent(var Index: integer);
    procedure DoOpUnary(var Index: integer);
    procedure DoValue(var Index: integer);
public
    constructor Create;
    destructor Destroy; override;
    function  Encode(Formula: AxUCString; ReturnClass: Byte; var Buf: PByteArray; BufSz: integer; FmlaVersion: TFormulaVersion): integer;

    property ExcelVersion: TExcelVersion read FExcelVersion write FExcelVersion;
    property LastSheetIndex: integer read FLastSheetIndex;

    property OnUnknownFunction: TUnknownNameEvent read FUnknownFuncEvent write FUnknownFuncEvent;
    property _OnUnknownName: TUnknownNameEvent read FUnknownNameEvent write FUnknownNameEvent;
    property OnExternName: TUnknownExternNameEvent read FExternNameEvent write FExternNameEvent;
    property OnFormulaError: TFormulaErrorEvent read FFormulaErrorEvent write FFormulaErrorEvent;
    end;

implementation

// const AC_MASK = $1F;  // Argument class mask

const charQuote1 = '"';
      charQuote2 = '''';

      SymbolFirstChar: AxUCString = '''_abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ';
      RefFirstChar   : AxUCString = '$';
      ExtRefFirstChar: AxUCString = '''[';

      EnclosingChars: AxUCString = '()[]{}';

      ERR_BADCONSTANT        = 1;
      ERR_INVALIDSYMBOL      = 2;
      ERR_PAR_MISS_FUNC      = 3;
      ERR_BAD_FUNCNAME       = 4;
      ERR_STR_MISSING_Q      = 5;
      ERR_STRLEN             = 6;
      ERR_UNKNOWN_NAME       = 7;
      ERR_ENCLOSING_CHAR     = 8;
      ERR_INVALID_FILENAME   = 9;
      ERR_UNARY_COUNT        = 10;
      ERR_MISSMATCH_PARAN    = 11;
      ERR_ILLEGAL_NUMBER     = 12;
      ERR_ILLEGAL_CHAR_IN_STR= 13;
      ERR_BAD_FILENAME       = 14;
      ERR_INVALID_3D_NAME    = 15;
      ERR_INVALID_NAME       = 16;
      ERR_FUNC_WRONG_ARGS    = 17;
      ERR_EMPTY_FORMULA      = 18;

{
function WCRPosLastStop(WC: WideChar; WS,StopChar: AxUCString): integer;
var
  i,j: integer;
  InQuote: boolean;
begin
  Result := -1;
  InQuote := False;
  for i := Length(WS) downto 1 do begin
    if Char(WS[Result]) in [charQuote1,charQuote2] then
      InQuote := not InQuote;
    if not InQuote then begin
      if WS[i] = WC then
        Result := i;
      for j := 1 to Length(StopChar) do begin
        if WS[i] =  StopChar[j] then
          Exit;
      end;
    end;
  end;
end;
}

procedure UppercaseSkipQuote(var S: AxUCString);
var
  i: integer;
  InQuote: boolean;
begin
  InQuote := False;
  for i := 1 to Length(S) do begin
    if CharInSet(S[i],[charQuote1,charQuote2]) then
      InQuote := not InQuote;
    if not InQuote then
      S[i] := MyWideUppercase(S[i])[1];
  end;
end;

{ TEncodeFormula }

function TEncodeFormula.GetNextFuncArg(Index: integer): integer;
var
  Par1,Par2: integer;
begin
  Par1 := 0;
  Par2 := 0;
  for Result := Index to FTokenizer.Count - 1 do begin
    case FTokenizer[Result].TokenType of
      ttFunc,
      ttParenL: Inc(Par1);
      ttParenR: begin
        if Par1 = 0 then
          Break;
        Dec(Par1);
      end;
      ttVectL : Inc(Par2);
      ttVectR : Dec(Par2);
//      '[': Inc(Par3);
//      ']': Dec(Par3);
    end;
    if (Par1 = 0) and (Par2 = 0) and (FTokenizer[Result].TokenType = ttListSep) then
      Exit;
  end;
  Result := -1;
end;

procedure TEncodeFormula.AddBuf(const Data; Size: integer);
begin
  if (FCurrBufSize + Size) > FBufSize then
    raise XLSRWException.Create('Formula buffer overflow.');
  if Size > 0 then begin
    Move(Data,FBuffer[FCurrBufSize],Size);
    Inc(FCurrBufSize,Size);
  end;
end;

procedure TEncodeFormula.AddPtg(Ptg: byte);
begin
  AddWS;
  if (FCurrBufSize + 1) > FBufSize then
    raise XLSRWException.Create('Formula buffer overflow.');
  FBuffer[FCurrBufSize] := Ptg;
  FLastPtgPos := FCurrBufSize;
  Inc(FCurrBufSize);
end;

procedure TEncodeFormula.AddPtg(Ptg: byte; const Data; Size: integer);
begin
  AddPtg(Ptg);
  AddBuf(Data,Size);
end;

procedure TEncodeFormula.CheckWS(var Index: integer);
var
  Count: integer;
begin
  Count := 0;
  while (Index < FTokenizer.Count) and (FTokenizer[Index].TokenType = ttSpace) do begin
    Inc(Count);
    Inc(Index);
  end;
  if Count > 255 then
    Count := 255;
  if (Count > 0) and (Index < FTokenizer.Count) then
    FWSStack.Push(Pointer((Byte(FTokenizer[Index].TokenType) shl 16) or Byte(Count)))
  else
    FWSStack.Push(Pointer(0));
end;

constructor TEncodeFormula.Create;
begin
  FExcelVersion := xvExcel97;
  FTokenizer := TXLSTokenizer.Create;
  FTokenizer.OnTokenError := ValueError;
  FWSStack := TStack.Create;
  FStrFALSE := MyWideUppercase(G_StrFALSE);
  FStrTRUE := MyWideUppercase(G_StrTRUE);
end;

destructor TEncodeFormula.Destroy;
begin
  FTokenizer.Free;
  FWSStack.Free;
  inherited;
end;

procedure TEncodeFormula.DoOpAddSub(var Index: integer);
var
  Token: TTokenType;
begin
  DoOpMulDiv(Index);
  while FTokenizer.NextToken(Index) in [ttAdd,ttSub] do begin
    Token := FTokenizer.NextToken(Index);
    CheckWS(Index);
    Inc(Index);
    DoOpMulDiv(Index);
    AddPtg(TokenToPtg(Token));
  end;
end;

procedure TEncodeFormula.DoOpCompare(var Index: integer);
var
  Token: TTokenType;
begin
  DoOpConcat(Index);
  while FTokenizer.NextToken(Index) in [ttNE,ttLE,ttGE,ttEQ,ttLT,ttGT] do begin
    Token := FTokenizer.NextToken(Index);
    CheckWS(Index);
    Inc(Index);
    DoOpConcat(Index);
    AddPtg(TokenToPtg(Token));
  end;
end;

procedure TEncodeFormula.DoOpConcat(var Index: integer);
begin
  DoOpAddSub(Index);
  while FTokenizer.NextToken(Index) = ttConcat do begin
    CheckWS(Index);
    Inc(Index);
    DoOpAddSub(Index);
    AddPtg(xptgOpConcat);
  end;
end;

procedure TEncodeFormula.DoOpExp(var Index: integer);
begin
  DoOpPercent(Index);
  while FTokenizer.NextToken(Index) = ttPower do begin
    CheckWS(Index);
    Inc(Index);
    DoOpPercent(Index);
    AddPtg(xptgOpPower);
  end;
end;

procedure TEncodeFormula.DoOpMulDiv(var Index: integer);
var
  Token: TTokenType;
begin
  DoOpExp(Index);
  while FTokenizer.NextToken(Index) in [ttMul,ttDiv] do begin
    Token := FTokenizer.NextToken(Index);
    CheckWS(Index);
    Inc(Index);
    DoOpExp(Index);
    AddPtg(TokenToPtg(Token));
  end;
end;

procedure TEncodeFormula.DoOpPercent(var Index: integer);
begin
  DoOpUnary(Index);
  while FTokenizer.NextToken(Index) = ttPercent do begin
    CheckWS(Index);
    Inc(Index);
    if FTokenizer.NextToken(Index) <> ttNone then
      DoOpUnary(Index);
    AddPtg(xptgOpPercent);
  end;
end;

procedure TEncodeFormula.DoOpUnary(var Index: integer);
var
  p1,p2: integer;
begin
  p1 := Index;
  while FTokenizer.NextToken(Index) in [ttAdd,ttSub] do begin
    CheckWS(Index);
    Inc(Index);
  end;
  p2 := Index;
  DoValue(Index);
  while p1 <= p2 do begin
    case FTokenizer[p2].TokenType of
      ttAdd: AddPtg(xptgOpUPlus);
      ttSub: AddPtg(xptgOpUMinus);
    end;
    Dec(p2);
  end;
end;

procedure TEncodeFormula.DoValue(var Index: integer);
begin
  if FTokenizer.NextToken(Index) = ttParenL then begin
    CheckWS(Index);
    Inc(Index);
    DoOpCompare(Index);
    if (Index >= FTokenizer.Count) or (FTokenizer[Index].TokenType <> ttParenR) then
      ValueError(ERR_MISSMATCH_PARAN,'');
    CheckWS(Index);
    Inc(Index);
    AddWS;
    AddPtg(xptgLPar);
  end
  else if FTokenizer.NextToken(Index) <> ttNone then begin
    CheckWS(Index);
    AddValue(Index);
  end;
end;

function TEncodeFormula.Encode(Formula: AxUCString; ReturnClass: Byte; var Buf: PByteArray; BufSz: integer; FmlaVersion: TFormulaVersion): integer;
var
  Index: integer;
begin
  Result := 0;
  FLastSheetIndex := -1;
  FFormulaVersion := FmlaVersion;

  if Trim(Formula) = '' then begin
    ValueError(ERR_EMPTY_FORMULA,'');
    Exit;
  end;

  FTokenizer.FormulaVersion := FmlaVersion;

  FTokenizer.Tokenize(Formula);
  if FTokenizer.LastError then
    raise XLSRWException.Create('Error in formula');

  FBuffer := Buf;
  FBufSize := BufSz;
  FCurrBufSize := 0;
  FLastPtgPos := -1;
//  BeginExpression(0,FTokenizer.Count -1);
//  Scan(0,FTokenizer.Count -1,ReturnClass);
  Index := 0;

  Expression(Index,ReturnClass);
//  if Index < FTokenizer.Count then
//    ValueError(-1,'');

  FTokenizer.Clear;
  Result := FCurrBufSize;
end;

procedure TEncodeFormula.Expression(var Index: integer; FuncArgClass: byte);
begin
  FFuncArgClass := FuncArgClass;
  if FFuncArgClass = 0 then
    FFuncArgClass := TokenClassDefault;
  DoOpCompare(Index);
end;

function TEncodeFormula.TokenToPtg(Token: TTokenType): byte;
begin
  case Token of
    ttPower:   Result := xptgOpPower;
    ttMul:     Result := xptgOpMult;
    ttListSep: Result := xptgOpUnion;
    ttDiv:     Result := xptgOpDiv;
    ttEQ:      Result := xptgOpEQ;
    ttNE:      Result := xptgOpNE;
    ttLT:      Result := xptgOpLT;
    ttGT:      Result := xptgOpGT;
    ttLE:      Result := xptgOpLE;
    ttGE:      Result := xptgOpGE;
    ttAdd:     Result := xptgOpAdd;
    ttConcat:  Result := xptgOpConcat;
    ttSub:     Result := xptgOpSub;
    ttPercent: Result := xptgOpPercent;
    else
      raise XLSRWException.CreateFmt('Can not find ptg token %d',[Integer(Token)]);
  end;
end;

procedure TEncodeFormula.AddConstant(i1: integer);
var
  vDouble: double;
  vInt: integer;
  S: AxUCString;
begin
  if FTokenizer[i1].TokenType = ttString then begin
    S := FTokenizer[i1].Token;
    if Length(S) > 255 then
      ValueError(ERR_STRLEN,'')
    else if S = '' then begin
      vInt := 0;
      AddPtg(xptgStr97,vInt,2);
    end
    else begin
      vInt := Length(S);
      AddPtg(xptgStr97,vInt,1);
      vInt := 1;
      AddBuf(vInt,1);
      AddBuf(Pointer(S)^,Length(S) * 2);
    end;
  end
  else begin
    try
      vDouble := StrToFloat(FTokenizer[i1].Token);
      if (Frac(vDouble) = 0) and (vDouble >= 0) and (vDouble <= $FFFF) then begin
        vInt := Round(vDouble);
        AddPtg(xptgInt97,vInt,2);
      end
      else
        AddPtg(xptgNum,vDouble,SizeOf(double));
    except
      ValueError(ERR_BADCONSTANT,FTokenizer[i1].Token);
    end;
  end;
end;

procedure TEncodeFormula.ValueError(Id: integer; S: AxUCString);
var
  ErrStr: AxUCString;
begin
  case Id of
    ERR_BADCONSTANT:
      ErrStr := Format('Invalid constant %s',[S]);
    ERR_INVALIDSYMBOL:
      ErrStr := Format('Invalid symbol %s',[S]);
    ERR_PAR_MISS_FUNC:
      ErrStr := 'Closing paranthese missing in function';
    ERR_BAD_FUNCNAME:
      ErrStr := Format('Invalid function name %s',[S]);
    ERR_STR_MISSING_Q:
      ErrStr := Format('Missing quote characther in string %s',[S]);
    ERR_STRLEN:
      ErrStr := 'Strings can have max 255 characthers';
    ERR_UNKNOWN_NAME:
      ErrStr := Format('Unknown name %s',[S]);
    ERR_ENCLOSING_CHAR:
      ErrStr := Format('Enclosing characther %s missing',[S]);
    ERR_INVALID_FILENAME:
      ErrStr := Format('Invalid filename %s',[S]);
    ERR_UNARY_COUNT:
      ErrStr := 'To many unary operators';
    ERR_MISSMATCH_PARAN:
      ErrStr := 'Missmatching parantheses';
    ERR_ILLEGAL_NUMBER:
      ErrStr := 'Illegal number';
    ERR_ILLEGAL_CHAR_IN_STR:
      ErrStr := Format('Illegal character in string %s',[S]);
    ERR_BAD_FILENAME:
      ErrStr := Format('Bad filename %s',[S]);
    ERR_INVALID_3D_NAME:
      ErrStr := Format('Invalid 3D name %s',[S]);
    ERR_INVALID_NAME:
      ErrStr := Format('Invalid name %s',[S]);
    ERR_FUNC_WRONG_ARGS:
      ErrStr := Format('Wrong number of arguments for function %s',[S]);
    ERR_EMPTY_FORMULA:
      ErrStr := Format('Empty formula',[S]);
    else
      ErrStr := 'Error in formula';
  end;
  ErrStr := Format('Error in formula %s' + #13,[FTokenizer.Value]) + ErrStr;
  if Assigned(FFormulaErrorEvent) then
    FFormulaErrorEvent(Self,Id,ErrStr)
  else
    raise XLSRWException.Create(ErrStr);
end;

function TEncodeFormula.AddBoolConst(i1: integer): boolean;
var
  V: byte;
begin
  V := $FF;
  if FTokenizer[i1].Token = FStrTRUE then
    V := 1
  else if FTokenizer[i1].Token = FStrFALSE then
    V := 0;
  Result := V < $FF;
  if Result then
    AddPtg(xptgBool,V,1);
end;

// Syntax: 'd:\path\[filename.xls]Sheet 1'!A1
function TEncodeFormula.Add3dCellRef(i1: integer): boolean;
var
  ExtIndex,NameIndex: integer;
  Col1,Row1,Col2,Row2: integer;
  AbsCol1,AbsRow1,AbsCol2,AbsRow2: boolean;
  Ref3d: TPTGRef3d8;
  Area3d: TPTGArea3d8;
  Namex: TPTGNameX8;
begin
  Result := False;
  if i1 >= (FTokenizer.Count - 1) then
    Exit;
  ExtIndex := -1;
  NameIndex := -1;
  case FTokenizer[i1].TokenType of
    ttSheetName: begin
      if (FTokenizer.Count - i1) >= 2 then begin
        FExternNameEvent('','',FTokenizer[i1].Token,FTokenizer[i1 + 1].Token,True,ExtIndex,NameIndex);
        Inc(i1);
      end
      else
        Exit;
    end;
    ttPath: begin
      if (FTokenizer.Count - i1) >= 2 then begin
        if FTokenizer[i1 + 1].TokenType = ttNameX then begin
          FExternNameEvent(FTokenizer[i1].Token,'','',FTokenizer[i1 + 1].Token,False,ExtIndex,NameIndex);
          Inc(i1);
        end
        else if (FTokenizer.Count - i1) >= 4 then begin
          FExternNameEvent(FTokenizer[i1].Token,FTokenizer[i1 + 1].Token,FTokenizer[i1 + 2].Token,FTokenizer[i1 + 3].Token,True,ExtIndex,NameIndex);
          Inc(i1,3);
        end
        else
          Exit;
      end
      else
        Exit;
    end;
  end;
//  FExternNameEvent(Path,Filename,SheetName,Ref,ExtIndex,NameIndex);
  if NameIndex < 0 then
    Exit;
  case FTokenizer[i1].TokenType of
    ttRef3d,ttRefX: begin
      if not RefStrToColRow(FTokenizer[i1].Token,Col1,Row1,AbsCol1,AbsRow1) then
        Exit;
      FLastSheetIndex := NameIndex;
      Ref3d.Index := ExtIndex;
      Ref3d.Col := Col1;
      Ref3d.Row := Row1;
      if not AbsRow1 then
        Ref3d.Col := Ref3d.Col + $8000;
      if not AbsCol1 then
        Ref3d.Col := Ref3d.Col + $4000;
      AddPtg(ptgRef3d_NAC or FFuncArgClass,Ref3d,SizeOf(TPTGRef3d8));
      Result := True;
    end;
    ttArea3d,ttAreaX: begin
      if not AreaStrToColRow(FTokenizer[i1].Token,Col1,Row1,Col2,Row2,AbsCol1,AbsRow1,AbsCol2,AbsRow2) then
        Exit;
      FLastSheetIndex := NameIndex;
      Area3d.Index := ExtIndex;
      Area3d.Col1 := Col1;
      Area3d.Row1 := Row1;
      Area3d.Col2 := Col2;
      Area3d.Row2 := Row2;
      if not AbsRow1 then
        Area3d.Col1 := Area3d.Col1 + $8000;
      if not AbsCol1 then
        Area3d.Col1 := Area3d.Col1 + $4000;
      if not AbsRow2 then
        Area3d.Col2 := Area3d.Col2 + $8000;
      if not AbsCol2 then
        Area3d.Col2 := Area3d.Col2 + $4000;
      AddPtg(ptgArea3d_NAC or FFuncArgClass,Area3d,SizeOf(TPTGArea3d8));
      Result := True;
    end;
    ttNameX: begin
      NameX.ExtSheet := ExtIndex;
      NameX.NameIndex := NameIndex + 1;
      NameX.Reserved := 0;
      AddPtg(ptgNameX_NAC or FFuncArgClass,NameX,SizeOf(TPTGNameX8));
      Result := True;
    end;
  end
end;

function TEncodeFormula.AddCellRef(i1: integer): boolean;
var
  Col1,Row1,Col2,Row2: integer;
  AbsCol1,AbsRow1,AbsCol2,AbsRow2: boolean;
  Area: TPTGArea8;
  Ref: TPTGRef8;
begin
  if Add3dCellRef(i1) then
    Result := True
  else if (FTokenizer[i1].TokenType = ttArea) and AreaStrToColRow(FTokenizer[i1].Token,Col1,Row1,Col2,Row2,AbsCol1,AbsRow1,AbsCol2,AbsRow2) then begin
    Area.Col1 := Col1;
    Area.Row1 := Row1;
    Area.Col2 := Col2;
    Area.Row2 := Row2;
    if not AbsRow1 then
      Area.Col1 := Area.Col1 + $8000;
    if not AbsCol1 then
      Area.Col1 := Area.Col1 + $4000;
    if not AbsRow2 then
      Area.Col2 := Area.Col2 + $8000;
    if not AbsCol2 then
      Area.Col2 := Area.Col2 + $4000;
    AddPtg(ptgArea_NAC or FFuncArgClass,Area,SizeOf(TPTGArea8));
    Result := True;
  end
  else if (FTokenizer[i1].TokenType = ttRef) and RefStrToColRow(FTokenizer[i1].Token,Col1,Row1,AbsCol1,AbsRow1) then begin
    Ref.Col := Col1;
    Ref.Row := Row1;
    if not AbsRow1 then
      Ref.Col := Ref.Col + $8000;
    if not AbsCol1 then
      Ref.Col := Ref.Col + $4000;
    AddPtg(ptgRef_NAC or FFuncArgClass,Ref,SizeOf(TPTGRef8));
    Result := True;
  end
  else
    Result := False;
end;

function TEncodeFormula.AddFunction(var Index: integer): boolean;
var
  p,ArgCount,Id: integer;
  FuncArgClass: byte;
begin
  Result := False;
  if FTokenizer[Index].TokenType <> ttFunc then
    Exit;
  FuncArgClass := 0;

  Id := GetFuncId(FTokenizer[Index].Token);
  ArgCount := 0;
  p := 0;
  Inc(Index);
  while p >= 0 do begin
    if (ArgCount <= High(ExcelFunctions[Id].ArgClass)) and (ExcelFunctions[Id].ArgClass[ArgCount] > 0) then
      FuncArgClass := ExcelFunctions[Id].ArgClass[ArgCount];
    p := GetNextFuncArg(Index);
    if p > 0 then begin
      if Index > (p - 1) then
        AddPtg(xptgMissArg97)
      else
        Expression(Index,FuncArgClass);
      Index := p + 1;
      Inc(ArgCount);
    end
    else if FTokenizer[Index].TokenType <> ttParenR then begin
      if (Index >= FTokenizer.Count) {or (FTokenizer[Index + 1].TokenType <> ttParenR)} then
        ValueError(ERR_PAR_MISS_FUNC,'');
      Expression(Index,FuncArgClass);
      Inc(Index);
      Inc(ArgCount);
    end;
  end;
  if Id >= 0 then begin
    if ExcelFunctions[Id].Min = ExcelFunctions[Id].Max then begin
      if ArgCount <> ExcelFunctions[Id].Min then
        ValueError(ERR_FUNC_WRONG_ARGS,ExcelFunctions[Id].Name);
      AddPtg(xptgFuncV97,Id,SizeOf(word));
    end
    else begin
      AddPtg(xptgFuncVarV97,ArgCount,SizeOf(byte));
      AddBuf(Id,SizeOf(word));
    end;
  end;
  FFuncArgClass := TokenClassDefault;
  Result := Id >= 0;
end;

// TODO Change this to a sorted list and binary search...
function TEncodeFormula.GetFuncId(S: AxUCString): integer;
begin
  for Result := 0 to High(ExcelFunctions) do begin
    if S = ExcelFunctions[Result].Name then begin
      if ExcelFunctions[Result].Min = $FF then
        ValueError(ERR_BAD_FUNCNAME,S);
      Exit;
    end;
  end;
  // User function
  Result := -1;
end;

procedure TEncodeFormula.AddValue(var Index: integer);
var
  Ok: boolean;
begin
  if FTokenizer[Index].TokenType = ttFunc then
    AddFunction(Index)
  else if FTokenizer[Index].TokenType in [ttRef,ttArea,ttName,ttNameX,ttPath,ttSheetName,ttSheetNameX] then begin
    Ok := AddCellRef(Index);
    if not Ok then
      Ok := AddBoolConst(Index);
    if not Ok then
      Ok := AddVariable(Index);
    if not Ok then
      ValueError(ERR_INVALIDSYMBOL,FTokenizer[Index].Token);
    if (Index < FTokenizer.Count) and (FTokenizer[Index].TokenType in [ttSheetName,ttSheetNameX]) then
      Inc(Index,2)
    else
      Inc(Index);
  end
  else begin
    if FTokenizer[Index].TokenType = ttVectL then
      AddVectorConst(Index)
    else
      AddConstant(Index);
    Inc(Index);
  end;
end;

function TEncodeFormula.AddVariable(i1: integer): boolean;
var
  NameId: integer;
  ptgName8: TPTGName8;
begin
  FUnknownNameEvent(FTokenizer[i1].Token,NameId);
  if NameId < 0 then
    ValueError(ERR_UNKNOWN_NAME,FTokenizer[i1].Token)
  else begin
    ptgName8.NameIndex := NameId + 1;
    ptgName8.Reserved := 0;
    AddPtg(ptgName_NAC or FFuncArgClass,ptgName8,SizeOf(TPTGName8));
  end;
  Result := True;
end;

procedure TEncodeFormula.AddVectorConst(i1: integer);
begin
  raise XLSRWException.Create('Vector constants not supported.');
end;

procedure TEncodeFormula.AddWS;
var
  LW: longword;
  Token: TTokenType;
  Buf: array[0..2] of byte;
begin
  if FWSStack.Count > 0 then begin
    LW := Longword(FWSStack.Pop);
    if LW <> 0 then begin
      Token := TTokenType(LW shr 16);
      LW := LW and $000000FF;
      case Token of
        ttParenL: Buf[1] := $02;
        ttParenR: Buf[1] := $04;
        else      Buf[1] := $00;
      end;
      Buf[0] := $40;
      Buf[2] := Byte(LW);
      if (FCurrBufSize + 1) > FBufSize then
        raise XLSRWException.Create('Formula buffer overflow.');
      FBuffer[FCurrBufSize] := xptgAttr97;
      FLastPtgPos := FCurrBufSize;
      Inc(FCurrBufSize);
      AddBuf(Buf,3);
    end;
  end;
end;

{ TXLSTokenizer }

procedure TXLSTokenizer.Add(Token: AxUCString; TType: TTokenType);
var
  P: PXLSToken;
begin
  New(P);
  P.Token := Token;
  P.TokenType := TType;
  inherited Add(P);
end;

procedure TXLSTokenizer.Clear;
var
  i: integer;
begin
  FError := False;
  for i := 0 to Count - 1 do begin
    Finalize(Items[i].Token);
    FreeMem(inherited Items[i]);
  end;
  inherited;
end;

function TXLSTokenizer.GetItems(Index: integer): PXLSToken;
begin
  Result := PXLSToken(inherited Items[Index])
end;

function TXLSTokenizer.DoNumber(i: integer): integer;
var
  DP: boolean;
begin
  Result := i;
  DP := False;
  while Result <= Length(FValue) do begin
    case FValue[Result] of
      '0'..'9':
        Inc(Result);
      '+','-','*','/','%','^','>','<','=',')',' ':
        Break;
      else begin
        if WideChar(FValue[Result]) = WideChar(FormatSettings.DecimalSeparator) then begin
          if DP then begin
            FTokenErrorEvent(ERR_ILLEGAL_NUMBER,'');
            FError := True;
            Break;
          end;
          DP := True;
          Inc(Result);
        end
        else if WideChar(FValue[Result]) = WideChar(FormatSettings.ListSeparator) then
          Break
        else begin
          FTokenErrorEvent(ERR_ILLEGAL_NUMBER,'');
          FError := True;
          Break;
        end;
      end;
    end;
  end;
  if not FError then
    Add(Copy(FValue,i,Result - i),ttNumber);
end;

function TXLSTokenizer.DoString(i: integer): integer;
begin
  Result := i;
  Inc(Result);
  while (Result <= Length(FValue)) and (FValue[Result] <> '"') do
    Inc(Result);
  FError := FValue[Result] <> '"';
  if not FError then begin
    Add(Copy(FValue,i + 1,Result - i - 1),ttString);
    Inc(Result);
  end
  else
    FTokenErrorEvent(ERR_STR_MISSING_Q,FOrigValue);
end;

function TXLSTokenizer.AddOp(Op: string; TType: TTokenType): integer;
begin
  Result := Length(Op);
  Add(Op,TType);
end;

function TXLSTokenizer.IsRef(i: integer): integer;
var
  j,v,pos: integer;
begin
  Result := 0;
  pos := i;
  if FValue[pos] = '$' then
    Inc(pos);
  if (pos = FLen) or not CharInSet(FValue[pos],['a'..'z','A'..'Z']) then
    Exit;
  Inc(pos);
  if CharInSet(FValue[pos - 1],['a'..'i','A'..'I']) and CharInSet(FValue[pos],['a'..'z','A'..'Z']) then
    Inc(pos);
  // maxcol = 255
  if CharInSet(FValue[pos - 1],['i','I']) and CharInSet(FValue[pos],['w'..'z','W'..'Z']) then
    Exit;
  if FValue[pos] = '$' then
    Inc(pos);
  j := pos;
  while j <= FLen do begin
    if not CharInSet(FValue[j],['0'..'9']) then
      Break;
    Inc(j);
  end;
  if j <= pos then
    Exit;
  // Check that next character is a separator, otherwise the "ref" may be the
  // first part of a name, such as "AD7SEK"
  if (j < FLen) and not CharInSet(FValue[j],[Char(FormatSettings.ListSeparator),'!',' ','=',':','0'..'9','+','&','-','*','/','%','^','>','<','(',')','{','}','"','''','[',']','#','$']) then
    Exit;
  v := StrToIntDef(Copy(FValue,pos,j - pos),-1);
  if (v < 1) or (v > (MAXROW + 1)) then
    Exit;
  Result := j - i;
end;

function TXLSTokenizer.NextToken(Index: integer): TTokenType;
begin
  while (Index < Count) and (Items[Index].TokenType = ttSpace) do
    Inc(Index);
  if (Index < Count) and (Items[Index].TokenType <> ttSpace) then
    Result := Items[Index].TokenType
  else
    Result := ttNone;
end;

function TXLSTokenizer.DoName(i: integer): integer;
begin
  Result := i;

  while (Result <= FLen) and not CharInSet(FValue[Result],[Char(FormatSettings.ListSeparator),'\','/','?','*','[',']',' ','+','&','-','*','/','%','^','<','>','=','(',')','#','{','}','''','"','!','.']) do
    Inc(Result);
{
  while (Result <= FLen) and (Char(FValue[Result]) in ['a'..'z','A'..'Z','0'..'9','_', 'Å','ä','.','ö']) do
    Inc(Result);
}
  if GetLast = ttPath then
    Add(Copy(FValue,i,Result - i),ttNameX)
  else
    Add(Copy(FValue,i,Result - i),ttName);
end;

function TXLSTokenizer.DoQuotedName(i: integer): integer;
var
  p: integer;
  S: AxUCString;
begin
  p := Pos('''',Copy(FValue,i + 1,MAXINT));
  // External path? Min length=9, 'z:\[x]n'
  if ((p - i) >= 7) and (FValue[i + 2] = ':') and (FValue[i + 3] = '\') then begin
    S := Copy(FValue,i + 1,p - i);
    p := Pos('[',S);
    if p > 1 then begin
      Result := i + p;
      Add(Copy(S,1,p - 1),ttPath);
    end
    else begin
      Result := i + Length(S) + 2;
      Add(S,ttPath);
    end;
  end
  else begin
    Result := i + 1;
    while (Result <= FLen) and (FValue[Result] <> '''') do begin
      if CharInSet(FValue[Result],['\','/','?','*','[',']']) then begin
        FTokenErrorEvent(ERR_ILLEGAL_CHAR_IN_STR,Copy(FValue,i,Result));
        FError := True;
        Exit;
      end;
      Inc(Result);
    end;
    Inc(Result);
    Add(Copy(FValue,i + 1,Result - i - 2),ttName);
  end;
end;

function TXLSTokenizer.DoFilename(i: integer): integer;
var
  p: integer;
begin
  p := Pos(']',Copy(FValue,i,MAXINT));
  if p < 3 then begin
    Result := i + 1;
    if p >= 1 then
      FTokenErrorEvent(ERR_BAD_FILENAME,Copy(FValue,i,p - i + 1))
    else
      FTokenErrorEvent(ERR_BAD_FILENAME,FValue);
    FError := True;
  end
  else begin
    Add(Copy(FValue,i + 1,p - 2),ttFile);
    Result := i + p;
  end;
end;

function TXLSTokenizer.DoError(i: integer): integer;
begin
  Result := i + 1;
  while (i <= FLen) and (CharInSet(FValue[Result],['A'..'Z','!'])) do
    Inc(Result);
end;

procedure TXLSTokenizer.Tokenize(Value: AxUCString);
var
  S: AxUCString;
  Token: TTokenType;
  i,j,n1,n2: integer;
begin
  Clear;
  FValue := Value;
  FOrigValue := Value;
  UppercaseSkipQuote(FValue);
  FError := False;
  FLen := Length(Value);
  i := 1;
  while (i <= FLen) and not FError do begin
    case FValue[i] of
      '0'..'9': i := DoNumber(i);
      '+': Inc(i,AddOp('+',ttAdd));
      '&': Inc(i,AddOp('&',ttConcat));
      '-': Inc(i,AddOp('-',ttSub));
      '*': Inc(i,AddOp('*',ttMul));
      '/': Inc(i,AddOp('/',ttDiv));
      '%': Inc(i,AddOp('%',ttPercent));
      '^': Inc(i,AddOp('%',ttPower));
      '>': begin
        if (i < FLen) and (FValue[i + 1] = '=') then
          i := i + AddOp('>=',ttGE)
        else
          i := i + AddOp('>',ttGT);
      end;
      '<': begin
        if (i < FLen) and (FValue[i + 1] = '=') then
          Inc(i,AddOp('<=',ttLE))
        else if (i < FLen) and (FValue[i + 1] = '>') then
          Inc(i,AddOp('<>',ttNE))
        else
          Inc(i,AddOp('<',ttLT));
      end;
      '=': Inc(i,AddOp('=',ttEQ));
      '(': begin
        if GetLast = ttName then begin
          ChangeLast(ttFunc);
          Inc(i);
        end
        else
          Inc(i,AddOp('(',ttParenL));
      end;
      ')': Inc(i,AddOp(')',ttParenR));
      // TODO Vectors unhandled.
      '{': FError := True;
      '}': FError := True;
      ' ': Inc(i,AddOp(' ',ttSpace));
      '"': i := DoString(i);
      '''': begin
        if GetLast = ttName then
          Inc(i)
        else
          i := DoQuotedName(i);
      end;
      '[': i := DoFilename(i);
      '!': begin
        if GetLast in [ttName,ttRef,ttArea] then begin
          if GetLast2 = ttFile then
//            Items[Count - 2].TokenType := ttPath
// Check this before changing. If above is used, autofilters (see sample) gives error.
            ChangeLast(ttSheetnameX)
          else
//            Items[Count - 2].TokenType := ttSheetname;
// Check this before changing. If above is used, autofilters (see sample) gives error.
            ChangeLast(ttSheetname);
          Inc(i);
        end
        else if GetLast = ttPath then
          Inc(i)
        else begin
          FTokenErrorEvent(ERR_INVALID_3D_NAME,'');
          FError := True;
        end;
      end;
      '#': i := DoError(i);
      '_','$',
      'a'..'z',
      'A'..'Z': begin
        n1 := IsRef(i);
        if n1 > 1 then begin
          Inc(i,n1);
          if (i < FLen) and (FValue[i] = ':') then begin
            n2 := IsRef(i + 1);
            if n2 > 1 then begin
              Inc(i,n2 + 1);
              S := Copy(FValue,i - n1 - 1 - n2,n1 + 1 + n2);
              Token := ttArea;
            end
            else begin
              S := Copy(FValue,i - n1,n1);
              Token := ttRef;
            end
          end
          else begin
            S := Copy(FValue,i - n1,n1);
            Token := ttRef;
          end;
          if GetLast = ttSheetName then begin
            if Token = ttRef then
              Token := ttRef3d
            else
              Token := ttArea3d;
          end
          else if GetLast = ttSheetNameX then begin
            if Token = ttRef then
              Token := ttRefX
            else
              Token := ttAreaX;
          end;
          Add(S,Token);
        end
        else if FValue[i] <> '$' then
          i := DoName(i)
        else begin
          FTokenErrorEvent(ERR_INVALID_NAME,'');
          FError := True;
        end;
      end;
      else begin
        if FValue[i] = FormatSettings.ListSeparator then begin
          Add(FormatSettings.ListSeparator,ttListSep);
          Inc(i);
        end
        else if FValue[i] = FormatSettings.DecimalSeparator then
          raise XLSRWException.Create('Illegal use of decimal separator')
        else begin
          j := DoName(i);
          if j = i then
            raise XLSRWException.Create('Invalid name');
          i := j;
        end;
      end;
    end;
  end;
end;

procedure TXLSTokenizer.ChangeLast(TType: TTokenType);
begin
  if Count > 0 then
    Items[Count - 1].TokenType := TType;
end;

function TXLSTokenizer.GetLast: TTokenType;
begin
  if Count > 0 then
    Result := Items[Count - 1].TokenType
  else
    Result := ttNone;
end;

function TXLSTokenizer.GetLast2: TTokenType;
begin
  if Count > 1 then
    Result := Items[Count - 2].TokenType
  else
    Result := ttNone;
end;

function TXLSTokenizer.GetName(Index: integer): AxUCString;
begin
  Result := TTokenName[Integer(GetItems(Index).TokenType)];
end;

end.
