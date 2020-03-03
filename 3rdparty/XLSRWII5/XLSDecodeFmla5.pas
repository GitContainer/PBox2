unit XLSDecodeFmla5;

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

uses Classes, SysUtils, IniFiles,
     Xc12Utils5, Xc12Manager5, Xc12DataTable5,
     XLSUtils5, XLSFormulaTypes5, XLSCellMMU5;

//§  While there are input symbol left
//    §  Read the next symbol from input.
//    §  If the symbol is an operand (i.e. value)
//        §  Push it onto the stack.
//    §  Otherwise, the symbol is an operator.
//        §  If there are fewer than 2 values on the stack
//            §  (Error) The user has not input sufficient values in the expression.
//        §  Else, Pop the top 2 values from the stack (operand1 & operand 2).
//        §  Put the operator, with the values as arguments and form a string (like : operand1 operator operand2).
//        §  Encapsulate the resulted string with parenthesis. (like: (a+b)  if operand1 =’a’, operand2 =’b’, operator = ‘+’ )
//        §  Push the resulted string back to stack.
//§  If there is only one value in the stack
//    §  That value in the stack is the desired infix string.
//§  If there are more values in the stack
//    §  (Error) The user input has too many values.

type TStringStack = class(TObject)
protected
     FStack   : array[0..255] of AxUCString;
     FStackPtr: integer;
     FManager : TXc12Manager;
public
     constructor Create(AManager: TXc12Manager);

     procedure Clear;
     function  Count: integer;

     procedure Push(AValue: AxUCString);
     procedure Op(AOperator: AxUCString);
     procedure Surround(ALeft,ARight: AxUCString);
     procedure Func(AName: AxUCString; AArgCount: integer);
     procedure Table(AName: AxUCString; AArgCount: integer);
     procedure Array_(ACols,ARows: integer);
     procedure UnaryLeft(AOperator: AxUCString);
     procedure UnaryRight(AOperator: AxUCString);
     function  Pop: AxUCString;
     function  Peek: AxUCString;
     end;

type TXLSFormulaDecoder = class(TObject)
protected
     FManager   : TXc12Manager;
     FCurrTables: TXc12Tables;
     FSheetIndex: integer;

     FStack     : TStringStack;

     function  CheckNameChars(const AName: AxUCString): AxUCString;
public
     constructor Create(AManager: TXc12Manager);
     destructor Destroy; override;

     function Decode(ACells: TXLSCellMMU; const ASheetIndex: integer; APtgs,AStopPtgs: PXLSPtgs; const APtgsSize: integer; const ASaveToFile: boolean): AxUCString;
     end;

function XLSDecodeFormula(AManager: TXc12Manager; ACells: TXLSCellMMU; const ASheetIndex: integer; APtgs: PXLSPtgs; const APtgsSize: integer): AxUCString;

implementation

function XLSDecodeFormula(AManager: TXc12Manager; ACells: TXLSCellMMU; const ASheetIndex: integer; APtgs: PXLSPtgs; const APtgsSize: integer): AxUCString;
var
  Decoder: TXLSFormulaDecoder;
begin
  Decoder := TXLSFormulaDecoder.Create(AManager);
  try
    Result := Decoder.Decode(ACells,ASheetIndex,APtgs,Nil,APtgsSize,False);
  finally
    Decoder.Free;
  end;
end;

{ TXLSFormulaDecoder }

function TXLSFormulaDecoder.CheckNameChars(const AName: AxUCString): AxUCString;
var
  i: integer;
begin
  for i := 1 to Length(AName) do begin
    if not CharInSet(AName[i],['a'..'z','A'..'Z','0'..'9']) then begin
      Result := '''' + AName + '''';
      Exit;
    end;
  end;
  Result := AName;
end;

constructor TXLSFormulaDecoder.Create(AManager: TXc12Manager);
begin
  FManager := AManager;
  FStack := TStringStack.Create(FManager);
end;

function TXLSFormulaDecoder.Decode(ACells: TXLSCellMMU; const ASheetIndex: integer; APtgs,AStopPtgs: PXLSPtgs; const APtgsSize: integer; const ASaveToFile: boolean): AxUCString;
var
  i: integer;
  Sz: integer;
  S,S2: AxUCString;
  Spaces: AxUCString;
  LastPtgs: NativeInt;
  Ref: TXLSCellItem;
  Ptgs: PXLSPtgs;
  DTF: PXLSPtgsDataTableFmla;
begin
  Result := '???';

  FSheetIndex := ASheetIndex;
  if FSheetIndex >= 0 then
    FCurrTables := FManager.Worksheets[FSheetIndex].Tables
  else
    FCurrTables := Nil;

  FStack.Clear;

  LastPtgs := NativeInt(APtgs) + APtgsSize;
  while NativeInt(APtgs) < LastPtgs do begin
    case APtgs.Id of
      xptgNone    : raise XLSRWException.CreateFmt('Illegal ptgs %.2X',[APtgs.Id]);
      xptgOpAdd   : begin
        FStack.Op(Spaces + '+');
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgs));
      end;
      xptgOpSub   : begin
        FStack.Op(Spaces + '-');
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgs));
      end;
      xptgOpMult  : begin
        FStack.Op(Spaces + '*');
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgs));
      end;
      xptgOpDiv   : begin
        FStack.Op(Spaces + '/');
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgs));
      end;
      xptgOpPower : begin
        FStack.Op(Spaces + '^');
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgs));
      end;
      xptgOpConcat: begin
        FStack.Op(Spaces + '&');
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgs));
      end;
      xptgOpLT: begin
        FStack.Op(Spaces + '<');
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgs));
      end;
      xptgOpLE: begin
        FStack.Op(Spaces + '<=');
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgs));
      end;
      xptgOpEQ: begin
        FStack.Op(Spaces + '=');
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgs));
      end;
      xptgOpGE: begin
        FStack.Op(Spaces + '>=');
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgs));
      end;
      xptgOpGT: begin
        FStack.Op(Spaces + '>');
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgs));
      end;
      xptgOpNE: begin
        FStack.Op(Spaces + '<>');
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgs));
      end;

      xptgOpIsect: begin
        SetLength(S,PXLSPtgsISect(APtgs).SpacesCount);
        for i := 1 to Length(S) do
          S[i] := ' ';
        FStack.Op(S);
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgsISect));
      end;
      xptgOpUnion: begin
        FStack.Op(Spaces + FormatSettings.ListSeparator);
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgs));
      end;
      xptgOpRange: begin
        FStack.Op(Spaces + ':');
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgs));
      end;

      xptgOpUPlus: begin
        FStack.UnaryLeft(Spaces + '+');
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgs));
      end;
      xptgOpUMinus: begin
        FStack.UnaryLeft(Spaces + '-');
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgs));
      end;

      xptgOpPercent: begin
        FStack.UnaryRight(Spaces + '%');
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgs));
      end;
      xptgLPar: begin
        FStack.Surround(Spaces + '(',')');
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgs));
      end;
      xptgArray: begin
        FStack.Array_(PXLSPtgsArray(APtgs).Cols,PXLSPtgsArray(APtgs).Rows);
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgsArray));
      end;

      xptgStr: begin
        SetLength(S,PXLSPtgsStr(APtgs).Len);
        Move(PXLSPtgsStr(APtgs).Str[0],Pointer(S)^,PXLSPtgsStr(APtgs).Len * 2);
        FStack.Push(Spaces + '"' + S + '"');
        APtgs := PXLSPtgs(NativeInt(APtgs) + FixedSzPtgsStr + PXLSPtgsStr(APtgs).Len * 2);
      end;
      xptgErr: begin
        FStack.Push(Spaces + Xc12CellErrorNames[TXc12CellError(PXLSPtgsErr(APtgs).Value)]);
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgsErr));
      end;
      xptgBool: begin
        if PXLSPtgsBool(APtgs).Value <> 0 then
          FStack.Push(Spaces + 'TRUE')
        else
          FStack.Push(Spaces + 'FALSE');
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgsBool));
      end;

      xptgInt: begin
        FStack.Push(Spaces + IntToStr(PXLSPtgsInt(APtgs).Value));
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgsInt));
      end;
      xptgNum: begin
        FStack.Push(Spaces + FloatToStr(PXLSPtgsNum(APtgs).Value));
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgsNum));
      end;

      xptgFunc: begin
        FStack.Func(Spaces + G_XLSExcelFuncNames.FindName(PXLSPtgsFunc(APtgs).FuncId),G_XLSExcelFuncNames.ArgCount(PXLSPtgsFunc(APtgs).FuncId));
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgsFunc));
      end;
      xptgFuncVar: begin
        FStack.Func(Spaces + G_XLSExcelFuncNames.FindName(PXLSPtgsFuncVar(APtgs).FuncId),PXLSPtgsFuncVar(APtgs).ArgCount);
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgsFuncVar));
      end;
      xptgName: begin
        if PXLSPtgsName(APtgs).NameId = XLS_NAME_UNKNOWN then
          FStack.Push(Spaces + Xc12CellErrorNames[errName])
        else
          FStack.Push(Spaces + FManager.Workbook.DefinedNames[PXLSPtgsName(APtgs).NameId].Name);
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgsName));
      end;
      xptgRef: begin
        FStack.Push(Spaces + ColRowToRefStrEnc(PXLSPtgsRef(APtgs).Col,PXLSPtgsRef(APtgs).Row));
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgsRef));
      end;
      xptgArea: begin
        FStack.Push(Spaces + AreaToRefStrEnc(PXLSPtgsArea(APtgs).Col1,PXLSPtgsArea(APtgs).Row1,PXLSPtgsArea(APtgs).Col2,PXLSPtgsArea(APtgs).Row2));
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgsArea));
      end;
      xptgRefErr: begin
        FStack.Push(Spaces + Xc12CellErrorNames[errRef]);
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgsRef));
      end;
      xptgAreaErr: begin
        FStack.Push(Spaces + Xc12CellErrorNames[errRef]);
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgsArea));
      end;

      xptgWS: begin
        SetLength(Spaces,PXLSPtgsWS(APtgs).Count);
        for i := 1 to Length(Spaces) do
          Spaces[i] := ' ';
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgsWS));
        Continue;
      end;

      xptgRPar: ;
      xptgRef1d: begin
        FStack.Push(Spaces + CheckNameChars(FManager.Worksheets[PXLSPtgsRef1d(APtgs).Sheet].Name) + '!' +
                    ColRowToRefStrEnc(PXLSPtgsRef1d(APtgs).Col,PXLSPtgsRef1d(APtgs).Row));
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgsRef1d));
      end;
      xptgRef1dErr: begin
        FStack.Push(Spaces + Xc12CellErrorNames[TXc12CellError(PXLSPtgsRef1dError(APtgs).Error)]);
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgsRef1dError));
      end;
      xptgRef3d: begin
        FStack.Push(Spaces + CheckNameChars(FManager.Worksheets[PXLSPtgsRef3d(APtgs).Sheet1].Name) + ':' +
                    CheckNameChars(FManager.Worksheets[PXLSPtgsRef3d(APtgs).Sheet2].Name) + '!' +
                    ColRowToRefStrEnc(PXLSPtgsRef3d(APtgs).Col,PXLSPtgsRef3d(APtgs).Row));
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgsRef3d));
      end;
      xptgRef3dErr: begin
        FStack.Push(Spaces + Xc12CellErrorNames[TXc12CellError(PXLSPtgsRef3dError(APtgs).Error1)]);
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgsRef3dError));
      end;
      xptgXRef1d: begin
        if ASaveToFile then
          FStack.Push(Spaces + '[' + IntToStr(PXLSPtgsXRef1d(APtgs).XBook + 1) + ']' +
                    FManager.XLinks[PXLSPtgsXRef1d(APtgs).XBook].ExternalBook.SheetNames[PXLSPtgsXRef1d(APtgs).Sheet] + '!' +
                    ColRowToRefStrEnc(PXLSPtgsXRef1d(APtgs).Col and not COL_ABSFLAG,PXLSPtgsXRef1d(APtgs).Row and not ROW_ABSFLAG))
        else
          FStack.Push(Spaces + '[' + FManager.XLinks[PXLSPtgsXRef1d(APtgs).XBook].ExternalBook.Filename + ']' +
                    FManager.XLinks[PXLSPtgsXRef1d(APtgs).XBook].ExternalBook.SheetNames[PXLSPtgsXRef1d(APtgs).Sheet] + '!' +
                    ColRowToRefStrEnc(PXLSPtgsXRef1d(APtgs).Col and not COL_ABSFLAG,PXLSPtgsXRef1d(APtgs).Row and not ROW_ABSFLAG));
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgsXRef1d));
      end;
      xptgXRef3d: begin
        if ASaveToFile then
                    FStack.Push(Spaces + '[' + IntToStr(PXLSPtgsXRef3d(APtgs).XBook + 1) + ']' +
                    FManager.XLinks[PXLSPtgsXRef3d(APtgs).XBook].ExternalBook.SheetNames[PXLSPtgsXRef3d(APtgs).Sheet1] + ':' +
                    FManager.XLinks[PXLSPtgsXRef3d(APtgs).XBook].ExternalBook.SheetNames[PXLSPtgsXRef3d(APtgs).Sheet2] + '!' +
                    ColRowToRefStrEnc(PXLSPtgsXRef3d(APtgs).Col,PXLSPtgsXRef3d(APtgs).Row))
        else
          FStack.Push(Spaces + '[' + FManager.XLinks[PXLSPtgsXRef3d(APtgs).XBook].ExternalBook.FileName + ']' +
                    FManager.XLinks[PXLSPtgsXRef3d(APtgs).XBook].ExternalBook.SheetNames[PXLSPtgsXRef3d(APtgs).Sheet1] + ':' +
                    FManager.XLinks[PXLSPtgsXRef3d(APtgs).XBook].ExternalBook.SheetNames[PXLSPtgsXRef3d(APtgs).Sheet2] + '!' +
                    ColRowToRefStrEnc(PXLSPtgsXRef3d(APtgs).Col,PXLSPtgsXRef3d(APtgs).Row));
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgsXRef3d));
      end;
      xptgArea1d: begin
        FStack.Push(Spaces + CheckNameChars(FManager.Worksheets[PXLSPtgsArea1d(APtgs).Sheet].Name) + '!' +
                    AreaToRefStrEnc(PXLSPtgsArea1d(APtgs).Col1,PXLSPtgsArea1d(APtgs).Row1,PXLSPtgsArea1d(APtgs).Col2,PXLSPtgsArea1d(APtgs).Row2));
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgsArea1d));
      end;
      xptgArea1dErr: begin
        FStack.Push(Spaces + CheckNameChars(FManager.Worksheets[PXLSPtgsArea1d(APtgs).Sheet].Name) + '!' + Xc12CellErrorNames[errRef]);

//        FStack.Push(Spaces + Xc12CellErrorNames[TXc12CellError(PXLSPtgsArea1dError(APtgs).Error)]);
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgsArea1dError));
      end;
      xptgArea3d: begin
        FStack.Push(Spaces + CheckNameChars(FManager.Worksheets[PXLSPtgsArea3d(APtgs).Sheet1].Name) + ':' +
                    CheckNameChars(FManager.Worksheets[PXLSPtgsArea3d(APtgs).Sheet2].Name) + '!' +
                    AreaToRefStrEnc(PXLSPtgsArea3d(APtgs).Col1,PXLSPtgsArea3d(APtgs).Row1,PXLSPtgsArea3d(APtgs).Col2,PXLSPtgsArea3d(APtgs).Row2));
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgsArea3d));
      end;
      xptgArea3dErr: begin
        FStack.Push(Spaces + Xc12CellErrorNames[TXc12CellError(PXLSPtgsArea3dError(APtgs).Error1)]);
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgsArea3dError));
      end;
      xptgXArea1d: begin
        if ASaveToFile then
          FStack.Push(Spaces + '[' + IntToStr(PXLSPtgsXArea1d(APtgs).XBook + 1) + ']' +
                    FManager.XLinks[PXLSPtgsXArea1d(APtgs).XBook].ExternalBook.SheetNames[PXLSPtgsXArea1d(APtgs).Sheet] + '!' +
                    AreaToRefStrEnc(PXLSPtgsXArea1d(APtgs).Col1,PXLSPtgsXArea1d(APtgs).Row1,PXLSPtgsXArea1d(APtgs).Col2,PXLSPtgsXArea1d(APtgs).Row2))
        else
          FStack.Push(Spaces + '[' + FManager.XLinks[PXLSPtgsXArea1d(APtgs).XBook].ExternalBook.Filename + ']' +
                    FManager.XLinks[PXLSPtgsXArea1d(APtgs).XBook].ExternalBook.SheetNames[PXLSPtgsXArea1d(APtgs).Sheet] + '!' +
                    AreaToRefStrEnc(PXLSPtgsXArea1d(APtgs).Col1,PXLSPtgsXArea1d(APtgs).Row1,PXLSPtgsXArea1d(APtgs).Col2,PXLSPtgsXArea1d(APtgs).Row2));
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgsXArea1d));
      end;
      xptgXArea3d: begin
        if ASaveToFile then
          FStack.Push(Spaces + '[' + IntToStr(PXLSPtgsXArea3d(APtgs).XBook + 1) + ']' +
                    FManager.XLinks[PXLSPtgsXArea3d(APtgs).XBook].ExternalBook.SheetNames[PXLSPtgsXArea3d(APtgs).Sheet1] + ':' +
                    FManager.XLinks[PXLSPtgsXArea3d(APtgs).XBook].ExternalBook.SheetNames[PXLSPtgsXArea3d(APtgs).Sheet2] + '!' +
                    AreaToRefStrEnc(PXLSPtgsXArea3d(APtgs).Col1,PXLSPtgsXArea3d(APtgs).Row1,PXLSPtgsXArea3d(APtgs).Col2,PXLSPtgsXArea3d(APtgs).Row2))
        else
          FStack.Push(Spaces + '[' + FManager.XLinks[PXLSPtgsXArea3d(APtgs).XBook].ExternalBook.FileName + ']' +
                    FManager.XLinks[PXLSPtgsXArea3d(APtgs).XBook].ExternalBook.SheetNames[PXLSPtgsXArea3d(APtgs).Sheet1] + ':' +
                    FManager.XLinks[PXLSPtgsXArea3d(APtgs).XBook].ExternalBook.SheetNames[PXLSPtgsXArea3d(APtgs).Sheet2] + '!' +
                    AreaToRefStrEnc(PXLSPtgsXArea3d(APtgs).Col1,PXLSPtgsXArea3d(APtgs).Row1,PXLSPtgsXArea3d(APtgs).Col2,PXLSPtgsXArea3d(APtgs).Row2));
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgsXArea3d));
      end;

      xptgArrayFmlaChild: begin
        // ACells = Nil when decoding a Name formula, and names can't have child
        // array formulas.
        if ACells = Nil then
          raise XLSRWException.Create('Decode formula with ACells = Nil');
        Ref := ACells.FindCell(PXLSPtgsArrayChildFmla(APtgs).ParentCol,PXLSPtgsArrayChildFmla(APtgs).ParentRow);
        if ACells.IsFormulaCompiled(@Ref) then begin
          Sz := ACells.GetFormulaPtgs(@Ref,Ptgs);
          FStack.Push(Decode(ACells,FSheetIndex,Ptgs,AStopPtgs,Sz,ASaveToFile));
        end
        else
          FStack.Push(ACells.GetFormula(@Ref));
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgsArrayChildFmla));
      end;

      xptgDataTableFmla: begin
        DTF := PXLSPtgsDataTableFmla(APtgs);
        if (Xc12FormulaTableOpt_DT2D and DTF.Options) <> 0 then begin
          if (Xc12FormulaTableOpt_del1 and DTF.Options) <> 0 then
            S := Xc12CellErrorNames[errRef]
          else
            S := ColRowToRefStrEnc(DTF.R1Col,DTF.R1Row);
          if (Xc12FormulaTableOpt_del2 and DTF.Options) <> 0 then
            S2 := Xc12CellErrorNames[errRef]
          else
            S2 := ColRowToRefStrEnc(DTF.R2Col,DTF.R2Row);
        end
        else if (Xc12FormulaTableOpt_DTR and DTF.Options) <> 0 then begin
          S2 := '';
          if (Xc12FormulaTableOpt_del1 and DTF.Options) <> 0 then
            S := Xc12CellErrorNames[errRef]
          else
            S := ColRowToRefStrEnc(DTF.R1Col,DTF.R1Row);
        end
        else begin
          S := '';
          if (Xc12FormulaTableOpt_del1 and DTF.Options) <> 0 then
            S2 := Xc12CellErrorNames[errRef]
          else
            S2 := ColRowToRefStrEnc(DTF.R1Col,DTF.R1Row);
        end;

        FStack.Push(Spaces + 'TABLE(' + S + FormatSettings.ListSeparator + S2 + ')');

        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgsDataTableFmla));
      end;

      xptgDataTableFmlaChild: begin
        // ACells = Nil when decoding a Name formula, and names can't have child
        // data table formulas.
        if ACells = Nil then
          raise XLSRWException.Create('Decode formula with ACells = Nil');
        Ref := ACells.FindCell(PXLSPtgsDataTableChildFmla(APtgs).ParentCol,PXLSPtgsDataTableChildFmla(APtgs).ParentRow);
        if ACells.IsFormulaCompiled(@Ref) then begin
          Sz := ACells.GetFormulaPtgs(@Ref,Ptgs);
          FStack.Push(Decode(ACells,FSheetIndex,Ptgs,AStopPtgs,Sz,ASaveToFile));
        end
        else
          FStack.Push(ACells.GetFormula(@Ref));
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgsDataTableChildFmla));
      end;

      xptgTable: begin
        if PXLSPtgsTable(APtgs).Sheet < XLS_MAXSHEETS then
          FCurrTables := FManager.Worksheets[PXLSPtgsTable(APtgs).Sheet].Tables;

        FStack.Table(Spaces + FCurrTables[PXLSPtgsTable(APtgs).TableId].Name,PXLSPtgsTable(APtgs).ArgCount);
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgsTable));
      end;
      xptgTableCol: begin
        FStack.Push(Spaces + FCurrTables[PXLSPtgsTableCol(APtgs).TableId].TableColumns[PXLSPtgsTableCol(APtgs).ColId].Name);
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgsTableCol));
      end;
      xptgTableSpecial: begin
        case TXLSTableSpecialSpecifier(PXLSPtgsTableSpecial(APtgs).SpecialId) of
          xtssNone   : S := '_ERROR_';
          xtssAll    : S := '#All';
          xtssData   : S := '#Data';
          xtssHeaders: S := '#Headers';
          xtssTotals : S := '#Totals';
          xtssThisRow: S := '#This Row';
        end;
        FStack.Push(S);
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgsTableSpecial));
      end;

      xptgMissingArg: begin
        FStack.Push(Spaces);
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgs));
      end;
      xptgUserFunc: begin
        SetLength(S,PXLSPtgsUserFunc(APtgs).Len);
        Move(PXLSPtgsUserFunc(APtgs).Name[0],Pointer(S)^,PXLSPtgsUserFunc(APtgs).Len * 2);
        FStack.Func(Spaces + S,PXLSPtgsUserFunc(APtgs).ArgCount);
        APtgs := PXLSPtgs(NativeInt(APtgs) + FixedSzPtgsUserFunc + PXLSPtgsUserFunc(APtgs).Len * 2);
      end;
      else raise XLSRWException.CreateFmt('Unknown Ptg "%.2X"',[APtgs.Id]);
    end;
    Spaces := '';
    if (AStopPtgs <> Nil) and (APtgs = AStopPtgs) then
      Break;
  end;
  Result := FStack.Peek;
end;

destructor TXLSFormulaDecoder.Destroy;
begin
  FStack.Free;
  inherited;
end;

{ TStringStack }

procedure TStringStack.Array_(ACols,ARows: integer);
var
  i,C,R: integer;
  S: AxUCString;
begin
  i := ACols * ARows;
  if i > 0 then begin
    Dec(i);
    S := '{';
    for R := 0 to ARows - 1 do begin
      if R <> 0 then
        S := S + FormatSettings.ListSeparator ;
      for C := 0 to ACols - 1 do begin
        if C = 0 then
          S := S + FStack[FStackPtr - i]
        else
          S := S + FManager.ColSeparator + FStack[FStackPtr - i];
        Dec(i);
      end;
    end;
    Dec(FStackPtr,ACols * ARows - 1);
    FStack[FStackPtr] := S + '}';
  end
  else
    Push('{}');
end;

procedure TStringStack.Clear;
begin
  FStackPtr := -1;
end;

function TStringStack.Count: integer;
begin
  Result := FStackPtr + 1;
end;

constructor TStringStack.Create(AManager: TXc12Manager);
begin
  FManager := AManager;
  FStackPtr := -1;
end;

procedure TStringStack.Func(AName: AxUCString; AArgCount: integer);
var
  i: integer;
  S: AxUCString;
begin
  if AArgCount > 0 then begin
    S := AName + '(' + FStack[FStackPtr - AArgCount + 1];
    for i := AArgCount - 1 downto 1 do
      S := S + FormatSettings.ListSeparator + FStack[FStackPtr - i + 1];
    Dec(FStackPtr,AArgCount - 1);
    FStack[FStackPtr] := S + ')';
  end
  else
    Push(AName + '()');
end;

procedure TStringStack.Op(AOperator: AxUCString);
begin
  Dec(FStackPtr);
  if FStackPtr < 0 then
    raise XLSRWException.Create('Empty decode stack');
  FStack[FStackPtr] := FStack[FStackPtr] + AOperator + FStack[FStackPtr + 1];
end;

function TStringStack.Peek: AxUCString;
begin
  Result := FStack[FStackPtr];
end;

function TStringStack.Pop: AxUCString;
begin
  if FStackPtr < 0 then
    raise XLSRWException.Create('Empty decode stack');
  Result := FStack[FStackPtr];
  Dec(FStackPtr);
end;

procedure TStringStack.Push(AValue: AxUCString);
begin
  Inc(FStackPtr);
  if FStackPtr > High(FStack) then
    raise XLSRWException.Create('Out of decode stack');
  FStack[FStackPtr] := AValue;
end;

procedure TStringStack.Surround(ALeft, ARight: AxUCString);
begin
  FStack[FStackPtr] := ALeft + FStack[FStackPtr] + ARight;
end;

procedure TStringStack.Table(AName: AxUCString; AArgCount: integer);
var
  i: integer;
  S: AxUCString;
begin
  if AArgCount > 0 then begin
    S := AName + '[' + FStack[FStackPtr - AArgCount + 1];
    for i := AArgCount - 1 downto 1 do
      S := S + FormatSettings.ListSeparator + FStack[FStackPtr - i + 1];
    Dec(FStackPtr,AArgCount - 1);
    FStack[FStackPtr] := S + ']';
  end
  else
    Push(AName + '[]');
end;

procedure TStringStack.UnaryLeft(AOperator: AxUCString);
begin
  FStack[FStackPtr] := AOperator + FStack[FStackPtr];
end;

procedure TStringStack.UnaryRight(AOperator: AxUCString);
begin
  FStack[FStackPtr] := FStack[FStackPtr] + AOperator;
end;

end.
