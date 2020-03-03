unit BIFF_DecodeFormula5;

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
     BIFF_Utils5, BIFF_RecsII5, BIFF_ExcelFuncII5,
     Xc12Utils5,
     XLSUtils5, XLSFormulaTypes5;

type TGetNameEvent = function(NameType: TFormulaNameType; SheetIndex,NameIndex,Col,Row: integer): AxUCString of object;

function  DecodeFmla(Buf: Pointer; Len: integer; SheetIndex,ACol,ARow: integer; GetNameMethod: TGetNameEvent; FuncArgSep: WideChar): AxUCString;
procedure ConvertShrFmla(BIFF8: boolean; Buf: Pointer; Len,ACol,ARow: integer);
procedure AdjustSheet(Buf: PByteArray; Len,NewExtIndex: integer);
procedure AdjustCell(BIFF8: boolean; Buf: Pointer; Len,DCol,DRow: integer; LockStartRow,ForceAdjust: Boolean);
procedure AdjustCell2(Buf: Pointer; Len,Col,Row,SheetIndex,DCol,DRow: integer);

type TAbsoluteRef = (arCol1,arRow1,arCol2,arRow2);
type TAbsoluteRefs = set of TAbsoluteRef;

var
  StrTRUE: AxUCString;
  StrFALSE: AxUCString;
  FuncArgSeparator: WideChar;

implementation

// There is a warning that is impossible to get rid of here...
{$WARNINGS OFF}

function IsAbsColEx12(C: word): boolean; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
begin
  Result := (C and $8000) = $8000;
end;

function IsAbsRowEx12(C: word): boolean; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
begin
  Result := (C and $80000000) = $80000000;
end;

function DecodeFmla(Buf: Pointer; Len: integer; SheetIndex,ACol,ARow: integer; GetNameMethod: TGetNameEvent; FuncArgSep: WideChar): AxUCString;
var
  i,C,R: integer;
  P,pArray: Pointer;
  B: byte;
  S: AxUCString;
  Stack: TStringList;

procedure Operator(O: AxUCString);
begin
  if Stack.Count < 2 then
    Stack.Add('<Val missing>')
  else begin
    Stack[Stack.Count - 2] := Stack[Stack.Count - 2] + O + Stack[Stack.Count - 1];
    Stack.Delete(Stack.Count - 1);
  end;
  P := Pointer(NativeInt(P) + 1);
end;

procedure UnaryOperator(O: AxUCString);
begin
  if Stack.Count < 1 then
    Stack.Add('<Val missing>')
  else
    Stack[Stack.Count - 1] := O + Stack[Stack.Count - 1];
  P := Pointer(NativeInt(P) + 1);
end;

function GetArray: AxUCString;
var
  i,j: integer;
begin
  Result := '';
  C := PPTGArray97(pArray).Cols;
  R := PPTGArray97(pArray).Rows;
  pArray := Pointer(NativeInt(pArray) + 3);
  for i := 0 to C do begin
    Result := Result + '{';
    for j := 0 to R do begin
      case Byte(pArray^) of
        $00: pArray := Pointer(NativeInt(pArray) + 9);
        $01: begin
          Result := Result + FloatToStr(TArrayFloat97(pArray^).Value) + FuncArgSep;
          pArray := Pointer(NativeInt(pArray) + 9);
        end;
        $02: begin
          Result := Result + '"' + ByteStrToWideString(@TArrayString97(pArray^).Data,TArrayString97(pArray^).Len) + '"' + FuncArgSep;
          pArray := Pointer(NativeInt(pArray) + TArrayString97(pArray^).Len + 4);
        end;
        $04: begin
          pArray := Pointer(NativeInt(pArray) + 1);
          if Byte(pArray^) = 1 then
            Result := Result + StrTRUE + FuncArgSep
          else
            Result := Result + StrFALSE + FuncArgSep;
          pArray := Pointer(NativeInt(pArray) + 8);
        end;
        $10: begin
          pArray := Pointer(NativeInt(pArray) + 1);
          Result := Result + ErrorCodeToText(Byte(pArray^)) + FuncArgSep;
          pArray := Pointer(NativeInt(pArray) + 8);
        end;
        else
          Result := 'Bad element ID#' + IntToStr(TArrayFloat97(pArray^).ID) + ' in array';
      end;
    end;
    Result := Copy(Result,1,Length(Result) - 1) + '}';
  end;
end;

procedure DecodeArea7(Cin: byte; Rin: word; var Cout,Rout: integer);
begin
  if (Rin and $8000) = 0 then
    Rout := Rin and $FF
  else
    Rout := ARow + (Rin and $FF);
  if (Rin and $4000) = 0 then
    Cout := Cin
  else
    Cout := ACol + Cin;
end;

procedure DecodeArea8(Cin,Rin: integer; var Cout,Rout: integer);
begin
  if (Cin and $4000) = 0 then
    Cout := Cin and $FF
  else
    Cout := ACol + Shortint(Cin and $FF);
  if (Cin and $8000) = 0 then
    Rout := Rin
  else
    Rout := ARow + Rin;
end;

function GetFuncArgs(Count: integer): AxUCString;
var
  i: integer;
begin
  Result := '';
  for i := Count downto 1 do begin
    if (Stack.Count - i) < 0 then
      raise XLSRWException.Create('Function arguments missing');
    Result := Result + Stack[Stack.Count - i] + FuncArgSep;
    Stack.Delete(Stack.Count - i);
  end;
  if Count > 0 then
    Result := Copy(Result,1,Length(Result) - 1);
end;


begin
  // The compiler waring "Return value of function 'DecodeFmla might be undefined"
  // is impossible to get rid of.
  // NOTE! Do not try to fix the warning by splitting up this function. It may
  // result in compiler bugs, i.e. some code may not be executed!
  Result := '';
  Stack := TStringList.Create;
  P := Buf;
  while (NativeInt(P) - NativeInt(Buf)) < Len do begin
    case GetBasePtgs(Byte(P^)) of
      0: Break;
      xptg_EXCEL_97:         P := Pointer(NativeInt(P) + 1);
      xptg_ARRAYCONSTS_97:   begin
        P := Pointer(NativeInt(P) + 1);
        Dec(Len,PWord(P)^);
        P := Pointer(NativeInt(P) + 2);
        pArray := Pointer(NativeInt(P) + Len - 3);
      end;
      xptgExtend97: begin
        P := Pointer(NativeInt(P) + 1);
        case Byte(P^) of
          eptgElfLel:        P := Pointer(NativeInt(P) + 4);
          eptgElfRw:         P := Pointer(NativeInt(P) + 4);
          eptgElfCol:        P := Pointer(NativeInt(P) + 4);
          eptgElfRwV: begin
            P := Pointer(NativeInt(P) + 1);
            Stack.Add(GetNameMethod(ntCellValue,SheetIndex,-1,PByteArray(P)[2],PWordArray(P)[0]));
            P := Pointer(NativeInt(P) + 3);
          end;
          eptgElfColV: begin
            P := Pointer(NativeInt(P) + 1);
            Stack.Add(GetNameMethod(ntCellValue,SheetIndex,-1,PByteArray(P)[2],PWordArray(P)[0]));
            P := Pointer(NativeInt(P) + 3);
          end;
          eptgElfRadical:    P := Pointer(NativeInt(P) + 13);
          eptgElfRadicalS:   P := Pointer(NativeInt(P) + 13);
          eptgElfRwS:        P := Pointer(NativeInt(P) + 4);
          eptgElfColS:       P := Pointer(NativeInt(P) + 4);
          eptgElfRwSV:       P := Pointer(NativeInt(P) + 4);
          eptgElfColSV:      P := Pointer(NativeInt(P) + 4);
          eptgElfRadicalLel: P := Pointer(NativeInt(P) + 4);
          eptgSxName:        P := Pointer(NativeInt(P) + 4);
          else
            Stack.Add(Format('Unknown eptg[%.2X]',[Byte(P^)]));
        end;
        P := Pointer(NativeInt(P) + 1);
      end;
//      xptgExp:
//      begin
//        P := Pointer(NativeInt(P) + 1);
//        Stack.Add(Format('[Row=%d Col=%d]',[PWordArray(P)[0],PWordArray(P)[1]]));
//        P := Pointer(NativeInt(P) + 4);
//      end;
      xptgOpAdd: Operator('+');
      xptgOpSub: Operator('-');
      xptgOpMult: Operator('*');
      xptgOpDiv: Operator('/');
      xptgOpPower: Operator('^');
      xptgOpConcat: Operator('&');
      xptgOpLT: Operator('<');
      xptgOpLE: Operator('<=');
      xptgOpEQ: Operator('=');
      xptgOpGE: Operator('>=');
      xptgOpGT: Operator('>');
      xptgOpNE: Operator('<>');
      xptgOpUnion: Operator(FormatSettings.ListSeparator);
      xptgOpRange: Operator(':');
      xptgOpUPlus: UnaryOperator('+'); //P := Pointer(NativeInt(P) + 1);
      xptgOpUMinus: UnaryOperator('-');
      xptgOpPercent: begin
        if Stack.Count < 1 then
          Stack.Add('<Val missing>')
        else
          Stack[Stack.Count - 1] := Stack[Stack.Count - 1] + '%';
        P := Pointer(NativeInt(P) + 1);
      end;
      xptgLPar:
      begin
        Stack[Stack.Count - 1] := '(' + Stack[Stack.Count - 1] + ')';
        P := Pointer(NativeInt(P) + 1);
      end;
      xptgMissArg97:
      begin
        P := Pointer(NativeInt(P) + 1);
        Stack.Add('');
      end;
      xptgStr97:
      begin
        P := Pointer(NativeInt(P) + 1);
        if PByteArray(P)[1] = 0 then begin
          B := Byte(P^);
          P := Pointer(NativeInt(P) + 1);
          S := ByteStrToWideString(P,B);
        end
        else begin
          B := Byte(P^) * 2;
          P := Pointer(NativeInt(P) + 1);
          S := ByteStrToWideString(P,B div 2);
        end;
        P := Pointer(NativeInt(P) + B);
        Stack.Add('"' + S + '"');
        P := Pointer(NativeInt(P) + 1);
      end;
      xptgAttr97:
      begin
        P := Pointer(NativeInt(P) + 1);
        if (Byte(P^) and $04) = $04 then begin
          P := Pointer(NativeInt(P) + 1);
          P := Pointer(NativeInt(P) + (Word(P^) + 2) * SizeOf(word) - 3);
        end
        else if (Byte(P^) and $10) = $10 then
          Stack[Stack.Count - 1] := 'SUM(' + Stack[Stack.Count - 1] + ')';
        P := Pointer(NativeInt(P) + 3);
      end;
      xptgErr:
      begin
        P := Pointer(NativeInt(P) + 1);
        Stack.Add(ErrorCodeToText(Byte(P^)));
        P := Pointer(NativeInt(P) + 1);
      end;
      xptgBool:
      begin
        P := Pointer(NativeInt(P) + 1);
        if Byte(P^) = 0 then
          Stack.Add(StrFALSE)
        else
          Stack.Add(StrTRUE);
        P := Pointer(NativeInt(P) + 1);
      end;
      xptgInt97:
      begin
        P := Pointer(NativeInt(P) + 1);
        Stack.Add(IntToStr(Word(P^)));
        P := Pointer(NativeInt(P) + 2);
      end;
      xptgNum:
      begin
        P := Pointer(NativeInt(P) + 1);
        Stack.Add(FloatToStr(Double(P^)));
        P := Pointer(NativeInt(P) + 8);
      end;
      xptgRef97:
      begin
        P := Pointer(NativeInt(P) + 1);
        with PPTGRef8(P)^ do begin
          Stack.Add(ColRowToRefStr(Col and $3FFF,Row,(Col and $4000) = 0,(Col and $8000) = 0));
          P := Pointer(NativeInt(P) + SizeOf(TPTGRef8));
        end;
      end;
      xptgRefN97:
      begin
        P := Pointer(NativeInt(P) + 1);
        with PPTGRef8(P)^ do begin
          DecodeArea8(Col,Row,C,R);
          Stack.Add(ColRowToRefStr(C,R,(Col and $4000) = 0,(Col and $8000) = 0));
          P := Pointer(NativeInt(P) + SizeOf(TPTGRef8));
        end;
      end;
      xptgArea97:
      begin
        P := Pointer(NativeInt(P) + 1);
        with PPTGArea8(P)^ do begin
          Stack.Add(ColRowToRefStr(Col1 and $3FFF,Row1,(Col1 and $4000) = 0,(Col1 and $8000) = 0) + ':' +
                    ColRowToRefStr(Col2 and $3FFF,Row2,(Col2 and $4000) = 0,(Col2 and $8000) = 0));
          P := Pointer(NativeInt(P) + SizeOf(TPTGArea8));
        end;
      end;
      xptgAreaN97:
      begin
        P := Pointer(NativeInt(P) + 1);
        with PPTGArea8(P)^ do begin
          DecodeArea8(Col1,Row1,C,R);
          S := ColRowToRefStr(C,R,(Col1 and $4000) = 0,(Col1 and $8000) = 0);
          DecodeArea8(Col2,Row2,C,R);
          Stack.Add(S + ':' + ColRowToRefStr(C,R,(Col2 and $4000) = 0,(Col2 and $8000) = 0));
          P := Pointer(NativeInt(P) + SizeOf(TPTGArea8));
        end;
      end;
      xptgRefErr97:
      begin
        P := Pointer(NativeInt(P) + 5);
        Stack.Add('#REF!');
      end;
      xptgAreaErr97:
      begin
        P := Pointer(NativeInt(P) + 9);
        Stack.Add('#REF!');
      end;
      xptgName97:
      begin
        P := Pointer(NativeInt(P) + 1);
        with PPTGName8(P)^ do begin
          if Assigned(GetNameMethod) then
            Stack.Add(GetNameMethod(ntName,-1,NameIndex,0,0))
          else
            Stack.Add('[EXTERNNAME]');
          P := Pointer(NativeInt(P) + SizeOf(TPTGName8));
        end;
      end;
      xptgNameX97:
      begin
        P := Pointer(NativeInt(P) + 1);
        with PPTGNameX8(P)^ do begin
          if Assigned(GetNameMethod) then
            S := GetNameMethod(ntExternSheet,ExtSheet,NameIndex,0,0)
          else
            S := '[EXTERNNAME X]';
          Stack.Add(S);
          P := Pointer(NativeInt(P) + SizeOf(TPTGNameX8));
        end;
      end;
{
      ptgNameV:
      begin
        P := Pointer(NativeInt(P) + 1);
        with PPTGName(P)^ do
          Stack.Add(FNames[ExtIndex - 1]);
        P := Pointer(NativeInt(P) + SizeOf(TPTGName));
      end;
}
      xptgArray97:
      begin
        P := Pointer(NativeInt(P) + 1);
        Stack.Add(GetArray);
        P := Pointer(NativeInt(P) + 7);
      end;
      xptgFunc97:
      begin
        P := Pointer(NativeInt(P) + 1);
        Stack.Add(ExcelFunctions[word(P^)].Name + '(' + GetFuncArgs(ExcelFunctions[word(P^)].Min) + ')');
        P := Pointer(NativeInt(P) + 2);
      end;
      xptgFuncVar97:
      begin
        P := Pointer(NativeInt(P) + 1);
        i := Byte(P^) and $7F;
        P := Pointer(NativeInt(P) + 1);
        S := '1';
        S := ExcelFunctions[word(P^) and $7FFF].Name;
        if (S[1] >= '0') and (S[1] <= '9') then begin
          S := Stack[Stack.Count - i];
          Stack.Delete(Stack.Count - i);
          Dec(i);
        end;
        Stack.Add(S + '(' + GetFuncArgs(i) + ')');
        P := Pointer(NativeInt(P) + 2);
      end;
      xptgRef3d97,xptgRefErr3d97:
      begin
        i := Byte(P^);
        P := Pointer(NativeInt(P) + 1);
        with PPTGRef3d8(P)^ do begin
          if Assigned(GetNameMethod) then
            S := GetNameMethod(ntExternSheet,Index,-1,0,0)
          else
            S := '[ExternSheet]';
          if i = xptgRefErr3d97 then
            Stack.Add(S + '!#REF!')
          else
            Stack.Add(S + ColRowToRefStr(Col and $3FFF,Row,(Col and $4000) = 0,(Col and $8000) = 0));
          P := Pointer(NativeInt(P) + SizeOf(TPTGRef3d8));
        end;
      end;
      xptgArea3d97,xptgAreaErr3d97:
      begin
        i := Byte(P^);
        P := Pointer(NativeInt(P) + 1);
        // Compiler bugg? add P,1 don't works.
//        P := Pointer(NativeInt(P) + 1);
        with PPTGArea3d8(P)^ do begin
          if Assigned(GetNameMethod) then
            S := GetNameMethod(ntExternSheet,Index,-1,0,0)
          else
            S := '[ExternSheet]';
          if i = xptgAreaErr3d97 then
            Stack.Add(S + '!#REF!')
          else
            Stack.Add(S + ColRowToRefStr(Col1 and $3FFF,Row1,(Col1 and $4000) = 0,(Col1 and $8000) = 0) + ':' +
                          ColRowToRefStr(Col2 and $3FFF,Row2,(Col2 and $4000) = 0,(Col2 and $8000) = 0));
          P := Pointer(NativeInt(P) + SizeOf(TPTGArea3d8));
        end;
      end;
      xptgMemErr97,
      xptgMemFunc97,
      xptgMemArea97: begin
        P := Pointer(NativeInt(P) + 1);
        P := Pointer(NativeInt(P) + PWord(P)^);
      end;
      else begin
        Stack.Add(Format('Unknown ptg[%.2X]',[Byte(P^)]));
        Break;
      end;
    end;
  end;
  if Stack.Count < 1 then
    Result := '[ERROR]'
  else
    Result := Stack[Stack.Count - 1];
  Stack.Free;
end;
{$WARNINGS ON}

procedure AdjustCell(BIFF8: boolean; Buf: Pointer; Len,DCol,DRow: integer; LockStartRow,ForceAdjust : Boolean);
var
  V,C1,R1,C2,R2: integer;
  Ok: boolean;
  IsEx12: boolean;
  P,P2: PByteArray;
begin
  P := Buf;
  IsEx12 := False;
  while (NativeInt(P) - NativeInt(Buf)) < Len do begin
    case GetBasePtgs(P[0]) of
      $02 {ptgTbl}   {02} : V := 5;
      xptgOpAdd      {03} : V := 1;
      xptgOpSub      {04} : V := 1;
      xptgOpMult     {05} : V := 1;
      xptgOpDiv      {06} : V := 1;
      xptgOpPower    {07} : V := 1;
      xptgOpConcat   {08} : V := 1;
      xptgOpLT       {09} : V := 1;
      xptgOpLE       {0A} : V := 1;
      xptgOpEQ       {0B} : V := 1;
      xptgOpGE       {0C} : V := 1;
      xptgOpGT       {0D} : V := 1;
      xptgOpNE       {0E} : V := 1;
      xptgOpIsect    {0F} : V := 1;
      xptgOpUnion    {10} : V := 1;
      xptgOpRange    {11} : V := 1;
      xptgOpUplus    {12} : V := 1;
      xptgOpUminus   {13} : V := 1;
      xptgOpPercent  {14} : V := 1;
      xptgLPar       {15} : V := 1;
      xptgMissArg97  {16} : V := 1;
      xptgStr97      {17} : begin
        P := Pointer(NativeInt(P) + 1);
        if BIFF8 then begin
          if PByteArray(P)[1] = 0 then
            V := PByteArray(P)[0] + 2
          else
            V := (PByteArray(P)[0] * 2) + 2;
        end
        else
          V := PByteArray(P)[0] + 1;
      end;
      xptgAttr97    {19} : begin
        P := Pointer(NativeInt(P) + 1);

        if P[0] = $04 then begin
          P := Pointer(NativeInt(P) + 1);
          V := (Word(Pointer(P)^) + 1) * SizeOf(word) + 2;
        end
        else
          V := 3;
      end;
      xptgErr        {1C} : V := 2;
      xptgBool       {1D} : V := 2;
      xptgInt97      {1E} : V := 3;
      xptgNum        {1F} : V := 9;

      xptgArray97    {60} : V := 8;

      xptgFunc97     {61} : V := 3;

      xptgFuncVar97  {62} : V := 4;

      xptgName97     {63} : if BIFF8 then V := 5 else V := 15;

      xptgRef97      {24} ,
      xptgRefN97     {6C} : begin
                              P2 := P;
                              P := Pointer(NativeInt(P) + 1);
                              if IsEx12 then begin
                                C1 := PPTGRef12(P).Col and $8FFF;
                                R1 := PPTGRef12(P).Row and $8FFFFFFF;
                                if ForceAdjust or not IsAbsRowEx12(PPTGRef12(P).Row) then
                                  If not LockStartRow Then
                                    R1 := NativeInt(PPTGRef12(P).Row) + DRow
                                  else
                                    R1 := PPTGRef12(P).Row;
                                if ForceAdjust or not IsAbsColEx12(PPTGRef12(P).Col) then
                                  C1 := PPTGRef12(P).Col + DCol;
                                Ok := (C1 >= 0) and (C1 <= MAXCOL) and (R1 >= 0) and (R1 <= MAXROW);
                                if Ok then begin
                                  PPTGRef12(P).Col := PPTGRef12(P).Col + C1;
                                  PPTGRef12(P).Row := R1;
                                end
                                else
                                  P2[0] := (P2[0] and $F0) + $0A;
                                V := SizeOf(TPTGRef12);
                              end
                              else if BIFF8 then begin
                                C1 := PPTGRef8(P).Col and $3FFF;
                                R1 := PPTGRef8(P).Row;
                                if ForceAdjust or ((PPTGRef8(P).Col and $8000) = $8000) then
                                  If not LockStartRow Then
                                    R1 := PPTGRef8(P).Row + DRow
                                  else
                                    R1 := PPTGRef8(P).Row;
                                if ForceAdjust or ((PPTGRef8(P).Col and $4000) = $4000) then
                                  C1 := (PPTGRef8(P).Col and $3FFF) + DCol;
                                Ok := (C1 >= 0) and (C1 <= $FF) and (R1 >= 0) and (R1 <= $FFFF);
                                if Ok then begin
                                  PPTGRef8(P).Col := (PPTGRef8(P).Col and $C000) + C1;
                                  PPTGRef8(P).Row := R1;
                                end
                                else
                                  P2[0] := (P2[0] and $F0) + $0A;
                                V := SizeOf(TPTGRef8);
                              end
                              else begin
                                C1 := PPTGRef7(P).Col;
                                R1 := PPTGRef7(P).Row;
                                if ForceAdjust or ((PPTGRef7(P).Row and $8000) = $8000) then
                                  C1 := PPTGRef7(P).Col + DCol;
                                if ForceAdjust or ((PPTGRef7(P).Row and $4000) = $4000) then
                                  If not LockStartRow Then
                                    R1 := (PPTGRef7(P).Row and $3FFF) + DRow
                                  else
                                    R1 := (PPTGRef7(P).Row and $3FFF);
                                Ok := (C1 >= 0) and (C1 <= $FF) and (R1 >= 0) and (R1 <= $3FFF);
                                if Ok then begin
                                  PPTGRef7(P).Col := C1;
                                  PPTGRef7(P).Row := (PPTGRef7(P).Row and $C000) + R1;
                                end
                                else
                                  P2[0] := (P2[0] and $F0) + $0A;
                                V := 3;
                              end;
                            end;

      xptgRefErr97   {6A} : if BIFF8 then V := 5 else V := 4;

      xptgArea97     {25} ,
      xptgAreaN97    {6D} : begin
                              P2 := P;
                              P := Pointer(NativeInt(P) + 1);
                              if IsEx12 then begin
                                C1 := PPTGArea12(P).Col1 and $8FFF;
                                R1 := PPTGArea12(P).Row1 and $8FFFFFFF;
                                C2 := PPTGArea12(P).Col2 and $8FFF;
                                R2 := PPTGArea12(P).Row2 and $8FFFFFFF;
                                if ForceAdjust or not IsAbsRowEx12(PPTGArea12(P).Row1) then
                                  If not LockStartRow Then
                                    R1 := Integer(PPTGArea12(P).Row1) + DRow
                                  else
                                    R1 := PPTGArea12(P).Row1;
                                if ForceAdjust or not IsAbsColEx12(PPTGArea12(P).Col1) then
                                  C1 := PPTGArea12(P).Col1 + DCol;
                                if ForceAdjust or not IsAbsRowEx12(PPTGArea12(P).Row2) then
                                  R2 := Integer(PPTGArea12(P).Row2) + DRow;
                                if ForceAdjust or not IsAbsColEx12(PPTGArea12(P).Col2) then
                                  C2 := PPTGArea12(P).Col2 + DCol;
                                Ok := (C1 >= 0) and (C1 <= MAXCOL) and (R1 >= 0) and (R1 <= MAXROW) and
                                      (C2 >= 0) and (C2 <= MAXCOL) and (R2 >= 0) and (R2 <= MAXROW);
                                if Ok then begin
                                  PPTGArea12(P).Col1 := PPTGArea12(P).Col1 + C1;
                                  PPTGArea12(P).Row1 := R1;
                                  PPTGArea12(P).Col2 := PPTGArea12(P).Col2 + C2;
                                  PPTGArea12(P).Row2 := R2;
                                end
                                else
                                  P2[0] := (P2[0] and $F0) + $0B;
                                V := SizeOf(TPTGArea12);
                              end
                              else if BIFF8 then begin
                                C1 := PPTGArea8(P).Col1 and $3FFF;
                                R1 := PPTGArea8(P).Row1;
                                C2 := PPTGArea8(P).Col2 and $3FFF;
                                R2 := PPTGArea8(P).Row2;
                                if ForceAdjust or ((PPTGArea8(P).Col1 and $8000) = $8000) then
                                  If not LockStartRow Then
                                    R1 := PPTGArea8(P).Row1 + DRow
                                  else
                                    R1 := PPTGArea8(P).Row1;
                                if ForceAdjust or ((PPTGArea8(P).Col1 and $4000) = $4000) then
                                  C1 := (PPTGArea8(P).Col1 and $3FFF) + DCol;
                                if ForceAdjust or ((PPTGArea8(P).Col2 and $8000) = $8000) then
                                  R2 := PPTGArea8(P).Row2 + DRow;
                                if ForceAdjust or ((PPTGArea8(P).Col2 and $4000) = $4000) then
                                  C2 := (PPTGArea8(P).Col2 and $3FFF) + DCol;
                                Ok := (C1 >= 0) and (C1 <= $FF) and (R1 >= 0) and (R1 <= $FFFF) and
                                      (C2 >= 0) and (C2 <= $FF) and (R2 >= 0) and (R2 <= $FFFF);
                                if Ok then begin
                                  PPTGArea8(P).Col1 := (PPTGArea8(P).Col1 and $C000) + C1;
                                  PPTGArea8(P).Row1 := R1;
                                  PPTGArea8(P).Col2 := (PPTGArea8(P).Col2 and $C000) + C2;
                                  PPTGArea8(P).Row2 := R2;
                                end
                                else
                                  P2[0] := (P2[0] and $F0) + $0B;
                                V := SizeOf(TPTGArea8);
                              end
                              else begin
                                C1 := PPTGArea7(P).Col1;
                                R1 := PPTGArea7(P).Row1;
                                C2 := PPTGArea7(P).Col2;
                                R2 := PPTGArea7(P).Row2;
                                if ForceAdjust or ((PPTGArea7(P).Row1 and $8000) = $8000) then
                                  C1 := PPTGArea7(P).Col1 + DCol;
                                if ForceAdjust or ((PPTGArea7(P).Row1 and $4000) = $4000) then
                                  If not LockStartRow Then
                                    R1 := (PPTGArea7(P).Row1 and $3FFF) + DRow
                                  else
                                    R1 := (PPTGArea7(P).Row1 and $3FFF);
                                if ForceAdjust or ((PPTGArea7(P).Row2 and $8000) = $8000) then
                                  C2 := PPTGArea7(P).Col2 + DCol;
                                if ForceAdjust or ((PPTGArea7(P).Row2 and $4000) = $4000) then
                                  R2 := (PPTGArea7(P).Row2 and $3FFF) + DRow;
                                Ok := (C1 >= 0) and (C1 <= $FF) and (R1 >= 0) and (R1 <= $3FFF) and
                                      (C2 >= 0) and (C2 <= $FF) and (R2 >= 0) and (R2 <= $3FFF);
                                if Ok then begin
                                  PPTGArea7(P).Col1 := C1;
                                  PPTGArea7(P).Row1 := (PPTGArea7(P).Row1 and $C000) + R1;
                                  PPTGArea7(P).Col2 := C2;
                                  PPTGArea7(P).Row2 := (PPTGArea7(P).Row2 and $C000) + R2;
                                end
                                else
                                  P2[0] := (P2[0] and $F0) + $0B;
                                V := 6;
                              end;
                            end;

      xptgAreaErr97  {6B} : if BIFF8 then V := 9 else V := 7;

      xptgMemArea97  {66} : V := Word(Pointer(NativeInt(P) + 1)^) + 7;

      $2E            {6E} : V := Word(Pointer(NativeInt(P) + 1)^) + 3;

      $27            {67} : V := Word(Pointer(NativeInt(P) + 1)^) + 7;

      $28            {68} : V := Word(Pointer(NativeInt(P) + 1)^) + 7;

      $2F            {6F} : V := Word(Pointer(NativeInt(P) + 1)^) + 3;

      xptgMemFunc97  {69} : V := Word(Pointer(NativeInt(P) + 1)^) + 3;

      xptgNameX97    {79} : if BIFF8 then V := 5 else V := 25;

      xptgRef3d97    {7A} : begin
                              P2 := P;
                              P := Pointer(NativeInt(P) + 1);
                              if IsEx12 then begin
                                C1 := PPTGRef3d12(P).Col and $8FFF;
                                R1 := PPTGRef3d12(P).Row and $8FFFFFFF;
                                if ForceAdjust or not IsAbsRowEx12(PPTGRef3d12(P).Col) then
                                  If not LockStartRow Then
                                    R1 := Integer(PPTGRef3d12(P).Row) + DRow
                                  else
                                    R1 := PPTGRef3d12(P).Row;
                                if ForceAdjust or not IsAbsColEx12(PPTGRef3d12(P).Col) then
                                  C1 := PPTGRef3d12(P).Col + DCol;
                                Ok := (C1 >= 0) and (C1 <= MAXCOL) and (R1 >= 0) and (R1 <= MAXROW);
                                if Ok then begin
                                  PPTGRef3d12(P).Col := PPTGRef3d12(P).Col + C1;
                                  PPTGRef3d12(P).Row := R1;
                                end
                                else
                                  P2[0] := (P2[0] and $F0) + $0C;
                                V := SizeOf(TPTGRef3d12);
                              end
                              else if BIFF8 then begin
                                C1 := PPTGRef3d8(P).Col and $3FFF;
                                R1 := PPTGRef3d8(P).Row;
                                if ForceAdjust or ((PPTGRef3d8(P).Col and $8000) = $8000) then
                                  If not LockStartRow Then
                                    R1 := PPTGRef3d8(P).Row + DRow
                                  else
                                    R1 := PPTGRef3d8(P).Row;
                                if ForceAdjust or ((PPTGRef3d8(P).Col and $4000) = $4000) then
                                  C1 := (PPTGRef3d8(P).Col and $3FFF) + DCol;
                                Ok := (C1 >= 0) and (C1 <= $FF) and (R1 >= 0) and (R1 <= $FFFF);
                                if Ok then begin
                                  PPTGRef3d8(P).Col := (PPTGRef3d8(P).Col and $C000) + C1;
                                  PPTGRef3d8(P).Row := R1;
                                end
                                else
                                  P2[0] := (P2[0] and $F0) + $0C;
                                V := SizeOf(TPTGRef3d8);
                              end
                              else begin
                                C1 := PPTGRef3d7(P).Col;
                                R1 := PPTGRef3d7(P).Row;
                                if ForceAdjust or ((PPTGRef3d7(P).Row and $8000) = $8000) then
                                  C1 := PPTGRef3d7(P).Col + DCol;
                                if ForceAdjust or ((PPTGRef3d7(P).Row and $4000) = $4000) then
                                  If not LockStartRow Then
                                    R1 := (PPTGRef3d7(P).Row and $3FFF) + DRow
                                  else
                                    R1 := (PPTGRef3d7(P).Row and $3FFF);
                                Ok := (C1 >= 0) and (C1 <= $FF) and (R1 >= 0) and (R1 <= $3FFF);
                                if Ok then begin
                                  PPTGRef3d7(P).Col := C1;
                                  PPTGRef3d7(P).Row := (PPTGRef3d7(P).Row and $C000) + R1;
                                end
                                else
                                  P2[0] := (P2[0] and $F0) + $0C;
                                V := 16;
                              end;
                            end;

      xptgRefErr3d97 {7C} : if BIFF8 then V := 7 else V := 17;

      xptgArea3d97   {7B} : begin
                              P2 := P;
                              P := Pointer(NativeInt(P) + 1);
                              if IsEx12 then begin
                                C1 := PPTGArea3d12(P).Col1 and $7FFF;
                                R1 := PPTGArea3d12(P).Row1 and $7FFFFFFF;
                                C2 := PPTGArea3d12(P).Col2 and $7FFF;
                                R2 := PPTGArea3d12(P).Row2 and $7FFFFFFF;
                                if ForceAdjust or not IsAbsRowEx12(PPTGArea3d12(P).Row1) then
                                  if not LockStartRow then
                                    R1 := Integer(PPTGArea3d12(P).Row1) + DRow
                                  else
                                    R1 := PPTGArea3d12(P).Row1;
                                if ForceAdjust or not IsAbsColEx12(PPTGArea3d12(P).Col1) then
                                  C1 := PPTGArea3d12(P).Col1 + DCol;
                                if ForceAdjust or not IsAbsRowEx12(PPTGArea3d12(P).Row2) then
                                  R2 := Integer(PPTGArea3d12(P).Row2) + DRow;
                                if ForceAdjust or not IsAbsColEx12(PPTGArea3d12(P).Col2) then
                                  C2 := PPTGArea3d12(P).Col2 + DCol;
                                Ok := (C1 >= 0) and (C1 <= MAXCOL) and (R1 >= 0) and (R1 <= MAXROW) and
                                      (C2 >= 0) and (C2 <= MAXCOL) and (R2 >= 0) and (R2 <= MAXROW);
                                if Ok then begin
                                  PPTGArea3d12(P).Col1 := PPTGArea3d12(P).Col1 + C1;
                                  PPTGArea3d12(P).Row1 := R1;
                                  PPTGArea3d12(P).Col2 := PPTGArea3d12(P).Col2 + C2;
                                  PPTGArea3d12(P).Row2 := R2;
                                end
                                else
                                  P2[0] := (P2[0] and $F0) + $0D;
                                V := SizeOf(TPTGArea3d12);
                              end
                              else if BIFF8 then begin
                                C1 := PPTGArea3d8(P).Col1 and $3FFF;
                                R1 := PPTGArea3d8(P).Row1;
                                C2 := PPTGArea3d8(P).Col2 and $3FFF;
                                R2 := PPTGArea3d8(P).Row2;
                                if ForceAdjust or ((PPTGArea3d8(P).Col1 and $8000) = $8000) then
                                  If not LockStartRow Then
                                    R1 := PPTGArea3d8(P).Row1 + DRow
                                  else
                                    R1 := PPTGArea3d8(P).Row1;
                                if ForceAdjust or ((PPTGArea3d8(P).Col1 and $4000) = $4000) then
                                  C1 := (PPTGArea3d8(P).Col1 and $3FFF) + DCol;
                                if ForceAdjust or ((PPTGArea3d8(P).Col2 and $8000) = $8000) then
                                  R2 := PPTGArea3d8(P).Row2 + DRow;
                                if ForceAdjust or ((PPTGArea3d8(P).Col2 and $4000) = $4000) then
                                  C2 := (PPTGArea3d8(P).Col2 and $3FFF) + DCol;
                                Ok := (C1 >= 0) and (C1 <= $FF) and (R1 >= 0) and (R1 <= $FFFF) and
                                      (C2 >= 0) and (C2 <= $FF) and (R2 >= 0) and (R2 <= $FFFF);
                                if Ok then begin
                                  PPTGArea3d8(P).Col1 := (PPTGArea3d8(P).Col1 and $C000) + C1;
                                  PPTGArea3d8(P).Row1 := R1;
                                  PPTGArea3d8(P).Col2 := (PPTGArea3d8(P).Col2 and $C000) + C2;
                                  PPTGArea3d8(P).Row2 := R2;
                                end
                                else
                                  P2[0] := (P2[0] and $F0) + $0D;
                                V := SizeOf(TPTGArea3d8);
                              end
                              else begin
                                C1 := PPTGArea3d7(P).Col1;
                                R1 := PPTGArea3d7(P).Row1;
                                C2 := PPTGArea3d7(P).Col2;
                                R2 := PPTGArea3d7(P).Row2;
                                if ForceAdjust or ((PPTGArea3d7(P).Row1 and $8000) = $8000) then
                                  C1 := PPTGArea3d7(P).Col1 + DCol;
                                if ForceAdjust or ((PPTGArea3d7(P).Row1 and $4000) = $4000) then
                                  If not LockStartRow Then
                                    R1 := (PPTGArea3d7(P).Row1 and $3FFF) + DRow
                                  else
                                    R1 := (PPTGArea3d7(P).Row1 and $3FFF);
                                if ForceAdjust or ((PPTGArea3d7(P).Row2 and $8000) = $8000) then
                                  C2 := PPTGArea3d7(P).Col2 + DCol;
                                if ForceAdjust or ((PPTGArea3d7(P).Row2 and $4000) = $4000) then
                                  R2 := (PPTGArea3d7(P).Row2 and $3FFF) + DRow;
                                Ok := (C1 >= 0) and (C1 <= $FF) and (R1 >= 0) and (R1 <= $3FFF) and
                                      (C2 >= 0) and (C2 <= $FF) and (R2 >= 0) and (R2 <= $3FFF);
                                if Ok then begin
                                  PPTGArea3d7(P).Col1 := C1;
                                  PPTGArea3d7(P).Row1 := (PPTGArea3d7(P).Row1 and $C000) + R1;
                                  PPTGArea3d7(P).Col2 := C2;
                                  PPTGArea3d7(P).Row2 := (PPTGArea3d7(P).Row2 and $C000) + R2;
                                end
                                else
                                  P2[0] := (P2[0] and $F0) + $0D;
                                V := 20;
                              end;
                            end;

      xptgAreaErr3d97{7D} : if BIFF8 then V := 11 else V := 21;

      $38            {38} : V := 3; // Not sure how to handle these.
      else
        raise XLSRWException.CreateFmt('Unknown ptg[%.2X] in AdjustCell',[P[0]]);
//        V := 1;
    end;
    P := Pointer(NativeInt(P) + V);
  end;
end;

procedure AdjustCell2(Buf: Pointer; Len,Col,Row,SheetIndex,DCol,DRow: integer);
var
  V,C1,R1,C2,R2: integer;
  Ok: boolean;
  P,P2: PByteArray;
begin
  P := Buf;
  while (NativeInt(P) - NativeInt(Buf)) < Len do begin
    case P[0] of
      $02            {02} : V := 5;
      xptgStr97     {17} : begin
        P := Pointer(NativeInt(P) + 1);
        if PByteArray(P)[1] = 0 then
          V := PByteArray(P)[0] + 2
        else
          V := (PByteArray(P)[0] * 2) + 2;
      end;
      xptgAttr97        {19} : begin
        P := Pointer(NativeInt(P) + 1);

        if P[0] = $04 then begin
          P := Pointer(NativeInt(P) + 1);
          V := (Word(Pointer(P)^) + 1) * SizeOf(word) + 2;
        end
        else
          V := 3;
      end;
      xptgErr         {1C} : V := 2;
      xptgBool        {1D} : V := 2;
      xptgInt97       {1E} : V := 3;
      xptgNum         {1F} : V := 9;

      xptgArray97    {60} : V := 8;

      xptgFunc97    {61} : V := 4;

      xptgFuncVar97    {62} : V := 5;

      xptgName97    {63} : V := 5;

      xptgRef97    {64} ,
      xptgRefN97    {6C} : begin
                              P2 := P;
                              P := Pointer(NativeInt(P) + 1);
                              C1 := PPTGRef8(P).Col and $3FFF;
                              R1 := PPTGRef8(P).Row;
                              if ((PPTGRef8(P).Col and $8000) = $8000) and (R1 >= Row) then
                                Inc(R1,DRow);
                              if ((PPTGRef8(P).Col and $4000) = $4000) and (C1 >= Col) then
                                Inc(C1,DCol);
                              Ok := (C1 >= 0) and (C1 <= $FF) and (R1 >= 0) and (R1 <= $FFFF);
                              if Ok then begin
                                PPTGRef8(P).Col := (PPTGRef8(P).Col and $C000) + C1;
                                PPTGRef8(P).Row := R1;
                              end
                              else
                                P2[0] := (P2[0] and $F0) + $0A;
                              V := 4;
                            end;

      xptgRefErr97    {6A} : V := 5;

      xptgArea97    {25} ,
      xptgAreaN97    {6D} : begin
                              P2 := P;
                              P := Pointer(NativeInt(P) + 1);
                              C1 := PPTGArea8(P).Col1 and $3FFF;
                              R1 := PPTGArea8(P).Row1;
                              C2 := PPTGArea8(P).Col2 and $3FFF;
                              R2 := PPTGArea8(P).Row2;
                              if ((PPTGArea8(P).Col1 and $8000) = $8000) and (R1 >= Row) then
                                Inc(R1,DRow);
                              if ((PPTGArea8(P).Col1 and $4000) = $4000) and (C1 >= Col) then
                                Inc(C1,DCol);
                              if ((PPTGArea8(P).Col2 and $8000) = $8000) and (R2 >= Row) then
                                Inc(R2,DRow);
                              if ((PPTGArea8(P).Col2 and $4000) = $4000) and (C2 >= Col) then
                                Inc(C2,DCol);
                              Ok := (C1 >= 0) and (C1 <= $FF) and (R1 >= 0) and (R1 <= $FFFF) and
                                    (C2 >= 0) and (C2 <= $FF) and (R2 >= 0) and (R2 <= $FFFF);
                              if Ok then begin
                                PPTGArea8(P).Col1 := (PPTGArea8(P).Col1 and $C000) + C1;
                                PPTGArea8(P).Row1 := R1;
                                PPTGArea8(P).Col2 := (PPTGArea8(P).Col2 and $C000) + C2;
                                PPTGArea8(P).Row2 := R2;
                              end
                              else
                                P2[0] := (P2[0] and $F0) + $0B;
                              V := 8;
                            end;

      xptgAreaErr97    {6B} : V := 9;

      xptgMemArea97    {66} : V := Word(Pointer(NativeInt(P) + 1)^) + 7;

      $2E              {6E} : V := Word(Pointer(NativeInt(P) + 1)^) + 3;

      $27              {67} : V := Word(Pointer(NativeInt(P) + 1)^) + 7;

      $28              {68} : V := Word(Pointer(NativeInt(P) + 1)^) + 7;

      $2F              {6F} : V := Word(Pointer(NativeInt(P) + 1)^) + 3;

      xptgMemFunc97    {69} : V := Word(Pointer(NativeInt(P) + 1)^) + 3;

      xptgNameX97    {79} : V := 7;

      xptgRef3d97    {7A} : begin
//                              P2 := P;
                              P := Pointer(NativeInt(P) + 1);
{
                              C1 := PPTGRef3d8(P).Col and $3FFF;
                              R1 := PPTGRef3d8(P).Row;
                              if ForceAdjust or ((PPTGRef3d8(P).Col and $8000) = $8000) then
                                If not LockStartRow Then
                                  R1 := PPTGRef3d8(P).Row + DRow
                                else
                                  R1 := PPTGRef3d8(P).Row;
                              if ForceAdjust or ((PPTGRef3d8(P).Col and $4000) = $4000) then
                                C1 := (PPTGRef3d8(P).Col and $3FFF) + DCol;
                              Ok := (C1 >= 0) and (C1 <= $FF) and (R1 >= 0) and (R1 <= $FFFF);
                              if Ok then begin
                                PPTGRef3d8(P).Col := (PPTGRef3d8(P).Col and $C000) + C1;
                                PPTGRef3d8(P).Row := R1;
                              end
                              else
                                P2[0] := (P2[0] and $F0) + $0C;
}
                              V := 6;
                            end;

      xptgRefErr3d97    {7C} : V := 7;

      xptgArea3d97    {7B} : begin
//                              P2 := P;
                              P := Pointer(NativeInt(P) + 1);
{
                              C1 := PPTGArea3d8(P).Col1 and $3FFF;
                              R1 := PPTGArea3d8(P).Row1;
                              C2 := PPTGArea3d8(P).Col2 and $3FFF;
                              R2 := PPTGArea3d8(P).Row2;
                              if ForceAdjust or ((PPTGArea3d8(P).Col1 and $8000) = $8000) then
                                If not LockStartRow Then
                                  R1 := PPTGArea3d8(P).Row1 + DRow
                                else
                                  R1 := PPTGArea3d8(P).Row1;
                              if ForceAdjust or ((PPTGArea3d8(P).Col1 and $4000) = $4000) then
                                C1 := (PPTGArea3d8(P).Col1 and $3FFF) + DCol;
                              if ForceAdjust or ((PPTGArea3d8(P).Col2 and $8000) = $8000) then
                                R2 := PPTGArea3d8(P).Row2 + DRow;
                              if ForceAdjust or ((PPTGArea3d8(P).Col2 and $4000) = $4000) then
                                C2 := (PPTGArea3d8(P).Col2 and $3FFF) + DCol;
                              Ok := (C1 >= 0) and (C1 <= $FF) and (R1 >= 0) and (R1 <= $FFFF) and
                                    (C2 >= 0) and (C2 <= $FF) and (R2 >= 0) and (R2 <= $FFFF);
                              if Ok then begin
                                PPTGArea3d8(P).Col1 := (PPTGArea3d8(P).Col1 and $C000) + C1;
                                PPTGArea3d8(P).Row1 := R1;
                                PPTGArea3d8(P).Col2 := (PPTGArea3d8(P).Col2 and $C000) + C2;
                                PPTGArea3d8(P).Row2 := R2;
                              end
                              else
                                P2[0] := (P2[0] and $F0) + $0D;
}
                              V := 10;
                            end;

      xptgAreaErr3d97    {7D} : V := 11;

      $38                {58} : V := 3; // Not sure how to handle these.
      else
//        raise XLSRWException.CreateFmt('Unknown ptg[%.2X] in AdjustCell',[P[0]]);
        V := 1;
    end;
    P := Pointer(NativeInt(P) + V);
  end;
end;

procedure AdjustSheet(Buf: PByteArray; Len,NewExtIndex: integer);
var
  V: integer;
  IsEx12: boolean;
  P: PByteArray;
  BIFF8: boolean;
begin
  BIFF8 := True;
  P := Buf;
  IsEx12 := False;
  while (NativeInt(P) - NativeInt(Buf)) < Len do begin
    case P[0] of
      $02            {02} : V := 5;
      xptgOpAdd         {03} : V := 1;
      xptgOpSub         {04} : V := 1;
      xptgOpMult         {05} : V := 1;
      xptgOpDiv         {06} : V := 1;
      xptgOpPower       {07} : V := 1;
      xptgOpConcat      {08} : V := 1;
      xptgOpLT          {09} : V := 1;
      xptgOpLE          {0A} : V := 1;
      xptgOpEQ          {0B} : V := 1;
      xptgOpGE          {0C} : V := 1;
      xptgOpGT          {0D} : V := 1;
      xptgOpNE          {0E} : V := 1;
      xptgOpIsect       {0F} : V := 1;
      xptgOpUnion       {10} : V := 1;
      xptgOpRange       {11} : V := 1;
      xptgOpUplus       {12} : V := 1;
      xptgOpUminus      {13} : V := 1;
      xptgOpPercent     {14} : V := 1;
      xptgLPar       {15} : V := 1;
      xptgMissArg97   {16} : V := 1;
      xptgStr97         {17} : begin
        P := Pointer(NativeInt(P) + 1);
        if BIFF8 then begin
          if PByteArray(P)[1] = 0 then
            V := PByteArray(P)[0] + 2
          else
            V := (PByteArray(P)[0] * 2) + 2;
        end
        else
          V := PByteArray(P)[0] + 1;
      end;
      xptgAttr97        {19} : begin
        P := Pointer(NativeInt(P) + 1);

        if P[0] = $04 then begin
          P := Pointer(NativeInt(P) + 1);
          V := (Word(Pointer(P)^) + 1) * SizeOf(word) + 2;
        end
        else
          V := 3;
      end;
      xptgErr         {1C} : V := 2;
      xptgBool        {1D} : V := 2;
      xptgInt97       {1E} : V := 3;
      xptgNum         {1F} : V := 9;

      xptgArray97    {60} : V := 8;

      xptgFunc97    {61} : V := 3;

      xptgFuncVar97    {62} : V := 4;

      xptgName97    {63} : if BIFF8 then V := 5 else V := 15;

      xptgRef97    {24} ,
      xptgRefN97    {6C} : begin
                              P := Pointer(NativeInt(P) + 1);
                              if IsEx12 then
                                V := SizeOf(TPTGRef12)
                              else if BIFF8 then
                                V := SizeOf(TPTGRef8)
                              else
                                V := 3;
                            end;

      xptgRefErr97    {6A} : if BIFF8 then V := 5 else V := 4;

      xptgArea97    {25} ,
      xptgAreaN97    {6D} : begin
                              P := Pointer(NativeInt(P) + 1);
                              if IsEx12 then
                                V := SizeOf(TPTGArea12)
                              else if BIFF8 then
                                V := SizeOf(TPTGArea8)
                              else
                                V := 6;
                            end;

      xptgAreaErr97    {6B} : if BIFF8 then V := 9 else V := 7;

      xptgMemArea97    {66} : V := Word(Pointer(NativeInt(P) + 1)^) + 7;

      $2E              {6E} : V := Word(Pointer(NativeInt(P) + 1)^) + 3;

      $27              {67} : V := Word(Pointer(NativeInt(P) + 1)^) + 7;

      $28              {68} : V := Word(Pointer(NativeInt(P) + 1)^) + 7;

      $2F              {6F} : V := Word(Pointer(NativeInt(P) + 1)^) + 3;

      xptgMemFunc97    {69} : V := Word(Pointer(NativeInt(P) + 1)^) + 3;

      xptgNameX97    {79} : if BIFF8 then V := 5 else V := 25;

      xptgRef3d97    {7A} : begin
                              P := Pointer(NativeInt(P) + 1);
                              if IsEx12 then begin
                                PPTGRef3d12(P).Index := NewExtIndex;
                                V := SizeOf(TPTGRef3d12);
                              end
                              else if BIFF8 then begin
                                PPTGRef3d8(P).Index := NewExtIndex;
                                V := SizeOf(TPTGRef3d8);
                              end
                              else begin
                                PPTGRef3d7(P).Index := NewExtIndex;
                                V := 16;
                              end;
                            end;

      xptgRefErr3d97    {7C} : if BIFF8 then V := 7 else V := 17;

      xptgArea3d97    {7B} : begin
                              P := Pointer(NativeInt(P) + 1);
                              if IsEx12 then begin
                                PPTGArea3d12(P).Index := NewExtIndex;
                                V := SizeOf(TPTGArea3d12);
                              end
                              else if BIFF8 then begin
                                PPTGArea3d8(P).Index := NewExtIndex;
                                V := SizeOf(TPTGArea3d8);
                              end
                              else begin
                                PPTGArea3d7(P).SheetIndex := NewExtIndex;
                                V := 20;
                              end;
                            end;

      xptgAreaErr3d97    {7D} : if BIFF8 then V := 11 else V := 21;

      $38                {38} : V := 3; // Not sure how to handle these.
      else
        raise XLSRWException.CreateFmt('Unknown ptg[%.2X] in AdjustCell',[P[0]]);
//        V := 1;
    end;
    P := Pointer(NativeInt(P) + V);
  end;
end;

procedure ConvertShrFmla(BIFF8: boolean; Buf: Pointer; Len,ACol,ARow: integer);
var
  V,C,R: integer;
  P: Pointer;

procedure DecodeArea7(Cin: byte; Rin: word; var Cout,Rout: integer);
begin
  if (Rin and $8000) = 0 then
    Rout := Rin and $FF
  else
    Rout := ARow + Rin;
  if (Rin and $4000) = 0 then
    Cout := Cin
  else
    Cout := ACol + Cin;
end;

procedure DecodeArea8(Cin,Rin: integer; var Cout,Rout: integer);
begin
  if (Cin and $4000) = 0 then
    Cout := Cin and $FF
  else
    Cout := ACol + Shortint(Cin and $FF);
  if (Cin and $8000) = 0 then
    Rout := Rin
  else
    Rout := ARow + Rin;
end;

begin
  P := Buf;
  while (NativeInt(P) - NativeInt(Buf)) < Len do begin
    case GetBasePtgs(Byte(P^)) of
      $02            {02} : V := 5;
      xptgOpAdd         {03} : V := 1;
      xptgOpSub         {04} : V := 1;
      xptgOpMult        {05} : V := 1;
      xptgOpDiv         {06} : V := 1;
      xptgOpPower       {07} : V := 1;
      xptgOpConcat      {08} : V := 1;
      xptgOpLT          {09} : V := 1;
      xptgOpLE          {0A} : V := 1;
      xptgOpEQ          {0B} : V := 1;
      xptgOpGE          {0C} : V := 1;
      xptgOpGT          {0D} : V := 1;
      xptgOpNE          {0E} : V := 1;
      xptgOpIsect       {0F} : V := 1;
      xptgOpUnion       {10} : V := 1;
      xptgOpRange       {11} : V := 1;
      xptgOpUplus       {12} : V := 1;
      xptgOpUminus      {13} : V := 1;
      xptgOpPercent     {14} : V := 1;
      xptgLPar       {15} : V := 1;
      xptgMissArg97     {16} : V := 1;
      xptgStr97         {17} : begin
        P := Pointer(NativeInt(P) + 1);
        if BIFF8 then begin
          if PByteArray(P)[1] = 0 then
            V := PByteArray(P)[0] + 2
          else
            V := (PByteArray(P)[0] * 2) + 2;
        end
        else
          V := PByteArray(P)[0] + 1;
      end;
      xptgAttr97       {19} : begin
        P := Pointer(NativeInt(P) + 1);

        if Byte(P^) = $04 then begin
          P := Pointer(NativeInt(P) + 1);
          V := (Word(P^) + 1) * SizeOf(word) + 2;
        end
        else
          V := 3;
      end;
      xptgErr         {1C} : V := 2;
      xptgBool        {1D} : V := 2;
      xptgInt97       {1E} : V := 3;
      xptgNum         {1F} : V := 9;

      xptgArray97    {60} : V := 8;

      xptgFunc97    {61} : V := 3;

      xptgFuncVar97    {62} : V := 4;

      xptgName97    {63} : if BIFF8 then V := 5 else V := 15;

      xptgRef97    {64} : if BIFF8 then V := 5 else V := 4;

      xptgArea97    {65} : if BIFF8 then V := 9 else V := 7;

      xptgMemArea97    {66} : V := Word(Pointer(NativeInt(P) + 1)^) + 7;

      $2E            {6E} : V := Word(Pointer(NativeInt(P) + 1)^) + 3;

      $27            {67} : V := Word(Pointer(NativeInt(P) + 1)^) + 7;

      $28            {68} : V := Word(Pointer(NativeInt(P) + 1)^) + 7;

      $2F            {6F} : V := Word(Pointer(NativeInt(P) + 1)^) + 3;

      xptgMemFunc97    {69} : V := Word(Pointer(NativeInt(P) + 1)^) + 3;

      xptgRefErr97    {6A} : if BIFF8 then V := 5 else V := 4;

      xptgAreaErr97    {6B} : if BIFF8 then V := 9 else V := 7;

      xptgRefN97    {6C} : begin
        case Byte(P^) of
          xptgRefN97:  Byte(P^) := xptgRef97;
          xptgRefNV97: Byte(P^) := xptgRefV97;
          xptgRefNA97: Byte(P^) := xptgRefA97;
        end;
        P := Pointer(NativeInt(P) + 1);
        if BIFF8 then begin
          DecodeArea8(PPTGRef8(P).Col,PPTGRef8(P).Row,C,R);
          PPTGRef8(P).Row := R;
          PPTGRef8(P).Col := (PPTGRef8(P).Col and $C000) + C;
          V := SizeOf(TPTGRef8);
        end
        else begin
          DecodeArea7(PPTGRef7(P).Col,PPTGRef7(P).Row,C,R);
          PPTGRef7(P).Row := (PPTGRef7(P).Row and $C000) + R;
          PPTGRef7(P).Col := C;
          V := SizeOf(TPTGRef7);
        end;
      end;
      xptgAreaN97    {6D} : begin
        case Byte(P^) of
          xptgAreaN97:  Byte(P^) := xptgArea97;
          xptgAreaNV97: Byte(P^) := xptgAreaV97;
          xptgAreaNA97: Byte(P^) := xptgAreaA97;
        end;
        P := Pointer(NativeInt(P) + 1);
        if BIFF8 then begin
          DecodeArea8(PPTGArea8(P).Col1,PPTGArea8(P).Row1,C,R);
          PPTGArea8(P).Row1 := R;
          PPTGArea8(P).Col1 := (PPTGArea8(P).Col1 and $C000) + C;

          DecodeArea8(PPTGArea8(P).Col2,PPTGArea8(P).Row2,C,R);
          PPTGArea8(P).Row2 := R;
          PPTGArea8(P).Col2 := (PPTGArea8(P).Col2 and $C000) + C;
          V := SizeOf(TPTGArea8);
        end
        else begin
          DecodeArea7(PPTGArea7(P).Col1,PPTGArea7(P).Row1,C,R);
          PPTGArea7(P).Row1 := (PPTGArea7(P).Row1 and $C000) + R;
          PPTGArea7(P).Col1 := C;

          DecodeArea7(PPTGArea7(P).Col2,PPTGArea7(P).Row2,C,R);
          PPTGArea7(P).Row2 := (PPTGArea7(P).Row2 and $C000) + R;
          PPTGArea7(P).Col2 := C;
          V := SizeOf(TPTGArea7);
        end;
      end;

      xptgNameX97    {79} : if BIFF8 then V := 7 else V := 25;

      xptgRef3d97    {7A} : if BIFF8 then V := 7 else V := 17;

      xptgArea3d97    {7B} : if BIFF8 then V := 11 else V := 21;

      xptgRefErr3d97    {7C} : if BIFF8 then V := 7 else V := 17;

      xptgAreaErr3d97    {7D} : if BIFF8 then V := 11 else V := 21;

      $38            {38} : V := 3; // Not sure how to handle these (ptgFuncCE).
      else
        raise XLSRWException.CreateFmt('Unknown ptg[%.2X] in Shared Formula',[Byte(P^)]);
//        V := 1;
    end;
    P := Pointer(NativeInt(P) + V);
  end;
end;

initialization
  StrTRUE := 'TRUE';
  StrFALSE := 'FALSE';
  FuncArgSeparator := FormatSettings.ListSeparator;

end.
