unit BIFF_Validate5;

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

uses Classes, SysUtils,
     BIFF_Utils5, BIFF_FormulaHandler5, BIFF_RecsII5, BIFF_CellAreas5, BIFF_Stream5,
     BIFF_MoveCopy5, BIFF_ExcelFuncII5,
     XLSUtils5, XLSFormulaTypes5;

//* Validation input value oprator.
type TValidationOperator = (vopBetween,     //* Input value must be between Expression1 and Expression2.
                            vopNotBetween,  //* Input value must outside Expression1 and Expression2.
                            vopEqual,       //* Input value must be equal to Expression1.
                            vopNotEqual,    //* Input value must not equal to Expression1.
                            vopGreater,     //* Input value must be greater than Expression1.
                            vopLess,        //* Input value must be less than Expression1.
                            vopGreateEqual, //* Input value must be greater or equal to Expression1.
                            vopLessEqual    //* Input value must be less or equal to Expression1.
                            );

type
//* Type of values that are accepted.
     TValidationType = (vtAny,        //* Any value.
                        vtInteger,    //* Integer Values.
                        vtNumber,     //* Floating point values.
                        vtList,       //* Choose a value from a list. Use ~[link TDataValidation.SetList] to define the list.
                        vtDate,       //* Date values.
                        vtTime,       //* Time values.
                        vtTextLength, //* Limit the text length of the value.
                        vtCustom      //* Custom. Any formula?
                        );

//* Message style when the enterd value not is valid.
type TValidationStyle = (vsStop,    //* Stop the user from enetering a value not valid.
                         vsWarning, //* Give the user a varning. A value not within the valid values can be enterd.
                         vsInfo     //* Inform the user that the value´not is valid.
                         );

type
    //* Validation options.
    TValidationOption = (voAllowEmptyCells, //* Allow empty cells in the validation areas.
                         voSupressDropDown, //* When ValidationType is vtList, prevent the combo box to be shown.
                         voShowPromptBox,   //* Show a hint text to the user.
                         voShowErrorBox     //* Show an error message when an invalid value is entered.
                         );
    //* Set of validation options.
    TValidationOptions = set of TValidationOption;

type
//* Validations of input values in cells. This implements the Excel Validation command.
//* Use TDataValidation in order to create a control of the cell values a user
//* enters in Excel. If a value not matches the criterias, the user can be
//* given a message, or the value can be refused.
//* See ~[link TDataValidations] for an example on how to use validations.
    TDataValidation = class(TCollectionItemMoveCopy)
private
     FAssigned: boolean;
     FOptions: longword;
     FValidationOptions: TValidationOptions;
     FInputTitle: AxUCString;
     FErrorTitle: AxUCString;
     FInputMsg: AxUCString;
     FErrorMsg: AxUCString;
     FFormula1: TRecPTGS;
     FFormula2: TRecPTGS;
     FValidationType: TValidationType;
     FValidationStyle: TValidationStyle;
     FValidationOperator: TValidationOperator;
     FAreas: TCellAreas97;

     procedure ClearExpr1;
     procedure ClearExpr2;
     function  GetExpression1: AxUCString;
     function  GetExpression2: AxUCString;
     procedure SetExpression1(const Value: AxUCString);
     procedure SetExpression2(const Value: AxUCString);
     procedure SetValidationOperator(const Value: TValidationOperator);
     procedure SetValidationStyle(const Value: TValidationStyle);
     procedure SetValidationType(const Value: TValidationType);
     procedure SetValidationOptions(const Value: TValidationOptions);
protected
     procedure LoadFromStream(Stream: TXLSStream; PBuf: PByteArray);
     procedure SaveToStream(Stream: TXLSStream; PBuf: PByteArray);
     procedure SetRawExpression1(Value: PByteArray; Len: integer);
     procedure SetRawExpression2(Value: PByteArray; Len: integer);
     function  Intersect(Col1,Row1,Col2,Row2: integer): boolean; override;
     procedure Copy(Col1,Row1,Col2,Row2,DeltaCol,DeltaRow: integer); override;
     procedure Delete(Col1,Row1,Col2,Row2: integer); override;
     procedure Include(Col1,Row1,Col2,Row2: integer); override;
     procedure Move(DeltaCol,DeltaRow: integer); override;
     procedure Move(Col1,Row1,Col2,Row2,DeltaCol,DeltaRow: integer); override;
public
     //* ~exclude
     constructor Create(Collection: TCollection); override;
     //* ~exclude
     destructor Destroy; override;
     //* Assigns another TDataValidation to this.
     //* ~param Source The source TDataValidation object.
     procedure Assign(Source: TPersistent); override;
     //* Sets the validation type to a list.
     //* Use SetList to set a list of string values that the user can choose of.
     //* SetList will also set the ~[link ValidationType] to vtList.
     //* ~param Values An array of strings that is the list.
     procedure SetList(Values: array of AxUCString);
     //* Get the list values.
     //* Use GetList to get the strings in the list when ~[link ValidationType]
     //* is vtList. GetList returns True if if the validationType is vtList,
     //* and there where any strings in the list.
     //* Use ~[link GetListSize] to get the number of elements in the list.
     //* ~param Values An array of strings to where the list strings are copied.
     //* ~result True if there where any strings returned. 
     function  GetList(var Values: array of AxUCString): boolean;
     //* Number of strings when the validation type is a list (vtList)
     //* ~result Number of elements in the list.
     function  GetListSize: integer;
published
     //* Options for a validation.
     property Options: TValidationOptions read FValidationOptions write SetValidationOptions;
     //* The cell areas that the TDataValidation shall be applied to.
     property Areas: TCellAreas97 read FAreas write FAreas;
     //* Message to the user.
     //* ErrorMsg is the message that the user will see if the cell value
     //* entered don't matches the criterias.
     property ErrorMsg: AxUCString read FErrorMsg write FErrorMsg;
     //* Title of the message box.
     property ErrorTitle: AxUCString read FErrorTitle write FErrorTitle;
     //* Expression for the first condition.
     property Expression1: AxUCString read GetExpression1 write SetExpression1;
     //* Expression for the second condition, if needed.
     property Expression2: AxUCString read GetExpression2 write SetExpression2;
     //* Prompt box shown for the user.
     //* In order to show a prompt box for the user, ValidationOptions must
     //* include the value voShowPromptBox.
     property InputMsg: AxUCString read FInputMsg write FInputMsg;
     //* Title of prompt box.
     property InputTitle: AxUCString read FInputTitle write FInputTitle;
     //* The validation operator.
     property ValidationOperator: TValidationOperator read FValidationOperator write SetValidationOperator;
     //* Validation style.
     property ValidationStyle: TValidationStyle read FValidationStyle write SetValidationStyle;
     //* Type of validation.
     property ValidationType: TValidationType read FValidationType write SetValidationType;
     end;

//* List of data validations.
//* ~example
//* ~[sample
//* // Mark the cells for the validation yellow.
//* XLS.Sheets[0~[].Range[2,4,3,8~[].FillPatternForeColor := xcYellow;
//* // Add a new validation.
//* with XLS.Sheets[0~[].Validations.Add do begin
//*   // Set the validation area to C5:D9
//*   Areas.Add(2,4,3,8);
//*   // First valid value.
//*   Expression1 := '100';
//*   // Optional second value.
//*   Expression2 := '200';
//*   // Operator, input value must be between Expression1 and Expression2.
//*   ValidationOperator := vopBetween;
//*   // Type of accepted data.
//*   ValidationType := vtNumber;
//*   // Optional error message.
//*   ErrorMsg := 'The number must be between 100 and 200';
//*   // Optional error title.
//*   ErrorTitle := 'No good value';
//* end;
//*
//* // Mark the cells for the validation green.
//* XLS.Sheets[0~[].Range[2,9,3,12~[].FillPatternForeColor := xcGreen;
//* // Add a new validation.
//* with XLS.Sheets[0~[].Validations.Add do begin
//*   // Set the validation area to C10:D13
//*   Areas.Add(2,9,3,12);
//*   // Set an array with strings that are valid to enter.
//*   SetList(['one','two','three','four'~[]);
//*   // Warning style. The user is given a warning, but any value can be enterd.
//*   ValidationStyle := vsWarning;
//*   // Optional error message.
//*   ErrorMsg := 'Select a value from the list';
//*   // Optional error title.
//*   ErrorTitle := 'No good value';
//* end;
//* ]
type TDataValidations = class(TCollectionMoveCopy)
private
     FRecDVAL: TRecDVAL;

     function  GetDataValidation(Index: integer): TDataValidation;
protected
     FOwner: TPersistent;
     FFormulaHandler: TFormulaHandler;

     function  GetOwner: TPersistent; override;
public
     //* ~exclude
     constructor Create(AOwner: TPersistent; FormulaHandler: TFormulaHandler);
     //* Add a new validation.
     //* ~result The new TDataValidation object.
     function  Add: TDataValidation;
     //* ~exclude
     procedure LoadFromStream(Stream: TXLSStream; PBuf: PByteArray);
     //* ~exclude
     procedure SaveToStream(Stream: TXLSStream; PBuf: PByteArray);
     //* ~exclude
     function  ComboBoxObjId: integer;

     // The validations in the list.
     property Items[Index: integer]: TDataValidation read GetDataValidation; default;
     end;

implementation

{ TDataValidation }

procedure TDataValidation.Assign(Source: TPersistent);
begin
  FAssigned := True;
  FOptions := TDataValidation(Source).FOptions;
  FInputTitle := TDataValidation(Source).FInputTitle;
  FErrorTitle := TDataValidation(Source).FErrorTitle;
  FInputMsg := TDataValidation(Source).FInputMsg;
  FErrorMsg := TDataValidation(Source).FErrorMsg;

  FFormula1.Size := TDataValidation(Source).FFormula1.Size;
  GetMem(FFormula1.PTGS,FFormula1.Size);
  System.Move(TDataValidation(Source).FFormula1.PTGS^,FFormula1.PTGS^,FFormula1.Size);

  FFormula2.Size := TDataValidation(Source).FFormula2.Size;
  GetMem(FFormula2.PTGS,FFormula2.Size);
  System.Move(TDataValidation(Source).FFormula2.PTGS^,FFormula2.PTGS^,FFormula2.Size);

  FValidationType := TDataValidation(Source).FValidationType;
  FValidationStyle := TDataValidation(Source).FValidationStyle;
  FValidationOperator := TDataValidation(Source).FValidationOperator;
  FAreas.Assign(TDataValidation(Source).Areas);
end;

procedure TDataValidation.ClearExpr1;
begin
  FreeMem(FFormula1.PTGS);
  FFormula1.PTGS := Nil;
  FFormula1.Size := 0;
end;

procedure TDataValidation.ClearExpr2;
begin
  FreeMem(FFormula2.PTGS);
  FFormula2.PTGS := Nil;
  FFormula2.Size := 0;
end;

procedure TDataValidation.Copy(Col1, Row1, Col2, Row2, DeltaCol, DeltaRow: integer);
begin
  FAreas.Copy(Col1, Row1, Col2, Row2,DeltaCol,DeltaRow);
end;

constructor TDataValidation.Create(Collection: TCollection);
begin
  inherited Create(Collection);
  FAreas := TCellAreas97.Create;
  FValidationType := vtNumber;
  FValidationStyle := vsStop;
  FValidationOperator := vopEqual;
  FValidationOptions := [voAllowEmptyCells,voShowPromptBox,voShowErrorBox];
end;

procedure TDataValidation.Delete(Col1, Row1, Col2, Row2: integer);
begin
  FAreas.Delete(Col1, Row1, Col2, Row2);
end;

destructor TDataValidation.Destroy;
begin
  inherited;
  ClearExpr1;
  ClearExpr2;
  FAreas.Free;
end;

function TDataValidation.GetExpression1: AxUCString;
begin
  if (FFormula1.PTGS <> Nil) and (FFormula1.Size > 0) then
    Result := TDataValidations(Collection).FFormulaHandler.DecodeFormula(FFormula1.PTGS,FFormula1.Size)
  else
    Result := '';
end;

function TDataValidation.GetExpression2: AxUCString;
begin
  if (FFormula2.PTGS <> Nil) and (FFormula2.Size > 0) then
    Result := TDataValidations(Collection).FFormulaHandler.DecodeFormula(FFormula1.PTGS,FFormula1.Size)
  else
    Result := '';
end;

function TDataValidation.GetList(var Values: array of AxUCString): boolean;
var
  i,j,k,Sz,Count: integer;
  pB: PByteArray;
  pW: PWordArray;
begin
  Result := False;
  Count := GetListSize;
  if Count <= 0 then
    Exit;
  Result := True;
  Sz := FFormula1.PTGS[1];
  if FFormula1.PTGS[2] = 0 then begin
    pB := PByteArray(@FFormula1.PTGS[3]);
    i := 0;
    k := 0;
    while i < Sz do begin
      SetLength(Values[k],255);
      j := 1;
      while (i < Sz) and (pB[i] <> 0) do begin
        Values[k][j] := WideChar(pB[i]);
        Inc(i);
        Inc(j);
      end;
      SetLength(Values[k],j - 1);
      Inc(i);
      Inc(k);
    end;
  end
  else begin
    pW := PWordArray(@FFormula1.PTGS[3]);
    i := 0;
    k := 0;
    while i < Sz do begin
      SetLength(Values[k],255);
      j := 1;
      while (i < Sz) and (pW[i] <> 0) do begin
        Values[k][j] := WideChar(pW[i]);
        Inc(i);
        Inc(j);
      end;
      SetLength(Values[k],j - 1);
      Inc(i);
      Inc(k);
    end;
  end;
end;

function TDataValidation.GetListSize: integer;
var
  i,Sz: integer;
  pB: PByteArray;
  pW: PWordArray;
begin
  Result := 0;
  if FValidationType <> vtList then
    Exit;
  if (FFormula1.Size < 3) or (FFormula1.PTGS[0] <> xptgStr97) then
    Exit;
  Sz := FFormula1.PTGS[1];
  if FFormula1.PTGS[2] = 0 then begin
    pB := @FFormula1.PTGS[3];
    for i := 0 to Sz - 1 do begin
      if pB[i] = 0 then
        Inc(Result);
    end;
    Inc(Result);
  end
  else begin
    pW := PWordArray(@FFormula1.PTGS[3]);
    for i := 0 to Sz - 1 do begin
      if pW[i] = 0 then
        Inc(Result);
    end;
    Inc(Result);
  end;
end;

procedure TDataValidation.Include(Col1, Row1, Col2, Row2: integer);
begin
  FAreas.Include(Col1, Row1, Col2, Row2);
end;

function TDataValidation.Intersect(Col1, Row1, Col2, Row2: integer): boolean;
begin
  Result := FAreas.AreaInAreas(Col1, Row1, Col2, Row2);
end;

procedure TDataValidation.LoadFromStream(Stream: TXLSStream; PBuf: PByteArray);
var
  P: PByteArray;
  i,Len: integer;
  Header: TBIFFHeader;

function GetPtrText: AxUCString;
var
  Len: integer;
begin
  Len := PWord(P)^;
  P := PByteArray(NativeInt(P) + 2);
  if Len > 0 then begin
    Result := Trim(ByteStrToWideString(P,Len));
    if P[0] = 0 then
      P := PByteArray(NativeInt(P) + Len + 1)
    else if P[0] = 1 then
      P := PByteArray(NativeInt(P) + Len * 2 + 1)
    else
      raise XLSRWException.Create('Unhandled string in DV');
  end
  else
    Result := '';
end;

begin
  Stream.ReadHeader(Header);
  if Header.RecID <> BIFFRECID_DV then
    raise XLSRWException.Create('Excpected record missing: DV');
  Stream.Read(Pbuf^,Header.Length);
  P := PBuf;
  FOptions := PInteger(P)^;
  FValidationType := TValidationType(FOptions and $000F);
  FValidationStyle := TValidationStyle((FOptions shr 4) and $0007);
  FValidationOperator := TValidationOperator((FOptions shr 20) and $000F);
  P := PByteArray(NativeInt(P) + 4);
  FInputTitle := GetPtrText;
  FErrorTitle := GetPtrText;
  FInputMsg := GetPtrText;
  FErrorMsg := GetPtrText;

  Len := PWord(P)^;
  P := PByteArray(NativeInt(P) + 4);
  SetRawExpression1(P,Len);
  P := PByteArray(NativeInt(P) + Len);

  Len := PWord(P)^;
  P := PByteArray(NativeInt(P) + 4);
  SetRawExpression2(P,Len);
  P := PByteArray(NativeInt(P) + Len);

  Len := PWord(P)^;
  if Len > $0FFF then
    Len := 0;
  P := PByteArray(NativeInt(P) + 2);
  for i := 0 to Len - 1 do begin
    Areas.Add.SetSize(PWordArray(P)[2],PWordArray(P)[0],PWordArray(P)[3],PWordArray(P)[1]);
    P := PByteArray(NativeInt(P) + SizeOf(TRecCellArea));
  end;
end;

procedure TDataValidation.Move(Col1, Row1, Col2, Row2, DeltaCol, DeltaRow: integer);
begin
  FAreas.Move(Col1, Row1, Col2, Row2, DeltaCol, DeltaRow);
end;

procedure TDataValidation.Move(DeltaCol, DeltaRow: integer);
begin
  FAreas.Move(DeltaCol, DeltaRow);
end;

procedure TDataValidation.SaveToStream(Stream: TXLSStream; PBuf: PByteArray);
var
  i,L: integer;

function StrSize(S: AxUCString): integer;
begin
  if S = '' then
    Result := 4
  else
    Result := 2 + 1 + Length(S) * 2;
end;

procedure WriteString(S: AxUCString);
begin
  if S = '' then
    Stream.WLWord(1)
  else
    Stream.WriteWideString(S);
end;

begin
  L := Sizeof(FOptions) +
       StrSize(FInputTitle) +
       StrSize(FErrorTitle) +
       StrSize(FInputMsg) +
       StrSize(FErrorMsg) +
       FFormula1.Size + 2 + 2 +
       FFormula2.Size + 2 + 2 +
       FAreas.Count * 8 + 2;
  Stream.WriteHeader(BIFFRECID_DV,L);
  if FAssigned then begin
    FOptions := $0000;
    if FValidationType = vtList then begin
      if (FFormula1.Size > 0) and (FFormula1.PTGS[0] = xptgStr97) then
        FOptions := FOptions or $00000080
      else
        FOptions := FOptions and not $00000080;
    end;
    Inc(FOptions,Integer(FValidationType));
    Inc(FOptions,Integer(FValidationStyle) shl 4);
    if voAllowEmptyCells in FValidationOptions then
      Inc(FOptions,$00000100);
    if voSupressDropDown in FValidationOptions then
      Inc(FOptions,$00000200);
    if voShowPromptBox in FValidationOptions then
      Inc(FOptions,$00040000);
    if voShowErrorBox in FValidationOptions then
      Inc(FOptions,$00080000);
    Inc(FOptions,Integer(FValidationOperator) shl 20);
  end;
  Stream.Write(FOptions,4);
  WriteString(FInputTitle);
  WriteString(FErrorTitle);
  WriteString(FInputMsg);
  WriteString(FErrorMsg);

  Stream.WWord(FFormula1.Size);
  Stream.WWord($0000);
  if FFormula1.Size > 0 then
    Stream.Write(FFormula1.PTGS^,FFormula1.Size);

  Stream.WWord(FFormula2.Size);
  Stream.WWord($0000);
  if FFormula2.Size > 0 then
    Stream.Write(FFormula2.PTGS^,FFormula2.Size);

  Stream.WWord(Areas.Count);
  for i := 0 to Areas.Count - 1 do begin
    Stream.WWord(Areas[i].Row1);
    Stream.WWord(Areas[i].Row2);
    Stream.WWord(Areas[i].Col1);
    Stream.WWord(Areas[i].Col2);
  end;
end;

procedure TDataValidation.SetExpression1(const Value: AxUCString);
var
  P: PByteArray;
  Sz: integer;
begin
   GetMem(P,512);
   try
    Sz := TDataValidations(Collection).FFormulaHandler.EncodeFormula(Value,TokenClassRef,P,512);
    SetRawExpression1(P,Sz);
  finally
    FreeMem(P);
  end;
  FAssigned := True;
end;

procedure TDataValidation.SetExpression2(const Value: AxUCString);
var
  P: PbyteArray;
  Sz: integer;
begin
  GetMem(P,512);
  try
    Sz := TDataValidations(Collection).FFormulaHandler.EncodeFormula(Value,TokenClassRef,P,512);
    SetRawExpression2(P,Sz);
  finally
    FreeMem(P);
  end;
  FAssigned := True;
end;

procedure TDataValidation.SetList(Values: array of AxUCString);
var
  i,Sz: integer;
  pW: PWordArray;
begin
  Sz := 0;
  for i := 0 to High(Values) do
    Inc(Sz,Length(Values[i]));
  if Sz <= 0 then
    Exit;
  // Separator zero.
  Inc(Sz,Length(Values));
  if Sz > 255 then
    raise XLSRWException.Create('There can be max 255 characthers in a list');
  SetValidationType(vtList);
  // Can propably be any value;
  SetValidationOperator(vopBetween);
  ClearExpr1;
  // ptgStr + Len + Unicode + Size
  FFormula1.Size := 1 + 1 + 1 + Sz * 2;
  GetMem(FFormula1.PTGS,FFormula1.Size);
  FFormula1.PTGS[0] := xptgStr97;
  FFormula1.PTGS[1] := Sz;
  FFormula1.PTGS[2] := 1;
  pW := PWordArray(@FFormula1.PTGS[3]);
  for i := 0 to High(Values) do begin
    System.Move(Pointer(Values[i])^,pW^,Length(Values[i]) * 2);
    pW[Length(Values[i])] := 0;
    pW := PWordArray(NativeInt(pW) + Length(Values[i]) * 2 + 2);
  end;
end;

procedure TDataValidation.SetRawExpression1(Value: PByteArray; Len: integer);
begin
  ClearExpr1;
  FFormula1.Size := Len;
  GetMem(FFormula1.PTGS,FFormula1.Size);
  System.Move(Value^,FFormula1.PTGS^,FFormula1.Size);
end;

procedure TDataValidation.SetRawExpression2(Value: PByteArray; Len: integer);
begin
  ClearExpr2;
  FFormula2.Size := Len;
  GetMem(FFormula2.PTGS,FFormula2.Size);
  System.Move(Value^,FFormula2.PTGS^,FFormula2.Size);
end;

procedure TDataValidation.SetValidationOperator(const Value: TValidationOperator);
begin
  FValidationOperator := Value;
  FAssigned := True;
end;

procedure TDataValidation.SetValidationOptions(const Value: TValidationOptions);
begin
  FValidationOptions := Value;
  FAssigned := True;
end;

procedure TDataValidation.SetValidationStyle(const Value: TValidationStyle);
begin
  FValidationStyle := Value;
  FAssigned := True;
end;

procedure TDataValidation.SetValidationType(const Value: TValidationType);
begin
  FValidationType := Value;
  FAssigned := True;
end;

{ TDataValidations }

function TDataValidations.Add: TDataValidation;
begin
  Result := TDataValidation(inherited Add);
end;

function TDataValidations.ComboBoxObjId: integer;
begin
  if FRecDVAL.DropDownId = $FFFFFFFF then
    Result := -1
  else
    Result := FRecDVAL.DropDownId;
end;

constructor TDataValidations.Create(AOwner: TPersistent; FormulaHandler: TFormulaHandler);
begin
  inherited Create(TDataValidation);
  FOwner := AOwner;
  FFormulaHandler := FormulaHandler;
  FRecDVAL.DropDownId := $FFFFFFFF;
end;

function TDataValidations.GetDataValidation(Index: integer): TDataValidation;
begin
  Result := TDataValidation(inherited Items[Index]);
end;

function TDataValidations.GetOwner: TPersistent;
begin
  Result := FOwner;
end;

procedure TDataValidations.LoadFromStream(Stream: TXLSStream; PBuf: PByteArray);
var
  i: integer;
begin
  Move(PBuf^,FRecDVAL,SizeOf(TRecDVAL));
  for i := 0 to Integer(FRecDVAL.DVCount) - 1 do
    Add.LoadFromStream(Stream,PBuf);
end;

procedure TDataValidations.SaveToStream(Stream: TXLSStream; PBuf: PByteArray);
var
  i: integer;
begin
  if Count > 0 then begin
    Stream.WriteHeader(BIFFRECID_DVAL,SizeOf(TRecDVAL));
    FRecDVAL.DVCount := Count;
    Stream.Write(FRecDVAL,SizeOf(TRecDVAL));
    for i := 0 to Count - 1 do
      Items[i].SaveToStream(Stream,PBuf);
  end;
end;

end.
