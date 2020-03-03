unit BIFF_FormulaHandler5;

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
     BIFF_RecsII5, BIFF_Utils5, BIFF_DecodeFormula5, BIFF_EncodeFormulaII5, BIFF_Names5,
     BIFF_ExcelFuncII5, BIFF_WideStrList5,
     Xc12Utils5,
     XLSUtils5, XLSFormulaTypes5, XLSMask5;

type TSheetDataType = (sdtName,sdtCell);

type TSheetNameEvent = procedure (Name: AxUCString; var Index,TabCount: integer) of object;
type TSheetDataEvent = function (DataType: TSheetDataType; SheetIndex,Col,Row: integer): AxUCString of object;

type TFormulaHandler = class(TObject)
private
     FOwner: TPersistent;
     FSheetNameEvent: TSheetNameEvent;
     FSheetDataEvent: TSheetDataEvent;
     FFormulaEncoder: TEncodeFormula;
     FExternalNames: TExternalNames;
     FInternalNames: TInternalNames;

     procedure FormulaUnknownFunction(Name: AxUCString; var ID: integer);
     //procedure FormulaUnknownName(Name: AxUCString; var ID: integer);
     procedure FormulaExternName(Path,Filename,SheetName,Name: AxUCString; IsRef: boolean; var ExtIndex,NameIndex: integer);
     function  GetNameEvent: TUnknownNameEvent;
     procedure SetNameEvent(const Value: TUnknownNameEvent);
public
     constructor Create(AOwner: TPersistent);
     destructor Destroy; override;
     procedure Clear;
     function  DecodeFormula(Buf: PByteArray; Len: integer): AxUCString;
     function  DecodeFormulaRel(Buf: PByteArray; Len: integer; SheetIndex,Col,Row: integer): AxUCString;
     function  EncodeFormula(Formula: AxUCString; ReturnClass: byte; var Buf: PByteArray; BufSz: integer): integer;
     function  GetName(NameType: TFormulaNameType; SheetIndex,NameIndex,Col,Row: integer): AxUCString;
     procedure GetNameRef(AName: string; out ASheetIndex,Col1,Row1,Col2,Row2: integer; AcceptArea: boolean = False);
     procedure DeleteSheet(SheetIndex: integer);
     procedure InsertSheet(SheetIndex: integer);
     procedure DeleteRows(SheetIndex,Row1,Row2: integer);
     procedure InsertRows(SheetIndex,Row,Count: integer);
     procedure DeleteColumns(SheetIndex,Col1,Col2: integer);
     procedure InsertColumns(SheetIndex,Col,Count: integer);

     property InternalNames: TInternalNames read FInternalNames write FInternalNames;
     property ExternalNames: TExternalNames read FExternalNames write FExternalNames;

     property _OnName: TUnknownNameEvent read GetNameEvent write SetNameEvent;
     property OnSheetName: TSheetNameEvent read FSheetNameEvent write FSheetNameEvent;
     property OnSheetData: TSheetDataEvent read FSheetDataEvent write FSheetDataEvent;
     end;

implementation

uses BIFF5;

{ TFormulaHandler }

procedure TFormulaHandler.Clear;
begin
  FExternalNames.Clear;
  FInternalNames.Clear;
end;

constructor TFormulaHandler.Create(AOwner: TPersistent);
begin
  FOwner := AOwner;

  FFormulaEncoder := TEncodeFormula.Create;
  FFormulaEncoder.OnUnknownFunction := FormulaUnknownFunction;
  FFormulaEncoder.OnExternName := FormulaExternName;

  FExternalNames := TExternalNames.Create;
  FInternalNames := TInternalNames.Create(FOwner,GetName,FFormulaEncoder);
end;

function TFormulaHandler.DecodeFormula(Buf: PByteArray; Len: integer): AxUCString;
begin
  Result := DecodeFmla(Buf,Len,-1,-1,-1,GetName,FuncArgSeparator);
end;

function TFormulaHandler.DecodeFormulaRel(Buf: PByteArray; Len, SheetIndex, Col, Row: integer): AxUCString;
begin
  Result := DecodeFmla(Buf,Len,SheetIndex,Col,Row,GetName,FuncArgSeparator);
end;

procedure TFormulaHandler.DeleteColumns(SheetIndex, Col1, Col2: integer);
var
  i,j,C1,C2,Count: integer;
begin
  Count := Col2 - Col1 + 1;
  for i := 0 to FInternalNames.Count - 1 do begin
    j := FInternalNames[i].ExtSheet;
    if (j >= 0) and (FExternalNames.IsSelf(j) = SheetIndex) then begin
      C1 := FInternalNames[i].Col1;
      C2 := FInternalNames[i].Col2;

      if C1 > Col2 then begin
        Dec(C1,Count);
        Dec(C2,Count);
      end
      else if (C1 < Col1) and (C2 > Col2) then begin
        Dec(C2,Count);
      end
      else if (C1 <= Col2) and (C2 > Col2) then begin
        C1 := Col1;
        Dec(C2,Count);
      end
      else if (C1 >= Col1) and (C2 <= Col2) then begin
        C1 := -1;
        C2 := -1;
      end
      else if (C2 >= Col1) and (C2 <= Col2) then begin
        C2 := Col1 - 1;
      end
      else if C2 < Col1 then begin
        Continue;
      end;
      if (C1 >= 0) and (C1 <= MAXCOL) and (C2 >= 0) and (C2 <= MAXCOL) then begin
        FInternalNames[i].Col1 := C1;
        FInternalNames[i].Col2 := C2;
      end
      else
        FInternalNames[i].SetError;
    end;
  end;
end;

procedure TFormulaHandler.DeleteRows(SheetIndex, Row1, Row2: integer);
var
  i,j,R1,R2,Count: integer;
begin
  Count := Row2 - Row1 + 1;
  for i := 0 to FInternalNames.Count - 1 do begin
    j := FInternalNames[i].ExtSheet;
    if (j >= 0) and (FExternalNames.IsSelf(j) = SheetIndex) then begin
      R1 := FInternalNames[i].Row1;
      R2 := FInternalNames[i].Row2;

      if R1 > Row2 then begin
        Dec(R1,Count);
        Dec(R2,Count);
      end
      else if (R1 < Row1) and (R2 > Row2) then begin
        Dec(R2,Count);
      end
      else if (R1 <= Row2) and (R2 > Row2) then begin
        R1 := Row1;
        Dec(R2,Count);
      end
      else if (R1 >= Row1) and (R2 <= Row2) then begin
        R1 := -1;
        R2 := -1;
      end
      else if (R2 >= Row1) and (R2 <= Row2) then begin
        R2 := Row1 - 1;
      end
      else if R2 < Row1 then begin
        Continue;
      end;
      if (R1 >= 0) and (R1 <= MAXROW) and (R2 >= 0) and (R2 <= MAXROW) then begin
        FInternalNames[i].Row1 := R1;
        FInternalNames[i].Row2 := R2;
      end
      else
        FInternalNames[i].SetError;
    end;
  end;
end;

procedure TFormulaHandler.DeleteSheet(SheetIndex: integer);
begin
  FExternalNames.DeleteSheet(SheetIndex);
end;

destructor TFormulaHandler.Destroy;
begin
  FInternalNames.Free;
  FExternalNames.Free;
  FFormulaEncoder.Free;
  inherited;
end;

function TFormulaHandler.EncodeFormula(Formula: AxUCString; ReturnClass: byte; var Buf: PByteArray; BufSz: integer): integer;
begin
  Result := FFormulaEncoder.Encode(Formula,ReturnClass,Buf,BufSz,fvExcel97);
end;

procedure TFormulaHandler.FormulaExternName(Path,Filename,SheetName,Name: AxUCString; IsRef: boolean; var ExtIndex, NameIndex: integer);
var
  Index,TabCount: integer;
begin
  if (Path = '') and (Filename = '') then begin
    FSheetNameEvent(SheetName,Index,TabCount);
    if Index >= 0 then begin
      ExtIndex := FExternalNames.AddSelf(Index,TabCount);
      NameIndex := Index;
    end;
  end
  else begin
    if IsRef then begin
      FExternalNames.RefIndexByName(Path,Filename,SheetName,ExtIndex,NameIndex);
      if NameIndex < 0 then
        FExternalNames.AddRef(Path,Filename,SheetName,ExtIndex,NameIndex);
    end
    else
      FExternalNames.NameIndexByName(Path,Filename,Name,ExtIndex,NameIndex);
  end;
end;

procedure TFormulaHandler.FormulaUnknownFunction(Name: AxUCString; var ID: integer);
begin
  ID := -1;
end;

//procedure TFormulaHandler.FormulaUnknownName(Name: AxUCString; var ID: integer);
//begin
//  ID := FInternalNames.FindName(Name);
//end;

function TFormulaHandler.GetName(NameType: TFormulaNameType; SheetIndex, NameIndex, Col, Row: integer): AxUCString;
var
  i: integer;
begin
  case NameType of
    ntExternSheet: begin
      i := FExternalNames.IsSelf(SheetIndex);
      // The sheet is deleted
      if i = $FFFF then
        Result := Xc12CellErrorNames[errRef] + '!'
      else if i = $FFFE then
        Result := FInternalNames[NameIndex - 1].Name
      else if i >= 0 then
        Result := FSheetDataEvent(sdtName,i,-1,-1) + '!'
      else
        Result := FExternalNames.AsString[SheetIndex,NameIndex];
    end;
    ntCellValue: begin
      Result := FSheetDataEvent(sdtName,SheetIndex,Col,Row)
    end;
    else
      Result := FInternalNames[NameIndex - 1].Name;
  end;
end;

function TFormulaHandler.GetNameEvent: TUnknownNameEvent;
begin
  Result := FFormulaEncoder._OnUnknownName;
end;

procedure TFormulaHandler.GetNameRef(AName: string; out ASheetIndex, Col1,Row1,Col2,Row2: integer; AcceptArea: boolean = False);
var
  i: integer;
begin
  i := FInternalNames.FindName(AName);
  if i < 0 then
    raise XLSRWException.CreateFmt('Can not find name "%s"',[AName]);
  if (Length(FInternalNames[i].NameDef) = 11) and (FInternalNames[i].NameDef[0] in [xptgArea3d97,xptgArea3dA97,xptgArea3dV97]) then begin
    if not AcceptArea then
      raise XLSRWException.CreateFmt('Name "%s" is an area',[AName]);
    ASheetIndex := FExternalNames.IsSelf(PPTGArea3d8(@FInternalNames[i].NameDef[1]).Index);
    if ASheetIndex < 0 then
      raise XLSRWException.CreateFmt('Name "%s" is not an internal name',[AName]);
    Col1 := PPTGArea3d8(@FInternalNames[i].NameDef[1]).Col1;
    Row1 := PPTGArea3d8(@FInternalNames[i].NameDef[1]).Row1;
    Col2 := PPTGArea3d8(@FInternalNames[i].NameDef[1]).Col2;
    Row2 := PPTGArea3d8(@FInternalNames[i].NameDef[1]).Row2;
  end
  else if (Length(FInternalNames[i].NameDef) = 7) and (FInternalNames[i].NameDef[0] in [xptgRef3d97,xptgRef3dV97,xptgRef3dA97]) then begin
    with PPTGRef3d8(@FInternalNames[i].NameDef[1])^ do begin
      ASheetIndex := FExternalNames.IsSelf(Index);
      if ASheetIndex < 0 then
        raise XLSRWException.CreateFmt('Name "%s" is not an internal name',[AName]);
      Col1 := Col;
      Row1 := Row;
      Col2 := Col;
      Row2 := Row;
    end;
  end
  else
    raise XLSRWException.CreateFmt('Name "%s" is not valid',[AName]);
end;

procedure TFormulaHandler.InsertColumns(SheetIndex, Col, Count: integer);
var
  i,j,C1,C2: integer;
begin
  for i := 0 to FInternalNames.Count - 1 do begin
    j := FInternalNames[i].ExtSheet;
    if (j >= 0) and (FExternalNames.IsSelf(j) = SheetIndex) then begin
      C1 := FInternalNames[i].Col1;
      C2 := FInternalNames[i].Col2;

      if C1 >= Col then begin
        Inc(C1,Count);
        Inc(C2,Count);
      end
      else if C2 < Col then
        Continue
      else
        Inc(C2,Count);

      if (C1 > MAXCOL) and (C2 > MAXCOL) then begin
        FInternalNames[i].SetError;
        Continue;
      end;
      if C1 > MAXCOL then
        C1 := MAXCOL;
      if C2 > MAXCOL then
        C2 := MAXCOL;
      FInternalNames[i].Col1 := C1;
      FInternalNames[i].Col2 := C2;
    end;
  end;
end;

procedure TFormulaHandler.InsertRows(SheetIndex, Row, Count: integer);
var
  i,j,R1,R2: integer;
begin
  for i := 0 to FInternalNames.Count - 1 do begin
    j := FInternalNames[i].ExtSheet;
    if (j >= 0) and (FExternalNames.IsSelf(j) = SheetIndex) then begin
      R1 := FInternalNames[i].Row1;
      R2 := FInternalNames[i].Row2;

      if R1 >= Row then begin
        Inc(R1,Count);
        Inc(R2,Count);
      end
      else if R2 < Row then
        Continue
      else
        Inc(R2,Count);

      if (R1 > MAXROW) and (R2 > MAXROW) then begin
        FInternalNames[i].SetError;
        Continue;
      end;
      if R1 > MAXROW then
        R1 := MAXROW;
      if R2 > MAXROW then
        R2 := MAXROW;
      FInternalNames[i].Row1 := R1;
      FInternalNames[i].Row2 := R2;
    end;
  end;
end;

procedure TFormulaHandler.InsertSheet(SheetIndex: integer);
begin
  FExternalNames.InsertSheet(SheetIndex);
end;

procedure TFormulaHandler.SetNameEvent(const Value: TUnknownNameEvent);
begin
  FFormulaEncoder._OnUnknownName := Value;
end;

end.
