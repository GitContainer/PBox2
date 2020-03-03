unit XLSNames5;

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
     Xc12Utils5, Xc12DataWorkbook5, Xc12Manager5,
     BIFF_Names5,
     XLSUtils5, XLSFormulaTypes5, XLSEncodeFmla5, XLSFormula5, XLSClassFactory5;

type TXLSNames = class;

     TXLSName = class(TXc12DefinedName)
private
     procedure SetName(const Value: AxUCString);
     function  GetArea: PXLS3dCellArea;
protected
     FParent: TXLSNames;

     procedure ClearDefinition;
     procedure Compile;

     function  GetContent: AxUCString; override;
     function  GetDefinition: AxUCString;
     // Do not overload SetContent of base class as SetDefinition also calls
     // Compile, and SetContent can't do that as SetContent is called when
     // file is read. At that time formulas can't be compiled until all data
     // is read.
     procedure SetDefinition(const Value: AxUCString);

     // Hide props from TXc12DefinedName
     property Ptgs;
     property PtgsSz;
     property Content;
     property OrigContent;
     property Deleted;
     property Unused97;
public
     constructor Create(AParent: TXLSNames);

     procedure Update;

     //* Returns True it the names area is on ASheetIndex and within ACol and ARow.
     function  Hit(const ASheetIndex, ACol, ARow: integer): boolean;

     property Name: AxUCString read FName write SetName;
     property Definition: AxUCString read GetDefinition write SetDefinition;
     property Area: PXLS3dCellArea read GetArea;
     end;

     TXLSNames = class(TXc12DefinedNames)
private
     function GetItems(Index: integer): TXLSName;
protected
     FManager: TXc12Manager;
     FFormulas: TXLSFormulaHandler;

     function  DoAdd(const AName,ADefinition: AxUCString; const ALocalSheetId: integer): TXLSName;
public
     constructor Create(AClassFactory: TXLSClassFactory; AManager: TXc12Manager; AFormulas: TXLSFormulaHandler);

     //* Copy the names that belongs to ASheetId to a Strings object.
     //* If ASetObject is True, the Objects property of AList is assigned with
     //* the relevant TXLSName object.
     procedure ToStrings(const ASheetId: integer; AList: TStrings; const ASetObject: boolean = False);
     procedure DumpNames(AList: TStrings);
     //* Returns a name that is on ASheetIndex and has an area is within ACol and ARow. If not found,
     //* Nil is returned.
     function  Hit(const ASheetIndex,ACol,ARow: integer): TXLSName;

     function  Add(const AName,ASheetName: AxUCString; const ACol,ARow: integer): TXLSName; overload;
     function  Add(const AName,ASheetName: AxUCString; const ACol1,ARow1,ACol2,ARow2: integer): TXLSName; overload;
     function  Add(const AName,ADefinition: AxUCString; const ALocalSheetId: integer = -1): TXLSName; overload;
     // Creates a local name on Sheet ASheetName.
     function  Add(const AName,ADefinition,ASheetName: AxUCString): TXLSName; overload;
     function  Find(const AName: AxUCString): TXLSName;

     property Items[Index: integer]: TXLSName read GetItems; default;
     end;

     TXLSNames_Int = class(TXLSNames)
protected
     procedure AdjustSheetInsert(const ASheetIndex, ACount: integer);
     procedure AdjustSheetDelete(const ASheetIndex, ACount: integer);
public
     procedure AfterRead;
     procedure AdjustSheetIndex(const ASheetIndex, ACount: integer);

     property HashValid: boolean read FHashValid write FHashValid;
     end;

implementation

{ TXLSNames }

function TXLSNames.Add(const AName,ADefinition: AxUCString; const ALocalSheetId: integer = -1): TXLSName;
begin
  Result := DoAdd(AName,ADefinition,ALocalSheetId);
end;

function TXLSNames.Add(const AName, ADefinition, ASheetName: AxUCString): TXLSName;
var
  i: integer;
begin
  i := FManager.Worksheets.Find(ASheetName);
  if i < 0 then begin
    FManager.Errors.Error(ASheetName,XLSERR_FMLA_UNKNOWNSHEET);
    Result := Nil;
  end
  else
    Result := DoAdd(AName,ADefinition,i);
end;

function TXLSNames.Add(const AName, ASheetName: AxUCString; const ACol, ARow: integer): TXLSName;
var
  Def: AxUCString;
begin
  if FManager.Worksheets.Find(ASheetName) < 0 then begin
    FManager.Errors.Error(ASheetName,XLSERR_FMLA_UNKNOWNSHEET);
    Result := Nil;
  end
  else begin
    Def := ASheetName + '!' + ColRowToRefStr(ACol,ARow,True,True);
    Result := Add(AName,Def);
  end;
end;

function TXLSNames.Add(const AName, ASheetName: AxUCString; const ACol1, ARow1, ACol2, ARow2: integer): TXLSName;
var
  Def: AxUCString;
begin
  if FManager.Worksheets.Find(ASheetName) < 0 then begin
    FManager.Errors.Error(ASheetName,XLSERR_FMLA_UNKNOWNSHEET);
    Result := Nil;
  end
  else begin
    Def := ASheetName + '!' + AreaToRefStr(ACol1,ARow1,ACol2,ARow2,True,True,True,True);
    Result := Add(AName,Def);
  end;
end;

constructor TXLSNames.Create(AClassFactory: TXLSClassFactory; AManager: TXc12Manager; AFormulas: TXLSFormulaHandler);
begin
  inherited Create(AClassFactory);
  FManager := AManager;
  FFormulas := AFormulas;
end;

function TXLSNames.DoAdd(const AName, ADefinition: AxUCString; const ALocalSheetId: integer): TXLSName;
begin
  Result := Nil;
  if Trim(AName) = '' then
    FManager.Errors.Error('',XLSERR_NAME_EMPTY)
  else begin
    if (ALocalSheetId >= 0) and (FindIdInSheet(AName,ALocalSheetId) >= 0) then
      FManager.Errors.Error(AName,XLSERR_NAME_DUPLICATE)
    else if Find(AName) <> Nil then
      FManager.Errors.Error(AName,XLSERR_NAME_DUPLICATE)
    else begin
      Result := TXLSName(inherited Add(AName,ALocalSheetId));
      Result.Definition := ADefinition;

      if FManager.Version = xvExcel97 then
{$ifdef XLS_BIFF}
        FManager.Names97.Add(AName,ADefinition);
{$endif}
//      Result.Compile;
    end;
  end;
end;

procedure TXLSNames.DumpNames(AList: TStrings);
var
  i: integer;
  S: AxUCString;
begin
  for i := 0 to Count - 1 do begin
    S := '"' + Items[i].Name + '","' + Items[i].Definition + '",';
    if (Items[i].LocalSheetId >= 0) then
      S := '"' + FManager.Worksheets[Items[i].LocalSheetId].Name + '"'
    else
      S := S + '"Workbook"';

    AList.Add(S);
  end;
end;

function TXLSNames.Find(const AName: AxUCString): TXLSName;
var
  i: integer;
begin
  i := FindId(AName,-1);
  if (i >= 0) and not Items[i].Deleted then
    Result := Items[i]
  else
    Result := Nil;
end;

function TXLSNames.GetItems(Index: integer): TXLSName;
begin
  Result := TXLSName(inherited Items[Index]);
end;

function TXLSNames.Hit(const ASheetIndex,ACol, ARow: integer): TXLSName;
var
  i: integer;
begin
  for i := 0 to Count - 1 do begin
    if Items[i].Hit(ASheetIndex,ACol,ARow) then begin
      Result := Items[i];
      Exit;
    end;
  end;
  Result := Nil;
end;

procedure TXLSNames.ToStrings(const ASheetId: integer; AList: TStrings; const ASetObject: boolean);
var
  i: integer;
  S: AxUCString;
begin
  for i := 0 to Count - 1 do begin
    S := Items[i].Name;
    if (Items[i].LocalSheetId < 0) or (Items[i].LocalSheetId = ASheetId) then begin
      if ASetObject then
        AList.AddObject(S,Items[i])
      else
        AList.Add(S);
    end;
  end;
end;

{ TXLSName }

procedure TXLSName.ClearDefinition;
begin
  if FPtgs <> Nil then
    FreeMem(FPtgs);
  FPtgs := Nil;
  FPtgsSz := 0;
  FSimpleName := xsntNone;
end;

procedure TXLSName.Compile;
begin
  if FPtgs = Nil then begin
    FPtgsSz := FParent.FFormulas.EncodeFormula(FContent,FPtgs,-1);
    if FPtgsSz = 0 then begin
      FParent.FManager.Errors.Error('',XLSERR_FMLA_BADSHEETNAME);
      Exit;
    end;
  end;
  FSimpleName := IsSimpleName(FParent.FManager,FPtgs,FPtgsSz,@FArea);
end;

constructor TXLSName.Create(AParent: TXLSNames);
begin
  FParent := AParent;
end;

function TXLSName.GetArea: PXLS3dCellArea;
begin
  if FSimpleName > xsntNone then
    Result := @FArea
  else
    Result := Nil;
end;

function TXLSName.GetContent: AxUCString;
begin
  case FSimpleName of
    xsntNone : begin
      Result := FParent.FFormulas.DecodeFormula(Nil,-1,FPtgs,FPtgsSz,True);
      Exit;
    end;
    xsntRef  : Result := FParent.FManager.Worksheets[FArea.SheetIndex].QuotedName + '!' + ColRowToRefStrEnc(FArea.Col1,FArea.Row1);
    xsntArea : Result := FParent.FManager.Worksheets[FArea.SheetIndex].QuotedName + '!' + AreaToRefStrEnc(FArea.Col1,FArea.Row1,FArea.Col2,FArea.Row2);
    xsntError: Result := Xc12CellErrorNames[errRef];
  end;
end;

function TXLSName.GetDefinition: AxUCString;
begin
  case FSimpleName of
    xsntNone : begin
      if FPtgsSz > 0 then
        Result := FParent.FFormulas.DecodeFormula(Nil,-1,FPtgs,FPtgsSz,False)
      else
        Result := '';
      Exit;
    end;
    xsntRef  : Result := FParent.FManager.Worksheets[FArea.SheetIndex].QuotedName + '!' + ColRowToRefStrEnc(FArea.Col1,FArea.Row1);
    xsntArea : Result := FParent.FManager.Worksheets[FArea.SheetIndex].QuotedName + '!' + AreaToRefStrEnc(FArea.Col1,FArea.Row1,FArea.Col2,FArea.Row2);
    xsntError: Result := Xc12CellErrorNames[errRef];
  end;
end;

function TXLSName.Hit(const ASheetIndex, ACol, ARow: integer): boolean;
begin
  case SimpleName of
    xsntNone : Result := False;
    xsntRef  : Result := (SimpleArea.SheetIndex = ASheetIndex) and
                        ((SimpleArea.Col1 and not COL_ABSFLAG) = ACol) and
                        ((SimpleArea.Row1 and not ROW_ABSFLAG) = ARow);
    xsntArea : Result := (SimpleArea.SheetIndex = ASheetIndex) and
                        ((SimpleArea.Col1 and not COL_ABSFLAG) >= ACol) and
                        ((SimpleArea.Row1 and not ROW_ABSFLAG) >= ARow) and
                        ((SimpleArea.Col2 and not COL_ABSFLAG) <= ACol) and
                        ((SimpleArea.Row2 and not ROW_ABSFLAG) <= ARow);
    xsntError: Result := False;
    else       Result := False;
  end;
end;

procedure TXLSName.SetDefinition(const Value: AxUCString);
begin
  ClearDefinition;

  FContent := Value;

  Compile;
end;

procedure TXLSName.SetName(const Value: AxUCString);
begin
  if FParent.Find(Value) <> Nil then
    FParent.FManager.Errors.Error(Value,XLSERR_NAME_DUPLICATE)
  else begin
    FName := Value;
    TXLSNames_Int(FParent).HashValid := False;
  end;
end;

procedure TXLSName.Update;
begin
  FSimpleName := IsSimpleName(FParent.FManager,FPtgs,FPtgsSz,@FArea);
end;

{ TXLSNames_Int }

procedure TXLSNames_Int.AdjustSheetDelete(const ASheetIndex, ACount: integer);
var
  i: integer;
  N: TXLSName;
begin
  for i := 0 to Count - 1 do begin
    N := Items[i];
    if (N.LocalSheetId >= ASheetIndex) and (N.LocalSheetId <= (ASheetIndex + ACount - 1)) then
      N.Deleted := True
    else begin
      case N.SimpleName of
        xsntNone: begin
          FFormulas.AdjustSheetIndex(N.Ptgs,N.PtgsSz,ASheetIndex,-ACount);
          N.LocalSheetId := N.LocalSheetId - ACount;
        end;
        xsntRef,
        xsntArea: begin
          if N.Area.SheetIndex >= ASheetIndex then begin
            if (N.Area.SheetIndex) < (ASheetIndex + ACount) then begin
              N.Area.SheetIndex := -1;
              N.FSimpleName := xsntError;
            end
            else
              N.Area.SheetIndex := N.Area.SheetIndex - ACount;
          end;
        end;
        xsntError: ;
      end;
    end;
  end;
end;

procedure TXLSNames_Int.AdjustSheetIndex(const ASheetIndex, ACount: integer);
begin
  if ACount < 0 then
    AdjustSheetDelete(ASheetIndex,-ACount)
  else
    AdjustSheetInsert(ASheetIndex,ACount);
end;

procedure TXLSNames_Int.AdjustSheetInsert(const ASheetIndex, ACount: integer);
var
  i,j: integer;
  N: TXLSName;
begin
  for i := 0 to Count - 1 do begin
    N := Items[i];
    case N.SimpleName of
      xsntNone: begin
          FFormulas.AdjustSheetIndex(N.Ptgs,N.PtgsSz,ASheetIndex,ACount);
          N.LocalSheetId := N.LocalSheetId + ACount;
        end;
      xsntRef,
      xsntArea: begin
        if N.Area.SheetIndex >= ASheetIndex then begin
          j := N.Area.SheetIndex + ACount;
          if (j >= 0) and (j < XLS_MAXSHEETS) then
            N.Area.SheetIndex := N.Area.SheetIndex + ACount
          else
            N.Area.SheetIndex := -1;
        end;
      end;
      xsntError: ;
    end;
  end;
end;

procedure TXLSNames_Int.AfterRead;
var
  i: integer;
begin
  for i := 0 to Count - 1 do
    Items[i].Compile;
end;

end.
