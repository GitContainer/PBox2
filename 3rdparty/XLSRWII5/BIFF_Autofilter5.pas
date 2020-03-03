unit BIFF_Autofilter5;

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
     BIFF_Utils5, BIFF_RecsII5, BIFF_Stream5, BIFF_Names5, BIFF_ControlObj5,
     Xc12Utils5,
     XLSUtils5;

const AUTOFILTER_UNUSED      = $00;
const AUTOFILTER_RK          = $02;
const AUTOFILTER_FLOAT       = $04;
const AUTOFILTER_STRING      = $06;
const AUTOFILTER_BOOLERR     = $08;
const AUTOFILTER_ALLBLANK    = $0C;
const AUTOFILTER_ALLNONBLANK = $0E;

//* What type of cell values the autofilter is filtering on.
type TAutoFilterDataType = (afdtUnused,      //* Filter data is unused.
                            afdtFloat,       //* Floating pint number.
                            afdtString,      //* String.
                            afdtBoolean,     //* Boolean.
                            afdtError,       //* Error value.
                            afdtAllBlanks,   //* Blank cells (cells without values).
                            afdtAllNonBlanks //* Non blank cells (cells with values).
                            );

//* What kind of comparision operator to use for the condition.
type TAutoFilterOperator = (afoUnassigned,  //* Operator is unasigned/unused
                            afoLess,        //* Less; <
                            afoEqual,       //* Equal; =
                            afoLessEqual,   //* Less or equal; <=
                            afoGreater,     //* Greater; >
                            afoNotEqual,    //* Not equal; <>
                            afoGreaterEqual //* Greater or equal; >=
                            );

//* Boolean filter operator for filters with two conditions.
type TAutoFilterType = (aftAndFilter, //* AND filter (both conditions must be met).
                        aftOrFilter   //* OR filter (one of the conditions must be met).
                        );

//* ~exclude
type TAutoFilterCondData = record
     CompOperator: TAutoFilterOperator;
     vString: AxUCString;
     case DataType: TAutoFilterDataType of
       afdtFloat:   (vFloat: double);
       afdtBoolean: (vBoolean: boolean);
       afdtError:   (vError: TXc12CellError);
     end;

//* An autofilter condition. A ~[link TAutofilter] has one or two conditions.
type TAutoFilterCondition = class(TObject)
private
     FData: TAutoFilterCondData;

     function  GetBooleanValue: boolean;
     function  GetCompOperator: TAutoFilterOperator;
     function  GetErrorValue: TXc12CellError;
     function  GetFloatValue: double;
     function  GetStringValue: AxUCString;
     procedure SetBooleanValue(const Value: boolean);
     procedure SetCompOperator(const Value: TAutoFilterOperator);
     procedure SetErrorValue(const Value: TXc12CellError);
     procedure SetFloatValue(const Value: double);
     procedure SetStringValue(const Value: AxUCString);
     function  GetBlankValue: boolean;
     procedure SetBlankValue(const Value: boolean);
     function  GetNonBlankValue: boolean;
     procedure SetNonBlankValue(const Value: boolean);
protected
     function  SetData(P,PStr: PByteArray): integer;
     procedure GetData(P,PStr: PByteArray);
     function  GetStringSize: integer;
public
     //* ~exclude
     function FilterFloatValue(Value: double): boolean;
     //* ~exclude
     function FilterStringValue(Value: AxUCString): boolean;
     //* ~exclude
     function FilterBooleanValue(Value: boolean): boolean;
     //* ~exclude
     function FilterErrorValue(Value: TXc12CellError): boolean;
     //* ~exclude
     function FilterBlankValue: boolean;

     //* Filtering comparision operator.
     property CompOperator: TAutoFilterOperator read GetCompOperator write SetCompOperator;
     //* Gets or sets the comparision value as a floationg point value.
     //* When reading; if the value not is a floating point type, zero is returned.
     property FloatValue: double read GetFloatValue write SetFloatValue;
     //* Gets or sets the comparision value as a string value.
     //* When reading; if the value not is a string type, an empty string is returned.
     property StringValue: AxUCString read GetStringValue write SetStringValue;
     //* Gets or sets the comparision value as a boolean value.
     //* When reading; if the value not is a boolean type, FALSE is returned.
     property BooleanValue: boolean read GetBooleanValue write SetBooleanValue;
     //* Gets or sets the comparision value as an error value.
     //* When reading; if the value not is an error type, errError is returned.
     property ErrorValue: TXc12CellError read GetErrorValue write SetErrorValue;
     //* Gets or sets the comparision value as a blank cell value.
     //* When reading; if the value not is a blank cell, False is returned.
     property BlankValue: boolean read GetBlankValue write SetBlankValue;
     //* Gets or sets the comparision value as a non blank cell value.
     //* When reading; if the value not is a non blank cell, False is returned.
     property NonBlankValue: boolean read GetNonBlankValue write SetNonBlankValue;
     //* ~exclude
     property ConditionData: TAutoFilterCondData read FData write FData;
     end;

type TAutoFilters = class;

//* An autofilter column condition. This is equal to one autofilter combo box
//* in Excel. The TAutoFilter objects are created with
//* ~[link TAutofilters.Add] or ~[link TAutofilters.QuickFilter]
//* See ~[link TAutoFilters] for an example on how to use autofilters.
     TAutoFilter = class(TObject)
private
     FParent: TAutoFilters;
     FId: longword;
     FIndex: word;
     FOptions: word;
     FFilterType: TAutoFilterType;
     FCondition1,FCondition2: TAutoFilterCondition;

     procedure SetFilterType(const Value: TAutoFilterType);
public
     //* ~exclude
     constructor Create(Parent: TAutoFilters);
     //* ~exclude
     destructor Destroy; override;
     //* ~exclude
     procedure LoadFromStream(Stream: TXLSStream; PBuf: PByteArray);
     //* ~exclude
     procedure SaveToStream(Stream: TXLSStream; PBuf: PByteArray);
     //* Clears the TAutofilter and sets all values to their default.
     procedure Clear;
     //* Checks if a floating point value meets the filter condition. ~[p]
     //* Mainly used by ~[link TAutofilters.Execute]
     //* ~param Value The value that the filter shall be tested against.
     //* ~result True it the value meets the filter conditions.
     function FilterFloatValue(Value: double): boolean;
     //* Checks if a string value meets the filter condition. ~[p]
     //* Mainly used by ~[link TAutofilters.Execute]
     //* ~param Value The value that the filter shall be tested against.
     //* ~result True it the value meets the filter conditions.
     function FilterStringValue(Value: AxUCString): boolean;
     //* Checks if a boolean value meets the filter condition. ~[p]
     //* Mainly used by ~[link TAutofilters.Execute]
     //* ~param Value The value that the filter shall be tested against.
     //* ~result True it the value meets the filter conditions.
     function FilterBooleanValue(Value: boolean): boolean;
     //* Checks if an error value meets the filter condition. ~[p]
     //* Mainly used by ~[link TAutofilters.Execute]
     //* ~param Value The value that the filter shall be tested against.
     //* ~result True it the value meets the filter conditions.
     function FilterErrorValue(Value: TXc12CellError): boolean;
     //* Checks if a blank value (a empty cell) meets the filter condition. ~[p]
     //* Mainly used by ~[link TAutofilters.Execute]
     //* ~result True it the value meets the filter conditions.
     function FilterBlankValue: boolean;

     //* Boolean condition operator when two conditions are used.
     property FilterType: TAutoFilterType read FFilterType write SetFilterType;
     //* First filtering condition.
     property Condition1: TAutoFilterCondition read FCondition1;
     //* Optional second filtering condition.
     property Condition2: TAutoFilterCondition read FCondition2;
     end;

//* Implements the Excel Autofilter command. Just as in Excel, there can only
//* be one autofilter on each worksheet. The autofilter can have multiple
//* conditions, one for each column in the filter area.
//* ~[br]
//* ~[br]
//* ~[b Example:]
//* ~[br]
//* ~[sample
//* var
//*   i: integer;
//* begin
//*   // Add some data
//*   for i := Low(LongMonthNames) to High(LongMonthNames) do
//*     XLS.Sheet[0~[].AsString[3,9 + i - Low(LongMonthNames)~[] := LongMonthNames[i~[];
//*   for i := 9 to 20 do
//*     XLS.Sheet[0~[].AsFloat[4,i~[] := Random(2000);
//*   // Create an autofilter on two columns
//*   XLS.Sheet[0~[].Autofilters.Add(3,5,4,21);
//*   // A string filter with one condition
//*   XLS.Sheet[0~[].Autofilters[0~[].Condition1.CompOperator := afoNotEqual;
//*   XLS.Sheet[0~[].Autofilters[0~[].Condition1.StringValue := 'april';
//*   // A numeric filter with two conditions
//*   XLS.Sheet[0~[].Autofilters[1~[].FilterType := aftAndFilter;
//*   XLS.Sheet[0~[].Autofilters[1~[].Condition1.CompOperator := afoGreaterEqual;
//*   XLS.Sheet[0~[].Autofilters[1~[].Condition1.FloatValue := 100;
//*   XLS.Sheet[0~[].Autofilters[1~[].Condition2.CompOperator := afoLessEqual;
//*   XLS.Sheet[0~[].Autofilters[1~[].Condition2.FloatValue := 900;
//*   // Execute the filter. Without executing the filter, the data on the
//*   // worksheet will be unfiltered.
//*   XLS.Sheet[0~[].Autofilters.Execute;
//* end;]
     TAutoFilters = class(TObjectList)
private
     FParent: TObject;
     FDropDownCount: integer;
     FLoadFromFile: boolean;
     FComboBoxes: TControlComboBoxes;
     FInternalNames: TInternalNames;
     FInternalName: TInternalName;
     FIsDynamicFilter: boolean;

     function GetCol1: integer;
     function GetCol2: integer;
     function GetRow1: integer;
     function GetRow2: integer;
     function GetItems(Index: integer): TAutoFilter;
protected
     procedure DeleteAll;
     procedure AddNil;
public
     //* ~exclude
     constructor Create(Parent: TObject; ComboBoxes: TControlComboBoxes; Names: TInternalNames);
     //* ~exclude
     destructor Destroy; override;
     //* ~exclude
     procedure LoadFromStream(Stream: TXLSStream; PBuf: PByteArray);
     //* ~exclude
     procedure SaveToStream(Stream: TXLSStream; PBuf: PByteArray);
     //* Adds a new set of autofilters. Any existing autofilters will be deleted.
     //* There will be as many autofilters as there are columns. If C1 is 3 and
     //* C2 is 6, there will be four autofilters.
     //* ~param C1 The first column for the filter area.
     //* ~param R1 The first row for the filter area.
     //* ~param C2 The last column for the filter area.
     //* ~param R2 The last row for the filter area.
     //* ~result The created ~[link TAutofilter]. If more than one filter is added,
     //* the last of the filters is returned. That is, if the filters are between
     //* columns 3 and 6, the one for column 6 is returned.
     //* See also ~[link .QuickFilter].
     function  Add(C1,R1,C2,R2: integer): TAutoFilter;
     //* ~exclude
     procedure IndexFilterVal(Index: integer; Op: TAutoFilterOperator; Value: AxUCString);
     //* ~exclude
     procedure IndexClear(Index: integer);
     //* Not used. Individual autofilters can't be deleted.
     //* ~param Index Index
     procedure Delete(Index: Integer);
     //* Deletes all autofilters.
     procedure Clear; override;
     //* Returns True if there are any active filter.
     function IsFiltered: boolean;
     //* QuickFilter is a fast way to create an autofilter.
     //* ~param C Column for the upper right corner of the filter area.
     //* ~param R1 The first row of the filter area.
     //* ~param R2 Last row of the filter area.
     //* ~param Conditions Array of strings that each is a filter condtion.
     //* There will be as many filters created as there are conditions.
     //* The format of the condoition strings are [Operator][Value].
     //* Example of condition strings: ['>=100','=target text','<>TRUE']
     //* Valid operators are: =, >=, <=, >, <, <> ~[p]
     //* QuickFilter will also call ~[link Execute]
     procedure QuickFilter(C,R1,R2: integer; Conditions: array of AxUCString);
     //* Returns True if the cell at C and R is within the autofilter area.
     //* ~param C Column for hit test.
     //* ~param R Row for the hit test.
     //* ~result True if C and R are within the autofilter area.
     function  CellHit(C,R: integer): boolean;
     //* The items in the list. If a filter at Index not is active
     // (the column is unfiltered), Items returns Nil.
     property Items[Index: integer]: TAutoFilter read GetItems; default;
     //* First column of the filter area.
     property Col1: integer read GetCol1;
     //* First row of the filter area.
     property Row1: integer read GetRow1;
     //* Last column of the filter area.
     property Col2: integer read GetCol2;
     //* Last row of the filter area.
     property Row2: integer read GetRow2;
     //* There is no defined filter area. The filter range starts on the
     //* Autofilter combo box row + 1, and stops on the first empty cell.
     property IsDynamicFilter: boolean read FIsDynamicFilter; // TODO not fully implemented.
     end;

//~endparse~

implementation

uses BIFF_SheetData5;

{ TAutoFilter }

procedure TAutoFilter.Clear;
begin
  FCondition1.SetStringValue('');
  FCondition1.CompOperator := afoUnassigned;
  FCondition2.SetStringValue('');
  FCondition2.CompOperator := afoUnassigned;
end;

constructor TAutoFilter.Create(Parent: TAutoFilters);
begin
  FParent := Parent;
  FIndex := 0;
  FOptions := $0000;
  FCondition1 := TAutoFilterCondition.Create;
  FCondition2 := TAutoFilterCondition.Create;
end;

destructor TAutoFilter.Destroy;
begin
  FCondition1.Free;
  FCondition2.Free;
  inherited;
end;

function TAutoFilter.FilterBlankValue: boolean;
begin
  Result := False;
  if FCondition1.CompOperator > afoUnassigned then begin
    Result := FCondition1.FilterBlankValue;
    if FCondition2.CompOperator > afoUnassigned then begin
      case FFilterType of
        aftAndFilter: Result := Result and FCondition2.FilterBlankValue;
        aftOrFilter:  Result := Result or FCondition2.FilterBlankValue;
      end;
    end;
  end;
end;

function TAutoFilter.FilterBooleanValue(Value: boolean): boolean;
begin
  Result := False;
  if FCondition1.CompOperator > afoUnassigned then begin
    Result := FCondition1.FilterBooleanValue(Value);
    if FCondition2.CompOperator > afoUnassigned then begin
      case FFilterType of
        aftAndFilter: Result := Result and FCondition2.FilterBooleanValue(Value);
        aftOrFilter:  Result := Result or FCondition2.FilterBooleanValue(Value);
      end;
    end;
  end;
end;

function TAutoFilter.FilterErrorValue(Value: TXc12CellError): boolean;
begin
  Result := False;
  if FCondition1.CompOperator > afoUnassigned then begin
    Result := FCondition1.FilterErrorValue(Value);
    if FCondition2.CompOperator > afoUnassigned then begin
      case FFilterType of
        aftAndFilter: Result := Result and FCondition2.FilterErrorValue(Value);
        aftOrFilter:  Result := Result or FCondition2.FilterErrorValue(Value);
      end;
    end;
  end;
end;

function TAutoFilter.FilterFloatValue(Value: double): boolean;
begin
  Result := False;
  if FCondition1.CompOperator > afoUnassigned then begin
    Result := FCondition1.FilterFloatValue(Value);
    if FCondition2.CompOperator > afoUnassigned then begin
      case FFilterType of
        aftAndFilter: Result := Result and FCondition2.FilterFloatValue(Value);
        aftOrFilter:  Result := Result or FCondition2.FilterFloatValue(Value);
      end;
    end;
  end;
end;

function TAutoFilter.FilterStringValue(Value: AxUCString): boolean;
begin
  Result := False;
  if FCondition1.CompOperator > afoUnassigned then begin
    Result := FCondition1.FilterStringValue(Value);
    if FCondition2.CompOperator > afoUnassigned then begin
      case FFilterType of
        aftAndFilter: Result := Result and FCondition2.FilterStringValue(Value);
        aftOrFilter:  Result := Result or FCondition2.FilterStringValue(Value);
      end;
    end;
  end;
end;

procedure TAutoFilter.LoadFromStream(Stream: TXLSStream; PBuf: PByteArray);
var
  StrSz1: integer;
  Header: TBIFFHeader;
begin
  Stream.ReadHeader(Header);
  Stream.Read(PBuf^,Header.Length);
  FIndex := PWordArray(PBuf)[0];
  FOptions := PWordArray(PBuf)[1];
  StrSz1 := FCondition1.SetData(PByteArray(NativeInt(PBuf) + 4),PByteArray(NativeInt(PBuf) + 4 + 10 + 10));
  FCondition2.SetData(PByteArray(NativeInt(PBuf) + 4 + 10),PByteArray(NativeInt(PBuf) + 4 + 10 + 10 + StrSz1));
end;

procedure TAutoFilter.SaveToStream(Stream: TXLSStream; PBuf: PByteArray);
var
  StrSz1,StrSz2: integer;
begin
{
  if not Assigned then
    Exit;
}
  StrSz1 := FCondition1.GetStringSize;
  StrSz2 := FCondition2.GetStringSize;
  Stream.WriteHeader(BIFFRECID_AUTOFILTER,4 + 10 + 10 + StrSz1 + StrSz2);
  Stream.WWord(FIndex);
  Stream.WWord(FOptions);
  FCondition1.GetData(PByteArray(NativeInt(PBuf)),PByteArray(NativeInt(PBuf) + 10 + 10));
  FCondition2.GetData(PByteArray(NativeInt(PBuf) + 10),PByteArray(NativeInt(PBuf) + 10 + 10 + StrSz1));
  Stream.Write(PBuf^,10 + 10 + StrSz1 + StrSz2);
end;

procedure TAutoFilter.SetFilterType(const Value: TAutoFilterType);
begin
  FFilterType := Value;
end;

{ TAutoFilterCondition }

function TAutoFilterCondition.FilterBlankValue: boolean;
begin
  case FData.DataType of
    afdtAllBlanks:    Result := True;
    afdtAllNonBlanks: Result := False;
    else              Result := False;
  end;
end;

function TAutoFilterCondition.FilterBooleanValue(Value: boolean): boolean;
begin
  case FData.DataType of
    afdtBoolean:      Result := True;
    afdtAllBlanks:    Result := False;
    afdtAllNonBlanks: Result := True;
    else              Result := False;
  end;
  if not Result then
    Exit;
  case FData.CompOperator of
    afoUnassigned:    Result := False;
    afoLess:          Result := Value < FData.vBoolean;
    afoEqual:         Result := Value = FData.vBoolean;
    afoLessEqual:     Result := Value <= FData.vBoolean;
    afoGreater:       Result := Value > FData.vBoolean;
    afoNotEqual:      Result := Value <> FData.vBoolean;
    afoGreaterEqual:  Result := Value >= FData.vBoolean;
  end;
end;

function TAutoFilterCondition.FilterErrorValue(Value: TXc12CellError): boolean;
begin
  case FData.DataType of
    afdtError:        Result := True;
    afdtAllBlanks:    Result := False;
    afdtAllNonBlanks: Result := True;
    else              Result := False;
  end;
  if not Result then
    Exit;
  case FData.CompOperator of
    afoUnassigned:    Result := False;
    afoLess:          Result := Value < FData.vError;
    afoEqual:         Result := Value = FData.vError;
    afoLessEqual:     Result := Value <= FData.vError;
    afoGreater:       Result := Value > FData.vError;
    afoNotEqual:      Result := Value <> FData.vError;
    afoGreaterEqual:  Result := Value >= FData.vError;
  end;
end;

function TAutoFilterCondition.FilterFloatValue(Value: double): boolean;
begin
  case FData.DataType of
    afdtFloat:        Result := True;
    afdtAllBlanks:    Result := False;
    afdtAllNonBlanks: Result := True;
    else              Result := False;
  end;
  if not Result then
    Exit;
  case FData.CompOperator of
    afoUnassigned:    Result := False;
    afoLess:          Result := Value < FData.vFloat;
    afoEqual:         Result := Value = FData.vFloat;
    afoLessEqual:     Result := Value <= FData.vFloat;
    afoGreater:       Result := Value > FData.vFloat;
    afoNotEqual:      Result := Value <> FData.vFloat;
    afoGreaterEqual:  Result := Value >= FData.vFloat;
  end;
end;

function TAutoFilterCondition.FilterStringValue(Value: AxUCString): boolean;
var
  V: integer;
begin
  case FData.DataType of
    afdtString:       Result := True;
    afdtAllBlanks: begin
      Result := Value = '';
      Exit;
    end;
    afdtAllNonBlanks: begin
      Result := Value <> '';
      Exit;
    end;
    else              Result := False;
  end;
  if not Result then
    Exit;
  V := CompareStr(MyWideUppercase(Value),MyWideUppercase(FData.vString));
  case FData.CompOperator of
    afoUnassigned:    Result := False;
    afoLess:          Result := V < 0;
    afoEqual:         Result := V = 0;
    afoLessEqual:     Result := V <= 0;
    afoGreater:       Result := V > 0;
    afoNotEqual:      Result := V <> 0;
    afoGreaterEqual:  Result := V >= 0;
  end;
end;

function TAutoFilterCondition.GetBlankValue: boolean;
begin
  if FData.DataType = afdtAllBlanks then
    Result := True
  else
    Result := False;
end;

function TAutoFilterCondition.GetBooleanValue: boolean;
begin
  if FData.DataType = afdtBoolean then
    Result := FData.vBoolean
  else
    Result := False;
end;

function TAutoFilterCondition.GetCompOperator: TAutoFilterOperator;
begin
  Result := FData.CompOperator;
end;

procedure TAutoFilterCondition.GetData(P, PStr: PByteArray);
var
  DataType: byte;
begin
  case FData.DataType of
    afdtFloat:        DataType := $04;
    afdtString:       DataType := $06;
    afdtBoolean,
    afdtError:        DataType := $08;
    afdtAllBlanks:    DataType := $0C;
    afdtAllNonBlanks: DataType := $0E;
    else              DataType := $00;
  end;
  P[0] := DataType;
  P[1] := Byte(FData.CompOperator);
  case DataType of
    AUTOFILTER_FLOAT: begin
      PDouble(NativeInt(P) + 2)^ := FData.vFloat;
    end;
    AUTOFILTER_STRING: begin
      P[6] := Length(FData.vString);
      WideStringToByteStr(FData.vString,PStr);
    end;
    AUTOFILTER_BOOLERR: begin
      if FData.DataType = afdtBoolean then begin
        P[2] := 0;
        P[3] := Byte(FData.vBoolean);
      end
      else begin
        P[2] := 1;
        P[3] := Byte(FData.vError);
      end;
    end;
  end;
end;

function TAutoFilterCondition.GetErrorValue: TXc12CellError;
begin
  if FData.DataType = afdtError then
    Result := FData.vError
  else
    Result := errUnknown;
end;

function TAutoFilterCondition.GetFloatValue: double;
begin
  if FData.DataType = afdtFloat then
    Result := FData.vFloat
  else
    Result := 0;
end;

function TAutoFilterCondition.GetNonBlankValue: boolean;
begin
  if FData.DataType = afdtAllNonBlanks then
    Result := True
  else
    Result := False;
end;

function TAutoFilterCondition.GetStringSize: integer;
begin
  if FData.DataType = afdtString then
    Result := Length(FData.vString) * 2 + 1
  else
    Result := 0;
end;

function TAutoFilterCondition.GetStringValue: AxUCString;
begin
  if FData.DataType = afdtString then
    Result := FData.vString
  else
    Result := '';
end;

procedure TAutoFilterCondition.SetBlankValue(const Value: boolean);
begin
  FData.DataType := afdtAllBlanks;
end;

procedure TAutoFilterCondition.SetBooleanValue(const Value: boolean);
begin
  FData.DataType := afdtBoolean;
  FData.vBoolean := Value;
end;

procedure TAutoFilterCondition.SetCompOperator(const Value: TAutoFilterOperator);
begin
  FData.CompOperator := Value;
end;

function TAutoFilterCondition.SetData(P,PStr: PByteArray): integer;
var
  DataType: byte;
begin
  Result := 0;
  DataType := P[0];
  FData.CompOperator := TAutoFilterOperator(P[1]);
  case DataType of
    AUTOFILTER_UNUSED: ;
    AUTOFILTER_RK: begin
      FData.DataType := afdtFloat;
      FData.vFloat := DecodeRK(PLongWordArray(NativeInt(P) + 2)[0]);
    end;
    AUTOFILTER_FLOAT: begin
      FData.DataType := afdtFloat;
      FData.vFloat := PDouble(NativeInt(P) + 2)^;
    end;
    AUTOFILTER_STRING: begin
      Result := P[6];
      FData.DataType := afdtString;
      FData.vString := ByteStrToWideString(PStr,Result);
    end;
    AUTOFILTER_BOOLERR: begin
      if P[2] = 0 then begin
        FData.DataType := afdtBoolean;
        FData.vBoolean := Boolean(P[3]);
      end
      else begin
        FData.DataType := afdtError;
        FData.vError := TXc12CellError(P[3]);
      end;
    end;
    AUTOFILTER_ALLBLANK: ;
    AUTOFILTER_ALLNONBLANK: ;
  end;
end;

procedure TAutoFilterCondition.SetErrorValue(const Value: TXc12CellError);
begin
  FData.DataType := afdtError;
  FData.vError := Value;
end;

procedure TAutoFilterCondition.SetFloatValue(const Value: double);
begin
  FData.DataType := afdtFloat;
  FData.vFloat := Value;
end;

procedure TAutoFilterCondition.SetNonBlankValue(const Value: boolean);
begin
  FData.DataType := afdtAllNonBlanks;
end;

procedure TAutoFilterCondition.SetStringValue(const Value: AxUCString);
begin
  FData.DataType := afdtString;
  FData.vString := Copy(Value,1,255);
end;

{ TAutoFilters }

function TAutoFilters.Add(C1,R1,C2,R2: integer): TAutoFilter;
var
  C: integer;
begin
  if not ValidArea(C1,R1,C2,R2) then
    raise XLSRWException.Create('Invalid cell area');
  Clear;
  FInternalName := FInternalNames.AddBuiltInName(xbnFilterDatabase,True,True,TSheet(FParent).Name,TSheet(FParent).Index,C1,R1,C2,R2);
  Result := Nil;
  for C := C1 to C2 do begin
    Result := TAutofilter.Create(Self);
    inherited Add(Result);
    with FComboBoxes.AddAutoFilter do begin
      Result.FId := Id;
      Col1 := C;
      Row1 := R1;
      Col2 := C + 1;
      Row2 := R1 + 1;
    end;
  end;
end;

procedure TAutoFilters.AddNil;
begin
  inherited Add(Nil);
end;

function TAutoFilters.CellHit(C, R: integer): boolean;
begin
  Result := (C >= GetCol1) and (C <= GetCol2) and (R >= GetRow1) and (R <= GetRow2); 
end;

procedure TAutoFilters.Clear;
var
  i: integer;
begin
  for i := 0 to Count - 1 do begin
    if Items[i] <> Nil then
      FComboBoxes.DeleteById(Items[i].FId);
  end;
  DeleteAll;
  FInternalNames.DeleteBuiltInName(xbnFilterDatabase,TSheet(FParent).Index + 1);
  FInternalName := Nil;
  inherited Clear;
end;

constructor TAutoFilters.Create(Parent: TObject; ComboBoxes: TControlComboBoxes; Names: TInternalNames);
begin
  inherited Create;
  OwnsObjects := False;
  FParent := Parent;
  FComboBoxes := ComboBoxes;
  FInternalNames := Names;
end;

procedure TAutoFilters.Delete(Index: Integer);
begin
  // Do nothing. Individual autofilters can't be deleted.
end;

procedure TAutoFilters.DeleteAll;
var
  i: integer;
begin
  for i := 0 to Count - 1 do begin
    if Items[i] <> Nil then
      Items[i].Free;
  end;
end;

destructor TAutoFilters.Destroy;
begin
  inherited;
end;

procedure TAutoFilters.IndexClear(Index: integer);
begin
  if Items[Index] <> Nil then begin
    Items[Index].Free;
    inherited Items[Index] := Nil;
  end;
end;

procedure TAutoFilters.IndexFilterVal(Index: integer; Op: TAutoFilterOperator;  Value: AxUCString);
var
  V: extended;
  S: string;
  Filter: TAutoFilter;
begin
  Filter := Items[Index];
  if Filter = Nil then begin
    Filter := TAutofilter.Create(Self);
    inherited Items[Index] := Filter;
  end
  else
    Filter.Clear;
  if Value = '' then begin
    Filter.Condition1.CompOperator := afoEqual;
    if OP = afoEqual then
      Filter.Condition1.BlankValue := True
    else
      Filter.Condition1.NonBlankValue := True;
  end
  else begin
    Filter.Condition1.CompOperator := Op;
    S := Value;
    if TextToFloat(PChar(S), V, fvExtended) then
      Filter.Condition1.FloatValue := V
    else
      Filter.Condition1.StringValue := Value;
  end;
end;

function TAutoFilters.GetCol1: integer;
begin
  Result := FInternalName.Col1;
end;

function TAutoFilters.GetCol2: integer;
begin
  Result := FInternalName.Col2;
end;

function TAutoFilters.GetItems(Index: integer): TAutoFilter;
begin
  Result := TAutoFilter(inherited Items[Index]);
end;

function TAutoFilters.GetRow1: integer;
begin
  Result := FInternalName.Row1;
end;

function TAutoFilters.GetRow2: integer;
begin
  Result := FInternalName.Row2;
end;

function TAutoFilters.IsFiltered: boolean;
var
  i: integer;
begin
  Result := False;
  for i := 0 to Count - 1 do begin
    Result := Items[i] <> Nil;
    if Result then
      Exit;
  end;
end;

procedure TAutoFilters.LoadFromStream(Stream: TXLSStream; PBuf: PByteArray);
var
  Header: TBIFFHeader;
  Filter: TAutoFilter;
begin
  if FInternalName = Nil then
    FInternalName := FInternalNames.FindBuiltInName(xbnFilterDatabase,TSheet(FParent).Index);
  FLoadFromFile := True;
  FIsDynamicFilter := False;
  // If next record is AUTOFILTERINFO, the current is FILTERMODE, and hence
  // there is an active autofilter.
  if Stream.PeekHeader = BIFFRECID_AUTOFILTERINFO then begin
    Stream.ReadHeader(Header);
    Stream.Read(PBuf^,Header.Length);
  end;
  FDropDownCount := PWordArray(PBuf)[0];
  while Stream.PeekHeader = BIFFRECID_AUTOFILTER do begin
    Filter := TAutofilter.Create(Self);
    Filter.LoadFromStream(Stream,PBuf);
    while Count < Filter.FIndex do
      AddNil;
    inherited Add(Filter);
  end;
  FLoadFromFile := False;
end;

procedure TAutoFilters.QuickFilter(C, R1, R2: integer; Conditions: array of AxUCString);
var
  i: integer;
  TempDS: WideChar;
  Values: array of TAutoFilterCondData;

procedure GetValue(Index: integer);
var
  S: AxUCString;
begin
  S := Copy(Conditions[Index],1,1);
  if S = '' then
    raise XLSRWException.Create('Missing condition');
  if S = '<' then begin
    if Copy(Conditions[Index],2,1) = '=' then begin
      Values[Index].CompOperator := afoLessEqual;
      S := Trim(Copy(Conditions[Index],3,MAXINT));
    end
    else if Copy(Conditions[Index],2,1) = '>' then begin
      Values[Index].CompOperator := afoNotEqual;
      S := Trim(Copy(Conditions[Index],3,MAXINT));
    end
    else begin
      Values[Index].CompOperator := afoLess;
      S := Trim(Copy(Conditions[Index],2,MAXINT));
    end;
  end
  else if S = '=' then begin
    Values[Index].CompOperator := afoEqual;
    S := Trim(Copy(Conditions[Index],2,MAXINT));
  end
  else if S = '>' then begin
    if Copy(Conditions[Index],2,1) = '=' then begin
      Values[Index].CompOperator := afoGreaterEqual;
      S := Trim(Copy(Conditions[Index],3,MAXINT));
    end
    else begin
      Values[Index].CompOperator := afoGreater;
      S := Trim(Copy(Conditions[Index],2,MAXINT));
    end;
  end;
  if S = '' then
    raise XLSRWException.Create('Missing value');
  if CharInSet(S[1],['0'..'9']) then begin
    Values[Index].vFloat := StrToFloat(S);
    Values[Index].DataType := afdtFloat;
  end
  else if (Uppercase(S) = 'TRUE') or (Uppercase(S) = 'FALSE') then begin
    if Uppercase(S) = 'TRUE' then
      Values[Index].vBoolean := True
    else
      Values[Index].vBoolean := False;
    Values[Index].DataType := afdtBoolean;
  end
  else begin
    Values[Index].vString := S;
    Values[Index].DataType := afdtString;
  end;
end;

begin
  if not ValidArea(C,R1,C + High(Conditions),R2) then
    raise XLSRWException.Create('Invalid cell area');
  SetLength(Values,Length(Conditions));
  TempDS := FormatSettings.DecimalSeparator;
  FormatSettings.DecimalSeparator := '.';
  try
    for i := 0 to High(Values) do
      GetValue(i);
  finally
    FormatSettings.DecimalSeparator := TempDS;
  end;
  Add(C,R1,C + High(Values),R2);
  for i := 0 to High(Values) do begin
    with Items[i] do
      Condition1.ConditionData := Values[i];
  end;
end;

procedure TAutoFilters.SaveToStream(Stream: TXLSStream; PBuf: PByteArray);
var
  i: integer;
begin
  if Count > 0 then begin
    Stream.WriteHeader(BIFFRECID_FILTERMODE,0);
    Stream.WriteWord(BIFFRECID_AUTOFILTERINFO,Count);
  end
  else begin
    // Files with "empty" filter ranges may have no filters but visible
    // drop down boxes.
    // Writing empty filters causes Excel2007 to crash on such files.
//    Stream.WriteWord(BIFFRECID_AUTOFILTERINFO,FDropDownCount);
  end;
  for i := 0 to Count - 1 do begin
    if Items[i] <> Nil then begin
      Items[i].FIndex := i;
      Items[i].SaveToStream(Stream,PBuf);
    end;
  end;
end;

end.
