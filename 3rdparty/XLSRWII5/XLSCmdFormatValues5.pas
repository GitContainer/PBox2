unit XLSCmdFormatValues5;

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
     Xc12Utils5, XLSUtils5;

type TXLSCmdFormatValueType = (xcfvtNone,xcfvtBoolean,xcfvtInteger,xcfvtFloat,xcfvtString);

type TXLSCmdFormatCommand = (
xcfcCellColor,
xcfcBorderBottomColor,
xcfcFontSize,
xcfcFontBold,
xcfcFontItalic,
xcfcIndent);

type TXLSCmdFormatValue = class(TObject)
private
     function  GetAsBoolean: boolean;
     function  GetAsFloat: double;
     function  GetAsInteger: integer;
     function  GetAsString: AxUCString;
     procedure SetAsBoolean(const Value: boolean);
     procedure SetAsFloat(const Value: double);
     procedure SetAsInteger(const Value: integer);
     procedure SetAsString(const Value: AxUCString);
protected
     FCommand: TXLSCmdFormatCommand;
public
     function  Type_: TXLSCmdFormatValueType; virtual; abstract;

     property AsBoolean: boolean read GetAsBoolean write SetAsBoolean;
     property AsInteger: integer read GetAsInteger write SetAsInteger;
     property AsFloat: double read GetAsFloat write SetAsFloat;
     property AsString: AxUCString read GetAsString write SetAsString;

     property Command: TXLSCmdFormatCommand read FCommand write FCommand;
     end;

type TXLSCmdFormatValueNone = class(TXLSCmdFormatValue)
protected
public
     function  Type_: TXLSCmdFormatValueType; override;
     end;

type TXLSCmdFormatValueBoolean = class(TXLSCmdFormatValue)
protected
     FValue: boolean;
public
     function  Type_: TXLSCmdFormatValueType; override;

     property Value: boolean read FValue write FValue;
     end;

type TXLSCmdFormatValueInteger = class(TXLSCmdFormatValue)
protected
     FValue: integer;
public
     function  Type_: TXLSCmdFormatValueType; override;

     property Value: integer read FValue write FValue;
     end;

type TXLSCmdFormatValueFloat = class(TXLSCmdFormatValue)
protected
     FValue: double;
public
     function  Type_: TXLSCmdFormatValueType; override;

     property Value: double read FValue write FValue;
     end;

type TXLSCmdFormatValueString = class(TXLSCmdFormatValue)
protected
     FValue: AxUCString;
public
     function  Type_: TXLSCmdFormatValueType; override;

     property Value: AxUCString read FValue write FValue;
     end;

type TXLSCmdFormatValues = class(TObjectList)
private
     function GetItems(Index: integer): TXLSCmdFormatValue;
protected
public
     constructor Create;

     procedure Add(ACommand: TXLSCmdFormatCommand); overload;
     procedure Add(ACommand: TXLSCmdFormatCommand; AValue: boolean); overload;
     procedure Add(ACommand: TXLSCmdFormatCommand; AValue: integer); overload;
     procedure Add(ACommand: TXLSCmdFormatCommand; AValue: double); overload;
     procedure Add(ACommand: TXLSCmdFormatCommand; AValue: AxUCString); overload;

     property Items[Index: integer]: TXLSCmdFormatValue read GetItems; default;
     end;

implementation

{ TXLSCmdFormatValue }

function TXLSCmdFormatValue.GetAsBoolean: boolean;
begin
  if Type_ <> xcfvtBoolean then
    raise XLSRWException.Create('Wrong value type');

  Result := TXLSCmdFormatValueBoolean(Self).Value;
end;

function TXLSCmdFormatValue.GetAsFloat: double;
begin
  if Type_ <> xcfvtFloat then
    raise XLSRWException.Create('Wrong value type');

  Result := TXLSCmdFormatValueFloat(Self).Value;
end;

function TXLSCmdFormatValue.GetAsInteger: integer;
begin
  if Type_ <> xcfvtInteger then
    raise XLSRWException.Create('Wrong value type');

  Result := TXLSCmdFormatValueInteger(Self).Value;
end;

function TXLSCmdFormatValue.GetAsString: AxUCString;
begin
  if Type_ <> xcfvtString then
    raise XLSRWException.Create('Wrong value type');

  Result := TXLSCmdFormatValueString(Self).Value;
end;

procedure TXLSCmdFormatValue.SetAsBoolean(const Value: boolean);
begin
  if Type_ <> xcfvtBoolean then
    raise XLSRWException.Create('Wrong value type');

  TXLSCmdFormatValueBoolean(Self).Value := Value;
end;

procedure TXLSCmdFormatValue.SetAsFloat(const Value: double);
begin
  if Type_ <> xcfvtFloat then
    raise XLSRWException.Create('Wrong value type');

  TXLSCmdFormatValueFloat(Self).Value := Value;
end;

procedure TXLSCmdFormatValue.SetAsInteger(const Value: integer);
begin
  if Type_ <> xcfvtInteger then
    raise XLSRWException.Create('Wrong value type');

  TXLSCmdFormatValueInteger(Self).Value := Value;
end;

procedure TXLSCmdFormatValue.SetAsString(const Value: AxUCString);
begin
  if Type_ <> xcfvtString then
    raise XLSRWException.Create('Wrong value type');

  TXLSCmdFormatValueString(Self).Value := Value;
end;

{ TXLSCmdFormatValues }

procedure TXLSCmdFormatValues.Add(ACommand: TXLSCmdFormatCommand; AValue: integer);
var
  Item: TXLSCmdFormatValueInteger;
begin
  Item := TXLSCmdFormatValueInteger.Create;
  Item.Command := ACommand;
  Item.Value := AValue;

  inherited Add(Item);
end;

procedure TXLSCmdFormatValues.Add(ACommand: TXLSCmdFormatCommand; AValue: boolean);
var
  Item: TXLSCmdFormatValueBoolean;
begin
  Item := TXLSCmdFormatValueBoolean.Create;
  Item.Command := ACommand;
  Item.Value := AValue;

  inherited Add(Item);
end;

procedure TXLSCmdFormatValues.Add(ACommand: TXLSCmdFormatCommand);
var
  Item: TXLSCmdFormatValueNone;
begin
  Item := TXLSCmdFormatValueNone.Create;
  Item.Command := ACommand;

  inherited Add(Item);
end;

procedure TXLSCmdFormatValues.Add(ACommand: TXLSCmdFormatCommand; AValue: double);
var
  Item: TXLSCmdFormatValueFloat;
begin
  Item := TXLSCmdFormatValueFloat.Create;
  Item.Command := ACommand;
  Item.Value := AValue;

  inherited Add(Item);
end;

procedure TXLSCmdFormatValues.Add(ACommand: TXLSCmdFormatCommand; AValue: AxUCString);
var
  Item: TXLSCmdFormatValueString;
begin
  Item := TXLSCmdFormatValueString.Create;
  Item.Command := ACommand;
  Item.Value := AValue;

  inherited Add(Item);
end;

constructor TXLSCmdFormatValues.Create;
begin
  inherited Create;
end;

function TXLSCmdFormatValues.GetItems(Index: integer): TXLSCmdFormatValue;
begin
  Result := TXLSCmdFormatValue(inherited Items[Index]);
end;

{ TXLSCmdFormatValueNone }

function TXLSCmdFormatValueNone.Type_: TXLSCmdFormatValueType;
begin
  Result := xcfvtNone;
end;

{ TXLSCmdFormatValueBoolean }

function TXLSCmdFormatValueBoolean.Type_: TXLSCmdFormatValueType;
begin
  Result := xcfvtBoolean;
end;

{ TXLSCmdFormatValueInteger }

function TXLSCmdFormatValueInteger.Type_: TXLSCmdFormatValueType;
begin
  Result := xcfvtInteger;
end;

{ TXLSCmdFormatValueFloat }

function TXLSCmdFormatValueFloat.Type_: TXLSCmdFormatValueType;
begin
  Result := xcfvtFloat;
end;

{ TXLSCmdFormatValueString }

function TXLSCmdFormatValueString.Type_: TXLSCmdFormatValueType;
begin
  Result := xcfvtString;
end;

{ TXLSCmdFormatValueArea }

end.
