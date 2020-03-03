unit XBookOptions2;

{-
********************************************************************************
******* XLSSpreadSheet V3.00                                             *******
*******                                                                  *******
******* Copyright(C) 2006,2017 Lars Arvidsson, Axolot Data               *******
*******                                                                  *******
******* email: components@axolot.com                                     *******
******* URL:   http://www.axolot.com                                     *******
********************************************************************************
** Users of the XLSSpreadSheet component must accept the following            **
** disclaimer of warranty:                                                    **
**                                                                            **
** XLSSpreadSheet is supplied as is. The author disclaims all warranties,     **
** expressed or implied, including, without limitation, the warranties of     **
** merchantability and of fitness for any purpose. The author assumes no      **
** liability for damages, direct or consequential, which may result from the  **
** use of XLSSpreadSheet.                                                     **
********************************************************************************
}

{$B-}
{$H+}
{$R-}
{$I AxCompilers.inc}

interface

uses Classes, SysUtils,
	   XLSUtils5, XLSReadWriteII5;

type TXLSBookOptionsView = class(TPersistent)
private
     FXLS: TXLSReadWriteII5;
     FRowsColumns: boolean;
     FGridlines: boolean;
     FHorizontalScroll: boolean;
     FVerticalScroll: boolean;

     procedure SetRowsColumns(const Value: boolean);
     procedure SetGridlines(const Value: boolean);
     procedure SetHorizontalScroll(const Value: boolean);
     procedure SetVerticalScroll(const Value: boolean);
     function GetGridlines: boolean;
     function GetHorizontalScroll: boolean;
     function GetRowsColumns: boolean;
     function GetVerticalScroll: boolean;
public
     constructor Create(XLS: TXLSReadWriteII5);
     destructor Destroy; override;
published
     property RowsColumns: boolean read GetRowsColumns write SetRowsColumns;
     property Gridlines: boolean read GetGridlines write SetGridlines;
     property HorizontalScroll: boolean read GetHorizontalScroll write SetHorizontalScroll;
     property VerticalScroll: boolean read GetVerticalScroll write SetVerticalScroll;
     end;

type TXLSBookOptionsGeneral = class(TPersistent)
private
     FXLS: TXLSReadWriteII5;

     procedure SetFontName(const Value: AxUCString);
     procedure SetFontSize(const Value: double);
     function  GetFontName: AxUCString;
     function GetFontSize: double;
public
     constructor Create(AXLS: TXLSReadWriteII5);
published
     property FontName: AxUCString read GetFontName write SetFontName;
     property FontSize: double read GetFontSize write SetFontSize;
     end;

type TXLSBookOptions = class(TPersistent)
private
     FGeneral: TXLSBookOptionsGeneral;
     FView: TXLSBookOptionsView;
     FReadOnly: boolean;
public
     constructor Create(XLS: TXLSReadWriteII5);
     destructor Destroy; override;
published
     property General: TXLSBookOptionsGeneral read FGeneral write FGeneral;
     property View: TXLSBookOptionsView read FView write FView;
     property ReadOnly: boolean read FReadOnly write FReadOnly;
     end;

implementation

{ TXLSBookOptions }

constructor TXLSBookOptions.Create(XLS: TXLSReadWriteII5);
begin
  FGeneral := TXLSBookOptionsGeneral.Create(XLS);
  FView := TXLSBookOptionsView.Create(XLS);
end;

destructor TXLSBookOptions.Destroy;
begin
  FGeneral.Free;
  FView.Free;
  inherited;
end;

{ TXLSBookOptionsView }

constructor TXLSBookOptionsView.Create(XLS: TXLSReadWriteII5);
begin
  FXLS := XLS;
  FRowsColumns := True;
  FGridlines := True;
  FHorizontalScroll := True;
  FVerticalScroll := True;
end;

destructor TXLSBookOptionsView.Destroy;
begin

  inherited;
end;

function TXLSBookOptionsView.GetGridlines: boolean;
begin
  Result := FGridlines;
end;

function TXLSBookOptionsView.GetHorizontalScroll: boolean;
begin
  Result := FHorizontalScroll;
end;

function TXLSBookOptionsView.GetRowsColumns: boolean;
begin
  Result := FRowsColumns;
end;

function TXLSBookOptionsView.GetVerticalScroll: boolean;
begin
  Result := FVerticalScroll;
end;

procedure TXLSBookOptionsView.SetGridlines(const Value: boolean);
begin
  FGridlines := Value;
end;

procedure TXLSBookOptionsView.SetHorizontalScroll(const Value: boolean);
begin
  FHorizontalScroll := Value;
end;

procedure TXLSBookOptionsView.SetRowsColumns(const Value: boolean);
begin
  FRowsColumns := Value;
end;
                       
procedure TXLSBookOptionsView.SetVerticalScroll(const Value: boolean);
begin
  FVerticalScroll := Value;
end;

{ TXLSBookOptionsGeneral }

constructor TXLSBookOptionsGeneral.Create(AXLS: TXLSReadWriteII5);
begin
  FXLS := AXLS;
end;

function TXLSBookOptionsGeneral.GetFontName: AxUCString;
begin
  Result := FXLS.Font.Name;
end;

function TXLSBookOptionsGeneral.GetFontSize: double;
begin
  Result := FXLS.Font.Size;
end;

procedure TXLSBookOptionsGeneral.SetFontName(const Value: AxUCString);
begin
  FXLS.Font.Name := Value;
end;

procedure TXLSBookOptionsGeneral.SetFontSize(const Value: double);
begin
  if (Value < 1) or (Value > 72) then
    Exit;
  FXLS.Font.Size := Value;
end;

end.
