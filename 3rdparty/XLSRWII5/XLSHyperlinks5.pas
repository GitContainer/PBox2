unit XLSHyperlinks5;

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
     Xc12Utils5, Xc12DataWorksheet5, Xc12Manager5,
     XLSUtils5, XLSMoveCopy5, XLSFormulaTypes5, XLSFormula5, XLSClassfactory5;

type TXLSHyperlinks = class;

     TXLSHyperlink = class(TXc12Hyperlink)
private
     function  GetAddress: AxUCString;
     procedure SetAddress(const Value: AxUCString);
     function  GetDescription: AxUCString;
     function  GetToolTip: AxUCString;
     procedure SetDescription(const Value: AxUCString);
     procedure SetToolTip(const Value: AxUCString);
protected
     FOwner: TXLSHyperlinks;

     function  EncodeFileHLink(ALink: AxUCString): AxUCString;
     function  GetCol1: integer; override;
     function  GetCol2: integer; override;
     function  GetRow1: integer; override;
     function  GetRow2: integer; override;
     procedure SetCol1(const Value: integer); override;
     procedure SetCol2(const Value: integer); override;
     procedure SetRow1(const Value: integer); override;
     procedure SetRow2(const Value: integer); override;
     procedure Parse; override;
public
     destructor Destroy; override;

     function  SheetIndex: integer; {$ifdef DELPHI_2007_OR_LATER} deprecated 'Do not use anymore. Take sheet index from the sheet that owns the hyperlink'; {$endif}

     //* If link type is xhltWorkbook, the link target column and row is returned in ACol and ARow.
     //* Returns True if the HyperlinkType is xhltWorkbook.
     function  GetWorkbookTarget(out ACol,ARow: integer): boolean;

     property Owner: TXLSHyperlinks read FOwner write FOwner;

     property Address: AxUCString read GetAddress write SetAddress;
     property Col1: integer read GetCol1 write SetCol1;
     property Row1: integer read GetRow1 write SetRow1;
     property Col2: integer read GetCol2 write SetCol2;
     property Row2: integer read GetRow2 write SetRow2;

     //* Text that is shown instead of the real address.
     property Description: AxUCString read GetDescription write SetDescription;
     //* Optional tooltip (hint) that is shown when the user holds the mouse
     //* pointer above the text.
     property ToolTip: AxUCString read GetToolTip write SetToolTip;
     end;

     TXLSHyperlinks = class(TXc12Hyperlinks)
private
     function GetItems(Index: integer): TXLSHyperlink;
protected
     FManager: TXc12Manager;
public
     constructor Create(AClassFactory: TXLSClassFactory; AManager: TXc12Manager);

     destructor Destroy; override;

     function  Add: TXLSHyperlink;
     function  Find(ACol,ARow: integer): TXLSHyperlink;

     property Items[Index: integer]: TXLSHyperlink read GetItems; default;
     end;

implementation

{ TXLSHyperlinks }

function TXLSHyperlinks.Add: TXLSHyperlink;
begin
  Result := TXLSHyperlink(inherited Add);
end;

constructor TXLSHyperlinks.Create(AClassFactory: TXLSClassFactory; AManager: TXc12Manager);
begin
  inherited Create(AClassFactory);

  FManager := AManager;
end;

destructor TXLSHyperlinks.Destroy;
begin

  inherited;
end;

function TXLSHyperlinks.Find(ACol, ARow: integer): TXLSHyperlink;
var
  i: integer;
begin
  i := CellInAreas(ACol,ARow);
  if i >= 0 then
    Result := GetItems(i)
  else
    Result := Nil;
end;

function TXLSHyperlinks.GetItems(Index: integer): TXLSHyperlink;
begin
  Result := TXLSHyperlink(inherited Items[Index]);
end;

{ TXLSHyperlink }

destructor TXLSHyperlink.Destroy;
begin

  inherited;
end;

function TXLSHyperlink.EncodeFileHLink(ALink: AxUCString): AxUCString;
var
  i: integer;
begin
  Result := '';
  for i := 1 to Length(ALink) do begin
    case ALink[i] of
      ' ': Result := Result + '%20';
      '#': Result := Result + '%23';
      '$': Result := Result + '%24';
      '%': begin
        if (StrToIntDef('$' + Copy(ALink,i + 1,2),MAXINT) = MAXINT) and not (Copy(ALink,i + 1,2)= '25')  then
          Result := Result + '%25'
        else
          Result := Result + ALink[i];
      end;
      '&': Result := Result + '%26';
      '+': Result := Result + '%2B';
      ',': Result := Result + '%2C';
      ';': Result := Result + '%3B';
      '=': Result := Result + '%3D';
      '@': Result := Result + '%40';
      '[': Result := Result + '%5B';
      ']': Result := Result + '%5D';
      '^': Result := Result + '%5E';
      '`': Result := Result + '%60';
      '{': Result := Result + '%7B';
      '}': Result := Result + '%7D';
      '~': Result := Result + '%7E';
      '': Result := Result + '%7F';
//      Reserved by windows, don not encode.
//      '/': Result := Result + '%2F';
//      ':': Result := Result + '%3A';
//      '<': Result := Result + '%3C';
//      '>': Result := Result + '%3E';
//      '?': Result := Result + '%3F';
//      '\': Result := Result + '%5C';
//      '|': Result := Result + '%7C';
      else Result := Result + ALink[i];
    end;
  end;
end;

function TXLSHyperlink.GetAddress: AxUCString;
begin
  Result := FAddress;
end;

function TXLSHyperlink.GetCol1: integer;
begin
  Result := FRef.Col1;
end;

function TXLSHyperlink.GetCol2: integer;
begin
  Result := FRef.Col2;
end;

function TXLSHyperlink.GetDescription: AxUCString;
begin
  Result := FDisplay;
end;

function TXLSHyperlink.GetRow1: integer;
begin
  Result := FRef.Row1;
end;

function TXLSHyperlink.GetRow2: integer;
begin
  Result := FRef.Row2;
end;

function TXLSHyperlink.GetToolTip: AxUCString;
begin
  Result := FTooltip;
end;

function TXLSHyperlink.GetWorkbookTarget(out ACol, ARow: integer): boolean;
var
  p: integer;
  S: AxUCString;
begin
  Result := FHyperlinkType = xhltWorkbook;
  if Result then begin
    p := RCPos('!',FAddress);
    if p > 1 then
      S := Copy(FAddress,p + 1,MAXINT)
    else
      S := FAddress;
    Result := RefStrToColRow(S,ACol,ARow);
  end;
end;

procedure TXLSHyperlink.Parse;
begin
  inherited;
  SetAddress(FAddress);
end;

procedure TXLSHyperlink.SetAddress(const Value: AxUCString);
var
  S,QS,TS,FS,VS: AxUCString;
  Ptgs: PXLSPtgs;
  Sz: integer;
  Drive: AxUCString;

function BeginWith(WS: AxUCString): boolean;
begin
  Result := AnsiLowercase(Copy(Value,1,Length(WS))) = WS;
end;

function IsURL(S: AxUCString): boolean;
var
  i: integer;
begin
  Result := False;
  S := Lowercase(S);
  for i := 1 to Length(S) do begin
    if not CharInSet(S[i],['a'..'z']) then begin
      Result := Copy(S,i,3) = '://';
      Break;
    end;
  end;
end;

function Fet(var Source: AxUCString; const Delim: AxUCString): AxUCString;
var
  i: integer;
begin
  Result := '';
  i := Pos(Delim,Source);
  if i > 0 then begin
    Result := Copy(Source,1,i-1);
    Delete(Source,1,i + (Length(Delim) - 1));
  end
  else begin
    Result := Source;
    Source := '';
  end;
end;

function StrReverse(Source: AxUCString): AxUCString;
var
  i: integer;
begin
  Result := '';
  for i := Length(Source) downto 1 do
    Result := Result + source[i];
end;

function RFet(var Source: AxUCString; Delim: AxUCString): AxUCString;
var
  i: integer;
  s: AxUCString;
begin
  s := StrReverse(Source);
  Result := '';
  i := Pos(StrReverse(Delim),s);
  if i > 0 then begin
    Result := StrReverse(Copy(s,1,i-1));
    Delete(s,1,i + (length(Delim) - 1));
    Source := StrReverse(s);
  end else begin
    Result := Source;
    Source := '';
  end;
end;

begin
  S := Value;
  Drive := ExtractFileDrive(S);
  if IsURL(Value) or BeginWith('www.') then begin
//  if BeginWith('http://') or BeginWith('https://') or BeginWith('www.') or BeginWith('ftp://') then begin

    FHyperlinkType := xhltURL;

    // We have a query string
    if CPos('?',S) > 0 then begin
      QS := RFet(S,'?');
      TS := '';
      while QS <> '' do begin
        //Parse out typical Fieldname=Value pairs then URL encode the values.
        FS := Fet(QS,'&');
        if CPos('=',FS) > 0 then begin
          VS := RFet(FS,'=');
          VS := EncodeFileHLink(VS);
          TS:=TS+FS+'='+VS;
          end
          else
            TS:=TS+EncodeFileHLink(FS);
          if QS <> '' then
            TS := TS + '&';
        end;
        S := S + '?' + TS;
      end
    else
      S := EncodeFileHLink(S);
  end
  else if BeginWith('\\') then begin
    FHyperlinkType := xhltUNC;
    S := 'file:///' + EncodeFileHLink(S);
  end
  else if BeginWith('file:///') or (Drive <> '') then begin
    if not BeginWith('file:///') then
      S := 'file:///' + EncodeFileHLink(S);
    FHyperlinkType := xhltFile;
  end
  else if BeginWith('mailto:') then begin
    FHyperlinkType := xhltURL;
  end
  else begin
    FOwner.FManager.Errors.IgnoreErrors := True;
    FOwner.FManager.Errors.LastError := 0;
    try
      Ptgs := Nil;
      Sz := NameEncodeFormula(FOwner.FManager,S,Ptgs,-1);
      if (Sz > 0) and (Ptgs <> Nil) then begin
        FOwner.FManager.Errors.IgnoreErrors := False;
//        if ((Sz = SizeOf(TXLSPtgsArea1d)) and (Ptgs.Id = xptgArea1d)) or ((Sz = SizeOf(TXLSPtgsRef1d)) and (Ptgs.Id = xptgRef1d)) then begin
          FHyperlinkType := xhltWorkbook;
          FToolTip := '#' + S;
//        end
//        else
//          FOwner.FManager.Errors.Error(S,XLSERR_HLINK_INVALID_CELLREF);
      end
      else begin
        // Do not preceed with 'file:///' will fail on local files. If needed
        // add it explicit to the url.
        S := {'file:///' + }EncodeFileHLink(S);
        FHyperlinkType := xhltFile;
      end;

      if Ptgs <> Nil then
        FreeMem(Ptgs);
    finally
      FOwner.FManager.Errors.IgnoreErrors := False;
    end;
  end;
  FAddress := S;

  FHyperlinkEnc := FHyperlinkType;
end;

procedure TXLSHyperlink.SetCol1(const Value: integer);
begin
  FRef.Col1 := Value;
end;

procedure TXLSHyperlink.SetCol2(const Value: integer);
begin
  FRef.Col2 := Value;
end;

procedure TXLSHyperlink.SetDescription(const Value: AxUCString);
begin
  FDisplay := Value;
end;

procedure TXLSHyperlink.SetRow1(const Value: integer);
begin
  FRef.Row1 := Value;
end;

procedure TXLSHyperlink.SetRow2(const Value: integer);
begin
  FRef.Row2 := Value;
end;

procedure TXLSHyperlink.SetToolTip(const Value: AxUCString);
begin
  FTooltip := Value;
end;

function TXLSHyperlink.SheetIndex: integer;
begin
  Result := -1;
end;

end.
