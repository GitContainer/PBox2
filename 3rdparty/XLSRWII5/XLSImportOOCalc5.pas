unit XLSImportOOCalc5;

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
     xpgPUtils, xpgPSimpleDOM, Xc12Utils5, Xc12DataStyleSheet5,
     XLSUtils5, XLSReadWriteII5, XLSRow5, XLSReadWriteZIP5, XLSSheetData5;

type TXLSImportOOC = class(TObject)
protected
     FZIP: TXLSZipArchiveIntfDelphi;
     FXLS: TXLSReadWriteII5;

     procedure DoContent;
     procedure DoBody(ANode: TXpgDOMNode);
     procedure DoSpreadsheet(ANode: TXpgDOMNode);
     procedure DoTable(ANode: TXpgDOMNode);
     procedure DoRow(ASheet: TXLSWorksheet; ANode: TXpgDOMNode; ACol: integer; var ARow: integer);
public
     constructor Create(AXLS: TXLSReadWriteII5);
     destructor Destroy; override;

     procedure Clear;

     procedure LoadFromFile(const AFilename: AxUCString);
     end;

implementation

{ TXLSImportOOC }

procedure TXLSImportOOC.Clear;
begin
  if FZip <> Nil then begin
    FZip.Close;
    FZip.Free;
  end;

  FZip := Nil;
end;

constructor TXLSImportOOC.Create(AXLS: TXLSReadWriteII5);
begin
  FXLS := AXLS;
end;

destructor TXLSImportOOC.Destroy;
begin
  Clear;

  inherited;
end;

procedure TXLSImportOOC.DoBody(ANode: TXpgDOMNode);
var
  Node: TXpgDOMNode;
begin
  Node := ANode.Find('office:spreadsheet');

  if Node <> Nil then
    DoSpreadsheet(Node);
end;

procedure TXLSImportOOC.DoContent;
var
  Stream: TStream;
  XML   : TXpgSimpleDOM;
  Node  : TXpgDOMNode;
begin
  Stream := FZip.OpenStream('content.xml');
  if Stream = Nil then
    raise Exception.Create('Can not find content.xml');

  XML := TXpgSimpleDOM.Create;
  try
    XML.LoadFromStream(Stream);

    Node := XML.Root.Find('office:document-content/office:body');

    if Node <> Nil then
      DoBody(Node);
  finally
    XML.Free;
  end;

  Stream.Free;
end;

procedure TXLSImportOOC.DoRow(ASheet: TXLSWorksheet; ANode: TXpgDOMNode; ACol: integer; var ARow: integer);
var
  i   : integer;
  S   : AxUCString;
  HH,
  MM,
  SS  : word;
  Fmla: AxUCString;
  cInc: integer;
  Node: TXpgDOMNode;
  N   : TXpgDOMNode;
  dVal: double;
begin

  for i := 0 to ANode.Count - 1 do begin
    Node := ANode[i];

    if Node.QName = 'table:table-cell' then begin
      cInc := Node.Attributes.FindValueIntDef('table:number-columns-repeated',0);
      Inc(ACol,cInc);

      S := Node.Attributes.FindValue('office:value-type');
      if S = 'float' then begin
        Fmla := Node.Attributes.FindValue('table:formula');

        dVal := Node.Attributes.FindValueFloatDef('office:value',0);

        ASheet.AsFloat[ACol,ARow] := dVal;
      end
      else if S = 'string' then begin
        N := Node.Find('text:p');
        if N <> Nil then
          ASheet.AsString[ACol,ARow] := N.Content;
      end
      else if S = 'boolean' then begin
        S := Uppercase(Node.Attributes.FindValueDef('office:boolean-value','FALSE'));

        ASheet.AsBoolean[ACol,ARow] := S = 'TRUE';
      end
      else if S = 'time' then begin
        S := Node.Attributes.FindValueDef('office:time-value','');

        if (S <> '') and (Copy(S,1,2) = 'PT') then begin
          // PT08H03M00S
          // 12345678901
          HH := StrToIntDef(Copy(S,3,2),0);
          MM := StrToIntDef(Copy(S,6,2),0);
          SS := StrToIntDef(Copy(S,9,2),0);

          ASheet.AsDateTime[ACol,ARow] := EncodeTime(HH,MM,SS,0);
        end;
      end
      else if S = 'date' then begin
        S := Node.Attributes.FindValueDef('office:date-value','');

        if S <> '' then
          ASheet.AsDateTime[ACol,ARow] := XmlStrToDateTime(S);
      end;

      if cInc = 0 then
        Inc(ACol);
    end;
  end;

  Inc(ARow);
end;

procedure TXLSImportOOC.DoSpreadsheet(ANode: TXpgDOMNode);
var
  Node: TXpgDOMNode;
begin
  Node := ANode.Find('table:table');

  if Node <> Nil then
    DoTable(Node);
end;

procedure TXLSImportOOC.DoTable(ANode: TXpgDOMNode);
var
  i    : integer;
  C,R  : integer;
//  cInc : integer;
//  rInc : integer;
  Node : TXpgDOMNode;
  Sheet: TXLSWorksheet;
begin
  Sheet := FXLS[0];

  Sheet.Name := ANode.Attributes.FindValueDef('table:name',Sheet.Name);

  C := 0;
  R := 0;

//  cInc := 1;
//  rInc := 1;

  for i := 0 to ANode.Count - 1 do begin
    Node := ANode[i];

    if Node.QName = 'table:table-column' then begin
//      cInc := Node.Attributes.FindValueIntDef('table:number-columns-repeated',1);
    end
    else if Node.QName = 'table:table-row' then begin
//      rInc := Node.Attributes.FindValueIntDef('table:number-rows-repeated',1);

      DoRow(Sheet,Node,C,R);
    end;

//    Inc(C,cInc);
//    Inc(R,rInc);
  end;
end;

procedure TXLSImportOOC.LoadFromFile(const AFilename: AxUCString);
begin
  Clear;

  FZip := TXLSZipArchiveIntfDelphi.Create;
  FZip.LoadFromFile(AFilename);

  DoContent;
end;

end.
