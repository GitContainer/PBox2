unit Xc12FileData5;

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

uses Classes, SysUtils, Contnrs, IniFiles,
     Xc12Utils5,
     XLSUtils5,
     xpgPSimpleDOM, XLSReadWriteOPC5, Xc12DefaultData5,
     xpgParseContentType, xpgParseDocPropsApp;

type TXLSSavedFileDataList = class;

     TXLSSavedFileData = class(TObject)
protected
     FType   : AxUCString;
     FTarget : AxUCString;
     FContent: AxUCString;
     FData   : TStream;
     FChilds : TXLSSavedFileDataList;
public
     constructor Create;
     destructor Destroy; override;

     procedure AddChilds;

     property Type_  : AxUCString read FType write FType;
     property Target : AxUCString read FTarget write FTarget;
     property Content: AxUCString read FContent write FContent;
     property Data   : TStream read FData write FData;
     property Childs : TXLSSavedFileDataList read FChilds;
     end;

     TXLSSavedFileDataList = class(TObjectList)
private
     function GetItems(Index: integer): TXLSSavedFileData;
protected
     FHasVBA      : boolean;
     FHasSignedVBA: boolean;
public
     constructor Create;
     destructor Destroy; override;

     function  FindByType(AType: AxUCString): TXLSSavedFileData;

     procedure Add(AType,ATarget,AContent: AxUCString; AData: TStream);

     property HasMacros   : boolean read FHasVBA;
     property HasSignedVBA: boolean read FHasSignedVBA;

     property Items[Index: integer]: TXLSSavedFileData read GetItems; default;
     end;

type TXLSSavedFileDataSheet = class(TObject)
protected
     FSheetData: TXLSSavedFileDataList;
     FTableData: TXLSSavedFileDataList;
public
     constructor Create;
     destructor Destroy; override;

     property SheetData: TXLSSavedFileDataList read FSheetData;
     property TableData: TXLSSavedFileDataList read FTableData;
     end;

type TXc12FileData = class(TObject)
private
     procedure SetUseAlternateZip(const Value: boolean);
     function  GetHasMacros: boolean;
protected
     FContentType  : TXPGDocContentType;
     FDocPropsApp  : TXPGDocDocPropsApp;
     FDocPropsCore : TXpgSimpleDOM;
     FOPC          : TOPC_XLSX;
     FSavedRoot    : TXLSSavedFileDataList;
     FSavedWorkbook: TXLSSavedFileDataList;
{$ifdef DELPHI_5}
     FSavedSheets  : TStringList;
{$else}
     FSavedSheets  : THashedStringList;
{$endif}
     FUseAlternateZip: boolean;

     procedure SetDefaultData;
     procedure ClearSavedSheets;

     procedure ReadContentTypes;
     procedure ReadDocPropsApp;
     procedure ReadDocPropsCore;
     procedure ReadTheme;

     procedure WriteDocPropsCore;
     procedure WriteTheme;
public
     constructor Create;
     destructor Destroy; override;

     procedure Clear;

     procedure LoadFromStream(AZIPStream: TStream);
     procedure ReadUnusedData;
     procedure WriteUnusedData;
     procedure WriteUnusedDataSheet(AIndex: integer; AOPCSheet: TOPCItem);
     procedure BeginSaveToStream(AZIPStream: TStream);
     procedure CommitSaveToStream;

     procedure AddSaveSheet(AId: AxUCString);

     function  StreamByName(const AName: AxUCString): TStream;

     property OPC            : TOPC_XLSX read FOPC;
     property DocPropsApp    : TXPGDocDocPropsApp read FDocPropsApp;
     property DocPropsCore   : TXpgSimpleDOM read FDocPropsCore;
     property HasMacros      : boolean read GetHasMacros;
     property UseAlternateZip: boolean read FUseAlternateZip write SetUseAlternateZip;
     end;

implementation

{ TXLSFileDataXLSX }

procedure TXc12FileData.AddSaveSheet(AId: AxUCString);
begin
  FSavedSheets.Add(AId);
end;

procedure TXc12FileData.BeginSaveToStream(AZIPStream: TStream);
begin
  FOPC.Clear;

  FOPC.OpenWrite(AZIPStream,FUseAlternateZip);

  WriteDocPropsCore;
end;

procedure TXc12FileData.Clear;
begin
  FDocPropsApp.Root.Clear;
  FDocPropsCore.Clear;

  FSavedRoot.Clear;
  FSavedWorkbook.Clear;
  ClearSavedSheets;
  FSavedSheets.Clear;

  FOPC.Clear;

  SetDefaultData;
end;

procedure TXc12FileData.ClearSavedSheets;
var
  i: integer;
begin
  for i := 0 to FSavedSheets.Count - 1 do
    TXLSSavedFileDataSheet(FSavedSheets.Objects[i]).Free;
end;

procedure TXc12FileData.CommitSaveToStream;
begin
  if FSavedWorkbook.FindByType(OPC_XLSX_THEME) = Nil then
    WriteTheme;
  WriteUnusedData;
  FOPC.Close;
end;

constructor TXc12FileData.Create;
begin
{$ifdef MSWINDOWS}
  FUseAlternateZip := False;
{$else}
  FUseAlternateZip := True;
{$endif}

  FOPC           := TOPC_XLSX.Create;
  FContentType   := TXPGDocContentType.Create;
  FDocPropsApp   := TXPGDocDocPropsApp.Create;
  FDocPropsCore  := TXpgSimpleDOM.Create;

  FSavedRoot     := TXLSSavedFileDataList.Create;
  FSavedWorkbook := TXLSSavedFileDataList.Create;
{$ifdef DELPHI_5}
  FSavedSheets   := TStringList.Create;
{$else}
  FSavedSheets   := THashedStringList.Create;
{$endif}

  SetDefaultData;
end;

destructor TXc12FileData.Destroy;
begin
  FOPC.Free;
  FContentType.Free;
  FDocPropsApp.Free;
  FDocPropsCore.Free;
  FSavedRoot.Free;
  FSavedWorkbook.Free;
  ClearSavedSheets;
  FSavedSheets.Free;

  inherited;
end;

function TXc12FileData.GetHasMacros: boolean;
begin
  Result := FSavedWorkbook.HasMacros;
end;

procedure TXc12FileData.LoadFromStream(AZIPStream: TStream);
begin
  FOPC.OpenRead(AZIPStream,FUseAlternateZip);

  ReadContentTypes;
  ReadDocPropsCore;
  ReadTheme;
end;

procedure TXc12FileData.ReadContentTypes;
var
  Stream: TStream;
begin
  Stream := FOPC.ReadContentType;
  if Stream <> Nil then begin
    FContentType.LoadFromStream(Stream);
    Stream.Free;
  end;
end;

procedure TXc12FileData.ReadDocPropsApp;
var
  Stream: TStream;
begin
  Stream := FOPC.ReadDocPropsApp;
  try
    FDocPropsApp.LoadFromStream(Stream);
  finally
    Stream.Free;
  end;
end;

procedure TXc12FileData.ReadDocPropsCore;
var
  Stream: TStream;
begin
  Stream := FOPC.ReadDocPropsCore;
  if Stream <> Nil then begin
    FDocPropsCore.LoadFromStream(Stream);
    Stream.Free;
  end;
end;

procedure TXc12FileData.ReadTheme;
var
  i     : TXc12ClrSchemeColor;
  DOM   : TXpgSimpleDOM;
  Stream: TStream;
  Node,N: TXpgDOMNode;
//  Attr  : TXpgDOMAttribute;
begin
  Stream := FOPC.ReadTheme;
  if Stream <> Nil then begin
    DOM := TXpgSimpleDOM.Create;
    try
      DOM.LoadFromStream(Stream);

      for i := Low(TXc12ClrSchemeColor) to High(TXc12ClrSchemeColor) do begin
        Node := DOM.Root.Find('a:theme/a:themeElements/a:clrScheme/a:' + TXc12ClrSchemeColorName[i]);
        if Node <> Nil then begin
          N := Node.Find('a:srgbClr');
          if N <> Nil then
            Xc12DefColorScheme[i] := N.Attributes.FindValueHexDef('val',Xc12DefColorScheme[i]);
        end;
      end;
    finally
      DOM.Free;
    end;

    Stream.Free;
  end;
end;

procedure TXc12FileData.ReadUnusedData;
var
  i          : integer;
  List       : TOPCItemList;
  SaveList   : TXLSSavedFileDataList;
  SavedSheets: TXLSSavedFileDataSheet;
  SavedData  : TXLSSavedFileData;

procedure AddSheet(AId: AxUCString);
var
  i: integer;
begin
  i := FSavedSheets.IndexOf(AId);
  if i < 0 then
    raise XLSRWException.CreateFmt('Unknown OPC sheet id "%s"',[AId]);
  SavedSheets := TXLSSavedFileDataSheet(FSavedSheets.Objects[i]);
  if SavedSheets = Nil then begin
    SavedSheets := TXLSSavedFileDataSheet.Create;
    FSavedSheets.Objects[i] := SavedSheets;
  end;
end;

begin
  SaveList := Nil;

  List := TOPCItemList.Create;
  try
    FOPC.SaveUnchecked(List);

    for i := 0 to List.Count - 1 do begin
      if (List[i].TargetMode <> otmInternal) or
         (List[i].Name = 'calcChain.xml') or
         (List[i].Name = 'app.xml') or
         (Copy(List[i].Name,1,15) = 'printerSettings') or
         (Copy(List[i].Name,1,9) = 'oleObject') then
        Continue;
      if List[i].Parent.Type_ = OPC_XLSX_WORKSHEET then begin
        AddSheet(List[i].Parent.Id);
        if List[i].Type_ = OPC_XLSX_TABLE then
          SaveList := SavedSheets.TableData
        else
          SaveList := SavedSheets.SheetData;
      end
      else if List[i].Parent.Type_ = OPC_XLSX_WORKBOOK then
        SaveList := FSavedWorkbook
      else if List[i].Parent.Type_ = OPC_ROOT then
        SaveList := FSavedRoot
      else begin
        SavedData := FSavedRoot.FindByType(List[i].Parent.Type_);
        if SavedData <> Nil then begin
          SavedData.AddChilds;
          SaveList := SavedData.Childs;
        end
        else begin
          SavedData := FSavedWorkbook.FindByType(List[i].Parent.Type_);
          if SavedData <> Nil then begin
            SavedData.AddChilds;
            SaveList := SavedData.Childs;
          end
        end;

        if SaveList = Nil then begin
          if Copy(List[i].Name,1,5) = 'image' then
            Continue;
          SaveList := FSavedRoot;
        end;
      end;
//        raise XLSRWException.CreateFmt('Unknown OPC type "%s" in save data',[List[i].Type_]);
      SaveList.Add(List[i].Type_,List[i].Target,List[i].Content,List[i].Data);
      List[i].Data := Nil;
    end;
  finally
    List.Free;
  end;
end;

procedure TXc12FileData.SetDefaultData;
begin
  FDocPropsApp.Root.RootAttributes.AddNameValue('xmlns',OOXML_URI_OFFICEDOC_EXTENDED_PROPERIES);
  FDocPropsApp.Root.RootAttributes.AddNameValue('xmlns:vt',OOXML_URI_OFFICEDOC_EXTENDED_DOCPROPSVTYPES);
  FDocPropsCore.LoadFromString(XLS_DEFAULT_DOCPROPS_CORE);
end;

procedure TXc12FileData.SetUseAlternateZip(const Value: boolean);
begin
// Must use built in zip when not windows.
{$ifdef MSWINDOWS}
  FUseAlternateZip := Value;
{$endif}
end;

function TXc12FileData.StreamByName(const AName: AxUCString): TStream;
var
  i: integer;
begin
  for i := 0 to FSavedWorkbook.Count - 1 do begin
    if FSavedWorkbook[i].Target = AName then begin
      Result := FSavedWorkbook[i].Data;
      Exit;
    end;
  end;
  Result := Nil;
end;

procedure TXc12FileData.WriteDocPropsCore;
var
  Stream: TMemoryStream;
begin
  Stream := TMemoryStream.Create;
  try
    FDocPropsCore.SaveToStream(Stream);
    FOPC.AddDocPropsCore(Stream);
  finally
    Stream.Free;
  end;
end;

procedure TXc12FileData.WriteTheme;
var
  Stream: TStringStream;
begin
  Stream := TStringStream.Create('');
  try
    Stream.WriteString(XLS_DEFAULT_THEME);
    FOPC.AddTheme(Stream,1);
  finally
    Stream.Free;
  end;
end;

procedure TXc12FileData.WriteUnusedData;
var
  i: integer;

procedure SaveData(AParent: TOPCItem; AData: TXLSSavedFileData);
var
  i  : integer;
  OPC: TOPCItem;
begin
  OPC := FOPC.ItemCreate(AParent,AData.Type_,AData.Target,AData.Content);
  FOPC.ItemWrite(OPC,AData.Data);
  FOPC.ItemClose(OPC);

  if AData.Childs <> Nil then begin
    for i := 0 to AData.Childs.Count - 1 do
      SaveData(OPC,AData.Childs[i]);
  end;
end;

begin
  for i := 0 to FSavedRoot.Count - 1 do
    SaveData(FOPC.Root,FSavedRoot[i]);
  for i := 0 to FSavedWorkbook.Count - 1 do
    SaveData(FOPC.Workbook,FSavedWorkbook[i]);

// Saved below
//  for i := 0 to FSavedSheets.Count - 1 do begin
//    if FSavedSheets.Objects[i] <> Nil then begin
//      List := TXLSSavedFileDataSheet(FSavedSheets.Objects[i]).SheetData;
//      OPCSheet := FOPC.FindSheet(i + 1);
//      if OPCSheet <> Nil then begin
//        for j := 0 to List.Count - 1 do
//          SaveData(OPCSheet,List[j]);
//      end;
//    end;
//  end;
end;

procedure TXc12FileData.WriteUnusedDataSheet(AIndex: integer; AOPCSheet: TOPCItem);
var
  i: integer;
  List: TXLSSavedFileDataList;
  OPC: TOPCItem;
begin
  if (AIndex >= 0) and (AIndex < FSavedSheets.Count) then begin
    if FSavedSheets.Objects[AIndex] <> Nil then begin

      List := TXLSSavedFileDataSheet(FSavedSheets.Objects[AIndex]).SheetData;
      for i := 0 to List.Count - 1 do begin
        OPC := FOPC.ItemCreate(AOPCSheet,List[i].Type_,List[i].Target,List[i].Content);
        FOPC.ItemWrite(OPC,List[i].Data);
        FOPC.ItemClose(OPC);
      end;

      List := TXLSSavedFileDataSheet(FSavedSheets.Objects[AIndex]).TableData;
      for i := 0 to List.Count - 1 do begin
        OPC := FOPC.ItemCreate(AOPCSheet,List[i].Type_,List[i].Target,List[i].Content);
        FOPC.ItemWrite(OPC,List[i].Data);
        FOPC.ItemClose(OPC);
      end;
    end;
  end;
end;

{ TXLSSavedFileData }

procedure TXLSSavedFileData.AddChilds;
begin
  if FChilds = Nil then
    FChilds := TXLSSavedFileDataList.Create;
end;

constructor TXLSSavedFileData.Create;
begin

end;

destructor TXLSSavedFileData.Destroy;
begin
  if FData <> Nil then
    FData.Free;

  if FChilds <> Nil then
    FChilds.Free;

  inherited;
end;

{ TXLSSavedFileDataList }

procedure TXLSSavedFileDataList.Add(AType, ATarget,AContent: AxUCString; AData: TStream);
var
  Item: TXLSSavedFileData;
begin
  Item := TXLSSavedFileData.Create;
  Item.Type_ := AType;
  Item.Target := ATarget;
  Item.Content := AContent;
  Item.Data := AData;

  FHasVBA := ATarget = 'vbaProject.bin';
  FHasSignedVBA := ATarget = 'vbaProjectSignature.bin';

  inherited Add(Item);
end;

constructor TXLSSavedFileDataList.Create;
begin
  inherited Create;
end;

destructor TXLSSavedFileDataList.Destroy;
begin

  inherited;
end;

function TXLSSavedFileDataList.FindByType(AType: AxUCString): TXLSSavedFileData;
var
  i: integer;
begin
  for i := 0 to Count - 1 do begin
    if Items[i].Type_ = AType then begin
      Result := Items[i];
      Exit;
    end;
  end;
  Result := Nil;
end;

function TXLSSavedFileDataList.GetItems(Index: integer): TXLSSavedFileData;
begin
  Result := TXLSSavedFileData(inherited Items[Index]);
end;

{ TXLSSavedFileDataSheet }

constructor TXLSSavedFileDataSheet.Create;
begin
  FSheetData := TXLSSavedFileDataList.Create;
  FTableData := TXLSSavedFileDataList.Create;
end;

destructor TXLSSavedFileDataSheet.Destroy;
begin
  FSheetData.Free;
  FTableData.Free;

  inherited;
end;

end.
