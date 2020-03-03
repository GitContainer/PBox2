unit BIFF_EscherCtrlObj5;

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

uses Classes, SysUtils, Math,
     BIFF_Utils5, BIFF_Stream5, BIFF_RecsII5, BIFF_RecordStorage5, BIFF_EscherTypes5,
     BIFF_RecordStorageChart5, BIFF_ExcelFuncII5, BIFF_Names5,
     XLSUtils5, XLSFormulaTypes5;

const SHP_NOTE_OPTION_ALWAYS_VISIBLE = $0002;

type TShapeOutsideMso = class(TObject)
protected
     FOwnsObjData: boolean;
     FObjId: word;
     FOptions: word;
     FShapeType: TMSOShapeType;

     function  GetDynArraySize(Arr: array of byte): integer;
     procedure WriteDynArray(Stream: TXLSStream; RecId: integer; Arr: array of byte);
public
     constructor Create(ObjId: integer; ST: TMSOShapeType);
     procedure Assign(Shape: TShapeOutsideMso); virtual; abstract;
     procedure Read(Stream: TXLSStream; PBuf: PByteArray); virtual; abstract;
     procedure Write(Stream: TXLSStream); virtual; abstract;
     function  ParseOBJ(OBJ: TBaseRecordStorage): boolean; virtual;

     property OwnsObjData: boolean read FOwnsObjData;
     property ObjId: word read FObjId write FObjId;
     property ShapeType: TMSOShapeType read FShapeType;
     end;

     TShapeOutsideMsoBaseText = class(TShapeOutsideMso)
private
     procedure SetText(const Value: AxUCString);
     procedure SetTXORuns(const Value: TDynTXORUNArray);
protected
     FpTXO: PRecTXO;
     FText: AxUCString;
     FTXORuns: TDynTXORUNArray;
public
     constructor Create(ObjId: integer; ST: TMSOShapeType);
     destructor Destroy; override;
     procedure Assign(Shape: TShapeOutsideMso); override;
     procedure SetDefaultTxo;
     procedure Read(Stream: TXLSStream; PBuf: PByteArray); override;
     procedure Write(Stream: TXLSStream); override;

     property TXO: PRecTXO read FpTXO;
     property Text: AxUCString read FText write SetText;
     property Formatting: TDynTXORUNArray read FTXORuns write SetTXORuns;
     property TXORuns: TDynTXORUNArray read FTXORuns write FTXORuns;
     end;

     TShapeOutsideMsoNote = class(TShapeOutsideMsoBaseText)
protected
     FCellCol,FCellRow: integer;
     FOptions: word;
     FAuthor: AxUCString;
public
     procedure Assign(Shape: TShapeOutsideMso); override;
     property CellCol: integer read FCellCol write FCellCol;
     property CellRow: integer read FCellRow write FCellRow;
     property Options: word read FOptions write FOptions;
     property Author: AxUCString read FAuthor write FAuthor;
     end;

     TShapeControl = class(TShapeOutsideMso)
protected
public
     end;

     TObjLBS = class(TObject)
protected
     FLBSOrig: array of byte;
     FUnknownFmlaData1: longword;
     FFormula: TDynByteArray;
     FUnknownFmlaData2: byte;
     FEntryCount: word;
     FSelectedEntry: word;
     FStyle: word;
     FLinkedEdit: word;
public
     constructor Create;
     function  Size(OrigSize: boolean): integer; virtual;
     function  LoadFromBuf(Buf: PByteArray; BufLen: integer): integer; virtual;
     procedure SaveToStream(Stream: TXLSStream; SaveOrig: boolean); virtual;
     procedure GetSourceArea(var Sheet,Col1,Row1,Col2,Row2: integer);
     procedure SetSourceArea(Sheet,Col1,Row1,Col2,Row2: integer);
     procedure GetSourceFmla(var Fmla: TDynByteArray);
     function  HasSourceArea: boolean;

     property Style: word read FStyle write FStyle;
     property LinkedEdit: word read FLinkedEdit write FLinkedEdit;
     property SelectedIndex: word read FSelectedEntry write FSelectedEntry;
     end;

     TObjLBSListBox = class(TObjLBS)
protected
     FMultiselect: array of byte;
public
     function  Size(OrigSize: boolean): integer; override;
     function  LoadFromBuf(Buf: PByteArray; BufLen: integer): integer; override;
     procedure SaveToStream(Stream: TXLSStream; SaveOrig: boolean); override;
     end;

     TObjLBSComboBox = class(TObjLBS)
private
     FLineCount: word;
     FComboStyle: word;
     FUnknown: word;
protected
public
     constructor Create;
     function  Size(OrigSize: boolean): integer; override;
     function  LoadFromBuf(Buf: PByteArray; BufLen: integer): integer; override;
     procedure SaveToStream(Stream: TXLSStream; SaveOrig: boolean); override;

     property ComboStyle: word read FComboStyle write FComboStyle;
     property LineCount: word read FLineCount write FLineCount;
     end;

     TShapeControlListBox = class(TShapeControl)
private
     function  GetSelectedEntry: word;
     procedure SetSelectedEntry(const Value: word);
protected
     FIsComboBox: boolean;
     FIsComboBoxAutoFilter: boolean;
     FSBS: TObjSBS;
     FSBSFormula: array of byte;
     FSBSFormulaOrig: array of byte;
     FLBS: TObjLBS;
     FMacroFormula: array of byte;
     FChanged: boolean;
public
     constructor Create(ObjId: integer; ST: TMSOShapeType);
     destructor Destroy; override;
     procedure Assign(Shape: TShapeOutsideMso); override;
     procedure Read(Stream: TXLSStream; PBuf: PByteArray); override;
     procedure Write(Stream: TXLSStream); override;
     function  ParseOBJ(OBJ: TBaseRecordStorage): boolean; override;
     procedure NameIndexChanged(Delta: integer);
     procedure GetDestCell(var Col,Row: integer);
     procedure SetDestCell(Col,Row: integer);
     procedure GetSourceArea(var Sheet,Col1,Row1,Col2,Row2: integer);
     procedure SetSourceArea(Sheet,Col1,Row1,Col2,Row2: integer);
     procedure GetSourceFmla(var Fmla: TDynByteArray);
     function  HasSourceArea: boolean;
     function  HasDestCell: boolean;

     property IsComboBox: boolean read FIsComboBox write FIsComboBox;
     property IsComboBoxAutoFilter: boolean read FIsComboBoxAutoFilter write FIsComboBoxAutoFilter;
     property SelectedIndex: word read GetSelectedEntry write SetSelectedEntry;
     end;

     TShapeControlComboBox = class(TShapeControlListBox)
private
     function  GetIsAutoFilter: boolean;
     procedure SetIsAutoFilter(Value: boolean);
public
     constructor Create(ObjId: integer; ST: TMSOShapeType);
     property IsAutoFilter: boolean read GetIsAutoFilter write SetIsAutoFilter;
     end;

     TShapeOutsideMsoBaseTextMacro = class(TShapeOutsideMsoBaseText)
protected
     FInternalNames: TInternalNames;
     FMacroFormula: array of byte;
public
     constructor Create(ObjId: integer; ST: TMSOShapeType; InternalNames: TInternalNames);
     function  SetMacro(Value: AxUCString): boolean;
     function  GetMacro: AxUCString;
     procedure NameIndexChanged(Delta: integer);
     end;

     TShapeControlButton = class(TShapeOutsideMsoBaseTextMacro)
protected
public
     constructor Create(ObjId: integer; InternalNames: TInternalNames);
     procedure Write(Stream: TXLSStream); override;
     function  ParseOBJ(OBJ: TBaseRecordStorage): boolean; override;
     end;

     TShapeControlCheckBox = class(TShapeOutsideMsoBaseTextMacro)
protected
     FCBLS: array of byte;
     FCBLSFMLA: array of byte;
     FCBLSDATA: array of byte;

     function  GetState: word;
     procedure SetState(const Value: word);
public
     constructor Create(ObjId: integer; InternalNames: TInternalNames);
     procedure Write(Stream: TXLSStream); override;
     function  ParseOBJ(OBJ: TBaseRecordStorage): boolean; override;
     function  HasDestCell: boolean;
     procedure GetDestCell(var Col,Row: integer);
     procedure SetDestCell(Col,Row: integer);

     property State: word read GetState write SetState;
     end;

     TShapeControlRadioButton = class(TShapeOutsideMsoBaseTextMacro)
protected
     FCBLS: array of byte;
     FRBO: array of byte;
     FCBLSDATA: array of byte;
     FCBLSFMLA: array of byte;
     FRBODATA: array of byte;

     function  GetState: word;
     procedure SetState(const Value: word);
public
     constructor Create(ObjId: integer; InternalNames: TInternalNames);
     procedure Write(Stream: TXLSStream); override;
     function  ParseOBJ(OBJ: TBaseRecordStorage): boolean; override;
     function  HasDestCell: boolean;
     procedure GetDestCell(var Col,Row: integer);
     procedure SetDestCell(Col,Row: integer);

     property State: word read GetState write SetState;
     end;

     TShapeControlScrollBar = class(TShapeOutsideMso)
private
     procedure SetIncChange(const Value: integer);
     procedure SetMaxValue(const Value: integer);
     procedure SetMinValue(const Value: integer);
     procedure SetPageChange(const Value: integer);
     procedure SetValue(const Value: integer);
     function GetIncChange: integer;
     function GetMaxValue: integer;
     function GetMinValue: integer;
     function GetPageChange: integer;
     function GetValue: integer;
protected
     FSBS: array of byte;
     FSBSFormula: array of byte;
     FSBSFormulaOrig: array of byte;
     FMacroFormula: array of byte;
public
     constructor Create(ObjId: integer);
     procedure Write(Stream: TXLSStream); override;
     function  ParseOBJ(OBJ: TBaseRecordStorage): boolean; override;
     procedure GetDestCell(var Col,Row: integer);
     procedure Assign(Shape: TShapeOutsideMso); override;
     procedure Read(Stream: TXLSStream; PBuf: PByteArray); override;

     property Value: integer read GetValue write SetValue;
     property MinValue: integer read GetMinValue write SetMinValue;
     property MaxValue: integer read GetMaxValue write SetMaxValue;
     property IncChange: integer read GetIncChange write SetIncChange;
     property PageChange: integer read GetPageChange write SetPageChange;
     end;

implementation

{ TShapeOutsideMso }

constructor TShapeOutsideMso.Create(ObjId: integer; ST: TMSOShapeType);
begin
  FObjId := ObjId;
  FShapeType := ST;
end;

function TShapeOutsideMso.GetDynArraySize(Arr: array of byte): integer;
begin
  if Length(Arr) > 0 then
    Result := Length(Arr) + SizeOf(TBIFFHeader)
  else
    Result := 0;
end;

function TShapeOutsideMso.ParseOBJ(OBJ: TBaseRecordStorage): boolean;
begin
  // Do nothing.
  Result := False;
end;

procedure TShapeOutsideMso.WriteDynArray(Stream: TXLSStream; RecId: integer; Arr: array of byte);
begin
  if Length(Arr) > 0 then begin
    Stream.WriteHeader(RecId,Length(Arr));
    Stream.Write(Arr[0],Length(Arr));
  end;
end;

{ TShapeControlListBox }

procedure TShapeControlListBox.Assign(Shape: TShapeOutsideMso);
begin
  raise XLSRWException.Create('TShapeControlListBox: Assign not implemented');
  inherited;
end;

constructor TShapeControlListBox.Create(ObjId: integer; ST: TMSOShapeType);
begin
  inherited Create(FObjId,ST);
  FOwnsObjData := True;
  FOptions := $2011;
  FSBS.Reserved[0] := $00;
  FSBS.Reserved[1] := $00;
  FSBS.Reserved[2] := $00;
  FSBS.Reserved[3] := $00;
  FSBS.ScrollPos := 0;
  FSBS.MinimumValue := 0;
  FSBS.MaximumValue := 4;
  FSBS.ScrollWhenArrowClicked := 1;
  FSBS.ScrollWhenBarClicked := 5;
  FSBS.Horizontal := 0;
  FSBS.Width := $10;
  FSBS.Options := $0001;
  // Last byte is unknown
  SetLength(FSBSFormula,6);
  FSBSFormula[0] := xptgRef97;
  FSBSFormula[5] := 0;
  PWordArray(@FSBSFormula[1])[0] := 1;
  PWordArray(@FSBSFormula[1])[1] := 1;

  if FIsComboBox then
    FLBS := TObjLBSComboBox.Create
  else
    FLBS := TObjLBSListBox.Create;
  FChanged := True;
end;

destructor TShapeControlListBox.Destroy;
begin
  FLBS.Free;
  inherited;
end;

procedure TShapeControlListBox.GetDestCell(var Col, Row: integer);
begin
  if FSBSFormula[0] <> xptgRef97 then
    raise XLSRWException.Create('Unknown SBS formula.');
  Row := PWordArray(@FSBSFormula[1])[0];
  Col := PWordArray(@FSBSFormula[1])[1];
end;

function TShapeControlListBox.GetSelectedEntry: word;
begin
  Result := FLBS.FSelectedEntry;
end;

procedure TShapeControlListBox.GetSourceArea(var Sheet,Col1, Row1, Col2, Row2: integer);
begin
  FLBS.GetSourceArea(Sheet,Col1,Row1,Col2,Row2);
end;

procedure TShapeControlListBox.GetSourceFmla(var Fmla: TDynByteArray);
begin
  FLBS.GetSourceFmla(Fmla);
end;

function TShapeControlListBox.HasDestCell: boolean;
begin
  Result := Length(FSBSFormula) > 0;
end;

function TShapeControlListBox.HasSourceArea: boolean;
begin
  Result := FLBS.HasSourceArea;
end;

procedure TShapeControlListBox.NameIndexChanged(Delta: integer);
begin
  //  !!!! The location of PPTGName8 in the formula can probably be elswhere.
  if Length(FMacroFormula) >= 10 then
    PPTGName8(@FMacroFormula[9]).NameIndex := PPTGName8(@FMacroFormula[9]).NameIndex + Delta;
end;

function TShapeControlListBox.ParseOBJ(OBJ: TBaseRecordStorage): boolean;
var
  i: integer;
  ObjType: word;
begin
  SetLength(FMacroFormula,0);
  SetLength(FSBSFormulaOrig,0);
  ObjType := 0;

  for i := 0 to OBJ.Count - 1 do begin
    case OBJ[i].RecId of
      OBJREC_CMO: begin
        ObjType := PObjCMO(@OBJ[i].Data).ObjType;
        FObjId := PObjCMO(@OBJ[i].Data).ObjId;
        FOptions := PObjCMO(@OBJ[i].Data).Options;
        // Autofilter
        if (FOptions and $0100) = $0100 then
          SetLength(FSBSFormula,0);
      end;
      OBJREC_SBS: begin
        Move(OBJ[i].Data,FSBS,SizeOf(TObjSBS));
      end;
      OBJREC_SBSFMLA: begin
        SetLength(FSBSFormulaOrig,OBJ[i].Length);
        Move(PByteArray(@OBJ[i].Data)[0],FSBSFormulaOrig[0],OBJ[i].Length);
        if PWordArray(@OBJ[i].Data)[0] = 5 then begin
          SetLength(FSBSFormula,PWordArray(@OBJ[i].Data)[0] + 1);
          Move(PByteArray(@OBJ[i].Data)[6],FSBSFormula[0],PWordArray(@OBJ[i].Data)[0] + 1);
        end;
      end;
      OBJREC_LBSDATA: begin
        FLBS.LoadFromBuf(PByteArray(@OBJ[i].Data),OBJ[i].Length);
        // LBSDATA list box has no END record
        if ObjType = OBJTYPE_LIST_BOX then
          Break;
      end;
      OBJREC_MACRO: begin
        SetLength(FMacroFormula,OBJ[i].Length);
        Move(PByteArray(@OBJ[i].Data)[0],FMacroFormula[0],OBJ[i].Length);
      end;
      OBJREC_END: begin
      end;
      else
        raise XLSRWException.CreateFmt('Unknown record %.4X in OBJ LBS',[OBJ[i].RecId]);
    end;
  end;
  Result := True;
  FOwnsObjData := Result;
  FChanged := False;
end;

procedure TShapeControlListBox.Read(Stream: TXLSStream; PBuf: PByteArray);
begin
  inherited;
end;

procedure TShapeControlListBox.SetDestCell(Col, Row: integer);
begin
  if (Length(FSBSFormula) <> 6) or (FSBSFormula[0] <> xptgRef97) then
    raise XLSRWException.Create('Unknown SBS formula.');
  if (Row <> PWordArray(@FSBSFormula[1])[0]) and (Col <> PWordArray(@FSBSFormula[1])[1]) then begin
    PWordArray(@FSBSFormula[1])[0] := Row;
    PWordArray(@FSBSFormula[1])[1] := Col;
    FChanged := True;
  end;
end;


procedure TShapeControlListBox.SetSelectedEntry(const Value: word);
begin
  FLBS.FSelectedEntry := Value;
end;

procedure TShapeControlListBox.SetSourceArea(Sheet, Col1, Row1, Col2, Row2: integer);
begin
  FLBS.SetSourceArea(Sheet,Col1,Row1,Col2,Row2);
  FChanged := True;
end;

procedure TShapeControlListBox.Write(Stream: TXLSStream);
var
  i,Sz: integer;
begin
  if FChanged then begin
    Sz := // CMO
          SizeOf(TObjCMO) + SizeOf(TBIFFHeader) +
          // SBS
          SizeOf(TObjSBS) + SizeOf(TBIFFHeader) +
          // LBSDATA
          FLBS.Size(False) + SizeOf(TBIFFHeader);
          // END is missing in list boxes
          if FIsComboBox then
            Inc(Sz,SizeOf(TBIFFHeader));
          // SBSFMLA
          if Length(FSBSFormula) > 0 then
            Inc(Sz,2 + 4 + Length(FSBSFormula) + SizeOf(TBIFFHeader));
          // MACRO
          if Length(FMacroFormula) > 0 then
            Inc(Sz,2 + 4 + Length(FMacroFormula) + SizeOf(TBIFFHeader));

    Stream.WriteHeader(BIFFRECID_OBJ,Sz);

    Stream.WriteHeader(OBJREC_CMO,SizeOf(TObjCMO));
    if FIsComboBox then
      Stream.WWord(OBJTYPE_COMBO_BOX)
    else
      Stream.WWord(OBJTYPE_LIST_BOX);
    Stream.WWord(FObjId);
    Stream.WWord(FOptions);
    for i := 1 to 6 do
      Stream.WWord(0);

    Stream.WriteHeader(OBJREC_SBS,SizeOf(TObjSBS));
    Stream.Write(FSBS,SizeOf(TObjSBS));

    if Length(FMacroFormula) > 0 then begin
      Stream.WriteHeader(OBJREC_MACRO,Length(FMacroFormula));
      Stream.Write(FMacroFormula[0],Length(FMacroFormula));
    end;
    if Length(FSBSFormula) > 0 then begin
      Stream.WriteHeader(OBJREC_SBSFMLA,2 + 4 + Length(FSBSFormula));
      // Last byte is not part of the formula
      Stream.WWord(Length(FSBSFormula) - 1);
      Stream.WWord($0000);
      Stream.WWord($0000);
      Stream.Write(FSBSFormula[0],Length(FSBSFormula));
    end;
    FLBS.SaveToStream(Stream,False);

    // END is missing in list boxes
    if FIsComboBox then
      // RecId = 0; Length = 0
      Stream.WLWord(0);
  end
  else begin
    Sz := // CMO
          SizeOf(TObjCMO) + SizeOf(TBIFFHeader) +
          // SBS
          SizeOf(TObjSBS) + SizeOf(TBIFFHeader) +
          // LBSDATA
          FLBS.Size(True) + SizeOf(TBIFFHeader);
          // END is missing in list boxes
          if FIsComboBox then
            Inc(Sz,SizeOf(TBIFFHeader));
          if Length(FMacroFormula) > 0 then
            // MACRO
            Inc(Sz,Length(FMacroFormula) + SizeOf(TBIFFHeader));
          if Length(FSBSFormulaOrig) > 0 then
            // SBSFMLA
            Inc(Sz,Length(FSBSFormulaOrig) + SizeOf(TBIFFHeader));
    Stream.WriteHeader(BIFFRECID_OBJ,Sz);

    Stream.WriteHeader(OBJREC_CMO,SizeOf(TObjCMO));
    if FIsComboBox then
      Stream.WWord(OBJTYPE_COMBO_BOX)
    else
      Stream.WWord(OBJTYPE_LIST_BOX);
    Stream.WWord(FObjId);
    Stream.WWord(FOptions);
    for i := 1 to 6 do
      Stream.WWord(0);

    Stream.WriteHeader(OBJREC_SBS,SizeOf(TObjSBS));
    Stream.Write(FSBS,SizeOf(TObjSBS));

    if Length(FMacroFormula) > 0 then begin
      Stream.WriteHeader(OBJREC_MACRO,Length(FMacroFormula));
      Stream.Write(FMacroFormula[0],Length(FMacroFormula));
    end;

    if Length(FSBSFormulaOrig) > 0 then begin
      Stream.WriteHeader(OBJREC_SBSFMLA,Length(FSBSFormulaOrig));
      Stream.Write(FSBSFormulaOrig[0],Length(FSBSFormulaOrig));
    end;

    FLBS.SaveToStream(Stream,True);

    // END is missing in list boxes
    if FIsComboBox then
      // RecId = 0; Length = 0
      Stream.WLWord(0);
  end;
end;

{ TShapeControlComboBox }

constructor TShapeControlComboBox.Create(ObjId: integer; ST: TMSOShapeType);
begin
  FIsComboBox := True;
  inherited Create(ObjId,ST);
end;

function TShapeControlComboBox.GetIsAutoFilter: boolean;
begin
  // Not 100% sure on how to check that a combo box belongs to an autofilter.
  // Bit 8 in Options seems to be right.
  Result := ((FOptions and $0100) = $0100) and (Length(FSBSFormula) = 0);
end;

procedure TShapeControlComboBox.SetIsAutoFilter(Value: boolean);
begin
  if not Value then
    raise XLSRWException.Create('An autofilter combo box can not be changed to a normal.');
  FOptions := $2101;
  TObjLBSComboBox(FLBS).Style := $0301;
  TObjLBSComboBox(FLBS).FComboStyle := $0002 + $0008;
  TObjLBSComboBox(FLBS).FLineCount := 8;
  SetLength(FSBSFormula,0);
end;

{ TShapeControlButton }

constructor TShapeControlButton.Create(ObjId: integer; InternalNames: TInternalNames);
begin
  inherited Create(ObjId,mstCtrlButton,InternalNames);
  FOwnsObjData := True;
  FOptions := $4001;
end;

function TShapeControlButton.ParseOBJ(OBJ: TBaseRecordStorage): boolean;
var
  i: integer;
begin
  SetLength(FMacroFormula,0);

  for i := 0 to OBJ.Count - 1 do begin
    case OBJ[i].RecId of
      OBJREC_CMO: begin
        FObjId := PObjCMO(@OBJ[i].Data).ObjId;
        FOptions := PObjCMO(@OBJ[i].Data).Options;
      end;
      OBJREC_MACRO: begin
        SetLength(FMacroFormula,OBJ[i].Length);
        Move(PByteArray(@OBJ[i].Data)[0],FMacroFormula[0],OBJ[i].Length);
      end;
      OBJREC_END: begin
      end;
      else
        raise XLSRWException.CreateFmt('Unknown record %.4X in OBJ LBS',[OBJ[i].RecId]);
    end;
  end;
  Result := True;
  FOwnsObjData := Result;
end;

procedure TShapeControlButton.Write(Stream: TXLSStream);
var
  i,Sz: integer;
begin
  Sz := // CMO
        SizeOf(TObjCMO) + SizeOf(TBIFFHeader);
        // MACRO
        if Length(FMacroFormula) > 0 then
          Inc(Sz,Length(FMacroFormula) + SizeOf(TBIFFHeader));
        // END
        Inc(Sz,SizeOf(TBIFFHeader));

  Stream.WriteHeader(BIFFRECID_OBJ,Sz);

  Stream.WriteHeader(OBJREC_CMO,SizeOf(TObjCMO));
  Stream.WWord(OBJTYPE_BUTTON);
  Stream.WWord(FObjId);
  Stream.WWord(FOptions);
  for i := 1 to 6 do
    Stream.WWord(0);

  if Length(FMacroFormula) > 0 then begin
    Stream.WriteHeader(OBJREC_MACRO,Length(FMacroFormula));
    Stream.Write(FMacroFormula[0],Length(FMacroFormula));
  end;

  Stream.WriteHeader(OBJREC_END,0);

  inherited Write(Stream);
end;

{ TShapeOutsideMsoBaseText }

procedure TShapeOutsideMsoBaseText.Assign(Shape: TShapeOutsideMso);
begin
  FreeMem(TShapeOutsideMsoBaseText(Shape).FpTXO);
  GetMem(TShapeOutsideMsoBaseText(Shape).FpTXO,SizeOf(TRecTXO) + FpTXO.ExSize);
  Move(FpTXO^,TShapeOutsideMsoBaseText(Shape).FpTXO^,SizeOf(TRecTXO) + FpTXO.ExSize);
  TShapeOutsideMsoBaseText(Shape).FText := FText;
  SetLength(TShapeOutsideMsoBaseText(Shape).FTXORuns,Length(FTXORuns));
  if Length(FTXORuns) > 0 then
    Move(Pointer(FTXORuns)^,Pointer(TShapeOutsideMsoBaseText(Shape).FTXORuns)^,SizeOf(TRecTXORUN) * Length(FTXORuns));
end;

constructor TShapeOutsideMsoBaseText.Create(ObjId: integer; ST: TMSOShapeType);
begin
  FText := '';
  GetMem(FpTXO,SizeOf(TRecTXO));
  FillChar(FpTXO^,SizeOf(TRecTXO),#0);
  inherited Create(ObjId,ST);
end;

destructor TShapeOutsideMsoBaseText.Destroy;
begin
  FreeMem(FpTXO);
  SetLength(FTXORuns,0);
  inherited;
end;

procedure TShapeOutsideMsoBaseText.Read(Stream: TXLSStream; PBuf: PByteArray);
var
  i: integer;
  FirstByte: byte;
  S,TempText: XLS8String;
  Count: integer;
  Header: TBIFFHeader;
begin
  // Read MSODRAWING
  Stream.ReadHeader(Header);
  if not Header.RecID in [BIFFRECID_MSODRAWING,BIFFRECID_CONTINUE] then
    raise XLSRWException.Create('Excpected record missing: MSODRAWING');
  Stream.Read(PBuf^,Header.Length);
//  FRecords.AddRec(Header,PBuf);

  Stream.ReadHeader(Header);
  if Header.RecID <> BIFFRECID_TXO then
    raise XLSRWException.Create('Excpected record missing: TXO');
  if FpTXO <> Nil then
    FreeMem(FpTXO);
  GetMem(FpTXO,Header.Length);
  Stream.Read(FpTXO^,Header.Length);

  if TXO.TextLen <= 0 then begin
    FText := '';
    SetLength(FTXORuns,0);
    Exit;
  end;

  // First CONTINUE (possibly more than one).
  Stream.ReadHeader(Header);
  Stream.Read(PBuf^,Header.Length);
//  FirstByte := PByte(PBuf)^;
  FirstByte := PBuf[0];
  if FirstByte = 0 then
    Count := TXO.TextLen div (MAXRECSZ_97 - 1)
  else
    Count := (TXO.TextLen * 2) div (MAXRECSZ_97 - 1);

  SetLength(TempText,Header.Length);
  System.Move(PBuf^,Pointer(TempText)^,Header.Length);

  for i := 1 to Count do begin
    Stream.ReadHeader(Header);
    Stream.Read(PBuf^,Header.Length);
//    if PByte(PBuf)^ <> FirstByte then begin
    if PBuf[0] <> FirstByte then begin
//      if (FirstByte = 1) and (PByte(PBuf)^ = 0) then begin
      if (FirstByte = 1) and (PBuf[0] = 0) then begin
        SetLength(S,(Header.Length - 1) * 2);
      end
      else
        // Don't forget Write...
        raise XLSRWException.Create('Unicode error in TXO');
    end
    else begin
      SetLength(S,Header.Length - 1);
      System.Move(Pointer(NativeInt(PBuf) + 1)^,Pointer(S)^,Header.Length - 1);
    end;
    SetLength(S,Header.Length - 1);
    System.Move(Pointer(NativeInt(PBuf) + 1)^,Pointer(S)^,Header.Length - 1);
    TempText := TempText + S;
  end;
  if TempText <> '' then begin
    if TempText[1] = #0 then begin
      TempText := Copy(TempText,2,MAXINT);
      FText := AxUCString(TempText);
    end
    else begin
      SetLength(FText,TXO.TextLen);
      TempText := Copy(TempText,2,MAXINT);
      Move(Pointer(TempText)^,Pointer(FText)^,TXO.TextLen * 2);
    end;
  end;
//   Read second CONTINUE with formatting.
  SetLength(FTXORuns,TXO.FormatLen div SizeOf(TRecTXORUN));
  Stream.ReadHeader(Header);
  Stream.Read(PBuf^,Header.Length);
  System.Move(PBuf^,Pointer(FTXORuns)^,Header.Length);
end;

procedure TShapeOutsideMsoBaseText.SetDefaultTxo;
begin
  FText := '';
  FpTXO.Options := $0212;
  FpTXO.Orientation := 0;
  FpTXO.TextLen := 0;
  FpTXO.FormatLen := 0;
  FpTXO.ExSize := 0;
  SetLength(FTXORuns,2);
  FTXORuns[0].CharIndex := 0;
  FTXORuns[0].FontIndex := 0;
  FTXORuns[0].Reserved := 0;
  FTXORuns[1].CharIndex := 0;
  FTXORuns[1].FontIndex := 0;
  FTXORuns[1].Reserved := 0;
end;

procedure TShapeOutsideMsoBaseText.SetText(const Value: AxUCString);
begin
  FText := Value;
end;

procedure TShapeOutsideMsoBaseText.SetTXORuns(const Value: TDynTXORUNArray);
var
  i: integer;
begin
  SetLength(FTXORuns,Length(Value));
  for i := 0 to High(Value) do begin
    FTXORuns[i].CharIndex := Value[i].CharIndex;
    FTXORuns[i].FontIndex := Value[i].FontIndex;
  end;
end;

procedure TShapeOutsideMsoBaseText.Write(Stream: TXLSStream);
var
  S1,S2: XLS8String;
begin
  FpTXO.TextLen := Length(FText);
  FpTXO.FormatLen := Length(FTXORuns) * SizeOf(TRecTXORUN);
  if Length(FTXORuns) > 0 then
    FTXORuns[High(FTXORuns)].CharIndex := Length(FText);

  Stream.WriteHeader(BIFFRECID_MSODRAWING,SizeOf(TMSOHeader));
  Stream.WriteMSOHeader($00,$0000,MSO_CLIENTTEXTBOX,0);

  Stream.WriteHeader(BIFFRECID_TXO,SizeOf(TRecTXO) + FpTXO.ExSize);
  Stream.Write(FpTXO^,SizeOf(TRecTXO) + FpTXO.ExSize);

  if FText <> '' then begin
    SetLength(S1,Length(FText) * 2);
    Move(Pointer(FText)^,Pointer(S1)^,Length(S1));
  end
  else
    S1 := '';

  while S1 <> '' do begin
    // Copy MAXRECSZ_97 - 2 to ensure that wide characters not are cut in two parts.
    S2 := #1 + Copy(S1,1,MAXRECSZ_97 - 2);
    Stream.WriteHeader(BIFFRECID_CONTINUE,Length(S2));
    Stream.Write(Pointer(S2)^,Length(S2));
    S1 := Copy(S1,MAXRECSZ_97 - 1,MAXINT);
  end;

  if Length(FTXORuns) > 0 then begin
    Stream.WriteHeader(BIFFRECID_CONTINUE,Length(FTXORuns) * SizeOf(TRecTXORUN));
    Stream.Write(Pointer(FTXORuns)^,Length(FTXORuns) * SizeOf(TRecTXORUN));
  end;
end;

{ TShapeOutsideMsoNote }

procedure TShapeOutsideMsoNote.Assign(Shape: TShapeOutsideMso);
begin
  inherited Assign(Shape);
  TShapeOutsideMsoNote(Shape).FCellCol := FCellCol;
  TShapeOutsideMsoNote(Shape).FCellRow := FCellRow;
  TShapeOutsideMsoNote(Shape).FOptions := FOptions;
  TShapeOutsideMsoNote(Shape).FAuthor := FAuthor;
end;

{ TObjLBS }

constructor TObjLBS.Create;
begin
  FEntryCount := 10;
  FSelectedEntry := 1;
  FStyle := $0008;
end;

procedure TObjLBS.GetSourceArea(var Sheet,Col1, Row1, Col2, Row2: integer);
begin
  if Length(FFormula) <= 0 then
    raise XLSRWException.Create('No source area assigned');
  if (Length(FFormula) = 9) and (FFormula[0] = xptgArea97) then begin
    Col1 := PPTGArea8(@FFormula[1]).Col1;
    Row1 := PPTGArea8(@FFormula[1]).Row1;
    Col2 := PPTGArea8(@FFormula[1]).Col2;
    Row2 := PPTGArea8(@FFormula[1]).Row2;
  end
  else if (Length(FFormula) = 11) and (FFormula[0] = xptgArea3d97) then begin
    Col1 := PPTGArea3d8(@FFormula[1]).Col1;
    Row1 := PPTGArea3d8(@FFormula[1]).Row1;
    Col2 := PPTGArea3d8(@FFormula[1]).Col2;
    Row2 := PPTGArea3d8(@FFormula[1]).Row2;
    Sheet := PPTGArea3d8(@FFormula[1]).Index;
  end
  else begin
{
    Col1 := 0;
    Row1 := 0;
    Col2 := 0;
    Row2 := 0;
    Sheet := -1;
}
    raise XLSRWException.Create('Unknown LBS formula.');
  end;
end;

procedure TObjLBS.GetSourceFmla(var Fmla: TDynByteArray);
begin
  Fmla := FFormula;
end;

function TObjLBS.HasSourceArea: boolean;
begin
  Result := Length(FFormula) > 0;
end;

function TObjLBS.LoadFromBuf(Buf: PByteArray; BufLen: integer): integer;
var
  Sz: integer;
  B: PByteArray;
begin
  B := Buf;
  SetLength(FLBSOrig,BufLen);
  Move(Buf[0],FLBSOrig[0],BufLen);
  Sz := PWordArray(Buf)[0];
  Buf := PByteArray(NativeInt(Buf) + 2);
  if Sz > 0 then begin
    Sz := PWordArray(Buf)[0];
    Buf := PByteArray(NativeInt(Buf) + 2);
    FUnknownFmlaData1 := PLongwordArray(Buf)[0];
    Buf := PByteArray(NativeInt(Buf) + 4);
    SetLength(FFormula,Sz);
    Move(Buf[0],FFormula[0],Sz);
    FUnknownFmlaData2 := Buf[Sz];
    Buf := PByteArray(NativeInt(Buf) + Sz + 1);
  end;
  FEntryCount := PWordArray(Buf)[0];
  FSelectedEntry := PWordArray(Buf)[1];
  FStyle := PWordArray(Buf)[2];
  FLinkedEdit := PWordArray(Buf)[3];
  Buf := PByteArray(NativeInt(Buf) + 2 + 2 + 2 + 2);
  Result := NativeInt(Buf) - NativeInt(B);
end;

procedure TObjLBS.SaveToStream(Stream: TXLSStream; SaveOrig: boolean);
begin
  // LBS must have this invalid length, otherwise excel craches.
  Stream.WriteHeader(OBJREC_LBSDATA,8158);
  if SaveOrig then
    Stream.Write(FLBSOrig[0],Length(FLBSOrig))
  else begin
    if Length(FFormula) > 0 then
      Stream.WWord(2 + 4 + Length(FFormula) + 1)
    else
      Stream.WWord(0);
    if Length(FFormula) > 0 then begin
      Stream.WWord(Length(FFormula));
      Stream.WLWord(FUnknownFmlaData1);
      Stream.Write(FFormula[0],Length(FFormula));
      Stream.WByte(FUnknownFmlaData2);
    end;
    Stream.WWord(FEntryCount);
    Stream.WWord(FSelectedEntry);
    Stream.WWord(FStyle);
    Stream.WWord(FLinkedEdit);
  end;
end;

procedure TObjLBS.SetSourceArea(Sheet,Col1, Row1, Col2, Row2: integer);
begin
  if (Length(FFormula) <> 9) or (FFormula[0] <> xptgArea97) then begin
    SetLength(FFormula,9);
    FFormula[0] := xptgArea97;
  end;
  if (PPTGArea8(@FFormula[1]).Col1 <> Col1) or (PPTGArea8(@FFormula[1]).Row1 <> Row1) or
     (PPTGArea8(@FFormula[1]).Col2 <> Col2) or (PPTGArea8(@FFormula[1]).Row2 <> Row2) then begin
    PPTGArea8(@FFormula[1]).Col1 := Col1;
    PPTGArea8(@FFormula[1]).Row1 := Row1;
    PPTGArea8(@FFormula[1]).Col2 := Col2;
    PPTGArea8(@FFormula[1]).Row2 := Row2;
  end;
end;

function TObjLBS.Size(OrigSize: boolean): integer;
begin
  if OrigSize then
    Result := Length(FLBSOrig)
  else begin
    Result := 2 + 2 + 2 + 2 + 2;
    if Length(FFormula) > 0 then
      Inc(Result,2 + 4 + Length(FFormula) + 1);
  end;
end;

{ TObjLBSComboBox }

constructor TObjLBSComboBox.Create;
begin
  FLineCount := 10;
end;

function TObjLBSComboBox.LoadFromBuf(Buf: PByteArray; BufLen: integer): integer;
begin
  Result := inherited LoadFromBuf(Buf,BufLen);
  Buf := PByteArray(NativeInt(Buf) + Result);
  FComboStyle := PWordArray(Buf)[0];
  FLineCount := PWordArray(Buf)[1];
  FUnknown := PWordArray(Buf)[2];
  Inc(Result,6);
{
  if (BufLen  - Result) <> 0 then
    raise XLSRWException.Create('OBJ LBS combo record size error');
}
end;

procedure TObjLBSComboBox.SaveToStream(Stream: TXLSStream; SaveOrig: boolean);
begin
  inherited SaveToStream(Stream,SaveOrig);
  if not SaveOrig then begin
    Stream.WWord(FComboStyle);
    Stream.WWord(FLineCount);
    Stream.WWord(FUnknown);
  end;
end;

function TObjLBSComboBox.Size(OrigSize: boolean): integer;
begin
  Result := inherited Size(OrigSize);
  if not OrigSize then
    Inc(Result,2 + 2 + 2);
end;

{ TObjLBSListBox }

function TObjLBSListBox.LoadFromBuf(Buf: PByteArray; BufLen: integer): integer;
begin
  Result := inherited LoadFromBuf(Buf,BufLen);
  Buf := PByteArray(NativeInt(Buf) + Result);
  // Multi or extended selection
  if (FStyle and $0030) <> 0 then begin
    SetLength(FMultiSelect,FEntryCount);
    Move(Buf[0],FMultiSelect[0],FEntryCount);
  end;
end;

procedure TObjLBSListBox.SaveToStream(Stream: TXLSStream; SaveOrig: boolean);
begin
  inherited SaveToStream(Stream,SaveOrig);
  if not SaveOrig and ((FStyle and $0030) <> 0) then
    Stream.Write(FMultiSelect[0],Length(FMultiSelect));
end;

function TObjLBSListBox.Size(OrigSize: boolean): integer;
begin
  Result := inherited Size(OrigSize);
  if not OrigSize and ((FStyle and $0030) <> 0) then
    Inc(Result,Length(FMultiSelect));
end;

{ TShapeOutsideMsoBaseTextMacro }

constructor TShapeOutsideMsoBaseTextMacro.Create(ObjId: integer; ST: TMSOShapeType; InternalNames: TInternalNames);
begin
  inherited Create(ObjId,ST);
  FInternalNames := InternalNames;
  FOwnsObjData := True;
  FObjId := ObjId;
  SetLength(FMacroFormula,0);
end;

function TShapeOutsideMsoBaseTextMacro.GetMacro: AxUCString;
begin
  if (Length(FMacroFormula) > 0) and (FMacroFormula[6] = xptgNameX97) then begin
    Result := FInternalNames[PPTGName8(@FMacroFormula[7]).NameIndex].Name;
  end
  else
    Result := '';
end;

procedure TShapeOutsideMsoBaseTextMacro.NameIndexChanged(Delta: integer);
begin
  //  !!!! The location of PPTGName8 in the formula can probably be elswhere.
  PPTGName8(@FMacroFormula[9]).NameIndex := PPTGName8(@FMacroFormula[9]).NameIndex + Delta;
end;

function TShapeOutsideMsoBaseTextMacro.SetMacro(Value: AxUCString): boolean;
var
  i: integer;
begin
  Result := True;
  i := FInternalNames.FindName(Value);
  if i >= 0 then begin
    SetLength(FMacroFormula,6 + 1 + SizeOf(TPTGNameX8) + 1);
    FillChar(Pointer(FMacroFormula)^,Length(FMacroFormula),#0);
    PWordArray(@FMacroFormula[0])[0] := 1 + SizeOf(TPTGNameX8);
    FMacroFormula[6] := xptgNameX97;

    PPTGNameX8(@FMacroFormula[7]).ExtSheet := 0;
    PPTGNameX8(@FMacroFormula[7]).NameIndex := i + 1;
    PPTGNameX8(@FMacroFormula[7]).Reserved := 0;
  end
  else
    raise XLSRWException.Create('Can not find Name ' + Value);
end;

{ TShapeControlCheckBox }

constructor TShapeControlCheckBox.Create(ObjId: integer; InternalNames: TInternalNames);
begin
  inherited Create(ObjId,mstCtrlCheckBox,InternalNames);
  SetLength(FCBLS,12);
  FOwnsObjData := True;
  FOptions := $0011;
end;

procedure TShapeControlCheckBox.GetDestCell(var Col, Row: integer);
begin
  if FCBLSFMLA[6] <> xptgRef97 then
    raise XLSRWException.Create('Unknown SBS formula.');
  Row := PWordArray(@FCBLSFMLA[7])[0];
  Col := PWordArray(@FCBLSFMLA[7])[1];
end;

function TShapeControlCheckBox.GetState: word;
begin
  Result := PWordArray(FCBLS)[0] and $0003;
end;

function TShapeControlCheckBox.HasDestCell: boolean;
begin
  Result := Length(FCBLSFMLA) > 0;
end;

function TShapeControlCheckBox.ParseOBJ(OBJ: TBaseRecordStorage): boolean;
var
  i: integer;
begin
  SetLength(FMacroFormula,0);
  SetLength(FCBLS,0);
  SetLength(FCBLSFMLA,0);
  SetLength(FCBLSDATA,0);

  for i := 0 to OBJ.Count - 1 do begin
    case OBJ[i].RecId of
      OBJREC_CMO: begin
        FObjId := PObjCMO(@OBJ[i].Data).ObjId;
        FOptions := PObjCMO(@OBJ[i].Data).Options;
      end;
      OBJREC_MACRO: begin
        SetLength(FMacroFormula,OBJ[i].Length);
        Move(PByteArray(@OBJ[i].Data)[0],FMacroFormula[0],OBJ[i].Length);
      end;
      OBJREC_CBLS: begin // Check box
        SetLength(FCBLS,OBJ[i].Length);
        Move(PByteArray(@OBJ[i].Data)[0],FCBLS[0],OBJ[i].Length);
      end;
      OBJREC_CBLSFMLA: begin // Check box link fmla-style macro
        SetLength(FCBLSFMLA,OBJ[i].Length);
        Move(PByteArray(@OBJ[i].Data)[0],FCBLSFMLA[0],OBJ[i].Length);
      end;
      OBJREC_CBLSDATA: begin
        SetLength(FCBLSDATA,OBJ[i].Length);
        Move(PByteArray(@OBJ[i].Data)[0],FCBLSDATA[0],OBJ[i].Length);
      end;
      OBJREC_END: begin
      end;
      else
        raise XLSRWException.CreateFmt('Unknown record %.4X in OBJ CBLS',[OBJ[i].RecId]);
    end;
  end;
  Result := True;
  FOwnsObjData := Result;
end;

procedure TShapeControlCheckBox.SetDestCell(Col, Row: integer);
begin

end;

procedure TShapeControlCheckBox.SetState(const Value: word);
begin
  PWordArray(FCBLS)[0] := PWordArray(FCBLS)[0] and $FFFC;
  PWordArray(FCBLS)[0] := PWordArray(FCBLS)[0] or Value;
end;

procedure TShapeControlCheckBox.Write(Stream: TXLSStream);
var
  i,Sz: integer;
begin
//     FCBLS: array of byte;
//     FCBLSFMLA: array of byte;
//     FCBLSDATA: array of byte;
  Sz := // CMO
        SizeOf(TObjCMO) + SizeOf(TBIFFHeader);

        Inc(Sz,GetDynArraySize(FCBLS));
        Inc(Sz,GetDynArraySize(FMacroFormula));
        Inc(Sz,GetDynArraySize(FCBLSFMLA));
        Inc(Sz,GetDynArraySize(FCBLSDATA));
        // END
        Inc(Sz,SizeOf(TBIFFHeader));

  Stream.WriteHeader(BIFFRECID_OBJ,Sz);

  Stream.WriteHeader(OBJREC_CMO,SizeOf(TObjCMO));
  Stream.WWord(OBJTYPE_CHECK_BOX);
  Stream.WWord(FObjId);
  Stream.WWord(FOptions);
  for i := 1 to 6 do
    Stream.WWord(0);

  WriteDynArray(Stream,OBJREC_CBLS,FCBLS);
  WriteDynArray(Stream,OBJREC_MACRO,FMacroFormula);
  WriteDynArray(Stream,OBJREC_CBLSFMLA,FCBLSFMLA);
  WriteDynArray(Stream,OBJREC_CBLSDATA,FCBLSDATA);

  Stream.WriteHeader(OBJREC_END,0);

  inherited Write(Stream);
end;

{ TShapeControlRadioButton }

constructor TShapeControlRadioButton.Create(ObjId: integer; InternalNames: TInternalNames);
begin
  inherited Create(ObjId,mstCtrlRadioButton,InternalNames);
  SetLength(FCBLS,12);
  FOwnsObjData := True;
  FOptions := $0011;
end;

procedure TShapeControlRadioButton.GetDestCell(var Col, Row: integer);
begin
  if FCBLSFMLA[6] <> xptgRef97 then
    raise XLSRWException.Create('Unknown SBS formula.');
  Row := PWordArray(@FCBLSFMLA[7])[0];
  Col := PWordArray(@FCBLSFMLA[7])[1];
end;

function TShapeControlRadioButton.GetState: word;
begin
  Result := PWordArray(FCBLS)[0] and $0003;
end;

function TShapeControlRadioButton.HasDestCell: boolean;
begin
  Result := Length(FCBLSFMLA) > 0;
end;

function TShapeControlRadioButton.ParseOBJ(OBJ: TBaseRecordStorage): boolean;
var
  i: integer;
begin
  SetLength(FMacroFormula,0);
  SetLength(FCBLS,0);
  SetLength(FCBLSFMLA,0);
  SetLength(FCBLSDATA,0);
  SetLength(FRBO,0);
  SetLength(FRBODATA,0);

  for i := 0 to OBJ.Count - 1 do begin
    case OBJ[i].RecId of
      OBJREC_CMO: begin
        FObjId := PObjCMO(@OBJ[i].Data).ObjId;
        FOptions := PObjCMO(@OBJ[i].Data).Options;
      end;
      OBJREC_MACRO: begin
        SetLength(FMacroFormula,OBJ[i].Length);
        Move(PByteArray(@OBJ[i].Data)[0],FMacroFormula[0],OBJ[i].Length);
      end;
      OBJREC_CBLS: begin // Check box link
        SetLength(FCBLS,OBJ[i].Length);
        Move(PByteArray(@OBJ[i].Data)[0],FCBLS[0],OBJ[i].Length);
      end;
      OBJREC_CBLSFMLA: begin // Check box link fmla-style macro
        SetLength(FCBLSFMLA,OBJ[i].Length);
        Move(PByteArray(@OBJ[i].Data)[0],FCBLSFMLA[0],OBJ[i].Length);
      end;
      OBJREC_CBLSDATA: begin
        SetLength(FCBLSDATA,OBJ[i].Length);
        Move(PByteArray(@OBJ[i].Data)[0],FCBLSDATA[0],OBJ[i].Length);
      end;
      OBJREC_RBO: begin // Radio button
        SetLength(FRBO,OBJ[i].Length);
        Move(PByteArray(@OBJ[i].Data)[0],FRBO[0],OBJ[i].Length);
      end;
      OBJREC_RBODATA: begin // Radio button data
        SetLength(FRBODATA,OBJ[i].Length);
        Move(PByteArray(@OBJ[i].Data)[0],FRBODATA[0],OBJ[i].Length);
      end;
      OBJREC_END: begin
      end;
      else
        raise XLSRWException.CreateFmt('Unknown record %.4X in OBJ RBO',[OBJ[i].RecId]);
    end;
  end;
  Result := True;
  FOwnsObjData := Result;
end;

procedure TShapeControlRadioButton.SetDestCell(Col, Row: integer);
begin

end;

procedure TShapeControlRadioButton.SetState(const Value: word);
begin
  PWordArray(FCBLS)[0] := PWordArray(FCBLS)[0] and $FFFC;
  PWordArray(FCBLS)[0] := PWordArray(FCBLS)[0] or Value;
end;

procedure TShapeControlRadioButton.Write(Stream: TXLSStream);
var
  i,Sz: integer;
begin
  Sz := // CMO
        SizeOf(TObjCMO) + SizeOf(TBIFFHeader);

        Inc(Sz,GetDynArraySize(FCBLS));
        Inc(Sz,GetDynArraySize(FRBO));
        Inc(Sz,GetDynArraySize(FMacroFormula));
        Inc(Sz,GetDynArraySize(FCBLSFMLA));
        Inc(Sz,GetDynArraySize(FCBLSDATA));
        Inc(Sz,GetDynArraySize(FRBODATA));
        // END
        Inc(Sz,SizeOf(TBIFFHeader));

  Stream.WriteHeader(BIFFRECID_OBJ,Sz);

  Stream.WriteHeader(OBJREC_CMO,SizeOf(TObjCMO));
  Stream.WWord(OBJTYPE_OPTION_BUTTON);
  Stream.WWord(FObjId);
  Stream.WWord(FOptions);
  for i := 1 to 6 do
    Stream.WWord(0);

  WriteDynArray(Stream,OBJREC_CBLS,FCBLS);
  WriteDynArray(Stream,OBJREC_RBO,FRBO);
  WriteDynArray(Stream,OBJREC_MACRO,FMacroFormula);
  WriteDynArray(Stream,OBJREC_CBLSFMLA,FCBLSFMLA);
  WriteDynArray(Stream,OBJREC_CBLSDATA,FCBLSDATA);
  WriteDynArray(Stream,OBJREC_RBODATA,FRBODATA);

  Stream.WriteHeader(OBJREC_END,0);

  inherited Write(Stream);
end;

{ TShapeControlScrollBar }

procedure TShapeControlScrollBar.Assign(Shape: TShapeOutsideMso);
begin

end;

constructor TShapeControlScrollBar.Create(ObjId: integer);
begin
  inherited Create(ObjId,mstCtrlScrollBar);
  SetLength(FSBS,20);
  FOwnsObjData := True;
  FOptions := $6011;
end;

procedure TShapeControlScrollBar.GetDestCell(var Col, Row: integer);
begin
  if FSBSFormula[0] <> xptgRef97 then
    raise XLSRWException.Create('Unknown SBS formula.');
  Row := PWordArray(@FSBSFormula[1])[0];
  Col := PWordArray(@FSBSFormula[1])[1];
end;

function TShapeControlScrollBar.GetIncChange: integer;
begin
  Result := PWordArray(FSBS)[5];
end;

function TShapeControlScrollBar.GetMaxValue: integer;
begin
  Result := PWordArray(FSBS)[4];
end;

function TShapeControlScrollBar.GetMinValue: integer;
begin
  Result := PWordArray(FSBS)[3];
end;

function TShapeControlScrollBar.GetPageChange: integer;
begin
  Result := PWordArray(FSBS)[6];
end;

function TShapeControlScrollBar.GetValue: integer;
begin
  Result := PWordArray(FSBS)[2];
end;

function TShapeControlScrollBar.ParseOBJ(OBJ: TBaseRecordStorage): boolean;
var
  i: integer;
begin
  SetLength(FMacroFormula,0);
  SetLength(FSBS,0);
  SetLength(FSBSFormula,0);
  SetLength(FSBSFormulaOrig,0);

  for i := 0 to OBJ.Count - 1 do begin
    case OBJ[i].RecId of
      OBJREC_CMO: begin
        FObjId := PObjCMO(@OBJ[i].Data).ObjId;
        FOptions := PObjCMO(@OBJ[i].Data).Options;
      end;
      OBJREC_MACRO: begin
        SetLength(FMacroFormula,OBJ[i].Length);
        Move(PByteArray(@OBJ[i].Data)[0],FMacroFormula[0],OBJ[i].Length);
      end;
      OBJREC_SBS: begin
        SetLength(FSBS,OBJ[i].Length);
        Move(PByteArray(@OBJ[i].Data)[0],FSBS[0],OBJ[i].Length);
      end;
      OBJREC_SBSFMLA: begin
        SetLength(FSBSFormulaOrig,OBJ[i].Length);
        Move(PByteArray(@OBJ[i].Data)[0],FSBSFormulaOrig[0],OBJ[i].Length);
        if PWordArray(@OBJ[i].Data)[0] = 5 then begin
          SetLength(FSBSFormula,PWordArray(@OBJ[i].Data)[0] + 1);
          Move(PByteArray(@OBJ[i].Data)[6],FSBSFormula[0],PWordArray(@OBJ[i].Data)[0] + 1);
        end;
      end;
      OBJREC_END: begin
      end;
      else
        raise XLSRWException.CreateFmt('Unknown record %.4X in OBJ CBLS',[OBJ[i].RecId]);
    end;
  end;
  Result := True;
  FOwnsObjData := Result;
end;

procedure TShapeControlScrollBar.Read(Stream: TXLSStream; PBuf: PByteArray);
begin

end;

procedure TShapeControlScrollBar.SetIncChange(const Value: integer);
begin
  PWordArray(FSBS)[5] := Value;
end;

procedure TShapeControlScrollBar.SetMaxValue(const Value: integer);
begin
  PWordArray(FSBS)[4] := Value;
end;

procedure TShapeControlScrollBar.SetMinValue(const Value: integer);
begin
  PWordArray(FSBS)[3] := Value;
end;

procedure TShapeControlScrollBar.SetPageChange(const Value: integer);
begin
  PWordArray(FSBS)[6] := Value;
end;

procedure TShapeControlScrollBar.SetValue(const Value: integer);
begin
  PWordArray(FSBS)[2] := Value;
end;

procedure TShapeControlScrollBar.Write(Stream: TXLSStream);
var
  i,Sz: integer;
begin
  Sz := // CMO
        SizeOf(TObjCMO) + SizeOf(TBIFFHeader);

        Inc(Sz,GetDynArraySize(FSBS));
        Inc(Sz,GetDynArraySize(FMacroFormula));
        Inc(Sz,GetDynArraySize(FSBSFormulaOrig));
        // END
        Inc(Sz,SizeOf(TBIFFHeader));

  Stream.WriteHeader(BIFFRECID_OBJ,Sz);

  Stream.WriteHeader(OBJREC_CMO,SizeOf(TObjCMO));
  Stream.WWord(OBJTYPE_SCROLL_BAR);
  Stream.WWord(FObjId);
  Stream.WWord(FOptions);
  for i := 1 to 6 do
    Stream.WWord(0);

  WriteDynArray(Stream,OBJREC_SBS,FSBS);
  WriteDynArray(Stream,OBJREC_MACRO,FMacroFormula);
  WriteDynArray(Stream,OBJREC_SBSFMLA,FSBSFormulaOrig);

  Stream.WriteHeader(OBJREC_END,0);
end;

end.
