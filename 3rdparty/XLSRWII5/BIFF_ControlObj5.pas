unit BIFF_ControlObj5;

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
     BIFF_DrawingObjAnchor5, BIFF_RecsII5, BIFF_Escher5, BIFF_EscherCtrlObj5,
     BIFF_Stream5, BIFF_FormulaHandler5,
     Xc12Utils5,
     XLSUtils5;

//* ~exclude
type TCtrlCheckState = (ccsUnchecked,ccsChecked,cchTristate);

//* ~exclude
type TControlObjList = class(TCollection)
private
     function GetItems(Index: integer): TDrwAnchor;
protected
     FOwner: TPersistent;
     FDrawing: TEscherDrawing;
     FFileAdd: boolean;

     function  GetOwner: TPersistent; override;
public
     //* ~exclude
     constructor Create(AOwner: TPersistent; Drawing: TEscherDrawing; ItemClass: TCollectionItemClass);
     //* ~exclude
     procedure Delete(Index: integer);
     //* ~exclude
     procedure DeleteById(Id: integer);

     property Items[Index: integer]: TDrwAnchor read GetItems; default;
     end;


//* A list box control. The list box is filled with the cell values given in
//* SourceArea. When one of the items in the list box is selected, the cell in
//* DestCell is assigned this value.
    TControlListBox = class(TDrwAnchor)
private
     function GetDestCell: AxUCString;
     function GetSourceArea: AxUCString;
     procedure SetDestCell(const Value: AxUCString);
     procedure SetSourceArea(const Value: AxUCString);
protected
public
     //* ~exclude
     constructor Create(Collection: TCollection); override;
published
     //* Area of cells that fills the list in the list box.
     property SourceArea: AxUCString read GetSourceArea write SetSourceArea;
     //* The destination cell that is assigned the selected line in the list
     //* box. The cell reference must be absolute, as: 'Sheet1!@A$1'
     property DestCell: AxUCString read GetDestCell write SetDestCell;
     end;

type
//* A list of TControlListBox objects.
    TControlListBoxes = class(TControlObjList)
private
     function  GetControlListBox(Index: integer): TControlListBox;
protected
public
     //* ~exclude
     constructor Create(AOwner: TPersistent; Drawing: TEscherDrawing);
     //* Add a new TControlListBox.
     //* ~result The new TControlListBox object.
     function  Add: TControlListBox;
     //* ~exclude
     procedure Clear;
     //* ~exclude
     procedure AddFromFile(Shape: TShapeClientAnchor);

     //* The TControlListBox objects in the list.
     property Items[Index: integer]: TControlListBox read GetControlListBox; default;
     end;

type
//* A combo box control. The combo box is filled with the cell values given in
//* SourceArea. When one of the items in the combo box is selected, the cell in
//* DestCell is assigned the index value of the selected item.
    TControlComboBox = class(TDrwAnchor)
private
     function GetDestCell: AxUCString;
     function GetSourceArea: AxUCString;
     procedure SetDestCell(const Value: AxUCString);
     procedure SetSourceArea(const Value: AxUCString);
protected
public
     //* ~exclude
     constructor Create(Collection: TCollection); override;
     //* ~exclude
//     constructor CreateAutoFilter(Collection: TCollection);
published
     //* Area of cells that fills the list in the conbo box.
     property SourceArea: AxUCString read GetSourceArea write SetSourceArea;
     //* The destination cell that is assigned the selected line in the conbo
     //* box.
     property DestCell: AxUCString read GetDestCell write SetDestCell;
     end;

type
//* A list of TControlComboBox objects.
    TControlComboBoxes = class(TControlObjList)
private
     function  GetControlComboBox(Index: integer): TControlComboBox;
protected
     FAddAutoFilter: boolean;
public
     //* ~exclude
     constructor Create(AOwner: TPersistent; Drawing: TEscherDrawing);
     //* Add a new TControlComboBox.
     //* ~result The new TControlComboBox object.
     function  Add: TControlComboBox;
     //* ~exclude
     function  AddAutoFilter: TControlComboBox;
     //* ~exclude
     procedure Clear;
     //* ~exclude
     procedure AddFromFile(Shape: TShapeClientAnchor);

     //* The TControlComboBox objects in the list.
     property Items[Index: integer]: TControlComboBox read GetControlComboBox; default;
     end;

//* Base class fo control objects with texts, such as buttons, check boxes etc.
type TControlText = class(TDrwAnchor)
private
     function  GetText: AxUCString;
     procedure SetText(const Value: AxUCString);
public
published
     //* Text on the object.
     property Text: AxUCString read GetText write SetText;
     end;

//* Base class fo control objects that can execute VBA macros, such as buttons etc.
//* ~[br]
//* ~[br]
//* ~[b Example:]
//* ~[br]
//* ~[sample
//*  // Button with macro example.
//*  // IMPORTANT! As XLSReadWriteII not can check if the syntax of the macro
//*  // code is correct, try to test run it in Excel first.
//*  // If there are errors in the code, Excel will open the file with an
//*  // error message.
//*  EditVBA must be set to True in order to edit macros.
//*  XLS.VBA.EditVBA := True;
//*  // Add a VBA module and a procedure.
//*  with XLS.VBA.AddModule('Module1',vmtMacro) do begin
//*    Source.Text :=
//*    'Sub MyTestMacro()' + #13#10 +
//*    'MsgBox "MyTestMacro executed"' + #13#10 +
//*    'End Sub';
//*  end;
//*
//*  // Add a button control.
//*  with XLS.Sheet[0~[].ControlsObjects.Buttons.Add do begin
//*    Text := 'Macro test';
//*    Col1 := 2;
//*    Row1 := 4;
//*    Col2 := 3;
//*    // Col2Offset is the distance in percent of the
//*    // following column that the control occupies.
//*    Col2Offset := 0.5;
//*    Row2 := 5;
//*    // Row2Offset is the distance in percent of the
//*    // following row that the control occupies.
//*    Row2Offset := 0.25;
//*    // The name of the macro must match the previous added macro procedure.
//*    MacroName := 'MyTestMacro';
//*  end;
//* ]
type TControlTextMacro = class(TControlText)
private
     function  GetMacroName: AxUCString;
     procedure SetMacroName(const Value: AxUCString);
public
published
     //* The name of the macro that is executed when the button is clicked.
     //* Use MacroName to assign an allready existing macro to the control.
     //* Add new macros with ~[link XLSReadWriteII4.TXLSReadWriteII4.VBA TXLSReadWriteII4.VBA]
     property MacroName: AxUCString read GetMacroName write SetMacroName;
     end;

type
//* A button object. A button can have a macro assigned to it, that is executed
//* when the button is clicked. See ~[link TControlTextMacro] for a macro example.
    TControlButton = class(TControlTextMacro)
private
public
     //* ~exclude
     constructor Create(Collection: TCollection); override;
     end;

type
//* A list of TControlButton objects.
    TControlButtons = class(TControlObjList)
private
     function  GetControlButton(Index: integer): TControlButton;
protected
public
     //* ~exclude
     constructor Create(AOwner: TPersistent; Drawing: TEscherDrawing);
     //* Add a new TControlButton.
     //* ~result The new TControlButton object.
     function  Add: TControlButton;
     //* ~exclude
     procedure Clear;
     //* ~exclude
     procedure AddFromFile(Shape: TShapeClientAnchor);

     //* The TControlButton objects in the list.
     property Items[Index: integer]: TControlButton read GetControlButton; default;
     end;

//* A check box object. A check box can have a macro assigned to it, that is executed
//* when the button is clicked. See ~[link TControlTextMacro] for a macro example.
type TControlCheckBox = class(TControlTextMacro)
private
     function  GetState: TCtrlCheckState;
     procedure SetState(const Value: TCtrlCheckState);
public
     //* ~exclude
     constructor Create(Collection: TCollection); override;
     //* The check state of the object.
     property State: TCtrlCheckState read GetState write SetState;
     end;

type
//* A list of TControlCheckBoxes objects.
    TControlCheckBoxes = class(TControlObjList)
private
     function  GetControlCheckBox(Index: integer): TControlCheckBox;
protected
public
     //* ~exclude
     constructor Create(AOwner: TPersistent; Drawing: TEscherDrawing);
     //* Add a new TControlCheckBox.
     //* ~result The new TControlCheckBox object.
     function  Add: TControlCheckBox;
     //* ~exclude
     procedure Clear;
     //* ~exclude
     procedure AddFromFile(Shape: TShapeClientAnchor);

     //* The TControlButton objects in the list.
     property Items[Index: integer]: TControlCheckBox read GetControlCheckBox; default;
     end;

//* A Radio button object. A check box can have a macro assigned to it, that is executed
//* when the button is clicked. See ~[link TControlTextMacro] for a macro example.
type TControlRadioButton = class(TControlTextMacro)
private
     function  GetState: TCtrlCheckState;
     procedure SetState(const Value: TCtrlCheckState);
public
     //* ~exclude
     constructor Create(Collection: TCollection); override;
     //* The check state of the object.
     property State: TCtrlCheckState read GetState write SetState;
     end;

type
//* A list of TControlRadioButtons objects.
    TControlRadioButtons = class(TControlObjList)
private
     function  GetControlRadioButton(Index: integer): TControlRadioButton;
protected
public
     //* ~exclude
     constructor Create(AOwner: TPersistent; Drawing: TEscherDrawing);
     //* Add a new TControlRadioButton.
     //* ~result The new TControlRadioButton object.
     function  Add: TControlRadioButton;
     //* ~exclude
     procedure Clear;
     //* ~exclude
     procedure AddFromFile(Shape: TShapeClientAnchor);

     //* The TControlButton objects in the list.
     property Items[Index: integer]: TControlRadioButton read GetControlRadioButton; default;
     end;

type
// Class with lists of all control objects.
   TControlObjects = class(TPersistent)
private
     FOwner: TPersistent;
     FDrawing: TEscherDrawing;
     FormulaHandler: TFormulaHandler;

     FListBoxes: TControlListBoxes;
     FComboBoxes: TControlComboBoxes;
     FButtons: TControlButtons;
     FCheckBoxes: TControlCheckBoxes;
     FRadioButtons: TControlRadioButtons;
protected
     function  GetOwner: TPersistent; override;
public
     //* ~exclude
     constructor Create(AOwner: TPersistent; Drawing: TEscherDrawing; FmlaHandler: TFormulaHandler);
     //* ~exclude
     destructor Destroy; override;
     //* Clear all control objects.
     procedure Clear;
published
     //* List box objects.
     property ListBoxes: TControlListBoxes read FListBoxes write FListBoxes;
     //* Combo box objects.
     property ComboBoxes: TControlComboBoxes read FComboBoxes write FComboBoxes;
     //* Button objects.
     property Buttons: TControlButtons read FButtons write FButtons;
     //* Check box objects.
     property CheckBoxes: TControlCheckBoxes read FCheckBoxes write FCheckBoxes;
     //* Radio button objects.
     property RadioButtons: TControlRadioButtons read FRadioButtons write FRadioButtons;
     end;

implementation

{ TControlListBox }

constructor TControlListBox.Create(Collection: TCollection);
begin
  if not TControlListBoxes(Collection).FFileAdd then
    FShape := TControlListBoxes(Collection).FDrawing.AddListBox;
  inherited Create(Collection);
  FName := 'ListBox ' + IntToStr(ID);
end;

function TControlListBox.GetDestCell: AxUCString;
var
  {Sht,}Col,Row: integer;
begin
  TShapeControlListBox(FShape.ExShape).GetDestCell(Col,Row);
  Result := ColRowToRefStr(Col,Row,True,True);
end;

function TControlListBox.GetSourceArea: AxUCString;
var
  Sheet,Col1,Row1,Col2,Row2: integer;
begin
  TShapeControlListBox(FShape.ExShape).GetSourceArea(Sheet,Col1,Row1,Col2,Row2);
  Result := AreaToRefStr(Col1,Row1,Col2,Row2,True,True,True,True);
end;

procedure TControlListBox.SetDestCell(const Value: AxUCString);
var
  Col,Row: integer;
  AbsCol,AbsRow: boolean;
begin
  if not RefStrToColRow(Value,Col,Row,AbsCol,AbsRow) then
    raise XLSRWException.Create('Invalid cell reference');
  TShapeControlListBox(FShape.ExShape).SetDestCell(Col,Row);
end;

procedure TControlListBox.SetSourceArea(const Value: AxUCString);
var
  Col1,Row1,Col2,Row2: integer;
  AbsCol1,AbsRow1,AbsCol2,AbsRow2: boolean;
begin
  if not AreaStrToColRow(Value,Col1,Row1,Col2,Row2,AbsCol1,AbsRow1,AbsCol2,AbsRow2) then
    raise XLSRWException.Create('Invalid area reference');
  TShapeControlListBox(FShape.ExShape).SetSourceArea(-1,Col1,Row1,Col2,Row2);
end;

{ TControlListBoxes }

function TControlListBoxes.Add: TControlListBox;
begin
  Result := TControlListBox(inherited Add);
end;

procedure TControlListBoxes.AddFromFile(Shape: TShapeClientAnchor);
begin
  FFileAdd := True;
  try
    TControlListBox(inherited Add).FShape := Shape;
  finally
    FFileAdd := False;
  end;
end;

procedure TControlListBoxes.Clear;
var
  i: integer;
begin
  for i := 0 to Count - 1 do
    FDrawing.DeleteShape(Items[i].FShape.SpId);
  inherited Clear;
end;

constructor TControlListBoxes.Create(AOwner: TPersistent; Drawing: TEscherDrawing);
begin
  inherited Create(AOwner,Drawing,TControlListBox);
end;

function TControlListBoxes.GetControlListBox(Index: integer): TControlListBox;
begin
  Result := TControlListBox(inherited Items[Index]);
end;

{ TControlObjects }

procedure TControlObjects.Clear;
begin
  FListBoxes.Clear;
  FButtons.Clear;
end;

constructor TControlObjects.Create(AOwner: TPersistent; Drawing: TEscherDrawing; FmlaHandler: TFormulaHandler);
begin
  FOwner := AOwner;
  FDrawing := Drawing;
  FormulaHandler := FmlaHandler;
  FListBoxes := TControlListBoxes.Create(Self,FDrawing);
  FComboBoxes := TControlComboBoxes.Create(Self,Drawing);
  FButtons := TControlButtons.Create(Self,FDrawing);
  FCheckBoxes := TControlCheckBoxes.Create(Self,FDrawing);
  FRadioButtons := TControlRadioButtons.Create(Self,FDrawing);
end;

destructor TControlObjects.Destroy;
begin
  FRadioButtons.Free;
  FCheckBoxes.Free;
  FButtons.Free;
  FComboBoxes.Free;
  FListBoxes.Free;
  inherited;
end;

function TControlObjects.GetOwner: TPersistent;
begin
  Result := FOwner;
end;

{ TControlButton }

constructor TControlButton.Create(Collection: TCollection);
begin
  if not TControlButtons(Collection).FFileAdd then
    FShape := TControlButtons(Collection).FDrawing.AddButton;
  inherited Create(Collection);
  FName := 'Button ' + IntToStr(ID);
end;

{ TControlButtons }

function TControlButtons.Add: TControlButton;
begin
  Result := TControlButton(inherited Add);
end;

procedure TControlButtons.AddFromFile(Shape: TShapeClientAnchor);
begin
  FFileAdd := True;
  try
    TControlButton(inherited Add).FShape := Shape;
  finally
    FFileAdd := False;
  end;
end;

procedure TControlButtons.Clear;
var
  i: integer;
begin
  for i := 0 to Count - 1 do
    FDrawing.DeleteShape(Items[i].FShape.SpId);
  inherited Clear;
end;

constructor TControlButtons.Create(AOwner: TPersistent; Drawing: TEscherDrawing);
begin
  inherited Create(AOwner,Drawing,TControlButton);
  FOwner := AOwner;
  FDrawing := Drawing;
end;

function TControlButtons.GetControlButton(Index: integer): TControlButton;
begin
  Result := TControlButton(inherited Items[Index]);
end;

{ TControlComboBox }

constructor TControlComboBox.Create(Collection: TCollection);
begin
  if TControlComboBoxes(Collection).FAddAutoFilter then begin
    FShape := TControlComboBoxes(Collection).FDrawing.AddComboBoxAutoFilter;
    inherited Create(Collection);
    FName := 'ComboBox Autofilter ' + IntToStr(ID);
//    CreateAutoFilter(Collection)
  end
  else begin
    if not TControlComboBoxes(Collection).FFileAdd then
      FShape := TControlComboBoxes(Collection).FDrawing.AddComboBox;
    inherited Create(Collection);
    FName := 'ComboBox ' + IntToStr(ID);
  end;
end;

{
constructor TControlComboBox.CreateAutoFilter(Collection: TCollection);
begin
  FShape := TControlComboBoxes(Collection).FDrawing.AddComboBoxAutoFilter;
  inherited Create(Collection);
  FName := 'ComboBox Autofilter ' + IntToStr(ID);
end;
}

function TControlComboBox.GetDestCell: AxUCString;
var
  {Sht,}Col,Row: integer;
begin
  TShapeControlComboBox(FShape.ExShape).GetDestCell(Col,Row);
  Result := ColRowToRefStr(Col,Row,True,True);
end;

function TControlComboBox.GetSourceArea: AxUCString;
var
  Sheet,Col1,Row1,Col2,Row2: integer;
begin
  TShapeControlComboBox(FShape.ExShape).GetSourceArea(Sheet,Col1,Row1,Col2,Row2);
  Result := AreaToRefStr(Col1,Row1,Col2,Row2,True,True,True,True);
end;

procedure TControlComboBox.SetDestCell(const Value: AxUCString);
var
  Col,Row: integer;
  AbsCol,AbsRow: boolean;
begin
  if not RefStrToColRow(Value,Col,Row,AbsCol,AbsRow) then
    raise XLSRWException.Create('Invalid cell reference');
  TShapeControlComboBox(FShape.ExShape).SetDestCell(Col,Row);
end;

procedure TControlComboBox.SetSourceArea(const Value: AxUCString);
var
  Col1,Row1,Col2,Row2: integer;
  AbsCol1,AbsRow1,AbsCol2,AbsRow2: boolean;
begin
  if not AreaStrToColRow(Value,Col1,Row1,Col2,Row2,AbsCol1,AbsRow1,AbsCol2,AbsRow2) then
    raise XLSRWException.Create('Invalid area reference');
  TShapeControlComboBox(FShape.ExShape).SetSourceArea(-1,Col1,Row1,Col2,Row2);
end;

{ TControlComboBoxes }

function TControlComboBoxes.Add: TControlComboBox;
begin
  FAddAutoFilter := False;
  Result := TControlComboBox(inherited Add);
end;

function TControlComboBoxes.AddAutoFilter: TControlComboBox;
begin
  FAddAutoFilter := True;
  Result := TControlComboBox(inherited Add);
end;

procedure TControlComboBoxes.AddFromFile(Shape: TShapeClientAnchor);
begin
  FFileAdd := True;
  try
    TControlComboBox(inherited Add).FShape := Shape;
  finally
    FFileAdd := False;
  end;
end;

procedure TControlComboBoxes.Clear;
var
  i: integer;
begin
  for i := 0 to Count - 1 do
    FDrawing.DeleteShape(Items[i].FShape.SpId);
  inherited Clear;
end;

constructor TControlComboBoxes.Create(AOwner: TPersistent; Drawing: TEscherDrawing);
begin
  inherited Create(AOwner,Drawing,TControlComboBox);
  FOwner := AOwner;
  FDrawing := Drawing;
end;

function TControlComboBoxes.GetControlComboBox(Index: integer): TControlComboBox;
begin
  Result := TControlComboBox(inherited Items[Index]);
end;

{ TControlObjList }

constructor TControlObjList.Create(AOwner: TPersistent; Drawing: TEscherDrawing; ItemClass: TCollectionItemClass);
begin
  inherited Create(ItemClass);
  FOwner := AOwner;
  FDrawing := Drawing;
end;

procedure TControlObjList.Delete(Index: integer);
begin
  FDrawing.DeleteShape(Items[Index].SpId);
  inherited Delete(Index);
end;

procedure TControlObjList.DeleteById(Id: integer);
var
  i: integer;
begin
  for i := 0 to Count - 1 do begin
    if Items[i].Id = Id then begin
      FDrawing.DeleteShape(Items[i].SpId);
      Delete(i);
      Exit;
    end;
  end;
end;

function TControlObjList.GetItems(Index: integer): TDrwAnchor;
begin
  Result := TDrwAnchor(inherited Items[Index]);
end;

function TControlObjList.GetOwner: TPersistent;
begin
  Result := FOwner;
end;

{ TControlText }

function TControlText.GetText: AxUCString;
begin
  Result := TShapeControlButton(FShape.ExShape).Text;
end;

procedure TControlText.SetText(const Value: AxUCString);
begin
  TShapeControlButton(FShape.ExShape).Text := Value;
end;

{ TControlTextMacro }

function TControlTextMacro.GetMacroName: AxUCString;
begin
  Result := TShapeControlButton(FShape.ExShape).GetMacro;
end;

procedure TControlTextMacro.SetMacroName(const Value: AxUCString);
var
  FH: TFormulaHandler;
begin
  FH := TControlObjects(TControlButtons(Collection).FOwner).FormulaHandler;
  with FH.InternalNames.Add do begin
    Name := Value;
    IsVBProcedure := True;
  end;
  FH.ExternalNames.AddSelf($FFFE,1);

  TShapeControlButton(FShape.ExShape).SetMacro(Value);
end;

{ TControlCheckBox }

constructor TControlCheckBox.Create(Collection: TCollection);
begin
  if not TControlCheckBoxes(Collection).FFileAdd then
    FShape := TControlCheckBoxes(Collection).FDrawing.AddCheckBox;
  inherited Create(Collection);
  FName := 'CheckBox ' + IntToStr(ID);
end;

function TControlCheckBox.GetState: TCtrlCheckState;
begin
  Result := TCtrlCheckState(TShapeControlCheckBox(FShape.ExShape).State);
end;

procedure TControlCheckBox.SetState(const Value: TCtrlCheckState);
begin
  TShapeControlCheckBox(FShape.ExShape).State := Word(Value);
end;

{ TControlCheckBoxes }

function TControlCheckBoxes.Add: TControlCheckBox;
begin
  Result := TControlCheckBox(inherited Add);
end;

procedure TControlCheckBoxes.AddFromFile(Shape: TShapeClientAnchor);
begin
  FFileAdd := True;
  try
    TControlCheckBox(inherited Add).FShape := Shape;
  finally
    FFileAdd := False;
  end;
end;

procedure TControlCheckBoxes.Clear;
var
  i: integer;
begin
  for i := 0 to Count - 1 do
    FDrawing.DeleteShape(Items[i].FShape.SpId);
  inherited Clear;
end;

constructor TControlCheckBoxes.Create(AOwner: TPersistent; Drawing: TEscherDrawing);
begin
  inherited Create(AOwner,Drawing,TControlCheckBox);
  FOwner := AOwner;
  FDrawing := Drawing;
end;

function TControlCheckBoxes.GetControlCheckBox(Index: integer): TControlCheckBox;
begin
  Result := TControlCheckBox(inherited Items[Index]);
end;

{ TControlRadioButton }

constructor TControlRadioButton.Create(Collection: TCollection);
begin
  if not TControlRadioButtons(Collection).FFileAdd then
    FShape := TControlRadioButtons(Collection).FDrawing.AddRadioButton;
  inherited Create(Collection);
  FName := 'RadioButton ' + IntToStr(ID);
end;

function TControlRadioButton.GetState: TCtrlCheckState;
begin
  Result := TCtrlCheckState(TShapeControlRadioButton(FShape.ExShape).State);
end;

procedure TControlRadioButton.SetState(const Value: TCtrlCheckState);
begin
  TShapeControlRadioButton(FShape.ExShape).State := Word(Value);
end;

{ TControlRadioButtons }

function TControlRadioButtons.Add: TControlRadioButton;
begin
  Result := TControlRadioButton(inherited Add);
end;

procedure TControlRadioButtons.AddFromFile(Shape: TShapeClientAnchor);
begin
  FFileAdd := True;
  try
    TControlRadioButton(inherited Add).FShape := Shape;
  finally
    FFileAdd := False;
  end;
end;

procedure TControlRadioButtons.Clear;
var
  i: integer;
begin
  for i := 0 to Count - 1 do
    FDrawing.DeleteShape(Items[i].FShape.SpId);
  inherited Clear;
end;

constructor TControlRadioButtons.Create(AOwner: TPersistent; Drawing: TEscherDrawing);
begin
  inherited Create(AOwner,Drawing,TControlRadioButton);
  FOwner := AOwner;
  FDrawing := Drawing;
end;

function TControlRadioButtons.GetControlRadioButton(Index: integer): TControlRadioButton;
begin
  Result := TControlRadioButton(inherited Items[Index]);
end;

end.
