unit XSSIEEditor;

interface

uses Classes, SysUtils, Contnrs, Types, Math,
     XBookPaintGDI2, XBookWindows2,
     XSSIERenderTextIE,
     XSSIEDefs, XSSIEUtils, XSSIEKeys, XSSIEDocProps, XSSIECharRun, XSSIELogPhyPosition,
     XSSIELogParas, XSSIELogEditor, XSSIEGDIText, XSSIEPhyRow, XSSIEPaginate, XSSIECaretRow;

type TWinCaretAction = (wcaShow,wcaHide,wcaPos);
type TWinCaretEvent = procedure(const AAction: TWinCaretAction; const AX,AY,AHeight: integer) of object;
type TAXWTextChangedEvent = procedure(const AText: AxUCString) of object;

type TAXWEditor = class(TAXWLogDocEditor)
private
     function  GetText: AxUCString;
     procedure SetText(const Value: AxUCString);
     function  GetTextHeight: double;
     function  GetTextWidth: double;
     function  GetCHP: TAXWCHP;
     function  GetPAP: TAXWPAP;
     procedure SetBgColor(const Value: longword);
     procedure SetNoFormatting(const Value: boolean);
     function  GetBgColor: longword;
protected
     FGDI                : TAXWGDI;
     FArea               : TAXWClientArea;

     FVertAlign          : TAXWVertAlign;

     // Used in Inplace Edit mode
     FNoFormatting       : boolean;
     FFormulaMode        : boolean;

     FTextRender         : TAXWTextRender;

     FPaginator          : TAXWPaginator;
     FTextPrint          : TAXWTextPrint;
     FCaretRow           : TAXWCaretRowEditor;

     FCHP                : TAXWCHP;
     FPrevCHP            : TAXWCHP;

     FIsUpdating         : boolean;

     FTextChangedEvent   : TAXWTextChangedEvent;
     FBeforePaintEvent   : TNotifyEvent;
     FAfterPaintEvent    : TNotifyEvent;
     FCaretEvent         : TWinCaretEvent;
     FCharFmtChangedEvent: TNotifyEvent;

     FDebugEvent         : TNotifyEvent;
     FDebugText          : AxUCString;

     procedure DoParaHeight(APara: TAXWLogPara; var Y: double);
     procedure Paginate; override;
     procedure PaginateCaretPara; override;
     procedure CaretPosChanged(ASender: TObject);
//     procedure CalcTopPosition;
     function  ParasHeight: double;
     procedure CheckCharFormatChanged;

     procedure Debug(const AText: AxUCString);
public
     constructor Create(AGDI: TAXWGDI; AArea: TAXWClientArea);
     destructor Destroy; override;

     procedure PaintText;

     //* Call this after text is added.
     procedure CommitText;

     procedure KeyPress(var Key: AxUCChar);

     procedure BeginUpdate;
     procedure EndUpdate;
     procedure Command(ACmd: TAXWCommand);
     procedure FillCharProps;

     procedure DEBUG_DumpCharRuns(ALines: TStrings);

     property TextWidth: double read GetTextWidth;
     property TextHeight: double read GetTextHeight;
     //* Call CommitText after text is added.
     property Text: AxUCString read GetText write SetText;
     property BgColor: longword read GetBgColor write SetBgColor;

     // Options when used as XLSSpreadSheet cell editor.
     property VertAlign: TAXWVertAlign read FVertAlign write FVertAlign;
     property NoFormatting: boolean read FNoFormatting write SetNoFormatting;
     property SetNoFormattingOnFirstEqual: boolean read FFormulaMode write FFormulaMode;

     property CaretPAP: TAXWPAP read GetPAP;
     // CHP is only updated if OnCharFmtChanged is assigned.
     property CaretCHP: TAXWCHP read GetCHP;

     property OnBeforePaint: TNotifyEvent read FBeforePaintEvent write FBeforePaintEvent;
     property OnAfterPaint: TNotifyEvent read FAfterPaintEvent write FAfterPaintEvent;
     property OnCaret : TWinCaretEvent read FCaretEvent write FCaretEvent;
     property OnTextChanged: TAXWTextChangedEvent read FTextChangedEvent write FTextChangedEvent;
     property OnCharFmtChanged: TNotifyEvent read FCharFmtChangedEvent write FCharFmtChangedEvent;

     property OnDebug: TNotifyEvent read FDebugEvent write FDebugEvent;
     property DebugText: AxUCString read FDebugText;
     end;

type TAXWWinEditor = class(TXSSClientWindow)
protected
     FEditor     : TAXWEditor;

     FDebugText  : TStringList;
     FDebugEvent : TNotifyEvent;

     procedure UpdateWinCaret(const AAction: TWinCaretAction; const AX,AY,AHeight: integer);

     procedure MoveAreaHoriz(const ADist: integer);
     procedure DoDebug;
     function  GetText: AxUCString; virtual;
     procedure SetText(const Value: AxUCString); virtual;
public
     constructor Create(AParent: TXSSWindow; AArea: TAXWClientArea);
     destructor Destroy; override;

     procedure Paint; override;

     procedure SetFocus; override;
     procedure KillFocus; override;
     procedure KeyPress(Key: AxUCChar); override;

     function  PosAsString: AxUCString;

     property Text: AxUCString read GetText write SetText;

     property DebugText: TStringList read FDebugText;
     property OnDebug: TNotifyEvent read FDebugEvent write FDebugEvent;
     end;

implementation

{ TAXWEditor }

procedure TAXWEditor.BeginUpdate;
begin
  FIsUpdating := True;
end;

procedure TAXWEditor.CaretPosChanged(ASender: TObject);
var
  p: integer;
  Y: integer;
begin
  FCaretRow.RowChanged;
//  FCaretRow.SaveX;

  if (FCaretPos.Para <> Nil) and Assigned(FCaretEvent) then begin
    p := FCaretPos.RelativePos;
    Y := FTextRender.CalcOriginY;
//    FCaretEvent(wcaPos,FArea.CX1 + Round(FCaretRow[p].X),Y + CaretPos.Row.Y - CaretPos.Row.Height + CaretPos.Row.Descent,CaretPos.Row.Height);
    FCaretEvent(wcaPos,FArea.CX1 + Round(FCaretRow[p].X),Round(Y - CaretPos.Row.Height + CaretPos.Row.Descent),Round(CaretPos.Row.Height));
//    FCaretEvent(wcaPos,FCaretRow.ScrCaretX,FCaretRow.ScrCaretY,FCaretRow.ScrCaretH);
  end;

  CheckCharFormatChanged;
end;

procedure TAXWEditor.CheckCharFormatChanged;
var
  Fire: boolean;
begin
  if Assigned(FCharFmtChangedEvent) then begin
    if FLogParas.CaretCHPX <> Nil then
      FCHP.Assign(FLogParas.CaretCHPX)
    else
      FCHP.Clear;
    if FPrevCHP = Nil then begin
      FPrevCHP := TAXWCHP.Create(Nil);
      Fire := True;
    end
    else
      Fire := not FCHP.Equal(FPrevCHP);
    if Fire then
      FCharFmtChangedEvent(Self);
    FPrevCHP.Assign(FCHP);
  end;
end;

procedure TAXWEditor.Command(ACmd: TAXWCommand);
begin
  if ACmd <> axcNone then begin
    case CommandClass(ACmd) of
      axccControl   : CmdControl(ACmd);
      axccMove      : begin
        CmdMove(ACmd);
        if ACmd in [axcMoveLineUp,axcMoveLineDown,axcMovePgUp,axcMovePgDown] then
          FCaretRow.RestoreX
        else
          FCaretRow.SaveX;
      end;
      axccFormat    : begin
        if not FNoFormatting then begin
          if ACmd = axcFormatCHP then
            FCHP.CopyToCHPX(FTempCHPX);
          CmdFormat(ACmd);
          CheckCharFormatChanged;
        end;
      end;
      axccClipboard : CmdClipboard(ACmd);
      axccSelect    : CmdSelect(ACmd);
      axccDeleteText: CmdDelete(ACmd);
    end;
    if Dirty then begin
      Paginate;
      PaintText;
    end;
    if FRepaint then
      PaintText;
  end;
end;

procedure TAXWEditor.CommitText;
begin
  Paginate;
  Command(axcMoveEndOfDoc);
  PaintText;
  CaretPosChanged(Self);
end;

constructor TAXWEditor.Create(AGDI: TAXWGDI; AArea: TAXWClientArea);
begin
  inherited Create;

  FGDI := AGDI;
  FArea := AArea;

  FCaretPos.OnPosChange := CaretPosChanged;

  FTextPrint := TAXWTextPrint.Create(FGDI,FArea,FSelections,FDOP);
  FPaginator := TAXWPaginator.Create(FTextPrint,FArea);
  FPaginator.KeepTrailingSpaces := True;
  FCaretRow := TAXWCaretRowEditor.Create(FTextPrint,FLogParas,FArea);
  FCaretRow.CaretPos := FCaretPos;

  FTextRender := TAXWTextRender.Create(FGDI,Self,FCaretRow,FArea);

  FCHP := TAXWCHP.Create(FMasterCHP);
end;

procedure TAXWEditor.Debug(const AText: AxUCString);
begin
  FDebugText := AText;
  if Assigned(FDebugEvent) then
    FDebugEvent(Self);
end;

procedure TAXWEditor.DEBUG_DumpCharRuns(ALines: TStrings);
var
  i,j: integer;
  S: string;
begin
  for i := 0 to FLogParas.Count - 1 do begin
    ALines.Add('Para #' + IntToStr(i));
    for j := 0 to FLogParas[i].Count - 1 do begin
      S := '  CR #' + IntToStr(j);
      S := S + ' "' + FLogParas[i][j].Text + '"';
      ALines.Add(S);
    end;
  end;
end;

destructor TAXWEditor.Destroy;
begin
  FCHP.Free;
  if FPrevCHP <> Nil then
    FPrevCHP.Free;

  FCaretRow.Free;

  FPaginator.Free;
  FTextPrint.Free;

  FTextRender.Free;
  inherited;
end;

procedure TAXWEditor.DoParaHeight(APara: TAXWLogPara; var Y: double);
var
  i: integer;
  Row: TAXWPhyRow;
begin
  Row := Nil;

  for i := 0 to APara.Rows.Count - 1 do begin
    Row := APara.Rows[i];
    Row.Y := Y + Row.Ascent;
    Y := Y + Row.Height;
  end;

  if Row <> Nil then
    Y := Y + Row.Height * 0.25;
end;

procedure TAXWEditor.EndUpdate;
begin
  FIsUpdating := False;

  Paginate;
  PaintText;
end;

procedure TAXWEditor.FillCharProps;
begin
  FCHP.Assign(FlogParas.CaretCHPX);
end;

function TAXWEditor.GetBgColor: longword;
begin
  Result := FDOP.Color;
end;

function TAXWEditor.GetCHP: TAXWCHP;
begin
  Result := FCHP;
end;

function TAXWEditor.GetPAP: TAXWPAP;
begin
  Result := FMasterPAP;
end;

function TAXWEditor.GetText: AxUCString;
begin
  Result := PlainText;
end;

function TAXWEditor.GetTextHeight: double;
begin
  if FLogParas.Count > 0 then
    Result := FLogParas[0].Rows.Height
  else
    Result := 0;
end;

function TAXWEditor.GetTextWidth: double;
begin
  if FLogParas.Count > 0 then
    Result := FLogParas[0].Rows.Width
  else
    Result := 0;
end;

procedure TAXWEditor.KeyPress(var Key: AxUCChar);
var
  TextAdded: boolean;
begin
  TextAdded := False;
  if CharIsPrintable(Key) then begin
    if FFormulaMode and not FNoFormatting then
      FNoFormatting := (Key = '=') and (FLogParas.Empty or ((FLogParas.Count = 1) and (FLogParas[0].Text = '')));
    InsertText(Key,True);
    TextAdded := True;
  end;
  if TextAdded then begin
    if Dirty then
      Paginate;
    PaintText;
    if Assigned(FTextChangedEvent) then
      FTextChangedEvent(Text);
  end;
end;

procedure TAXWEditor.Paginate;
const
  Chars = 'abcdefghijklmnopqrstuvwxyz';
var
  i: integer;
  Y: double;
  Para: TAXWLogPara;
  Stack: TListStack;
begin
  if FIsUpdating then
    Exit;

  FLogParas.PPIX := FGDI.PixelsPerInchX;
  FLogParas.PPIY := FGDI.PixelsPerInchY;

  FLogParas.ClearRows;

  Stack := TListStack.Create;
  try
    Y := 0;
    for i := 0 to FLogParas.Count - 1 do begin
      Para := FLogParas[i];
      FPaginator.Paginate(Para);

      DoParaHeight(Para,Y);
    end;
  finally
    Stack.Free;
  end;

  ClearDirty;

//  CaretPosChanged(Nil);
end;

procedure TAXWEditor.PaginateCaretPara;
var
  Y: double;
  Para: TAXWLogPara;
begin
  if FIsUpdating then
    Exit;

  FLogParas.PPIX := FGDI.PixelsPerInchX;
  FLogParas.PPIY := FGDI.PixelsPerInchY;

  Para := FCaretRow.CaretPos.Para;
  if Para = Nil then
    Paginate
  else begin
    Para.Rows.Clear;
    FPaginator.Paginate(Para);

    if (Para.Prev <> Nil) and not Para.Empty then
      Y := Para.Prev.Rows[0].Y + Para.Prev.Rows.Height
    else
      Y := 0;
    DoParaHeight(Para,Y);
  end;

    ClearDirty;
  //  CaretPosChanged(Nil);
end;

procedure TAXWEditor.PaintText;
begin
  if FIsUpdating then
    Exit;

  if Assigned(FCaretEvent) then
    FCaretEvent(wcaHide,0,0,0);

  if Assigned(FBeforePaintEvent) then
    FBeforePaintEvent(Self);

  FGDI.PenColor := FDOP.Color;
  FGDI.BrushColor := FDOP.Color;
  FGDI.Rectangle(FArea.PadX1,FArea.PadY1,FArea.PadX2,FArea.PadY2);

  if not FLogParas.Empty then begin
    FTextRender.Render;

    CaretPosChanged(Nil);

    FRepaint := False;
  end;

  if Assigned(FAfterPaintEvent) then
    FAfterPaintEvent(Self);

  if Assigned(FCaretEvent) then
    FCaretEvent(wcaShow,0,0,0);
end;

function TAXWEditor.ParasHeight: double;
var
  i,j: integer;
begin
  Result := 0;

  for i := 0 to FLogParas.Count - 1 do begin
    for j := 0 to FLogParas[i].Rows.Count - 1 do
      Result := Result + FLogParas[i].Rows.Height;
  end;
end;

procedure TAXWEditor.SetBgColor(const Value: longword);
begin
  FDOP.Color := Value;
end;

procedure TAXWEditor.SetNoFormatting(const Value: boolean);
begin
  FNoFormatting := Value;
end;

procedure TAXWEditor.SetText(const Value: AxUCString);
begin
  AppendText(Value);
end;

{ TAXWWinEditor }

constructor TAXWWinEditor.Create(AParent: TXSSWindow; AArea: TAXWClientArea);
begin
  inherited Create(AParent);

  FEditor := TAXWEditor.Create(FSkin.GDI,AArea);
  FEditor.OnCaret := UpdateWinCaret;

  FDebugText := TStringList.Create;

  FSystem.CreateCaret;
end;

destructor TAXWWinEditor.Destroy;
begin
  FEditor.Free;
  FSystem.HideCaret;
  FDebugText.Free;
  FSystem.DestroyCaret;
  inherited;
end;

procedure TAXWWinEditor.DoDebug;
begin
  FDebugText.Clear;
end;

function TAXWWinEditor.GetText: AxUCString;
begin
  Result := FEditor.Text;
end;

procedure TAXWWinEditor.KeyPress(Key: AxUCChar);
begin
  inherited;
  FEditor.KeyPress(Key);
end;

procedure TAXWWinEditor.KillFocus;
begin
  FSystem.DestroyCaret;
end;

procedure TAXWWinEditor.Paint;
begin
  inherited;

  FEditor.PaintText;
end;

function TAXWWinEditor.PosAsString: AxUCString;
var
  p: integer;
begin
  if FEditor.CaretPos.Para <> Nil then begin
    p := FEditor.CaretPos.RelativePos;
    Result := Format('[%d:%d %d:%d]',[FEditor.CaretPos.Para.Index,FEditor.CaretPos.CharPos,FEditor.FCaretRow[p].X,FEditor.CaretPos.Row.Y]);
  end
  else
    Result := '';
end;

procedure TAXWWinEditor.MoveAreaHoriz(const ADist: integer);
begin
  Inc(FX1,ADist);
  Inc(FX2,ADist);
  FEditor.FArea.X1 := FEditor.FArea.X1 + ADist;
  FEditor.FArea.X2 := FEditor.FArea.X2 + ADist;
end;

procedure TAXWWinEditor.SetFocus;
begin
  inherited SetFocus;

  FSystem.CreateCaret;
end;

procedure TAXWWinEditor.SetText(const Value: AxUCString);
begin
  FEditor.Text := Value;
end;

procedure TAXWWinEditor.UpdateWinCaret(const AAction: TWinCaretAction; const AX,AY,AHeight: integer);
begin
  case AAction of
    wcaShow: FSystem.ShowCaret;
    wcaHide: FSystem.HideCaret;
    wcaPos : FSystem.ShowCaret(1,AHeight,AX,AY);
  end;
end;

end.
