unit uMainForm;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs, Grids, UITypes, uHexEditor, db.uCommon, StdCtrls, ExtCtrls, Buttons, ComCtrls, Menus, clipbrd, Printers;

type
  TfrmHexEdit = class(TForm)
    StatusBar1: TStatusBar;
    MainMenu1: TMainMenu;
    Datei1: TMenuItem;
    Neu1: TMenuItem;
    ffnen1: TMenuItem;
    Speichern1: TMenuItem;
    Speichernunter1: TMenuItem;
    N1: TMenuItem;
    Beenden1: TMenuItem;
    Bearbeiten1: TMenuItem;
    Rckgngig1: TMenuItem;
    N3: TMenuItem;
    Ausschneiden1: TMenuItem;
    Kopieren1: TMenuItem;
    Einfgen1: TMenuItem;
    Hilfe1: TMenuItem;
    Info1: TMenuItem;
    View1: TMenuItem;
    _16: TMenuItem;
    _32: TMenuItem;
    _64: TMenuItem;
    SaveDialog1: TSaveDialog;
    OpenDialog1: TOpenDialog;
    N2: TMenuItem;
    Blocksize1: TMenuItem;
    N2Bytesperblock1: TMenuItem;
    N4Bytesperblock1: TMenuItem;
    N1Byteperblock1: TMenuItem;
    N5: TMenuItem;
    Goto1: TMenuItem;
    Caretstyle1: TMenuItem;
    Fullblock1: TMenuItem;
    Leftline1: TMenuItem;
    Bottomline1: TMenuItem;
    Linesize1: TMenuItem;
    Grid1: TMenuItem;
    Offsetdisplay1: TMenuItem;
    Hex1: TMenuItem;
    Dec1: TMenuItem;
    None1: TMenuItem;
    Showmarkers1: TMenuItem;
    Find1: TMenuItem;
    FindNext1: TMenuItem;
    Jump1: TMenuItem;
    Amount40001: TMenuItem;
    N4: TMenuItem;
    Jumpforward1: TMenuItem;
    Jumpbackward1: TMenuItem;
    N6: TMenuItem;
    Printpreview1: TMenuItem;
    PrinterSetupDialog1: TPrinterSetupDialog;
    Printsetup1: TMenuItem;
    Print1: TMenuItem;
    N7: TMenuItem;
    Translation1: TMenuItem;
    Ansi1: TMenuItem;
    ASCII7Bit1: TMenuItem;
    DOS8Bit1: TMenuItem;
    Mac1: TMenuItem;
    IBMEBCDICcp381: TMenuItem;
    SwapNibbles1: TMenuItem;
    InsertNibble1: TMenuItem;
    DeleteNibble1: TMenuItem;
    N8: TMenuItem;
    Convertfile1: TMenuItem;
    MaskWhitespaces1: TMenuItem;
    AnyFile1: TMenuItem;
    HexText1: TMenuItem;
    ImportfromHexText1: TMenuItem;
    FixedFilesize1: TMenuItem;
    Octal1: TMenuItem;
    PasteText1: TMenuItem;
    N9: TMenuItem;
    Auto1: TMenuItem;
    Printlayout1: TMenuItem;
    Hex2: TMenuItem;
    Decimal1: TMenuItem;
    Octal2: TMenuItem;
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure Bearbeiten1Click(Sender: TObject);
    procedure Datei1Click(Sender: TObject);
    procedure Einfgen1Click(Sender: TObject);
    procedure Rckgngig1Click(Sender: TObject);
    procedure HexEditor1StateChanged(Sender: TObject);
    procedure Kopieren1Click(Sender: TObject);
    procedure Ausschneiden1Click(Sender: TObject);
    procedure _64Click(Sender: TObject);
    procedure View1Click(Sender: TObject);
    procedure Neu1Click(Sender: TObject);
    procedure ffnen1Click(Sender: TObject);
    procedure Beenden1Click(Sender: TObject);
    procedure Speichern1Click(Sender: TObject);
    procedure HexEditor1KeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure FormActivate(Sender: TObject);
    procedure Blocksize1Click(Sender: TObject);
    procedure N4Bytesperblock1Click(Sender: TObject);
    procedure Goto1Click(Sender: TObject);
    procedure Caretstyle1Click(Sender: TObject);
    procedure Bottomline1Click(Sender: TObject);
    procedure Grid1Click(Sender: TObject);
    procedure Offsetdisplay1Click(Sender: TObject);
    procedure None1Click(Sender: TObject);
    procedure Showmarkers1Click(Sender: TObject);
    procedure FindNext1Click(Sender: TObject);
    procedure Find1Click(Sender: TObject);
    procedure Jumpforward1Click(Sender: TObject);
    procedure Jumpbackward1Click(Sender: TObject);
    procedure Amount40001Click(Sender: TObject);
    procedure Jump1Click(Sender: TObject);
    procedure Info1Click(Sender: TObject);
    procedure Printpreview1Click(Sender: TObject);
    procedure Printsetup1Click(Sender: TObject);
    procedure Print1Click(Sender: TObject);
    procedure Ansi1Click(Sender: TObject);
    procedure SwapNibbles1Click(Sender: TObject);
    procedure InsertNibble1Click(Sender: TObject);
    procedure DeleteNibble1Click(Sender: TObject);
    procedure Convertfile1Click(Sender: TObject);
    procedure MaskWhitespaces1Click(Sender: TObject);
    procedure AnyFile1Click(Sender: TObject);
    procedure HexText1Click(Sender: TObject);
    procedure ImportfromHexText1Click(Sender: TObject);
    procedure FixedFilesize1Click(Sender: TObject);
    procedure changelinelengths1Click(Sender: TObject);
    procedure PasteText1Click(Sender: TObject);
    procedure Auto1Click(Sender: TObject);
    procedure Printlayout1Click(Sender: TObject);
    procedure Hex2Click(Sender: TObject);
    procedure Decimal1Click(Sender: TObject);
    procedure Octal2Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    HexToCanvas1: THexToCanvas;
    HexEditor1  : THexEditor;
    function CheckChanges: Boolean;
    function SaveFile: Boolean;
    procedure Error(aMsg: string);
    procedure AppendByte;
  public
    { Public-Deklarationen }
  end;

procedure db_ShowDllForm_Plugins(var frm: TFormClass; var ft: TSPFileType; var strParentModuleName, strModuleName, strClassName, strWindowName, strIconFileName: PAnsiChar; const bShow: Boolean = True); stdcall;

implementation

uses sample2, sample3;

{$R *.DFM}

procedure db_ShowDllForm_Plugins(var frm: TFormClass; var ft: TSPFileType; var strParentModuleName, strModuleName, strClassName, strWindowName, strIconFileName: PAnsiChar; const bShow: Boolean = True); stdcall;
begin
  frm                     := TfrmHexEdit;
  ft                      := ftDelphiDll;
  strParentModuleName     := '程序员工具';
  strModuleName           := 'HexEdit';
  strIconFileName         := '';
  strClassName            := '';
  strWindowName           := '';
  Application.Handle      := GetMainFormApplication.Handle;
  Application.Icon.Handle := GetMainFormApplication.Icon.Handle;
end;

const
  cCaption = 'HexEdit v2.0';

var
  FindPos      : Integer             = -1;
  FindBuf      : PChar               = nil;
  FindLen      : Integer             = 0;
  FindStr      : string              = '';
  FindICase    : Boolean             = False;
  JumpOffs     : Integer             = 4000;
  CanvasDisplay: TOffsetDisplayStyle = odHex;

function SaveText(const aText, aName: string): Boolean;
var
  f: System.Text;
begin
{$I-}
  AssignFile(f, aName);
  Rewrite(f);
  write(f, aText);
  CloseFile(f);
  Result := IOResult = 0;
{$I+}
end;

function LoadText(const aName: string): string;
var
  fST: TMemoryStream;
begin
  Result := '';
  if not FileExists(aName) then
    Exit;
  fST := TMemoryStream.Create;
  try
    fST.LoadFromFile(aName);
    SetLength(Result, fST.Size);
    fST.Position := 0;
    fST.Read(Result[1], fST.Size);
  finally
    fST.Free;
  end;
end;

procedure TfrmHexEdit.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  if FindBuf <> nil then
  begin
    FreeMem(FindBuf);
    FindBuf := nil;
  end;
  HexToCanvas1.Free;
  HexEditor1.Free;

  Action := caFree;
end;

procedure TfrmHexEdit.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
  CanClose := CheckChanges;
end;

procedure TfrmHexEdit.FormCreate(Sender: TObject);
begin
  HexEditor1                           := THexEditor.Create(Self);
  HexEditor1.Parent                    := Self;
  HexEditor1.Align                     := alClient;
  HexEditor1.BytesPerColumn            := 1;
  HexEditor1.Translation               := ttAnsi;
  HexEditor1.BackupExtension           := '.bak';
  HexEditor1.OffsetDisplay             := odHex;
  HexEditor1.BytesPerLine              := 16;
  HexEditor1.Colors.Background         := clWindow;
  HexEditor1.Colors.PositionBackground := clMaroon;
  HexEditor1.Colors.PositionText       := clRed;
  HexEditor1.Colors.ChangedBackground  := clWindow;
  HexEditor1.Colors.ChangedText        := clRed;
  HexEditor1.Colors.CursorFrame        := clBlack;
  HexEditor1.Colors.Offset             := clBlack;
  HexEditor1.Colors.OddColumn          := clBlue;
  HexEditor1.Colors.EvenColumn         := clNavy;
  HexEditor1.FocusFrame                := False;
  HexEditor1.Font.Charset              := ANSI_CHARSET;
  HexEditor1.Font.Color                := clBlack;
  HexEditor1.Font.Height               := -16;
  HexEditor1.Font.Name                 := '宋体';
  HexEditor1.Font.Pitch                := fpFixed;
  HexEditor1.Font.Style                := [];
  HexEditor1.OffsetSeparator           := ':';
  HexEditor1.AllowInsertMode           := True;
  HexEditor1.OnStateChanged            := HexEditor1StateChanged;
  HexEditor1.OnKeyDown                 := HexEditor1KeyDown;
  HexToCanvas1                         := THexToCanvas.Create(Self);
  HexToCanvas1.hexeditor               := HexEditor1;
end;

procedure TfrmHexEdit.Bearbeiten1Click(Sender: TObject);
begin
  with Rckgngig1 do
  begin
    Enabled := HexEditor1.CanUndo;
    Caption := '撤销 : ' + HexEditor1.UndoDescription;
  end;
  Ausschneiden1.Enabled := HexEditor1.SelCount > 0;
  Kopieren1.Enabled     := HexEditor1.SelCount > 0;
  Einfgen1.Enabled      := Clipboard.HasFormat(CF_TEXT);
  PasteText1.Enabled    := Einfgen1.Enabled;
  Goto1.Enabled         := HexEditor1.DataSize > 0;
  Find1.Enabled         := HexEditor1.DataSize > 0;
  FindNext1.Enabled     := HexEditor1.DataSize > 0;
  InsertNibble1.Enabled := HexEditor1.DataSize > 0;
  DeleteNibble1.Enabled := HexEditor1.DataSize > 0;
  Convertfile1.Enabled  := HexEditor1.DataSize > 0;
  if HexEditor1.SelCount = 0 then
    Convertfile1.Caption := '文件编码转换'
  else
    Convertfile1.Caption := '选区编码转换'
end;

procedure TfrmHexEdit.Datei1Click(Sender: TObject);
begin
  Speichern1.Enabled      := HexEditor1.Modified and (not HexEditor1.ReadOnly);
  Speichernunter1.Enabled := HexEditor1.DataSize > 0;
  Printpreview1.Enabled   := HexEditor1.DataSize > 0;
  Print1.Enabled          := HexEditor1.DataSize > 0;
end;

procedure TfrmHexEdit.Einfgen1Click(Sender: TObject);
var
  sr: string;
  BT: Integer;
begin
  Bearbeiten1Click(Sender);
  if Einfgen1.Enabled then
  begin
    sr := Clipboard.AsText;
    HexEditor1.ReplaceSelection(ConvertHexToBin(@sr[1], @sr[1], Length(sr), False, BT), BT);
  end;
end;

procedure TfrmHexEdit.Rckgngig1Click(Sender: TObject);
begin
  if HexEditor1.CanUndo then
    HexEditor1.Undo;
end;

procedure TfrmHexEdit.HexEditor1StateChanged(Sender: TObject);
var
  pSS, pSE: Integer;
begin
  with HexEditor1, StatusBar1 do
  begin
    Panels[0].Text := 'Pos : ' + IntToStr(GetCursorPos);
    if SelCount <> 0 then
    begin
      pSS            := SelStart;
      pSE            := SelEnd;
      Panels[1].Text := 'Sel : ' + IntToStr(Min(pSS, pSE)) + ' - ' + IntToStr(Max(pSS, pSE));
    end
    else
      Panels[1].Text := '';
    if Modified then
      Panels[2].Text := '*'
    else
      Panels[2].Text := '';
    if ReadOnly then
      Panels[3].Text := 'R'
    else
      Panels[3].Text := '';
    if IsInsertMode then
      Panels[4].Text := 'INS'
    else
      Panels[4].Text := 'OVW';
    Panels[5].Text   := 'Size : ' + IntToStr(DataSize);
    Caption          := cCaption + '[' + FileName + ']';
  end;
end;

procedure SetCBText(aP: PChar; aCount: Integer);
var
  sr: string;
begin
  SetLength(sr, aCount * 2);
  ConvertBinToHex(aP, @sr[1], aCount, False);
  Clipboard.AsText := sr;
  SetLength(sr, 0);
end;

procedure TfrmHexEdit.Kopieren1Click(Sender: TObject);
var
  pct: Integer;
  pPC: PChar;
begin
  Bearbeiten1Click(Sender);
  if Kopieren1.Enabled then
    with HexEditor1 do
    begin
      pct := SelCount;
      pPC := BufferFromFile(Min(SelStart, SelEnd), pct);
      SetCBText(pPC, pct);
      FreeMem(pPC, pct);
    end;
end;

procedure TfrmHexEdit.Ausschneiden1Click(Sender: TObject);
var
  pct: Integer;
  pPC: PChar;
begin
  Bearbeiten1Click(Sender);
  if Ausschneiden1.Enabled then
    with HexEditor1 do
    begin
      pct := SelCount;
      pPC := BufferFromFile(Min(SelStart, SelEnd), pct);
      SetCBText(pPC, pct);
      FreeMem(pPC, pct);
      DeleteSelection;
    end;
end;

procedure TfrmHexEdit._64Click(Sender: TObject);
begin
  HexEditor1.BytesPerLine := TMenuItem(Sender).Tag;
end;

procedure TfrmHexEdit.View1Click(Sender: TObject);
begin
  Case HexEditor1.BytesPerLine of
    16:
      _16.Checked := True;
    32:
      _32.Checked := True;
    64:
      _64.Checked := True;
  end;
  Grid1.Checked            := HexEditor1.GridLineWidth = 1;
  Showmarkers1.Checked     := HexEditor1.ShowMarkerColumn;
  SwapNibbles1.Checked     := HexEditor1.SwapNibbles;
  MaskWhitespaces1.Checked := HexEditor1.MaskWhiteSpaces;
  Case HexEditor1.Translation of
    ttAnsi:
      Ansi1.Checked := True;
    ttASCII:
      ASCII7Bit1.Checked := True;
    ttDOS8:
      DOS8Bit1.Checked := True;
    ttMac:
      Mac1.Checked := True;
    ttEBCDIC:
      IBMEBCDICcp381.Checked := True;
  end;
  FixedFilesize1.Checked := HexEditor1.NoSizeChange;
end;

procedure TfrmHexEdit.Neu1Click(Sender: TObject);
begin
  if CheckChanges then
    HexEditor1.CreateEmptyFile('未命名');
end;

function TfrmHexEdit.CheckChanges: Boolean;
var
  psr: string;
begin
  Result := True;
  if not HexEditor1.Modified then
    Exit;

  Result := False;
  psr    := HexEditor1.FileName;
  if not FileExists(psr) then
    psr := '未命名文件';
  Case MessageDlg('保存修改'#13#10 + psr + ' ?', mtConfirmation, [mbNo, mbYes, mbCancel], 0) of
    IDNo:
      Result := True;
    IDYES:
      Result := SaveFile;
  end;
end;

procedure TfrmHexEdit.ffnen1Click(Sender: TObject);
begin
  if CheckChanges then
    if OpenDialog1.Execute then
    begin
      HexEditor1.LoadFromFile(OpenDialog1.FileName);
      if ofReadOnly in OpenDialog1.Options then
        HexEditor1.ReadOnly := True;
    end;
end;

procedure TfrmHexEdit.Beenden1Click(Sender: TObject);
begin
  Close;
end;

procedure TfrmHexEdit.Speichern1Click(Sender: TObject);
begin
  SaveFile;
end;

function TfrmHexEdit.SaveFile: Boolean;
begin
  Result := False;
  if (not FileExists(HexEditor1.FileName)) or HexEditor1.ReadOnly then
  begin
    if SaveDialog1.Execute then
    begin
      if HexEditor1.SaveToFile(SaveDialog1.FileName) then
        Result := True
      else
      begin
        Error('无法保存文件'#13#10 + HexEditor1.FileName);
      end;
    end
  end
  else if HexEditor1.SaveToFile(HexEditor1.FileName) then
    Result := True
  else
    Error('无法保存文件'#13#10 + HexEditor1.FileName);
end;

procedure TfrmHexEdit.Error(aMsg: string);
begin
  Windows.MessageBox(Handle, PChar(aMsg), nil, MB_ICONHAND);
end;

procedure TfrmHexEdit.HexEditor1KeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if (Shift = [ssCTRL]) and (Key = Ord('A')) then
  begin
    Key := 0;
    AppendByte;
  end;
end;

procedure TfrmHexEdit.AppendByte;
var
  pBY: Byte;
begin
  pBY := 255;
  HexEditor1.AppendBuffer(@pBY, 1);
end;

procedure TfrmHexEdit.FormActivate(Sender: TObject);
{$J+}
const
  flFirst: Boolean = True;
{$J-}
begin
  if flFirst then
  begin
    flFirst := False;
    if ParamCount > 0 then
      if FileExists(ParamStr(1)) then
      begin
        Application.ProcessMessages;
        HexEditor1.LoadFromFile(ParamStr(1));
      end;
  end;
end;

procedure TfrmHexEdit.Blocksize1Click(Sender: TObject);
begin
  case HexEditor1.BytesPerColumn of
    1:
      N1Byteperblock1.Checked := True;
    2:
      N2Bytesperblock1.Checked := True;
    4:
      N4Bytesperblock1.Checked := True;
  end;
end;

procedure TfrmHexEdit.N4Bytesperblock1Click(Sender: TObject);
begin
  HexEditor1.BytesPerColumn := TMenuItem(Sender).Tag;
end;

procedure TfrmHexEdit.Goto1Click(Sender: TObject);
var
  sr: string;
  s1: LongInt;
begin
  if HexEditor1.DataSize < 1 then
    Exit;

  s1 := HexEditor1.GetCursorPos;
  sr := IntToStr(s1);
  if InputQuery('转到...', '(输入：( 0x or $ 格式) )', sr) then
  begin
    if Pos('0x', AnsiLowerCase(sr)) = 1 then
      sr := '$' + Copy(sr, 3, MaxInt);
    s1   := StrToIntDef(sr, -1);
    if not HexEditor1.Seek(s1, soFromBeginning, True) then
      Error('不正确的位置')
  end;
end;

procedure TfrmHexEdit.Caretstyle1Click(Sender: TObject);
begin
  case HexEditor1.CaretStyle of
    csFull:
      Fullblock1.Checked := True;
    csLeftLine:
      Leftline1.Checked := True;
    csBottomLine:
      Bottomline1.Checked := True;
  end;
end;

procedure TfrmHexEdit.Bottomline1Click(Sender: TObject);
begin
  HexEditor1.CaretStyle := TCaretStyle(TMenuItem(Sender).Tag);
end;

procedure TfrmHexEdit.Grid1Click(Sender: TObject);
begin
  HexEditor1.GridLineWidth := 1 - HexEditor1.GridLineWidth;
end;

procedure TfrmHexEdit.Offsetdisplay1Click(Sender: TObject);
begin
  case HexEditor1.OffsetDisplay of
    odHex:
      Hex1.Checked := True;
    odDec:
      Dec1.Checked := True;
    odOctal:
      Octal1.Checked := True;
    odNone:
      None1.Checked := True;
  end;
  Auto1.Checked := HexEditor1.AutoCaretMode;
end;

procedure TfrmHexEdit.None1Click(Sender: TObject);
begin
  HexEditor1.OffsetDisplay := TOffsetDisplayStyle(TMenuItem(Sender).Tag);
end;

procedure TfrmHexEdit.Showmarkers1Click(Sender: TObject);
begin
  HexEditor1.ShowMarkerColumn := not HexEditor1.ShowMarkerColumn;
end;

procedure TfrmHexEdit.FindNext1Click(Sender: TObject);
begin
  if HexEditor1.DataSize < 1 then
    Exit;

  if FindStr = '' then
    Find1.Click
  else
  begin
    if FindPos = HexEditor1.SelEnd then
      Inc(FindPos, 1);
    FindPos := HexEditor1.Find(FindBuf, FindLen, FindPos, HexEditor1.DataSize - 1, FindICase, False);
    if FindPos = -1 then
      ShowMessage('Data not found.')
    else
    begin
      HexEditor1.SelStart := FindPos + FindLen - 1;
      HexEditor1.SelEnd   := FindPos;
    end;
  end;
end;

procedure TfrmHexEdit.Find1Click(Sender: TObject);
const
  cHexChars = '0123456789abcdef';
var
  pSTR, pTMP: String;
  pct, pCT1 : Integer;
begin
  if HexEditor1.DataSize < 1 then
    Exit;

  FindPos   := -1;
  FindICase := False;
  if FindBuf <> nil then
  begin
    FreeMem(FindBuf);
    FindBuf := nil;
  end;
  if not InputQuery('Find Data', '"t.." ascii, "T.." + ignore case, else search hex', FindStr) then
    Exit;

  if FindStr = '' then
    Exit;

  pSTR := '';
  if UpCase(FindStr[1]) = 'T' then
  begin
    pSTR := Copy(FindStr, 2, MaxInt);
    pCT1 := Length(pSTR);
  end
  else
  begin
    pTMP    := AnsiLowerCase(FindStr);
    for pct := Length(pTMP) downto 1 do
      if Pos(pTMP[pct], cHexChars) = 0 then
        Delete(pTMP, pct, 1);
    while (Length(pTMP) mod 2) <> 0 do
      pTMP := '0' + pTMP;
    if pTMP = '' then
      Exit;

    pSTR    := '';
    pCT1    := Length(pTMP) div 2;
    for pct := 0 to (Length(pTMP) div 2) - 1 do
      pSTR  := pSTR + Char((Pos(pTMP[pct * 2 + 1], cHexChars) - 1) * 16 + (Pos(pTMP[pct * 2 + 2], cHexChars) - 1));
  end;
  if pCT1 = 0 then
    Exit;

  GetMem(FindBuf, pCT1);
  try
    if FindStr[1] = 'T' then
      FindICase := True;
    FindLen     := pCT1;
    Move(pSTR[1], FindBuf^, pCT1);
    FindPos := HexEditor1.Find(FindBuf, FindLen, HexEditor1.GetCursorPos, HexEditor1.DataSize - 1, FindICase, UpCase(FindStr[1]) = 'T');
    if FindPos = -1 then
      ShowMessage('没有发现任何数据.')
    else
    begin
      HexEditor1.SelStart := FindPos + FindLen - 1;
      HexEditor1.SelEnd   := FindPos;
    end;
  finally
  end;
end;

procedure TfrmHexEdit.Jumpforward1Click(Sender: TObject);
begin
  HexEditor1.Seek(JumpOffs, soFromCurrent, False);
end;

procedure TfrmHexEdit.Jumpbackward1Click(Sender: TObject);
begin
  HexEditor1.Seek(-JumpOffs, soFromCurrent, False);
end;

procedure TfrmHexEdit.Amount40001Click(Sender: TObject);
var
  sr: string;
begin
  sr := IntToStr(JumpOffs);
  if InputQuery('偏移：', '输入值：', sr) then
    JumpOffs := StrToIntDef(sr, JumpOffs);
end;

procedure TfrmHexEdit.Jump1Click(Sender: TObject);
begin
  Amount40001.Caption   := '偏移：' + IntToStr(JumpOffs);
  Jumpforward1.Enabled  := HexEditor1.DataSize > 0;
  Jumpbackward1.Enabled := HexEditor1.DataSize > 0;
end;

procedure TfrmHexEdit.Info1Click(Sender: TObject);
begin
  ShowMessage('HexEdit v2.0' + #13#10 + '2019-11-11' + #13#10 + 'dbyoung@sina.com' + #13#10 + 'https://blog.csdn.net/dbyoung');
end;

procedure TfrmHexEdit.Printpreview1Click(Sender: TObject);
var
  pBMP: TBitMap;
  l1  : Integer;
begin
  pBMP := TBitMap.Create;
  try
    pBMP.Width  := 1000;
    pBMP.Height := Round(Printer.PageHeight / Printer.PageWidth * 1000);
    with pBMP do
    begin
      Canvas.Brush.Color := clWhite;
      Canvas.Brush.Style := bsSolid;
      Canvas.FillRect(Rect(0, 0, Width, Height));
      HexToCanvas1.GetLayout;
      HexToCanvas1.MemFieldDisplay := CanvasDisplay;
      HexToCanvas1.StretchToFit    := False;
      HexToCanvas1.TopMargin       := Height div 20;
      HexToCanvas1.BottomMargin    := Height - (Height div 20);
      HexToCanvas1.LeftMargin      := Width div 20;
      HexToCanvas1.RightMargin     := Width - (Width div 20);
      HexToCanvas1.Draw(Canvas, 0, HexEditor1.DataSize - 1, 'HexEdit : ' + ExtractFileName(HexEditor1.FileName), '页码 1');
    end;
    with tfmPreview.Create(Application) do
      try
        Image1.Width := ClientWidth;
        l1           := Round(Image1.Width / pBMP.Width * pBMP.Height);
        if l1 > ClientHeight then
        begin
          Image1.Height := ClientHeight;
          Image1.Width  := Round(Image1.Height / pBMP.Height * pBMP.Width);
        end
        else
          Image1.Height := l1;
        Image1.Picture.Bitmap.Assign(pBMP);
        ShowModal;
      finally
        Free;
      end;
  finally
    pBMP.Free;
  end;
end;

procedure TfrmHexEdit.Printsetup1Click(Sender: TObject);
begin
  PrinterSetupDialog1.Execute;
end;

procedure TfrmHexEdit.Print1Click(Sender: TObject);
var
  l1, l2: Integer;
begin
  if HexEditor1.DataSize < 1 then
    Exit;

  l1 := 0;
  with Printer do
  begin
    HexToCanvas1.StretchToFit := False;
    HexToCanvas1.GetLayout;
    HexToCanvas1.MemFieldDisplay := CanvasDisplay;
    HexToCanvas1.TopMargin       := PageHeight div 20;
    HexToCanvas1.BottomMargin    := PageHeight - (PageHeight div 20);
    HexToCanvas1.LeftMargin      := PageWidth div 20;
    HexToCanvas1.RightMargin     := PageWidth - (PageWidth div 20);
    l2                           := 1;
    BeginDoc;
    repeat
      Canvas.Brush.Color := clWhite;
      Canvas.Brush.Style := bsSolid;
      Canvas.FillRect(Rect(0, 0, PageWidth, PageHeight));
      l1 := HexToCanvas1.Draw(Canvas, l1, HexEditor1.DataSize - 1, 'HexEdit : ' + ExtractFileName(HexEditor1.FileName), '页码 ' + IntToStr(l2));
      if l1 < HexEditor1.DataSize then
        NewPage;
      l2 := l2 + 1;
    until l1 >= HexEditor1.DataSize;
    EndDoc;
  end;
end;

procedure TfrmHexEdit.Ansi1Click(Sender: TObject);
begin
  HexEditor1.Translation := TTranslationType(TMenuItem(Sender).Tag);
end;

procedure TfrmHexEdit.SwapNibbles1Click(Sender: TObject);
begin
  HexEditor1.SwapNibbles := not HexEditor1.SwapNibbles;
end;

procedure TfrmHexEdit.InsertNibble1Click(Sender: TObject);
begin
  HexEditor1.InsertNibble(HexEditor1.GetCursorPos, HexEditor1.InCharField or ((HexEditor1.Col mod 2) = 0));
end;

procedure TfrmHexEdit.DeleteNibble1Click(Sender: TObject);
begin
  HexEditor1.DeleteNibble(HexEditor1.GetCursorPos, HexEditor1.InCharField or ((HexEditor1.Col mod 2) = 0));
end;

procedure TfrmHexEdit.Convertfile1Click(Sender: TObject);
var
  pFrom, pTo: Integer;
begin
  pFrom := 0;
  pTo   := HexEditor1.DataSize - 1;
  if HexEditor1.SelCount <> 0 then
  begin
    pFrom := Min(HexEditor1.SelStart, HexEditor1.SelEnd);
    pTo   := Max(HexEditor1.SelStart, HexEditor1.SelEnd);
  end;
  with TForm2.Create(Application) do
    try
      Caption := Convertfile1.Caption;
      if ShowModal = IDOK then
        HexEditor1.ConvertRange(pFrom, pTo, TTranslationType(ListBox1.ItemIndex), TTranslationType(ListBox2.ItemIndex));
    finally
      Free;
    end;
end;

procedure TfrmHexEdit.MaskWhitespaces1Click(Sender: TObject);
begin
  HexEditor1.MaskWhiteSpaces := not HexEditor1.MaskWhiteSpaces;
end;

procedure TfrmHexEdit.AnyFile1Click(Sender: TObject);
begin
  if SaveDialog1.Execute then
    if not HexEditor1.SaveToFile(SaveDialog1.FileName) then
      Error('无法保存文件'#13#10 + SaveDialog1.FileName);
end;

procedure TfrmHexEdit.HexText1Click(Sender: TObject);
begin
  if SaveDialog1.Execute then
    if not SaveText(HexEditor1.AsHex, SaveDialog1.FileName) then
      Error('无法保存文件'#13#10 + SaveDialog1.FileName);
end;

procedure TfrmHexEdit.ImportfromHexText1Click(Sender: TObject);
begin
  if OpenDialog1.Execute then
  begin
    HexEditor1.AsHex := LoadText(OpenDialog1.FileName);
    if ofReadOnly in OpenDialog1.Options then
      HexEditor1.ReadOnly := True;
  end;
end;

procedure TfrmHexEdit.FixedFilesize1Click(Sender: TObject);
begin
  HexEditor1.NoSizeChange := not HexEditor1.NoSizeChange;
end;

procedure TfrmHexEdit.changelinelengths1Click(Sender: TObject);
var
  pct: Integer;
  pCU: Integer;
  pLL: TList;
begin
  pLL := TList.Create;
  try
    pct     := HexEditor1.BytesPerLine;
    for pCU := 0 to 1000 do
      pLL.Add(Pointer(Random(pct) + 1));
    HexEditor1.SetLineLengths(pLL);
  finally
    pLL.Free;
  end;
end;

procedure TfrmHexEdit.PasteText1Click(Sender: TObject);
var
  sr: string;
begin
  Bearbeiten1Click(Sender);
  if Einfgen1.Enabled then
  begin
    sr := Clipboard.AsText;
    HexEditor1.ReplaceSelection(@sr[1], Length(sr));
  end;
end;

procedure TfrmHexEdit.Auto1Click(Sender: TObject);
begin
  HexEditor1.AutoCaretMode := not HexEditor1.AutoCaretMode;
end;

procedure TfrmHexEdit.Printlayout1Click(Sender: TObject);
begin
  case CanvasDisplay of
    odHex:
      Hex2.Checked := True;
    odDec:
      Decimal1.Checked := True;
    odOctal:
      Octal2.Checked := True;
  end;
end;

procedure TfrmHexEdit.Hex2Click(Sender: TObject);
begin
  CanvasDisplay := odHex;
end;

procedure TfrmHexEdit.Decimal1Click(Sender: TObject);
begin
  CanvasDisplay := odDec;
end;

procedure TfrmHexEdit.Octal2Click(Sender: TObject);
begin
  CanvasDisplay := odOctal;
end;

end.
