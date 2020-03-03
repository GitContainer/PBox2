unit uMainForm;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.StrUtils, System.Variants, System.Classes, System.Math, Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.WinXCtrls, Vcl.ComCtrls, Vcl.ExtCtrls, Vcl.Grids,
  db.uCommon, uHexEditor, uResource, Vcl.Menus;

type
  TfrmPEInfo = class(TForm)
    srchbxSelectPEFile: TSearchBox;
    lbl1: TLabel;
    dlgOpenSelectPEFile: TOpenDialog;
    pgcPEInfo: TPageControl;
    tsAll: TTabSheet;
    tsDosHeader: TTabSheet;
    tsNTHeader: TTabSheet;
    tsSectionTable: TTabSheet;
    tsSectionData: TTabSheet;
    lbl39: TLabel;
    lbl38: TLabel;
    lbl37: TLabel;
    lbl40: TLabel;
    lbl43: TLabel;
    lbl42: TLabel;
    lbl41: TLabel;
    lbl32: TLabel;
    lbl31: TLabel;
    lbl30: TLabel;
    lbl33: TLabel;
    lbl36: TLabel;
    lbl35: TLabel;
    lbl34: TLabel;
    lbl53: TLabel;
    lbl52: TLabel;
    lbl51: TLabel;
    lbl54: TLabel;
    lbl56: TLabel;
    lbl55: TLabel;
    lbl46: TLabel;
    lbl45: TLabel;
    lbl44: TLabel;
    lbl47: TLabel;
    lbl50: TLabel;
    lbl49: TLabel;
    lbl48: TLabel;
    lbl29: TLabel;
    lbl10: TLabel;
    lbl9: TLabel;
    lbl8: TLabel;
    lbl11: TLabel;
    lbl14: TLabel;
    lbl13: TLabel;
    lbl12: TLabel;
    lbl3: TLabel;
    lbl2: TLabel;
    Label1: TLabel;
    lbl4: TLabel;
    lbl7: TLabel;
    lbl6: TLabel;
    lbl5: TLabel;
    lbl24: TLabel;
    lbl23: TLabel;
    lbl22: TLabel;
    lbl25: TLabel;
    lbl28: TLabel;
    lbl27: TLabel;
    lbl26: TLabel;
    lbl18: TLabel;
    lbl16: TLabel;
    lbl15: TLabel;
    lbl19: TLabel;
    lbl21: TLabel;
    lbl17: TLabel;
    lbl20: TLabel;
    lbl57: TLabel;
    Label2: TLabel;
    lbl58: TLabel;
    lbl59: TLabel;
    lbl60: TLabel;
    lbl61: TLabel;
    lbl62: TLabel;
    lbl63: TLabel;
    lbl64: TLabel;
    lbl65: TLabel;
    lbl66: TLabel;
    lbl67: TLabel;
    lbl68: TLabel;
    lbl69: TLabel;
    lbl70: TLabel;
    lbl71: TLabel;
    lbl72: TLabel;
    lbl73: TLabel;
    lbl74: TLabel;
    lbl75: TLabel;
    lbl76: TLabel;
    scrlbx1: TScrollBox;
    lbl77: TLabel;
    lbl78: TLabel;
    lbl79: TLabel;
    lbl80: TLabel;
    lbl81: TLabel;
    lbl82: TLabel;
    lbl83: TLabel;
    lbl84: TLabel;
    lbl85: TLabel;
    lbl86: TLabel;
    lbl87: TLabel;
    lbl88: TLabel;
    lbl89: TLabel;
    lbl90: TLabel;
    lbl91: TLabel;
    lbl92: TLabel;
    lbl93: TLabel;
    lbl94: TLabel;
    lbl95: TLabel;
    lbl96: TLabel;
    lbl97: TLabel;
    lbl98: TLabel;
    lbl99: TLabel;
    lbl100: TLabel;
    lbl101: TLabel;
    lbl102: TLabel;
    lbl103: TLabel;
    lbl104: TLabel;
    lbl105: TLabel;
    lbl106: TLabel;
    lbl107: TLabel;
    lbl108: TLabel;
    lbl109: TLabel;
    lbl110: TLabel;
    lbl111: TLabel;
    lbl112: TLabel;
    lbl113: TLabel;
    lbl114: TLabel;
    lbl115: TLabel;
    lbl116: TLabel;
    lbl117: TLabel;
    lbl118: TLabel;
    lbl119: TLabel;
    lbl120: TLabel;
    lbl121: TLabel;
    lbl122: TLabel;
    lbl123: TLabel;
    lbl124: TLabel;
    lbl125: TLabel;
    lbl126: TLabel;
    lbl127: TLabel;
    lbl128: TLabel;
    lbl129: TLabel;
    lbl130: TLabel;
    lbl131: TLabel;
    lbl132: TLabel;
    lbl133: TLabel;
    lbl134: TLabel;
    lbl135: TLabel;
    lbl136: TLabel;
    lbl137: TLabel;
    lbl138: TLabel;
    lbl139: TLabel;
    lbl140: TLabel;
    lbl141: TLabel;
    lbl142: TLabel;
    lbl143: TLabel;
    lbl144: TLabel;
    lbl145: TLabel;
    lbl146: TLabel;
    lbl147: TLabel;
    lbl148: TLabel;
    lbl149: TLabel;
    lbl150: TLabel;
    lbl151: TLabel;
    lbl152: TLabel;
    lbl153: TLabel;
    lbl154: TLabel;
    lbl155: TLabel;
    lbl156: TLabel;
    lbl157: TLabel;
    lbl158: TLabel;
    lbl160: TLabel;
    lbl159: TLabel;
    lbl161: TLabel;
    lbl162: TLabel;
    lbl164: TLabel;
    lbl166: TLabel;
    lbl168: TLabel;
    lbl163: TLabel;
    lbl165: TLabel;
    lbl167: TLabel;
    lbl169: TLabel;
    lbl170: TLabel;
    lbl172: TLabel;
    lbl174: TLabel;
    lbl176: TLabel;
    lbl178: TLabel;
    lbl180: TLabel;
    lbl182: TLabel;
    lbl184: TLabel;
    lbl171: TLabel;
    lbl173: TLabel;
    lbl175: TLabel;
    lbl177: TLabel;
    lbl179: TLabel;
    lbl181: TLabel;
    lbl183: TLabel;
    lbl185: TLabel;
    lbl186: TLabel;
    lbl187: TLabel;
    lbl188: TLabel;
    lbl189: TLabel;
    lbl190: TLabel;
    lbl191: TLabel;
    lbl192: TLabel;
    lbl193: TLabel;
    lbl195: TLabel;
    lbl196: TLabel;
    lbl197: TLabel;
    lbl198: TLabel;
    lbl199: TLabel;
    lbl200: TLabel;
    lbl201: TLabel;
    lbl194: TLabel;
    lbl202: TLabel;
    lbl203: TLabel;
    lbl204: TLabel;
    lbl205: TLabel;
    lbl206: TLabel;
    lbl207: TLabel;
    lbl208: TLabel;
    lbl209: TLabel;
    lbl210: TLabel;
    lbl211: TLabel;
    lbl212: TLabel;
    lbl213: TLabel;
    lbl214: TLabel;
    lbl215: TLabel;
    lbl216: TLabel;
    lvSectionTable: TListView;
    pnl2: TPanel;
    pnl3: TPanel;
    chk1: TCheckBox;
    chk2: TCheckBox;
    chk3: TCheckBox;
    chk4: TCheckBox;
    chk5: TCheckBox;
    chk6: TCheckBox;
    chk7: TCheckBox;
    chk8: TCheckBox;
    chk9: TCheckBox;
    chk10: TCheckBox;
    chk11: TCheckBox;
    chk12: TCheckBox;
    chk13: TCheckBox;
    chk14: TCheckBox;
    chk15: TCheckBox;
    lvSectionData: TListView;
    tsExport: TTabSheet;
    tsImport: TTabSheet;
    tsResource: TTabSheet;
    lvFunc: TListView;
    grpExport: TGroupBox;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    Label11: TLabel;
    Label12: TLabel;
    Label13: TLabel;
    Label14: TLabel;
    Label15: TLabel;
    Label16: TLabel;
    Label17: TLabel;
    Label18: TLabel;
    Label19: TLabel;
    Label20: TLabel;
    Label21: TLabel;
    Label22: TLabel;
    lvImportDllFileName: TListView;
    lvImportFuncList: TListView;
    pnlHex: TPanel;
    btnMagic: TButton;
    pmDosMagic: TPopupMenu;
    IMAGEDOSSIGNATURE5A4DMZ1: TMenuItem;
    IMAGEOS2SIGNATURE454ENE1: TMenuItem;
    IMAGEOS2SIGNATURELE454CLE1: TMenuItem;
    IMAGEVXDSIGNATURE454CLE1: TMenuItem;
    btnMachine: TButton;
    pmOptionHeaderMagic: TPopupMenu;
    IMAGENTOPTIONALHDR32MAGIC010B1: TMenuItem;
    IMAGENTOPTIONALHDR64MAGIC020B1: TMenuItem;
    IMAGEROMOPTIONALHDRMAGIC01071: TMenuItem;
    pmOptionHeaderSubSystem: TPopupMenu;
    IMAGESUBSYSTEMUNKNOWN0Unknownsubsystem1: TMenuItem;
    IMAGESUBSYSTEMNATIVE1Imagedoesntrequireasubsystem1: TMenuItem;
    IMAGESUBSYSTEMWINDOWSGUI2ImagerunsintheWindowsGUIsubsystem1: TMenuItem;
    IMAGESUBSYSTEMWINDOWSCUI3ImagerunsintheWindowscharactersubsystem1: TMenuItem;
    IMAGESUBSYSTEMOS2CUI5imagerunsintheOS2charactersubsystem1: TMenuItem;
    IMAGESUBSYSTEMPOSIXCUI7imagerunsinthePosixcharactersubsystem1: TMenuItem;
    IMAGESUBSYSTEMNATIVEWINDOWS8imageisanativeWin9xdriver1: TMenuItem;
    IMAGESUBSYSTEMWINDOWSCEGUI9ImagerunsintheWindowsCEsubsystem1: TMenuItem;
    IMAGESUBSYSTEMEFIAPPLICATION101: TMenuItem;
    IMAGESUBSYSTEMEFIBOOTSERVICEDRIVER111: TMenuItem;
    IMAGESUBSYSTEMEFIRUNTIMEDRIVER121: TMenuItem;
    IMAGESUBSYSTEMEFIROM131: TMenuItem;
    IMAGESUBSYSTEMXBOX141: TMenuItem;
    IMAGESUBSYSTEMWINDOWSBOOTAPPLICATION161: TMenuItem;
    pmDll: TPopupMenu;
    IMAGEDLLCHARACTERISTICSDYNAMICBASE0x0040DLLcanmove1: TMenuItem;
    IMAGEDLLCHARACTERISTICSFORCEINTEGRITY0x0080CodeIntegrityImage1: TMenuItem;
    IMAGEDLLCHARACTERISTICSNXCOMPAT0x0100ImageisNXcompatible1: TMenuItem;
    IMAGEDLLCHARACTERISTICSNOISOLATION0x0200Imageunderstandsisolationanddoesntwantit1: TMenuItem;
    IMAGEDLLCHARACTERISTICSNOSEH0x0400ImagedoesnotuseSEHNoSEhandlermayresidei1: TMenuItem;
    IMAGEDLLCHARACTERISTICSNOBIND0x0800Donotbindthisimage1: TMenuItem;
    N0x1000Reserved1: TMenuItem;
    IMAGEDLLCHARACTERISTICSWDMDRIVER0x2000DriverusesWDMmodel1: TMenuItem;
    N0x4000Reserved1: TMenuItem;
    mniIMAGEDLLCHARACTERISTICSTERMINALSERVERAWARE0x80001: TMenuItem;
    btnOptionHeaderMagic: TButton;
    btnSubSystem: TButton;
    pmMachine: TPopupMenu;
    MenuItem1: TMenuItem;
    MenuItem2: TMenuItem;
    MenuItem3: TMenuItem;
    MenuItem4: TMenuItem;
    MenuItem5: TMenuItem;
    MenuItem6: TMenuItem;
    MenuItem7: TMenuItem;
    MenuItem8: TMenuItem;
    MenuItem9: TMenuItem;
    MenuItem10: TMenuItem;
    btnDll: TButton;
    tsDosStub: TTabSheet;
    tvResource: TTreeView;
    procedure srchbxSelectPEFileInvokeSearch(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure lvSectionTableClick(Sender: TObject);
    procedure lvSectionDataClick(Sender: TObject);
    procedure lbl154Click(Sender: TObject);
    procedure lbl156Click(Sender: TObject);
    procedure lbl158Click(Sender: TObject);
    procedure scrlbx1MouseWheel(Sender: TObject; Shift: TShiftState; WheelDelta: Integer; MousePos: TPoint; var Handled: Boolean);
    procedure lvFuncClick(Sender: TObject);
    procedure lvImportDllFileNameClick(Sender: TObject);
    procedure lvImportFuncListClick(Sender: TObject);
    procedure pgcPEInfoChange(Sender: TObject);
    procedure btnMagicClick(Sender: TObject);
    procedure btnMachineClick(Sender: TObject);
    procedure btnOptionHeaderMagicClick(Sender: TObject);
    procedure btnSubSystemClick(Sender: TObject);
    procedure IMAGESUBSYSTEMWINDOWSBOOTAPPLICATION161DrawItem(Sender: TObject; ACanvas: TCanvas; ARect: TRect; Selected: Boolean);
    procedure IMAGESUBSYSTEMWINDOWSBOOTAPPLICATION161MeasureItem(Sender: TObject; ACanvas: TCanvas; var Width, Height: Integer);
    procedure IMAGEFILEMACHINEAMD648664AMD64K81DrawItem(Sender: TObject; ACanvas: TCanvas; ARect: TRect; Selected: Boolean);
    procedure IMAGEFILEMACHINEAMD648664AMD64K81MeasureItem(Sender: TObject; ACanvas: TCanvas; var Width, Height: Integer);
    procedure IMAGEVXDSIGNATURE454CLE1DrawItem(Sender: TObject; ACanvas: TCanvas; ARect: TRect; Selected: Boolean);
    procedure IMAGEVXDSIGNATURE454CLE1MeasureItem(Sender: TObject; ACanvas: TCanvas; var Width, Height: Integer);
    procedure mniIMAGEDLLCHARACTERISTICSTERMINALSERVERAWARE0x80001DrawItem(Sender: TObject; ACanvas: TCanvas; ARect: TRect; Selected: Boolean);
    procedure mniIMAGEDLLCHARACTERISTICSTERMINALSERVERAWARE0x80001MeasureItem(Sender: TObject; ACanvas: TCanvas; var Width, Height: Integer);
    procedure btnDllClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    FbX64   : Boolean;
    FTempHex: THexEditor;
    procedure SelectNodeRect(const intStart, intEnd: Integer);
    { PE 简介 }
    procedure AnalyzePE_Note();
    { PE DOS Header }
    procedure AnalyzePE_DosHeader;
    { PE NT Header }
    procedure AnalyzePE_NTHeader;
    { PE NT Header  X64 }
    procedure AnalyzePE_NTHeader_X64;
    { PE NT Header X86 }
    procedure AnalyzePE_NTHeader_X86;
    { PE Section Table }
    procedure AnalyzePE_SectionTable;
    { PE Export Data }
    procedure AnalyzePE_ExportFunc;
    { PE Import Data }
    procedure AnalyzePE_ImportFunc;
    { PE Resource Data }
    procedure AnalyzePE_Resource;
    procedure GetPESectionsInfo(var sts: TArray<TImageSectionHeader>; var intVA: Cardinal; const intDataDirectoryIndex: Integer = 0);
    procedure GetDllFuncList(const intVA, intVA1, intVA2, intRA: Integer);
  public
    { Public declarations }
  end;

procedure db_ShowDllForm_Plugins(var frm: TFormClass; var ft: TSPFileType; var strParentModuleName, strModuleName, strClassName, strWindowName, strIconFileName: PAnsiChar; const bShow: Boolean = True); stdcall;

implementation

{$R *.dfm}

procedure db_ShowDllForm_Plugins(var frm: TFormClass; var ft: TSPFileType; var strParentModuleName, strModuleName, strClassName, strWindowName, strIconFileName: PAnsiChar; const bShow: Boolean = True); stdcall;
begin
  frm                     := TfrmPEInfo;
  ft                      := ftDelphiDll;
  strParentModuleName     := '程序员工具';
  strModuleName           := 'PE查看器';
  strIconFileName         := '';
  strClassName            := '';
  strWindowName           := '';
  Application.Handle      := GetMainFormApplication.Handle;
  Application.Icon.Handle := GetMainFormApplication.Icon.Handle;
end;

procedure TfrmPEInfo.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  FTempHex.Free;
  Action := caFree;
end;

procedure TfrmPEInfo.FormCreate(Sender: TObject);
begin
  FTempHex                := THexEditor.Create(pnlHex);
  FTempHex.Parent         := pnlHex;
  FTempHex.Align          := alClient;
  FTempHex.BorderStyle    := bsNone;
  FTempHex.Ctl3D          := False;
  FTempHex.Visible        := False;
  FTempHex.OffsetDisplay  := odHex;
  FTempHex.BytesPerColumn := 1;
  FTempHex.BytesPerLine   := 16;
  FTempHex.Font.Size      := 5;

  FbX64                     := False;
  pgcPEInfo.ActivePageIndex := 0;
  pgcPEInfo.Enabled         := False;
end;

procedure AddSubDosHeader(tnDosHeader: TTreeNode; tv: TTreeView);
begin
  tv.Items.AddChild(tnDosHeader, 'e_magic');
  tv.Items.AddChild(tnDosHeader, 'e_cblp');
  tv.Items.AddChild(tnDosHeader, 'e_cp');
  tv.Items.AddChild(tnDosHeader, 'e_crlc');
  tv.Items.AddChild(tnDosHeader, 'e_cparhdr');
  tv.Items.AddChild(tnDosHeader, 'e_minalloc');
  tv.Items.AddChild(tnDosHeader, 'e_maxalloc');
  tv.Items.AddChild(tnDosHeader, 'e_ss');
  tv.Items.AddChild(tnDosHeader, 'e_sp');
  tv.Items.AddChild(tnDosHeader, 'e_csum');
  tv.Items.AddChild(tnDosHeader, 'e_ip');
  tv.Items.AddChild(tnDosHeader, 'e_cs');
  tv.Items.AddChild(tnDosHeader, 'e_lfarlc');
  tv.Items.AddChild(tnDosHeader, 'e_ovno');
  tv.Items.AddChild(tnDosHeader, 'e_res');
  tv.Items.AddChild(tnDosHeader, 'e_oemid');
  tv.Items.AddChild(tnDosHeader, 'e_oeminfo');
  tv.Items.AddChild(tnDosHeader, 'e_res2');
  tv.Items.AddChild(tnDosHeader, '_lfanew');
end;

procedure AddSubPEHeader(tnPEHeader: TTreeNode; tv: TTreeView);
begin
  tv.Items.AddChild(tnPEHeader, 'Signature');
  tv.Items.AddChild(tnPEHeader, 'File Header');
  tv.Items.AddChild(tnPEHeader, 'OptionalHeader');
end;

procedure TfrmPEInfo.GetPESectionsInfo(var sts: TArray<TImageSectionHeader>; var intVA: Cardinal; const intDataDirectoryIndex: Integer = 0);
var
  intOffset       : DWORD;
  intLen          : Integer;
  ntHeader32      : PImageNtHeaders32;
  ntHeader64      : PImageNtHeaders64;
  intSectionTables: Integer;
begin
  intLen     := 4;
  intOffset  := PDWORD(FTempHex.BufferFromFile(SizeOf(TImageDosHeader) - 4, intLen))^;
  intLen     := SizeOf(TImageNtHeaders32);
  ntHeader32 := PImageNtHeaders32(FTempHex.BufferFromFile(intOffset, intLen));
  if ntHeader32^.FileHeader.Machine = IMAGE_FILE_MACHINE_AMD64 then
  begin
    ntHeader64       := PImageNtHeaders64(FTempHex.BufferFromFile(intOffset, intLen));
    intSectionTables := ntHeader64^.FileHeader.NumberOfSections;
    SetLength(sts, intSectionTables);
    intLen := intSectionTables * SizeOf(TImageSectionHeader);
    CopyMemory(@sts[0], FTempHex.BufferFromFile(intOffset + SizeOf(TImageNtHeaders64), intLen), intLen);
    if intDataDirectoryIndex <> -1 then
      intVA := ntHeader64^.OptionalHeader.DataDirectory[intDataDirectoryIndex].VirtualAddress
    else
      intVA := 0;
  end
  else
  begin
    intSectionTables := ntHeader32.FileHeader.NumberOfSections;
    SetLength(sts, intSectionTables);
    intLen := intSectionTables * SizeOf(TImageSectionHeader);
    CopyMemory(@sts[0], FTempHex.BufferFromFile(intOffset + SizeOf(TImageNtHeaders32), intLen), intLen);
    if intDataDirectoryIndex <> -1 then
      intVA := ntHeader32^.OptionalHeader.DataDirectory[intDataDirectoryIndex].VirtualAddress
    else
      intVA := 0;
  end;
end;

procedure TfrmPEInfo.IMAGEFILEMACHINEAMD648664AMD64K81DrawItem(Sender: TObject; ACanvas: TCanvas; ARect: TRect; Selected: Boolean);
begin
  ACanvas.Font.Name := '宋体';
  ACanvas.Font.Size := 11;
  ACanvas.TextOut(ARect.Left, ARect.Top, (Sender as TMenuItem).Caption);
end;

procedure TfrmPEInfo.IMAGEFILEMACHINEAMD648664AMD64K81MeasureItem(Sender: TObject; ACanvas: TCanvas; var Width, Height: Integer);
begin
  Width := 700;
end;

procedure TfrmPEInfo.IMAGESUBSYSTEMWINDOWSBOOTAPPLICATION161DrawItem(Sender: TObject; ACanvas: TCanvas; ARect: TRect; Selected: Boolean);
begin
  ACanvas.Font.Name := '宋体';
  ACanvas.Font.Size := 11;
  ACanvas.TextOut(ARect.Left, ARect.Top, (Sender as TMenuItem).Caption);
end;

procedure TfrmPEInfo.IMAGESUBSYSTEMWINDOWSBOOTAPPLICATION161MeasureItem(Sender: TObject; ACanvas: TCanvas; var Width, Height: Integer);
begin
  Width := 800;
end;

procedure TfrmPEInfo.IMAGEVXDSIGNATURE454CLE1DrawItem(Sender: TObject; ACanvas: TCanvas; ARect: TRect; Selected: Boolean);
begin
  ACanvas.Font.Name := '宋体';
  ACanvas.Font.Size := 11;
  ACanvas.TextOut(ARect.Left, ARect.Top, (Sender as TMenuItem).Caption);
end;

procedure TfrmPEInfo.IMAGEVXDSIGNATURE454CLE1MeasureItem(Sender: TObject; ACanvas: TCanvas; var Width, Height: Integer);
begin
  Width := 300;
end;

{ PE 简介 }
procedure TfrmPEInfo.AnalyzePE_Note();
var
  intOffset: DWORD;
  intLen   : Integer;
  ntHeader : PImageNtHeaders32;
begin
  intLen    := 4;
  intOffset := PDWORD(FTempHex.BufferFromFile(SizeOf(TImageDosHeader) - 4, intLen))^;
  intLen    := SizeOf(TImageNtHeaders32);
  ntHeader  := PImageNtHeaders32(FTempHex.BufferFromFile(intOffset, intLen));
  FbX64     := ntHeader^.FileHeader.Machine = IMAGE_FILE_MACHINE_AMD64;
  if FbX64 then
  begin
    if ntHeader^.FileHeader.Characteristics and IMAGE_FILE_DLL = IMAGE_FILE_DLL then
      lbl57.Caption := '64位DLL'
    else
      lbl57.Caption := '64位EXE';
  end
  else
  begin
    if ntHeader^.FileHeader.Characteristics and IMAGE_FILE_DLL = IMAGE_FILE_DLL then
      lbl57.Caption := '32位DLL'
    else
      lbl57.Caption := '32位EXE';
  end;
end;

{ PE DOS Header }
procedure TfrmPEInfo.AnalyzePE_DosHeader;
var
  idh   : PImageDosHeader;
  intLen: Integer;
begin
  intLen        := SizeOf(TImageDosHeader);
  idh           := PImageDosHeader(FTempHex.BufferFromFile(0, intLen));
  lbl17.Caption := Format('$%0.4x', [idh^.e_magic]);
  lbl21.Caption := Format('$%0.4x', [idh^.e_cblp]);
  lbl22.Caption := Format('$%0.4x', [idh^.e_cp]);
  lbl23.Caption := Format('$%0.4x', [idh^.e_crlc]);
  lbl24.Caption := Format('$%0.4x', [idh^.e_cparhdr]);
  lbl25.Caption := Format('$%0.4x', [idh^.e_minalloc]);
  lbl26.Caption := Format('$%0.4x', [idh^.e_maxalloc]);
  lbl27.Caption := Format('$%0.4x', [idh^.e_ss]);
  lbl28.Caption := Format('$%0.4x', [idh^.e_sp]);
  lbl29.Caption := Format('$%0.4x', [idh^.e_csum]);
  lbl30.Caption := Format('$%0.4x', [idh^.e_ip]);
  lbl31.Caption := Format('$%0.4x', [idh^.e_cs]);
  lbl32.Caption := Format('$%0.4x', [idh^.e_lfarlc]);
  lbl33.Caption := Format('$%0.4x', [idh^.e_ovno]);
  lbl34.Caption := Format('$%0.8x', [0]);
  lbl35.Caption := Format('$%0.4x', [idh^.e_oemid]);
  lbl36.Caption := Format('$%0.4x', [idh^.e_oeminfo]);
  lbl37.Caption := Format('$%0.8x%0.8x%0.4x', [0, 0, 0]);
  lbl38.Caption := Format('$%0.8x', [idh^._lfanew]);
end;

{ PE NT Header  X64 }
procedure TfrmPEInfo.AnalyzePE_NTHeader_X64;
var
  intLen   : Integer;
  intOffset: DWORD;
  inh      : PImageNtHeaders64;
  I        : Integer;
begin
  intLen    := 4;
  intOffset := PDWORD(FTempHex.BufferFromFile(SizeOf(TImageDosHeader) - 4, intLen))^;
  intLen    := SizeOf(TImageNtHeaders64);
  inh       := PImageNtHeaders64(FTempHex.BufferFromFile(intOffset, intLen));

  lbl67.Caption := Format('$%0.8x', [inh^.Signature]);
  lbl68.Caption := Format('$%0.4x', [inh^.FileHeader.Machine]);
  lbl69.Caption := Format('$%0.4x', [inh^.FileHeader.NumberOfSections]);
  lbl70.Caption := Format('$%0.8x', [inh^.FileHeader.TimeDateStamp]);
  lbl71.Caption := Format('$%0.8x', [inh^.FileHeader.PointerToSymbolTable]);
  lbl72.Caption := Format('$%0.8x', [inh^.FileHeader.NumberOfSymbols]);
  lbl73.Caption := Format('$%0.4x', [inh^.FileHeader.SizeOfOptionalHeader]);
  lbl74.Caption := Format('$%0.4x', [inh^.FileHeader.Characteristics]);

  lbl124.Caption := Format('$%0.4x', [inh^.OptionalHeader.Magic]);
  lbl125.Caption := Format('$%0.2x', [inh^.OptionalHeader.MajorLinkerVersion]);
  lbl126.Caption := Format('$%0.2x', [inh^.OptionalHeader.MinorLinkerVersion]);
  lbl127.Caption := Format('$%0.8x', [inh^.OptionalHeader.SizeOfCode]);
  lbl128.Caption := Format('$%0.8x', [inh^.OptionalHeader.SizeOfInitializedData]);
  lbl129.Caption := Format('$%0.8x', [inh^.OptionalHeader.SizeOfUninitializedData]);
  lbl130.Caption := Format('$%0.8x', [inh^.OptionalHeader.AddressOfEntryPoint]);
  lbl131.Caption := Format('$%0.8x', [inh^.OptionalHeader.BaseOfCode]);
  lbl132.Visible := False;
  lbl85.Visible  := False;
  lbl133.Caption := Format('$%0.8x', [inh^.OptionalHeader.ImageBase]);
  lbl134.Caption := Format('$%0.8x', [inh^.OptionalHeader.SectionAlignment]);
  lbl135.Caption := Format('$%0.8x', [inh^.OptionalHeader.FileAlignment]);
  lbl136.Caption := Format('$%0.4x', [inh^.OptionalHeader.MajorOperatingSystemVersion]);
  lbl137.Caption := Format('$%0.4x', [inh^.OptionalHeader.MinorOperatingSystemVersion]);
  lbl138.Caption := Format('$%0.4x', [inh^.OptionalHeader.MajorImageVersion]);
  lbl139.Caption := Format('$%0.4x', [inh^.OptionalHeader.MinorImageVersion]);
  lbl140.Caption := Format('$%0.4x', [inh^.OptionalHeader.MajorSubsystemVersion]);
  lbl141.Caption := Format('$%0.4x', [inh^.OptionalHeader.MinorSubsystemVersion]);
  lbl142.Caption := Format('$%0.8x', [inh^.OptionalHeader.Win32VersionValue]);
  lbl143.Caption := Format('$%0.8x', [inh^.OptionalHeader.SizeOfImage]);
  lbl144.Caption := Format('$%0.4x', [inh^.OptionalHeader.SizeOfHeaders]);
  lbl145.Caption := Format('$%0.8x', [inh^.OptionalHeader.CheckSum]);
  lbl146.Caption := Format('$%0.4x', [inh^.OptionalHeader.Subsystem]);
  lbl147.Caption := Format('$%0.4x', [inh^.OptionalHeader.DllCharacteristics]);
  lbl148.Caption := Format('$%0.8x', [inh^.OptionalHeader.SizeOfStackReserve]);
  lbl149.Caption := Format('$%0.8x', [inh^.OptionalHeader.SizeOfStackCommit]);
  lbl150.Caption := Format('$%0.8x', [inh^.OptionalHeader.SizeOfHeapReserve]);
  lbl151.Caption := Format('$%0.8x', [inh^.OptionalHeader.SizeOfHeapCommit]);
  lbl152.Caption := Format('$%0.8x', [inh^.OptionalHeader.LoaderFlags]);
  lbl153.Caption := Format('$%0.8x', [inh^.OptionalHeader.NumberOfRvaAndSizes]);
  for I          := 0 to 15 do
  begin
    TLabel(FindComponent('lbl' + IntToStr(154 + 2 * I + 0))).Caption := Format('$%0.8x', [inh^.OptionalHeader.DataDirectory[I].VirtualAddress]);
    TLabel(FindComponent('lbl' + IntToStr(154 + 2 * I + 1))).Caption := Format('$%0.8x', [inh^.OptionalHeader.DataDirectory[I].Size]);
  end;
end;

{ PE NT Header X86 }
procedure TfrmPEInfo.AnalyzePE_NTHeader_X86;
var
  intLen   : Integer;
  intOffset: DWORD;
  inh      : PImageNtHeaders32;
  I        : Integer;
begin
  intLen    := 4;
  intOffset := PDWORD(FTempHex.BufferFromFile(SizeOf(TImageDosHeader) - 4, intLen))^;
  intLen    := SizeOf(TImageNtHeaders32);
  inh       := PImageNtHeaders32(FTempHex.BufferFromFile(intOffset, intLen));

  lbl67.Caption := Format('$%0.8x', [inh^.Signature]);
  lbl68.Caption := Format('$%0.4x', [inh^.FileHeader.Machine]);
  lbl69.Caption := Format('$%0.4x', [inh^.FileHeader.NumberOfSections]);
  lbl70.Caption := Format('$%0.8x', [inh^.FileHeader.TimeDateStamp]);
  lbl71.Caption := Format('$%0.8x', [inh^.FileHeader.PointerToSymbolTable]);
  lbl72.Caption := Format('$%0.8x', [inh^.FileHeader.NumberOfSymbols]);
  lbl73.Caption := Format('$%0.4x', [inh^.FileHeader.SizeOfOptionalHeader]);
  lbl74.Caption := Format('$%0.4x', [inh^.FileHeader.Characteristics]);

  lbl124.Caption := Format('$%0.4x', [inh^.OptionalHeader.Magic]);
  lbl125.Caption := Format('$%0.2x', [inh^.OptionalHeader.MajorLinkerVersion]);
  lbl126.Caption := Format('$%0.2x', [inh^.OptionalHeader.MinorLinkerVersion]);
  lbl127.Caption := Format('$%0.8x', [inh^.OptionalHeader.SizeOfCode]);
  lbl128.Caption := Format('$%0.8x', [inh^.OptionalHeader.SizeOfInitializedData]);
  lbl129.Caption := Format('$%0.8x', [inh^.OptionalHeader.SizeOfUninitializedData]);
  lbl130.Caption := Format('$%0.8x', [inh^.OptionalHeader.AddressOfEntryPoint]);
  lbl131.Caption := Format('$%0.8x', [inh^.OptionalHeader.BaseOfCode]);
  lbl132.Caption := Format('$%0.8x', [inh.OptionalHeader.BaseOfData]);
  lbl132.Visible := True;
  lbl85.Visible  := True;
  lbl133.Caption := Format('$%0.8x', [inh^.OptionalHeader.ImageBase]);
  lbl134.Caption := Format('$%0.8x', [inh^.OptionalHeader.SectionAlignment]);
  lbl135.Caption := Format('$%0.8x', [inh^.OptionalHeader.FileAlignment]);
  lbl136.Caption := Format('$%0.4x', [inh^.OptionalHeader.MajorOperatingSystemVersion]);
  lbl137.Caption := Format('$%0.4x', [inh^.OptionalHeader.MinorOperatingSystemVersion]);
  lbl138.Caption := Format('$%0.4x', [inh^.OptionalHeader.MajorImageVersion]);
  lbl139.Caption := Format('$%0.4x', [inh^.OptionalHeader.MinorImageVersion]);
  lbl140.Caption := Format('$%0.4x', [inh^.OptionalHeader.MajorSubsystemVersion]);
  lbl141.Caption := Format('$%0.4x', [inh^.OptionalHeader.MinorSubsystemVersion]);
  lbl142.Caption := Format('$%0.8x', [inh^.OptionalHeader.Win32VersionValue]);
  lbl143.Caption := Format('$%0.8x', [inh^.OptionalHeader.SizeOfImage]);
  lbl144.Caption := Format('$%0.4x', [inh^.OptionalHeader.SizeOfHeaders]);
  lbl145.Caption := Format('$%0.8x', [inh^.OptionalHeader.CheckSum]);
  lbl146.Caption := Format('$%0.4x', [inh^.OptionalHeader.Subsystem]);
  lbl147.Caption := Format('$%0.4x', [inh^.OptionalHeader.DllCharacteristics]);
  lbl148.Caption := Format('$%0.8x', [inh^.OptionalHeader.SizeOfStackReserve]);
  lbl149.Caption := Format('$%0.8x', [inh^.OptionalHeader.SizeOfStackCommit]);
  lbl150.Caption := Format('$%0.8x', [inh^.OptionalHeader.SizeOfHeapReserve]);
  lbl151.Caption := Format('$%0.8x', [inh^.OptionalHeader.SizeOfHeapCommit]);
  lbl152.Caption := Format('$%0.8x', [inh^.OptionalHeader.LoaderFlags]);
  lbl153.Caption := Format('$%0.8x', [inh^.OptionalHeader.NumberOfRvaAndSizes]);
  for I          := 0 to 15 do
  begin
    TLabel(FindComponent('lbl' + IntToStr(154 + 2 * I + 0))).Caption := Format('$%0.8x', [inh^.OptionalHeader.DataDirectory[I].VirtualAddress]);
    TLabel(FindComponent('lbl' + IntToStr(154 + 2 * I + 1))).Caption := Format('$%0.8x', [inh^.OptionalHeader.DataDirectory[I].Size]);
  end;
end;

{ PE NT Header }
procedure TfrmPEInfo.AnalyzePE_NTHeader;
var
  intOffset: DWORD;
  intLen   : Integer;
  ntHeader : PImageNtHeaders32;
begin
  intLen    := 4;
  intOffset := PDWORD(FTempHex.BufferFromFile(SizeOf(TImageDosHeader) - 4, intLen))^;
  intLen    := SizeOf(TImageNtHeaders32);
  ntHeader  := PImageNtHeaders32(FTempHex.BufferFromFile(intOffset, intLen));
  if ntHeader^.FileHeader.Machine = IMAGE_FILE_MACHINE_AMD64 then
    AnalyzePE_NTHeader_X64
  else
    AnalyzePE_NTHeader_X86;
end;

{ PE Section Table }
procedure TfrmPEInfo.AnalyzePE_SectionTable;
var
  intSectionTables: Integer;
  sts             : TArray<TImageSectionHeader>;
  I               : Integer;
  arrName         : TBytes;
  intVA           : Cardinal;
begin
  GetPESectionsInfo(sts, intVA, -1);
  intSectionTables := Length(sts);
  lvSectionTable.Items.Clear;
  for I := 0 to intSectionTables - 1 do
  begin
    with lvSectionTable.Items.Add do
    begin
      SetLength(arrName, 8);
      CopyMemory(@arrName[0], @sts[I].Name[0], 8);
      Caption := StringOf(arrName);
      SubItems.Add(Format('$%0.8x', [sts[I].Misc.PhysicalAddress]));
      SubItems.Add(Format('$%0.8x', [sts[I].VirtualAddress]));
      SubItems.Add(Format('$%0.8x', [sts[I].SizeOfRawData]));
      SubItems.Add(Format('$%0.8x', [sts[I].PointerToRawData]));
      SubItems.Add(Format('$%0.8x', [sts[I].Characteristics]));
    end;
  end;

  lvSectionData.Items.Clear;
  for I := 0 to intSectionTables - 1 do
  begin
    with lvSectionData.Items.Add do
    begin
      SetLength(arrName, 8);
      CopyMemory(@arrName[0], @sts[I].Name[0], 8);
      Caption := StringOf(arrName);
      SubItems.Add(Format('$%0.8x', [sts[I].Misc.PhysicalAddress]));
      SubItems.Add(Format('$%0.8x', [sts[I].VirtualAddress]));
      SubItems.Add(Format('$%0.8x', [sts[I].SizeOfRawData]));
      SubItems.Add(Format('$%0.8x', [sts[I].PointerToRawData]));
      SubItems.Add(Format('$%0.8x', [sts[I].Characteristics]));
    end;
  end;
end;

procedure TfrmPEInfo.btnDllClick(Sender: TObject);
var
  pt: TPoint;
begin
  GetCursorPos(pt);
  pmDll.Popup(pt.x, pt.y);
end;

procedure TfrmPEInfo.btnMachineClick(Sender: TObject);
var
  pt: TPoint;
begin
  GetCursorPos(pt);
  pmMachine.Popup(pt.x, pt.y);
end;

procedure TfrmPEInfo.btnMagicClick(Sender: TObject);
var
  pt: TPoint;
begin
  GetCursorPos(pt);
  pmDosMagic.Popup(pt.x, pt.y);
end;

procedure TfrmPEInfo.btnOptionHeaderMagicClick(Sender: TObject);
var
  pt: TPoint;
begin
  GetCursorPos(pt);
  pmOptionHeaderMagic.Popup(pt.x, pt.y);
end;

procedure TfrmPEInfo.btnSubSystemClick(Sender: TObject);
var
  pt: TPoint;
begin
  GetCursorPos(pt);
  pmOptionHeaderSubSystem.Popup(pt.x, pt.y);
end;

procedure TfrmPEInfo.lvSectionTableClick(Sender: TObject);
var
  I      : Integer;
  intTag : int64;
  chkTemp: TCheckBox;
  intAttr: Integer;
begin
  if lvSectionTable.ItemIndex = -1 then
    Exit;

  intAttr := StrToInt(lvSectionTable.Selected.SubItems[4]);
  for I   := 1 to 15 do
  begin
    chkTemp := TCheckBox(FindComponent('chk' + IntToStr(I)));
    if I <> 4 then
    begin
      chkTemp.Checked := chkTemp.Tag and intAttr = chkTemp.Tag;
    end
    else
    begin
      intTag          := chk4.Tag + 1;
      chkTemp.Checked := intTag and intAttr = intTag;
    end;
  end;
end;

procedure TfrmPEInfo.mniIMAGEDLLCHARACTERISTICSTERMINALSERVERAWARE0x80001DrawItem(Sender: TObject; ACanvas: TCanvas; ARect: TRect; Selected: Boolean);
begin
  ACanvas.Font.Name := '宋体';
  ACanvas.Font.Size := 11;
  ACanvas.TextOut(ARect.Left, ARect.Top, (Sender as TMenuItem).Caption);
end;

procedure TfrmPEInfo.mniIMAGEDLLCHARACTERISTICSTERMINALSERVERAWARE0x80001MeasureItem(Sender: TObject; ACanvas: TCanvas; var Width, Height: Integer);
begin
  Width := 1050;
end;

procedure TfrmPEInfo.scrlbx1MouseWheel(Sender: TObject; Shift: TShiftState; WheelDelta: Integer; MousePos: TPoint; var Handled: Boolean);
begin
  if WheelDelta < 0 then
    SendMessage(scrlbx1.Handle, WM_VSCROLL, SB_LINEDOWN, 0) // 向下滚
  else
    SendMessage(scrlbx1.Handle, WM_VSCROLL, SB_LINEUP, 0); // 向上滚
end;

procedure TfrmPEInfo.lvFuncClick(Sender: TObject);
var
  intOffset: Integer;
begin
  if lvFunc.ItemIndex = -1 then
    Exit;

  intOffset := StrToInt(lvFunc.Items[lvFunc.ItemIndex].SubItems[1]);
  SelectNodeRect(intOffset, intOffset + Length(lvFunc.Items[lvFunc.ItemIndex].SubItems[2]) - 1);
end;

procedure TfrmPEInfo.lvSectionDataClick(Sender: TObject);
var
  intOffset: Integer;
  intSize  : Integer;
begin
  if lvSectionData.ItemIndex = -1 then
    Exit;

  intOffset := StrToInt(lvSectionData.Items[lvSectionData.ItemIndex].SubItems[3]);
  intSize   := StrToInt(lvSectionData.Items[lvSectionData.ItemIndex].SubItems[2]) + StrToInt(lvSectionData.Items[lvSectionData.ItemIndex].SubItems[3]);
  if intSize = 0 then
    intSize := 1;
  SelectNodeRect(intOffset, intSize - 1);
end;

procedure TfrmPEInfo.srchbxSelectPEFileInvokeSearch(Sender: TObject);
begin
  if not dlgOpenSelectPEFile.Execute(GetMainFormHandle) then
    Exit;

  lvImportDllFileName.Clear;
  srchbxSelectPEFile.Text := dlgOpenSelectPEFile.FileName;
  FTempHex.LoadFromFile(dlgOpenSelectPEFile.FileName);
  FTempHex.Visible := True;
  AnalyzePE_Note();
  pgcPEInfo.ActivePage := tsAll;
  pgcPEInfo.Enabled    := True;
end;

procedure TfrmPEInfo.SelectNodeRect(const intStart, intEnd: Integer);
begin
  FTempHex.SelStart := intStart;
  FTempHex.SelEnd   := intEnd;
end;

procedure TfrmPEInfo.pgcPEInfoChange(Sender: TObject);
var
  intOffset: DWORD;
  intLen   : Integer;
  ntHeader : PImageNtHeaders32;
begin
  if pgcPEInfo.ActivePage = tsAll then
  begin
    SelectNodeRect(0, 0);
  end;

  if pgcPEInfo.ActivePage = tsDosHeader then
  begin
    SelectNodeRect(0, SizeOf(TImageDosHeader) - 1);
    AnalyzePE_DosHeader;
  end;

  if pgcPEInfo.ActivePage = tsDosStub then
  begin
    intLen := 4;
    SelectNodeRect(SizeOf(TImageDosHeader), PDWORD(FTempHex.BufferFromFile(SizeOf(TImageDosHeader) - 4, intLen))^ - 1);
  end;

  if pgcPEInfo.ActivePage = tsNTHeader then
  begin
    intLen    := 4;
    intOffset := PDWORD(FTempHex.BufferFromFile(SizeOf(TImageDosHeader) - 4, intLen))^;
    intLen    := SizeOf(TImageNtHeaders32);
    ntHeader  := PImageNtHeaders32(FTempHex.BufferFromFile(intOffset, intLen));
    if ntHeader^.FileHeader.Machine = IMAGE_FILE_MACHINE_AMD64 then
      SelectNodeRect(intOffset, intOffset + SizeOf(TImageNtHeaders64) - 1)
    else
      SelectNodeRect(intOffset, intOffset + SizeOf(TImageNtHeaders32) - 1);
    AnalyzePE_NTHeader;
  end;

  if (pgcPEInfo.ActivePage = tsSectionTable) or (pgcPEInfo.ActivePage = tsSectionData) then
  begin
    intLen    := 4;
    intOffset := PDWORD(FTempHex.BufferFromFile(SizeOf(TImageDosHeader) - 4, intLen))^;
    intLen    := SizeOf(TImageNtHeaders32);
    ntHeader  := PImageNtHeaders32(FTempHex.BufferFromFile(intOffset, intLen));
    if ntHeader^.FileHeader.Machine = IMAGE_FILE_MACHINE_AMD64 then
      SelectNodeRect(intOffset + SizeOf(TImageNtHeaders64), ntHeader^.FileHeader.NumberOfSections * SizeOf(TImageSectionHeader) + intOffset + SizeOf(TImageNtHeaders64) - 1)
    else
      SelectNodeRect(intOffset + SizeOf(TImageNtHeaders32), ntHeader^.FileHeader.NumberOfSections * SizeOf(TImageSectionHeader) + intOffset + SizeOf(TImageNtHeaders32) - 1);
    AnalyzePE_SectionTable;
  end;

  if pgcPEInfo.ActivePage = tsExport then
  begin
    AnalyzePE_ExportFunc;
  end;

  if pgcPEInfo.ActivePage = tsImport then
  begin
    AnalyzePE_ImportFunc;
  end;

  if pgcPEInfo.ActivePage = tsResource then
  begin
    AnalyzePE_Resource;
  end;
end;

procedure FindTableFilePos(const sts: array of TImageSectionHeader; const intVA: Cardinal; var intRA: Cardinal);
var
  III, Count: Integer;
begin
  intRA := 0;

  Count   := Length(sts);
  for III := 0 to Count - 2 do
  begin
    if (intVA >= sts[III + 0].VirtualAddress) and (intVA < sts[III + 1].VirtualAddress) then
    begin
      intRA := (intVA - sts[III].VirtualAddress) + sts[III].PointerToRawData;
      Break;
    end;
  end;
end;

{ PE Export Data }
procedure TfrmPEInfo.AnalyzePE_ExportFunc;
var
  intLen         : Integer;
  sts            : TArray<TImageSectionHeader>;
  I              : Integer;
  intVA          : Cardinal;
  intRA          : Cardinal;
  eft            : PImageExportDirectory;
  arrDllFileName : array [0 .. 255] of AnsiChar;
  arrFunctionName: array [0 .. 255] of AnsiChar;
  intFuncRA      : Cardinal;
  strFuncName    : String;
begin
  GetPESectionsInfo(sts, intVA, 0);
  if intVA = 0 then
  begin
    grpExport.Visible := False;
    lvFunc.Visible    := False;
    Exit;
  end;

  grpExport.Visible := True;
  lvFunc.Visible    := True;
  intLen            := SizeOf(TImageExportDirectory);
  FindTableFilePos(sts, intVA, intRA);
  SelectNodeRect(intRA, intRA + Cardinal(intLen) - 1);

  eft             := PImageExportDirectory(FTempHex.BufferFromFile(intRA, intLen));
  Label12.Caption := Format('$%0.8x', [intRA]);
  Label11.Caption := Format('$%0.8x', [eft^.Characteristics]);
  Label10.Caption := Format('$%0.8x', [eft^.Base]);
  Label9.Caption  := Format('$%0.8x', [eft^.Name]);
  intLen          := 256;
  CopyMemory(@arrDllFileName[0], FTempHex.BufferFromFile(eft^.Name - intVA + intRA, intLen), intLen);
  Label8.Caption  := String((arrDllFileName));
  Label18.Caption := Format('$%0.8x', [eft^.NumberOfFunctions]);
  Label19.Caption := Format('$%0.8x', [eft^.NumberOfNames]);
  Label20.Caption := Format('$%0.8x', [eft^.AddressOfFunctions]);
  Label21.Caption := Format('$%0.8x', [eft^.AddressOfNames]);
  Label22.Caption := Format('$%0.8x', [eft^.AddressOfNameOrdinals]);

  lvFunc.Items.Clear;
  for I := 0 to eft.NumberOfNames - 1 do
  begin
    intLen    := 4;
    intFuncRA := PInteger(FTempHex.BufferFromFile(eft.AddressOfNames - intVA + intRA + DWORD(4 * I), intLen))^;

    intLen := 256;
    CopyMemory(@arrFunctionName[0], FTempHex.BufferFromFile(intFuncRA - intVA + intRA, intLen), intLen);

    with lvFunc.Items.Add do
    begin
      Caption := Format('%0.2d', [I + Integer(eft^.Base)]);
      SubItems.Add(Format('$%0.8x', [eft^.AddressOfNames - intVA + intRA + DWORD(4 * I)]));
      SubItems.Add(Format('$%0.8x', [intFuncRA - intVA + intRA]));
      strFuncName := string(arrFunctionName);
      SubItems.Add(GetFullFuncNameCpp(strFuncName));
    end;
  end;
end;

procedure TfrmPEInfo.GetDllFuncList(const intVA, intVA1, intVA2, intRA: Integer);
var
  intEntry   : Cardinal;
  itd        : PImageThunkData;
  intLen     : Integer;
  intIndex   : Integer;
  pFuncName  : PImageImportByName;
  arrFuncName: array [0 .. 255] of AnsiChar;
  strFuncName: String;
  intFlag    : NativeUInt;
begin
  intFlag := Ifthen(FbX64, IMAGE_ORDINAL_FLAG64, IMAGE_ORDINAL_FLAG32);

  lvImportFuncList.Clear;
  intLen := SizeOf(TImageThunkData);
  if intVA1 <> 0 then
    intEntry := intVA1 - intVA + intRA
  else
    intEntry := intVA2 - intVA + intRA;

  intIndex := 0;
  itd      := PImageThunkData(FTempHex.BufferFromFile(intEntry, intLen));
  while itd^.AddressOfData <> 0 do
  begin

    with lvImportFuncList.Items.Add do
    begin
      Caption := '$' + IntToHex(intVA1 + intIndex * intLen);
      SubItems.Add('$' + IntToHex(itd^.AddressOfData));

      if itd^.Ordinal and intFlag = intFlag then
      begin
        Inc(intIndex);
        intLen := SizeOf(TImageThunkData);
        itd    := PImageThunkData(FTempHex.BufferFromFile(intEntry + Cardinal(intIndex * intLen), intLen));
        SubItems.Add('$' + IntToHex(itd^.Ordinal));
        SubItems.Add('N/A');
      end
      else
      begin
        intLen    := SizeOf(TImageImportByName);
        pFuncName := PImageImportByName(FTempHex.BufferFromFile(Integer(itd^.AddressOfData) - intVA + intRA, intLen));
        SubItems.Add('$' + IntToHex(pFuncName^.Hint));

        intLen := 256;
        CopyMemory(@arrFuncName[0], FTempHex.BufferFromFile(Integer(itd^.AddressOfData) - intVA + intRA + 2, intLen), 256);
        strFuncName := string(arrFuncName);
        SubItems.Add(GetFullFuncNameCpp(strFuncName));
      end;
    end;

    Inc(intIndex);
    intLen := SizeOf(TImageThunkData);
    itd    := PImageThunkData(FTempHex.BufferFromFile(intEntry + Cardinal(intIndex * intLen), intLen));
  end;
end;

procedure TfrmPEInfo.lvImportDllFileNameClick(Sender: TObject);
var
  sts  : TArray<TImageSectionHeader>;
  intVA: Cardinal;
  intRA: Cardinal;
begin
  if lvImportDllFileName.ItemIndex = -1 then
    Exit;

  GetPESectionsInfo(sts, intVA, 1);
  FindTableFilePos(sts, intVA, intRA);
  SelectNodeRect(                                                                           //
    intRA + Cardinal((lvImportDllFileName.ItemIndex + 0) * SizeOf(TImageImportDescriptor)), //
    intRA + Cardinal((lvImportDllFileName.ItemIndex + 1) * SizeOf(TImageImportDescriptor)) - 1);
  GetDllFuncList(intVA, StrToInt(lvImportDllFileName.Items[lvImportDllFileName.ItemIndex].SubItems[0]), StrToInt(lvImportDllFileName.Items[lvImportDllFileName.ItemIndex].SubItems[4]), intRA);
end;

procedure TfrmPEInfo.lvImportFuncListClick(Sender: TObject);
var
  sts  : TArray<TImageSectionHeader>;
  intVA: Cardinal;
  intRA: Cardinal;
begin
  if lvImportFuncList.ItemIndex = -1 then
    Exit;

  if lvImportFuncList.Items[lvImportFuncList.ItemIndex].SubItems[2] = 'N/A' then
    Exit;

  GetPESectionsInfo(sts, intVA, 1);
  FindTableFilePos(sts, intVA, intRA);
  SelectNodeRect(                                                                                           //
    intRA + Cardinal(StrToInt(lvImportFuncList.Items[lvImportFuncList.ItemIndex].SubItems[0])) - intVA + 2, //
    intRA + Cardinal(StrToInt(lvImportFuncList.Items[lvImportFuncList.ItemIndex].SubItems[0])) - intVA + 2 + Cardinal(lvImportFuncList.Items[lvImportFuncList.ItemIndex].SubItems[2].Length) - 1);
end;

{ PE Import Data }
procedure TfrmPEInfo.AnalyzePE_ImportFunc;
var
  sts           : TArray<TImageSectionHeader>;
  intVA         : Cardinal;
  intRA         : Cardinal;
  intLen        : Integer;
  imageEntry    : PImageImportDescriptor;
  intIndex      : Integer;
  arrDllFileName: array [0 .. 255] of AnsiChar;
begin
  GetPESectionsInfo(sts, intVA, 1);
  if intVA = 0 then
    Exit;

  FindTableFilePos(sts, intVA, intRA);
  SelectNodeRect(intRA, intRA + Cardinal(SizeOf(TImageImportDescriptor)) - 1);

  intIndex   := 0;
  intLen     := SizeOf(TImageImportDescriptor);
  imageEntry := PImageImportDescriptor(FTempHex.BufferFromFile(intRA, intLen));
  while imageEntry^.Name <> 0 do
  begin
    intLen := 256;
    CopyMemory(@arrDllFileName[0], FTempHex.BufferFromFile(imageEntry^.Name - intVA + intRA, intLen), intLen);
    with lvImportDllFileName.Items.Add do
    begin
      Caption := string(arrDllFileName);
      SubItems.Add('$' + IntToHex(imageEntry^.Characteristics, 8));
      SubItems.Add('$' + IntToHex(imageEntry^.TimeDateStamp, 8));
      SubItems.Add('$' + IntToHex(imageEntry^.ForwarderChain, 8));
      SubItems.Add('$' + IntToHex(imageEntry^.Name, 8));
      SubItems.Add('$' + IntToHex(imageEntry^.FirstThunk, 8));
    end;
    Inc(intIndex);
    intLen     := SizeOf(TImageImportDescriptor);
    imageEntry := PImageImportDescriptor(FTempHex.BufferFromFile(intRA + Cardinal(intIndex * SizeOf(TImageImportDescriptor)), intLen));
  end;
end;

{ PE Resource Data }
procedure TfrmPEInfo.AnalyzePE_Resource;
var
  resNode: TTreeNode;
begin
  tvResource.Items.Clear;
  resNode := tvResource.Items.Add(nil, 'Resource');
  LoadPEResource(srchbxSelectPEFile.Text, resNode, True, True);
  tvResource.Items[0].Expanded := True;
end;

{ 导出表 }
procedure TfrmPEInfo.lbl154Click(Sender: TObject);
begin
  if StrToInt(TLabel(Sender).Caption) = 0 then
  begin
    ShowMessage('没有导出表');
    Exit;
  end;

  AnalyzePE_ExportFunc;
  pgcPEInfo.ActivePage := tsExport;
end;

{ 导入表 }
procedure TfrmPEInfo.lbl156Click(Sender: TObject);
begin
  if StrToInt(TLabel(Sender).Caption) = 0 then
  begin
    ShowMessage('没有导入表');
    Exit;
  end;

  AnalyzePE_ImportFunc;
  pgcPEInfo.ActivePage := tsImport;
end;

{ 资源表 }
procedure TfrmPEInfo.lbl158Click(Sender: TObject);
begin
  if StrToInt(TLabel(Sender).Caption) = 0 then
  begin
    ShowMessage('没有资源表');
    Exit;
  end;

  AnalyzePE_Resource;
  pgcPEInfo.ActivePage := tsResource;
end;

end.
