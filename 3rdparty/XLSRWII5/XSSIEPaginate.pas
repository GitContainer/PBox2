unit XSSIEPaginate;

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

uses { Delphi } Classes, SysUtils, Contnrs, Math,
     { AXWord } XSSIEDefs, XSSIEUtils, XSSIEDocProps, XSSIECharRun, XSSIEPhyRow,
                XSSIELogParas, XSSIEGDIText,
                XBookPaintGDI2;

type TAXWWordDataNewCharRunEvent = procedure(var CRIndex: integer; var CR: TAXWCharRun) of object;

type TAXWParaObjSize = class(TObject)
protected
     FWidth     : double;
     FHeight    : double;
     FCurrHeight: double;
     FAlignRight: boolean;
public
     property Width: double read FWidth write FWidth;
     property Height: double read FHeight write FHeight;
     property CurrHeight: double read FCurrHeight write FCurrHeight;
     property AlignRigth: boolean read FAlignRight write FAlignRight;
     end;

type TAXWParaObjSizes = class(TObjectList)
private
     function  GetItems(Index: integer): TAXWParaObjSize;
     function  GetWidth: double;
     function  GetXOffset: double;
protected
public
     procedure Add(const AWidth,AHeight: integer; AAlignRigth: boolean);
     procedure Dec(const AHeight: double);

     property Width: double read GetWidth;
     property XOffset: double read GetXOffset;
     property Items[Index: integer]: TAXWParaObjSize read GetItems;
     end;

type TAXWWordData = class(TObject)
private
     procedure SetBreakChar(const Value: AxUCChar);
protected
     FPara                   : TAXWLogPara;
     FArea                   : TAXWClientArea;
     FPrint                  : TAXWTextPrint;
     FCR1,FCR2               : TAXWCharRun;
     FCRPos1,FCRPos2         : integer;
     FLeadingSpaceCount      : integer;
     FLeadingSpaceWidth      : double;
     FCharCount              : integer;
     FWidth                  : double;
     FTrailingSpaceCount     : integer;
     FTrailingSpaceWidth     : double;
     FHasLineBreak           : boolean;
     FCollectBreakChar       : boolean;
     FBreakChar              : AxUCChar;
     FBreakCharWidth         : double;
     FLeadingBreakChars,
     FTrailingBreakChars     : array of TBreakCharData;
     FLeadingBreakCharsCount,
     FTrailingBreakCharsCount: integer;
     FNewCharRunEvent        : TAXWWordDataNewCharRunEvent;
     FTabStop                : TAXWTabStop;

     function  TextWidth(const AText: AxUCString): double;
     procedure AddLeadingBC;
     procedure AddTrailingBC;
     function  CharIsWhitespaceLeading(C: AxUCChar): boolean;
     function  CharIsWhitespaceTrailing(C: AxUCChar): boolean;
public
     constructor Create(Print: TAXWTextPrint; Area: TAXWClientArea);

     procedure Clear(ACRPos: integer);

     function  TotWidth: double; {$ifdef D2006PLUS} inline; {$endif}
     function  WidthBeginning: double; {$ifdef D2006PLUS} inline; {$endif}
     function  WidthEnd: double; {$ifdef D2006PLUS} inline; {$endif}
     function  TotCharCount: integer; {$ifdef D2006PLUS} inline; {$endif}
     function  NextWord(var CR: TAXWCharRun; var Source: AxUCString; var CRIndex: integer): boolean;

     property Para: TAXWLogPara read FPara write FPara;

     property CR1: TAXWCharRun read FCR1;
     property CR2: TAXWCharRun read FCR2;
     property CRPos1: integer read FCRPos1;
     property CRPos2: integer read FCRPos2;

     property LeadingSpaceCount: integer read FLeadingSpaceCount;
     property LeadingSpaceWidth: double read FLeadingSpaceWidth;
     property CharCount: integer read FCharCount;
     property Width: double read FWidth;
     property TrailingSpaceCount: integer read FTrailingSpaceCount;
     property TrailingSpaceWidth: double read FTrailingSpaceWidth;
     property HasLineBreak: boolean read FHasLineBreak;
     property CollectBreakChar: boolean read FCollectBreakChar write FCollectBreakChar;
     property BreakChar: AxUCChar read FBreakChar write SetBreakChar;
     property LeadingBreakCharsCount: integer read FLeadingBreakCharsCount;
     property TrailingBreakCharsCount: integer read FTrailingBreakCharsCount;

     property OnNewCharRun: TAXWWordDataNewCharRunEvent read FNewCharRunEvent write FNewCharRunEvent;
     end;

type TAXWRowPaginateStack = class(TObject)
protected
     FGDI: TAXWGDI;
     FArea: TAXWClientArea;
     FWordData: TAXWWordData;
     FPara: TAXWLogPara;
     FRow: TAXWPhyRow;
     FRowWidth: double;
     FRowIndex: integer;
     FLogPos: integer;
     FDeltaChars: integer;
     FMaxAscent: double;
     FMaxDescent: double;
     FBreakChars: array of TBreakCharData;
     FBreakCharsCount: integer;
     FLastTrailingWidth: double;
     FLastTrailingBreakCharsCount: integer;
     FKeepTrailingSpaces: boolean;
     FParaWidth: TAXWParaObjSizes;

     procedure AddBreakChars(Chars: array of TBreakCharData; CharCount: integer);
public
     constructor Create(Print: TAXWTextPrint; Area: TAXWClientArea);
     destructor Destroy; override;

     procedure Clear(Para: TAXWLogPara);
     procedure Push;
     procedure BeginRow(CR: TAXWCharRun; CRPos: integer; ANewPara: boolean = False);
     procedure UpdateRow(CR: TAXWCharRun; CRPos: integer);
     procedure EndRow; overload;
     procedure EndRow(CR: TAXWCharRun; CRPos: integer); overload;
     procedure NewCharRun;
     procedure Cleanup;

     property Para: TAXWLogPara read FPara;
     property Row: TAXWPhyRow read FRow;
     property Curr: TAXWWordData read FWordData;
     property MaxAscent: double read FMaxAscent write FMaxAscent;
     property MaxDescent: double read FMaxDescent write FMaxDescent;
     property DeltaChars: integer read FDeltaChars write FDeltaChars;
     end;

type TAXWPaginator = class(TObject)
private
     function  GetKeepTrailingSpaces: boolean;
     procedure SetKeepTrailingSpaces(const Value: boolean);
protected
     FPrint: TAXWTextPrint;
     FArea: TAXWClientArea;

     FText: AxUCString;
     FRowStack: TAXWRowPaginateStack;
     FSpaceWidths: array of integer;
     FSpaceWidthsCount: integer;

     function  BeginCharRun(Index: integer): TAXWCharRun;
     procedure Cleanup;
     procedure WordDataNewCharRun(var CRIndex: integer; var CR: TAXWCharRun);
public
     constructor Create(Print: TAXWTextPrint; Area: TAXWClientArea);
     destructor Destroy; override;
     procedure Clear;

     procedure Paginate(Para: TAXWLogPara);

     property KeepTrailingSpaces: boolean read GetKeepTrailingSpaces write SetKeepTrailingSpaces;
     end;


implementation

{ TAXWPaginator }

procedure TAXWPaginator.Paginate(Para: TAXWLogPara);
var
  i: integer;
  R: TAXWPhyRow;
  CR: TAXWCharRun;
begin
  Clear;

  if Para.Empty then begin
    R := Para.Rows.Add;
    FPrint.SetupFormat(Para.CHP,Nil);
    FRowStack.FGDI.GetTextMetric;
    R.Height := FPrint.GDI.TM.Height;
    Exit;
  end;

  Para.Rows.ClearFirstLast;

  FRowStack.Clear(Para);

  FRowStack.BeginRow(Para[0],1,True);
  i := 0;
  while i < Para.Count do begin
    CR := BeginCharRun(i);
    FRowStack.Curr.Clear(1);

    while FRowStack.Curr.NextWord(CR,FText,i) do
      FRowStack.Push;

    Inc(i);
  end;
  FRowStack.EndRow(Para.CRLast,Para.CRLast.Len);

//  Para.LogPara.SetRowsDim(Para);

  Para.Rows.SetFirstLast;
  Cleanup;
end;

function TAXWPaginator.BeginCharRun(Index: integer): TAXWCharRun;
begin
  Result := FRowStack.Para[Index];

  FText := Result.Text;
  FPrint.SetupFormat(FRowStack.Para.CHP,Result.CHPX);
  FPrint.GDI.GetTextMetric;
  FRowStack.NewCharRun;
  FRowStack.MaxDescent := Max(FRowStack.MaxDescent,FPrint.GDI.TM.Descent);
  FRowStack.MaxAscent := Max(FRowStack.MaxAscent,FPrint.GDI.TM.Ascent);
end;

procedure TAXWPaginator.Cleanup;
begin
  FArea.GraphicWidth := 0;
{
  while FCurrPara.Count > (FCurrRowIndex + 1) do
    FCurrPara.Delete(FCurrPara.Count - 1);
}
end;

procedure TAXWPaginator.Clear;
begin
//  FillChar(FTextmetric,SizeOf(TEXTMETRIC),#0);
end;

constructor TAXWPaginator.Create(Print: TAXWTextPrint; Area: TAXWClientArea);
begin
  FPrint := Print;
  FArea := Area;
  FRowStack := TAXWRowPaginateStack.Create(FPrint,FArea);
  FRowStack.Curr.OnNewCharRun := WordDataNewCharRun;
end;

destructor TAXWPaginator.Destroy;
begin
  FreeAndNil(FRowStack);
  inherited;
end;

function TAXWPaginator.GetKeepTrailingSpaces: boolean;
begin
  Result := FRowStack.FKeepTrailingSpaces;
end;

procedure TAXWPaginator.SetKeepTrailingSpaces(const Value: boolean);
begin
  FRowStack.FKeepTrailingSpaces := Value;
end;

procedure TAXWPaginator.WordDataNewCharRun(var CRIndex: integer; var CR: TAXWCharRun);
begin
  CR := BeginCharRun(CRIndex);
end;

{ TAXWWordData }

procedure TAXWWordData.AddLeadingBC;
begin
  if FLeadingBreakCharsCount >= High(FLeadingBreakChars) then
    SetLength(FLeadingBreakChars,Length(FLeadingBreakChars) + 10);
  FLeadingBreakChars[FLeadingBreakCharsCount].Char := FBreakChar;
  FLeadingBreakChars[FLeadingBreakCharsCount].Width := FBreakCharWidth;
  Inc(FLeadingBreakCharsCount);
end;

procedure TAXWWordData.AddTrailingBC;
begin
  if FTrailingBreakCharsCount >= High(FTrailingBreakChars) then
    SetLength(FTrailingBreakChars,Length(FTrailingBreakChars) + 10);
  FTrailingBreakChars[FTrailingBreakCharsCount].Char := FBreakChar;
  FTrailingBreakChars[FTrailingBreakCharsCount].Width := FBreakCharWidth;
  Inc(FTrailingBreakCharsCount);
end;

procedure TAXWWordData.Clear(ACRPos: integer);
begin
  FCRPos1 := ACRPos;
  FCRPos2 := ACRPos;
  FCR1 := Nil;
  FCR2 := Nil;
  FLeadingSpaceCount := 0;
  FLeadingSpaceWidth := 0;
  FCharCount := 0;
  FWidth := 0;
  FTrailingSpaceCount := 0;
  FTrailingSpaceWidth := 0;
  FTabStop := Nil;
//  FWidthAfterTab := 0;
//  FCollectBreakChar := False;
//  FBreakChar := #0;
end;

constructor TAXWWordData.Create(Print: TAXWTextPrint; Area: TAXWClientArea);
begin
  FPrint := Print;
  FArea := Area;
  SetLength(FLeadingBreakChars,10);
  SetLength(FTrailingBreakChars,10);
end;

function TAXWWordData.CharIsWhitespaceLeading(C: AxUCChar): boolean;
begin
  if FCollectBreakChar and (C = FBreakChar) then begin
    AddLeadingBC;
    Result := True;
  end
  else
    Result := CharIsWhitespace(C);
end;

function TAXWWordData.CharIsWhitespaceTrailing(C: AxUCChar): boolean;
begin
  if FCollectBreakChar and (C = FBreakChar) then begin
    AddTrailingBC;
    Result := True;
  end
  else
    Result := CharIsWhitespace(C);
end;

function TAXWWordData.NextWord(var CR: TAXWCharRun; var Source: AxUCString; var CRIndex: integer): boolean;
var
  S: AxUCString;
  i,j,p: integer;
  W,CharW: double;
  Done: boolean;
begin
  Result := Source <> '';
  if not Result then
    Exit;

  FLeadingBreakCharsCount := 0;
  FTrailingBreakCharsCount := 0;

  FCR1 := CR;
  FCRPos1 := FCRPos2;

  i := 1;
  while (i <= Length(Source)) and CharIsWhitespaceLeading(Source[i]) do
    Inc(i);
  if i > 1 then begin
    FLeadingSpaceCount := i - 1;
    FLeadingSpaceWidth := TextWidth(Copy(Source,1,i - 1));
  end
  else begin
    FLeadingSpaceCount := 0;
    FLeadingSpaceWidth := 0;
  end;

  FHasLineBreak := False;
  while (i <= Length(Source)) and not CharIsWhitespace(Source[i]) do begin
    if Source[i] = AXW_CHAR_SPECIAL_HARDLINEBREAK then begin
      FHasLineBreak := True;
      Inc(i);
      Break;
    end;
    Inc(i);
    if CharIsBreakingChar(Source[i]) then begin
      Inc(i);
      Break;
    end;
  end;

  S := Copy(Source,FLeadingSpaceCount + 1,i - FLeadingSpaceCount - 1);
  W := TextWidth(S);

  if W > FArea.PaginateWidth then begin
    j := Length(S);
    while (j >= 1) and (W > FArea.PaginateWidth) do begin
      CharW := TextWidth(S[j]);
      W := W - CharW;
      Dec(j);
      Dec(i);
    end;
  end;

  FWidth := W;
  FCharCount := i - FLeadingSpaceCount - 1;

  if (i > Length(Source)) and (acrfPartOfNext in CR.Flags) then begin
    Done := False;
    repeat
      Inc(CRIndex);
      FNewCharRunEvent(CRIndex,CR);
      Source := CR.Text;
      i := 1;
      FCRPos2 := 1;

      while (i <= Length(Source)) and not CharIsWhitespace(Source[i]) do begin
        if Source[i] = AXW_CHAR_SPECIAL_HARDLINEBREAK then begin
          FHasLineBreak := True;
          Inc(i);
          Break;
        end;
        Inc(i);
      end;

      S := Copy(Source,1,i - 1);
      W := TextWidth(S);

      if (W + FWidth) > FArea.PaginateWidth then begin
        j := Length(S);
        while (j >= 1) and ((W + FWidth) > FArea.PaginateWidth) do begin
          CharW := TextWidth(S[j]);
          W := W - CharW;
          Dec(j);
          Dec(i);
        end;
        Done := True;
      end;

      Inc(FCharCount,i - 1);
      FWidth := FWidth + W;
    until not (acrfNoWhitespace in CR.Flags) or Done;
  end;

  j := i;
  while (i <= Length(Source)) and CharIsWhitespaceTrailing(Source[i]) do
    Inc(i);
  FTrailingSpaceCount := i - j;
  p := i - FTrailingSpaceCount;
  FTrailingSpaceWidth := TextWidth(Copy(Source,p,i - p));

  if FHasLineBreak then
    Inc(FCRPos2,i - 2)
  else
    Inc(FCRPos2,i - 1);

  FCR2 := CR;

  if i <= Length(Source) then
    Source := Copy(Source,i,MAXINT)
  else
    Source := '';
end;

procedure TAXWWordData.SetBreakChar(const Value: AxUCChar);
begin
  FBreakChar := Value;
  FBreakCharWidth := TextWidth(Value);
end;

function TAXWWordData.TextWidth(const AText: AxUCString): double;
var
  L: integer;
begin
  L := Length(AText);
  while (L >= 1) and (AText[L] = AXW_CHAR_SPECIAL_HARDLINEBREAK) do
    Dec(L);

  Result := FPrint.GDI.TextWidthF(Copy(AText,1,L));
end;

function TAXWWordData.TotCharCount: integer;
begin
  Result := FLeadingSpaceCount + FCharCount + FTrailingSpaceCount;
end;

function TAXWWordData.TotWidth: double;
begin
  Result := FLeadingSpaceWidth + FWidth + FTrailingSpaceWidth;
end;

function TAXWWordData.WidthBeginning: double;
begin
  Result := FLeadingSpaceWidth + FWidth;
end;

function TAXWWordData.WidthEnd: double;
begin
  Result := FWidth + FTrailingSpaceWidth;
end;

{ TAXWRowPaginateStack }

procedure TAXWRowPaginateStack.AddBreakChars(Chars: array of TBreakCharData; CharCount: integer);
begin
  if CharCount > 0 then begin
    if (CharCount + FBreakCharsCount) > Length(FBreakChars) then
      SetLength(FBreakChars,Length(FBreakChars) + CharCount);
    System.Move(Chars[0],FBreakChars[FBreakCharsCount],CharCount * SizeOf(TBreakCharData));
    Inc(FBreakCharsCount,CharCount);
  end;
end;

procedure TAXWRowPaginateStack.BeginRow(CR: TAXWCharRun; CRPos: integer; ANewPara: boolean = False);
begin
  FRowWidth := 0;

  FBreakCharsCount := 0;
  if ANewPara then begin
    FMaxAscent := 0;
    FMaxDescent := 0;
  end
  else begin
    FMaxAscent := FGDI.TM.Ascent;
    FMaxDescent := FGDI.TM.Descent;
  end;

  Inc(FRowIndex);
  if FRowIndex < FPara.Rows.Count then begin
    FRow := FPara.Rows[FRowIndex];
  end
  else
    FRow := FPara.Rows.Add;
  if (FRow.LogPos1 <> (FLogPos - FDeltaChars)) then
    FRow.Flags := [aprfDirty];

{$ifdef AXW_DEBUG}
  FRow.Flags := FRow.Flags + [aprfDebug];
{$endif}

  FRow.CR1 := CR;
  FRow.CRPos1 := CRPos;

  FRow.LogPos1 := FLogPos;
end;

procedure TAXWRowPaginateStack.Cleanup;
begin
  while FPara.Rows.Count > (FRowIndex + 1) do
    FPara.Rows.Delete(FPara.Rows.Count - 1);
end;

constructor TAXWRowPaginateStack.Create(Print: TAXWTextPrint; Area: TAXWClientArea);
begin
  inherited Create;
  FGDI := Print.GDI;
  FArea := Area;
  FWordData := TAXWWordData.Create(Print,FArea);
  FParaWidth := TAXWParaObjSizes.Create;
end;

destructor TAXWRowPaginateStack.Destroy;
begin
  FreeAndNil(FWordData);
  FreeAndNil(FParaWidth);
  inherited;
end;

procedure TAXWRowPaginateStack.EndRow;
var
  NewRow: TAXWPhyRowComplex;
begin
//  if (FRow.Ascent + FRow.Descent) <= 0 then begin
//    FRow.Ascent := 14;
//    FRow.Descent := 8;
//    FRow.Height := FRow.Ascent + FRow.Descent;
//  end;

  if not FKeepTrailingSpaces then
    FRow.Width := FRow.Width - FLastTrailingWidth;
  Dec(FBreakCharsCount,FLastTrailingBreakCharsCount);
  // Single line paras are not justified.
  if FBreakCharsCount > 0 then begin
    if not (FRow.SameCR and (FWordData.BreakChar = AXW_CHAR_BREAKCHAR)) then begin
      SetLength(FBreakChars,FBreakCharsCount);
      NewRow := TAXWPhyRowComplex.Create(FBreakChars);
      NewRow.Assign(FRow);
      NewRow.Flags := NewRow.Flags + [aprfComplex];
      // Replace will free FRow (owned TObjectList).
      FPara.Rows.Replace(FRowIndex,NewRow);
      FRow := NewRow;
    end;
    FRow.BreakCharCount := FBreakCharsCount;
  end;

  FPara.Rows.UpdateArea;

  FRow.XOffset := FParaWidth.XOffset;
  if (FRowIndex = 0) and (FPara.PAPX.IndentHanging <> 0) then
    FParaWidth.Add(FGDI.PtToPixX(FPara.PAPX.Indent + FPara.PAPX.IndentHanging),MAXINT,False);
  FParaWidth.Dec(FRow.Height);
  FArea.GraphicWidth := Round(FParaWidth.Width);
end;

procedure TAXWRowPaginateStack.EndRow(CR: TAXWCharRun; CRPos: integer);
begin
  FRow.Height := (FMaxAscent + FMaxDescent); //FMaxHeight;
  FRow.Ascent := FMaxAscent;
  FRow.Descent := FMaxDescent;

  FRow.CR2 := CR;
  FRow.CRPos2 := CRPos;

  EndRow;
end;

procedure TAXWRowPaginateStack.Clear(Para: TAXWLogPara);
begin
  FPara := Para;
  FRowIndex := -1;
  FMaxAscent := 0;
  FMaxDescent := 0;
  FRowWidth := 0;
  FDeltaChars := 0;

  FWordData.Para := Para;

  FParaWidth.Clear;
  if FPara.PAPX.Indent <> 0 then
    FParaWidth.Add(FGDI.PtToPixX(FPara.PAPX.Indent),MAXINT,False);
  if FPara.PAPX.IndentRight <> 0 then
    FParaWidth.Add(FGDI.PtToPixX(FPara.PAPX.IndentRight),MAXINT,True);
  if FPara.PAPX.IndentFirstLine <> 0 then
    FParaWidth.Add(FGDI.PtToPixX(FPara.PAPX.Indent + FPara.PAPX.IndentFirstLine),1,False);

  FArea.GraphicWidth := Round(FParaWidth.Width);

  SetLength(FBreakChars,0);
  FBreakCharsCount := 0;
  FLastTrailingWidth := 0;
  FLastTrailingBreakCharsCount := 0;

  if FRowIndex < 0 then begin
    FRow := Nil;
    FLogPos := 1;
    FWordData.Clear(1);
    FPara.Rows.Clear;
  end
  else begin
    FRow := FPara.Rows[FRowIndex];
    FRow.Flags := FRow.Flags + [aprfDirty];
{$ifdef AXW_DEBUG}
    FRow.Flags := FRow.Flags + [aprfDebug];
{$endif}
    FWordData.Clear(FRow.CRPos1);
    FLogPos := FRow.LogPos1;
  end;
  FWordData.CollectBreakChar := Para.PAPX.Alignment = axptaJustify;
end;

procedure TAXWRowPaginateStack.NewCharRun;
begin
  FWordData.BreakChar := AxUCChar(FGDI.TM.BreakChar);
end;

procedure TAXWRowPaginateStack.Push;
begin
  if FWordData.HasLineBreak then begin
    FRowWidth := FRowWidth + FWordData.Width;
    UpdateRow(FWordData.CR1,FWordData.CRPos2 - 1);
    EndRow;
    BeginRow(FWordData.CR1,FWordData.CRPos2 + 1);
    FWordData.Clear(FWordData.CRPos2 + 1);
  end
  else if ((FRowWidth + FWordData.WidthBeginning) > FArea.PaginateWidth) and (FWordData.CharCount > 0) then begin
    if (FRowWidth + FWordData.WidthBeginning) > FArea.PaginateWidth then begin
      EndRow;
      BeginRow(FWordData.CR1,FWordData.CRPos1);
//      FRowWidth := FWordData.Width;
      UpdateRow(FWordData.CR2,FWordData.CRPos2 - 1);
    end
    else
      Exit;
  end;
  // Except for the first row in a para, leading spaces are not permitted at the
  // beginning of a row.

  if (FWordData.LeadingSpaceCount > 0) and (FRowIndex > 0) and (FRowWidth = 0) then begin
    FPara.Rows[FRowIndex - 1].CR2 := FPara.Rows[FRowIndex].CR1;
    FPara.Rows[FRowIndex - 1].CRPos2 := FWordData.LeadingSpaceCount;
    FPara.Rows[FRowIndex].CRPos1 := FPara.Rows[FRowIndex].CRPos1 + FWordData.LeadingSpaceCount;
    FPara.Rows[FRowIndex].LogPos1 := FPara.Rows[FRowIndex].LogPos1 + FWordData.LeadingSpaceCount;
    FRowWidth := FRowWidth + FWordData.WidthEnd;
    if FWordData.CollectBreakChar then
      AddBreakChars(FWordData.FTrailingBreakChars,FWordData.TrailingBreakCharsCount);
  end
  else begin
    FRowWidth := FRowWidth + FWordData.TotWidth;
    FLastTrailingWidth := FWordData.TrailingSpaceWidth;
    UpdateRow(FWordData.CR2,FWordData.CRPos2 - 1);
    if FWordData.CollectBreakChar then begin
      AddBreakChars(FWordData.FLeadingBreakChars,FWordData.LeadingBreakCharsCount);
      AddBreakChars(FWordData.FTrailingBreakChars,FWordData.TrailingBreakCharsCount);
      FLastTrailingBreakCharsCount := FWordData.TrailingBreakCharsCount;
    end;
  end;
  Inc(FLogPos,FWordData.TotCharCount);
end;

procedure TAXWRowPaginateStack.UpdateRow(CR: TAXWCharRun; CRPos: integer);
begin
  FRow.CR2 := CR;
  FRow.CRPos2 := CRPos;
  FRow.Height := (FMaxAscent + FMaxDescent); //FMaxHeight;
  FRow.Ascent := FMaxAscent;
  FRow.Descent := FMaxDescent;
  FRow.Width := FRowWidth;
end;

{ TAXWParaObjSizes }

procedure TAXWParaObjSizes.Add(const AWidth, AHeight: integer; AAlignRigth: boolean);
var
  Item: TAXWParaObjSize;
begin
  Item := TAXWParaObjSize.Create;
  Item.Width := AWidth;
  Item.Height := AHeight;
  Item.CurrHeight := AHeight;
  Item.AlignRigth := AAlignRigth;
  inherited Add(Item);
end;

procedure TAXWParaObjSizes.Dec(const AHeight: double);
var
  i: integer;
begin
  i := 0;
  while i < Count do begin
    Items[i].CurrHeight := Items[i].CurrHeight - AHeight;
    if Items[i].CurrHeight <= 0 then
      Delete(i)
    else
      Inc(i);
  end;
end;

function TAXWParaObjSizes.GetItems(Index: integer): TAXWParaObjSize;
begin
  Result := TAXWParaObjSize(inherited Items[Index]);
end;

function TAXWParaObjSizes.GetWidth: double;
var
  i: integer;
  L,R: double;
begin
  L := 0;
  R := 0;
  for i := 0 to Count - 1 do begin
    if Items[i].AlignRigth then
      R := Max(R,Items[i].Width)
    else
      L := Max(L,Items[i].Width);
  end;
  Result := L + R;
end;

function TAXWParaObjSizes.GetXOffset: double;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do begin
    if not Items[i].AlignRigth then
      Result := Max(Result,Items[i].Width);
  end;
end;

end.
