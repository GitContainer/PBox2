unit XLSTools5;

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
     Xc12Utils5, Xc12Common5, Xc12Manager5, Xc12DataStyleSheet5,
     XLSUtils5, XLSFormulaTypes5;

// <b> <i> <u> <s:12> <f:Font Name> <c:XXXXXX> <c:Auto>

const XLS_LT_PLACEHOLDER = #1;

type TXLSFontRunData = class(TObject)
end;

type TXLSFontRunBuilder = class(TObject)
protected
     FManager  : TXc12Manager;
     FStyles   : TXc12DataStyleSheet;

     FText     : AxUCString;
     FRuns     : TXc12DynFontRunArray;

     FBold     : boolean;
     FItalic   : boolean;
     FUnderline: boolean;
     FSize     : integer;
     FFontName : AxUCString;
     FColor    : TXc12RGBColor;

     function  MakeFont: TXc12Font;
     function  IncFontRuns: PXc12FontRun;
     function  ExtractTag(var AText: AxUCString; out ATag: AxUCString; out AIndex: integer): boolean;
     procedure DoSimpleTag(const ATag: AxUCString);
public
     constructor Create(AManager: TXc12Manager);

     procedure Clear;

     procedure FromSimpleTags(const AText: AxUCString);

     property Text: AxUCString read FText;
     property FontRuns: TXc12DynFontRunArray read FRuns;
     end;

type TXLSFormattedTextItem = class(TObject)
protected
     FText   : AxUCString;
     FFont   : TXc12Font;

     FNewLine: boolean;
public
     constructor Create(AText: AxUCString; AFont: TXc12Font);

     function AsHTML: AxUCString;
     function FontAsRTF: AxUCString;

     procedure AddSpecialCR(ACR: AxUCChar);

     property Text: AxUCString read FText;
     property Font: TXc12Font read FFont;
     // Only assigned when TXLSFormattedText.SplitAtCR is True.
     property NewLine  : boolean read FNewLine;
     end;

type TXLSFormattedText = class(TObjectList)
private
protected
     FFirstFont: TXc12Font;
     FSplitAtCR: boolean;

     function GetItems(Index: integer): TXLSFormattedTextItem;
public
     constructor Create;

     procedure Assign(AText: TXLSFormattedText);

     procedure Add(const AText: AxUCString; AFont: TXc12Font); overload;
     procedure Add(const AText: AxUCString; AFirstFont: TXc12Font; AFontRuns: PXc12FontRunArray; AFontRunsCount: integer); overload;

     // Set AIncludeFirstFont = True if the first font not is set elsewhere
     // such as in a <td> tag.
     function AsHTML(const AIncludeFirstFont: boolean = False): AxUCString;

     property SplitAtCR: boolean read FSplitAtCR write FSplitAtCR;
     property Items[Index: integer]: TXLSFormattedTextItem read GetItems; default;
     end;

implementation

{ TXLSFontRunBuilder }

procedure TXLSFontRunBuilder.Clear;
begin
  SetLength(FRuns,0);
  FBold := False;
  FItalic := False;
  FUnderline := False;
  FSize := 0;
  FFontName := '';
  FColor := XLSCOLOR_AUTO;
end;

constructor TXLSFontRunBuilder.Create(AManager: TXc12Manager);
begin
  FManager := AManager;
  FStyles := FManager.StyleSheet;

  Clear;
end;

procedure TXLSFontRunBuilder.DoSimpleTag(const ATag: AxUCString);
var
  Off: boolean;
  iVal: integer;
  sVal: AxUCString;
begin
  Off := Copy(ATag,1,1) = '/';
  if Off then begin
    if Length(ATag) <> 2 then
      FManager.Errors.Error('</>',XLSERR_INVALID_SIMPLE_TAG)
    else begin
      case ATag[2] of
        'b': FBold := False;
        'i': FItalic := False;
        'u': FUnderline := False;
        else FManager.Errors.Error('<' + ATag + '>',XLSERR_INVALID_SIMPLE_TAG);
      end;
    end;
  end
  else begin
    case ATag[1] of
      'b': FBold := True;
      'i': FItalic := True;
      'u': FUnderline := True;
      's': begin
        if TryStrToInt(Copy(ATag,3,MAXINT),iVal) then
          FSize := iVal
        else
          FManager.Errors.Error('<' + ATag + '>',XLSERR_INVALID_SIMPLE_TAG);
     end;
      'f': begin
        sVal := Copy(ATag,3,MAXINT);
        if Trim(sVal) <> '' then
          FFontName := sVal
        else
          FManager.Errors.Error('<' + ATag + '>',XLSERR_INVALID_SIMPLE_TAG);
      end;
      'c': begin
        sVal := Copy(ATag,3,MAXINT);
        if Lowercase(sVal) = 'auto' then
          FColor := XLSCOLOR_AUTO
        else if TryStrToInt('$' + Copy(ATag,3,MAXINT),iVal) then
          FColor := iVal
        else
          FManager.Errors.Error('<' + ATag + '>',XLSERR_INVALID_SIMPLE_TAG);
      end;
    end;
  end;
end;

function TXLSFontRunBuilder.ExtractTag(var AText: AxUCString; out ATag: AxUCString; out AIndex: integer): boolean;
var
  p1,p2: integer;
begin
  p1 := CPos('<',AText);
  p2 := CPos('>',AText);
  if (p1 >= 1) or (p2 >= 1) then begin
    if ((p1 >= 1) xor (p2 >= 1)) or (p2 < p1) then begin
      // Error;
      Result := False;
      Exit;
    end;
    ATag := Copy(AText,p1 + 1,p2 - p1 - 1);
    AText := Copy(AText,1,p1 - 1) + Copy(AText,p2 + 1,MAXINT);
    AIndex := p1;
    Result := True;
  end
  else
    Result := False;
end;

procedure TXLSFontRunBuilder.FromSimpleTags(const AText: AxUCString);
var
  i: integer;
  CurrI: integer;
  Tag: AxUCString;
  FR: PXc12FontRun;
begin
  FText := AText;
  repeat
    i := Pos('<>',FText);
    if i >= 1 then
      FText := Copy(FText,1,i - 1) + XLS_LT_PLACEHOLDER + Copy(FText,i + 2,MAXINT);
  until (i <= 0);

  CurrI := 1;
  while ExtractTag(FText,Tag,i) do begin
    if i = CurrI then begin
      CurrI := i;
    end else begin
      FR := IncFontRuns;
      FR.Index := CurrI - 1;
      FR.Font := MakeFont;
      CurrI := i;
    end;
    DoSimpleTag(Tag);
  end;
  if CurrI < Length(FText) then begin
    FR := IncFontRuns;
    FR.Index := CurrI - 1;
    FR.Font := MakeFont;
  end;

  for i := 1 to Length(FText) do begin
    if FText[i] = XLS_LT_PLACEHOLDER then
      FText[i] := '<';
  end;
end;

function TXLSFontRunBuilder.IncFontRuns: PXc12FontRun;
begin
  SetLength(FRuns,Length(FRuns) + 1);
  Result := @FRuns[High(FRuns)];
end;

function TXLSFontRunBuilder.MakeFont: TXc12Font;
var
  F: TXc12Font;
begin
  Result := TXc12Font.Create(Nil);
  Result.Assign(FStyles.Fonts.DefaultFont);
  if FBold then
    Result.Style := Result.Style + [xfsBold]
  else
    Result.Style := Result.Style - [xfsBold];
  if FItalic then
    Result.Style := Result.Style + [xfsItalic]
  else
    Result.Style := Result.Style - [xfsItalic];
  if FUnderline then
    Result.Underline := xulSingle
  else
    Result.Underline := xulNone;

  if FFontName <> '' then
    Result.Name := FFontName;

  if FSize > 0 then
    Result.Size := FSize;

  Result.Color := RGBColorToXc12(FColor);

  F := FStyles.Fonts.Find(Result);
  if F <> Nil then begin
    Result.Free;
    Result := F;
  end
  else
    FStyles.Fonts.Add(Result);

  FStyles.XFEditor.UseFont(Result);
end;

{ TXLSFormattedText }

function TXLSFormattedText.AsHTML(const AIncludeFirstFont: boolean = False): AxUCString;
var
  i: integer;
begin
  Result := '';

  if Count <= 0 then
    Exit;

  if AIncludeFirstFont or ((FFirstFont <> Nil) and (Items[0].Font.Index <> FFirstFont.Index)) then
    Result := Result + Items[0].AsHTML
  else
    Result := Result + AxUCString(XLSUTF8Encode(Items[0].Text));

  for i := 1 to Count - 1 do
    Result := Result + Items[i].AsHTML;
end;

procedure TXLSFormattedText.Assign(AText: TXLSFormattedText);
var
  i   : integer;
begin
  Clear;

  FFirstFont := AText.FFirstFont;
  FSplitAtCR := AText.FSplitAtCR;

  for i := 0 to AText.Count - 1 do
    inherited Add(TXLSFormattedTextItem.Create(AText[i].Text,AText[i].Font));
end;

procedure TXLSFormattedText.Add(const AText: AxUCString; AFirstFont: TXc12Font; AFontRuns: PXc12FontRunArray; AFontRunsCount: integer);
var
  i,j,k: integer;
  S: AxUCString;
  S2: AxUCString;
  F: TXc12Font;
  FoundCR: boolean;
  Item: TXLSFormattedTextItem;
begin
  FFirstFont := AFirstFont;

  if AText = '' then
    Exit;
  if AFontRunsCount <= 0 then begin
    Item := TXLSFormattedTextItem.Create(AText,AFirstFont);
    Add(Item);
  end
  else begin
    j := 1;
    if (AFontRuns[0].Index > 0) and (AFirstFont <> Nil) then begin
      k := 0;
      F := AFirstFont;
    end
    else begin
      k := 1;
      F := AFontRuns[0].Font;
    end;

    if FSplitAtCR then begin
      for i := k to AFontRunsCount - 1 do begin
        S := Copy(AText,j,AFontRuns[i].Index - j + 1);
        while S <> '' do begin
          S2 := SplitAtCRLF(S,FoundCR);
          // This shall never happend.
          if F = Nil then
            Item := TXLSFormattedTextItem.Create(S2,AFirstFont)
          else
            Item := TXLSFormattedTextItem.Create(S2,F);
          Item.FNewLine := FoundCR;
          Add(Item);
          F := AFontRuns[i].Font;
          j := AFontRuns[i].Index + 1;
        end;
      end;
    end
    else begin
      for i := k to AFontRunsCount - 1 do begin
        S := Copy(AText,j,AFontRuns[i].Index - j + 1);
        // This shall never happend.
        if F = Nil then
          Item := TXLSFormattedTextItem.Create(S,AFirstFont)
        else
          Item := TXLSFormattedTextItem.Create(S,F);
        Add(Item);
        F := AFontRuns[i].Font;
        j := AFontRuns[i].Index + 1;
      end;
    end;

    if FSplitAtCR then begin
      S := Copy(AText,j,MAXINT);
      while S <> '' do begin
        S2 := SplitAtCRLF(S,FoundCR);
        Item := TXLSFormattedTextItem.Create(S2,F);
        Item.FNewLine := FoundCR;
        Add(Item);
      end;
    end
    else begin
      S := Copy(AText,j,MAXINT);
      Item := TXLSFormattedTextItem.Create(S,F);
      Add(Item);
    end;
  end;
end;

procedure TXLSFormattedText.Add(const AText: AxUCString; AFont: TXc12Font);
var
  Item: TXLSFormattedTextItem;
  S: AxUCString;
  S2: AxUCString;
  FoundCR: boolean;
begin
  if AText = '' then
    Exit;
  if FSplitAtCR then begin
    S := AText;
    while S <> '' do begin
      S2 := SplitAtCRLF(S,FoundCR);
      Item := TXLSFormattedTextItem.Create(S2,AFont);
      Item.FNewLine := FoundCR;
      Add(Item);
    end;
  end
  else begin
    Item := TXLSFormattedTextItem.Create(AText,AFont);
    Add(Item);
  end;
end;

constructor TXLSFormattedText.Create;
begin
  inherited Create;
end;

function TXLSFormattedText.GetItems(Index: integer): TXLSFormattedTextItem;
begin
  Result := TXLSFormattedTextItem(inherited Items[Index]);
end;

{ TXLSFormattedTextItem }

procedure TXLSFormattedTextItem.AddSpecialCR(ACR: AxUCChar);
begin
  FText := FText + ACR;
end;

function TXLSFormattedTextItem.AsHTML: AxUCString;
begin
  Result := Format('<span class="%s">%s</span>',[FFont.CSSSelector,XLSUTF8Encode(FText)]);
end;

constructor TXLSFormattedTextItem.Create(AText: AxUCString; AFont: TXc12Font);
begin
  FText := AText;
  FFont := AFont;
end;

function TXLSFormattedTextItem.FontAsRTF: AxUCString;
begin
  Result := Format('\fs%d',[Round(Font.Size * 2)]);
  if xfsBold in Font.Style then
    Result := Result + '\b1'
  else
    Result := Result + '\b0';
  if xfsItalic in Font.Style then
    Result := Result + '\i1'
  else
    Result := Result + '\i0' ;
  if Font.Underline <> xulNone then
    Result := Result + '\ul1'
  else
    Result := Result + '\ul0';
end;

end.
