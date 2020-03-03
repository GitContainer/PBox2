unit XLSWriteRTF;

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

uses Classes, SysUtils, Contnrs, StrUtils, IniFiles, Math,
     XLSUtils5,
     XLSAXWEditor;

type TAXWWriteRTF = class(TObject)
protected
     FDoc         : TAXWLogDocEditor;
     FStream      : TStream;
     FFontTable   : THashedStringList;
     FColorTable  : TList;
     FStyles      : THashedStringList;
     FHasNumbering: boolean;

     procedure WChar(AChar: AnsiChar); {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     procedure WStr(AString: AxUCString); overload; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     procedure WStr(AString: AnsiString); overload; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     procedure WStrUC(AString: AxUCString); {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     procedure WStrCR(AString: AxUCString); {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     procedure WCR; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     procedure WProp(AName: AxUCString); overload; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     procedure WProp(AName: AxUCString; AValue: integer); overload; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     procedure WProp(AName,AValue: AxUCString); overload; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     procedure WBeginGroup; overload; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     procedure WBeginGroup(AName: AxUCString); overload; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     procedure WBeginGroup(AName,AValue: AxUCString); overload; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     procedure WBeginGroup(AName: AxUCString; AValue: integer); overload; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     // Pt -> Character units
     procedure WPropPt(AName: AxUCString; AValue: double); {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     // Pt -> Twips
     procedure WPropTwips(AName: AxUCString; AValue: double); {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     procedure WPropTwipsNZ(AName: AxUCString; AValue: double); {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     procedure WStream(AStream: TStream);
     procedure WShpProp(AName,AValue: AxUCString); overload;
     procedure WShpProp(AName:AxUCString; AValue: integer); overload;
     procedure WShpPropNZ(AName:AxUCString; AValue: integer);
     procedure WBeginShpProp(AName,AValue: AxUCString);
     procedure WBeginField(AName,AValue: AxUCString);
     procedure WEndGroup(ACount: integer = 1);

     function  ColorIndex(AColor: longword): integer; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     procedure AddColor(AColor: longword);

     procedure CollectBorderData(ABorder: TAXWDocPropBorder);
     procedure CollectTableData(ATable: TAXWTable);
     procedure CollectStylesData;
     procedure CollectPAPXData(APAPX: TAXWPAPX);
     procedure CollectCHPXData(ACHPX: TAXWCHPX);
     procedure CollectParasData(AParas: TAXWLogParas);
     procedure CollectNumberingsData;

     procedure DoMainDocument;
     procedure DoHeader;
     procedure DoFontTable;
     procedure DoColorTable;
     procedure DoStyles;
     procedure DoDocument(ADoc: TAXWLogDocEditor);
     procedure DoHeadersFooters;
     procedure DoNumberings;
     procedure DoParas(AParas: TAXWLogParas; AInTable: boolean = False);
     procedure DoPara(APara: TAXWLogPara);
     procedure DoCharRun(APara: TAXWLogPara; var ACR: TAXWCharRun);
     procedure DoTable(ATable: TAXWTable);
//     procedure DoSimpleField(APara: TAXWLogPara; var ACR: TAXWCharRunSimpleField);
     procedure DoBorder(AName: AxUCString; ABorder: TAXWDocPropBorder);
     procedure DoCellBorder(ATAPX: TAXWTAPX; ACellProps: TAXWTableCellProps; ABorder: TAXWTapId);
     procedure DoCHPX(APAPX: TAXWPAPX; ACHPX: TAXWCHPX);
     procedure DoPAPX(APara: TAXWLogPara);
     procedure DoSEP(ASEP: TAXWSEP);
//     procedure DoGraphic(AGraphic: TAXWGraphicRotate);
//     procedure DoPicture(AGrPicture: TAXWGraphicPicture); overload;
//     procedure DoPicture(ACRPict: TAXWCharRunPicture); overload;
public
     constructor Create(ADoc: TAXWLogDocEditor);
     destructor Destroy; override;

     procedure SaveToStream(AStream: TStream);
     procedure SaveToFile(const AFilename: AxUCString);
     end;

function EMU(AValue: double): integer; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}

implementation

function EMU(AValue: double): integer;
begin
  Result := Round(AValue * EMU_PER_PT);
end;

{ TAXWWriteRTF }

procedure TAXWWriteRTF.AddColor(AColor: longword);
begin
  if (AColor <> 0) and (FColorTable.IndexOf(Pointer(AColor)) < 0) then
    FColorTable.Add(Pointer(AColor));
end;

procedure TAXWWriteRTF.CollectBorderData(ABorder: TAXWDocPropBorder);
begin
  if ABorder <> Nil then
    AddColor(ABorder.Color);
end;

procedure TAXWWriteRTF.CollectCHPXData(ACHPX: TAXWCHPX);
begin
  if FFontTable.IndexOf(ACHPX.FontName) < 0 then
    FFontTable.Add(ACHPX.FontName);
  AddColor(ACHPX.Color);
  AddColor(ACHPX.UnderlineColor);
end;

procedure TAXWWriteRTF.CollectNumberingsData;
//var
//  i,j : integer;
//  Numb: TAXWNumbering;
begin
//  for i := 0 to FDoc.Numberings.Count - 1 do begin
//    for j := 0 to FDoc.Numberings[i].Count - 1 do begin
//      Numb := FDoc.Numberings[i][j];
//
//      CollectPAPXData(Numb.PAPX);
//      CollectCHPXData(Numb.CHPX);
//    end;
//  end;
end;

procedure TAXWWriteRTF.CollectParasData(AParas: TAXWLogParas);
var
  i,j : integer;
  Para: TAXWLogPara;
begin
  for i := 0 to AParas.Count - 1 do begin
    case AParas.Items[i].Type_ of
      alptPara: begin
        Para := AParas[i];
        if Para.PAPX <> Nil then
          CollectPAPXData(Para.PAPX);
//        FHasNumbering := FHasNumbering or (Para.Numbering <> Nil);
        for j := 0 to Para.Runs.Count - 1 do begin
          if Para.Runs[j].CHPX <> Nil then
            CollectCHPXData(Para.Runs[j].CHPX);
        end;
      end;
      alptTable: CollectTableData(AParas.Tables[i]);
    end;
  end;
end;

procedure TAXWWriteRTF.CollectPAPXData(APAPX: TAXWPAPX);
begin
  if (APAPX.Color <> 0) and (FColorTable.IndexOf(Pointer(APAPX.Color)) < 0) then
    FColorTable.Add(Pointer(APAPX.Color));

  CollectBorderData(APAPX.BorderLeft);
  CollectBorderData(APAPX.BorderTop);
  CollectBorderData(APAPX.BorderRight);
  CollectBorderData(APAPX.BorderBottom);
end;

procedure TAXWWriteRTF.CollectStylesData;
var
  i    : integer;
  Style: TAXWStyle;
begin
  FDoc.Styles.GetAllStyles(FStyles);
  for i := 0 to FStyles.Count - 1 do begin
    Style := TAXWStyle(FStyles.Objects[i]);
    case Style.Type_ of
      astPara     : begin
        CollectPAPXData(TAXWParaStyle(Style).PAPX);
//        CollectCHPXData(TAXWParaStyle(Style).CHPX);
      end;
      astCharRun  : CollectCHPXData(TAXWCharStyle(Style).CHPX);
//      astTable    : CollectPAPXData(TAXWTableStyle(Style).PAPX);
      astNumbering: ;
    end;
  end;
end;

procedure TAXWWriteRTF.CollectTableData(ATable: TAXWTable);
var
  R,C : integer;
  Row : TAXWTableRow;
  Cell: TAXWTableCell;
begin
  CollectBorderData(ATable.TAPX.BorderLeft);
  CollectBorderData(ATable.TAPX.BorderTop);
  CollectBorderData(ATable.TAPX.BorderRight);
  CollectBorderData(ATable.TAPX.BorderBottom);
  CollectBorderData(ATable.TAPX.BorderInsideVert);
  CollectBorderData(ATable.TAPX.BorderInsideHoriz);

  for R := 0 to ATable.Count - 1 do begin
    Row := ATable[R];
    for C := 0 to Row.Count - 1 do begin
      Cell := Row[C];

      if Cell.Props <> Nil then begin
        CollectBorderData(Cell.Props.Borders.Left);
        CollectBorderData(Cell.Props.Borders.Top);
        CollectBorderData(Cell.Props.Borders.Right);
        CollectBorderData(Cell.Props.Borders.Bottom);
        CollectBorderData(Cell.Props.BorderInsideHoriz);
        CollectBorderData(Cell.Props.BorderInsideVert);
        CollectBorderData(Cell.Props.BorderTl2br);
        CollectBorderData(Cell.Props.BorderTr2bl);
      end;

      CollectParasData(Cell.Paras);
    end;
  end;
end;

function TAXWWriteRTF.ColorIndex(AColor: longword): integer;
begin
  Result := FColorTable.IndexOf(Pointer(AColor)) + 1;
end;

constructor TAXWWriteRTF.Create(ADoc: TAXWLogDocEditor);
begin
  FDoc := ADoc;

  FFontTable := THashedStringList.Create;
  FColorTable := TList.Create;
  FStyles := THashedStringList.Create;
end;

destructor TAXWWriteRTF.Destroy;
begin
  FFontTable.Free;
  FColorTable.Free;
  FStyles.Free;

  inherited;
end;

procedure TAXWWriteRTF.DoBorder(AName: AxUCString; ABorder: TAXWDocPropBorder);
begin
  WProp(AName);

  case ABorder.Style of
    stbNil: ;
    stbNone: ;
    stbSingle                : WProp('brdrs');
    stbThick                 : WProp('brdrth');
    stbDouble                : WProp('brdrdb');
    stbDotted                : WProp('brdrdot');
    stbDashed                : WProp('brdrdash');
    stbDotDash               : WProp('brdrdashd');
    stbDotDotDash            : WProp('brdrdashdd');
    stbTriple                : WProp('brdrtriple');
    stbThinThickSmallGap     : WProp('brdrthtnsg');
    stbThickThinSmallGap     : WProp('brdrtnthsg');
    stbThinThickThinSmallGap : WProp('brdrtnthtnsg');
    stbThinThickMediumGap    : WProp('brdrthtnmg');
    stbThickThinMediumGap    : WProp('brdrtnthmg');
    stbThinThickThinMediumGap: WProp('brdrtnthtnmg');
    stbThinThickLargeGap     : WProp('brdrthtnlg');
    stbThickThinLargeGap     : WProp('brdrtnthlg');
    stbThinThickThinLargeGap : WProp('brdrtnthtnlg');
    stbWave                  : WProp('brdrwavy');
    stbDoubleWave            : WProp('brdrwavydb');
    stbDashSmallGap          : WProp('brdrdashsm');
    stbOutset                : WProp('brdroutset');
    stbInset                 : WProp('brdrinset');
    else                       WProp('brdrs');
  end;
  WProp('brdrw',Round(ABorder.Width * 20));
  WProp('brsp',Round(ABorder.Space * 20));
  WProp('brdrcf',ColorIndex(ABorder.Color));
end;

procedure TAXWWriteRTF.DoCellBorder(ATAPX: TAXWTAPX; ACellProps: TAXWTableCellProps; ABorder: TAXWTapId);
var
  Brd: TAXWDocPropBorder;
begin
  Brd := Nil;
  case ABorder of
    axtiBorderLeft       : begin
      WProp('clbrdrl');
      if (ACellProps <> Nil) and (ACellProps.Borders.Left <> Nil) then
        Brd := ACellProps.Borders.Left
      else if ATAPX.BorderLeft <> Nil then
        Brd := ATAPX.BorderLeft;
    end;
    axtiBorderTop        : begin
      WProp('clbrdrt');
      if (ACellProps <> Nil) and (ACellProps.Borders.Top <> Nil) then
        Brd := ACellProps.Borders.Top
      else if ATAPX.BorderTop <> Nil then
        Brd := ATAPX.BorderTop;
    end;
    axtiBorderRight      : begin
      WProp('clbrdrr');
      if (ACellProps <> Nil) and (ACellProps.Borders.Right <> Nil) then
        Brd := ACellProps.Borders.Right
      else if ATAPX.BorderRight <> Nil then
        Brd := ATAPX.BorderRight;
    end;
    axtiBorderBottom     : begin
      WProp('clbrdrb');
      if (ACellProps <> Nil) and (ACellProps.Borders.Bottom <> Nil) then
        Brd := ACellProps.Borders.Bottom
      else if ATAPX.BorderBottom <> Nil then
        Brd := ATAPX.BorderBottom;
    end;
    axtiBorderInsideVert : begin
      WProp('clbrdrr');
      if (ACellProps <> Nil) and (ACellProps.BorderInsideVert <> Nil) then
        Brd := ACellProps.BorderInsideVert
      else if ATAPX.BorderInsideVert <> Nil then
        Brd := ATAPX.BorderInsideVert;
    end;
    axtiBorderInsideHoriz: begin
      WProp('clbrdrb');
      if (ACellProps <> Nil) and (ACellProps.BorderInsideHoriz <> Nil) then
        Brd := ACellProps.BorderInsideHoriz
      else if ATAPX.BorderInsideHoriz <> Nil then
        Brd := ATAPX.BorderInsideHoriz;
    end;
  end;
  if Brd = Nil then
    WProp('brdrs')
  else if Brd.Style <> stbNone then
    DoBorder('brdrs',Brd);
end;

procedure TAXWWriteRTF.DoCharRun(APara: TAXWLogPara; var ACR: TAXWCharRun);
// var
//  Para  : TAXWLogPara;
//  FNote : TAXWFootnoteItem;
//  CRLink: TAXWCharRunLinked;
begin
  if ACR.Type_ in AXWCharRunTypeText then
    WProp('plain');

  DoCHPX(APara.PAPX,ACR.CHPX);

  case ACR.Type_ of
    acrtText,
    acrtHyperlink  : begin
      WStrUC(' ');
      WStrUC(ACR.RawText);
    end;
    acrtTab        : WProp('tab');
    acrtBreak      : begin
      case TAXWCharRunBreak(ACR).BreakType of
        acrbtLineBreak            : WProp('line');
        acrbtPageBreak            : WProp('page');
        acrbtColumn               : WProp('column');
        acrbtSeparator            : ;
        acrbtContinuationSeparator: ;
      end;
    end;
//    acrtPicture    : DoPicture(TAXWCharRunPicture(ACR));
//    acrtGraphic    : begin
//      if TAXWCharRunGraphic(ACR).Graphic is TAXWGraphicRotate then
//        DoGraphic(TAXWGraphicRotate(TAXWCharRunGraphic(ACR).Graphic));
//    end;
//    acrtSimpleField: DoSimpleField(APara,TAXWCharRunSimpleField(ACR));
//    acrtLinked     : begin
//      CRLink := TAXWCharRunLinked(ACR);
//      case CRLink.LinkedType of
//        acrltBookmark: begin
//          if CRLink.LinkedPrev = Nil then begin
//            WBeginGroup('*\bkmkstart ',CRLink.AsBookmark.Name);
//            WEndGroup;
//          end;
//          if CRLink.LinkedNext = Nil then begin
//            WBeginGroup('*\bkmkend ',CRLink.GetFirst.AsBookmark.Name);
//            WEndGroup;
//          end;
//        end;
//      end;
//    end;
//    acrtFootnote   : ;
//    acrtFootnoteRef: begin
//      FNote := TAXWFootnoteItem(TAXWCharRunFootnote(ACR).Footnote);
//
//      WProp('plain');
//      WProp('super');
//      WProp('chftn');
//
//      if TAXWCharRunFootnote(ACR).IsEndnote then
//        WBeginGroup('endnote')
//      else
//        WBeginGroup('footnote');
//      WStr('{}');
//
//      WProp('super');
//
//      WProp('chftn');
//
//      Para := FNote.StartPara;
//      while (Para <> Nil) and (Para.Index <= FNote.EndPara.Index) do begin
//        DoPara(Para);
//        Para := Para.Next;
//      end;
//
//      WEndGroup;
//    end;
  end;
end;

procedure TAXWWriteRTF.DoCHPX(APAPX: TAXWPAPX; ACHPX: TAXWCHPX);
var
  FontData: TAXWFontDataDoc;
begin
  FontData := TAXWFontDataDoc.Create;
  try
    if (ACHPX <> Nil) or ((APAPX <> Nil) and (APAPX.Style <> Nil)) then begin
      FontData.Apply(FDoc.MasterCHP);

      if (APAPX <> Nil) and (APAPX.Style <> Nil) then
        FontData.Apply(APAPX.Style);

//      if (APAPX <> Nil) and (APAPX.CHPX <> Nil) then
//        FontData.Apply(APAPX.CHPX);

      if ACHPX <> Nil then
        FontData.Apply(ACHPX);

      if FontData.Bold then
        WProp('b');
      if FontData.Italic then
        WProp('i');
      if FontData.StrikeOut then
        WProp('strike');
      WProp('cf',ColorIndex(FontData.Color));
      case FontData.Underline of
        axcuNone           : ;
        axcuSingle         : WProp('ul');
        axcuWords          : WProp('ulw');
        axcuDouble         : WProp('uldb');
        axcuThick          : WProp('ulth');
        axcuDotted         : WProp('uld');
        axcuDottedHeavy    : WProp('ulthd');
        axcuDash           : WProp('uldash');
        axcuDashedHeavy    : WProp('ulthdash');
        axcuDashLong       : WProp('ulldash');
        axcuDashLongHeavy  : WProp('ulthldash');
        axcuDotDash        : WProp('uldashd');
        axcuDashDotHeavy   : WProp('ulthdashd');
        axcuDotDotDash     : WProp('uldashdd');
        axcuDashDotDotHeavy: WProp('ulthdashdd');
        axcuWave           : WProp('ulwave');
        axcuWavyHeavy      : WProp('ulhwave');
        axcuWavyDouble     : WProp('ululdbwave');
        else                 WProp('ul');
      end;
      WProp('ulc',ColorIndex(FontData.ULColor));
      case FontData.SubSuperscript of
        axcssSubscript  : WProp('sub');
        axcssSuperscript: WProp('super');
      end;
      WProp('fs',Round(FontData.Size) * 2);
      if (ACHPX <> Nil) and (ACHPX.Style <> Nil) then
        WProp('cs',FStyles.IndexOf(ACHPX.Style.Name));
      WProp('f',FFontTable.IndexOf(FontData.Name));
    end
    else begin
      WProp('f',FFontTable.IndexOf(FDoc.MasterCHP.FontName));
      WProp('fs',Round(FDoc.MasterCHP.Size * 2));
    end;
  finally
    FontData.Free;
  end;
end;

procedure TAXWWriteRTF.DoColorTable;
var
  i: integer;
  C: longword;
begin
  if FColorTable.Count > 0 then begin
    WStr('{\colortbl;');
    for i := 0 to FColorTable.Count - 1 do begin
      C := Longword(FColorTable[i]);
      WProp('red',(C and $00FF0000) shr 16);
      WProp('green',(C and $0000FF00) shr 8);
      WProp('blue',C and $000000FF);
      WStrCR(';');
    end;
    WStrCR('}');
  end;
end;

procedure TAXWWriteRTF.DoDocument(ADoc: TAXWLogDocEditor);
begin
  WBeginGroup('generator ',ComponentName + '_' + CurrentVersionNumber + ';');
  WEndGroup;
  DoParas(ADoc.Paras);
end;

procedure TAXWWriteRTF.DoFontTable;
var
  i: integer;
begin
  WStrCR('{\fonttbl');
  for i := 0 to FFontTable.Count - 1 do
    WStrCR(Format('{\f%d\fnil\fcharset0 %s;}',[i,FFontTable[i]]));
  WStrCR('}');
end;

//procedure TAXWWriteRTF.DoGraphic(AGraphic: TAXWGraphicRotate);
//var
//  Shp  : TAXWGraphicShape;
//begin
//  WStr('{');
//  if AGraphic.Wrap <> agowInLine then begin
//    WProp('shp');
//    WStr('{');
//    WProp('*\shpinst');
//    WPropTwips('shpleft',AGraphic.PosLeft);
//    WPropTwips('shptop',AGraphic.PosTop);
//    WPropTwips('shpright',AGraphic.PosLeft + AGraphic.Width);
//    WPropTwips('shpbottom',AGraphic.PosTop + AGraphic.Height);
//
//    case AGraphic.Type_ of
//      agtPicture  : WShpProp('shapeType','75');
//      agtRectangle: ;
//      agtRoundRect: ;
//      agtEllipse  : ;
//      agtShape    : if TAXWGraphicShape(AGraphic).TextBox <> Nil then
//        WShpProp('shapeType','202');
//    end;
//
//    case AGraphic.HorizPosRel of
//      agohprMargin       : WProp('shpbxmargin');
//      agohprPage         : WProp('shpbxpage');
//      agohprColumn       : WProp('shpbxcolumn');
////      agohprCharacter    : ;
////      agohprLeftMargin   : ;
////      agohprRightMargin  : ;
////      agohprInsideMargin : ;
////      agohprOutsideMargin: ;
//      else                 WProp('shpbxcolumn');
//    end;
//    if AGraphic.Wrap in [agow_ParaLeft,agow_ParaRight] then
//      WProp('shpbxignore');
//
//    case AGraphic.VertPosRel of
//      agovprMargin       : WProp('shpbymargin');
//      agovprPage         : WProp('shpbypage');
//      agovprParagraph    : WProp('shpbypara');
////      agovprLine         : ;
////      agovprTopMargin    : ;
////      agovprBottomMargin : ;
////      agovprInsideMargin : ;
////      agovprOutsideMargin: ;
//      else                 WProp('shpbypara');
//    end;
//    if AGraphic.Wrap in [agow_ParaLeft,agow_ParaRight] then
//      WProp('shpbyignore');
//
//    if AGraphic.Angle <> 0 then
//      WShpProp('rotation',Round(RadToDeg(-AGraphic.Angle * 60000)));
//  end;
//
//  case AGraphic.Type_ of
//    agtPicture  : begin
//      if AGraphic.Wrap <> agowInLine then
//        WBeginShpProp('pib','{\shpwr2\shpwrk3');
//      DoPicture(TAXWGraphicPicture(AGraphic));
//      if AGraphic.Wrap <> agowInLine then
//        WEndGroup(2);
//    end;
//    agtRectangle: ;
//    agtRoundRect: ;
//    agtEllipse  : ;
//    agtShape    : begin
//      Shp := TAXWGraphicShape(AGraphic);
//      WShpProp('fillColor',RevRGB(Shp.FillColor));
//      WShpPropNZ('lineColor',RevRGB(Shp.StrokeColor));
//      WShpPropNZ('lineWidth',EMU(Shp.StrokeWeight));
//      if TAXWGraphicShape(AGraphic).TextBox <> Nil then begin
//        WBeginGroup('shptxt');
//        DoParas(TAXWLogDocEditor(Shp.TextBox.Owner).Paras);
//        WEndGroup;
//      end;
//    end;
//  end;
//
//  if AGraphic.Wrap <> agowInLine then
//    WEndGroup;
//  WEndGroup;
//end;

procedure TAXWWriteRTF.DoHeader;
begin
  WStrCR('{\rtf1\fbidis\ansi\ansicpg0\uc1\deff0\fet2');

  DoSEP(FDOC.SEP);
end;

procedure TAXWWriteRTF.DoHeadersFooters;
begin
//  if FDoc.Headers.First <> Nil then begin
//    WBeginGroup('headerf ');
//    WProp('titlepg ');
//    DoParas(FDoc.Headers.First.Paras);
//    WEndGroup;
//  end;
//
//  if FDoc.Headers.Default_ <> Nil then begin
//    WBeginGroup('header ');
//    DoParas(FDoc.Headers.Default_.Paras);
//    WEndGroup;
//  end;
//
//  if FDoc.Headers.Even <> Nil then begin
//    WBeginGroup('headerl ');
//    WProp('facingp ');
//    DoParas(FDoc.Headers.Even.Paras);
//    WEndGroup;
//  end;
//
//  if FDoc.Footers.First <> Nil then begin
//    WBeginGroup('footerf ');
//    WProp('titlepg ');
//    DoParas(FDoc.Footers.First.Paras);
//    WEndGroup;
//  end;
//
//  if FDoc.Footers.Default_ <> Nil then begin
//    WBeginGroup('footer ');
//    DoParas(FDoc.Footers.Default_.Paras);
//    WEndGroup;
//  end;
//
//  if FDoc.Footers.Even <> Nil then begin
//    WBeginGroup('footerl ');
//    WProp('facingp ');
//    DoParas(FDoc.Footers.Even.Paras);
//    WEndGroup;
//  end;
end;

procedure TAXWWriteRTF.DoMainDocument;
begin
//  FDoc.Bookmarks.Reindex;

  FFontTable.Add(FDoc.MasterCHP.FontName);
  CollectParasData(FDoc.Paras);

//  if FDoc.Headers.Default_ <> Nil then CollectParasData(FDoc.Headers.Default_.Paras);
//  if FDoc.Headers.First    <> Nil then CollectParasData(FDoc.Headers.First.Paras);
//  if FDoc.Headers.Even     <> Nil then CollectParasData(FDoc.Headers.Even.Paras);
//
//  if FDoc.Footers.Default_ <> Nil then CollectParasData(FDoc.Footers.Default_.Paras);
//  if FDoc.Footers.First    <> Nil then CollectParasData(FDoc.Footers.First.Paras);
//  if FDoc.Footers.Even     <> Nil then CollectParasData(FDoc.Footers.Even.Paras);

  CollectStylesData;

  CollectNumberingsData;

  DoHeader;
  WStr('{\*\ftnsep\pard\plain\chftnsep }');
  WStr('{\*\ftnsepc\pard\plain\chftnsepc}');

  DoFontTable;
  DoColorTable;
  DoStyles;

  WStr('{\*\generator ' + ComponentName + '_' + CurrentVersionNumber + '}');

  if FHasNumbering then
    DoNumberings;

  DoHeadersFooters;

  DoDocument(FDoc);

  WEndGroup;
end;

procedure TAXWWriteRTF.DoNumberings;
//var
//  i,j : integer;
//  Numb: TAXWNumbering;
begin
//  WBeginGroup('*\listtable');
//  WCR;
//
//  for i := 0 to FDoc.Numberings.Count - 1 do begin
//    WBeginGroup('list');
//    for j := 0 to FDoc.Numberings[i].Count - 1 do begin
//      Numb := FDoc.Numberings[i][j];
//
//      WBeginGroup('listlevel');
//
//      case Numb.Style of
//        anstNone       : WProp('levelnfc',255);
//        anstBullet     : WProp('levelnfc',23);
//        anstDecimal    : WProp('levelnfc',0);
//        anstLowerRoman : WProp('levelnfc',2);
//        anstUpperRoman : WProp('levelnfc',1);
//        anstLowerLetter: WProp('levelnfc',3);
//        anstUpperLetter: WProp('levelnfc',4);
//      end;
//
//      case Numb.Align of
//        axptaLeft  : ;
//        axptaCenter: WProp('leveljc',1);
//        axptaRight : WProp('leveljc',2);
//      end;
//
//      WPropTwips('li',Numb.PAPX.IndentLeft);
//      WPropTwips('fi',Numb.PAPX.IndentFirstLine);
//      WProp('jclisttab');
//      WPropTwips('tx',Numb.PAPX.IndentLeft);
//      WProp('levelstartat1');
//
//      if Numb.Style in [anstNone,anstBullet] then begin
//        WBeginGroup('leveltext');
//        WStr('\''01\u183 \''b7;');
//        WEndGroup;
//
//        WBeginGroup('levelnumbers');
//        WStr(';');
//        WEndGroup;
//      end
//      else begin
//        WBeginGroup('leveltext');
//        WStr('\''02\''0' + IntToStr(j) + ';');
//        WEndGroup;
//
//        WBeginGroup('levelnumbers');
//        WStr('\''01;');
//        WEndGroup;
//      end;
//
//      DoCHPX(Numb.PAPX,Numb.CHPX);
//
//      WEndGroup;
//      WCR;
//    end;
//    WProp('listid',1000 + i);
//    WCR;
//    WEndGroup;
//    WCR;
//  end;
//  WEndGroup;
//  WCR;
//
//  WBeginGroup('*\listoverridetable');
//  for i := 0 to FDoc.Numberings.Count - 1 do begin
//    WBeginGroup('listoverride');
//    WProp('listid',1000 + i);
//    WProp('listoverridecount',0);
//    WProp('ls',i + 1);
//    WEndGroup;
//    WCR;
//  end;
//  WEndGroup;
//  WCR;
end;

procedure TAXWWriteRTF.DoPAPX(APara: TAXWLogPara);
var
  BL,BT,BR,BB: TAXWDocPropBorder;
  ParaData   : TAXWParaData;
begin
  ParaData := TAXWParaData.Create;
  try
    BL := Nil;
    BT := Nil;
    BR := Nil;
    BB := Nil;

    APara.SetupParaData(ParaData);

    case ParaData.Alignment of
      axptaLeft   : WProp('ql');
      axptaCenter : WProp('qc');
      axptaRight  : WProp('qr');
      axptaJustify: WProp('qj');
    end;
    WProp('cbpat',ColorIndex(ParaData.Color));
    if ParaData.IndentLeft <> 0 then
      WPropPt('li',ParaData.IndentLeft);
    if ParaData.IndentRight      <> 0 then
      WPropPt('ri',ParaData.IndentRight);
    if ParaData.IndentFirstLine  <> 0 then
      WPropPt('fi',ParaData.IndentFirstLine);
    if ParaData.IndentHanging    <> 0 then
      WPropPt('fi',-ParaData.IndentHanging);
    if ParaData.LineSpacing      <> 0 then begin
      WProp('sl',Round(ParaData.LineSpacing * 240));
      WProp('slmult',1);
    end;
    if ParaData.SpaceBefore      <> 0 then
      WPropTwips('sb',ParaData.SpaceBefore);
    if ParaData.SpaceAfter       <> 0 then
      WPropTwips('sa',ParaData.SpaceAfter);

    if ParaData.Borders.Left <> Nil then
      BL := ParaData.Borders.Left;
    if ParaData.Borders.Top <> Nil then
      BT := ParaData.Borders.Top;
    if ParaData.Borders.Right <> Nil then
      BR := ParaData.Borders.Right;
    if ParaData.Borders.Bottom <> Nil then
      BB := ParaData.Borders.Bottom;

    if (BL <> Nil) or (BT <> Nil) or (BR <> Nil) or (BB <> Nil) then begin
//      WProp('brdrbtw');
      if BL <> Nil then
        DoBorder('brdrl',BL);
      if BT <> Nil then
        DoBorder('brdrt',BT);
      if BR <> Nil then
        DoBorder('brdrr',BR);
      if BB <> Nil then
        DoBorder('brdrb',BB);
    end;
  finally
    ParaData.Free;
  end;
end;

procedure TAXWWriteRTF.DoPara(APara: TAXWLogPara);
var
  i    : integer;
  CR   : TAXWCharRun;
  HLink: boolean;
begin
//  if APara.SEP <> Nil then begin
//    WProp('sectd');
//    if APara.SEP.Cols.Count > 0 then begin
//      WProp('cols',APara.SEP.Cols.Count);
//      WPropTwips('colsx',APara.SEP.Cols.AverageColSpace);
//    end;
//  end;

  WStrCR('\pard');

  // {\listtext\pard\plain\f1\fs24\lang1024 \u183 \'b7\tab}\ls1\ilvl0 \plain \f0\fs20 Sugga
//  if APara.Numbering <> Nil then begin
//    WBeginGroup('listtext');
//
//    WProp('pard');
//    WProp('plain');
//
//    WStrUC(APara.Numbering.Text);
//
//    WProp('tab');
//
//    WEndGroup;
//
//    WProp('ls',APara.Numbering.NumList.Index + 1);
//    WProp('ilvl',APara.Numbering.Level);
//  end;

//  if APara.TabStops <> Nil then begin
//    for i := 0 to APara.TabStops.Count - 1 do begin
//      case APara.TabStops[i].Alignment of
//        atsaClear  : ;
//        atsaLeft   : ;
//        atsaCenter : WProp('tqc');
//        atsaRight  : WProp('tqr');
//        atsaDecimal: WProp('tqdec');
//        atsaNum    : ;
//      end;
//
//      case APara.TabStops[i].Leader of
//        atslDot       : WProp('tldot');
//        atslHyphen    : WProp('tlhyph');
//        atslUnderscore: WProp('tlul');
//        atslHeavy     : WProp('tlth');
//        atslMiddleDot : WProp('tlmdot');
//      end;
//
//      WPropTwips('tx',APara.TabStops[i].Position);
//    end;
//  end;

  DoPAPX(APara);

  HLink := False;

  for i := 0 to APara.Runs.Count - 1 do begin
    CR := APara.Runs[i];

    if not HLink and (CR.Type_ = acrtHyperlink) then begin
      WBeginField('HYPERLINK',TAXWCharRunHyperlink(CR).Hyperlink.Address);
      HLink := True;
    end
    else if HLink and (CR.Type_ <> acrtHyperlink) then begin
      WEndGroup(2);
      HLink := False;
    end;

    DoCharRun(APara,CR);
  end;

  if HLink then
    WEndGroup(2);

//  if APara.Graphics <> Nil then begin
//    for i := 0 to APara.Graphics.Count - 1 do begin
//      if APara.Graphics[i] is TAXWGraphicRotate then
//        DoGraphic(TAXWGraphicRotate(APara.Graphics[i]));
//    end;
//  end;
end;

procedure TAXWWriteRTF.DoParas(AParas: TAXWLogParas; AInTable: boolean = False);
var
  i: integer;
begin
  for i := 0 to AParas.Count - 1 do begin
    case AParas.Items[i].Type_ of
      alptPara : begin
        if AInTable then
          WProp('intbl');

        DoPara(AParas[i]);

        if not AInTable or (AInTable and (i < (AParas.Count - 1))) then
          WProp('par');
      end;
      alptTable: DoTable(AParas.Tables[i]);
    end;
  end;
end;

//procedure TAXWWriteRTF.DoPicture(ACRPict: TAXWCharRunPicture);
//begin
//  WStr('{');
//  WProp('pict');
//  WPropTwips('picwgoal',ACRPict.Bitmap.Width);
//  WPropTwips('pichgoal',ACRPict.Bitmap.Height);
//
//  case ACRPict.OrigType of
//    aptBMP     : WProp('dibitmap0');
//    aptJPG     : WProp('jpegblip');
//    aptPNG     : WProp('pngblip');
//    aptGIF     : ; // TODO convert.
//    aptEMF     : WProp('emfblip');
//  end;
//
//  WProp('picw',ACRPict.Bitmap.PixWidth);
//  WProp('pich',ACRPict.Bitmap.PixHeight);
//
//  WStr(' ');
//  WStream(ACRPict.OrigData);
//  WStr('}');
//end;

//procedure TAXWWriteRTF.DoPicture(AGrPicture: TAXWGraphicPicture);
//begin
//// '{\pict\picwgoal1020\pichgoal675\jpegblip\picw68\pich45 '
//// \pict\picscalex100\picscaley100\piccropl0\piccropr0\piccropt0\piccropb0\picw1799\pich1191\picwgoal1020\pichgoal675\jpegblip\bliptag260238629
//
//  WStr('{');
//  WProp('pict');
//  WPropTwips('picwgoal',AGrPicture.Width);
//  WPropTwips('pichgoal',AGrPicture.Height);
//
//  case AGrPicture.Picture.OrigType of
//    aptBMP     : WProp('dibitmap0');
//    aptJPG     : WProp('jpegblip');
//    aptPNG     : WProp('pngblip');
//    aptGIF     : ; // TODO convert.
//    aptEMF     : WProp('emfblip');
//  end;
//
//  WProp('picw',AGrPicture.Picture.Width);
//  WProp('pich',AGrPicture.Picture.Height);
//
//  WPropTwipsNZ('piccropl',AGrPicture.CropLeft);
//  WPropTwipsNZ('piccropt',AGrPicture.CropTop);
//  WPropTwipsNZ('piccropr',AGrPicture.CropRight);
//  WPropTwipsNZ('piccropb',AGrPicture.CropBottom);
//
//  WStr(' ');
//  WStream(AGrPicture.Picture.OrigData);
//  WStr('}');
//end;

procedure TAXWWriteRTF.DoSEP(ASEP: TAXWSEP);
begin
//  WPropTwips('paperw',ASEP.PageWidth);
//  WPropTwips('paperh',ASEP.PageHeight);
//  WPropTwips('margl',ASEP.MargLeft);
//  WPropTwips('margt',ASEP.MargTop);
//  WPropTwips('margr',ASEP.MargRight);
//  WPropTwips('margb',ASEP.MargBottom);
end;

//procedure TAXWWriteRTF.DoSimpleField(APara: TAXWLogPara; var ACR: TAXWCharRunSimpleField);
//var
//  TempCR: TAXWCharRun;
//begin
//  TempCR := TAXWCharRun.Create(ACR.Parent,Nil);
//  try
//    WBeginGroup('field');
//    WBeginGroup('*\fldinst ');
//    WStrUC(' ' + ACR.SimpField.InstrText);
//    WEndGroup;
//    WBeginGroup('fldrslt ');
//    while ACR <> Nil do begin
//      WBeginGroup;
//
//      TempCR.Text := ACR.RawText;
//      TempCR.CHPX := ACR.CHPX;
//      DoCharRun(APara,TempCR);
//
//      WEndGroup;
//
//      if (ACR.Next <> Nil) and (ACR.Next.Type_ = acrtSimpleField) then
//        ACR := TAXWCharRunSimpleField(ACR.Next)
//      else
//        Break;
//    end;
//    WEndGroup(2);
//  finally
//    TempCR.Free;
//  end;
//end;

procedure TAXWWriteRTF.DoStyles;
var
  i    : integer;
  Style: TAXWStyle;
begin
  FDoc.Styles.GetAllStyles(FStyles);
  if FStyles.Count > 0 then begin
    WBeginGroup('stylesheet');
    WCR;
    for i := 0 to FStyles.Count - 1 do begin
      Style := TAXWStyle(FStyles.Objects[i]);
      case Style.Type_ of
        astPara     : begin
        end;
        astCharRun  : begin
          WBeginGroup('*\cs',i);
          WStr(' ');
//          WStr('\additive ');
          DoCHPX(Nil,TAXWCharStyle(STyle).CHPX);
          WStr(' ' + Style.Name + ';');
          WEndGroup;
          WCR;
        end;
        astTable    : ;
        astNumbering: ;
      end;
    end;
    WEndGroup;
  end;
end;

procedure TAXWWriteRTF.DoTable(ATable: TAXWTable);
var
  R,C : integer;
  X   : double;
  Row : TAXWTableRow;
  Cell: TAXWTableCell;
begin
  for R := 0 to ATable.Count - 1 do begin
    Row := ATable[R];

    WBeginGroup('parad');
    WProp('intbl');
    WProp('trowd');

    WProp('trftsWidth3');
//    WPropTwips('trwWidth',Row.Width);

    if Row.AHeight > 0 then
      WPropTwipsNZ('trrh',Row.AHeight);

    X := 0;
    for C := 0 to Row.Count - 1 do begin
      Cell := Row[C];

      X := X + Cell.Width;

      if R = 0 then begin
        if C = 0 then
          DoCellBorder(ATable.TAPX,Cell.Props,axtiBorderLeft);
        if C = (Row.Count - 1) then
          DoCellBorder(ATable.TAPX,Cell.Props,axtiBorderRight);
        if C < (Row.Count - 1)  then
          DoCellBorder(ATable.TAPX,Cell.Props,axtiBorderInsideVert);
        DoCellBorder(ATable.TAPX,Cell.Props,axtiBorderTop);
      end;
      if R = (ATable.Count - 1) then begin
        if C = 0 then
          DoCellBorder(ATable.TAPX,Cell.Props,axtiBorderLeft);
        if C = (Row.Count - 1) then
          DoCellBorder(ATable.TAPX,Cell.Props,axtiBorderRight);
        if C < (Row.Count - 1) then
          DoCellBorder(ATable.TAPX,Cell.Props,axtiBorderInsideVert);
        DoCellBorder(ATable.TAPX,Cell.Props,axtiBorderBottom);
      end;
      if R < (ATable.Count - 1) then begin
        if C = 0 then
          DoCellBorder(ATable.TAPX,Cell.Props,axtiBorderLeft);
        if C = (Row.Count - 1) then
          DoCellBorder(ATable.TAPX,Cell.Props,axtiBorderRight);
        if C < (Row.Count - 1) then
          DoCellBorder(ATable.TAPX,Cell.Props,axtiBorderInsideVert);
        DoCellBorder(ATable.TAPX,Cell.Props,axtiBorderInsideHoriz);
      end;

      if atcfMergedRoot in Cell.Flags then
        WProp('clvmgf')
      else if (atcfMerged in Cell.Flags) or (atcfMergedLast in Cell.Flags) then
        WProp('clvmrg');

      WProp('clftsWidth3');
      WPropTwips('clwWidth',Cell.Width);
      WPropTwips('cellx',Round(X));

//      DoParas(Cell.Paras,True);
//
//      WProp('cell');
    end;

    for C := 0 to Row.Count - 1 do begin
      Cell := Row[C];
      DoParas(Cell.Paras,True);

      WProp('cell');
    end;

    WProp('row');
    WCR;
    WEndGroup;
  end;
end;

procedure TAXWWriteRTF.SaveToFile(const AFilename: AxUCString);
var
  Stream: TFileStream;
begin
  Stream := TFileStream.Create(AFilename,fmCreate);
  try
    SaveToStream(Stream);
  finally
    Stream.Free;
  end;
end;

procedure TAXWWriteRTF.SaveToStream(AStream: TStream);
begin
  FStream := AStream;

  DoMainDocument;
end;

procedure TAXWWriteRTF.WBeginField(AName, AValue: AxUCString);
begin
// {\field{\*\fldinst HYPERLINK "www.sugga.com"}{\fldrslt \plain \cs8\f0\ul\fs20\cf1 Sugga}}!
  WBeginGroup('field');
  WBeginGroup('*\fldinst ');
  WStr(AName + ' ');
  WStr(AValue);
  WEndGroup;
  WBeginGroup('fldrslt ');
end;

procedure TAXWWriteRTF.WBeginGroup;
begin
  WStr('{');
end;

procedure TAXWWriteRTF.WBeginGroup(AName, AValue: AxUCString);
begin
  WStr('{\' + AName + AValue);
end;

procedure TAXWWriteRTF.WBeginGroup(AName: AxUCString; AValue: integer);
begin
  WStr('{\' + AName + IntToStr(AValue));
end;

procedure TAXWWriteRTF.WBeginShpProp(AName, AValue: AxUCString);
begin
  WStr('{\sp {\sn ' + AName + '}{\sv ' + AValue);
end;

procedure TAXWWriteRTF.WChar(AChar: AnsiChar);
begin
  FStream.Write(AChar,1);
end;

procedure TAXWWriteRTF.WCR;
begin
{$ifdef DEBUG}
  WStr(#13#10);
{$endif}
end;

procedure TAXWWriteRTF.WEndGroup(ACount: integer);
begin
  while ACount > 0 do begin
    WStr('}');
    Dec(ACount);
  end;
end;

procedure TAXWWriteRTF.WProp(AName: AxUCString);
begin
  WStr('\' + AName);
end;

procedure TAXWWriteRTF.WProp(AName: AxUCString; AValue: integer);
begin
  WProp(AName);
  WStr(IntToStr(AValue));
end;

procedure TAXWWriteRTF.WBeginGroup(AName: AxUCString);
begin
  WStr('{\' + AName);
end;

procedure TAXWWriteRTF.WProp(AName, AValue: AxUCString);
begin
  WProp(AName);
  WStr(AValue);
end;

procedure TAXWWriteRTF.WPropPt(AName: AxUCString; AValue: double);
begin
  WProp(AName,Round(AValue * 18));
end;

procedure TAXWWriteRTF.WPropTwips(AName: AxUCString; AValue: double);
begin
  WProp(AName,Round(AValue * 20));
end;

procedure TAXWWriteRTF.WPropTwipsNZ(AName: AxUCString; AValue: double);
begin
  if AValue <> 0 then
    WProp(AName,Round(AValue * 20));
end;

procedure TAXWWriteRTF.WStr(AString: AxUCString);
var
  S: AnsiString;
begin
  S := AnsiString(AString);

  FStream.Write(Pointer(S)^,Length(S));
end;

procedure TAXWWriteRTF.WShpProp(AName, AValue: AxUCString);
begin
  WStr('{\sp {\sn ' + AName + '}{\sv ' + AValue + '}}');
end;

procedure TAXWWriteRTF.WShpProp(AName: AxUCString; AValue: integer);
begin
  WShpProp(AName,IntToStr(AValue));
end;

procedure TAXWWriteRTF.WShpPropNZ(AName: AxUCString; AValue: integer);
begin
  if AValue <> 0 then
    WShpProp(AName,AValue);
end;

procedure TAXWWriteRTF.WStr(AString: AnsiString);
begin
  FStream.Write(Pointer(AString)^,Length(AString));
end;

procedure TAXWWriteRTF.WStrCR(AString: AxUCString);
begin
  WStr(AString);
  WCR;
end;

procedure TAXWWriteRTF.WStream(AStream: TStream);
var
  i: integer;
  b: byte;
begin
  AStream.Seek(0,soFromBeginning);

  for i := 0 to AStream.Size - 1 do begin
    AStream.Read(b,1);

    WStr(HexStrs255[b]);
  end;
end;

procedure TAXWWriteRTF.WStrUC(AString: AxUCString);
var
  i: integer;
  C: AxUCChar;
begin
  for i := 1 to Length(AString) do begin
    C := AString[i];
    if C = '\' then
      WStr('\\')
    else if Word(C) > $7F then begin
      WProp('u',Word(C));

      // Dummy. If zero is written, the (u)nicode value is used.
      WStr(' \''00');
      // If the character code is written, the correct codepage must be written
      // as well (Prop: langXXXX).
//      WChar(' ');
//      WProp('''',Format('%x',[Word(C)]));
    end
    else
      WChar(AnsiChar(C));
  end;
end;

end.
