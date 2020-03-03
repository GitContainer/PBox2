unit XLSHTMLTypes5;

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

uses Classes, IniFiles;

type THTMLElementID = (heTEXT,
heDOCUMENT,heA,heABBR,heACRONYM,heADDRESS,heAPPLET,heAREA,heB,heBASE,heBASEFONT,
heBDO,heBIG,heBLOCKQUOTE,heBODY,heBR,heBUTTON,heCAPTION,heCENTER,heCITE,heCODE,
heCOL,heCOLGROUP,heDD,heDEL,heDFN,heDIR,heDIV,heDL,heDT,heEM,heFIELDSET,heFONT,
heFORM,heFRAME,heFRAMESET,heH1,heH2,heH3,heH4,heH5,heH6,heHEAD,heHR,heHTML,heI,
heIFRAME,heILAYER,heIMG,heINPUT,heINS,heISINDEX,heKBD,heLABEL,heLAYER,heLEGEND,
heLI,heLINK,heMAP,heMENU,heMETA,heNOFRAMES,heNOLAYER,heNOSCRIPT,heOBJECT,heOL,
heOPTGROUP,heOPTION,heP,hePARAM,hePRE,heQ,heS,heSAMP,heSCRIPT,heSELECT,heSMALL,
heSPAN,heSTRIKE,heSTRONG,heSTYLE,heSUB,heSUP,heTABLE,heTBODY,heTD,heTEXTAREA,
heTFOOT,heTH,heTHEAD,heTITLE,heTR,heTT,heU,heUL,heVAR,heDOCTYPE,

heUNKNOWN,heEOF,

heBeginEndtags,

heEndTEXT,
heEndDOCUMENT,heEndA,heEndABBR,heEndACRONYM,heEndADDRESS,heEndAPPLET,heEndAREA,
heEndB,heEndBASE,heEndBASEFONT,heEndBDO,heEndBIG,heEndBLOCKQUOTE,heEndBODY,
heEndBR,heEndBUTTON,heEndCAPTION,heEndCENTER,heEndCITE,heEndCODE,heEndCOL,heEndCOLGROUP,
heEndDD,heEndDEL,heEndDFN,heEndDIR,heEndDIV,heEndDL,heEndDT,heEndEM,heEndFIELDSET,
heEndFONT,heEndFORM,heEndFRAME,heEndFRAMESET,heEndH1,heEndH2,heEndH3,heEndH4,
heEndH5,heEndH6,heEndHEAD,heEndHR,heEndHTML,heEndI,heEndIFRAME,heEndILAYER,heEndIMG,
heEndINPUT,heEndINS,heEndISINDEX,heEndKBD,heEndLABEL,heEndLAYER,heEndLEGEND,
heEndLI,heEndLINK,heEndMAP,heEndMENU,heEndMETA,heEndNOFRAMES,heEndNOLAYER,
heEndNOSCRIPT,heEndOBJECT,heEndOL,heEndOPTGROUP,heEndOPTION,heEndP,heEndPARAM,
heEndPRE,heEndQ,heEndS,heEndSAMP,heEndSCRIPT,heEndSELECT,heEndSMALL,heEndSPAN,
heEndSTRIKE,heEndSTRONG,heEndSTYLE,heEndSUB,heEndSUP,heEndTABLE,heEndTBODY,heEndTD,
heEndTEXTAREA,heEndTFOOT,heEndTH,heEndTHEAD,heEndTITLE,heEndTR,heEndTT,heEndU,
heEndUL,heEndVAR,heEndDOCTYPE,

heEndUNKNOWN,heEndEOF);

type TElementAttributeID = (
eaAbbr,eaAcceptCharset,eaAccept,eaAccesskey,eaAction,eaAlign,eaAlt,eaArchive,
eaAxis,eaBackground,eaBgcolor,eaBorder,eaCellpadding,eaChar,eaCharoff,eaCharset,
eaChecked,eaCite,eaClass,eaClassid,eaClear,eaCode,eaCodebase,eaCodetype,eaColor,
eaCols,eaColspan,eaCompact,eaContent,eaCoords,eaData,eaDatetime,eaDeclare,eaDefer,
eaDir,eaDisabled,eaEnctype,eaFace,eaFor,eaFrame,eaFrameborder,eaHeaders,eaHeight,
eaHref,eaHreflang,eaHspace,eaHttpEquiv,eaId,eaIsmap,eaLabel,eaLang,eaLanguage,
eaLink,eaLeftmargin,eaLongdesc,eaMarginheight,eaMarginwidth,eaMaxlength,eaMedia,
eaMethod,eaMultiple,eaName,eaNoHref,eaNoResize,eaNoShade,eaNoWrap,eaObject,eaOnBlur,
eaOnChange,eaOnClick,eaOnDblclick,eaOnFocus,eaOnKeydown,eaOnKeypress,eaOnKeyup,
eaOnLoad,eaOnMousedown,eaOnMousemove,eaOnMouseout,eaOnMouseover,eaOnMouseup,
eaOnReset,eaOnSelect,eaOnSubmit,eaOnUnload,eaProfile,eaPrompt,eaAdonl,eaRel,
eaRev,eaRows,eaRowspan,eaRules,eaScheme,eaScope,eaScrolling,eaSelected,eaShape,
eaSize,eaSpan,eaSrc,eaStandby,eaStart,eaStyle,eaSummary,eaTabindex,eaTarget,eaText,
eaTitle,eaTopmargin,
eaType,eaUsemap,eaValign,eaValue,eaValuetype,eaVersion,eaVlink,eaVspace,
eaWidth,

eaUnknown);

type TAttrAlign = (aaLeft,aaCenter,aaRight,aaJustify,aaChar,aaTop,aaMiddle,aaBottom);

type TElementFlag = (efEndtag);
     TElementFlags = set of TElementFlag;

type TAttributeType = (atString,atNumber,atPercent,atColor,atAlign);

type PElementAttribute = ^TElementAttribute;
     TElementAttribute = record
     ID: TElementAttributeID;
     case AttType: TAttributeType of
       atString: (StrVal: PChar);
       atNumber: (NumVal: integer);
       atPercent:(PercentVal: integer);
       atColor:  (ColorVal: longword);
       atAlign:  (AlignVal: TAttrAlign);
     end;

type THTMLElement = record
     ID: THTMLElementID;
//     Flags: TElementFlags;
     Attributes: array of TElementAttribute;
     end;

type THTMLElementArray = array of THTMLElement;

type PHTMLElementArray = ^THTMLElementArray;

const ElementNames: array[0..98] of string = ('text',
'document','a','abbr','acronym','address','applet','area',
'b','base','basefont','bdo','big','blockquote','body','br','button','caption',
'center','cite','code','col','colgroup','dd','del','dfn','dir','div','dl',
'dt','em','fieldset','font','form','frame','frameset','h1','h2','h3','h4',
'h5','h6','head','hr','html','i','iframe','ilayer','img','input','ins','isindex','kbd',
'label','layer','legend','li','link','map','menu','meta','noframes','nolayer','noscript',
'object','ol','optgroup','option','p','param','pre','q','s','samp','script',
'select','small','span','strike','strong','style','sub','sup','table',
'tbody','td','textarea','tfoot','th','thead','title','tr','tt','u','ul','var',
'doctype',

'unknown','eof');


const AttributeNames: array[0..119] of string = (
'abbr','acceptcharset','accept','accesskey','action','align','alt','archive',
'axis','background','bgcolor','border','cellpadding','char','charoff','charset',
'checked','cite','class','classid','clear','code','codebase','codetype','color',
'cols','colspan','compact','content','coords','data','datetime','declare','defer',
'dir','disabled','enctype','face','for','frame','frameborder','headers','height',
'href','hreflang','hspace','http-equiv','id','ismap','label','lang','language',
'link','leftmargin','longdesc','marginheight','marginwidth','maxlength','media',
'method','multiple','name','nohref','noresize','noshade','nowrap','object','onblur',
'onchange','onclick','ondblclick','onfocus','onkeydown','onkeypress','onkeyup',
'onload','onmousedown','onmousemove','onmouseout','onmouseover','onmouseup',
'onreset','onselect','onsubmit','onunload','profile','prompt','adonl','rel','rev',
'rows','rowspan','rules','scheme','scope','scrolling','selected','shape','size',
'span','src','standby','start','style','summary','tabindex','target','text','title',
'topmargin','type','usemap','valign','value','valuetype','version','vlink','vspace',
'width',

'unknown');

type THTMLAlign = (haLeft,haCenter,haRight);

var
{$ifdef DELPHI_5}
  HTMLColors: TStringList;
  HTMLElements: TStringList;
  HTMLAttributes: TStringList;
{$else}
  HTMLColors: THashedStringList;
  HTMLElements: THashedStringList;
  HTMLAttributes: THashedStringList;
{$endif}

implementation

var
  i: integer;

initialization
{$ifdef DELPHI_5}
  HTMLColors := TStringList.Create;
  HTMLElements := TStringList.Create;
  HTMLAttributes := TStringList.Create;
{$else}
  HTMLColors := THashedStringList.Create;
  HTMLElements := THashedStringList.Create;
  HTMLAttributes := THashedStringList.Create;
{$endif}

  for i := 0 to High(ElementNames) do
    HTMLElements.Add(ElementNames[i]);

  for i := 0 to High(AttributeNames) do
    HTMLAttributes.Add(AttributeNames[i]);

  HTMLColors.AddObject('aliceblue',TObject($FFF8F0));
  HTMLColors.AddObject('antiquewhite',TObject($D7EBFA));
  HTMLColors.AddObject('aqua',TObject($FFFF00));
  HTMLColors.AddObject('aquamarine',TObject($D4FF7F));
  HTMLColors.AddObject('azure',TObject($FFFFF0));
  HTMLColors.AddObject('beige',TObject($DCF5F5));
  HTMLColors.AddObject('bisque',TObject($C4E4FF));
  HTMLColors.AddObject('black',TObject($000000));
  HTMLColors.AddObject('blanchedalmond',TObject($CDEBFF));
  HTMLColors.AddObject('blue',TObject($FF0000));
  HTMLColors.AddObject('blueviolet',TObject($E22B8A));
  HTMLColors.AddObject('brown',TObject($2A2AA5));
  HTMLColors.AddObject('burlywood',TObject($87B8DE));
  HTMLColors.AddObject('cadetblue',TObject($A09E5F));
  HTMLColors.AddObject('chartreuse',TObject($00FF7F));
  HTMLColors.AddObject('chocolate',TObject($1E69D2));
  HTMLColors.AddObject('coral',TObject($507FFF));
  HTMLColors.AddObject('cornflowerblue',TObject($ED9564));
  HTMLColors.AddObject('cornsilk',TObject($DCF8FF));
  HTMLColors.AddObject('crimson',TObject($3C14DC));
  HTMLColors.AddObject('cyan',TObject($FFFF00));
  HTMLColors.AddObject('darkblue',TObject($8B0000));
  HTMLColors.AddObject('darkcyan',TObject($8B8B00));
  HTMLColors.AddObject('darkgoldenrod',TObject($0B86B8));
  HTMLColors.AddObject('darkgray',TObject($A9A9A9));
  HTMLColors.AddObject('darkgrey',TObject($A9A9A9));
  HTMLColors.AddObject('darkgreen',TObject($006400));
  HTMLColors.AddObject('darkkhaki',TObject($6BB7BD));
  HTMLColors.AddObject('darkmagenta',TObject($8B008B));
  HTMLColors.AddObject('darkolivegreen',TObject($2F6B55));
  HTMLColors.AddObject('darkorange',TObject($008CFF));
  HTMLColors.AddObject('darkorchid',TObject($CC3299));
  HTMLColors.AddObject('darkred',TObject($00008B));
  HTMLColors.AddObject('darksalmon',TObject($7A96E9));
  HTMLColors.AddObject('darkseagreen',TObject($8FBC8F));
  HTMLColors.AddObject('darkslateblue',TObject($8B3D48));
  HTMLColors.AddObject('darkslategray',TObject($4F4F2F));
  HTMLColors.AddObject('darkslategrey',TObject($4F4F2F));
  HTMLColors.AddObject('darkturquoise',TObject($D1CE00));
  HTMLColors.AddObject('darkviolet',TObject($D30094));
  HTMLColors.AddObject('deeppink',TObject($9314FF));
  HTMLColors.AddObject('deepskyblue',TObject($FFBF00));
  HTMLColors.AddObject('dimgray',TObject($696969));
  HTMLColors.AddObject('dimgrey',TObject($696969));
  HTMLColors.AddObject('dodgerblue',TObject($FF901E));
  HTMLColors.AddObject('firebrick',TObject($2222B2));
  HTMLColors.AddObject('floralwhite',TObject($F0FAFF));
  HTMLColors.AddObject('forestgreen',TObject($228B22));
  HTMLColors.AddObject('fuchsia',TObject($FF00FF));
  HTMLColors.AddObject('gainsboro',TObject($DCDCDC));
  HTMLColors.AddObject('ghostwhite',TObject($FFF8F8));
  HTMLColors.AddObject('gold',TObject($00D7FF));
  HTMLColors.AddObject('goldenrod',TObject($20A5DA));
  HTMLColors.AddObject('gray',TObject($808080));
  HTMLColors.AddObject('grey',TObject($808080));
  HTMLColors.AddObject('green',TObject($008000));
  HTMLColors.AddObject('greenyellow',TObject($2FFFAD));
  HTMLColors.AddObject('honeydew',TObject($F0FFF0));
  HTMLColors.AddObject('hotpink',TObject($B469FF));
  HTMLColors.AddObject('indianred',TObject($5C5CCD));
  HTMLColors.AddObject('indigo',TObject($82004B));
  HTMLColors.AddObject('ivory',TObject($F0FFFF));
  HTMLColors.AddObject('khaki',TObject($8CE6F0));
  HTMLColors.AddObject('lavender',TObject($FAE6E6));
  HTMLColors.AddObject('lavenderblush',TObject($F5F0FF));
  HTMLColors.AddObject('lawngreen',TObject($00FC7C));
  HTMLColors.AddObject('lemonchiffon',TObject($CDFAFF));
  HTMLColors.AddObject('lightblue',TObject($E6D8AD));
  HTMLColors.AddObject('lightcoral',TObject($8080F0));
  HTMLColors.AddObject('lightcyan',TObject($FFFFE0));
  HTMLColors.AddObject('lightgoldenrodyellow',TObject($D2FAFA));
  HTMLColors.AddObject('lightgray',TObject($D3D3D3));
  HTMLColors.AddObject('lightgrey',TObject($D3D3D3));
  HTMLColors.AddObject('lightgreen',TObject($90EE90));
  HTMLColors.AddObject('lightpink',TObject($C1B6FF));
  HTMLColors.AddObject('lightsalmon',TObject($7AA0FF));
  HTMLColors.AddObject('lightseagreen',TObject($AAB220));
  HTMLColors.AddObject('lightskyblue',TObject($FACE87));
  HTMLColors.AddObject('lightslategray',TObject($998877));
  HTMLColors.AddObject('lightslategrey',TObject($998877));
  HTMLColors.AddObject('lightsteelblue',TObject($DEC4B0));
  HTMLColors.AddObject('lightyellow',TObject($E0FFFF));
  HTMLColors.AddObject('lime',TObject($00FF00));
  HTMLColors.AddObject('limegreen',TObject($32CD32));
  HTMLColors.AddObject('linen',TObject($E6F0FA));
  HTMLColors.AddObject('magenta',TObject($FF00FF));
  HTMLColors.AddObject('maroon',TObject($000080));
  HTMLColors.AddObject('mediumaquamarine',TObject($AACD66));
  HTMLColors.AddObject('mediumblue',TObject($CD0000));
  HTMLColors.AddObject('mediumorchid',TObject($D355BA));
  HTMLColors.AddObject('mediumpurple',TObject($D87093));
  HTMLColors.AddObject('mediumseagreen',TObject($71B33C));
  HTMLColors.AddObject('mediumslateblue',TObject($EE687B));
  HTMLColors.AddObject('mediumspringgreen',TObject($9AFA00));
  HTMLColors.AddObject('mediumturquoise',TObject($CCD148));
  HTMLColors.AddObject('mediumvioletred',TObject($8515C7));
  HTMLColors.AddObject('midnightblue',TObject($701919));
  HTMLColors.AddObject('mintcream',TObject($FAFFF5));
  HTMLColors.AddObject('mistyrose',TObject($E1E4FF));
  HTMLColors.AddObject('moccasin',TObject($B5E4FF));
  HTMLColors.AddObject('navajowhite',TObject($ADDEFF));
  HTMLColors.AddObject('navy',TObject($800000));
  HTMLColors.AddObject('oldlace',TObject($E6F5FD));
  HTMLColors.AddObject('olive',TObject($008080));
  HTMLColors.AddObject('olivedrab',TObject($238E6B));
  HTMLColors.AddObject('orange',TObject($00A5FF));
  HTMLColors.AddObject('orangered',TObject($0045FF));
  HTMLColors.AddObject('orchid',TObject($D670DA));
  HTMLColors.AddObject('palegoldenrod',TObject($AAE8EE));
  HTMLColors.AddObject('palegreen',TObject($98FB98));
  HTMLColors.AddObject('paleturquoise',TObject($EEEEAF));
  HTMLColors.AddObject('palevioletred',TObject($9370D8));
  HTMLColors.AddObject('papayawhip',TObject($D5EFFF));
  HTMLColors.AddObject('peachpuff',TObject($B9DAFF));
  HTMLColors.AddObject('peru',TObject($3F85CD));
  HTMLColors.AddObject('pink',TObject($CBC0FF));
  HTMLColors.AddObject('plum',TObject($DDA0DD));
  HTMLColors.AddObject('powderblue',TObject($E6E0B0));
  HTMLColors.AddObject('purple',TObject($800080));
  HTMLColors.AddObject('red',TObject($0000FF));
  HTMLColors.AddObject('rosybrown',TObject($8F8FBC));
  HTMLColors.AddObject('royalblue',TObject($E16941));
  HTMLColors.AddObject('saddlebrown',TObject($13458B));
  HTMLColors.AddObject('salmon',TObject($7280FA));
  HTMLColors.AddObject('sandybrown',TObject($60A4F4));
  HTMLColors.AddObject('seagreen',TObject($578B2E));
  HTMLColors.AddObject('seashell',TObject($EEF5FF));
  HTMLColors.AddObject('sienna',TObject($2D52A0));
  HTMLColors.AddObject('silver',TObject($C0C0C0));
  HTMLColors.AddObject('skyblue',TObject($EBCE87));
  HTMLColors.AddObject('slateblue',TObject($CD5A6A));
  HTMLColors.AddObject('slategray',TObject($908070));
  HTMLColors.AddObject('slategrey',TObject($908070));
  HTMLColors.AddObject('snow',TObject($FAFAFF));
  HTMLColors.AddObject('springgreen',TObject($7FFF00));
  HTMLColors.AddObject('steelblue',TObject($B48246));
  HTMLColors.AddObject('tan',TObject($8CB4D2));
  HTMLColors.AddObject('teal',TObject($808000));
  HTMLColors.AddObject('thistle',TObject($D8BFD8));
  HTMLColors.AddObject('tomato',TObject($4763FF));
  HTMLColors.AddObject('turquoise',TObject($D0E040));
  HTMLColors.AddObject('violet',TObject($EE82EE));
  HTMLColors.AddObject('wheat',TObject($B3DEF5));
  HTMLColors.AddObject('white',TObject($FFFFFF));
  HTMLColors.AddObject('whitesmoke',TObject($F5F5F5));
  HTMLColors.AddObject('yellow',TObject($00FFFF));
  HTMLColors.AddObject('yellowgreen',TObject($32CD9A));
finalization
  HTMLColors.Free;
  HTMLElements.Free;
  HTMLAttributes.Free;

end.

