unit BIFF_RecsII5;

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

//* ~exclude
const MAXCOL_97 = 255;
//* ~exclude
const MAXROW_97 = 65535;

const MAXSHEET = 255;
//* The highest column value that can be used.
const MAXCOL = MAXCOL_97;
//* The highest row value that can be used.
const MAXROW = MAXROW_97;


const
  GUID_HLINK_STDLINK: array[0..15] of byte = ($D0,$C9,$EA,$79,$F9,$BA,$CE,$11,$8C,$82,$00,$AA,$00,$4B,$A9,$0B);
  GUID_HLINK_URL:     array[0..15] of byte = ($E0,$C9,$EA,$79,$F9,$BA,$CE,$11,$8C,$82,$00,$AA,$00,$4B,$A9,$0B);
  GUID_HLINK_FILE:    array[0..15] of byte = ($03,$03,$00,$00,$00,$00,$00,$00,$C0,$00,$00,$00,$00,$00,$00,$46);


//* Options for calculations of formulas.
type TCalculateOption = (coNotCalcDependent, //* Do not calculate dependent cells.
                                             //* That is, if the formula is 'SUM(B4)' and cell B4 is a formula cell,
                                             //* that cell (B4) will not be calculated in order to obtain a correct result.
                         coExternalSeek      //* Not used anymore. Replaced with ~[link XLSNames4.TExternalNames.ExternalLookup TExternalNames.ExternalLookup]
                         );
     TCalculateOptions = set of TCalculateOption;

//* Excel file version. Each change in the file version is also a change in
//* the file format. That is, different file versions has incompatible file
//* formats. Please note that the file version is NOT the same as the Excel
//* program version. The XLS file format has remaind unchaged since Excel 97 in
//* order to ensure compatibility with future versions.
//* To start with Excel 2007 a new XML based file format was intoduced. This
//* file format stores the file components in a ZIP archive with the extension
//* XLSX. There is also a binary version of the XLSX files, with the extension
//* XLSB. The file components are still in a ZIP arcive, but are here in a
//* binary format, similar to the old versions.
//* There has never been an Excel 1.x version. Excel 2.x was choosen in order
//* to match the Word version.
// type TExcelVersion = (xvNone,xvExcelUnknown,xvExcel21,xvExcel30,xvExcel40,xvExcel50,xvExcel97,xvExcel2007);

//* Excel fill patterns. The numbers in the image corresponds to the
//* pattern constants. ~[image PatternsNum.bmp]
//type TExcelFillPattern = (efpNone,efpSolid,efpMediumGray,efpDarkGray,efpLightGray,
//                         efpDarkHorizontal,efpDarkVertical,efpDarkDown,efpDarkUp,
//                         efpDarkGrid,efpDarkTrellis,efpLightHorizontal,
//                         efpLightVertical,efpLightDown,efpLightUp,efpLightGrid,
//                         efpLightTrellis,efpGray125,efpGray0625);
//~endparse~

type TFormulaVersion = (fvExcel95,fvExcel97);

// This must be a memory from excel 4.
const MAXSELROW = $3FFF;


const MAXRECSZ_40 = 2080;
const MAXRECSZ_97 = 8224;

const MAX_EXCEL_STRSZ = 32767;

//const DEFAULT_FORMAT_97 = 15;
//const DEFAULT_FORMAT_07 = 0;
const DEFAULT_FORMATS_COUNT_50 = 16;
const DEFAULT_FORMATS_COUNT_97 = 21;
const DEFAULT_FORMAT40 = 20;

const DEFAULT_FONT_97 = 0;
const DEFAULT_FONTS_COUNT_97 = 5;

// BIFF Records
const BIFFRECID_INTEGER_20 =        $0002;
const BIFFRECID_NUMBER_20 =         $0003;
const BIFFRECID_LABEL_20 =          $0004;
const BIFFRECID_STRING_20 =         $0007;

const BIFFRECID_00EF =              $00EF;

const BIFFRECID_1904 =              $0022;
const BIFFRECID_ARRAY =             $0221;
const BIFFRECID_AUTOFILTER =        $009E;
const BIFFRECID_AUTOFILTERINFO =    $009D;
const BIFFRECID_BACKUP =            $0040;
const BIFFRECID_BLANK =             $0201;
const BIFFRECID_BOF =               $0009;
const BIFFRECID_BOF4 =              $0409;
const BIFFRECID_BOF8 =              $0809;
const BIFFRECID_BOOKBOOL =          $00DA;
const BIFFRECID_BOOLERR =           $0205;
const BIFFRECID_BOTTOMMARGIN =      $0029;
const BIFFRECID_BOUNDSHEET =        $0085;
const BIFFRECID_CALCMODE =          $000D;
const BIFFRECID_CALCCOUNT =         $000C;
const BIFFRECID_CODENAME =          $01BA; // Name according to OO Calc.
const BIFFRECID_CODENAME_VBE =      $0042; // Yes, they have the same id.
const BIFFRECID_CODEPAGE =          $0042; //
const BIFFRECID_COLINFO =           $007D;
const BIFFRECID_COLWIDTH =          $0024;
const BIFFRECID_CONTINUE =          $003C;
const BIFFRECID_CONDFMT =           $01B0;
const BIFFRECID_CF =                $01B1;
const BIFFRECID_COUNTRY =           $008C;
const BIFFRECID_CRN =               $005A;
const BIFFRECID_DBCELL =            $00D7;
const BIFFRECID_DCONREF =           $0051;
const BIFFRECID_DEFAULTROWHEIGHT =  $0225;
const BIFFRECID_DEFCOLWIDTH =       $0055;
const BIFFRECID_DELTA =             $0010;
const BIFFRECID_DIMENSIONS_20 =     $0000;
const BIFFRECID_DIMENSIONS =        $0200;
const BIFFRECID_DSF =               $0161;
const BIFFRECID_DV =                $01BE;
const BIFFRECID_DVAL =              $01B2;
const BIFFRECID_EOF =               $000A;
const BIFFRECID_EXCEL9FILE =        $01C0;
const BIFFRECID_EXTERNCOUNT =       $0016;
const BIFFRECID_EXTERNNAME =        $0023;
const BIFFRECID_EXTERNSHEET =       $0017;
const BIFFRECID_FILEPASS =          $002F;
const BIFFRECID_FILESHAREING =      $005B;
const BIFFRECID_FILTERMODE =        $009B;

const BIFFRECID_EXTSST =            $00FF;
const BIFFRECID_FILESHARING =       $005B;
const BIFFRECID_FNGROUPCOUNT =      $009C;
const BIFFRECID_FONT =              $0031;
const BIFFRECID_FONT_40 =           $0231;
const BIFFRECID_FOOTER =            $0015;
const BIFFRECID_FORMAT =            $041E;
const BIFFRECID_FORMULA =           $0006;
const BIFFRECID_FORMULA_30 =        $0206;
const BIFFRECID_FORMULA_40 =        $0406;
const BIFFRECID_GRIDSET =           $0082;
const BIFFRECID_GUTS =              $0080;
const BIFFRECID_HCENTER =           $0083;
const BIFFRECID_HEADER =            $0014;
const BIFFRECID_HIDEOBJ =           $008D;
const BIFFRECID_HLINK =             $01B8;
const BIFFRECID_HLINKTOOLTIP =      $0800;
const BIFFRECID_HORIZONTALPAGEBREAKS = $001B;
const BIFFRECID_IMDATA =            $007F; 
const BIFFRECID_INDEX =             $020B;
const BIFFRECID_INTERFACEHDR =      $00E1;
const BIFFRECID_INTERFACEEND =      $00E2;
const BIFFRECID_ITERATION =         $0011;
const BIFFRECID_LABEL =             $0204;
const BIFFRECID_LABELSST =          $00FD;
const BIFFRECID_LEFTMARGIN =        $0026;
const BIFFRECID_MERGEDCELLS =       $00E5;
const BIFFRECID_MMSADDMENUDELMENU = $00C1;
const BIFFRECID_MSO_0866 =          $0866;
const BIFFRECID_MSODRAWING =        $00EC;
const BIFFRECID_MSODRAWINGGROUP =   $00EB;
const BIFFRECID_MSODRAWINGSELECTION=$00ED;
const BIFFRECID_MULRK =             $00BD;
const BIFFRECID_MULBLANK =          $00BE;
const BIFFRECID_NAME =              $0018;
const BIFFRECID_NOTE =              $001C;
const BIFFRECID_NUMBER =            $0203;
const BIFFRECID_OBJ =               $005D;
const BIFFRECID_OBJPROTECT =        $0063;
const BIFFRECID_OBPROJ =            $00D3;
const BIFFRECID_PALETTE =           $0092;
const BIFFRECID_PANE =              $0041;
const BIFFRECID_PASSWORD =          $0013;
const BIFFRECID_PRECISION =         $000E;
const BIFFRECID_PRINTGRIDLINES =    $002B;
const BIFFRECID_PRINTHEADERS =      $002A;
const BIFFRECID_PROT4REV =          $01AF;
const BIFFRECID_PROT4REVPASS =      $01BC;
const BIFFRECID_PROTECT =           $0012;
const BIFFRECID_QSI =               $01AD;
const BIFFRECID_RECALCID =          $01C1;
const BIFFRECID_REFRESHALL =        $01B7;
const BIFFRECID_REFMODE =           $000F;
const BIFFRECID_RIGHTMARGIN =       $0027;
const BIFFRECID_RK =                $007E;
const BIFFRECID_RK7 =               $027E;
const BIFFRECID_PLS =               $004D;
const BIFFRECID_ROW =               $0208;
const BIFFRECID_RSTRING =           $00D6;
const BIFFRECID_SAVERECALC =        $005F;
const BIFFRECID_SCL =               $00A0;
const BIFFRECID_SCENPROTECT =       $00DD;
const BIFFRECID_SELECTION =         $001D;
const BIFFRECID_SETUP =             $00A1;
const BIFFRECID_SHEETEXT =          $0862;
const BIFFRECID_FEATHEADR =         $0867;
const BIFFRECID_SHRFMLA =           $00BC;
const BIFFRECID_SHRFMLA_20 =        $04BC;
const BIFFRECID_SST =               $00FC;
const BIFFRECID_STRING =            $0207;
const BIFFRECID_STYLE =             $0293;
const BIFFRECID_SUPBOOK =           $01AE;
const BIFFRECID_SXIDSTM =           $00D5;
const BIFFRECID_SXVS =              $00E3;
const BIFFRECID_TABID =             $013D;
const BIFFRECID_TOPMARGIN =         $0028;
const BIFFRECID_TXO =               $01B6;
const BIFFRECID_UNCALCED =          $005E;
const BIFFRECID_USERBVIEW =         $01A9;
const BIFFRECID_USERSVIEWBEGIN =    $01AA;
const BIFFRECID_USERSVIEWEND =      $01AB;
const BIFFRECID_USESELFS =          $0160;
const BIFFRECID_VCENTER =           $0084;
const BIFFRECID_VERTICALPAGEBREAKS =$001A;
const BIFFRECID_WINDOW1 =           $003D;
const BIFFRECID_WINDOW2 =           $023E;
const BIFFRECID_WINDOWPROTECT =     $0019;
const BIFFRECID_WRITEACCESS =       $005C;
const BIFFRECID_WSBOOL =            $0081;
const BIFFRECID_XCT =               $0059;
const BIFFRECID_XF =                $00E0;
const BIFFRECID_XF_30 =             $0243;
const BIFFRECID_XF_40 =             $0443;

const XBIFFRECID_CHART =            $0851;              

const CHARTRECID_3D =               $103A;
const CHARTRECID_AI =               $1051;
const CHARTRECID_ALRUNS =           $1050;
const CHARTRECID_AREA =             $101A;
const CHARTRECID_AREAFORMAT =       $100A;
const CHARTRECID_ATTACHEDLABEL =    $100C;
const CHARTRECID_AXCEXT =           $1062;
const CHARTRECID_AXESUSED =         $1046;
const CHARTRECID_AXIS =             $101D;
const CHARTRECID_AXISLINEFORMAT =   $1021;
const CHARTRECID_AXISPARENT =       $1041;
const CHARTRECID_BAR =              $1017;
const CHARTRECID_BEGIN =            $1033;
const CHARTRECID_BOPPOP =           $1061;
const CHARTRECID_BOPPOPCUSTOM =     $1067;
const CHARTRECID_CATSERRANGE =      $1020;
const CHARTRECID_CHART =            $1002;
const CHARTRECID_CHARTFORMAT =      $1014;
const CHARTRECID_CHARTFORMATLINK =  $1022;
const CHARTRECID_CHARTLINE =        $101C;
const CHARTRECID_DAT =              $1063;
const CHARTRECID_DATAFORMAT =       $1006;
const CHARTRECID_DEFAULTTEXT =      $1024;
const CHARTRECID_DROPBAR =          $103D;
const CHARTRECID_END =              $1034;
const CHARTRECID_FBI =              $1060;
const CHARTRECID_FONTX =            $1026;
const CHARTRECID_FRAME =            $1032;
const CHARTRECID_GELFRAME =         $1066;
const CHARTRECID_GEOMETRY =         $105F;
const CHARTRECID_IFMT =             $104E;
const CHARTRECID_LEGEND =           $1015;
const CHARTRECID_LEGENDXN =         $1043;
const CHARTRECID_LINE =             $1018;
const CHARTRECID_LINEFORMAT =       $1007;
const CHARTRECID_MARKERFORMAT =     $1009;
const CHARTRECID_OBJECTLINK =       $1027;
const CHARTRECID_PICF =             $103C;
const CHARTRECID_PIE =              $1019;
const CHARTRECID_PIEFORMAT =        $100B;
const CHARTRECID_PLOTAREA =         $1035;
const CHARTRECID_PLOTGROWTH =       $1064;
const CHARTRECID_POS =              $104F;
const CHARTRECID_RADAR =            $103E;
const CHARTRECID_RADARAREA =        $1040;
const CHARTRECID_SBASEREF =         $1048;
const CHARTRECID_SCATTER =          $101B;
const CHARTRECID_SERAUXERRBAR =     $105B;
const CHARTRECID_SERAUXTREND =      $104B;
const CHARTRECID_SERFMT =           $105D;
const CHARTRECID_SERIES =           $1003;
const CHARTRECID_SERIESLIST =       $1016;
const CHARTRECID_SERIESTEXT =       $100D;
const CHARTRECID_SERPARENT =        $104A;
const CHARTRECID_SERTOCRT =         $1045;
const CHARTRECID_SHTPROPS =         $1044;
const CHARTRECID_SIINDEX =          $1065;
const CHARTRECID_SURFACE =          $103F;
const CHARTRECID_TEXT =             $1025;
const CHARTRECID_TICK =             $101E;
const CHARTRECID_UNITS =            $1001;
const CHARTRECID_VALUERANGE =       $101F;

const OBJREC_END      = $00;
const OBJREC_MACRO    = $04;
const OBJREC_BUTTON   = $05;
const OBJREC_GMO      = $06;
const OBJREC_CF	      = $07;
const OBJREC_PIOGRBIT = $08;
const OBJREC_PICTFMLA = $09;
const OBJREC_CBLS     = $0A;
const OBJREC_RBO      = $0B;
const OBJREC_SBS      = $0C;
const OBJREC_NTS      = $0D;
const OBJREC_SBSFMLA  = $0E;
const OBJREC_GBODATA  = $0F;
const OBJREC_EDODATA  = $10;
const OBJREC_RBODATA  = $11;
const OBJREC_CBLSDATA = $12;
const OBJREC_LBSDATA  = $13;
const OBJREC_CBLSFMLA = $14;
const OBJREC_CMO      = $15;


const OBJTYPE_GROUP         = $0000;
const OBJTYPE_LINE          = $0001;
const OBJTYPE_RECTANGLE     = $0002;
const OBJTYPE_OVAL          = $0003;
const OBJTYPE_ARC           = $0004;
const OBJTYPE_CHART         = $0005;
const OBJTYPE_TEXT          = $0006;
const OBJTYPE_BUTTON        = $0007;
const OBJTYPE_PICTURE       = $0008;
const OBJTYPE_POLYGON       = $0009;
const OBJTYPE_RESERVED1     = $000A;
const OBJTYPE_CHECK_BOX     = $000B;
const OBJTYPE_OPTION_BUTTON = $000C;
const OBJTYPE_EDIT_BOX      = $000D;
const OBJTYPE_LABEL         = $000E;
const OBJTYPE_DIALOG_BOX    = $000F;
const OBJTYPE_SPINNER       = $0010;
const OBJTYPE_SCROLL_BAR    = $0011;
const OBJTYPE_LIST_BOX      = $0012;
const OBJTYPE_GROUP_BOX     = $0013;
const OBJTYPE_COMBO_BOX     = $0014;
const OBJTYPE_RESERVED2     = $0015;
const OBJTYPE_RESERVED3     = $0016;
const OBJTYPE_RESERVED4     = $0017;
const OBJTYPE_RESERVED5     = $0018;
const OBJTYPE_COMMENT       = $0019;
const OBJTYPE_RESERVED6     = $001A;
const OBJTYPE_RESERVED7     = $001B;
const OBJTYPE_RESERVED8     = $001C;
const OBJTYPE_RESERVED9     = $001D;
const OBJTYPE_MICROSOFT_OFFICE_DRAWING = $001E;

//const TokenClassRef   = $20;
//const TokenClassValue = $40;
//const TokenClassArray = $60;

// NAC = No Argument Class
ptgArea_NAC        = $05;
ptgAreaN_NAC       = $0D;
ptgAreaErr_NAC     = $0B;
ptgArea3d_NAC      = $1B;
ptgAreaErr3d_NAC   = $1D;
ptgArray_NAC       = $00;
ptgFunc_NAC        = $01;
ptgFuncCEV_NAC     = $18;
ptgFuncVar_NAC     = $02;
ptgMemArea_NAC     = $06;
ptgMemAreaN_NAC    = $0E;
ptgMemErr_NAC      = $07;
ptgMemFunc_NAC     = $09;
ptgMemNoMem_NAC    = $08;
ptgMemNoMemN_NAC   = $0F;
ptgName_NAC        = $03;
ptgNameX_NAC       = $19;
ptgRef_NAC         = $04;
ptgRefN_NAC        = $0C;
ptgRef3d_NAC       = $1A;
ptgRefErr_NAC      = $0A;
ptgRefErr3d_NAC    = $1C;

//  ptgArea        = $25;
//  ptgAreaV       = $45;
//  ptgAreaA       = $65;

//  ptgAreaN       = $2D;
//  ptgAreaNV      = $4D;
//  ptgAreaNA      = $6D;

//  ptgAreaErr     = $2B;
//  ptgAreaErrV    = $4B;
//  ptgAreaErrA    = $6B;

//  ptgArea3d      = $3B;
//  ptgArea3dV     = $5B;
//  ptgArea3dA     = $7B;

//  ptgAreaErr3d   = $3D;
//  ptgAreaErr3dV  = $5D;
//  ptgAreaErr3dA  = $7D;

//  ptgArray       = $20;
//  ptgArrayV      = $40;
//  ptgArrayA      = $60;

//  ptgFunc        = $21;
//  ptgFuncV       = $41;
//  ptgFuncA       = $61;

//  ptgFuncCEV     = $58;
//  ptgFuncCEA     = $78;

//  ptgFuncVar     = $22;
//  ptgFuncVarV    = $42;
//  ptgFuncVarA    = $62;

//  ptgMemArea     = $26;
//  ptgMemAreaV    = $46;
//  ptgMemAreaA    = $66;

//  ptgMemAreaN    = $2E;
//  ptgMemAreaNV   = $4E;
//  ptgMemAreaNA   = $6E;

//  ptgMemErr      = $27;
//  ptgMemErrV     = $47;
//  ptgMemErrA     = $67;

//  ptgMemFunc     = $29;
//  ptgMemFuncV    = $49;
//  ptgMemFuncA    = $69;

//  ptgMemNoMem    = $28;
//  ptgMemNoMemV   = $48;
//  ptgMemNoMemA   = $68;

//  ptgMemNoMemN   = $2F;
//  ptgMemNoMemNV  = $4F;
//  ptgMemNoMemNA  = $6F;

//  ptgName        = $23;
//  ptgNameV       = $43;
//  ptgNameA       = $63;

//  ptgNameX       = $39;
//  ptgNameXV      = $59;
//  ptgNameXA      = $79;

//  ptgRef         = $24;
//  ptgRefV        = $44;
//  ptgRefA        = $64;

//  ptgRefN        = $2C;
//  ptgRefNV       = $4C;
//  ptgRefNA       = $6C;

//  ptgRef3d       = $3A;
//  ptgRef3dV      = $5A;
//  ptgRef3dA      = $7A;

//  ptgRefErr      = $2A;
//  ptgRefErrV     = $4A;
//  ptgRefErrA     = $6A;

//  ptgRefErr3d    = $3C;
//  ptgRefErr3dV   = $5C;
//  ptgRefErr3dA   = $7C;


//const
//  ptgExp         = $01;
//  ptgTbl         = $02;
//  ptgAdd         = $03;
//  ptgSub         = $04;
//  ptgMul         = $05;
//  ptgDiv         = $06;
//  ptgPower       = $07;
//  ptgConcat      = $08;
//  ptgLT	         = $09;
//  ptgLE	         = $0A;
//  ptgEQ	         = $0B;
//  ptgGE	         = $0C;
//  ptgGT	         = $0D;
//  ptgNE	         = $0E;
//  ptgIsect       = $0F;
//  ptgUnion       = $10;
//  ptgRange       = $11;
//  ptgUplus       = $12;
//  ptgUminus      = $13;
//  ptgPercent     = $14;
//  ptgParen       = $15;
//  ptgMissArg     = $16;
//  ptgStr         = $17;
//  ptgExtend      = $18;
//  ptgAttr        = $19;
//  ptgSheet       = $1A;
//  ptgEndSheet    = $1B;
//  ptgErr         = $1C;
//  ptgBool        = $1D;
//  ptgInt         = $1E;
//  ptgNum         = $1F;
//  ptgArray       = $20;
//  ptgFunc        = $21;
//  ptgFuncVar     = $22;
//  ptgName        = $23;
//  ptgRef         = $24;
//  ptgArea        = $25;
//  ptgMemArea     = $26;
//  ptgMemErr      = $27;
//  ptgMemNoMem    = $28;
//  ptgMemFunc     = $29;
//  ptgRefErr      = $2A;
//  ptgAreaErr     = $2B;
//  ptgRefN        = $2C;
//  ptgAreaN       = $2D;
//  ptgMemAreaN    = $2E;
//  ptgMemNoMemN   = $2F;
//  ptgNameX       = $39;
//  ptgRef3d       = $3A;
//  ptgArea3d      = $3B;
//  ptgRefErr3d    = $3C;
//  ptgAreaErr3d   = $3D;
//  ptgArrayV      = $40;
//  ptgFuncV       = $41;
//  ptgFuncVarV    = $42;
//  ptgNameV       = $43;
//  ptgRefV        = $44;
//  ptgAreaV       = $45;
//  ptgMemAreaV    = $46;
//  ptgMemErrV     = $47;
//  ptgMemNoMemV   = $48;
//  ptgMemFuncV    = $49;
//  ptgRefErrV     = $4A;
//  ptgAreaErrV    = $4B;
//  ptgRefNV       = $4C;
//  ptgAreaNV      = $4D;
//  ptgMemAreaNV   = $4E;
//  ptgMemNoMemNV  = $4F;
//  ptgFuncCEV     = $58;
//  ptgNameXV      = $59;
//  ptgRef3dV      = $5A;
//  ptgArea3dV     = $5B;
//  ptgRefErr3dV   = $5C;
//  ptgAreaErr3dV  = $5D;
//  ptgArrayA      = $60;
//  ptgFuncA       = $61;
//  ptgFuncVarA    = $62;
//  ptgNameA       = $63;
//  ptgRefA        = $64;
//  ptgAreaA       = $65;
//  ptgMemAreaA    = $66;
//  ptgMemErrA     = $67;
//  ptgMemNoMemA   = $68;
//  ptgMemFuncA    = $69;
//  ptgRefErrA     = $6A;
//  ptgAreaErrA    = $6B;
//  ptgRefNA       = $6C;
//  ptgAreaNA      = $6D;
//  ptgMemAreaNA   = $6E;
//  ptgMemNoMemNA  = $6F;
//  ptgFuncCEA     = $78;
//  ptgNameXA      = $79;
//  ptgRef3dA      = $7A;
//  ptgArea3dA     = $7B;
//  ptgRefErr3dA   = $7C;
//  ptgAreaErr3dA  = $7D;
//// internal used by the tokenizer
//  ptgNone        = $00;
//  ptgParenR      = $F0;
//  ptgSpace       = $F1;
//  ptg3d          = $F2;
//  ptg3dX         = $F3;

                            // Size
  eptgElfLel	      = $01;  // 4
  eptgElfRw	        = $02;  // 4
  eptgElfCol	      = $03;  // 4
  eptgElfRwV	      = $06;  // 4
  eptgElfColV	      = $07;  // 4
  eptgElfRadical    = $0A;  // 13
  eptgElfRadicalS   = $0B;  // 13
  eptgElfRwS	      = $0C;  // 4
  eptgElfColS	      = $0D;  // 4
  eptgElfRwSV	      = $0E;  // 4
  eptgElfColSV	    = $0F;  // 4
  eptgElfRadicalLel = $10;  // 4
  eptgSxName	      = $1D;  // 4

type PRecPtgsExp = ^TRecPtgsExp;
     TRecPtgsExp = packed record
     Id: byte;
     Row: word;
     Col: word;
     end;

const ERR_NULL  = 1;
const ERR_DIV   = 2;
const ERR_VALUE = 3;
const ERR_REF   = 4;
const ERR_NAME  = 5;
const ERR_NUM   = 6;
const ERR_NA    = 7;


const FONT_COLOR_SYS_WINTEXT = $7FFF;

type TSubStreamType = (stGlobals,stVBModule,stWorksheet,stChart,stExcel4Macro,stWorkspace);

type TWordBit = (b0,b1,b2,b3,b4,b5,b6,b7,b8,b9,b10,b11,b12,b13,b14,b15);
type TWordBits = set of TWordBit;

type T2ByteWord = packed record
     case word of
       0: (W: word);
       1: (B: array[0..1] of byte);
     end;

type TRecCellAreaI = record
     Row1,Row2: integer;
     Col1,Col2: integer;
     end;

type PRecCellArea = ^TRecCellArea;
     TRecCellArea = packed record
     Row1,Row2: word;
     Col1,Col2: word;
     end;

type PRec32CellArea = ^TRecCellArea;
     TRec32CellArea = packed record
     Row1,Row2: longword;
     Col1,Col2: longword;
     end;

type PRecCellAreaShort = ^TRecCellAreaShort;
     TRecCellAreaShort = packed record
     Row1,Row2: word;
     Col1,Col2: byte;
     end;

type PRecCellRef = ^TRecCellRef;
     TRecCellRef = packed record
     Row: word;
     Col: word;
     end;

type TCellRefArray = array of TRecCellRef;

type TPageBreak = packed record
     Val1: word;
     Val2: word;
     Val3: word;
     end;

type PRecTXORUN = ^TRecTXORUN;
     TRecTXORUN = packed record
     CharIndex: word;
     FontIndex: word;
     Reserved: longword;
     end;

type TDynTXORUNArray = array of TRecTXORUN;

type PObjCMO = ^TObjCMO;
     TObjCMO = packed record
     ObjType: word;
     ObjId: word;
     Options: word;
     Reserved: array[0..11] of byte;
     end;

type PObjSBS = ^TObjSBS;
     TObjSBS = packed record
     Reserved: array[0..3] of byte;
     ScrollPos: word;
     MinimumValue: word;
     MaximumValue: word;
     ScrollWhenArrowClicked: word;
     ScrollWhenBarClicked: word;
     Horizontal: word;
     Width: word;
     Options: word;
     end;

// This is the part AFTER the formula

type PObjLBS = ^TObjLBS;
     TObjLBS = packed record
//   Reserved1: word;
//   FmlaLen: word;
//   Reserved2: longword;
//   Fmla: array of byte;

     Reserved3: byte;
     Elements: word;
     Selected: word;
     Options: word;
     end;



// ********** BIFF Records *******

type PFormatRun = ^TFormatRun;
     TFormatRun = packed record
     Index: word;
     FontIndex: word;
     end;

type PBIFFHeader = ^TBIFFHeader;
     TBIFFHeader = packed record
     RecID: word;
     Length: word;
     end;

type PRec2INTEGER = ^TRec2INTEGER;
     TRec2INTEGER = packed record
     Row: word;
     Col: byte;
     Res: byte;
     Attribute: array[0..2] of byte;
     Value: word;
     end;

type PRec2NUMBER = ^TRec2NUMBER;
     TRec2NUMBER = packed record
     Row: word;
     Col: byte;
     Res: byte;
     Attribute: array[0..2] of byte;
     Value: double;
     end;

type PRec2LABEL = ^TRec2LABEL;
     TRec2LABEL = packed record
     Row: word;
     Col: byte;
     Res: byte;
     Attribute: array[0..2] of byte;
     Len: byte;
     Data: array[0..255] of byte;
     end;

type PRec2FORMULA = ^TRec2FORMULA;
     TRec2FORMULA = packed record
     Row: word;
     Col: byte;
     Res: byte;
     Attribute: array[0..2] of byte;
     Value: double;
     Options: byte;
     ParseLen: byte;
     Data: array[0..255] of byte;
     end;

type PRec2STRING = ^TRec2STRING;
     TRec2STRING = packed record
     Len: byte;
     Data: array[0..255] of byte;
     end;

type PRecARRAY = ^TRecARRAY;
     TRecARRAY = packed record
     Row1,Row2: word;
     Col1,Col2: byte;
     Options: word;
     Unused: longword;
     DataSize: word;
     Data: record end;
     end;

type PRecBLANK = ^TRecBLANK;
     TRecBLANK = packed record
     Row: word;
     Col: word;
     FormatIndex: word;
     end;

type PRecBOF8 = ^TRecBOF8;
     TRecBOF8 = packed record
     VersionNumber: word;
     SubstreamType: word;
     BuildIdentifier: word;
     BuildYear: word;
     FileHistoryFlags: longint;
     LowBIFF: longint;
     end;

type PRecBOF7 = ^TRecBOF7;
     TRecBOF7 = packed record
     VersionNumber: word;
     SubstreamType: word;
     BuildIdentifier: word;
     BuildYear: word;
     end;

type PRecBOF4 = ^TRecBOF4;
     TRecBOF4 = packed record
     A: word;  // Don't know what this is.
     B: word;  //
     C: word;  //
     end;

type PRecBOOLERR = ^TRecBOOLERR;
     TRecBOOLERR = packed record
     Row: word;
     Col: word;
     FormatIndex: word;
     BoolErr: byte;
     Error: byte;
     end;

type PRecBOUNDSHEET7 = ^TRecBOUNDSHEET7;
     TRecBOUNDSHEET7 = packed record
     BOFPos: longint;
     Options: word;
     NameLen: byte;
     Name: array[0..255] of byte;
     end;

type PRecBOUNDSHEET8 = ^TRecBOUNDSHEET8;
     TRecBOUNDSHEET8 = packed record
     BOFPos: longint;
     Options: word;
     NameLen: byte;
     Name: array[0..65535] of byte;
     end;

type PRecCODEPAGE = ^TRecCODEPAGE;
     TRecCODEPAGE = packed record
     Codepage: word;
     end;


type PRecCOLWIDTH = ^TRecCOLWIDTH;
     TRecCOLWIDTH = packed record
     Col1,Col2: byte;
     Width: word;
     end;

type PRecCOLINFO = ^TRecCOLINFO;
     TRecCOLINFO = packed record
     Col1,Col2: word;
     Width: word;
     FormatIndex: word;
     Options: word;
     Reserved: word;
     end;

type PRecCOUNTRY = ^TRecCOUNTRY;
     TRecCOUNTRY = packed record
     DefaultCountryIndex: word;
     WinIniCountry: word;
     end;

type POPERNum = ^TOPERNum;
     TOPERNum = packed record
     OperType: byte;
     Value: double;
     end;

type POPERStr = ^TOPERStr;
     TOPERStr = packed record
     OperType: byte;
     StrLen: byte;
     Value: array[0..1023] of char;
     end;

type POPERBool = ^TOPERBool;
     TOPERBool = packed record
     OperType: byte;
     Value: word;
     Unused: array[0..5] of byte;
     end;

type POPERErr = ^TOPERErr;
     TOPERErr = packed record
     OperType: byte;
     Value: word;
     Unused: array[0..5] of byte;
     end;

type PRecCRN = ^TRecCRN;
     TRecCRN = packed record
     LastCol, FirstCol: byte;
     Row: word;
     OPER: array[0..1023] of TOperNum;
     end;

type PCFFont = ^TCFFont;
     TCFFont = packed record
     Reserved1: array[0..63] of byte;
     Height: longword;
     Options: longword;
     Weight: word;
     Escapement: word;
     Underline: byte;
     Reserved2: array[0..2] of byte;
     ColorIndex: longword;
     Reserved3: array[0..3] of byte;
     ModifiedFlags: longword;
     EscapementModified: longword;
     UnderlineModified: longword;
     Reserved4: array[0..15] of byte;
     Is0001: word;
     end;

type PCFBorder = ^TCFBorder;
     TCFBorder = packed record
     Linestyle: word;
     ColorIndex: longword;
     Reserved: word;
     end;

type PCFPattern = ^TCFPattern;
     TCFPattern = packed record
     PatternStyle: word;
     ColorIndex: word;
     end;

type PRecCF = ^TRecCF;
     TRecCF = packed record
     CFType: byte;
     Operator_: byte;
     Formula1Size: word;
     Formula2Size: word;
     Flags: longword;
     Reserved: word;
     VarData: record end;
     end;

type PRecCONDFMT = ^TRecCONDFMT;
     TRecCONDFMT = packed record
     CFCount: word;
     Flags: word;
     AreaExt: TRecCellArea;
     AreaCount: word;
     Areas: array[0..0] of TRecCellArea;
     end;

type PRecDEFAULTROWHEIGHT = ^TRecDEFAULTROWHEIGHT;
     TRecDEFAULTROWHEIGHT = packed record
     Options: word;
     Height: word;
     end;

type PRecDBCELL = ^TRecDBCELL;
     TRecDBCELL = packed record
     RowOffset: longint;
     Offsets: array[0..32768] of smallint;
     end;

type PRecDIMENSIONS2 = ^TRecDIMENSIONS2;
     TRecDIMENSIONS2 = packed record
     FirstRow,LastRow: word;
     FirstCol,LastCol: word;
     end;

type PRecDIMENSIONS7 = ^TRecDIMENSIONS7;
     TRecDIMENSIONS7 = packed record
     FirstRow,LastRow: word;
     FirstCol,LastCol: word;
     Reserved: word;
     end;

type PRecDIMENSIONS8 = ^TRecDIMENSIONS8;
     TRecDIMENSIONS8 = packed record
     FirstRow,LastRow: longint;
     FirstCol,LastCol: word;
     Reserved: word;
     end;

type PRecDVAL = ^TRecDVAL;
     TRecDVAL = packed record
     Options: word;
     X,Y: longword;
     DropDownId: longword;
     DVCount: longword;
     end;

type TRecISSTINF = packed record
     Pos: longint;
     Offset: word;
     Reserved: word;
     end;

type PRecEXTSST = ^TRecEXTSST;
     TRecEXTSST = packed record
     BucketSize: word;
     ISSTINF: record end;
     end;

type PRecEXTERNNAME8 = ^TRecEXTERNNAME8;
     TRecEXTERNNAME8  = packed record
     Options: word;
     Reserved: longint;
     LenName: byte;
     Data: array[0..65535] of byte;
     end;

type PRecEXTERNNAME5 = ^TRecEXTERNNAME5;
     TRecEXTERNNAME5  = packed record
     Options: word;
     Reserved: longint;
     NameLen: byte;
     Name: array[0..255] of byte;
     end;

type PRecEXTERNNAME3 = ^TRecEXTERNNAME3;
     TRecEXTERNNAME3  = packed record
     Options: word;
     NameLen: byte;
     Name: array[0..255] of byte;
     end;

type PRecEXTERNSHEET7 = ^TRecEXTERNSHEET7;
     TRecEXTERNSHEET7  = packed record
     Len: byte;
     Data: array[0..255] of byte;
     end;

type TXTI = packed record
     SupBook: word;
     FirstTab: word;
     LastTab: word;
     end;

type PRecEXTERNSHEET8 = ^TRecEXTERNSHEET8;
     TRecEXTERNSHEET8  = packed record
     XTICount: word;
     XTI: array[0..255] of TXTI;
     end;

type PRecFILEPASS = ^TRecFILEPASS;
     TRecFILEPASS = packed record
     Options: word;                      // Zero when SillyPassword is set,
                                         // otherwiese one. When SillyPassword
                                         // is used, the entire record is 6 bytes.
     SillyPassword: array[0..3] of byte; // XOR:ed password for french excel.
     DocId: array[0..15] of byte;
     Salt: array[0..15] of byte;
     HashedSalt: array[0..15] of byte;
     end;

type PRecFONT4 = ^TRecFONT4;
     TRecFONT4 = packed record
     Height: word;
     Attributes: word;
     Color: word; // $7FFF = default
     NameLen: byte;
     Name: array[0..255] of byte;
     end;

type PRecFONT = ^TRecFONT;
     TRecFONT = packed record
     Height: word;
     Attributes: word;
     ColorIndex: word;
     Bold: word;
     SubSuperScript: word;
     Underline: byte;
     Family: byte;
     CharSet: byte;
     Reserved: byte;
     NameLen: byte;
     Name: array[0..255] of byte; 
     end;

type PRecFORMAT4 = ^TRecFORMAT4;
     TRecFORMAT4 = packed record
     Unused: word;
     Len: byte;
     Data: array[0..255] of byte;
     end;

type PRecFORMAT7 = ^TRecFORMAT7;
     TRecFORMAT7 = packed record
     Index: word;
     Len: byte;
     Data: record end;
     end;

type PRecFORMAT8 = ^TRecFORMAT8;
     TRecFORMAT8 = packed record
     Index: word;
     Len: word;
     Data: record end;
     end;

type PRecFORMAT2 = ^TRecFORMAT2;
     TRecFORMAT2 = packed record
     Len: byte;
     Data: record end;
     end;

type PRecFORMULA = ^TRecFORMULA;
     TRecFORMULA = packed record
     Row,Col: word;
     FormatIndex: word;
     Value: double;
     Options: word;
     Reserved: longint;
     ParseLen: word;
     Data: array[0..High(word)] of byte;
     end;

type TRecFORMULA_ = packed record
     Row,Col: word;
     FormatIndex: word;
     Value: double;
     Options: word;
     Reserved: longint;
     ParseLen: word;
     end;

type PRecFORMULA3 = ^TRecFORMULA3;
     TRecFORMULA3 = packed record
     Row,Col: word;
     FormatIndex: word;
     Value: double;
     Options: word;
     ParseLen: word;
     Data: array[0..High(word)] of byte;
     end;

type PRecGUTS = ^TRecGUTS;
     TRecGUTS = packed record
     SizeRow: word;
     SizeCol: word;
     LevelRow: word;
     LevelCol: word;
     end;

type PRecHLINK = ^TRecHLINK;
     TRecHLINK = packed record
     Row1,Row2: word;
     Col1,Col2: word;
     GUID: array[0..15] of byte;
     Reserved: longword;
     Options: longword;
     end;

type PRecHLINKTOOLTIP = ^TRecHLINKTOOLTIP;
     TRecHLINKTOOLTIP = packed record
     RecId: word;
     Row1,Row2: word;
     Col1,Col2: word;
     Text: record end;
     end;

type PRecHORIZONTALPAGEBREAKS = ^TRecHORIZONTALPAGEBREAKS;
     TRecHORIZONTALPAGEBREAKS = packed record
     Count: word;
     Breaks: array[0..1024] of TPageBreak;
     end;

type PRecINDEX7 = ^TRecINDEX7;
     TRecINDEX7 = packed record
     Reserved1: longint;
     Row1: word;
     Row2: word;
     Reserved2: longint;
     Offsets: array[0..(High(integer) div 4) - 4] of integer;
     end;

type PRecINDEX8 = ^TRecINDEX8;
     TRecINDEX8 = packed record
     Reserved1: longint;
     Row1: longint;
     Row2: longint;
     Reserved2: longint;
     Offsets: array[0..(High(integer) div 4) - 40] of integer;
     end;

type PRecLABELSST = ^TRecLABELSST;
     TRecLABELSST = packed record
     Row,Col: word;
     FormatIndex: word;
     SSTIndex: longint;
     end;

type PRecLABEL = ^TRecLABEL;
     TRecLABEL = packed record
     Row,Col: word;
     FormatIndex: word;
     Len: word;   // String lengths may not exceed 255 bytes.
     Data: array[0..255] of byte;
     end;

// Left, Right, Top, Bottom margins
type PRecMARGIN = ^TRecMARGIN;
     TRecMARGIN = packed record
     Value: double;
     end;

type PRecMERGEDCELLS = ^TRecMERGEDCELLS;
     TRecMERGEDCELLS = packed record
     Count: word;
     Cells: array[0..8191] of TRecCellArea;
     end;

type PRecMULBLANK = ^TRecMULBLANK;
     TRecMULBLANK = packed record
     Row: word;
     Col1: word;
     FormatIndexes: array[0..65535] of word;
     end;

type TMULRK = packed record
     XF: word;
     RK: longint;
     end;

type PRecMULRK = ^TRecMULRK;
     TRecMULRK = packed record
     Row: word;
     Col1: word;
     RKs: array[0..65535] of TMulRK;
     end;

type PRecNUMBER = ^TRecNUMBER;
     TRecNUMBER = packed record
     Row,Col: word;
     FormatIndex: word;
     Value: double;
     end;

type PRecNAME = ^TRecNAME;
     TRecNAME = packed record
     Options: word;
     KeyShortcut: byte;
     LenName: byte;
     LenNameDef: word;
     // Doc's are wrong. LocalSheetIndex is the second value.
     Unused: word;
     LocalSheetIndex: word;

     LenCustMenu: byte;
     LenDescText: byte;
     LenHelpText: byte;
     LenStatusText: byte;
     Data: record end;
     end;

type PRecNAME3 = ^TRecNAME3;
     TRecNAME3 = packed record
     Options: word;
     Res1: byte;
     LenName: byte;
     Res2: word;
     Data: array[0..255] of byte;
     end;

const TRecNOTE_FIXEDLEN = 8;
type PRecNOTE = ^TRecNOTE;
     TRecNOTE = packed record
     Row: word;
     Col: word;
     Options: word;
     ObjId: word;
     AuthorNameLen: word;
     AuthorName: array[0..32767] of byte;
     end;

// Only first subrecord OBJ record.
type PRecOBJ = ^TRecOBJ;
     TRecOBJ = packed record
     RecId: word;
     Length: word;
     ObjType: word;
     ObjId: word;
     Options: word;
     Reserved: array[0..11] of byte;
     end;

type PRecPALETTE = ^TRecPALETTE;
     TRecPALETTE = packed record
     Count: word;
     Color: array[0..65535] of longint;
     end;

type PRecPANE = ^TRecPANE;
     TRecPANE = packed record
     X,Y: word;
     TopRow: word;
     LeftCol: word;
     PaneNumber: word;
     end;

type PRecRECALCID = ^TRecRECALCID;
     TRecRECALCID = packed record
     RecordIdRepeated: word;
     Reserved: word;
     RecalcEngineId: longword;
     end;

type PRecROW = ^TRecROW;
     TRecROW = packed record
     Row: word;
     Col1,Col2: word;
     Height: word;
     Reserved1,Reserved2: word;
     Options: word;
     FormatIndex: word;
     end;

type TRK = packed record
     case integer of
       0: (V: double);
       1: (DW: array [0..1] of longint);
     end;

type PRecRK = ^TRecRK;
     TRecRK = packed record
     Row,Col: word;
     FormatIndex: word;
     Value: longword;
     end;

type PRecRSTRING = ^TRecRSTRING;
     TRecRSTRING = packed record
     Row,Col: word;
     FormatIndex: word;
     Len: word;
     Data: array[0..65535] of byte;
     end;

type PRecSCL = ^TRecSCL;
     TRecSCL = packed record
     Numerator: word;
     Denominator: word;
     end;

type PRecSELECTION = ^TRecSELECTION;
     TRecSELECTION = packed record
     Pane: byte;
     ActiveRow: word;
     ActiveCol: word;
     ActiveRef: word;
     Refs: word;
     Row1,Row2: word;
     Col1,Col2:byte
     end;

// This is a better definition than the previous.
type PRecSELECTION_2 = ^TRecSELECTION_2;
     TRecSELECTION_2 = packed record
     Pane: byte;
     ActiveRow: word;
     ActiveCol: word;
     ActiveRef: word;
     RefCount: word;
     Refs: record end;
     end;

type PRecSETUP = ^TRecSETUP;
     TRecSETUP = packed record
     PaperSize: word;
     Scale: word;
     PageStart: word;
     FitWidth: word;
     FitHeight: word;
     Options: word;
     Resolution: word;
     VertResolution: word;
     HeaderMargin: double;
     FooterMargin: double;
     Copies: word;
     end;

type PRecSHEETEXT = ^TRecSHEETEXT;
     TRecSHEETEXT = packed record
     Frt: word;
     Flags: word;
     Unused1: longword;
     Unused2: longword;
     RecSizeMinus4: longword;
     TabColor: longword;
     end;

type PRecFEATHEADR = ^TRecFEATHEADR;
     TRecFEATHEADR = packed record
     Frt: word;
     Flags: word;
     Unused1: array[0..7] of byte;
     SharedFeatureType: word; // 2 for enhanced protections, the only thing of interrest.
     Hdr: byte;
     HdrDataSz: longword;
     Protections: longword;
     end;

type PRecSHRFMLA = ^TRecSHRFMLA;
     TRecSHRFMLA = packed record
     Row1,Row2: word;
     Col1,Col2: byte;
     Reserved: word;
     ParseLen: word;
     Data: array[0..65535] of byte;
     end;

type PRecSST = ^TRecSST;
     TRecSST = packed record
     Total: longword;
     Unique: longword;
     Data: array[0..MAXINT div 2] of byte;
     end;

type PRecSTRING = ^TRecSTRING;
     TRecSTRING = packed record
     Len: word;
     Data: array[0..65535] of byte;
     end;

type PRecSTRING1Byte = ^TRecSTRING1Byte;
     TRecSTRING1Byte = packed record
     Len: byte;
     Data: array[0..255] of byte;
     end;

type PRecSTYLE = ^TRecSTYLE;
     TRecSTYLE = packed record
     FormatIndex: word;
     BuiltInStyle: byte;
     Level: byte;
     end;

type PRecSTYLE_USER = ^TRecSTYLE_USER;
     TRecSTYLE_USER = packed record
     FormatIndex: word;
     Len: word;
     Data: record end;
     end;

type PRecSUPBOOK = ^TRecSUPBOOK;
     TRecSUPBOOK = packed record
     Tabs: word;
     Data: array[0..65535] of byte;
     end;

type PRecTXO = ^TRecTXO;
     TRecTXO = packed record
     Options: word;
     Orientation: word;
     Reserved: array[0..5] of byte;
     TextLen: word;
     FormatLen: word;
     // Undocumented: The size of the TXO record may be 32 bytes (normal 18).
     // This seems to occure when the text is in a Rectangle OBJ. If the
     // size is 32 bytes, Reserved2 has the value 000E0000. Reserved2 is
     // in two parts, as the last word seems to be the size of the extra data.
     Reserved2a: word;
     ExSize: word;
     ExData:  record end;
     end;

type PRecVERTICALPAGEBREAKS = ^TRecVERTICALPAGEBREAKS;
     TRecVERTICALPAGEBREAKS = packed record
     Count: word;
     Breaks: array[0..1024] of TPageBreak;
     end;

type PRecWINDOW1 = ^TRecWINDOW1;
     TRecWINDOW1 = packed record
     Left: word;
     Top: word;
     Width: word;
     Height: word;
     Options: word;
     SelectedTabIndex: word;
     FirstDispTabIndex: word;
     SelectedTabs: word;
     TabRatio: word;
     end;

type PRecWINDOW2_7 = ^TRecWINDOW2_7;
     TRecWINDOW2_7 = packed record
     Options: word;
     TopRow: word;
     LeftCol: word;
     HeaderColorIndex: longint;
     end;

type PRecWINDOW2_8 = ^TRecWINDOW2_8;
     TRecWINDOW2_8 = packed record
     Options: word;
     TopRow: word;
     LeftCol: word;
     HeaderColorIndex: longint;
     ZoomPreview: word;
     Zoom: word;
     Reserved: longint;
     end;

type PRecWINDOW2_3 = ^TRecWINDOW2_3;
     TRecWINDOW2_3 = packed record
     Options: word;
     TopRow: word;
     LeftCol: word;
     Unknown: longint;
     end;

type PRecXCT = ^TRecXCT;
     TRecXCT = packed record
     CRNCount: word;
     SheetIndex: word;
     end;

type PRecXF3 = ^TRecXF3;
     TRecXF3 = packed record
     FontIndex: byte;
     Data: array[0..10] of byte;
     end;

type PRecXF4 = ^TRecXF4;
     TRecXF4 = packed record
     FontIndex: byte;
     FormatIndex: byte;
     Data1: word;
     Alignment: byte;
     UsedAttributes: byte;
     CellColor: word;
     TopBorder: byte;
     LeftBorder: byte;
     BottomBorder: byte;
     RightBorder: byte;
     end;

type PRecXF7 = ^TRecXF7;
     TRecXF7 = packed record
     FontIndex: word;
     FormatIndex: word;
     Data1: word;
     Data2: word;
     Data3: word;
     Data4: word;
     Data5: word;
     Data6: word;
     end;

type PRecXF8 = ^TRecXF8;
     TRecXF8 = packed record
     FontIndex: word;
     NumFmtIndex: word;
     Data1: word;
     Data2: word;
     Data3: word;
     Data4: word;
     Data5: word;
     Data6: longword;
     Data7: word;
     end;

// *** Chart Records ***

type PCRec3D = ^TCRec3D;
     TCRec3D = packed record
     Rotation: word;
     Elevation: word;
     Distance: word;
     PlotHeight: word;
     Depth: word;
     Space: word;
     Flags: word;
     end;

type PCRecAI = ^TCRecAI;
     TCRecAI = packed record
     LinkType: byte;
     ReferenceType: byte;
     Flags: word;
     FormatIndex: word;
     FormulaSize: word;
     Formula: record end;
     end;

type PCRecAREA = ^TCRecAREA;
     TCRecAREA = packed record
     Flags: word;
     end;

type PCRecAREAFORMAT = ^TCRecAREAFORMAT;
     TCRecAREAFORMAT = packed record
     ColorFgRGB: longword;
     ColorBgRGB: longword;
     Pattern: word;
     Format: word;
     ColorFgIndex: word;
     ColorBgIndex: word;
     end;

type PCRecATTACHEDLABEL = ^TCRecATTACHEDLABEL;
     TCRecATTACHEDLABEL = packed record
     Flags: word;
     end;

type PCRecAXCEXT = ^TCRecAXCEXT;
     TCRecAXCEXT = packed record
     MinCategory: word;
     MaxCategory: word;
     ValueMajor: word;
     UnitsMajor: word;
     ValueMinor: word;
     UnitsMinor: word;
     BaseUnit: word;
     CrossingPoint: word;
     Flags: word;
     end;

type PCRecAXESUSED = ^TCRecAXESUSED;
     TCRecAXESUSED = packed record
     AxesSets: word;
     end;

type PCRecAXISLINEFORMAT = ^TCRecAXISLINEFORMAT;
     TCRecAXISLINEFORMAT = packed record
     LineId: word;
     end;

type PCRecAXISPARENT = ^TCRecAXISPARENT;
     TCRecAXISPARENT = packed record
     AxixIndex: word;
     Top: longword;
     Left: longword;
     Width: longword;
     Height: longword;
     end;

type PCRecAXIS = ^TCRecAXIS;
     TCRecAXIS = packed record
     AxisType: word;
     Reserved: array[0..15] of byte;
     end;

type PCRecBAR = ^TCRecBAR;
     TCRecBAR = packed record
     SpaceBars: word;
     SpaceCategories: word;
     Flags: word;
     end;

type PCRecCATSERRANGE = ^TCRecCATSERRANGE;
     TCRecCATSERRANGE = packed record
     CrossingPoint: word;
     FreqLabels: word;
     FreqMarks: word;
     Flags: word;
     end;

type PCRecCHART = ^TCRecCHART;
     TCRecCHART = packed record
     Top:    array[0..1] of word;
     Left:   array[0..1] of word;
     Width:  array[0..1] of word;
     Height: array[0..1] of word;
     end;

type PCRecCHARTFORMAT = ^TCRecCHARTFORMAT;
     TCRecCHARTFORMAT = packed record
     Reserved: array[0..15] of byte;
     Flags: word;
     DrawingOrder: word;
     end;

type PCRecDATAFORMAT = ^TCRecDATAFORMAT;
     TCRecDATAFORMAT = packed record
     PointNumber: word;
     SeriesIndex: word;
     SeriesNumber: word;
     Flags: word;
     end;

type PCRecDEFAULTTEXT = ^TCRecDEFAULTTEXT;
     TCRecDEFAULTTEXT = packed record
     Id: word;
     end;

type PCRecFBI = ^TCRecFBI;
     TCRecFBI = packed record
     Width: word;
     Height: word;
     HeightApplied: word;
     Scale: word;
     Index: word;
     end;

type PCRecFBI_Font = ^TCRecFBI_Font;
     TCRecFBI_Font = packed record
     Width: word;
     Height: word;
     HeightApplied: word;
     Scale: word;
     Index: word;
     Xc12Font: TObject;
     end;

type PCRecFONTX = ^TCRecFONTX;
     TCRecFONTX = packed record
     FontIndex: word;
     end;

type PCRecFRAME = ^TCRecFRAME;
     TCRecFRAME = packed record
     FrameType: word;
     Flags: word;
     end;

// Undocumented
type PCRecGEOMETRY = ^TCRecGEOMETRY;
     TCRecGEOMETRY = packed record
     ItemType: word;
     end;

type PCRecLEGEND_ = ^TCRecLEGEND_;
     TCRecLEGEND_ = packed record
     Top: longword;
     Left: longword;
     Width: longword;
     Height: longword;
     LegendType: byte;
     Spacing: byte;
     Flags: word;
     end;

type PCRecLINE = ^TCRecLINE;
     TCRecLINE = packed record
     Flags: word;
     end;

type PCRecLINEFORMAT = ^TCRecLINEFORMAT;
     TCRecLINEFORMAT = packed record
     ColorRGB: longword;
     Pattern: word;
     Weight: word;
     Format: word;
     ColorIndex: word;
     end;

type PCRecMARKERFORMAT = ^TCRecMARKERFORMAT;
     TCRecMARKERFORMAT = packed record
     ColorFgRGB: longword;
     ColorBgRGB: longword;
     TypeOfMarker: word;
     Format: word;
     ColorFgIndex: word;
     ColorBgIndex: word;
     Size: longword;
     end;

type PCRecOBJECTLINK = ^TCRecOBJECTLINK;
     TCRecOBJECTLINK = packed record
     LinkedTo: word;
     IndexSeries: word;
     IndexData: word;
     end;

type PCRecPLOTGROWTH = ^TCRecPLOTGROWTH;
     TCRecPLOTGROWTH = packed record
     Horizontal: array[0..1] of word;
     Vertical: array[0..1] of word;
     end;

type PCRecPICF = ^TCRecPICF;
     TCRecPICF = packed record
     PaintMode: word;
     ImageFormat: word;
     Environment: byte;
     Flags: byte;
     Scaling: double;
     end;

type PCRecPIE = ^TCRecPIE;
     TCRecPIE = packed record
     Angle: word;
     Donut: word;
     Flags: word;
     end;

type PCRecPIEFORMAT = ^TCRecPIEFORMAT;
     TCRecPIEFORMAT = packed record
     Percent: word;
     end;

type PCRecPOS = ^TCRecPOS;
     TCRecPOS = packed record
     TopLt: word;
     TopRt: word;
     X1: longword;
     Y1: longword;
     X2: longword;
     Y2: longword;
     end;

type PCRecRADAR = ^TCRecRADAR;
     TCRecRADAR = packed record
     Flags: word;
     end;

type PCRecRADARAREA = ^TCRecRADARAREA;
     TCRecRADARAREA = packed record
     Flags: word;
     end;

type PCRecSCATTER = ^TCRecSCATTER;
     TCRecSCATTER = packed record
     BubblePercent: word;
     BubbleSizeIs: word;
     Flags: word;
     end;

type PCRecSERIES = ^TCRecSERIES;
     TCRecSERIES = packed record
     CategoriesType: word;
     ValuesType: word;
     CategoriesCount: word;
     ValuesCount: word;
     BubbleType: word;
     BubbleCount: word;
     end;

type PCRecSERIESTEXT = ^TCRecSERIESTEXT;
     TCRecSERIESTEXT = packed record
     TextId: word;
     Length: byte;
     Text: record end;
     end;

type PCRecSERTOCRT = ^TCRecSERTOCRT;
     TCRecSERTOCRT = packed record
     ChartGroupIndex: word;
     end;

type PCRecSHTPROPS = ^TCRecSHTPROPS;
     TCRecSHTPROPS = packed record
     Flags: word;
     BlanksAs: byte;
     Unknown: byte;
     end;

type PCRecSIINDEX = ^TCRecSIINDEX;
     TCRecSIINDEX = packed record
     Index: word;
     end;

type PCRecSURFACE = ^TCRecSURFACE;
     TCRecSURFACE = packed record
     Flags: word;
     end;

type PCRecTEXT = ^TCRecTEXT;
     TCRecTEXT = packed record
     HorizAlignment: byte;
     VertAlignment: byte;
     BkgMode: word;
     ColorRGB: longword;
     Left: longword;
     Top: longword;
     Width: longword;
     Height: longword;
     Options1: word;
     ColorIndex: word;
     Options2: word;
     Rotation: word;
     end;

type PCRecTICK = ^TCRecTICK;
     TCRecTICK = packed record
     MajorMark: byte;
     MinorMark: byte;
     LabelPos: byte;
     BkgMode: byte;
     ColorRGB: longword;
     Reserved1: array[0..15] of byte;
     Flags: word;
     ColorIndex: word;
     Rotation: word;
     end;

type PCRecVALUERANGE = ^TCRecVALUERANGE;
     TCRecVALUERANGE = packed record
     MinValue: double;
     MaxValue: double;
     MajorInc: double;
     MinorInc: double;
     AxisCrosses: double;
     Flags: word;
     end;

implementation

end.


