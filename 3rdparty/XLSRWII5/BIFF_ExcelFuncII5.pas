unit BIFF_ExcelFuncII5;

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

uses XLSUtils5;

type PPTGName7 = ^TPTGName7;
     TPTGName7 = packed record
     NameIndex: word;
     Reserved: array[0..11] of byte;
     end;

type PPTGName8 = ^TPTGName8;
     TPTGName8 = packed record
     NameIndex: word;
     Reserved: word;
     end;

type PPTGNameX7 = ^TPTGNameX7;
     TPTGNameX7 = packed record
     ExtSheet: word;
     Reserved1: array[0..7] of byte;
     NameIndex: word;
     Reserved2: array[0..11] of byte;
     end;

type PPTGNameX8 = ^TPTGNameX8;
     TPTGNameX8 = packed record
     ExtSheet: word;
     NameIndex: word;
     Reserved: word;
     end;

type PPTGTbl = ^TPTGTbl;
     TPTGTbl = packed record
     Row: word;
     Col: word;
     end;

type PPTGRef7 = ^TPTGRef7;
     TPTGRef7 = packed record
     Row: word;
     Col: byte;
     end;

type PPTGRef8 = ^TPTGRef8;
     TPTGRef8 = packed record
     Row: word;
     Col: word;
     end;

type PPTGRef12 = ^TPTGRef12;
     TPTGRef12 = packed record
     Row: longword;
     Col: word;
     end;

type PPTGRef3d8 = ^TPTGRef3d8;
     TPTGRef3d8 = packed record
     Index: word;
     Row: word;
     Col: word;
     end;

type PPTGRef3d12 = ^TPTGRef3d12;
     TPTGRef3d12 = packed record
     Index: word;
     Row: longword;
     Col: word;
     end;

type PPTGArea3d8 = ^TPTGArea3d8;
     TPTGArea3d8 = packed record
     Index: word;
     Row1,Row2: word;
     Col1,Col2: word;
     end;

type PPTGArea3d12 = ^TPTGArea3d12;
     TPTGArea3d12 = packed record
     Index: word;
     Row1,Row2: longword;
     Col1,Col2: word;
     end;

type PPTGRef3d7 = ^TPTGRef3d7;
     TPTGRef3d7 = packed record
     Index: word;
     Res: array[0..7] of byte;
     IndexFirst: word;
     IndexLast: word;
     Row: word;
     Col: byte;
     end;

type PPTGArea3d7 = ^TPTGArea3d7;
     TPTGArea3d7 = packed record
     SheetIndex: word;
     Res: array[0..7] of byte;
     IndexFirst: word;
     IndexLast: word;
     Row1,Row2: word;
     Col1,Col2: byte;
     end;

type PPTGArea7 = ^TPTGArea7;
     TPTGArea7 = packed record
     Row1: word;
     Row2: word;
     Col1: byte;
     Col2: byte;
     end;

type PPTGArea8 = ^TPTGArea8;
     TPTGArea8 = packed record
     Row1: word;
     Row2: word;
     Col1: word;
     Col2: word;
     end;

type PPTGArea12 = ^TPTGArea12;
     TPTGArea12 = packed record
     Row1: longword;
     Row2: longword;
     Col1: word;
     Col2: word;
     end;

type TFunctionId = (
fiCount,fiIf,fiIsNa,fiIsError,fiSum,fiAverage,fiMin,fiMax,fiRow,fiColumn,fiNA,
fiNPV,fiSTDEV,fiDollar,fiFixed,fiSin,fiCos,fiTan,fiATan,fiPI,fiSqrt,fiExp,fiLN,
fiLog10,fiAbs,fiInt,fiSign,fiRound,fiLookup,fiIndex,fiRept,fiMid,fiLen,fiValue,
fiTrue,fiFalse,fiAnd,fiOr,fiNot,fiMod,fiDCount,fiDSum,fiDAverage,fiDMin,fiDMax,
fiDStdev,fiVar,fiDVar,fiText,fiLinest,fiTrend,fiLogest,fiGrowth,fiInt1,fiInt2,
fiInt3,fiPV,fiFV,fiNPER,fiPMT,fiRate,fiMirr,fiIrr,fiRand,fiMatch,fiDate,fiTime,
fiDay,fiMonth,fiYear,fiWeekday,fiHour,fiMinute,fiSecond,fiNow,fiAreas,fiRows,
fiColumns,fiOffset,fiInt4,fiInt5,fiInt6,fiSearch,fiTranspose,fiInt7,fiInt8,
fiType,fiInt9,fiInt10,fiInt11,fiInt12,fiInt13,fiInt14,fiInt15,fiInt16,fiInt17,
fiInt18,fiATan2,fiASin,fiACos,fiChoose,fiHLookup,fiVLookup,fiInt19,fiInt20,
fiIsref,fiInt21,fiInt22,fiInt23,fiLog,fiInt24,fiChar,fiLower,fiUpper,fiProper,
fiLeft,fiRight,fiExact,fiTrim,fiReplace,fiSubstitute,fiCode,fiInt25,fiInt26,
fiFind,fiCell,fiIsErr,fiIsText,fiIsNumber,fiIsblank,fiT,fiN,fiInt27,fiInt28,
fiInt29,fiInt30,fiInt31,fiInt32,fiInt33,fiInt34,fiDatevalue,fiTimevalue,fiSln,
fiSyd,fiDdb,fiInt35,fiInt36,fiInt37,fiIndirect,fiInt38,fiInt39,fiInt40,fiInt41,
fiInt42,fiInt43,fiInt44,fiInt45,fiInt46,fiInt47,fiInt48,fiInt49,fiInt50,fiClean,
fiMdeterm,fiMinverse,fiMmult,fiInt51,fiIpmt,fiPpmt,fiCounta,fiInt52,fiInt53,
fiInt54,fiInt55,fiInt56,fiInt57,fiInt58,fiInt59,fiInt60,fiInt61,fiInt62,fiInt63,
fiInt64,fiProduct,fiFact,fiInt65,fiInt66,fiInt67,fiInt68,fiDproduct,fiIsnontext,
fiInt69,fiInt70,fiStdevp,fiVarp,fiDstdevp,fiDvarp,fiTrunc,fiIslogical,fiDcounta,
fiInt71,fiInt72,fiInt73,fiInt74,fiInt75,fiInt76,fiInt77,fiInt78,fiInt79,fiInt80,
fiInt81,fiInt82,fiRoundup,fiRounddown,fiInt83,fiInt84,fiRank,fiInt85,fiInt86,
fiAddress,fiDays360,fiToday,fiVdb,fiInt87,fiInt88,fiInt89,fiInt90,fiMedian,
fiSumproduct,fiSinh,fiCosh,fiTanh,fiAsinh,fiAcosh,fiAtanh,fiDget,fiInt91,fiInt92,
fiInt93,fiInt94,fiInt95,fiInt96,fiInt97,fiInt98,fiInt100,fiInt101,fiInt102,fiDb,
fiInt103,fiInt104,fiInt105,fiInt106,fiFrequency,fiInt107,fiInt108,fiInt109,
fiInt110,fiInt111,fiInt112,fiInt113,fiInt114,fiInt115,fiInt116,fiInt117,fiInt118,
fiInt119,fiInt120,fiInt121,fiInt122,fiAvedev,fiBetadist,fiGammaln,fiBetainv,
fiBinomdist,fiChidist,fiChiinv,fiCombin,fiConfidence,fiCritbinom,fiEven,
fiExpondist,fiFdist,fiFinv,fiFisher,fiFisherinv,fiFloor,fiGammadist,fiGammainv,
fiCeiling,fiHypgeomdist,fiLognormdist,fiLoginv,fiNegbinomdist,fiNormdist,
fiNormsdist,fiNorminv,fiNormsinv,fiStandardize,fiOdd,fiPermut,fiPoisson,fiTdist,
fiWeibull,fiSumxmy2,fiSumx2my2,fiSumx2py2,fiChitest,fiCorrel,fiCovar,fiForecast,
fiFtest,fiIntercept,fiPearson,fiRsq,fiSteyx,fiSlope,fiTtest,fiProb,fiDevsq,
fiGeomean,fiHarmean,fiSumsq,fiKurt,fiSkew,fiZtest,fiLarge,fiSmall,fiQuartile,
fiPercentile,fiPercentrank,fiMode,fiTrimmean,fiTinv,fiInt123,fiInt124,fiInt125,
fiConcatenate,fiPower,
fiUserFunc // fiUserFunc must be an unused function id (here 338=PIVOT.ADD.DATA)
);

const TokenClassRef     = $20;
const TokenClassValue   = $40;
const TokenClassArray   = $60;
const TokenClassDefault = TokenClassValue;

type TExcelFunctions = packed record
     Min,Max: byte;
     Name: AxUCString;
     ArgClass: array[0..5] of byte;
     end;


type TExcelFunctionsArray = array[0..380] of TExcelFunctions;

// Min and Max = $FF; illegal function name.
var ExcelFunctions: TExcelFunctionsArray = (
{000} (Min: 01; Max: 99; Name: 'COUNT';    ArgClass: (TokenClassRef,0,0,0,0,0)),
{001} (Min: 02; Max: 03; Name: 'IF';       ArgClass: (TokenClassValue,TokenClassValue,TokenClassValue,0,0,0)),
{002} (Min: 01; Max: 01; Name: 'ISNA';     ArgClass: (TokenClassValue,0,0,0,0,0)),
{003} (Min: 01; Max: 01; Name: 'ISERROR';  ArgClass: (TokenClassValue,0,0,0,0,0)),
{004} (Min: 01; Max: 99; Name: 'SUM';      ArgClass: (TokenClassRef,0,0,0,0,0)),
{005} (Min: 01; Max: 99; Name: 'AVERAGE';  ArgClass: (TokenClassRef,0,0,0,0,0)),
{006} (Min: 01; Max: 99; Name: 'MIN';      ArgClass: (TokenClassRef,0,0,0,0,0)),
{007} (Min: 01; Max: 99; Name: 'MAX';      ArgClass: (TokenClassRef,0,0,0,0,0)),
{008} (Min: 00; Max: 01; Name: 'ROW';      ArgClass: (TokenClassRef,0,0,0,0,0)),
{009} (Min: 00; Max: 01; Name: 'COLUMN';   ArgClass: (TokenClassRef,0,0,0,0,0)),
{010} (Min: 00; Max: 00; Name: 'NA';       ArgClass: (0,0,0,0,0,0)),
{011} (Min: 02; Max: 99; Name: 'NPV';      ArgClass: (TokenClassValue,TokenClassRef,0,0,0,0)),
{012} (Min: 01; Max: 99; Name: 'STDEV';    ArgClass: (TokenClassValue,0,0,0,0,0)),
{013} (Min: 01; Max: 02; Name: 'DOLLAR';   ArgClass: (TokenClassValue,TokenClassValue,0,0,0,0)),
{014} (Min: 01; Max: 03; Name: 'FIXED';    ArgClass: (TokenClassValue,TokenClassValue,TokenClassValue,0,0,0)),
{015} (Min: 01; Max: 01; Name: 'SIN';      ArgClass: (TokenClassValue,0,0,0,0,0)),
{016} (Min: 01; Max: 01; Name: 'COS';      ArgClass: (TokenClassValue,0,0,0,0,0)),
{017} (Min: 01; Max: 01; Name: 'TAN';      ArgClass: (TokenClassValue,0,0,0,0,0)),
{018} (Min: 01; Max: 01; Name: 'ATAN';     ArgClass: (TokenClassValue,0,0,0,0,0)),
{019} (Min: 00; Max: 00; Name: 'PI';       ArgClass: (0,0,0,0,0,0)),
{020} (Min: 01; Max: 01; Name: 'SQRT';     ArgClass: (TokenClassValue,0,0,0,0,0)),
{021} (Min: 01; Max: 01; Name: 'EXP';      ArgClass: (TokenClassValue,0,0,0,0,0)),
{022} (Min: 01; Max: 01; Name: 'LN';       ArgClass: (TokenClassValue,0,0,0,0,0)),
{023} (Min: 01; Max: 01; Name: 'LOG10';    ArgClass: (TokenClassValue,0,0,0,0,0)),
{024} (Min: 01; Max: 01; Name: 'ABS';      ArgClass: (TokenClassValue,0,0,0,0,0)),
{025} (Min: 01; Max: 01; Name: 'INT';      ArgClass: (TokenClassValue,0,0,0,0,0)),
{026} (Min: 01; Max: 01; Name: 'SIGN';     ArgClass: (TokenClassValue,0,0,0,0,0)),
{027} (Min: 02; Max: 02; Name: 'ROUND';    ArgClass: (TokenClassValue,TokenClassValue,0,0,0,0)),
{028} (Min: 02; Max: 03; Name: 'LOOKUP';   ArgClass: (TokenClassValue,TokenClassRef,TokenClassRef,0,0,0)),
{029} (Min: 01; Max: 03; Name: 'INDEX';    ArgClass: (TokenClassRef,TokenClassValue,TokenClassValue,TokenClassValue,0,0)),
{030} (Min: 02; Max: 02; Name: 'REPT';     ArgClass: (TokenClassValue,TokenClassValue,0,0,0,0)),
{031} (Min: 01; Max: 03; Name: 'MID';      ArgClass: (TokenClassValue,TokenClassValue,TokenClassValue,0,0,0)),
{032} (Min: 01; Max: 01; Name: 'LEN';      ArgClass: (TokenClassValue,0,0,0,0,0)),
{033} (Min: 01; Max: 01; Name: 'VALUE';    ArgClass: (TokenClassValue,0,0,0,0,0)),
{034} (Min: 00; Max: 00; Name: 'TRUE';     ArgClass: (0,0,0,0,0,0)),
{035} (Min: 00; Max: 00; Name: 'FALSE';    ArgClass: (0,0,0,0,0,0)),
{036} (Min: 01; Max: 99; Name: 'AND';      ArgClass: (TokenClassValue,0,0,0,0,0)),
{037} (Min: 01; Max: 99; Name: 'OR';       ArgClass: (TokenClassValue,0,0,0,0,0)),
{038} (Min: 01; Max: 01; Name: 'NOT';      ArgClass: (TokenClassValue,0,0,0,0,0)),
{039} (Min: 02; Max: 02; Name: 'MOD';      ArgClass: (TokenClassValue,TokenClassValue,0,0,0,0)),
{040} (Min: 03; Max: 03; Name: 'DCOUNT';   ArgClass: (TokenClassRef,TokenClassRef,TokenClassRef,0,0,0)),
{041} (Min: 03; Max: 03; Name: 'DSUM';     ArgClass: (TokenClassRef,TokenClassRef,TokenClassRef,0,0,0)),
{042} (Min: 03; Max: 03; Name: 'DAVERAGE'; ArgClass: (TokenClassRef,TokenClassRef,TokenClassRef,0,0,0)),
{043} (Min: 03; Max: 03; Name: 'DMIN';     ArgClass: (TokenClassRef,TokenClassRef,TokenClassRef,0,0,0)),
{044} (Min: 03; Max: 03; Name: 'DMAX';     ArgClass: (TokenClassRef,TokenClassRef,TokenClassRef,0,0,0)),
{045} (Min: 03; Max: 03; Name: 'DSTDEV';   ArgClass: (TokenClassRef,TokenClassRef,TokenClassRef,0,0,0)),
{046} (Min: 01; Max: 99; Name: 'VAR';      ArgClass: (TokenClassRef,0,0,0,0,0)),
{047} (Min: 03; Max: 03; Name: 'DVAR';     ArgClass: (TokenClassRef,TokenClassRef,TokenClassRef,0,0,0)),
{048} (Min: 02; Max: 02; Name: 'TEXT';     ArgClass: (TokenClassValue,TokenClassValue,0,0,0,0)),
{049} (Min: 01; Max: 04; Name: 'LINEST';   ArgClass: (TokenClassRef,TokenClassRef,TokenClassValue,TokenClassValue,0,0)),
{050} (Min: 01; Max: 04; Name: 'TREND';    ArgClass: (TokenClassRef,TokenClassRef,TokenClassRef,TokenClassValue,0,0)),
{051} (Min: 01; Max: 04; Name: 'LOGEST';   ArgClass: (TokenClassRef,TokenClassRef,TokenClassValue,TokenClassValue,0,0)),
{052} (Min: 01; Max: 04; Name: 'GROWTH';   ArgClass: (TokenClassRef,TokenClassRef,TokenClassRef,TokenClassValue,0,0)),
{053} (Min: $FF; Max: $FF; Name: 'GOTO'),
{054} (Min: $FF; Max: $FF; Name: 'HALT'),
{055} (Min: $FF; Max: $FF; Name: 'RETURN'),
{056} (Min: 02; Max: 05; Name: 'PV';       ArgClass: (TokenClassValue,TokenClassValue,TokenClassValue,TokenClassValue,TokenClassValue,0)),
{057} (Min: 02; Max: 05; Name: 'FV';       ArgClass: (TokenClassValue,TokenClassValue,TokenClassValue,TokenClassValue,TokenClassValue,0)),
{058} (Min: 02; Max: 05; Name: 'NPER';     ArgClass: (TokenClassValue,TokenClassValue,TokenClassValue,TokenClassValue,TokenClassValue,0)),
{059} (Min: 03; Max: 05; Name: 'PMT';      ArgClass: (TokenClassValue,TokenClassValue,TokenClassValue,TokenClassValue,TokenClassValue,0)),
{060} (Min: 03; Max: 05; Name: 'RATE';     ArgClass: (TokenClassValue,TokenClassValue,TokenClassValue,TokenClassValue,TokenClassValue,TokenClassValue)),
{061} (Min: 03; Max: 03; Name: 'MIRR';     ArgClass: (TokenClassRef,TokenClassValue,TokenClassValue,0,0,0)),
{062} (Min: 01; Max: 02; Name: 'IRR';      ArgClass: (TokenClassRef,TokenClassValue,0,0,0,0)),
{063} (Min:00; Max: 00; Name: 'RAND';      ArgClass: (0,0,0,0,0,0)),
{064} (Min:02; Max: 99; Name: 'MATCH';     ArgClass: (TokenClassValue,TokenClassRef,0,0,0,0)),
{065} (Min:03; Max: 03; Name: 'DATE';      ArgClass: (TokenClassValue,TokenClassValue,TokenClassValue,0,0,0)),
{066} (Min:03; Max: 03; Name: 'TIME';      ArgClass: (TokenClassValue,TokenClassValue,TokenClassValue,0,0,0)),
{067} (Min:01; Max: 01; Name: 'DAY';       ArgClass: (TokenClassValue,0,0,0,0,0)),
{068} (Min:01; Max: 01; Name: 'MONTH';     ArgClass: (TokenClassValue,0,0,0,0,0)),
{069} (Min:01; Max: 01; Name: 'YEAR';      ArgClass: (TokenClassValue,0,0,0,0,0)),
{070} (Min:01; Max: 99; Name: 'WEEKDAY';   ArgClass: (TokenClassValue,0,0,0,0,0)),
{071} (Min:01; Max: 01; Name: 'HOUR';      ArgClass: (TokenClassValue,0,0,0,0,0)),
{072} (Min:01; Max: 01; Name: 'MINUTE';    ArgClass: (TokenClassValue,0,0,0,0,0)),
{073} (Min:01; Max: 01; Name: 'SECOND';    ArgClass: (TokenClassValue,0,0,0,0,0)),
{074} (Min:00; Max: 00; Name: 'NOW';       ArgClass: (0,0,0,0,0,0)),
{075} (Min:01; Max: 01; Name: 'AREAS';     ArgClass: (TokenClassRef,0,0,0,0,0)),
{076} (Min:01; Max: 01; Name: 'ROWS';      ArgClass: (TokenClassRef,0,0,0,0,0)),
{077} (Min:01; Max: 01; Name: 'COLUMNS';   ArgClass: (TokenClassRef,0,0,0,0,0)),
{078} (Min:02; Max: 99; Name: 'OFFSET';    ArgClass: (TokenClassRef,TokenClassValue,0,0,0,0)),
{079} (Min:255; Max: 255; Name: 'ABSREF'),
{080} (Min:255; Max: 255; Name: 'RELREF'),
{081} (Min:255; Max: 255; Name: 'ARGUMENT'),
{082} (Min:01; Max: 99; Name: 'SEARCH';    ArgClass: (TokenClassValue,0,0,0,0,0)),
{083} (Min:01; Max: 01; Name: 'TRANSPOSE'; ArgClass: (TokenClassArray,0,0,0,0,0)),
{084} (Min:255; Max: 255; Name: 'ERROR'),
{085} (Min:255; Max: 255; Name: 'STEP'),
{086} (Min:01; Max: 01; Name: 'TYPE';      ArgClass: (TokenClassValue,0,0,0,0,0)),
{087} (Min:255; Max: 255; Name: 'ECHO'),
{088} (Min:255; Max: 255; Name: 'SET.NAME'),
{089} (Min:255; Max: 255; Name: 'CALLER'),
{090} (Min:255; Max: 255; Name: 'DEREF'),
{091} (Min:255; Max: 255; Name: 'WINDOWS'),
{092} (Min:255; Max: 255; Name: 'SERIES'),
{093} (Min:255; Max: 255; Name: 'DOCUMENTS'),
{094} (Min:255; Max: 255; Name: 'ACTIVE.CELL'),
{095} (Min:255; Max: 255; Name: 'SELECTION'),
{096} (Min:255; Max: 255; Name: 'RESULT'),
{097} (Min:02; Max: 02; Name: 'ATAN2';     ArgClass: (TokenClassValue,TokenClassValue,0,0,0,0)),
{098} (Min:01; Max: 01; Name: 'ASIN';      ArgClass: (TokenClassValue,0,0,0,0,0)),
{099} (Min:01; Max: 01; Name: 'ACOS';      ArgClass: (TokenClassValue,0,0,0,0,0)),
{100} (Min:02; Max: 99; Name: 'CHOOSE';    ArgClass: (TokenClassValue,TokenClassRef,0,0,0,0)),
{101} (Min:03; Max: 99; Name: 'HLOOKUP';   ArgClass: (TokenClassValue,TokenClassRef,TokenClassRef,TokenClassValue,0,0)),
{102} (Min:03; Max: 99; Name: 'VLOOKUP';   ArgClass: (TokenClassValue,TokenClassRef,TokenClassRef,TokenClassValue,0,0)),
{103} (Min:255; Max: 255; Name: 'LINKS'),
{104} (Min:255; Max: 255; Name: 'INPUT'),
{105} (Min:01; Max: 01; Name: 'ISREF';     ArgClass: (TokenClassRef,0,0,0,0,0)),
{106} (Min:255; Max: 255; Name: 'GET.FORMULA'),
{107} (Min:255; Max: 255; Name: 'GET.NAME'),
{108} (Min:255; Max: 255; Name: 'SET.VALUE'),
{109} (Min:01; Max: 99; Name: 'LOG';       ArgClass: (TokenClassValue,0,0,0,0,0)),
{110} (Min:255; Max: 255; Name: 'EXEC'),
{111} (Min:01; Max: 01; Name: 'CHAR';      ArgClass: (TokenClassValue,0,0,0,0,0)),
{112} (Min:01; Max: 01; Name: 'LOWER';     ArgClass: (TokenClassValue,0,0,0,0,0)),
{113} (Min:01; Max: 01; Name: 'UPPER';     ArgClass: (TokenClassValue,0,0,0,0,0)),
{114} (Min:01; Max: 01; Name: 'PROPER';    ArgClass: (TokenClassValue,0,0,0,0,0)),
{115} (Min:01; Max: 99; Name: 'LEFT';      ArgClass: (TokenClassValue,0,0,0,0,0)),
{116} (Min:01; Max: 99; Name: 'RIGHT';     ArgClass: (TokenClassValue,0,0,0,0,0)),
{117} (Min:02; Max: 02; Name: 'EXACT';     ArgClass: (TokenClassValue,TokenClassValue,0,0,0,0)),
{118} (Min:01; Max: 01; Name: 'TRIM';      ArgClass: (TokenClassValue,0,0,0,0,0)),
{119} (Min:04; Max: 04; Name: 'REPLACE';   ArgClass: (TokenClassValue,TokenClassValue,TokenClassValue,TokenClassValue,0,0)),
{120} (Min:01; Max: 99; Name: 'SUBSTITUTE'; ArgClass: (TokenClassValue,0,0,0,0,0)),
{121} (Min:01; Max: 01; Name: 'CODE';      ArgClass: (TokenClassValue,0,0,0,0,0)),
{122} (Min:255; Max: 255; Name: 'NAMES'),
{123} (Min:255; Max: 255; Name: 'DIRECTORY'),
{124} (Min:01; Max: 99; Name: 'FIND';      ArgClass: (TokenClassValue,0,0,0,0,0)),
{125} (Min:02; Max: 99; Name: 'CELL';      ArgClass: (TokenClassValue,TokenClassRef,0,0,0,0)),
{126} (Min:01; Max: 01; Name: 'ISERR';     ArgClass: (TokenClassValue,0,0,0,0,0)),
{127} (Min:01; Max: 01; Name: 'ISTEXT';    ArgClass: (TokenClassValue,0,0,0,0,0)),
{128} (Min:01; Max: 01; Name: 'ISNUMBER';  ArgClass: (TokenClassValue,0,0,0,0,0)),
{129} (Min:01; Max: 01; Name: 'ISBLANK';   ArgClass: (TokenClassValue,0,0,0,0,0)),
{130} (Min:01; Max: 01; Name: 'T';         ArgClass: (TokenClassRef,0,0,0,0,0)),
{131} (Min:01; Max: 01; Name: 'N';         ArgClass: (TokenClassRef,0,0,0,0,0)),
{132} (Min:255; Max: 255; Name: 'FOPEN'),
{133} (Min:255; Max: 255; Name: 'FCLOSE'),
{134} (Min:255; Max: 255; Name: 'FSIZE'),
{135} (Min:255; Max: 255; Name: 'FREADLN'),
{136} (Min:255; Max: 255; Name: 'FREAD'),
{137} (Min:255; Max: 255; Name: 'FWRITELN'),
{138} (Min:255; Max: 255; Name: 'FWRITE'),
{139} (Min:255; Max: 255; Name: 'FPOS'),
{140} (Min:01; Max: 01; Name: 'DATEVALUE'; ArgClass: (TokenClassValue,0,0,0,0,0)),
{141} (Min:01; Max: 01; Name: 'TIMEVALUE'; ArgClass: (TokenClassValue,0,0,0,0,0)),
{142} (Min:03; Max: 03; Name: 'SLN';       ArgClass: (TokenClassValue,TokenClassValue,TokenClassValue,0,0,0)),
{143} (Min:04; Max: 04; Name: 'SYD';       ArgClass: (TokenClassValue,TokenClassValue,TokenClassValue,TokenClassValue,0,0)),
{144} (Min:01; Max: 99; Name: 'DDB';       ArgClass: (TokenClassValue,0,0,0,0,0)),
{145} (Min:255; Max: 255; Name: 'GET.DEF'),
{146} (Min:255; Max: 255; Name: 'REFTEXT'),
{147} (Min:255; Max: 255; Name: 'TEXTREF'),
{148} (Min:01; Max: 99; Name: 'INDIRECT';  ArgClass: (TokenClassValue,0,0,0,0,0)),
{149} (Min:255; Max: 255; Name: 'REGISTER'),
{150} (Min:255; Max: 255; Name: 'CALL'),
{151} (Min:255; Max: 255; Name: 'ADD.BAR'),
{152} (Min:255; Max: 255; Name: 'ADD.MENU'),
{153} (Min:255; Max: 255; Name: 'ADD.COMMAND'),
{154} (Min:255; Max: 255; Name: 'ENABLE.COMMAND'),
{155} (Min:255; Max: 255; Name: 'CHECK.COMMAND'),
{156} (Min:255; Max: 255; Name: 'RENAME.COMMAND'),
{157} (Min:255; Max: 255; Name: 'SHOW.BAR'),
{158} (Min:255; Max: 255; Name: 'DELETE.MENU'),
{159} (Min:255; Max: 255; Name: 'DELETE.COMMAND'),
{160} (Min:255; Max: 255; Name: 'GET.CHART.ITEM'),
{161} (Min:255; Max: 255; Name: 'DIALOG.BOX'),
{162} (Min:01; Max: 01; Name: 'CLEAN';     ArgClass: (TokenClassValue,0,0,0,0,0)),
{163} (Min:01; Max: 01; Name: 'MDETERM';   ArgClass: (TokenClassArray,0,0,0,0,0)),
{164} (Min:01; Max: 01; Name: 'MINVERSE';  ArgClass: (TokenClassArray,0,0,0,0,0)),
{165} (Min:02; Max: 02; Name: 'MMULT';     ArgClass: (TokenClassArray,TokenClassArray,0,0,0,0)),
{166} (Min:255; Max: 255; Name: 'FILES'),
{167} (Min:01; Max: 99; Name: 'IPMT';      ArgClass: (TokenClassValue,0,0,0,0,0)),
{168} (Min:01; Max: 99; Name: 'PPMT';      ArgClass: (TokenClassValue,0,0,0,0,0)),
{169} (Min:01; Max: 99; Name: 'COUNTA';    ArgClass: (TokenClassRef,0,0,0,0,0)),
{170} (Min:255; Max: 255; Name: 'CANCEL.KEY'),
{171} (Min:255; Max: 255; Name: 'FOR'),
{172} (Min:255; Max: 255; Name: 'WHILE'),
{173} (Min:255; Max: 255; Name: 'BREAK'),
{174} (Min:255; Max: 255; Name: 'NEXT'),
{175} (Min:255; Max: 255; Name: 'INITIATE'),
{176} (Min:255; Max: 255; Name: 'REQUEST'),
{177} (Min:255; Max: 255; Name: 'POKE'),
{178} (Min:255; Max: 255; Name: 'EXECUTE'),
{179} (Min:255; Max: 255; Name: 'TERMINATE'),
{180} (Min:255; Max: 255; Name: 'RESTART'),
{181} (Min:255; Max: 255; Name: 'HELP'),
{182} (Min:255; Max: 255; Name: 'GET.BAR'),
{183} (Min:01; Max: 99; Name: 'PRODUCT';   ArgClass: (TokenClassRef,0,0,0,0,0)),
{184} (Min:01; Max: 01; Name: 'FACT';      ArgClass: (TokenClassValue,0,0,0,0,0)),
{185} (Min:255; Max: 255; Name: 'GET.CELL'),
{186} (Min:255; Max: 255; Name: 'GET.WORKSPACE'),
{187} (Min:255; Max: 255; Name: 'GET.WINDOW'),
{188} (Min:255; Max: 255; Name: 'GET.DOCUMENT'),
{189} (Min:03; Max: 03; Name: 'DPRODUCT';  ArgClass: (TokenClassRef,TokenClassRef,TokenClassRef,0,0,0)),
{190} (Min:01; Max: 01; Name: 'ISNONTEXT'; ArgClass: (TokenClassValue,0,0,0,0,0)),
{191} (Min:255; Max: 255; Name: 'GET.NOTE'),
{192} (Min:255; Max: 255; Name: 'NOTE'),
{193} (Min:01; Max: 99; Name: 'STDEVP';    ArgClass: (TokenClassRef,0,0,0,0,0)),
{194} (Min:01; Max: 99; Name: 'VARP';      ArgClass: (TokenClassRef,0,0,0,0,0)),
{195} (Min:03; Max: 03; Name: 'DSTDEVP';   ArgClass: (TokenClassRef,TokenClassRef,TokenClassRef,0,0,0)),
{196} (Min:03; Max: 03; Name: 'DVARP';     ArgClass: (TokenClassRef,TokenClassRef,TokenClassRef,0,0,0)),
{197} (Min:01; Max: 99; Name: 'TRUNC';     ArgClass: (TokenClassValue,0,0,0,0,0)),
{198} (Min:01; Max: 01; Name: 'ISLOGICAL'; ArgClass: (TokenClassValue,0,0,0,0,0)),
{199} (Min:03; Max: 03; Name: 'DCOUNTA';   ArgClass: (TokenClassRef,TokenClassRef,TokenClassRef,0,0,0)),
{200} (Min:255; Max: 255; Name: 'DELETE.BAR'),
{201} (Min:255; Max: 255; Name: 'UNREGISTER'),
{202} (Min:255; Max: 255; Name: '202'),
{203} (Min:255; Max: 255; Name: '203'),
{204} (Min:255; Max: 255; Name: 'USDOLLAR'),
{205} (Min:255; Max: 255; Name: 'FINDB';   ArgClass: (TokenClassValue,0,0,0,0,0)),
{206} (Min:255; Max: 255; Name: 'SEARCHB'; ArgClass: (TokenClassValue,0,0,0,0,0)),
{207} (Min:255; Max: 255; Name: 'REPLACEB';ArgClass: (TokenClassValue,TokenClassValue,TokenClassValue,TokenClassValue,0,0)),
{208} (Min:255; Max: 255; Name: 'LEFTB';   ArgClass: (TokenClassValue,0,0,0,0,0)),
{209} (Min:255; Max: 255; Name: 'RIGHTB';  ArgClass: (TokenClassValue,0,0,0,0,0)),
{210} (Min:255; Max: 255; Name: 'MIDB';    ArgClass: (TokenClassValue,TokenClassValue,TokenClassValue,0,0,0)),
{211} (Min:255; Max: 255; Name: 'LENB';    ArgClass: (TokenClassValue,0,0,0,0,0)),
{212} (Min:02; Max: 02; Name: 'ROUNDUP';   ArgClass: (TokenClassValue,TokenClassValue,0,0,0,0)),
{213} (Min:02; Max: 02; Name: 'ROUNDDOWN'; ArgClass: (TokenClassValue,TokenClassValue,0,0,0,0)),
{214} (Min:255; Max: 255; Name: 'ASC'),
{215} (Min:255; Max: 255; Name: 'DBCS'),
{216} (Min:03; Max: 99; Name: 'RANK';      ArgClass: (TokenClassValue,TokenClassRef,TokenClassValue,0,0,0)),
{217} (Min:255; Max: 255; Name: '217'),
{218} (Min:255; Max: 255; Name: '218'),
{219} (Min:05; Max: 99; Name: 'ADDRESS';   ArgClass: (TokenClassValue,TokenClassValue,TokenClassValue,TokenClassValue,TokenClassValue,0)),
{220} (Min:01; Max: 99; Name: 'DAYS360';   ArgClass: (TokenClassValue,0,0,0,0,0)),
{221} (Min:00; Max: 00; Name: 'TODAY';     ArgClass: (0,0,0,0,0,0)),
{222} (Min:01; Max: 99; Name: 'VDB';       ArgClass: (TokenClassValue,0,0,0,0,0)),
{223} (Min:255; Max: 255; Name: 'ELSE'),
{224} (Min:255; Max: 255; Name: 'ELSE.IF'),
{225} (Min:255; Max: 255; Name: 'END.IF'),
{226} (Min:255; Max: 255; Name: 'FOR.CELL'),
{227} (Min:01; Max: 99; Name: 'MEDIAN';    ArgClass: (TokenClassRef,0,0,0,0,0)),
{228} (Min:01; Max: 99; Name: 'SUMPRODUCT'; ArgClass: (TokenClassArray,0,0,0,0,0)),
{229} (Min:01; Max: 01; Name: 'SINH';      ArgClass: (TokenClassValue,0,0,0,0,0)),
{230} (Min:01; Max: 01; Name: 'COSH';      ArgClass: (TokenClassValue,0,0,0,0,0)),
{231} (Min:01; Max: 01; Name: 'TANH';      ArgClass: (TokenClassValue,0,0,0,0,0)),
{232} (Min:01; Max: 01; Name: 'ASINH';     ArgClass: (TokenClassValue,0,0,0,0,0)),
{233} (Min:01; Max: 01; Name: 'ACOSH';     ArgClass: (TokenClassValue,0,0,0,0,0)),
{234} (Min:01; Max: 01; Name: 'ATANH';     ArgClass: (TokenClassValue,0,0,0,0,0)),
{235} (Min:03; Max: 03; Name: 'DGET';      ArgClass: (TokenClassRef,TokenClassRef,TokenClassRef,0,0,0)),
{236} (Min:255; Max: 255; Name: 'CREATE.OBJECT'),
{237} (Min:255; Max: 255; Name: 'VOLATILE'),
{238} (Min:255; Max: 255; Name: 'LAST.ERROR'),
{239} (Min:255; Max: 255; Name: 'CUSTOM.UNDO'),
{240} (Min:255; Max: 255; Name: 'CUSTOM.REPEAT'),
{241} (Min:255; Max: 255; Name: 'FORMULA.CONVERT'),
{242} (Min:255; Max: 255; Name: 'GET.LINK.INFO'),
{243} (Min:255; Max: 255; Name: 'TEXT.BOX'),
{244} (Min:01; Max: 01; Name: 'INFO';      ArgClass: (TokenClassValue,0,0,0,0,0)),
{245} (Min:255; Max: 255; Name: 'GROUP'),
{246} (Min:255; Max: 255; Name: 'GET.OBJECT'),
{247} (Min:01; Max: 99; Name: 'DB';        ArgClass: (TokenClassValue,0,0,0,0,0)),
{248} (Min:255; Max: 255; Name: 'PAUSE'),
{249} (Min:255; Max: 255; Name: '249'),
{250} (Min:255; Max: 255; Name: '250'),
{251} (Min:255; Max: 255; Name: 'RESUME'),
{252} (Min:02; Max: 02; Name: 'FREQUENCY'; ArgClass: (TokenClassRef,TokenClassRef,0,0,0,0)),
{253} (Min:255; Max: 255; Name: 'ADD.TOOLBAR'),
{254} (Min:255; Max: 255; Name: 'DELETE.TOOLBAR'),
{255} (Min:255; Max: 255; Name: '255'),
{256} (Min:255; Max: 255; Name: 'RESET.TOOLBAR'),
{257} (Min:255; Max: 255; Name: 'EVALUATE'),
{258} (Min:255; Max: 255; Name: 'GET.TOOLBAR'),
{259} (Min:255; Max: 255; Name: 'GET.TOOL'),
{260} (Min:255; Max: 255; Name: 'SPELLING.CHECK'),
{261} (Min:01; Max: 01; Name: 'ERROR.TYPE'; ArgClass: (TokenClassRef,0,0,0,0,0)),
{262} (Min:255; Max: 255; Name: 'APP.TITLE'),
{263} (Min:255; Max: 255; Name: 'WINDOW.TITLE'),
{264} (Min:255; Max: 255; Name: 'SAVE.TOOLBAR'),
{265} (Min:255; Max: 255; Name: 'ENABLE.TOOL'),
{266} (Min:255; Max: 255; Name: 'PRESS.TOOL'),
{267} (Min:255; Max: 255; Name: 'REGISTER.ID'),
{268} (Min:255; Max: 255; Name: 'GET.WORKBOOK'),
{269} (Min:01; Max: 99; Name: 'AVEDEV';    ArgClass: (TokenClassRef,0,0,0,0,0)),
{270} (Min:01; Max: 99; Name: 'BETADIST';  ArgClass: (TokenClassValue,0,0,0,0,0)),
{271} (Min:01; Max: 01; Name: 'GAMMALN';   ArgClass: (TokenClassValue,0,0,0,0,0)),
{272} (Min:01; Max: 99; Name: 'BETAINV';   ArgClass: (TokenClassValue,0,0,0,0,0)),
{273} (Min:04; Max: 04; Name: 'BINOMDIST'; ArgClass: (TokenClassValue,TokenClassValue,TokenClassValue,TokenClassValue,0,0)),
{274} (Min:02; Max: 02; Name: 'CHIDIST';   ArgClass: (TokenClassValue,TokenClassValue,0,0,0,0)),
{275} (Min:02; Max: 02; Name: 'CHIINV';    ArgClass: (TokenClassValue,TokenClassValue,0,0,0,0)),
{276} (Min:02; Max: 02; Name: 'COMBIN';    ArgClass: (TokenClassValue,TokenClassValue,0,0,0,0)),
{277} (Min:03; Max: 03; Name: 'CONFIDENCE'; ArgClass: (TokenClassValue,TokenClassValue,TokenClassValue,0,0,0)),
{278} (Min:03; Max: 03; Name: 'CRITBINOM'; ArgClass: (TokenClassValue,TokenClassValue,TokenClassValue,0,0,0)),
{279} (Min:01; Max: 01; Name: 'EVEN';      ArgClass: (TokenClassValue,0,0,0,0,0)),
{280} (Min:03; Max: 03; Name: 'EXPONDIST'; ArgClass: (TokenClassValue,TokenClassValue,TokenClassValue,0,0,0)),
{281} (Min:03; Max: 03; Name: 'FDIST';     ArgClass: (TokenClassValue,TokenClassValue,TokenClassValue,0,0,0)),
{282} (Min:03; Max: 03; Name: 'FINV';      ArgClass: (TokenClassValue,TokenClassValue,TokenClassValue,0,0,0)),
{283} (Min:01; Max: 01; Name: 'FISHER';    ArgClass: (TokenClassValue,0,0,0,0,0)),
{284} (Min:01; Max: 01; Name: 'FISHERINV'; ArgClass: (TokenClassValue,0,0,0,0,0)),
{285} (Min:02; Max: 02; Name: 'FLOOR';     ArgClass: (TokenClassValue,TokenClassValue,0,0,0,0)),
{286} (Min:04; Max: 04; Name: 'GAMMADIST'; ArgClass: (TokenClassValue,TokenClassValue,TokenClassValue,TokenClassValue,0,0)),
{287} (Min:03; Max: 03; Name: 'GAMMAINV';  ArgClass: (TokenClassValue,TokenClassValue,TokenClassValue,0,0,0)),
{288} (Min:02; Max: 02; Name: 'CEILING';   ArgClass: (TokenClassValue,TokenClassValue,0,0,0,0)),
{289} (Min:04; Max: 04; Name: 'HYPGEOMDIST'; ArgClass: (TokenClassValue,TokenClassValue,TokenClassValue,TokenClassValue,0,0)),
{290} (Min:03; Max: 03; Name: 'LOGNORMDIST'; ArgClass: (TokenClassValue,TokenClassValue,TokenClassValue,0,0,0)),
{291} (Min:03; Max: 03; Name: 'LOGINV';    ArgClass: (TokenClassValue,TokenClassValue,TokenClassValue,0,0,0)),
{292} (Min:03; Max: 03; Name: 'NEGBINOMDIST'; ArgClass: (TokenClassValue,TokenClassValue,TokenClassValue,0,0,0)),
{293} (Min:04; Max: 04; Name: 'NORMDIST';  ArgClass: (TokenClassValue,TokenClassValue,TokenClassValue,TokenClassValue,0,0)),
{294} (Min:01; Max: 01; Name: 'NORMSDIST'; ArgClass: (TokenClassValue,0,0,0,0,0)),
{295} (Min:03; Max: 03; Name: 'NORMINV';   ArgClass: (TokenClassValue,TokenClassValue,TokenClassValue,0,0,0)),
{296} (Min:01; Max: 01; Name: 'NORMSINV';  ArgClass: (TokenClassValue,0,0,0,0,0)),
{297} (Min:03; Max: 03; Name: 'STANDARDIZE'; ArgClass: (TokenClassValue,TokenClassValue,TokenClassValue,0,0,0)),
{298} (Min:01; Max: 01; Name: 'ODD';       ArgClass: (TokenClassValue,0,0,0,0,0)),
{299} (Min:02; Max: 02; Name: 'PERMUT';    ArgClass: (TokenClassValue,TokenClassValue,0,0,0,0)),
{300} (Min:03; Max: 03; Name: 'POISSON';   ArgClass: (TokenClassValue,TokenClassValue,TokenClassValue,0,0,0)),
{301} (Min:03; Max: 03; Name: 'TDIST';     ArgClass: (TokenClassValue,TokenClassValue,TokenClassValue,0,0,0)),
{302} (Min:04; Max: 04; Name: 'WEIBULL';   ArgClass: (TokenClassValue,TokenClassValue,TokenClassValue,TokenClassValue,0,0)),
{303} (Min:02; Max: 02; Name: 'SUMXMY2';   ArgClass: (TokenClassArray,TokenClassArray,0,0,0,0)),
{304} (Min:02; Max: 02; Name: 'SUMX2MY2';  ArgClass: (TokenClassArray,TokenClassArray,0,0,0,0)),
{305} (Min:02; Max: 02; Name: 'SUMX2PY2';  ArgClass: (TokenClassArray,TokenClassArray,0,0,0,0)),
{306} (Min:02; Max: 02; Name: 'CHITEST';   ArgClass: (TokenClassArray,TokenClassArray,0,0,0,0)),
{307} (Min:02; Max: 02; Name: 'CORREL';    ArgClass: (TokenClassArray,TokenClassArray,0,0,0,0)),
{308} (Min:02; Max: 02; Name: 'COVAR';     ArgClass: (TokenClassArray,TokenClassArray,0,0,0,0)),
{309} (Min:03; Max: 03; Name: 'FORECAST';  ArgClass: (TokenClassValue,TokenClassArray,TokenClassArray,0,0,0)),
{310} (Min:02; Max: 02; Name: 'FTEST';     ArgClass: (TokenClassArray,TokenClassArray,0,0,0,0)),
{311} (Min:02; Max: 02; Name: 'INTERCEPT'; ArgClass: (TokenClassArray,TokenClassArray,0,0,0,0)),
{312} (Min:02; Max: 02; Name: 'PEARSON';   ArgClass: (TokenClassArray,TokenClassArray,0,0,0,0)),
{313} (Min:02; Max: 02; Name: 'RSQ';       ArgClass: (TokenClassArray,TokenClassArray,0,0,0,0)),
{314} (Min:02; Max: 02; Name: 'STEYX';     ArgClass: (TokenClassArray,TokenClassArray,0,0,0,0)),
{315} (Min:02; Max: 02; Name: 'SLOPE';     ArgClass: (TokenClassArray,TokenClassArray,0,0,0,0)),
{316} (Min:04; Max: 04; Name: 'TTEST';     ArgClass: (TokenClassArray,TokenClassArray,TokenClassValue,TokenClassValue,0,0)),
{317} (Min:03; Max: 99; Name: 'PROB';      ArgClass: (TokenClassArray,TokenClassArray,TokenClassValue,0,0,0)),
{318} (Min:01; Max: 99; Name: 'DEVSQ';     ArgClass: (TokenClassRef,0,0,0,0,0)),
{319} (Min:01; Max: 99; Name: 'GEOMEAN';   ArgClass: (TokenClassRef,0,0,0,0,0)),
{320} (Min:01; Max: 99; Name: 'HARMEAN';   ArgClass: (TokenClassRef,0,0,0,0,0)),
{321} (Min:01; Max: 99; Name: 'SUMSQ';     ArgClass: (TokenClassRef,0,0,0,0,0)),
{322} (Min:01; Max: 99; Name: 'KURT';      ArgClass: (TokenClassRef,0,0,0,0,0)),
{323} (Min:01; Max: 99; Name: 'SKEW';      ArgClass: (TokenClassRef,0,0,0,0,0)),
{324} (Min:02; Max: 99; Name: 'ZTEST';     ArgClass: (TokenClassRef,TokenClassValue,0,0,0,0)),
{325} (Min:02; Max: 02; Name: 'LARGE';     ArgClass: (TokenClassRef,TokenClassValue,0,0,0,0)),
{326} (Min:02; Max: 02; Name: 'SMALL';     ArgClass: (TokenClassRef,TokenClassValue,0,0,0,0)),
{327} (Min:02; Max: 02; Name: 'QUARTILE';  ArgClass: (TokenClassRef,TokenClassValue,0,0,0,0)),
{328} (Min:02; Max: 02; Name: 'PERCENTILE'; ArgClass: (TokenClassRef,TokenClassValue,0,0,0,0)),
{329} (Min:02; Max: 99; Name: 'PERCENTRANK'; ArgClass: (TokenClassRef,TokenClassValue,0,0,0,0)),
{330} (Min:01; Max: 99; Name: 'MODE';      ArgClass: (TokenClassArray,0,0,0,0,0)),
{331} (Min:02; Max: 02; Name: 'TRIMMEAN';  ArgClass: (TokenClassRef,TokenClassValue,0,0,0,0)),
{332} (Min:02; Max: 02; Name: 'TINV';      ArgClass: (TokenClassValue,TokenClassValue,0,0,0,0)),
{333} (Min:255; Max: 255; Name: '333'),
{334} (Min:255; Max: 255; Name: 'MOVIE.COMMAND'),
{335} (Min:255; Max: 255; Name: 'GET.MOVIE'),
{336} (Min:01; Max: 99; Name: 'CONCATENATE'; ArgClass: (TokenClassValue,0,0,0,0,0)),
{337} (Min:02; Max: 02; Name: 'POWER';     ArgClass: (TokenClassValue,TokenClassValue,0,0,0,0)),
{338} (Min:255; Max: 255; Name: 'PIVOT.ADD.DATA'),
{339} (Min:255; Max: 255; Name: 'GET.PIVOT.TABLE'),
{340} (Min:255; Max: 255; Name: 'GET.PIVOT.FIELD'),
{341} (Min:255; Max: 255; Name: 'GET.PIVOT.ITEM'),
{342} (Min:01; Max: 01; Name: 'RADIANS';   ArgClass: (TokenClassValue,0,0,0,0,0)),
{343} (Min:01; Max: 01; Name: 'DEGREES';   ArgClass: (TokenClassValue,0,0,0,0,0)),
{344} (Min:02; Max: 99; Name: 'SUBTOTAL';  ArgClass: (TokenClassValue,TokenClassRef,0,0,0,0)),
{345} (Min:03; Max: 99; Name: 'SUMIF';     ArgClass: (TokenClassRef,TokenClassValue,TokenClassRef,0,0,0)),
{346} (Min:02; Max: 02; Name: 'COUNTIF';   ArgClass: (TokenClassRef,TokenClassValue,0,0,0,0)),
{347} (Min:01; Max: 01; Name: 'COUNTBLANK'; ArgClass: (TokenClassRef,0,0,0,0,0)),
{348} (Min:255; Max: 255; Name: 'SCENARIO.GET'),
{349} (Min:255; Max: 255; Name: 'OPTIONS.LISTS.GET'),
{350} (Min:04; Max: 04; Name: 'ISPMT';     ArgClass: (TokenClassValue,TokenClassValue,TokenClassValue,TokenClassValue,0,0)),
{351} (Min:255; Max: 255; Name: 'DATEDIF'),
{352} (Min:255; Max: 255; Name: 'DATESTRING'),
{353} (Min:255; Max: 255; Name: 'NUMBERSTRING'),
{354} (Min:02; Max: 99; Name: 'ROMAN';     ArgClass: (TokenClassValue,TokenClassValue,0,0,0,0)),
{355} (Min:255; Max: 255; Name: 'OPEN.DIALOG'),
{356} (Min:255; Max: 255; Name: 'SAVE.DIALOG'),
{357} (Min:255; Max: 255; Name: 'VIEW.GET'),
{358} (Min:255; Max: 255; Name: 'GETPIVOTDATA'),
{359} (Min:01; Max: 02; Name: 'HYPERLINK'; ArgClass: (0,0,0,0,0,0)),
{360} (Min:01; Max: 01; Name: 'PHONETIC';  ArgClass: (0,0,0,0,0,0)),
{361} (Min:01; Max: 99; Name: 'AVERAGEA';  ArgClass: (TokenClassRef,0,0,0,0,0)),
{362} (Min:01; Max: 99; Name: 'MAXA';      ArgClass: (TokenClassRef,0,0,0,0,0)),
{363} (Min:01; Max: 99; Name: 'MINA';      ArgClass: (TokenClassRef,0,0,0,0,0)),
{364} (Min:01; Max: 99; Name: 'STDEVPA';   ArgClass: (TokenClassRef,0,0,0,0,0)),
{365} (Min:01; Max: 99; Name: 'VARPA';     ArgClass: (TokenClassRef,0,0,0,0,0)),
{366} (Min:01; Max: 99; Name: 'STDEVA';    ArgClass: (TokenClassRef,0,0,0,0,0)),
{367} (Min:01; Max: 99; Name: 'VARA';      ArgClass: (TokenClassRef,0,0,0,0,0)),
{368} (Min:01; Max: 01; Name: 'BAHTTEXT';  ArgClass: (0,0,0,0,0,0)),
{369} (Min:01; Max: 01; Name: 'THAIDAYOFWEEK'; ArgClass: (0,0,0,0,0,0)),
{370} (Min:01; Max: 01; Name: 'THAIDIGIT'; ArgClass: (0,0,0,0,0,0)),
{371} (Min:01; Max: 01; Name: 'THAIMONTHOFYEAR'; ArgClass: (0,0,0,0,0,0)),
{372} (Min:01; Max: 01; Name: 'THAINUMSOUND'; ArgClass: (0,0,0,0,0,0)),
{373} (Min:01; Max: 01; Name: 'THAINUMSTRING'; ArgClass: (0,0,0,0,0,0)),
{374} (Min:01; Max: 01; Name: 'THAISTRINGLENGTH'; ArgClass: (0,0,0,0,0,0)),
{375} (Min:01; Max: 01; Name: 'ISTHAIDIGIT'; ArgClass: (0,0,0,0,0,0)),
{376} (Min:01; Max: 01; Name: 'ROUNDBAHTDOWN'; ArgClass: (0,0,0,0,0,0)),
{377} (Min:01; Max: 01; Name: 'ROUNDBAHTUP'; ArgClass: (0,0,0,0,0,0)),
{378} (Min:01; Max: 01; Name: 'THAIYEAR';  ArgClass: (0,0,0,0,0,0)),
{379} (Min:01; Max: 99; Name: 'RTD';       ArgClass: (0,0,0,0,0,0)),
{380} (Min:01; Max: 01; Name: 'ISHYPERLINK'; ArgClass: (0,0,0,0,0,0))
);


implementation

end.
