unit XSSIEDefs;

interface

const AXWComponentVersion = '2.00';

const AXW_COLOR_DEFAULT   = $F0000000;

// TODO variables
const AXW_COLOR_SELECTION = $00D7E4BC;
const AXW_COLOR_FONT      = $00000000;

const AXW_LOGPHY_REQUEST_TOPROW     = 1;
const AXW_LOGPHY_REQUEST_BOTTOMROW  = 2;
const AXW_LOGPHY_REQUEST_PAGEDOWN   = 3;
const AXW_LOGPHY_REQUEST_PAGEUP     = 4;

const AXW_CHAR_EOL                      = #0;
const AXW_CHAR_BREAKCHAR                = #32;

// Same as MS Word.
const AXW_CHAR_SPECIAL_ENDOFPARA       = #13;
const AXW_CHAR_SPECIAL_HARDLINEBREAK   = #11;
const AXW_CHAR_SPECIAL_NONREQHYPEN     = #45;
const AXW_CHAR_SPECIAL_NONBREAKHYPEN   = #30;
const AXW_CHAR_SPECIAL_NONBREAPSPACE   = #160;
const AXW_CHAR_SPECIAL_PAGEBREAK       = #12;
const AXW_CHAR_SPECIAL_COLUMNBREAK     = #14;
const AXW_CHAR_SPECIAL_TAB             = #9;
const AXW_CHAR_SPECIAL_FIELDBEGIN      = #19;
const AXW_CHAR_SPECIAL_FIELDEND        = #21;
const AXW_CHAR_SPECIAL_FIELDSEP        = #20;
const AXW_CHAR_SPECIAL_FIELDESCAPE     = '\';
const AXW_CHAR_SPECIAL_CELLMARK        = #7;

// TODO Will these chars work with all fonts?
const AXW_CHAR_REPLACE_ENDOFPARA = '¶';
const AXW_CHAR_REPLACE_SPACE     = '·';

{
Types of Non-Printing Characters
There are several different non-printing characters that you can display. Among them:
    * The paragraph symbol or pilcrow (which contains formatting codes for the preceding paragraph)
    * Tab markers, inserted when you press the Tab key (depicted as arrows)
    * Spaces, inserted when you press the space bar (depicted as dots toward the vertical center of a line)
    * Non-breaking spaces, inserted when you press Ctrl Shift space bar (depicted as a degree symbol)
    * Line breaks (depicted as a bent left-pointing arrow)
    * Page breaks (depicted as a densely dotted line with the words “Page Break” in the middle)
    * Section breaks (depicted as a two densely dotted lines with the words “Section Break,” followed by the type of break (Next Page) or (Continuous), in the middle); note that section breaks contain the instructions for the page formatting — margins, headers and footers, page orientation, etc. — of the section preceding the break
    * Column breaks (depicted as a dotted line with the words “Column Break” in the middle)
    * Hidden text (depicted as a dense line of dots — usually colored a shade of purple — immediately underneath the text)**
    * Optional or conditional (“soft”) hyphens, inserted when you press Ctrl hyphen (depicted as a hyphen with a short vertical extension at the right side)
    * Non-breaking (“hard”) hyphens, inserted when you press Ctrl Shift hyphen (which looks almost exactly like an en dash but is slightly higher up)
    * Object anchors, used to pin graphics or other “floating” items to a particular location in a document (depicted as an anchor)
    * End-of-cell and end-of row markers in tables (depicted as a circle with lines coming out of it; to me the symbol looks something like a mini-sun); these markers contain formatting codes for the individual cell and row, respectively
}
const NonPrintableCharMap: array[0..127] of AnsiChar = (
#0,#1,#2,#3,#4,#5,#6,#7,#8,#9,#10,#11,#12,AXW_CHAR_REPLACE_ENDOFPARA,#14,#15,#16,#17,#18,#19,#20,#21,
#22,#23,#24,#25,#26,#27,#28,#29,#30,#31,AXW_CHAR_REPLACE_SPACE,'!','"','#','$','%','&','''','(',
')','*','+',',','-','.','/','0','1','2','3','4','5','6','7','8','9',':',';','<',
'=','>','?','@','A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P',
'Q','R','S','T','U','V','W','X','Y','Z','[','\',']','^','_','`','a','b','c','d',
'e','f','g','h','i','j','k','l','m','n','o','p','q','r','s','t','u','v','w','x',
'y','z','{','|','}','~',#127);

implementation

end.
