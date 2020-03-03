unit XSSIEKeys;

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

uses { Delphi } Classes, SysUtils
{$ifdef MSWINDOWS}
     ,Windows
{$endif}
;

{$ifdef MACOS}
const VK_RETURN   = 1;
const VK_LEFT     = 37;
const VK_RIGHT    = 39;
const VK_UP       = 4;
const VK_DOWN     = 5;
const VK_END      = 6;
const VK_HOME     = 7;
const VK_PRIOR    = 80;
const VK_NEXT     = 9;
const VK_F5       = 10;
const VK_BACK     = 8;
const VK_DELETE   = 12;
{$endif}

type TAXWCommandClass = (axccNone,axccControl,axccMove,axccFormat,axccClipboard,axccSelect,axccDeleteText,axccUnknown);

type TAXWCommand = (
// *** Special ***
axcNoneStart,
axcNone,axcCurrent,
axcNoneEnd,

// *** Control ***
axcControlStart,
axcControlNewPara,
axcControlLineBreak,
axcControlEnd,

// *** Move ***
axcMoveStart,
axcMoveCharLeft,axcMoveCharRight,axcMoveWordLeft,axcMoveWordRight,axcMoveParaUp,
axcMoveParaDown,axcMoveCellLeft,axcMoveCellRight,axcMoveLineUp,axcMoveLineDown,
axcMoveEndOfLIne,axcMoveBeginningOfLine,axcMoveTopOfWindow,axcMoveEndOfWindow,
axcMovePgUp,axcMovePgDown,axcMoveTopOfNextPg,axcMoveTopOfPrevPg,axcMoveEndOfDoc,
axcMoveBeginningOfDoc,axcMoveLastLocation,
axcMoveEnd,

axcMoveBeginningOfPara,
axcMoveEndOfPara,

// *** Format ***
axcFormatStart,
axcFormatCHP,axcFormatBold,axcFormatItalic,axcFormatUnderline,axcClearAllCharFmt,
axcFormatDecFont,axcFormatIncFont,
axcFormatParaLeft,axcFormatParaCenter,axcFormatParaRight,axcFormatParaJustify,
axcFormatEnd,

// *** Clipboard ***
axcClipboardStart,
axcClipboardCopy,axcClipboardCut,axcClipboardPaste,axcClipboardPastePlain,
axcClipboardEnd,

// *** Select ***
axcSelectStart,
axcSelectCharLeft,axcSelectCharRight,axcSelectWordLeft,axcSelectWordRight,axcSelectParaUp,
axcSelectParaDown,axcSelectLineUp,axcSelectLineDown,axcSelectEndOfLine,axcSelectBeginningOfLine,
axcSelectPgUp,axcSelectPgDown,axcSelectTopOfNextPg,axcSelectTopOfPrevPg,axcSelectEndOfDoc,
axcSelectBeginningOfDoc,axcSelectBookmark,

axcSelectWord,axcSelectLine,axcSelectPara,axcSelectAll,
axcSelectEnd,

// *** Delete ***
axcDeleteStart,
axcDelCharLeft,axcDelCharRight,axcDelWordLeft,axcDelWordRight,axcDelSelection {no key for this},
axcDeleteEnd
);

const AXW_VK_NOKEY = 0; // TODO Check that this is ok.

type TAXWKeyMapRec = record
     Key: word;
     Shift: TShiftState;
     Command: TAXWCommand;
     end;

const NavKeysMap: array[0..59] of TAXWKeyMapRec = (
(Key: VK_RETURN; Shift: [ssCtrl]         ; Command: axcControlNewPara),
(Key: VK_RETURN; Shift: [ssAlt]          ; Command: axcControlLineBreak),

(Key: VK_LEFT;   Shift: []               ; Command: axcMoveCharLeft),
(Key: VK_RIGHT;  Shift: []               ; Command: axcMoveCharRight),
(Key: VK_LEFT;   Shift: [ssCtrl]         ; Command: axcMoveWordLeft),
(Key: VK_RIGHT;  Shift: [ssCtrl]         ; Command: axcMoveWordRight),
(Key: VK_UP;     Shift: [ssCtrl]         ; Command: axcMoveParaUp),
(Key: VK_DOWN;   Shift: [ssCtrl]         ; Command: axcMoveParaDown),
// (Key: VK_TAB;    Shift: [ssShift]     ; Command: axcMoveCellLeft),  // Only when in a table
// (Key: VK_TAB;    Shift: []            ; Command: axcMoveCellRight), // Only when in a table
(Key: VK_UP;     Shift: []               ; Command: axcMoveLineUp),
(Key: VK_DOWN;   Shift: []               ; Command: axcMoveLineDown),
(Key: VK_END;    Shift: []               ; Command: axcMoveEndOfLine),
(Key: VK_HOME;   Shift: []               ; Command: axcMoveBeginningOfLine),
(Key: VK_PRIOR;  Shift: [ssAlt,ssCtrl]   ; Command: axcMoveTopOfWindow),
(Key: VK_NEXT;   Shift: [ssAlt,ssCtrl]   ; Command: axcMoveEndOfWindow),
(Key: VK_PRIOR;  Shift: []               ; Command: axcMovePgUp),
(Key: VK_NEXT;   Shift: []               ; Command: axcMovePgDown),
(Key: VK_PRIOR;  Shift: [ssCtrl]         ; Command: axcMoveTopOfNextPg),
(Key: VK_NEXT;   Shift: [ssCtrl]         ; Command: axcMoveTopOfPrevPg),
(Key: VK_END;    Shift: [ssCtrl]         ; Command: axcMoveEndOfDoc),
(Key: VK_HOME;   Shift: [ssCtrl]         ; Command: axcMoveBeginningOfDoc),
(Key: VK_F5;     Shift: [ssCtrl]         ; Command: axcMoveLastLocation),

(Key: Word('B'); Shift: [ssCtrl]         ; Command: axcFormatBold),
(Key: Word('I'); Shift: [ssCtrl]         ; Command: axcFormatItalic),
(Key: Word('U'); Shift: [ssCtrl]         ; Command: axcFormatUnderline),
(Key: Word('Q'); Shift: [ssCtrl]         ; Command: axcFormatDecFont),
(Key: Word('W'); Shift: [ssCtrl]         ; Command: axcFormatIncFont),
(Key: Word('0'); {AXW_VK_NOKEY;} Shift: [ssCtrl]            ; Command: axcFormatCHP),

(Key: Word('L'); Shift: [ssCtrl]         ; Command: axcFormatParaLeft),
(Key: Word('E'); Shift: [ssCtrl]         ; Command: axcFormatParaCenter),
(Key: Word('R'); Shift: [ssCtrl]         ; Command: axcFormatParaRight),
(Key: Word('J'); Shift: [ssCtrl]         ; Command: axcFormatParaJustify),

(Key: Word('C'); Shift: [ssCtrl]         ; Command: axcClipboardCopy),
(Key: Word('X'); Shift: [ssCtrl]         ; Command: axcClipboardCut),
(Key: Word('V'); Shift: [ssCtrl]         ; Command: axcClipboardPaste),

(Key: VK_LEFT;   Shift: [ssShift]        ; Command: axcSelectCharLeft),
(Key: VK_RIGHT;  Shift: [ssShift]        ; Command: axcSelectCharRight),
(Key: Word('7'); Shift: [ssCtrl]         ; Command: axcSelectWord),
(Key: Word('8'); Shift: [ssCtrl]         ; Command: axcSelectLine),
(Key: Word('9'); Shift: [ssCtrl]         ; Command: axcSelectPara),
(Key: Word('A'); Shift: [ssCtrl]         ; Command: axcSelectAll),
(Key: VK_LEFT;   Shift: [ssShift]        ; Command: axcSelectCharLeft),
(Key: VK_RIGHT;  Shift: [ssShift]        ; Command: axcSelectCharRight),
(Key: VK_LEFT;   Shift: [ssShift,ssCtrl] ; Command: axcSelectWordLeft),
(Key: VK_RIGHT;  Shift: [ssShift,ssCtrl] ; Command: axcSelectWordRight),
(Key: VK_UP;     Shift: [ssShift,ssCtrl] ; Command: axcSelectParaUp),
(Key: VK_DOWN;   Shift: [ssShift,ssCtrl] ; Command: axcSelectParaDown),
(Key: VK_UP;     Shift: [ssShift]        ; Command: axcSelectLineUp),
(Key: VK_DOWN;   Shift: [ssShift]        ; Command: axcSelectLineDown),
(Key: VK_END;    Shift: [ssShift]        ; Command: axcSelectEndOfLine),
(Key: VK_HOME;   Shift: [ssShift]        ; Command: axcSelectBeginningOfLine),
(Key: VK_PRIOR;  Shift: [ssShift]        ; Command: axcSelectPgUp),
(Key: VK_NEXT;   Shift: [ssShift]        ; Command: axcSelectPgDown),
(Key: VK_PRIOR;  Shift: [ssShift,ssCtrl] ; Command: axcSelectTopOfNextPg),
(Key: VK_NEXT;   Shift: [ssShift,ssCtrl] ; Command: axcSelectTopOfPrevPg),
(Key: VK_END;    Shift: [ssShift,ssCtrl] ; Command: axcSelectEndOfDoc),
(Key: VK_HOME;   Shift: [ssShift,ssCtrl] ; Command: axcSelectBeginningOfDoc),

(Key: VK_BACK;   Shift: []               ; Command: axcDelCharLeft),
(Key: VK_DELETE; Shift: []               ; Command: axcDelCharRight), // Deletes selection if there is one.
(Key: VK_BACK;   Shift: [ssCtrl]         ; Command: axcDelWordLeft),
(Key: VK_DELETE; Shift: [ssCtrl]         ; Command: axcDelWordRight)
);

function KeyToCommand(Key: Word; Shift: TShiftState): TAXWCommand;
function CommandClass(Command: TAXWCommand): TAXWCommandClass;

implementation

function KeyToCommand(Key: Word; Shift: TShiftState): TAXWCommand;
var
  i: integer;
begin
  for i := 0 to High(NavKeysMap) do begin
    if (Key = NavKeysMap[i].Key) and (Shift = NavKeysMap[i].Shift) then begin
      Result := NavKeysMap[i].Command;
      Exit;
    end;
  end;
  Result := axcNone;
end;

function CommandClass(Command: TAXWCommand): TAXWCommandClass;
begin
  if (Command > axcNoneStart) and (Command < axcNoneEnd) then
    Result := axccNone
  else if (Command > axcControlStart) and (Command < axcControlEnd) then
    Result := axccControl
  else if (Command > axcMoveStart) and (Command < axcMoveEnd) then
    Result := axccMove
  else if (Command > axcFormatStart) and (Command < axcFormatEnd) then
    Result := axccFormat
  else if (Command > axcClipboardStart) and (Command < axcClipboardEnd) then
    Result := axccClipboard
  else if (Command > axcSelectStart) and (Command < axcSelectEnd) then
    Result := axccSelect
  else if (Command > axcDeleteStart) and (Command < axcDeleteEnd) then
    Result := axccDeleteText
  else
    Result := axccUnknown;
end;

end.
