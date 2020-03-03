call "C:\Program Files (x86)\Microsoft Visual Studio\2017\Enterprise\VC\Auxiliary\Build\vcvars32.bat"

set "include=C:\Program Files (x86)\Windows Kits\10\Include\10.0.17763.0\ucrt;C:\Program Files (x86)\Windows Kits\10\Include\10.0.17763.0\shared;C:\Program Files (x86)\Windows Kits\10\Include\10.0.17763.0\um;%include%"
set "lib=C:\Program Files (x86)\Windows Kits\10\Lib\10.0.17763.0\ucrt\x86;C:\Program Files (x86)\Windows Kits\10\Lib\10.0.17763.0\um\x86;%lib%"

cl /c  /Zi /nologo /W3 /WX- /diagnostics:classic /O2 /Oy- /GL /D WIN32 /D STATIC_BUILD /D SCI_LEXER /D BOOKMARK_EDITION /D NDEBUG /D _CRT_SECURE_NO_WARNINGS /D _UNICODE /D UNICODE /Gm- /EHsc /MT /GS /arch:SSE2 /fp:precise /Zc:wchar_t /Zc:forScope /Zc:inline /Gd /analyze- /FC  NP2.cpp
pause

Link /dll NP2.obj x86\*.obj x86\Notepad2.res /SUBSYSTEM:WINDOWS /MACHINE:X86 x86\Scintilla.lib imm32.lib Shlwapi.lib Comctl32.lib Msimg32.lib kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib OpenGL32.Lib
pause


