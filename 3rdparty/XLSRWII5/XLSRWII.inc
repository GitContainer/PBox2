{$I Axcompilers.inc}

{$Q-}
{$R-}

{.$define COMPILE_LAZARUS}
{.$define BABOON}
{.$define DEMO_TIMELIMIT}

{$define XLSRWII}
{$define XLS_BIFF}
{$define XLS_XSS}
{.$define AXLINUX}

{$ifdef DELPHI_XE_OR_LATER}
  {$define XLS_CRYPTO_SUPPORT}
{$endif}

{$ifndef BABOON}
  {$define XLS_HAS_METAFILE_SUPPORT}
  {$ifdef DELPHI_2009_OR_LATER}
    {$define XLS_HAS_PNG_SUPPORT}
    {$define XLS_HAS_JPG_SUPPORT}
  {$endif}
{$endif}


