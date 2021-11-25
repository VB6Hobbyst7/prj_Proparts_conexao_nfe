set Today=%Date: =0%
set Year=%Today:~-4%
set Month=%Today:~-7,2%
set Day=%Today:~-10,2%

set hr=%TIME: =0%
set hr=%hr:~0,2%
set min=%TIME:~3,2%

set sFolderBkp=*
set sAppSource="7za.exe"
set sSecret="41L70N#$"

for /d %%X in (%sFolderBkp%) do (
  
  rem COMPACTACAO
  %sAppSource% a "%%X"_%Year%%Month%%Day%-%hr%%min%".7z" "%%X\" -p%sSecret%
  
  rem TESTE EM ARQUIVO COMPACTADO
  %sAppSource% t "%%X"_%Year%%Month%%Day%-%hr%%min%".7z" -p%sSecret% > "%%X"_%Year%%Month%%Day%-%hr%%min%".log"
  
)

REM https://ss64.com/nt/forfiles.html