set Today=%Date: =0%
set Year=%Today:~-4%
set Month=%Today:~-7,2%
set Day=%Today:~-10,2%

set hr=%TIME: =0%
set hr=%hr:~0,2%
set min=%TIME:~3,2%

set sFolderBkp=%Year%%Month%
set sAppSource="7za.exe"
set sSecret="41L70N@@"
REM set sFileJpg="logo_tasks.jpg"

for /d %%X in (%sFolderBkp%) do (
  %sAppSource% a "%%X"_%Year%%Month%%Day%-%hr%%min%".7z" "%%X\" -p%sSecret%
  REM copy /b %sFileJpg% + %Year%%Month%%Day%-%hr%%min%_"%%X.7z" logo-%Year%%Month%%Day%-%hr%%min%_"%%X.jpg"
)

MOVE *.7z .\referencia\

explorer .\referencia\

REM RMDIR "%sFolderBkp%" /s/q