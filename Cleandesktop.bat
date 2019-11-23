@echo off
@SETLOCAL enabledelayedexpansion
Set SP=%USERPROFILE%\Desktop

REM ***Making directory for Backup files on desktop***
if exist %SP%\BackUp goto mk_dir
mkdir %SP%\BackUp
REM ***Making Log Directory to store LOG FILE***
mkdir %SP%\BackUp\Log

:mk_dir
Set BackUp=%SP%\BackUp
Set BAT_LOG=%BackUp%\Log\%COMPUTERNAME%_%date:~0,4%_%date:~5,-3%_%date:~8,2%.log

if exist %BackUp%\PPT goto copy_files 
REM ***Making directory ***
mkdir %BackUp%\PPT
mkdir %BackUp%\Excel
mkdir %BackUp%\Word
mkdir %BackUp%\Text

:copy_files
REM ***Making Path Variable***
Set PPT_Path=%BackUp%\PPT
Set XLSX_Path=%BackUp%\Excel
Set DOC_Path=%BackUp%\Word
Set Text_Path=%BackUp%\Text

REM ***Copy files as excel,word and ppts***
COPY %SP%\*.xlsx %XLSX_Path% >> %BAT_LOG%
COPY %SP%\*.pptx %PPT_Path% >> %BAT_LOG%
COPY %SP%\*.docx %DOC_Path% >> %BAT_LOG%
COPY %SP%\*.txt %Text_Path% >> %BAT_LOG%

REM ***Cleaning desktop by deleting excel, word and ppt files***
del %SP%\*.xlsx >> %BAT_LOG%
del %SP%\*.pptx >> %BAT_LOG%
del %SP%\*.docx >> %BAT_LOG%
del %SP%\*.txt >> %BAT_LOG%


REM ***Count Variable ***
set n_x=0
set n_p=0
set n_d=0
set n_t=0

REM ***Setting xlsx files to array***
for %%A in (%XLSX_Path%\*.xlsx) do (
    Set array_x[!n_x!]=%%~nA
    Set /a n_x=n_x+1
)
for %%A in (%PPT_Path%\*.pptx) do (
    Set array_p[!n_p!]=%%~nA
    Set /a n_p=n_p+1
)
for %%A in (%DOC_Path%\*.docx) do (
    Set array_d[!n_d!]=%%~nA
    Set /a n_d=n_d+1
)
for %%A in (%Text_Path%\*.txt) do (
    Set array_t[!n_t!]=%%~nA
    Set /a n_t=n_t+1
)

REM ***To count how many excel files exist in Excel directory***
REM ***Q.Why xlsx_number starts at -1?***
REM ***That's why roop process ends proper point.***
Set xlsx_number=-1
Set PPT_number=-1
Set Doc_number=-1
Set TXT_number=-1
for %%f in (%XLSX_Path%\*.xlsx) do (
    if exist %%f (set /a xlsx_number=xlsx_number+1)
)
for %%f in (%PPT_Path%\*.pptx) do (
    if exist %%f (set /a PPT_number=PPT_number+1)
)
for %%f in (%DOC_Path%\*.docx) do (
    if exist %%f (set /a DOC_number=DOC_number+1)
)
for %%f in (%Text_Path%\*.txt) do (
    if exist %%f (set /a TXT_number=TXT_number+1)
)

REM ***TO check what files array is contained***
ECHO.
ECHO "Below Files are completely Backuped!"
ECHO ***Excel********************
for /l %%n in (0,1,%xlsx_number%) do (
    echo !array_x[%%n]!.xlsx
)
ECHO ***PowerPoint********************

for /l %%n in (0,1,%PPT_number%) do (
    echo !array_p[%%n]!.pptx
)
ECHO ***Word********************

for /l %%n in (0,1,%DOC_number%) do (
    echo !array_d[%%n]!.docx
)
ECHO ***Text********************

for /l %%n in (0,1,%TXT_number%) do (
    echo !array_t[%%n]!.txt
)
ECHO.