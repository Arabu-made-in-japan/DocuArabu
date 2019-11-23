@echo off
@SETLOCAL enabledelayedexpansion
Set SP=%USERPROFILE%\Desktop

cd %SP%
REM ***Making directory for Backup files on desktop***
if exist %SP%\BackUp goto mk_dir
mkdir %SP%\BackUp

:mk_dir
Set BackUp=%SP%\BackUp
Set PATH_LOG=%BackUp%\%COMPUTERNAME%_%date%.log

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
COPY %SP%\*.xlsx %XLSX_Path% >>%PATH_LOG
COPY %SP%\*.ppt %PPT_Path% >>%PATH_LOG
COPY %SP%\*.docx %DOC_Path% >>%PATH_LOG
COPY %SP%\*.txt %Text_Path% >>%PATH_LOG

REM ***Cleaning desktop by deleting excel, word and ppt files***
del %SP%\*.xlsx >>%PATH_LOG
del %SP%\*.ppt >>%PATH_LOG
del %SP%\*.docx >>%PATH_LOG
del %SP%\*.txt >>%PATH_LOG


REM ***Count Variable ***
set n=0

REM ***Setting xlsx files to array***
for %%A in (%XLSX_Path%\*.xlsx) do (
    Set array[!n!]=%%~nA
    Set /a n=n+1
)

REM ***To count how many excel files exist in Excel directory***
REM ***Q.Why xlsx_number starts at -1?***
REM ***That's why roop process ends proper point.***
Set xlsx_number=-1
for %%f in (%XLSX_Path%\*.xlsx) do (
    if exist %%f (set /a xlsx_number=xlsx_number+1)
)
REM ***TO check what files array is contained***
for /l %%n in (0,1,%xlsx_number%) do (
    echo !array[%%n]!
)