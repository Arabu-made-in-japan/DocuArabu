@echo off
@SETLOCAL enabledelayedexpansion
Set SP=%USERPROFILE%\Desktop

cd %SP%
REM ***Making directory for Backup files on desktop***
mkdir %SP%\BackUp

Set BackUp=%SP%\BackUp

REM ***Making directory ***
mkdir %BackUp%\PPT
mkdir %BackUp%\Excel
mkdir %BackUp%\Word

REM ***Making Path Variable***
Set PPT_Path=%BackUp%\PPT
Set XLSX_Path=%BackUp%\Excel
Set DOC_Path=%BackUp%\Word

REM ***Copy files as excel,word and ppts***
COPY %SP%\*.xlsx %XLSX_Path%
COPY %SP%\*.ppt %PPT_Path%
COPY %SP%\*.docx %DOC_Path%

REM ***Cleaning desktop by deleting excel, word and ppt files***
del %SP%\*.xlsx
del %SP%\*.ppt
del %SP%\*.docx

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