@echo off
chcp 65001 >nul
SETLOCAL ENABLEDELAYEDEXPANSION

REM ==============================================
REM Universal batch to register radmin.dll for 1C console (64-bit)
REM Checks HKLM, then HKCU, then standard Program Files directories
REM Filters only numeric folders in standard directories
REM Automatically elevates to admin if needed
REM ==============================================

REM Check administrator rights
NET SESSION >nul 2>&1
IF %ERRORLEVEL% NEQ 0 (
    echo Administrator rights required. Restarting with elevated privileges...
    powershell -Command "Start-Process '%~f0' -Verb RunAs"
    exit /b
)

SET MaxVersion=0
SET OneCPath=
SET CurrentRegPath=

REM Define registry paths to check
SET RegPathsHKLM=HKLM\SOFTWARE\1C\1Cv8 HKLM\SOFTWARE\Wow6432Node\1C\1Cv8
SET RegPathsHKCU=HKCU\Software\1C\1Cv8

REM 1️⃣ Check HKLM
FOR %%R IN (%RegPathsHKLM%) DO (
    FOR /F "tokens=*" %%V IN ('reg query "%%R" 2^>nul') DO (
        SET Ver=%%~nxV
        SET Ver=!Ver: =!
        IF !Ver! GTR !MaxVersion! (
            SET MaxVersion=!Ver!
            SET CurrentRegPath=%%R
        )
    )
)

REM 2️⃣ Check HKCU if nothing found
IF !MaxVersion! EQU 0 (
    FOR %%R IN (%RegPathsHKCU%) DO (
        FOR /F "tokens=*" %%V IN ('reg query "%%R" 2^>nul') DO (
            SET Ver=%%~nxV
            SET Ver=!Ver: =!
            IF !Ver! GTR !MaxVersion! (
                SET MaxVersion=!Ver!
                SET CurrentRegPath=%%R
            )
        )
    )
)

REM 3️⃣ If still not found, try standard directories
IF !MaxVersion! EQU 0 (
    FOR /F "tokens=*" %%D IN ('dir "C:\Program Files\1cv8" /AD /B 2^>nul') DO (
        SET Ver=%%D
        REM Only consider folders starting with a digit
        SET FirstChar=!Ver:~0,1!
        IF "!FirstChar!" GEQ "0" IF "!FirstChar!" LEQ "9" (
            IF !Ver! GTR !MaxVersion! (
                SET MaxVersion=!Ver!
                SET OneCPath=C:\Program Files\1cv8\%%D
            )
        )
    )
)

REM Exit if nothing found
IF !MaxVersion! EQU 0 (
    echo No installed 1C versions found in registry or standard paths!
    pause
    exit /b
)

echo Latest 1C version found: !MaxVersion!

REM Get installation path from registry if not already set
IF "!OneCPath!"=="" (
    FOR /F "tokens=2*" %%A IN ('reg query "!CurrentRegPath!\!MaxVersion!" /v InstallPath 2^>nul') DO SET OneCPath=%%B
)

REM If still empty, try default path
IF "!OneCPath!"=="" (
    SET OneCPath=C:\Program Files\1cv8\!MaxVersion!
)

echo Install path: !OneCPath!
echo.

REM Check if radmin.dll exists
IF EXIST "!OneCPath!\bin\radmin.dll" (
    echo Registering radmin.dll...
    regsvr32 /s "!OneCPath!\bin\radmin.dll"
    echo Registration completed.
) ELSE (
    echo radmin.dll not found in !OneCPath!\bin\
)

echo.
echo All steps completed. Press any key to exit...
pause >nul
