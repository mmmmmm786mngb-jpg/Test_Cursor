@echo off
chcp 65001 >nul
SETLOCAL ENABLEDELAYEDEXPANSION

REM ==============================================
REM Batch to register radmin.dll for 1C console (64-bit)
REM Allows user to select version from a list
REM Starts 1CV8 Servers.msc automatically
REM ==============================================

REM Check administrator rights
NET SESSION >nul 2>&1
IF %ERRORLEVEL% NEQ 0 (
    powershell -Command "Start-Process '%~f0' -Verb RunAs"
    exit /b
)

del versions.tmp >nul 2>&1
SET Count=0

REM Registry paths
SET RegPathsHKLM=HKLM\SOFTWARE\1C\1Cv8 HKLM\SOFTWARE\Wow6432Node\1C\1Cv8
SET RegPathsHKCU=HKCU\Software\1C\1Cv8

REM 1️⃣ Check HKLM
FOR %%R IN (%RegPathsHKLM%) DO (
    FOR /F "tokens=*" %%V IN ('reg query "%%R" 2^>nul') DO (
        SET Ver=%%~nxV
        SET Ver=!Ver: =!
        IF NOT "!Ver!"=="" (
            SET /A Count+=1
            echo !Count!^|!Ver!^|%%R>>versions.tmp
            echo !Count!. !Ver!
        )
    )
)

REM 2️⃣ Check HKCU
FOR %%R IN (%RegPathsHKCU%) DO (
    FOR /F "tokens=*" %%V IN ('reg query "%%R" 2^>nul') DO (
        SET Ver=%%~nxV
        SET Ver=!Ver: =!
        IF NOT "!Ver!"=="" (
            SET /A Count+=1
            echo !Count!^|!Ver!^|%%R>>versions.tmp
            echo !Count!. !Ver!
        )
    )
)

REM 3️⃣ Check Program Files directories
FOR /F "tokens=*" %%D IN ('dir "C:\Program Files\1cv8" /AD /B 2^>nul') DO (
    SET Ver=%%D
    SET FirstChar=!Ver:~0,1!
    IF "!FirstChar!" GEQ "0" IF "!FirstChar!" LEQ "9" (
        SET /A Count+=1
        echo !Count!^|!Ver!^|STANDARD_PATH>>versions.tmp
        echo !Count!. !Ver!
    )
)

IF %Count% EQU 0 (
    echo No 1C versions found!
    pause
    exit /b
)

echo.
:AskVersion
SET /P Sel="Enter number of 1C version to use: "
SET "ChosenVer="
SET "ChosenPathType="

FOR /F "tokens=1,2,3 delims=|" %%A IN (versions.tmp) DO (
    IF "%%A"=="%Sel%" (
        SET ChosenVer=%%B
        SET ChosenPathType=%%C
    )
)

IF "!ChosenVer!"=="" (
    echo Invalid selection. Try again.
    goto AskVersion
)

REM Determine installation path
IF "!ChosenPathType!"=="STANDARD_PATH" (
    SET OneCPath=C:\Program Files\1cv8\!ChosenVer!
) ELSE (
    FOR /F "tokens=2*" %%A IN ('reg query "!ChosenPathType!\!ChosenVer!" /v InstallPath 2^>nul') DO SET OneCPath=%%B
)

IF "!OneCPath!"=="" (
    SET OneCPath=C:\Program Files\1cv8\!ChosenVer!
)

echo.
echo Selected version: !ChosenVer!
echo Install path: !OneCPath!
echo.

REM Register radmin.dll
IF EXIST "!OneCPath!\bin\radmin.dll" (
    regsvr32 /s "!OneCPath!\bin\radmin.dll"
    echo radmin.dll registration completed.
) ELSE (
    echo ERROR: radmin.dll not found in !OneCPath!\bin\
)

REM Start 1C Administration Console
IF EXIST "%SystemRoot%\System32\1CV8 Servers.msc" (
    start mmc.exe "%SystemRoot%\System32\1CV8 Servers.msc"
) ELSE IF EXIST "!OneCPath!\bin\1CV8 Servers.msc" (
    start mmc.exe "!OneCPath!\bin\1CV8 Servers.msc"
) ELSE IF EXIST "!OneCPath!\1CV8 Servers.msc" (
    start mmc.exe "!OneCPath!\1CV8 Servers.msc"
) ELSE (
    echo WARNING: 1CV8 Servers.msc not found
)

del versions.tmp >nul 2>&1

echo.
echo Done. Press any key to exit...
pause >nul
