@echo off
setlocal EnableExtensions

REM Очистка кэша Cursor на Windows
REM 1) Закрывает процессы Cursor
REM 2) Удаляет индекс проекта .cursor (в текущем каталоге запуска)
REM 3) Чистит системные кэши в %APPDATA% и %LOCALAPPDATA%

REM --- Закрыть Cursor ---
echo [1/4] Closing Cursor processes...
for /f "tokens=2 delims=," %%P in ('tasklist /FI "IMAGENAME eq Cursor.exe" /FO CSV /NH 2^>NUL') do (
    taskkill /F /PID %%~P >NUL 2>&1
)
REM Повтор: на случай под-процессов
taskkill /IM Cursor.exe /F >NUL 2>&1

REM --- Удалить индекс проекта .cursor (если есть) ---
echo [2/4] Removing project index .\.cursor (if exists)...
if exist ".\.cursor" (
    attrib -R -H -S ".\.cursor" /S /D >NUL 2>&1
    rmdir /S /Q ".\.cursor" >NUL 2>&1
)

REM --- Пути к кэшам ---
set "ROAMING=%APPDATA%\Cursor"
set "LOCAL=%LOCALAPPDATA%\Cursor"

REM --- Удалить кэши в Roaming ---
echo [3/4] Clearing Roaming caches...
for %%D in ("Cache" "CachedData" "Code Cache" "GPUCache" "Service Worker" "User\workspaceStorage") do (
    if exist "%ROAMING%\%%~D" (
        attrib -R -H -S "%ROAMING%\%%~D" /S /D >NUL 2>&1
        rmdir /S /Q "%ROAMING%\%%~D" >NUL 2>&1
    )
)

REM --- Удалить кэши в Local ---
echo [4/4] Clearing Local caches...
for %%D in ("Cache" "CachedData" "Code Cache" "GPUCache" "Service Worker" "User\workspaceStorage") do (
    if exist "%LOCAL%\%%~D" (
        attrib -R -H -S "%LOCAL%\%%~D" /S /D >NUL 2>&1
        rmdir /S /Q "%LOCAL%\%%~D" >NUL 2>&1
    )
)

echo Done. You can restart Cursor.
endlocal
exit /b 0









