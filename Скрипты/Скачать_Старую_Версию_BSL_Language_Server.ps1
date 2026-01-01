#!/usr/bin/env pwsh
# -*- coding: utf-8 -*-
<#
.SYNOPSIS
    Скачивание и установка старой версии BSL Language Server

.DESCRIPTION
    Этот скрипт скачивает указанную версию BSL Language Server с GitHub
    и заменяет текущий JAR файл. Полезно для отката на стабильную версию.

.PARAMETER Version
    Версия для скачивания (например: "0.25.0", "0.24.0"). По умолчанию: "0.25.0"

.PARAMETER JarPath
    Путь для сохранения JAR файла. По умолчанию: "C:\bsl\bsl-language-server.jar"

.EXAMPLE
    .\Скачать_Старую_Версию_BSL_Language_Server.ps1 -Version "0.25.0"
    .\Скачать_Старую_Версию_BSL_Language_Server.ps1 -Version "0.24.0"
#>

param(
    [string]$Version = "0.25.0",
    [string]$JarPath = "C:\bsl\bsl-language-server.jar"
)

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Скачивание BSL Language Server" -ForegroundColor Cyan
Write-Host "Версия: $Version" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# URL для скачивания
$DownloadUrl = "https://github.com/1c-syntax/bsl-language-server/releases/download/v$Version/bsl-language-server-$Version-exec.jar"

# Временный файл
$TempFile = "$env:TEMP\bsl-language-server-$Version-exec.jar"

Write-Host "URL для скачивания:" -ForegroundColor Yellow
Write-Host "  $DownloadUrl" -ForegroundColor Gray
Write-Host ""

# Проверка существования папки
$JarDir = Split-Path -Parent $JarPath
if (-not (Test-Path $JarDir)) {
    Write-Host "Создание папки: $JarDir" -ForegroundColor Yellow
    New-Item -ItemType Directory -Path $JarDir -Force | Out-Null
}

# Резервная копия текущего файла
if (Test-Path $JarPath) {
    $BackupPath = "$JarPath.backup.$(Get-Date -Format 'yyyyMMdd_HHmmss')"
    Write-Host "Создание резервной копии текущего файла..." -ForegroundColor Yellow
    Copy-Item -Path $JarPath -Destination $BackupPath -Force
    Write-Host "  Резервная копия: $BackupPath" -ForegroundColor Gray
    
    # Показываем версию текущего файла
    $CurrentVersion = & "C:\Program Files\Eclipse Adoptium\jdk-17.0.16.8-hotspot\bin\java.exe" -jar $JarPath --version 2>&1 | Select-Object -First 1
    Write-Host "  Текущая версия: $CurrentVersion" -ForegroundColor Gray
    Write-Host ""
}

# Скачивание
Write-Host "Скачивание версии $Version..." -ForegroundColor Yellow
try {
    $ProgressPreference = 'SilentlyContinue'
    Invoke-WebRequest -Uri $DownloadUrl -OutFile $TempFile -ErrorAction Stop
    Write-Host "Файл успешно скачан!" -ForegroundColor Green
    Write-Host ""
}
catch {
    Write-Host "ОШИБКА при скачивании: $_" -ForegroundColor Red
    Write-Host ""
    Write-Host "Возможные причины:" -ForegroundColor Yellow
    Write-Host "1. Неверная версия (проверьте доступные версии на GitHub)" -ForegroundColor White
    Write-Host "2. Проблемы с интернет-соединением" -ForegroundColor White
    Write-Host "3. Файл не существует для этой версии" -ForegroundColor White
    Write-Host ""
    Write-Host "Проверьте доступные версии:" -ForegroundColor Cyan
    Write-Host "  https://github.com/1c-syntax/bsl-language-server/releases" -ForegroundColor Gray
    exit 1
}

# Проверка скачанного файла
if (-not (Test-Path $TempFile)) {
    Write-Host "ОШИБКА: Файл не был скачан!" -ForegroundColor Red
    exit 1
}

$FileSize = (Get-Item $TempFile).Length / 1MB
Write-Host "Размер файла: $([math]::Round($FileSize, 2)) MB" -ForegroundColor Gray

# Проверка версии скачанного файла
Write-Host "Проверка версии скачанного файла..." -ForegroundColor Yellow
try {
    $DownloadedVersion = & "C:\Program Files\Eclipse Adoptium\jdk-17.0.16.8-hotspot\bin\java.exe" -jar $TempFile --version 2>&1 | Select-Object -First 1
    Write-Host "  Версия: $DownloadedVersion" -ForegroundColor Green
    Write-Host ""
}
catch {
    Write-Host "  Предупреждение: Не удалось проверить версию, но файл скачан" -ForegroundColor Yellow
    Write-Host ""
}

# Замена файла
Write-Host "Замена текущего файла..." -ForegroundColor Yellow
try {
    Move-Item -Path $TempFile -Destination $JarPath -Force -ErrorAction Stop
    Write-Host "Файл успешно установлен!" -ForegroundColor Green
    Write-Host "  Путь: $JarPath" -ForegroundColor Gray
    Write-Host ""
}
catch {
    Write-Host "ОШИБКА при замене файла: $_" -ForegroundColor Red
    Write-Host "Временный файл сохранен в: $TempFile" -ForegroundColor Yellow
    exit 1
}

# Финальная проверка
Write-Host "Финальная проверка..." -ForegroundColor Yellow
if (Test-Path $JarPath) {
    $FinalVersion = & "C:\Program Files\Eclipse Adoptium\jdk-17.0.16.8-hotspot\bin\java.exe" -jar $JarPath --version 2>&1 | Select-Object -First 1
    Write-Host "  Установленная версия: $FinalVersion" -ForegroundColor Green
    Write-Host ""
    Write-Host "Готово! Версия $Version успешно установлена." -ForegroundColor Green
    Write-Host ""
    Write-Host "Следующие шаги:" -ForegroundColor Cyan
    Write-Host "1. Закройте Cursor полностью" -ForegroundColor White
    Write-Host "2. Откройте Cursor заново" -ForegroundColor White
    Write-Host "3. Дождитесь инициализации BSL Language Server" -ForegroundColor White
}
else {
    Write-Host "ОШИБКА: Файл не найден после установки!" -ForegroundColor Red
    exit 1
}






