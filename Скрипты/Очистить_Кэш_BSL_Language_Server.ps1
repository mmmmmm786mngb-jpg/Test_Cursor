#!/usr/bin/env pwsh
# -*- coding: utf-8 -*-
<#
.SYNOPSIS
    Очистка кэша BSL Language Server для исправления ошибок EPIPE

.DESCRIPTION
    Этот скрипт очищает кэш расширения BSL Language Server в Cursor,
    что помогает исправить проблемы с запуском сервера (ошибки EPIPE).

.NOTES
    Автор: Auto-generated
    Дата: 2025-01-27
    Версия: 1.0
#>

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Очистка кэша BSL Language Server" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Путь к кэшу расширения
$CachePath = "$env:APPDATA\Cursor\User\globalStorage\1c-syntax.language-1c-bsl"

Write-Host "Проверка пути кэша: $CachePath" -ForegroundColor Yellow

if (Test-Path $CachePath) {
    Write-Host "Кэш найден. Удаление..." -ForegroundColor Yellow
    
    try {
        Remove-Item -Path $CachePath -Recurse -Force -ErrorAction Stop
        Write-Host "Кэш успешно удален!" -ForegroundColor Green
    }
    catch {
        Write-Host "ОШИБКА при удалении кэша: $_" -ForegroundColor Red
        Write-Host ""
        Write-Host "Попробуйте удалить вручную:" -ForegroundColor Yellow
        Write-Host "  $CachePath" -ForegroundColor White
    }
}
else {
    Write-Host "Кэш не найден (это нормально - кэш еще не создан)" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "Проверка путей к Java и JAR файлу..." -ForegroundColor Cyan

# Проверка Java
$JavaPath = "C:\Program Files\Eclipse Adoptium\jdk-17.0.16.8-hotspot\bin\java.exe"
if (Test-Path $JavaPath) {
    Write-Host "  Java: OK - $JavaPath" -ForegroundColor Green
    $JavaVersion = & $JavaPath -version 2>&1 | Select-Object -First 1
    Write-Host "    Версия: $JavaVersion" -ForegroundColor Gray
}
else {
    Write-Host "  Java: НЕ НАЙДЕН - $JavaPath" -ForegroundColor Red
}

# Проверка JAR
$JarPath = "C:\bsl\bsl-language-server.jar"
if (Test-Path $JarPath) {
    Write-Host "  JAR: OK - $JarPath" -ForegroundColor Green
    $JarVersion = & $JavaPath -jar $JarPath --version 2>&1 | Select-Object -First 1
    Write-Host "    Версия: $JarVersion" -ForegroundColor Gray
}
else {
    Write-Host "  JAR: НЕ НАЙДЕН - $JarPath" -ForegroundColor Red
}

Write-Host ""
Write-Host "Следующие шаги:" -ForegroundColor Cyan
Write-Host "1. Убедитесь, что Cursor закрыт полностью" -ForegroundColor White
Write-Host "2. Примените упрощенные настройки из файла:" -ForegroundColor White
Write-Host "   Настройки_Cursor/User_Settings/04_settings_minimal.json" -ForegroundColor Yellow
Write-Host "3. Откройте Cursor заново" -ForegroundColor White
Write-Host "4. Дождитесь инициализации BSL Language Server (10-30 секунд)" -ForegroundColor White

Write-Host ""
Write-Host "Готово!" -ForegroundColor Green

