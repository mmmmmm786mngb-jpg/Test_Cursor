#!/usr/bin/env pwsh
# -*- coding: utf-8 -*-
# Скрипт для установки русского языка в EDT (Eclipse Development Tools)
# Автор: Auto
# Дата: 2025-01-27

param(
    [string]$EclipsePath = "",
    [string]$EclipseVersion = "2023-12"
)

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Установка русского языка в EDT" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Функция для поиска установки Eclipse/EDT
function Find-EclipseInstallation {
    Write-Host "Поиск установки Eclipse/EDT..." -ForegroundColor Yellow
    
    # Проверка стандартной установки 1С:EDT
    $edtPaths = @(
        "C:\Program Files\1C\1CE\components\1c-edt-start-0.9.0+229-x86_64",
        "C:\Program Files (x86)\1C\1CE\components\1c-edt-start-0.9.0+229-x86_64",
        "C:\Program Files\1C\1CE\components\*",
        "C:\Program Files (x86)\1C\1CE\components\*"
    )
    
    foreach ($pathPattern in $edtPaths) {
        $paths = Get-ChildItem -Path $pathPattern -ErrorAction SilentlyContinue -Directory
        foreach ($path in $paths) {
            if (Test-Path "$($path.FullName)\1cedtstart.exe") {
                Write-Host "Найдено 1С:EDT: $($path.FullName)\1cedtstart.exe" -ForegroundColor Green
                return $path.FullName
            }
        }
    }
    
    # Стандартные пути Eclipse
    $searchPaths = @(
        "C:\Program Files\Eclipse",
        "C:\Program Files (x86)\Eclipse",
        "$env:USERPROFILE\eclipse",
        "$env:USERPROFILE\Desktop\eclipse",
        "C:\eclipse",
        "D:\eclipse",
        "E:\eclipse"
    )
    
    foreach ($path in $searchPaths) {
        if (Test-Path "$path\eclipse.exe") {
            Write-Host "Найдено: $path\eclipse.exe" -ForegroundColor Green
            return $path
        }
    }
    
    # Поиск через реестр
    $regPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    
    foreach ($regPath in $regPaths) {
        $items = Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue | 
                 Where-Object { $_.DisplayName -like "*Eclipse*" -or $_.DisplayName -like "*EDT*" }
        
        if ($items) {
            foreach ($item in $items) {
                if ($item.InstallLocation -and (Test-Path "$($item.InstallLocation)\eclipse.exe")) {
                    Write-Host "Найдено через реестр: $($item.InstallLocation)" -ForegroundColor Green
                    return $item.InstallLocation
                }
            }
        }
    }
    
    return $null
}

# Определение версии Eclipse
function Get-EclipseVersion {
    param([string]$EclipsePath)
    
    if (-not $EclipsePath) {
        return $null
    }
    
    # Проверка наличия исполняемого файла
    $hasExecutable = (Test-Path "$EclipsePath\1cedtstart.exe") -or (Test-Path "$EclipsePath\eclipse.exe")
    if (-not $hasExecutable) {
        return $null
    }
    
    try {
        # Проверка версии через плагины (для 1С:EDT основан на Eclipse 2023-09 или 2023-12)
        $pluginsPath = "$EclipsePath\plugins"
        if (Test-Path $pluginsPath) {
            $osgiPlugin = Get-ChildItem -Path $pluginsPath -Filter "org.eclipse.osgi_*.jar" -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($osgiPlugin) {
                # Версия OSGI может указывать на версию Eclipse
                if ($osgiPlugin.Name -match "org\.eclipse\.osgi_(\d+)\.(\d+)\.(\d+)") {
                    $major = [int]$matches[1]
                    $minor = [int]$matches[2]
                    
                    # Маппинг версий OSGI на версии Eclipse
                    if ($major -eq 3 -and $minor -ge 18) {
                        return "2023-12"
                    } elseif ($major -eq 3 -and $minor -ge 17) {
                        return "2023-09"
                    } elseif ($major -eq 3 -and $minor -ge 16) {
                        return "2023-06"
                    }
                }
            }
        }
        
        # Проверка через about.ini или readme
        $aboutFile = "$EclipsePath\readme\readme_eclipse.html"
        if (Test-Path $aboutFile) {
            $content = Get-Content $aboutFile -Raw -ErrorAction SilentlyContinue
            if ($content -match "Eclipse IDE (\d{4}-\d{2})") {
                return $matches[1]
            }
        }
        
        # Проверка через about.ini
        $aboutIni = Get-ChildItem -Path $EclipsePath -Filter "about.ini" -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($aboutIni) {
            $content = Get-Content $aboutIni.FullName -ErrorAction SilentlyContinue
            if ($content -match "(\d{4}-\d{2})") {
                return $matches[1]
            }
        }
    } catch {
        Write-Host "Не удалось определить версию: $_" -ForegroundColor Yellow
    }
    
    return "2023-09" # Версия по умолчанию для 1С:EDT 0.9.0
}

# Установка русского языка через Babel
function Install-RussianLanguagePack {
    param(
        [string]$EclipsePath,
        [string]$Version
    )
    
    Write-Host ""
    Write-Host "Установка русского языкового пакета..." -ForegroundColor Yellow
    
    # Репозитории Babel для разных версий
    $babelRepos = @{
        "2023-12" = "http://download.eclipse.org/technology/babel/update-site/R0.21.0/2023-12"
        "2023-09" = "http://download.eclipse.org/technology/babel/update-site/R0.21.0/2023-09"
        "2023-06" = "http://download.eclipse.org/technology/babel/update-site/R0.20.1/2023-06"
        "2022-12" = "http://download.eclipse.org/technology/babel/update-site/R0.20.0/2022-12"
        "2022-09" = "http://download.eclipse.org/technology/babel/update-site/R0.19.0/2022-09"
    }
    
    $repoUrl = $babelRepos[$Version]
    if (-not $repoUrl) {
        Write-Host "Версия $Version не поддерживается. Используется версия 2023-12" -ForegroundColor Yellow
        $repoUrl = $babelRepos["2023-12"]
    }
    
    Write-Host "Версия Eclipse: $Version" -ForegroundColor Cyan
    Write-Host "Репозиторий Babel: $repoUrl" -ForegroundColor Cyan
    Write-Host ""
    
    # Создание команды для установки через Eclipse
    $installScript = @"
import org.eclipse.equinox.p2.director.app.DirectorApplication;
import org.eclipse.equinox.p2.core.ProvisionException;
import java.util.ArrayList;
import java.util.List;

public class InstallRussian {
    public static void main(String[] args) {
        DirectorApplication app = new DirectorApplication();
        List<String> installArgs = new ArrayList<String>();
        installArgs.add("-repository");
        installArgs.add("$repoUrl");
        installArgs.add("-installIUs");
        installArgs.add("org.eclipse.babel.runtime.feature.group");
        installArgs.add("-destination");
        installArgs.add("$EclipsePath");
        installArgs.add("-profile");
        installArgs.add("SDKProfile");
        app.run((String[])installArgs.toArray(new String[0]));
    }
}
"@
    
    Write-Host "Для установки русского языка выполните следующие шаги:" -ForegroundColor Green
    Write-Host ""
    Write-Host "1. Запустите EDT" -ForegroundColor White
    Write-Host "2. Перейдите в меню: Help → Install New Software..." -ForegroundColor White
    Write-Host "3. Нажмите кнопку 'Add...'" -ForegroundColor White
    Write-Host "4. В поле 'Name' введите: Babel" -ForegroundColor White
    Write-Host "5. В поле 'Location' введите: $repoUrl" -ForegroundColor White
    Write-Host "6. Нажмите 'Add' и дождитесь загрузки списка" -ForegroundColor White
    Write-Host "7. В дереве выберите: Babel Language Packs → Russian" -ForegroundColor White
    Write-Host "8. Нажмите 'Next' → 'Next' → примите лицензию → 'Finish'" -ForegroundColor White
    Write-Host "9. После установки перезапустите EDT" -ForegroundColor White
    Write-Host ""
    
    # Определение исполняемого файла
    $executable = if (Test-Path "$EclipsePath\1cedtstart.exe") { 
        "1cedtstart.exe" 
    } elseif (Test-Path "$EclipsePath\eclipse.exe") { 
        "eclipse.exe" 
    } else { 
        "eclipse.exe" 
    }
    
    # Альтернативный способ через командную строку
    Write-Host "Альтернативный способ (через командную строку):" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Запустите команду:" -ForegroundColor Cyan
    Write-Host "& '$EclipsePath\$executable' -application org.eclipse.equinox.p2.director -repository $repoUrl -installIUs org.eclipse.babel.runtime.feature.group -destination '$EclipsePath' -profile DefaultProfile" -ForegroundColor White
    Write-Host ""
}

# Изменение настройки локали в конфигурации
function Set-EclipseLocale {
    param([string]$EclipsePath)
    
    Write-Host ""
    Write-Host "Настройка локали в конфигурации Eclipse..." -ForegroundColor Yellow
    
    # Файлы конфигурации
    $configFile = "$EclipsePath\configuration\config.ini"
    
    # Определение INI файла (1cedtstart.ini для 1С:EDT или eclipse.ini для стандартного Eclipse)
    $iniFile = if (Test-Path "$EclipsePath\1cedtstart.ini") { 
        "$EclipsePath\1cedtstart.ini" 
    } elseif (Test-Path "$EclipsePath\eclipse.ini") { 
        "$EclipsePath\eclipse.ini" 
    } else { 
        $null 
    }
    
    # Добавление параметра локали в ini файл
    if ($iniFile -and (Test-Path $iniFile)) {
        $iniFileName = Split-Path $iniFile -Leaf
        Write-Host "Обновление $iniFileName..." -ForegroundColor Cyan
        
        $content = Get-Content $iniFile -Raw
        $newContent = $content
        
        # Проверка наличия параметра локали
        if ($content -notmatch "-Duser\.language=ru") {
            # Добавляем параметр после -vmargs или в конец файла
            if ($content -match "(-vmargs)") {
                $newContent = $content -replace "(-vmargs)", "`$1`n-Duser.language=ru`n-Duser.country=RU"
            } elseif ($content -match "(-Xmx\d+M)") {
                # Для 1cedtstart.ini добавляем перед последним параметром
                $newContent = $content -replace "(-Xmx\d+M)", "`$1`n-Duser.language=ru`n-Duser.country=RU"
            } else {
                $newContent = $content + "`n-Duser.language=ru`n-Duser.country=RU"
            }
            
            try {
                # Сохраняем в исходной кодировке (обычно это ASCII или UTF-8 без BOM)
                $encoding = [System.Text.Encoding]::UTF8
                [System.IO.File]::WriteAllText($iniFile, $newContent, $encoding)
                Write-Host "Параметры локали добавлены в $iniFileName" -ForegroundColor Green
            } catch {
                Write-Host "Не удалось обновить $iniFileName: $_" -ForegroundColor Red
                Write-Host "Добавьте вручную в $iniFileName следующие строки:" -ForegroundColor Yellow
                Write-Host "-Duser.language=ru" -ForegroundColor White
                Write-Host "-Duser.country=RU" -ForegroundColor White
            }
        } else {
            Write-Host "Параметры локали уже присутствуют в $iniFileName" -ForegroundColor Green
        }
    } else {
        Write-Host "INI файл не найден, пропускаем обновление" -ForegroundColor Yellow
    }
    
    # Обновление config.ini
    if (Test-Path $configFile) {
        Write-Host "Обновление config.ini..." -ForegroundColor Cyan
        
        $content = Get-Content $configFile -Raw
        $newContent = $content
        
        if ($content -notmatch "osgi\.nl=ru") {
            if ($content -match "(osgi\.nl=)") {
                $newContent = $content -replace "(osgi\.nl=)[^\r\n]*", "`$1ru"
            } else {
                # Добавляем в конец файла
                $newContent = $content.TrimEnd() + "`nosgi.nl=ru`n"
            }
            
            try {
                $encoding = [System.Text.Encoding]::UTF8
                [System.IO.File]::WriteAllText($configFile, $newContent, $encoding)
                Write-Host "Параметр локали добавлен в config.ini" -ForegroundColor Green
            } catch {
                Write-Host "Не удалось обновить config.ini: $_" -ForegroundColor Red
                Write-Host "Добавьте вручную в config.ini строку:" -ForegroundColor Yellow
                Write-Host "osgi.nl=ru" -ForegroundColor White
            }
        } else {
            Write-Host "Параметр локали уже присутствует в config.ini" -ForegroundColor Green
        }
    } else {
        Write-Host "config.ini не найден" -ForegroundColor Yellow
    }
}

# Основная логика
if (-not $EclipsePath) {
    $EclipsePath = Find-EclipseInstallation
}

if (-not $EclipsePath) {
    Write-Host "EDT/Eclipse не найден в системе!" -ForegroundColor Red
    Write-Host ""
    Write-Host "Укажите путь к установке EDT вручную:" -ForegroundColor Yellow
    Write-Host ".\install_russian_language_edt.ps1 -EclipsePath 'C:\путь\к\eclipse'" -ForegroundColor White
    Write-Host ""
    exit 1
}

# Проверка наличия исполняемого файла
$hasExecutable = (Test-Path "$EclipsePath\1cedtstart.exe") -or (Test-Path "$EclipsePath\eclipse.exe")
if (-not $hasExecutable) {
    Write-Host "Ошибка: исполняемый файл (1cedtstart.exe или eclipse.exe) не найден по пути $EclipsePath" -ForegroundColor Red
    exit 1
}

Write-Host "Путь к EDT: $EclipsePath" -ForegroundColor Green

# Определение версии
if (-not $EclipseVersion) {
    $EclipseVersion = Get-EclipseVersion -EclipsePath $EclipsePath
}

# Установка языкового пакета
Install-RussianLanguagePack -EclipsePath $EclipsePath -Version $EclipseVersion

# Настройка локали
Set-EclipseLocale -EclipsePath $EclipsePath

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Готово!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Следующие шаги:" -ForegroundColor Yellow
Write-Host "1. Установите языковой пакет через Help → Install New Software (см. инструкцию выше)" -ForegroundColor White
Write-Host "2. Перезапустите EDT" -ForegroundColor White
Write-Host "3. Если интерфейс не переключился, перейдите в:" -ForegroundColor White
Write-Host "   Window → Preferences → General → Appearance → Locale → Russian (ru)" -ForegroundColor White
Write-Host ""




