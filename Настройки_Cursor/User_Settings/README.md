# User Settings для Cursor

Эта папка содержит пользовательские настройки для Cursor.

## Файлы

### 1. settings.json

**Файл:** `01_settings.json`

Основные настройки Cursor для работы с 1С:Enterprise:

- **BSL Language Server** - настройки для работы с языком 1С (BSL)
  - Java путь: `C:\Program Files\Eclipse Adoptium\jdk-17.0.16.8-hotspot\bin\java.exe`
  - JAR путь: `C:\bsl\bsl-language-server.jar`
  - Оптимизация памяти: `-Xmx4g`
  
- **Редактор**
  - Размер табуляции: 4 пробела
  - Автосохранение: после задержки (1000ms)
  - Форматирование при сохранении: включено
  - Цветовая тема: Visual Studio Light
  - Шрифт: Fira Code
  
- **Производительность**
  - Исключение файлов: `.bak`, `.tmp`, `.git`
  - Кодировка файлов: UTF-8 с BOM
  - Подсветка строк: до 120 символов

- **Git**
  - Автополучение: включено
  - Умная фиксация: включена

### 2. settings_with_terminal_encoding.json

**Файл:** `02_settings_with_terminal_encoding.json`

Расширенные настройки с дополнительными параметрами терминала для работы с кириллицей:

**Дополнительные настройки:**

- **Терминал**
  - Профиль по умолчанию: PowerShell
  - Кодировка вывода: UTF-8
  - Переменная окружения: `PYTHONIOENCODING=utf-8`
  - Автоматическая настройка кодовой страницы: chcp 65001

- **Дополнительные аргументы терминала**

  ```powershell
  [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
  $OutputEncoding = [System.Text.Encoding]::UTF8
  $env:PYTHONIOENCODING = 'utf-8'
  chcp 65001 | Out-Null
  ```

- **Шрифт терминала**
  - Семейство: Consolas, Courier New, monospace
  - Размер: 14
  - История: 10000 строк

## Ключевые особенности

### BSL Language Server

- Запуск через JAR файл с указанной Java
- Диагностика кода на русском языке
- Строгий режим разбора запросов
- Подсветка пар скобок
- Форматирование кода с автоматическим удалением пробелов

### Редактор

- Настройки для удобной работы с кодом 1С
- Поддержка синтаксиса BSL
- Подсветка ключевых слов и операторов
- Цветовая дифференциация строк, комментариев и функций

### Работа с русским языком

- Поддержка UTF-8 кодировки
- Автоматическая настройка терминала для кириллицы
- Правильное отображение русских символов в консоли Python

## Как использовать

### Базовые настройки

Скопируйте содержимое `01_settings.json` в файл настроек Cursor:

- Windows: `%APPDATA%\Cursor\User\settings.json`
- Linux: `~/.config/Cursor/User/settings.json`
- Mac: `~/Library/Application Support/Cursor/User/settings.json`

### Расширенные настройки с поддержкой кириллицы

Используйте `02_settings_with_terminal_encoding.json` если работаете с Python скриптами, которые выводят текст на русском языке.

## Примечания

### Конфликтующие настройки

Некоторые настройки закомментированы, так как вызывали ошибку:

```
Unrecognized option: -
```

Это происходило из-за двойного вызова Java. Правильное решение - использование `language-1c-bsl.java.options` вместо `language-1c-bsl.languageServerExternalJarJavaOpts`.

### Пути для Java и BSL

Убедитесь, что пути к Java и JAR файлу существуют на вашей системе:

- Java: `C:\Program Files\Eclipse Adoptium\jdk-17.0.16.8-hotspot\bin\java.exe`
- BSL Server: `C:\bsl\bsl-language-server.jar`

При необходимости измените эти пути в настройках.

## Последнее обновление

Настройки актуальны на 25.10.2025
