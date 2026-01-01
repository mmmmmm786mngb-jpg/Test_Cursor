# Исправление проблемы с BSL Language Server

## Проблема

BSL Language Server постоянно падает с ошибками:
- `write EPIPE` - ошибка записи в pipe (соединение разорвано)
- `Pending response rejected since connection got disposed` - соединение закрыто
- Кракозябры в логах (строки 15, 18, 27, 61, 62) - проблема с кодировкой
- Сервер упал 5 раз за 3 минуты и больше не перезапускается

## Причина

Проблема вызвана:
1. **Отсутствием настроек кодировки UTF-8 для Java процесса** - это приводит к кракозябрам в логах
2. **Проблемами с соединением** между клиентом и сервером из-за неправильной кодировки
3. **Избыточным логированием** ("verbose"), которое может перегружать соединение

## Решение

### 1. Обновлены настройки Java для BSL Language Server

Добавлены параметры кодировки UTF-8 в `language-1c-bsl.java.options`:

```json
"language-1c-bsl.java.options": [
  "-Xmx4g",
  "-Dfile.encoding=UTF-8",
  "-Dconsole.encoding=UTF-8",
  "-Dsun.stdout.encoding=UTF-8",
  "-Dsun.stderr.encoding=UTF-8",
  "-Duser.language=ru",
  "-Duser.country=RU"
]
```

### 2. Уменьшен уровень логирования

Изменено с `"verbose"` на `"off"` для стабильности:

```json
"language-1c-bsl.trace.client": "off"
```

## КРИТИЧЕСКИ ВАЖНО: Проверка путей и очистка кэша (если существует)

Перед применением исправлений **ОБЯЗАТЕЛЬНО** проверьте пути и очистите кэш (если он существует):

1. **Закройте Cursor полностью**

2. **Запустите скрипт проверки:**
   ```powershell
   .\Скрипты\Очистить_Кэш_BSL_Language_Server.ps1
   ```
   
   Скрипт автоматически:
   - Проверит существование Java и JAR файла
   - Покажет версии компонентов
   - Очистит кэш, если он существует
   - Если кэша нет - это нормально (кэш создается только после первого использования)

3. **Альтернативный способ (вручную):**
   - Откройте проводник Windows
   - Перейдите в: `C:\Users\Acer\AppData\Roaming\Cursor\User\globalStorage\`
   - Если папка `1c-syntax.language-1c-bsl` существует - удалите её
   - Если папки нет - это нормально, значит кэш еще не создан

4. **Откройте Cursor заново**

## Применение исправлений

### Шаг 1: Скопировать настройки

Скопируйте обновленные настройки из файла:
- `Настройки_Cursor/User_Settings/00_settings_current.json`
- или `Настройки_Cursor/settings.json`

В ваш файл настроек Cursor (обычно `.vscode/settings.json` в корне проекта или глобальные настройки).

### Шаг 2: Перезапустить Cursor

1. Закройте Cursor полностью
2. Откройте Cursor заново
3. Дождитесь инициализации BSL Language Server

### Шаг 3: Проверить работу

1. Откройте файл `.bsl`
2. Проверьте, что:
   - Подсветка синтаксиса работает
   - Автодополнение работает
   - Нет ошибок в Output панели (View → Output → BSL Language Server)

### Шаг 4: Если проблемы остались

Если проблемы сохраняются после очистки кэша:

1. **Попробуйте минимальные настройки (для диагностики):**
   ```json
   "language-1c-bsl.java.options": [
     "-Xmx1g",
     "-Dfile.encoding=UTF-8"
   ]
   ```

2. **Попробуйте использовать встроенный сервер (временно для проверки):**
   ```json
   "language-1c-bsl.downloadLanguageServer": true,
   "language-1c-bsl.server.mode": "auto"
   ```
   Если это работает, значит проблема в настройках внешнего JAR.

3. **Проверьте пути к Java и JAR файлу:**
   ```json
   "language-1c-bsl.java.executablePath": "C:\\Program Files\\Eclipse Adoptium\\jdk-17.0.16.8-hotspot\\bin\\java.exe",
   "language-1c-bsl.server.jarPath": "C:\\bsl\\bsl-language-server.jar"
   ```

2. **Проверьте, что файлы существуют:**
   - Откройте PowerShell
   - Выполните:
     ```powershell
     Test-Path "C:\Program Files\Eclipse Adoptium\jdk-17.0.16.8-hotspot\bin\java.exe"
     Test-Path "C:\bsl\bsl-language-server.jar"
     ```

3. **Попробуйте уменьшить память:**
   ```json
   "language-1c-bsl.java.options": [
     "-Xmx2g",  // Вместо 4g
     "-Dfile.encoding=UTF-8",
     "-Dconsole.encoding=UTF-8",
     "-Dsun.stdout.encoding=UTF-8",
     "-Dsun.stderr.encoding=UTF-8",
     "-Duser.language=ru",
     "-Duser.country=RU"
   ]
   ```

4. **Проверьте версию расширения:**
   - Откройте Extensions (Ctrl+Shift+X)
   - Найдите "1C (BSL) Language Server"
   - Убедитесь, что установлена последняя версия
   - При необходимости переустановите расширение

## Дополнительные настройки для стабильности

Если проблемы продолжаются, можно добавить дополнительные параметры:

```json
"language-1c-bsl.java.options": [
  "-Xmx4g",
  "-Xms512m",
  "-Dfile.encoding=UTF-8",
  "-Dconsole.encoding=UTF-8",
  "-Dsun.stdout.encoding=UTF-8",
  "-Dsun.stderr.encoding=UTF-8",
  "-Duser.language=ru",
  "-Duser.country=RU",
  "-XX:+UseG1GC",
  "-XX:MaxGCPauseMillis=200"
]
```

## Решение проблемы EPIPE (Broken Pipe)

Если вы видите ошибку `write EPIPE` или `Pending response rejected since connection got disposed`, выполните следующие шаги **В УКАЗАННОМ ПОРЯДКЕ**:

### Шаг 1: Очистка кэша (ОБЯЗАТЕЛЬНО!)

1. **Закройте Cursor полностью** (File → Exit)

2. **Запустите скрипт очистки кэша:**
   ```powershell
   .\Скрипты\Очистить_Кэш_BSL_Language_Server.ps1
   ```
   
   Или вручную удалите папку:
   ```
   C:\Users\Acer\AppData\Roaming\Cursor\User\globalStorage\1c-syntax.language-1c-bsl
   ```

3. **Откройте Cursor заново**

### Шаг 2: Применение упрощенных настроек

Если после очистки кэша проблема сохраняется, используйте минимальные настройки:

1. Скопируйте настройки из `Настройки_Cursor/User_Settings/04_settings_minimal.json`
2. Вставьте в ваш `settings.json`
3. Перезапустите Cursor

### Шаг 3: Альтернатива - встроенный сервер

Если внешний JAR не работает, попробуйте встроенный сервер:

1. Скопируйте настройки из `Настройки_Cursor/User_Settings/05_settings_builtin_server.json`
2. Вставьте в ваш `settings.json`
3. Перезапустите Cursor
4. Дождитесь автоматической загрузки сервера (может занять несколько минут)

### Шаг 4: Откат на старую версию BSL Language Server (РЕКОМЕНДУЕТСЯ)

Если проблема сохраняется, попробуйте установить более старую, стабильную версию:

1. **Закройте Cursor полностью**

2. **Запустите скрипт для установки версии 0.25.0:**
   ```powershell
   .\Скрипты\Скачать_Старую_Версию_BSL_Language_Server.ps1 -Version "0.25.0"
   ```

3. **Откройте Cursor заново**

Подробная инструкция: `Документация/Откат_на_Старую_Версию_BSL_Language_Server.md`

### Шаг 5: Проверка версии расширения

Убедитесь, что у вас установлена последняя версия расширения:

1. Откройте Extensions (Ctrl+Shift+X)
2. Найдите "1C (BSL) Language Server"
3. Проверьте версию (должна быть не ниже 1.32.1)
4. Если есть обновления - обновите
5. Перезапустите Cursor

## Проверка успешности исправления

После применения исправлений:

1. ✅ Нет кракозябр в логах BSL Language Server
2. ✅ Сервер запускается без ошибок
3. ✅ Подсветка синтаксиса работает
4. ✅ Автодополнение работает
5. ✅ Диагностика кода работает

## Откат изменений

Если исправления не помогли или вызвали новые проблемы:

1. Верните уровень логирования:
   ```json
   "language-1c-bsl.trace.client": "verbose"
   ```

2. Упростите параметры Java:
   ```json
   "language-1c-bsl.java.options": [
     "-Xmx4g"
   ]
   ```

3. Перезапустите Cursor

## Контакты и поддержка

Если проблема не решается:
- Проверьте Issues на GitHub расширения: https://github.com/1c-syntax/bsl-language-server
- Проверьте документацию: https://1c-syntax.github.io/bsl-language-server/

## Дата создания

2025-01-27

## Версия исправления

1.0


