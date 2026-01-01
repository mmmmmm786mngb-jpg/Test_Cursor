# Откат на старую версию BSL Language Server

## Проблема

Текущая версия BSL Language Server (0.26.0) вызывает ошибки `EPIPE` и проблемы с соединением.

## Решение: Установка стабильной версии

Рекомендуется откатиться на более старую, стабильную версию.

### Рекомендуемые стабильные версии

- **0.25.0** - последняя версия серии 0.25.x (рекомендуется)
- **0.24.0** - стабильная версия серии 0.24.x
- **0.23.0** - более старая, но проверенная версия

## Установка старой версии

### Способ 1: Автоматический (рекомендуется)

1. **Закройте Cursor полностью**

2. **Запустите скрипт:**
   ```powershell
   .\Скрипты\Скачать_Старую_Версию_BSL_Language_Server.ps1 -Version "0.25.0"
   ```

   Скрипт автоматически:
   - Создаст резервную копию текущего файла
   - Скачает указанную версию с GitHub
   - Заменит текущий JAR файл
   - Проверит установленную версию

3. **Откройте Cursor заново**

### Способ 2: Ручной

1. **Создайте резервную копию текущего файла:**
   ```powershell
   Copy-Item "C:\bsl\bsl-language-server.jar" "C:\bsl\bsl-language-server.jar.backup"
   ```

2. **Скачайте нужную версию:**
   - Откройте браузер
   - Перейдите на: https://github.com/1c-syntax/bsl-language-server/releases
   - Найдите нужную версию (например, v0.25.0)
   - Скачайте файл `bsl-language-server-0.25.0-exec.jar`

3. **Замените файл:**
   ```powershell
   Move-Item -Path "$env:USERPROFILE\Downloads\bsl-language-server-0.25.0-exec.jar" -Destination "C:\bsl\bsl-language-server.jar" -Force
   ```

4. **Проверьте версию:**
   ```powershell
   & "C:\Program Files\Eclipse Adoptium\jdk-17.0.16.8-hotspot\bin\java.exe" -jar "C:\bsl\bsl-language-server.jar" --version
   ```

5. **Перезапустите Cursor**

## Проверка работы

После установки старой версии:

1. Откройте Cursor
2. Откройте файл `.bsl`
3. Откройте панель Output (View → Output)
4. Выберите "BSL Language Server"
5. Убедитесь, что нет ошибок `EPIPE`
6. Проверьте работу подсветки синтаксиса

## Восстановление предыдущей версии

Если новая версия не помогла, можно вернуться к предыдущей:

1. **Найдите резервную копию:**
   ```powershell
   Get-ChildItem "C:\bsl\bsl-language-server.jar.backup*" | Sort-Object LastWriteTime -Descending
   ```

2. **Восстановите файл:**
   ```powershell
   Copy-Item "C:\bsl\bsl-language-server.jar.backup.20250127_133000" "C:\bsl\bsl-language-server.jar" -Force
   ```

3. **Перезапустите Cursor**

## Список доступных версий

Все версии доступны на GitHub:
https://github.com/1c-syntax/bsl-language-server/releases

### Рекомендуемый порядок тестирования

1. **0.25.0** - попробуйте сначала эту версию
2. **0.24.0** - если 0.25.0 не работает
3. **0.23.0** - если 0.24.0 не работает
4. **0.22.0** - последний вариант

## Дополнительные настройки

После установки старой версии рекомендуется использовать минимальные настройки:

Скопируйте настройки из:
```
Настройки_Cursor/User_Settings/04_settings_minimal.json
```

В ваш `settings.json`.

## Отчет о проблемах

Если проблема сохраняется даже на старой версии:

1. Проверьте версию Java (должна быть 17 или выше)
2. Проверьте пути к Java и JAR файлу
3. Попробуйте использовать встроенный сервер (см. `05_settings_builtin_server.json`)
4. Создайте Issue на GitHub: https://github.com/1c-syntax/bsl-language-server/issues

## Дата создания

2025-01-27

## Версия инструкции

1.0






