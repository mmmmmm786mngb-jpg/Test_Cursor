#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Создание Markdown файла с правильной кодировкой UTF-8 с BOM
"""

import sys
import os
from pathlib import Path

def create_md_file_with_bom(file_path, content):
    """Создает markdown файл с кодировкой UTF-8 с BOM"""
    try:
        # Добавляем BOM (Byte Order Mark) для UTF-8
        bom = '\ufeff'
        
        # Записываем файл с BOM
        with open(file_path, 'w', encoding='utf-8-sig', newline='\n') as f:
            f.write(content)
        
        print(f"✓ Файл создан: {file_path}")
        print(f"  Кодировка: UTF-8 с BOM")
        print(f"  Размер: {len(content)} символов")
        return True
        
    except Exception as e:
        print(f"✗ Ошибка создания файла: {e}")
        return False

def create_simple_test_file():
    """Создает простой тестовый файл"""
    content = """# Тестовый файл

Это простой тестовый файл Markdown для проверки открытия в редакторе.

## Заголовок 2

Обычный текст на русском языке.

### Заголовок 3

**Жирный текст**

*Курсив*

## Список

1. Пункт 1
2. Пункт 2
3. Пункт 3

## Код

```
print("Hello")
```

---

Готово.
"""
    return content

if __name__ == '__main__':
    # Определяем базовую директорию проекта
    base_dir = Path(__file__).parent.parent.resolve()
    
    # Путь к файлу
    if len(sys.argv) > 1:
        file_arg = sys.argv[1]
        if os.path.isabs(file_arg):
            file_path = Path(file_arg)
        else:
            file_path = base_dir / file_arg
    else:
        file_path = base_dir / 'Проекты' / 'test_fixed.md'
    
    # Получаем содержимое
    content = create_simple_test_file()
    
    print(f"Создание файла: {file_path}")
    print("-" * 60)
    
    if create_md_file_with_bom(file_path, content):
        print("-" * 60)
        print("✓ Файл создан успешно")
        print(f"  Попробуйте открыть: {file_path}")
        sys.exit(0)
    else:
        print("-" * 60)
        print("✗ Не удалось создать файл")
        sys.exit(1)


