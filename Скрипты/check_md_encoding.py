#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Проверка и исправление кодировки Markdown файлов
"""

import sys
import os
from pathlib import Path

def check_and_fix_file(file_path):
    """Проверяет и исправляет кодировку файла"""
    try:
        # Пробуем прочитать как UTF-8
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()
        print(f"✓ Файл {file_path} успешно прочитан в UTF-8")
        print(f"  Размер: {len(content)} символов")
        
        # Перезаписываем файл с явным указанием UTF-8 без BOM
        with open(file_path, 'w', encoding='utf-8', newline='\n') as f:
            f.write(content)
        print(f"✓ Файл перезаписан в UTF-8 без BOM")
        return True
        
    except UnicodeDecodeError as e:
        print(f"✗ Ошибка декодирования: {e}")
        print("  Пробую другие кодировки...")
        
        # Пробуем CP1251 (Windows-1251)
        try:
            with open(file_path, 'r', encoding='cp1251') as f:
                content = f.read()
            print("✓ Файл прочитан в CP1251, конвертирую в UTF-8...")
            
            with open(file_path, 'w', encoding='utf-8', newline='\n') as f:
                f.write(content)
            print("✓ Файл конвертирован и сохранен в UTF-8")
            return True
        except Exception as e2:
            print(f"✗ Не удалось конвертировать: {e2}")
            return False
            
    except Exception as e:
        print(f"✗ Ошибка: {e}")
        return False

if __name__ == '__main__':
    # Определяем базовую директорию проекта
    base_dir = Path(__file__).parent.parent.resolve()
    
    # Путь к файлу
    if len(sys.argv) > 1:
        # Если путь относительный, делаем его абсолютным относительно базовой директории
        file_arg = sys.argv[1]
        if os.path.isabs(file_arg):
            file_path = Path(file_arg)
        else:
            file_path = base_dir / file_arg
    else:
        # Путь по умолчанию
        file_path = base_dir / 'Проекты' / 'тест_копия.md'
    
    # Преобразуем в строку для работы
    file_path_str = str(file_path)
    
    if not os.path.exists(file_path_str):
        print(f"✗ Файл не найден: {file_path_str}")
        print(f"  Рабочая директория: {os.getcwd()}")
        print(f"  Базовая директория: {base_dir}")
        sys.exit(1)
    
    print(f"Проверка файла: {file_path_str}")
    print("-" * 60)
    
    if check_and_fix_file(file_path_str):
        print("-" * 60)
        print("✓ Файл исправлен, можно открывать в редакторе")
        sys.exit(0)
    else:
        print("-" * 60)
        print("✗ Не удалось исправить файл")
        sys.exit(1)

