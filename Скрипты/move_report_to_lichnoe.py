#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Скрипт для перемещения Word документа отчета в папку Личное
"""
import os
import shutil
import glob

# Абсолютные пути
base_dir = r'c:\CURSOR_Projects\AChmykhalov\GitHub_Home\Test_Cursor'
src_file = 'Отчет_по_целям_2025.docx'
dst_dir = os.path.join(base_dir, 'Личное')
dst_file = os.path.join(dst_dir, src_file)

print(f'Ищем файл: {src_file}')
print(f'В директории: {base_dir}')

# Ищем файл в корне проекта
os.chdir(base_dir)
if os.path.exists(src_file):
    print(f'Файл найден: {os.path.abspath(src_file)}')
    # Проверяем целевую директорию
    if not os.path.exists(dst_dir):
        os.makedirs(dst_dir)
        print(f'Создана директория: {dst_dir}')
    
    # Перемещаем файл
    if os.path.exists(dst_file):
        print(f'Файл уже существует в целевой директории, перезаписываем...')
        os.remove(dst_file)
    
    shutil.move(src_file, dst_file)
    print(f'✓ Файл успешно перемещен в: {dst_file}')
    print(f'  Размер: {os.path.getsize(dst_file)} байт')
else:
    # Ищем в подпапках
    print('Файл не найден в корне, ищем в подпапках...')
    for root, dirs, files in os.walk(base_dir):
        if src_file in files:
            found_path = os.path.join(root, src_file)
            print(f'Файл найден: {found_path}')
            if not os.path.exists(dst_dir):
                os.makedirs(dst_dir)
            if os.path.exists(dst_file):
                os.remove(dst_file)
            shutil.move(found_path, dst_file)
            print(f'✓ Файл успешно перемещен в: {dst_file}')
            break
    else:
        print(f'✗ Файл {src_file} не найден в проекте')





