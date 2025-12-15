#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Скрипт для создания Word документа отчета в папке Личное
"""
import os
import sys

# Добавляем путь к библиотеке python-docx если нужно
# Но лучше использовать MCP напрямую

target_dir = r'c:\CURSOR_Projects\AChmykhalov\GitHub_Home\Test_Cursor\Личное'
target_file = os.path.join(target_dir, 'Отчет_по_целям_2025.docx')

print(f'Целевая директория: {target_dir}')
print(f'Целевой файл: {target_file}')
print(f'Директория существует: {os.path.exists(target_dir)}')



