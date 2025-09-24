#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Скрипт для чтения и анализа Excel файла materials.xlsx
"""

import pandas as pd
import sys
import os

def read_excel_file(file_path):
    """
    Читает Excel файл и выводит информацию о его содержимом
    """
    try:
        # Проверяем существование файла
        if not os.path.exists(file_path):
            print(f"Файл {file_path} не найден!")
            return
        
        print(f"Анализ файла: {file_path}")
        print("=" * 50)
        
        # Получаем список листов
        xl_file = pd.ExcelFile(file_path)
        sheet_names = xl_file.sheet_names
        
        print(f"Количество листов: {len(sheet_names)}")
        print(f"Названия листов: {sheet_names}")
        print()
        
        # Читаем каждый лист
        for i, sheet_name in enumerate(sheet_names, 1):
            print(f"ЛИСТ {i}: {sheet_name}")
            print("-" * 40)
            
            try:
                # Читаем лист
                df = pd.read_excel(file_path, sheet_name=sheet_name)
                
                print(f"Размер данных: {df.shape[0]} строк, {df.shape[1]} столбцов")
                
                if not df.empty:
                    print("\nСтолбцы:")
                    for col in df.columns:
                        print(f"  - {col}")
                    
                    print(f"\nПервые 5 строк:")
                    print(df.head().to_string())
                    
                    # Информация о типах данных
                    print(f"\nТипы данных:")
                    for col, dtype in df.dtypes.items():
                        print(f"  {col}: {dtype}")
                    
                    # Статистика для числовых столбцов
                    numeric_cols = df.select_dtypes(include=['number']).columns
                    if len(numeric_cols) > 0:
                        print(f"\nСтатистика числовых столбцов:")
                        print(df[numeric_cols].describe().to_string())
                    
                    # Проверяем на пустые значения
                    null_counts = df.isnull().sum()
                    if null_counts.sum() > 0:
                        print(f"\nПустые значения:")
                        for col, count in null_counts[null_counts > 0].items():
                            print(f"  {col}: {count} пустых значений")
                    
                else:
                    print("Лист пустой")
                
            except Exception as e:
                print(f"Ошибка при чтении листа {sheet_name}: {e}")
            
            print("\n" + "=" * 50 + "\n")
        
        # Дополнительный анализ всех листов вместе
        print("ОБЩИЙ АНАЛИЗ")
        print("-" * 40)
        
        all_data = []
        for sheet_name in sheet_names:
            try:
                df = pd.read_excel(file_path, sheet_name=sheet_name)
                if not df.empty:
                    df['Источник_лист'] = sheet_name
                    all_data.append(df)
            except:
                continue
        
        if all_data:
            combined_df = pd.concat(all_data, ignore_index=True)
            print(f"Общее количество записей: {len(combined_df)}")
            print(f"Общее количество уникальных столбцов: {len(combined_df.columns)}")
            
            # Ищем потенциальные материалы или товары
            text_columns = combined_df.select_dtypes(include=['object']).columns
            for col in text_columns:
                if any(keyword in col.lower() for keyword in ['название', 'наименование', 'материал', 'товар', 'name']):
                    unique_values = combined_df[col].dropna().unique()
                    print(f"\nУникальные значения в столбце '{col}' ({len(unique_values)} шт.):")
                    for val in unique_values[:20]:  # Показываем первые 20
                        print(f"  - {val}")
                    if len(unique_values) > 20:
                        print(f"  ... и еще {len(unique_values) - 20} значений")
        
    except Exception as e:
        print(f"Общая ошибка при анализе файла: {e}")

if __name__ == "__main__":
    # Путь к файлу
    excel_file = "/Users/ablagovestnov/CursorProjects/ai-renovation/materials.xlsx"
    
    # Анализируем файл
    read_excel_file(excel_file)
