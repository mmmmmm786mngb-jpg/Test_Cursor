#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Extract DU processing times from HTML report with smart scenario detection.

Input: Документы/ИзменениеСкоростиОбработки.htm
Output: data/du_tasks_times.csv (semicolon-separated, UTF-8)

Scenario boundaries (auto-detected by date):
  - Типовой: до 08.10.2025 включительно
  - Без дублей обменов: 09.10.2025 - 16.10.2025
  - Без дублей обменов + Параллельные портфели: с 17.10.2025
"""

from datetime import datetime
from pathlib import Path

import pandas as pd


PROJECT_ROOT = Path(__file__).resolve().parents[2]
HTML_PATH = PROJECT_ROOT / "Документы" / "ИзменениеСкоростиОбработки.htm"
CSV_OUT = PROJECT_ROOT / "data" / "du_tasks_times.csv"


def assign_scenario(date):
    """Assign scenario based on date boundaries."""
    if date <= datetime(2025, 10, 8).date():
        return "Типовой"
    elif date <= datetime(2025, 10, 16).date():
        return "Без дублей обменов"
    else:
        return "Без дублей обменов + Параллельные портфели"


def extract() -> pd.DataFrame:
    if not HTML_PATH.exists():
        raise FileNotFoundError(f"HTML file not found: {HTML_PATH}")
    
    tables = pd.read_html(HTML_PATH, encoding='utf-8')
    df = tables[3]  # Main data table
    
    # Rename columns based on structure
    df.columns = ['datetime_start', 'datetime_end', 'has_errors', 'reference', 'minutes', 'col5', 'col6']
    
    # Remove header rows and empty data
    df = df[df['minutes'].notna()].copy()
    df = df[~df['minutes'].astype(str).str.contains('Длительность', na=False)].copy()
    
    # Parse dates and minutes
    df['date_end'] = pd.to_datetime(df['datetime_end'], format='%d.%m.%Y %H:%M:%S', errors='coerce')
    df = df[df['date_end'].notna()].copy()
    df['minutes'] = pd.to_numeric(df['minutes'].astype(str).str.replace(' ', ''), errors='coerce')
    df = df[df['minutes'].notna()].copy()
    
    # Extract date (no time)
    df['date'] = df['date_end'].dt.date
    
    # Group by date to get daily totals
    daily = df.groupby('date')['minutes'].sum().reset_index()
    daily['date'] = pd.to_datetime(daily['date'])
    
    # Assign scenario labels based on date boundaries
    daily['scenario'] = daily['date'].dt.date.apply(assign_scenario)
    daily['date'] = daily['date'].dt.strftime('%Y-%m-%d')
    
    return daily[['date', 'scenario', 'minutes']]


def main() -> int:
    try:
        df = extract()
        CSV_OUT.parent.mkdir(parents=True, exist_ok=True)
        df.to_csv(CSV_OUT, sep=';', index=False, encoding='utf-8')
        
        # Print summary
        print(f"\n✅ Извлечено: {len(df)} дней")
        print(f"\nРаспределение по сценариям:")
        for scenario, count in df['scenario'].value_counts().items():
            avg = df[df['scenario'] == scenario]['minutes'].mean()
            print(f"  {scenario}: {count} дней, среднее {avg:.1f} мин/день")
        print(f"\nСохранено в: {CSV_OUT}")
        
        return 0
    except Exception as exc:
        print(f"Error: {exc}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())



