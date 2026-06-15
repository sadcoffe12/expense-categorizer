#!/usr/bin/env python3
"""
Script de prueba para validar el parsing de fechas
"""
import sys
sys.path.insert(0, '/home/aeon95/Downloads/Expense Categorizer/expense-categorizer/backend')

from app.utils.file_handler import TextUtils
import pandas as pd

# Ejemplos de fechas del usuario
test_dates = [
    "11-28-24",   # November 28, 2024 (MM-DD-YY)
    "11-30-24",   # November 30, 2024
    "12-01-24",   # December 1, 2024
    "12-02-24",   # December 2, 2024
    "06-10-26",   # Current format in logs (June 10, 2026)
    "2024-11-28", # ISO format
    "28-11-24",   # DD-MM-YY
    "28/11/2024", # DD/MM/YYYY
]

print("🧪 Prueba de Parse_Date()")
print("=" * 60)

for date_str in test_dates:
    try:
        result = TextUtils.parse_date(date_str)
        print(f"✅ '{date_str}' → {result.strftime('%Y-%m-%d')} ({result.strftime('%A')})")
    except Exception as e:
        print(f"❌ '{date_str}' → ERROR: {str(e)}")

print("\n" + "=" * 60)
print("Notas:")
print("- Los años 24 se interpretan como 2024 (2000-2030)")
print("- Los formatos MM-DD-YY son priorizados (formato americano)")
print("- Si un número > 12, se asume que es día")
