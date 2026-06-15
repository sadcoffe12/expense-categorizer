#!/usr/bin/env python3
"""
Test para verificar que el flujo de importación funciona correctamente
"""

import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'backend'))

import pandas as pd
import tempfile
from app.utils.file_handler import FileHandler, TextUtils
from app.utils.column_mapper import ColumnMapper

print("=" * 70)
print("TEST: Pipeline de Importación Completo")
print("=" * 70)

# Crear un CSV de prueba
print("\n1️⃣  CREAR CSV DE PRUEBA")
print("-" * 70)

test_data = {
    'Fecha': ['11-28-24', '11-30-24', '12-01-24', '12-02-24', '12-03-24'],
    'Concepto': ['Compra', 'Pago', 'Gasto', 'Ingreso', 'Sueldo'],
    'Monto': ['150.50', '45,00', '$100', '1000', 'ABC'],
    'Categoria': ['Comida', 'Servicios', 'Otros', 'Sueldo', 'Comida'],
    'Tipo': ['Variable', 'Fijo', 'Variable', 'Ingreso', 'Ingreso']
}

df = pd.DataFrame(test_data)
print(f"✅ CSV creado: {len(df)} filas")
print(df)

# Guardar a temporal
with tempfile.NamedTemporaryFile(delete=False, suffix='.csv', mode='w') as f:
    df.to_csv(f, index=False)
    csv_path = f.name

print(f"✅ Archivo temporal: {csv_path}")

# Test: Parse CSV
print("\n2️⃣  PARSEAR CSV")
print("-" * 70)

headers, preview = FileHandler.parse_csv(csv_path)
print(f"✅ Headers: {headers}")
print(f"✅ Preview (primeras 2 filas): {preview[:2]}")

# Test: Mapeo sugerido
print("\n3️⃣  GENERAR MAPEO SUGERIDO")
print("-" * 70)

suggested = ColumnMapper.get_suggested_mapping(headers)
print(f"✅ Mapeo sugerido: {suggested}")

# Test: Leer completo
print("\n4️⃣  LEER CSV COMPLETO")
print("-" * 70)

df_full = FileHandler.read_csv_full(csv_path)
print(f"✅ DataFrame completo: {len(df_full)} filas")
print(df_full)

# Test: Transform cada fila
print("\n5️⃣  TRANSFORMAR FILAS")
print("-" * 70)

success = 0
failed = 0
for idx, row in df_full.iterrows():
    transformed = ColumnMapper.transform_row(row, suggested)
    if transformed:
        success += 1
        print(f"✅ Fila {idx+1}: {transformed['date']} - {transformed['description']} (${transformed['amount']})")
    else:
        failed += 1
        print(f"❌ Fila {idx+1}: No se pudo transformar")

print(f"\n✅ Éxito: {success}/{len(df_full)}")
print(f"❌ Fallaron: {failed}/{len(df_full)}")

# Limpiar
os.unlink(csv_path)

print("\n" + "=" * 70)
print("✅ TEST COMPLETADO")
print("=" * 70)
