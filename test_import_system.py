#!/usr/bin/env python3
"""
Script de prueba para validar la importación mejorada
"""
import sys
from pathlib import Path

# Agregar backend al path
backend_path = Path(__file__).parent / "backend"
sys.path.insert(0, str(backend_path))

from app.utils.format_detector import DataFormatDetector
from app.utils.error_messages import ErrorMessageBuilder
from app.utils.column_mapper import ColumnMapper
import pandas as pd

print("=" * 70)
print("🧪 PRUEBA: Sistema de Importación Mejorado")
print("=" * 70)

# Test 1: Detector de Formatos
print("\n1️⃣  DETECTOR DE FORMATOS")
print("-" * 70)

test_dates = ["11-28-24", "28-11-24", "2024-11-28", "28/11/2024"]
print(f"Prueba de detección de fechas:")
for date_str in test_dates:
    result = DataFormatDetector.detect_date_format([date_str])
    if result:
        print(f"  ✅ '{date_str}' → {result.format_type} (confianza: {result.confidence})")
    else:
        print(f"  ❌ '{date_str}' → No se pudo detectar")

# Test 2: Mensajes de Error
print("\n2️⃣  MENSAJES DE ERROR")
print("-" * 70)

error1 = ErrorMessageBuilder.invalid_date("25/13/2024", 5, "fecha")
print(f"✅ {error1}")
print()

error2 = ErrorMessageBuilder.currency_symbol_detected("$1,500.50", 10, "monto")
print(f"✅ {error2}")
print()

# Test 3: Column Mapper con Diagnósticos
print("\n3️⃣  COLUMN MAPPER CON DIAGNÓSTICOS")
print("-" * 70)

# Crear DataFrame de prueba
test_data = {
    'Fecha': ['11-28-24', '11-30-24', '12-01-24', '12-02-24', 'INVALID'],
    'Concepto': ['Compra', 'Pago', 'Gasto', 'Ingreso', 'Test'],
    'Monto': ['150.50', '45,00', '$100', '1000', 'ABC'],
    'Categoria': ['Comida', 'Servicios', 'Otros', 'Sueldo', '']
}

df = pd.DataFrame(test_data)
print(f"DataFrame de prueba ({len(df)} filas):")
print(df)
print()

mapping = {
    'fecha': 'Fecha',
    'concepto': 'Concepto',
    'monto': 'Monto',
    'categoria': 'Categoria',
    'tipo': 'Tipo'
}

validation = ColumnMapper.validate_with_diagnostics(df, mapping)
print(f"Resultado de validación:")
print(f"  ✅ Válido: {validation.is_valid}")
print(f"  📊 Filas válidas: {validation.stats.get('valid_rows', 0)}/{validation.stats.get('total_rows', 0)}")
print(f"  ⚠️  Issues encontrados: {len(validation.issues)}")

if validation.issues:
    print(f"\n  Primeros 3 issues:")
    for issue in validation.issues[:3]:
        print(f"    • {issue.error_type}: {issue.message}")
        if issue.suggestion:
            print(f"      💡 {issue.suggestion}")

if validation.format_hints:
    print(f"\n  Format Hints:")
    for field, hint in validation.format_hints.items():
        print(f"    • {field}: {hint.get('format_string', 'N/A')}")

print("\n" + "=" * 70)
print("✅ Pruebas completadas")
print("=" * 70)
