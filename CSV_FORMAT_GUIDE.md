# 📋 Formato Recomendado para Importar Gastos

## ✅ Formato Correcto (ISO 8601)

El formato **RECOMENDADO y MÁS COMPATIBLE** es:

| Fecha | Concepto | Monto | Categoría | Tipo |
|-------|----------|-------|-----------|------|
| 2024-12-25 | Compra en supermercado | 150.50 | Comida | Variable |
| 2024-12-26 | Pago de internet | 45.00 | Servicios | Fijo |
| 2024-12-27 | Sueldo | 3000.00 | Sueldo | Ingreso |

### Especificaciones:

**Fecha:**
- Formato: `YYYY-MM-DD` (ISO 8601)
- Ejemplos: `2024-12-25`, `2025-06-14`
- ✅ Soportado
- ❌ Evitar: `25/12/2024`, `25-12-24`, `12/25/2024`

**Monto:**
- Formato: Números con punto o coma como separador decimal
- Ejemplos: `150.50`, `150,50`, `3000`
- ✅ Soporta ambos separadores
- ❌ Evitar: `$150.50`, `[150.50]`, `150` (sin decimales si hay centavos)

**Categoría:**
- Texto sin espacios especiales al inicio/final
- Ejemplos: `Comida`, `Servicios`, `Sueldo`, `Transporte`
- ✅ Sin valores nulos/vacíos
- ❌ Evitar: campos vacíos, "nan", "null"

**Tipo:**
- Valores permitidos: `Variable`, `Fijo`, `Ingreso`
- ✅ Uno de estos tres
- ❌ Evitar: otros valores

---

## 📊 Ejemplo de CSV Completo

```csv
Fecha,Concepto,Monto,Categoría,Tipo,Localización,Notas
2024-12-01,Café,5.50,Comida,Variable,Downtown,Mañana
2024-12-01,Internet,45.00,Servicios,Fijo,Casa,
2024-12-02,Transporte,2.50,Transporte,Variable,,Subte
2024-12-05,Alquiler,800.00,Alquiler,Fijo,Casa,Departamento
2024-12-10,Sueldo,3000.00,Sueldo,Ingreso,,
```

---

## ⚠️ Formatos Soportados (Compatibles)

El sistema automáticamente detecta y convierte estos formatos:

| Formato | Ejemplo | Prioridad |
|---------|---------|-----------|
| ISO 8601 (Recomendado) | 2024-12-25 | ⭐⭐⭐⭐⭐ |
| Americano | 12/25/2024 | ⭐⭐⭐ |
| Europeo | 25/12/2024 | ⭐⭐⭐⭐ |
| ISO Corto | 2024-12-25 | ⭐⭐⭐⭐⭐ |
| Americano Corto | 12/25/24 | ⭐⭐⭐ |
| Europeo Corto | 25/12/24 | ⭐⭐⭐ |

---

## 🔧 Cómo Preparar tu CSV

### En Excel:
1. Abre tu archivo Excel
2. Asegúrate que las columnas sean: Fecha, Concepto, Monto, Categoría, Tipo
3. Convierte las fechas al formato `YYYY-MM-DD` (Formato → Celdas → Personalizado → `YYYY-MM-DD`)
4. Guarda como CSV (Archivo → Guardar como → Formato CSV)

### En Google Sheets:
1. Selecciona la columna de fechas
2. Formato → Número → Más formatos → Formato personalizado
3. Ingresa: `YYYY-MM-DD`
4. Descarga como CSV (Archivo → Descargar → CSV)

### En LibreOffice Calc:
1. Columna Fecha → Click derecho → Formato de Celdas
2. Categoría: Fecha
3. Formato: `YYYY-MM-DD`
4. Guardar como .csv

---

## ❌ Problemas Comunes

| Problema | Causa | Solución |
|----------|-------|----------|
| "Fecha inválida" | Formato no reconocido | Usa `YYYY-MM-DD` |
| "Monto no numérico" | Caracteres especiales ($, []) | Elimina símbolos |
| "Categoría vacía" | Celdas sin contenido | Rellena todas las celdas requeridas |
| Importación vacía | Separador CSV incorrecto | Verifica que sea coma (,) o tabulación |
| "NaN" o "NULL" en datos | Valores nulos en Excel | Reemplaza con valor válido o elimina fila |

---

## 📌 Checklist Antes de Importar

- [ ] Fechas en formato `YYYY-MM-DD` (2024-12-25)
- [ ] Montos son números (150.50 o 150,50)
- [ ] Categorías no están vacías
- [ ] Tipo es uno de: Variable, Fijo, Ingreso
- [ ] No hay valores "NaN", "NULL" o campos vacíos en columnas requeridas
- [ ] Archivo guardado como CSV
- [ ] No hay caracteres especiales ilegales

---

## ✨ Ventajas del Formato Recomendado

✅ Estándar internacional ISO 8601  
✅ Compatible con todas las herramientas  
✅ No tiene ambigüedad en fechas  
✅ Excel, Google Sheets y LibreOffice lo reconocen automáticamente  
✅ Sorting y búsqueda más rápida  
✅ Menos errores de importación  

