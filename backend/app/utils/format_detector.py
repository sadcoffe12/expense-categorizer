"""
Detector de formatos de datos - Identifica automáticamente tipos y formatos de columnas
"""
from typing import Dict, List, Tuple, Any, Optional
import pandas as pd
import re
from datetime import datetime

class ColumnType:
    """Tipos de datos soportados"""
    DATE = "date"
    NUMBER = "number"
    TEXT = "text"
    BOOLEAN = "boolean"
    UNKNOWN = "unknown"

class FormatInfo:
    """Información sobre formato detectado"""
    def __init__(self, format_type: str, format_string: str = "", separator: str = "", 
                 confidence: float = 0.0, is_ambiguous: bool = False, notes: str = ""):
        self.format_type = format_type
        self.format_string = format_string
        self.separator = separator
        self.confidence = confidence
        self.is_ambiguous = is_ambiguous
        self.notes = notes
    
    def to_dict(self):
        return {
            "format_type": self.format_type,
            "format_string": self.format_string,
            "separator": self.separator,
            "confidence": self.confidence,
            "is_ambiguous": self.is_ambiguous,
            "notes": self.notes
        }

class DataFormatDetector:
    """Detecta automáticamente formatos y tipos de datos en columnas"""
    
    # Patrones para detección de fechas
    DATE_PATTERNS = [
        (r'^\d{4}[-/\.]\d{2}[-/\.]\d{2}$', 'ISO 8601', '-/.'),  # 2024-12-25 o 2024/12/25 o 2024.12.25
        (r'^\d{2}[-/\.]\d{2}[-/\.]\d{4}$', 'DD-MM-YYYY o MM-DD-YYYY', '-/.'),  # 25-12-2024
        (r'^\d{2}[-/\.]\d{2}[-/\.]\d{2}$', 'YY-MM-DD o MM-DD-YY', '-/.'),  # 24-12-25
        (r'^\d{1,2}\s[A-Za-z]{3}\s\d{4}$', 'DD MMM YYYY', ' '),  # 25 Dec 2024
        (r'^\d{1,2}\s[A-Za-z]+\s\d{4}$', 'DD MMMM YYYY', ' '),  # 25 December 2024
    ]
    
    # Patrones para detección de números
    NUMBER_PATTERNS = [
        (r'^-?\d+\.?\d*$', 'integer o decimal con punto', '.'),  # 123 o 123.45
        (r'^-?\d+,?\d*$', 'integer o decimal con coma', ','),  # 123 o 123,45
        (r'^-?\$?\d+[.,]\d{2}$', 'currency format', '.,'),  # $123.45
        (r'^-?\d{1,3}[.,]\d{3}[.,]\d{2}$', 'formato europeo con miles', '.,'),  # 1.234,56
        (r'^-?\d+([.,])\d+$', 'decimal con punto o coma', '.,'),  # 123.45 o 123,45
    ]
    
    # Patrones para booleanos
    BOOLEAN_PATTERNS = [
        (r'^(true|false|si|no|yes|no|1|0)$', 'boolean', ''),
    ]
    
    @staticmethod
    def detect_column_types(df: pd.DataFrame, sample_size: int = 100) -> Dict[str, str]:
        """
        Detecta el tipo de datos de cada columna analizando una muestra.
        
        Args:
            df: DataFrame a analizar
            sample_size: Número de filas a muestrear (máximo)
            
        Returns:
            Dict con nombre_columna -> tipo_detectado
        """
        result = {}
        sample = df.head(sample_size)
        
        for col in df.columns:
            col_type = DataFormatDetector._detect_single_column_type(sample[col])
            result[col] = col_type
        
        return result
    
    @staticmethod
    def _detect_single_column_type(series: pd.Series) -> str:
        """Detecta el tipo de una columna individual"""
        
        # Filtrar valores NaN/vacíos
        non_null = series.dropna()
        non_null = non_null[non_null != '']
        
        if len(non_null) == 0:
            return ColumnType.UNKNOWN
        
        # Convertir todo a string para análisis
        values = [str(v).strip() for v in non_null.head(50)]
        
        # Contar coincidencias por tipo
        date_matches = sum(1 for v in values if DataFormatDetector._is_date(v))
        bool_matches = sum(1 for v in values if DataFormatDetector._is_boolean(v))
        number_matches = sum(1 for v in values if DataFormatDetector._is_number(v))
        
        # Determinar tipo basado en porcentaje de coincidencias
        match_threshold = len(values) * 0.7  # 70% de coincidencias
        
        if date_matches >= match_threshold:
            return ColumnType.DATE
        elif bool_matches >= match_threshold:
            return ColumnType.BOOLEAN
        elif number_matches >= match_threshold:
            return ColumnType.NUMBER
        else:
            return ColumnType.TEXT
    
    @staticmethod
    def detect_date_format(values: List[str]) -> Optional[FormatInfo]:
        """Detecta el formato de fecha de una columna"""
        
        if not values:
            return None
        
        # Filtrar valores válidos
        values = [str(v).strip() for v in values if pd.notna(v) and str(v).strip()]
        
        if not values:
            return None
        
        # Analizar los primeros valores para detectar patrón
        sample = values[:min(10, len(values))]
        
        # Detectar separador
        separators = set()
        for val in sample:
            for sep in ['-', '/', '.']:
                if sep in val:
                    separators.add(sep)
        
        separator = separators.pop() if separators else '-'
        
        # Intentar detectar formato
        for val in sample:
            parts = val.split(separator) if separator in val else None
            
            if parts and len(parts) == 3:
                try:
                    # Intentar diferentes formatos
                    p1, p2, p3 = int(parts[0]), int(parts[1]), int(parts[2])
                    
                    # Análisis de rango
                    if p1 > 31:  # p1 es año (ISO 8601: YYYY-MM-DD)
                        if p2 <= 12 and p3 <= 31:
                            return FormatInfo(ColumnType.DATE, "ISO 8601 (YYYY-MM-DD)", separator, 0.99)
                    elif p3 > 31:  # p3 es año
                        if p1 > 12:
                            return FormatInfo(ColumnType.DATE, "DD-MM-YYYY", separator, 0.9)
                        elif p2 > 12:
                            return FormatInfo(ColumnType.DATE, "MM-DD-YYYY", separator, 0.9)
                        else:
                            return FormatInfo(ColumnType.DATE, "DD-MM-YYYY o MM-DD-YYYY", separator, 0.7, True)
                    elif p1 <= 31 and p2 <= 31 and p3 <= 99:
                        # Años de 2 dígitos - ambiguo
                        return FormatInfo(ColumnType.DATE, "DD-MM-YY o MM-DD-YY", separator, 0.6, True,
                                        f"Año de 2 dígitos detectado: {p3}")
                except (ValueError, IndexError):
                    continue
        
        return None
    
    @staticmethod
    def detect_decimal_format(values: List[str]) -> Optional[FormatInfo]:
        """Detecta el separador decimal de una columna numérica"""
        
        if not values:
            return None
        
        values = [str(v).strip() for v in values if pd.notna(v) and str(v).strip()]
        
        if not values:
            return None
        
        # Contar ocurrencias de separadores
        comma_count = sum(1 for v in values if ',' in v and not v.startswith(',') and not v.endswith(','))
        dot_count = sum(1 for v in values if '.' in v and not v.startswith('.') and not v.endswith('.'))
        
        # Detectar formato
        if comma_count > dot_count * 2:
            return FormatInfo(ColumnType.NUMBER, "decimal_comma", ",", 0.95, False, "Separador decimal: coma")
        elif dot_count > comma_count * 2:
            return FormatInfo(ColumnType.NUMBER, "decimal_point", ".", 0.95, False, "Separador decimal: punto")
        elif comma_count > 0 and dot_count > 0:
            return FormatInfo(ColumnType.NUMBER, "mixed", "", 0.5, True, "Ambiguo: contiene tanto puntos como comas")
        else:
            return FormatInfo(ColumnType.NUMBER, "integer", "", 0.95, False, "Números enteros sin decimales")
    
    @staticmethod
    def _is_date(value: str) -> bool:
        """Verifica si un valor parece una fecha"""
        value = str(value).strip()
        for pattern, _, _ in DataFormatDetector.DATE_PATTERNS:
            if re.match(pattern, value):
                return True
        return False
    
    @staticmethod
    def _is_number(value: str) -> bool:
        """Verifica si un valor parece un número"""
        value = str(value).strip()
        value_clean = value.replace('$', '').replace('€', '').replace(' ', '')
        for pattern, _, _ in DataFormatDetector.NUMBER_PATTERNS:
            if re.match(pattern, value_clean):
                return True
        return False
    
    @staticmethod
    def _is_boolean(value: str) -> bool:
        """Verifica si un valor parece booleano"""
        value = str(value).strip().lower()
        for pattern, _, _ in DataFormatDetector.BOOLEAN_PATTERNS:
            if re.match(pattern, value):
                return True
        return False
    
    @staticmethod
    def suggest_transformations(df: pd.DataFrame, column_mapping: Dict[str, str]) -> Dict[str, Any]:
        """
        Sugiere cómo transformar cada columna basándose en análisis de formato.
        
        Args:
            df: DataFrame a analizar
            column_mapping: Mapeo de campos DB a columnas CSV
            
        Returns:
            Dict con sugerencias de transformación
        """
        result = {}
        
        for db_field, csv_col in column_mapping.items():
            if csv_col not in df.columns:
                continue
            
            col_data = df[csv_col].astype(str)
            col_type = DataFormatDetector._detect_single_column_type(col_data)
            
            transformation = {
                "detected_type": col_type,
                "sample_values": col_data.dropna().head(3).tolist(),
            }
            
            # Sugerencias específicas por tipo
            if col_type == ColumnType.DATE:
                date_format = DataFormatDetector.detect_date_format(col_data.tolist())
                if date_format:
                    transformation["format_info"] = date_format.to_dict()
            
            elif col_type == ColumnType.NUMBER:
                decimal_format = DataFormatDetector.detect_decimal_format(col_data.tolist())
                if decimal_format:
                    transformation["format_info"] = decimal_format.to_dict()
            
            result[db_field] = transformation
        
        return result
