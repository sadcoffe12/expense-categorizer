from typing import Dict, List, Tuple, Optional
import pandas as pd
from datetime import datetime
from .format_detector import DataFormatDetector, ColumnType
from .error_messages import ErrorMessageBuilder, ErrorType, ErrorMessage

class ValidationResult:
    """Resultado de validación con detalles completos"""
    def __init__(self, is_valid: bool, issues: List[ErrorMessage] = None, stats: Dict = None, format_hints: Dict = None):
        self.is_valid = is_valid
        self.issues = issues or []
        self.stats = stats or {}
        self.format_hints = format_hints or {}
    
    def to_dict(self):
        return {
            "is_valid": self.is_valid,
            "issues": [issue.to_dict() for issue in self.issues],
            "stats": self.stats,
            "format_hints": self.format_hints
        }

class ColumnMapper:
    """Valida y mapea columnas de archivos a schema esperado"""
    
    REQUIRED_COLUMNS = ['fecha', 'concepto', 'monto', 'categoria', 'tipo']
    OPTIONAL_COLUMNS = ['localizacion', 'notas', 'año', 'mes', 'dia']
    
    @staticmethod
    def validate_mapping(df: pd.DataFrame, mapping: Dict[str, str]) -> Tuple[bool, List[str]]:
        """Valida que las columnas mapeadas existan en el DF"""
        errors = []
        
        for db_col, csv_col in mapping.items():
            if csv_col and csv_col not in df.columns:
                errors.append(f"Columna '{csv_col}' no encontrada en archivo")
        
        if errors:
            return False, errors
        
        if len(df) == 0:
            errors.append("Archivo no contiene datos")
        
        return len(errors) == 0, errors
    
    @staticmethod
    def validate_with_diagnostics(df: pd.DataFrame, mapping: Dict[str, str]) -> ValidationResult:
        """
        Validación completa con diagnósticos detallados.
        Retorna ValidationResult con issues, stats y format hints.
        """
        from .file_handler import TextUtils
        
        issues: List[ErrorMessage] = []
        stats = {
            'total_rows': len(df),
            'valid_rows': 0,
            'invalid_rows': 0,
            'column_analysis': {}
        }
        format_hints = {}
        
        # ETAPA 1: Validar estructura de columnas
        for db_field, csv_col in mapping.items():
            if not csv_col:
                continue
            if csv_col not in df.columns:
                issues.append(ErrorMessageBuilder.missing_column(csv_col))
                continue
        
        if not issues and len(df) == 0:
            issues.append(ErrorMessageBuilder.no_data_rows())
            return ValidationResult(False, issues, stats, format_hints)
        
        # ETAPA 2: Detectar formatos de cada columna
        column_types = DataFormatDetector.detect_column_types(df)
        
        # Para fecha y monto, detectar formatos específicos
        if mapping.get('fecha') and mapping['fecha'] in df.columns:
            fecha_vals = df[mapping['fecha']].astype(str).tolist()
            date_format = DataFormatDetector.detect_date_format(fecha_vals)
            if date_format:
                format_hints['fecha'] = date_format.to_dict()
        
        if mapping.get('monto') and mapping['monto'] in df.columns:
            monto_vals = df[mapping['monto']].astype(str).tolist()
            decimal_format = DataFormatDetector.detect_decimal_format(monto_vals)
            if decimal_format:
                format_hints['monto'] = decimal_format.to_dict()
        
        # ETAPA 3: Validar muestra de datos (50 filas)
        sample_size = min(50, len(df))
        problematic_rows = []
        success_count = {'fecha': 0, 'monto': 0, 'categoria': 0}
        total_fields = {'fecha': 0, 'monto': 0, 'categoria': 0}
        
        for idx in range(sample_size):
            row = df.iloc[idx]
            row_issues = []
            
            # Validar fecha
            if mapping.get('fecha'):
                fecha_str = str(row.get(mapping['fecha'], '')).strip()
                total_fields['fecha'] += 1
                
                if pd.isna(fecha_str) or fecha_str in ['', 'nan', 'nat', 'none']:
                    row_issues.append(ErrorMessageBuilder.required_field_empty(idx + 2, mapping['fecha']))
                else:
                    try:
                        TextUtils.parse_date(fecha_str)
                        success_count['fecha'] += 1
                    except ValueError as e:
                        # Detectar si es ambiguo
                        if '–' in str(e) or 'ambiguo' in str(e).lower():
                            row_issues.append(ErrorMessageBuilder.date_ambiguous(
                                fecha_str, idx + 2, mapping['fecha'],
                                ['DD-MM-YY', 'MM-DD-YY']
                            ))
                        else:
                            row_issues.append(ErrorMessageBuilder.invalid_date(
                                fecha_str, idx + 2, mapping['fecha']
                            ))
            
            # Validar monto
            if mapping.get('monto'):
                monto_str = str(row.get(mapping['monto'], '')).strip()
                total_fields['monto'] += 1
                
                if pd.isna(monto_str) or monto_str in ['', 'nan', 'none']:
                    row_issues.append(ErrorMessageBuilder.required_field_empty(idx + 2, mapping['monto']))
                elif '$' in monto_str or '€' in monto_str or '¥' in monto_str:
                    row_issues.append(ErrorMessageBuilder.currency_symbol_detected(
                        monto_str, idx + 2, mapping['monto']
                    ))
                else:
                    try:
                        amount = TextUtils.parse_amount(monto_str)
                        if pd.notna(amount):
                            success_count['monto'] += 1
                    except:
                        row_issues.append(ErrorMessageBuilder.invalid_number(
                            monto_str, idx + 2, mapping['monto']
                        ))
            
            # Validar categoría
            if mapping.get('categoria'):
                categoria_str = str(row.get(mapping['categoria'], '')).strip()
                total_fields['categoria'] += 1
                
                if pd.isna(categoria_str) or categoria_str in ['', 'nan', 'none']:
                    row_issues.append(ErrorMessageBuilder.required_field_empty(idx + 2, mapping['categoria']))
                else:
                    success_count['categoria'] += 1
            
            if row_issues:
                problematic_rows.append({
                    'row': idx + 2,
                    'issues': row_issues
                })
                stats['invalid_rows'] += 1
            else:
                stats['valid_rows'] += 1
        
        # Calcular tasas de éxito por columna
        stats['column_analysis'] = {
            'fecha': {
                'success_rate': (success_count['fecha'] / total_fields['fecha'] * 100) if total_fields['fecha'] > 0 else 0,
                'issues_count': sum(1 for r in problematic_rows for issue in r['issues'] if 'fecha' in str(issue.column).lower())
            },
            'monto': {
                'success_rate': (success_count['monto'] / total_fields['monto'] * 100) if total_fields['monto'] > 0 else 0,
                'issues_count': sum(1 for r in problematic_rows for issue in r['issues'] if 'monto' in str(issue.column).lower())
            },
            'categoria': {
                'success_rate': (success_count['categoria'] / total_fields['categoria'] * 100) if total_fields['categoria'] > 0 else 0,
                'issues_count': sum(1 for r in problematic_rows for issue in r['issues'] if 'categoria' in str(issue.column).lower())
            }
        }
        
        # Agregar issues
        if problematic_rows:
            # Top 5 filas problemáticas
            for row_info in problematic_rows[:5]:
                issues.extend(row_info['issues'])
            
            # Si hay más, mostrar resumen
            if len(problematic_rows) > 5:
                issues.append(ErrorMessage(
                    ErrorType.UNKNOWN,
                    f"... y {len(problematic_rows) - 5} filas más tienen problemas",
                    "Revisa el CSV para asegurar que todos los datos sigan el formato correcto"
                ))
        
        is_valid = len(issues) == 0 and stats['invalid_rows'] == 0
        return ValidationResult(is_valid, issues, stats, format_hints)
    
    @staticmethod
    def transform_row(row: pd.Series, mapping: Dict[str, str]) -> Optional[Dict]:
        """Transforma fila del CSV al formato esperado por la BD con fallbacks inteligentes"""
        from .file_handler import TextUtils
        from ..logger import get_logger
        
        logger = get_logger("column_mapper")
        
        # Validar mapeo
        if not mapping:
            logger.warning(f"Mapeo vacío recibido")
            return None
        
        if mapping.get('fecha') is None:
            logger.warning(f"Campo 'fecha' no está mapeado. Mapeo: {mapping}")
            return None
        
        fecha_str = row.get(mapping.get('fecha'), '')
        concepto = row.get(mapping.get('concepto'), '')
        monto_str = row.get(mapping.get('monto'), 0)
        categoria = row.get(mapping.get('categoria'), 'Sin categoría')
        tipo = row.get(mapping.get('tipo'), 'Variable')
        localizacion = row.get(mapping.get('localizacion'), '') if mapping.get('localizacion') else ''
        notas = row.get(mapping.get('notas'), '') if mapping.get('notas') else ''
        
        # Validar que la fecha no sea NaN/NaT/None/vacía
        if pd.isna(fecha_str) or fecha_str == '' or str(fecha_str).lower() in ['nan', 'nat', 'none']:
            return None
        
        try:
            fecha = TextUtils.parse_date(str(fecha_str))
            # Validar que parse_date retornó un valor válido
            if pd.isna(fecha):
                return None
        except Exception as e:
            logger.debug(f"Error parseando fecha '{fecha_str}': {str(e)}")
            return None
        
        # Validar monto válido
        try:
            monto = TextUtils.parse_amount(str(monto_str))
        except Exception as e:
            logger.debug(f"Error parseando monto '{monto_str}': {str(e)}")
            return None
        
        if pd.isna(monto_str) or monto == 0 and str(monto_str).strip() == '':
            return None
        
        descripcion_cleaned = TextUtils.normalize_text(str(concepto), is_transaction=True)
        
        # Validar categoría
        if pd.isna(categoria) or str(categoria).lower() in ['nan', 'none', '']:
            categoria = 'Sin categoría'
        
        return {
            'date': fecha.date(),
            'description': str(concepto),
            'description_cleaned': descripcion_cleaned,
            'amount': float(monto),
            'category': str(categoria).strip(),
            'type': str(tipo).strip(),
            'location': str(localizacion).strip() if localizacion else '',
            'notes': str(notas).strip() if notas else '',
            'source': 'import'
        }
    
    @staticmethod
    def get_suggested_mapping(headers: List[str]) -> Dict[str, str]:
        """Intenta sugerir mapeo automáticamente"""
        mapping = {}
        headers_lower = [h.lower() for h in headers]
        
        fecha_keywords = ['fecha', 'date', 'fecha_transaccion', 'transaction_date']
        for keyword in fecha_keywords:
            for i, h in enumerate(headers_lower):
                if keyword in h:
                    mapping['fecha'] = headers[i]
                    break
        
        concepto_keywords = ['concepto', 'description', 'descripcion', 'descripción', 'detail']
        for keyword in concepto_keywords:
            for i, h in enumerate(headers_lower):
                if keyword in h:
                    mapping['concepto'] = headers[i]
                    break
        
        monto_keywords = ['monto', 'amount', 'valor', 'importe']
        for keyword in monto_keywords:
            for i, h in enumerate(headers_lower):
                if keyword in h:
                    mapping['monto'] = headers[i]
                    break
        
        categoria_keywords = ['categoria', 'category', 'clase', 'type_']
        for keyword in categoria_keywords:
            for i, h in enumerate(headers_lower):
                if keyword in h:
                    mapping['categoria'] = headers[i]
                    break
        
        tipo_keywords = ['tipo', 'type', 'class']
        for keyword in tipo_keywords:
            for i, h in enumerate(headers_lower):
                if keyword in h and mapping.get('categoria') != headers[i]:
                    mapping['tipo'] = headers[i]
                    break
        
        return mapping
