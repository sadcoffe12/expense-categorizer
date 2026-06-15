import pandas as pd
import sqlite3
from pathlib import Path
from typing import Dict, List, Tuple, Optional
import re
import unicodedata

class FileHandler:
    """Maneja parseo y validación de archivos"""
    
    @staticmethod
    def parse_csv(file_path: str) -> Tuple[List[str], List[List]]:
        """Lee CSV y devuelve headers + preview - como STRINGS para evitar conversión automática"""
        df = pd.read_csv(file_path, nrows=5, dtype=str)
        headers = df.columns.tolist()
        preview = df.values.tolist()
        return headers, preview
    
    @staticmethod
    def parse_xlsx(file_path: str) -> Tuple[List[str], List[List]]:
        """Lee XLSX y devuelve headers + preview - como STRINGS para evitar conversión automática"""
        df = pd.read_excel(file_path, nrows=5, dtype=str)
        headers = df.columns.tolist()
        preview = df.values.tolist()
        return headers, preview
    
    @staticmethod
    def validate_sql(file_path: str) -> Dict:
        """Valida archivo SQL"""
        errors = []
        table_count = 0
        record_count = 0
        
        try:
            conn = sqlite3.connect(file_path)
            cursor = conn.cursor()
            
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
            tables = cursor.fetchall()
            table_count = len(tables)
            table_names = [t[0] for t in tables]
            
            if 'expenses' not in table_names:
                errors.append("Tabla 'expenses' no encontrada")
            else:
                cursor.execute("SELECT COUNT(*) FROM expenses")
                record_count = cursor.fetchone()[0]
                
                cursor.execute("PRAGMA table_info(expenses)")
                columns = [col[1] for col in cursor.fetchall()]
                required_cols = ['date', 'description', 'amount', 'category_id']
                missing = [col for col in required_cols if col not in columns]
                if missing:
                    errors.append(f"Columnas faltantes: {', '.join(missing)}")
            
            conn.close()
            return {
                'valid': len(errors) == 0,
                'table_count': table_count,
                'record_count': record_count,
                'errors': errors
            }
        
        except Exception as e:
            return {
                'valid': False,
                'table_count': 0,
                'record_count': 0,
                'errors': [f"Error al leer archivo: {str(e)}"]
            }
    
    @staticmethod
    def read_csv_full(file_path: str) -> pd.DataFrame:
        """Lee CSV completo - TODAS las columnas como strings para evitar conversión automática de fechas"""
        # Leer con dtype=str para que pandas NO convierte automaticamente fechas
        df = pd.read_csv(file_path, dtype=str)
        return df
    
    @staticmethod
    def read_xlsx_full(file_path: str) -> pd.DataFrame:
        """Lee XLSX completo - TODAS las columnas como strings para evitar conversión automática de fechas"""
        # Leer con dtype=str para que pandas NO convierte automaticamente fechas
        df = pd.read_excel(file_path, dtype=str)
        return df


class TextUtils:
    """Utilidades para procesar texto"""
    
    @staticmethod
    def normalize_text(text: str, is_transaction: bool = False) -> str:
        """Normaliza texto"""
        if not isinstance(text, str):
            return str(text) if text is not None else ""

        text = text.lower().strip()
        text = ''.join(c for c in unicodedata.normalize('NFD', text)
                      if unicodedata.category(c) != 'Mn')

        if is_transaction:
            text = re.sub(r'tarj\s?nro\.?\s?\d+', '', text)
            text = re.sub(r'\d{5,}', '', text)

        text = re.sub(r'\s+', ' ', text).strip()
        return text
    
    @staticmethod
    def parse_date(date_str: str):
        """Parsea string de fecha en múltiples formatos - MEJORADO para formatos ambiguos"""
        if pd.isna(date_str) or date_str == '' or str(date_str).lower() in ['nan', 'nat', 'none']:
            return pd.NaT
        
        date_str = str(date_str).strip()
        
        # Intentar convertir directamente con pd.to_datetime
        try:
            result = pd.to_datetime(date_str, infer_datetime_format=True)
            if not pd.isna(result):
                # IMPORTANTE: Validar que el año sea razonable (2020-2025 para gastos)
                if 2020 <= result.year <= 2030:
                    return result
        except:
            pass
        
        # Intentar formatos específicos
        # IMPORTANTE: El orden importa - probar formatos más específicos primero
        formats = [
            # Formato ISO (más específico)
            '%Y-%m-%d',    # 2024-12-12
            '%Y/%m/%d',    # 2024/12/12
            
            # Formatos con slash
            '%d/%m/%Y',    # 12/12/2024
            '%m/%d/%Y',    # 12/12/2024
            '%d/%m/%y',    # 12/12/24
            '%m/%d/%y',    # 12/12/24 (MM-DD-YY - formato americano)
            
            # Formatos con guión - IMPORTANTE: MM-DD-YY ANTES que DD-MM-YY
            '%m-%d-%Y',    # 12-12-2024 (MM-DD-YYYY)
            '%m-%d-%y',    # 11-28-24 (MM-DD-YY - formato americano, PRIORITARIO)
            '%d-%m-%Y',    # 12-12-2024 (DD-MM-YYYY)
            '%d-%m-%y',    # 12-12-24 (DD-MM-YY - formato europeo)
            
            # Formatos con punto
            '%d.%m.%Y',    # 12.12.2024
            '%d.%m.%y',    # 12.12.24
            
            # Formatos con nombre de mes
            '%d %b %Y',    # 12 Dec 2024
            '%d %B %Y',    # 12 December 2024
            '%d %b %y',    # 12 Dec 24
        ]
        
        for fmt in formats:
            try:
                result = pd.to_datetime(date_str, format=fmt)
                if not pd.isna(result):
                    # IMPORTANTE: Validar que el año sea razonable
                    if 2020 <= result.year <= 2030:
                        return result
            except:
                continue
        
        # Si nada funcionó, intentar una estrategia especial para formatos ambiguos
        # Ejemplos: "11-28-24" podría ser MM-DD-YY o DD-MM-YY
        try:
            parts = date_str.replace('/', '-').replace('.', '-').split('-')
            if len(parts) == 3:
                # Intentar identificar qué es qué basándose en ranges lógicos
                p1, p2, p3 = int(parts[0]), int(parts[1]), int(parts[2])
                
                # Si p3 <= 99, es un año de 2 dígitos
                if p3 <= 99:
                    # Convertir año de 2 dígitos a 4 dígitos
                    # Asumir años recientes: 00-30 = 2000-2030
                    year = 2000 + p3 if p3 <= 30 else 1900 + p3
                else:
                    year = p3
                
                # Ahora determinar cuál es mes y cuál es día
                # Si p1 > 12, tiene que ser día (formato DD-MM-YY)
                # Si p2 > 12, tiene que ser día (formato MM-DD-YY)
                
                if p1 > 12 and p2 <= 12:
                    # p1 es día, p2 es mes -> DD-MM-YY
                    month, day = p2, p1
                elif p2 > 12 and p1 <= 12:
                    # p1 es mes, p2 es día -> MM-DD-YY
                    month, day = p1, p2
                elif p1 <= 12 and p2 <= 12:
                    # Ambiguo - PRIORITARIO: MM-DD-YY (formato americano)
                    month, day = p1, p2
                else:
                    raise ValueError(f"No se puede determinar formato: {date_str}")
                
                # Validar rango
                if 1 <= month <= 12 and 1 <= day <= 31:
                    result = pd.Timestamp(year=year, month=month, day=day)
                    if 2020 <= result.year <= 2030:
                        return result
        except:
            pass
        
        # Si llegamos aquí, no se pudo parsear - retornar NaT en lugar de lanzar excepción
        return pd.NaT
    
    @staticmethod
    def parse_amount(amount_str: str) -> float:
        """Parsea string de monto"""
        amount_str = str(amount_str).replace('$', '').replace('[', '').replace(']', '').strip()
        try:
            return float(amount_str)
        except:
            return 0.0
