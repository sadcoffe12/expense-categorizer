"""
Mensajes de error amigables para importación de datos
"""
from typing import Optional, List
from enum import Enum

class ErrorType(str, Enum):
    """Tipos de errores que pueden ocurrir durante importación"""
    
    # Errores de archivo
    FILE_NOT_FOUND = "FILE_NOT_FOUND"
    FILE_CORRUPT = "FILE_CORRUPT"
    FILE_EMPTY = "FILE_EMPTY"
    FILE_TOO_LARGE = "FILE_TOO_LARGE"
    INVALID_FORMAT = "INVALID_FORMAT"
    
    # Errores de estructura
    MISSING_COLUMN = "MISSING_COLUMN"
    DUPLICATE_COLUMN = "DUPLICATE_COLUMN"
    NO_DATA_ROWS = "NO_DATA_ROWS"
    
    # Errores de datos
    INVALID_DATE = "INVALID_DATE"
    INVALID_NUMBER = "INVALID_NUMBER"
    INVALID_BOOLEAN = "INVALID_BOOLEAN"
    INVALID_TEXT = "INVALID_TEXT"
    ENCODING_ERROR = "ENCODING_ERROR"
    CURRENCY_SYMBOL = "CURRENCY_SYMBOL"
    MIXED_DECIMAL_FORMAT = "MIXED_DECIMAL_FORMAT"
    
    # Errores de validación
    REQUIRED_FIELD_EMPTY = "REQUIRED_FIELD_EMPTY"
    VALUE_OUT_OF_RANGE = "VALUE_OUT_OF_RANGE"
    UNKNOWN = "UNKNOWN"

class ErrorMessage:
    """Mensaje de error estructurado"""
    
    def __init__(self, error_type: ErrorType, message: str, suggestion: str = "", row: Optional[int] = None, 
                 column: Optional[str] = None, value: Optional[str] = None):
        self.error_type = error_type
        self.message = message
        self.suggestion = suggestion
        self.row = row
        self.column = column
        self.value = value
    
    def to_dict(self):
        return {
            "error_type": self.error_type.value,
            "message": self.message,
            "suggestion": self.suggestion,
            "row": self.row,
            "column": self.column,
            "value": self.value
        }
    
    def __str__(self):
        parts = [self.message]
        if self.suggestion:
            parts.append(f"💡 {self.suggestion}")
        if self.row:
            parts.insert(0, f"Fila {self.row}:")
        return "\n".join(parts)

class ErrorMessageBuilder:
    """Constructor de mensajes de error"""
    
    @staticmethod
    def file_not_found(filename: str) -> ErrorMessage:
        return ErrorMessage(
            ErrorType.FILE_NOT_FOUND,
            f"Archivo '{filename}' no encontrado",
            "Verifica que el archivo existe y está en la ubicación correcta"
        )
    
    @staticmethod
    def file_corrupt(filename: str) -> ErrorMessage:
        return ErrorMessage(
            ErrorType.FILE_CORRUPT,
            f"Archivo '{filename}' está corrupto o no es válido",
            "Intenta descargar el archivo nuevamente y verifica que sea un CSV o XLSX válido"
        )
    
    @staticmethod
    def file_empty(filename: str) -> ErrorMessage:
        return ErrorMessage(
            ErrorType.FILE_EMPTY,
            f"Archivo '{filename}' está vacío",
            "Asegúrate que el archivo contiene datos (headers + al menos 1 fila)"
        )
    
    @staticmethod
    def file_too_large(filename: str, max_size_mb: int = 50) -> ErrorMessage:
        return ErrorMessage(
            ErrorType.FILE_TOO_LARGE,
            f"Archivo '{filename}' es demasiado grande (máximo {max_size_mb}MB)",
            "Divide el archivo en partes más pequeñas"
        )
    
    @staticmethod
    def invalid_format(filename: str) -> ErrorMessage:
        return ErrorMessage(
            ErrorType.INVALID_FORMAT,
            f"Formato de archivo '{filename}' no soportado",
            "Solo se soportan archivos CSV (.csv) y Excel (.xlsx, .xls)"
        )
    
    @staticmethod
    def missing_column(column_name: str) -> ErrorMessage:
        return ErrorMessage(
            ErrorType.MISSING_COLUMN,
            f"Columna requerida '{column_name}' no encontrada",
            f"Asegúrate que tu archivo CSV/XLSX tenga una columna llamada '{column_name}'"
        )
    
    @staticmethod
    def duplicate_column(column_name: str) -> ErrorMessage:
        return ErrorMessage(
            ErrorType.DUPLICATE_COLUMN,
            f"Columna '{column_name}' aparece múltiples veces",
            f"Elimina las columnas duplicadas y deja solo una"
        )
    
    @staticmethod
    def no_data_rows() -> ErrorMessage:
        return ErrorMessage(
            ErrorType.NO_DATA_ROWS,
            "Archivo no contiene filas de datos (solo headers)",
            "Asegúrate que el archivo tiene al menos 1 fila de datos además de los headers"
        )
    
    @staticmethod
    def invalid_date(value: str, row: int, column: str, suggested_format: str = "YYYY-MM-DD") -> ErrorMessage:
        return ErrorMessage(
            ErrorType.INVALID_DATE,
            f"Fecha inválida '{value}' en fila {row}, columna '{column}'",
            f"Usa el formato {suggested_format} (ejemplo: 2024-12-25, 25-12-2024 o 12-25-24)",
            row=row,
            column=column,
            value=value
        )
    
    @staticmethod
    def date_ambiguous(value: str, row: int, column: str, possible_formats: List[str]) -> ErrorMessage:
        formats_str = " o ".join(possible_formats)
        return ErrorMessage(
            ErrorType.INVALID_DATE,
            f"Fecha ambigua '{value}' en fila {row} - no se puede determinar el formato",
            f"Puede ser {formats_str}. Por favor, usa un formato claro como YYYY-MM-DD (2024-12-25)",
            row=row,
            column=column,
            value=value
        )
    
    @staticmethod
    def invalid_number(value: str, row: int, column: str) -> ErrorMessage:
        return ErrorMessage(
            ErrorType.INVALID_NUMBER,
            f"Número inválido '{value}' en fila {row}, columna '{column}'",
            "Usa punto (.) o coma (,) como separador decimal. Ejemplo: 1500.50 o 1.500,50",
            row=row,
            column=column,
            value=value
        )
    
    @staticmethod
    def currency_symbol_detected(value: str, row: int, column: str) -> ErrorMessage:
        return ErrorMessage(
            ErrorType.CURRENCY_SYMBOL,
            f"Valor '{value}' en fila {row} contiene símbolo de moneda",
            f"Elimina símbolos ($, €, etc.). Ejemplo: cambia '${value}' a '{value.replace('$', '').replace('€', '').strip()}'",
            row=row,
            column=column,
            value=value
        )
    
    @staticmethod
    def mixed_decimal_format(row: int, column: str) -> ErrorMessage:
        return ErrorMessage(
            ErrorType.MIXED_DECIMAL_FORMAT,
            f"Columna '{column}' mezcla separadores decimales (. y ,) en fila {row}",
            "Usa consistentemente punto (.) O coma (,), pero no ambos en la misma columna",
            row=row,
            column=column
        )
    
    @staticmethod
    def required_field_empty(row: int, column: str) -> ErrorMessage:
        return ErrorMessage(
            ErrorType.REQUIRED_FIELD_EMPTY,
            f"Campo requerido '{column}' está vacío en fila {row}",
            "Asegúrate que todos los campos requeridos tengan valores",
            row=row,
            column=column
        )
    
    @staticmethod
    def encoding_error(row: int, column: str, detail: str = "") -> ErrorMessage:
        msg = f"Caracteres especiales no reconocidos en fila {row}, columna '{column}'"
        if detail:
            msg += f" ({detail})"
        return ErrorMessage(
            ErrorType.ENCODING_ERROR,
            msg,
            "Asegúrate que el archivo está guardado en formato UTF-8. En Excel: Archivo → Guardar Como → Formato CSV → Opciones → Codificación UTF-8",
            row=row,
            column=column
        )
    
    @staticmethod
    def unknown_error(detail: str = "") -> ErrorMessage:
        return ErrorMessage(
            ErrorType.UNKNOWN,
            f"Error desconocido durante la importación" + (f": {detail}" if detail else ""),
            "Intenta con un archivo más simple o contacta al soporte"
        )
    
    @staticmethod
    def success_summary(total_rows: int, imported_rows: int, skipped_rows: int) -> str:
        """Genera un resumen de importación exitosa"""
        return f"✅ Importación completada: {imported_rows}/{total_rows} filas importadas" + \
               (f" ({skipped_rows} filas omitidas)" if skipped_rows > 0 else "")
    
    @staticmethod
    def warning_summary(total_rows: int, problematic_rows: int) -> str:
        """Genera un resumen de advertencia"""
        percentage = (problematic_rows / total_rows * 100) if total_rows > 0 else 0
        return f"⚠️  {problematic_rows} de {total_rows} filas ({percentage:.1f}%) tienen problemas de formato"
