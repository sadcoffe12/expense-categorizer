from pydantic import BaseModel
from datetime import date, datetime
from typing import Optional, List, Dict, Any

class ValidationIssue(BaseModel):
    """Detalle de un error o advertencia de validación"""
    error_type: str
    message: str
    suggestion: str = ""
    row: Optional[int] = None
    column: Optional[str] = None
    value: Optional[str] = None

class ColumnAnalysis(BaseModel):
    """Análisis de una columna"""
    success_rate: float  # 0-100
    issues_count: int

class ValidationResult(BaseModel):
    """Resultado detallado de validación de archivo"""
    is_valid: bool
    issues: List[ValidationIssue] = []
    stats: Dict[str, Any] = {}
    format_hints: Dict[str, Any] = {}

class ParseFileResponse(BaseModel):
    """Respuesta al parsear archivo"""
    headers: List[str]
    preview: List[List]
    row_count: int
    suggested_mapping: Optional[Dict[str, str]] = None
    validation_result: Optional[ValidationResult] = None  # NUEVO

class CategoryBase(BaseModel):
    name: str
    type: str
    color_hex: Optional[str] = "#666666"
    icon: Optional[str] = "💰"

class CategoryCreate(CategoryBase):
    pass

class CategoryResponse(CategoryBase):
    id: int
    created_at: datetime
    class Config:
        from_attributes = True

class ExpenseBase(BaseModel):
    date: date
    description: str
    amount: float
    category_id: int
    type: str
    location: Optional[str] = None
    notes: Optional[str] = None

class ExpenseCreate(ExpenseBase):
    pass

class ColumnMapping(BaseModel):
    fecha: str
    concepto: str
    monto: str
    categoria: str
    tipo: str
    localizacion: Optional[str] = None
    notas: Optional[str] = None

class ValidateSQLResponse(BaseModel):
    valid: bool
    table_count: int
    record_count: int
    errors: List[str]

class CreateDatabaseResponse(BaseModel):
    success: bool
    records_imported: int
    database_path: str
    errors: List[str]

class BudgetCreate(BaseModel):
    category_id: int
    amount: float
    period: str  # "month", "year", or "custom"
    start_date: Optional[date] = None
    end_date: Optional[date] = None

class BudgetResponse(BaseModel):
    id: int
    category_id: int
    category: str
    amount: float
    period: str
    start_date: Optional[str] = None
    end_date: Optional[str] = None
    created_at: str
    class Config:
        from_attributes = True
