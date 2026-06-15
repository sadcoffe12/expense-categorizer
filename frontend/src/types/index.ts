export interface Expense {
  id: number;
  date: string;
  description: string;
  amount: number;
  category_id: number;
  category?: string;
  type: string;
  location?: string;
  notes?: string;
  created_at: string;
}

export interface Category {
  id: number;
  name: string;
  type: string;
  color_hex: string;
  icon: string;
}

export interface Budget {
  id: number;
  name: string;
  limit: number;
  spent: number;
  period: string;
}

export interface ColumnMapping {
  fecha: string;
  concepto: string;
  monto: string;
  categoria: string;
  tipo: string;
  localizacion?: string;
  notas?: string;
}

export interface ValidationIssue {
  error_type: string;
  message: string;
  suggestion: string;
  row?: number;
  column?: string;
  value?: string;
}

export interface ColumnAnalysis {
  success_rate: number;
  issues_count: number;
}

export interface ValidationResult {
  is_valid: boolean;
  issues: ValidationIssue[];
  stats: {
    total_rows: number;
    valid_rows: number;
    invalid_rows: number;
    column_analysis: Record<string, ColumnAnalysis>;
  };
  format_hints: Record<string, any>;
}

export interface ParseFileResponse {
  headers: string[];
  preview: any[];
  row_count: number;
  suggested_mapping?: Record<string, string>;
  validation_result?: ValidationResult;
}

export interface CreateDatabaseResponse {
  success: boolean;
  records_imported: number;
  database_path: string;
  errors: string[];
}

export interface ConfigStatus {
  configured: boolean;
  database_path?: string;
  records_count: number;
  categories_count: number;
}
