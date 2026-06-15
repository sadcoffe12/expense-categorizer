import React from 'react';
import {
  Alert,
  AlertTitle,
  Box,
  Card,
  CardContent,
  Chip,
  LinearProgress,
  Paper,
  Table,
  TableBody,
  TableCell,
  TableContainer,
  TableHead,
  TableRow,
  Typography,
  Collapse,
  IconButton,
} from '@mui/material';
import {
  ExpandMore as ExpandMoreIcon,
  ExpandLess as ExpandLessIcon,
  CheckCircle as CheckCircleIcon,
  Warning as WarningIcon,
  Error as ErrorIcon,
} from '@mui/icons-material';

interface ValidationIssue {
  error_type: string;
  message: string;
  suggestion: string;
  row?: number;
  column?: string;
  value?: string;
}

interface ColumnAnalysis {
  success_rate: number;
  issues_count: number;
}

interface ValidationResult {
  is_valid: boolean;
  issues: ValidationIssue[];
  stats: {
    total_rows?: number;
    valid_rows?: number;
    invalid_rows?: number;
    column_analysis?: {
      [key: string]: ColumnAnalysis;
    };
  };
  format_hints: {
    [key: string]: {
      format_type?: string;
      format_string?: string;
      separator?: string;
      confidence?: number;
      is_ambiguous?: boolean;
      notes?: string;
    };
  };
}

interface ValidationResultsProps {
  validation: ValidationResult | null | undefined;
  isLoading?: boolean;
  onRetry?: () => void;
}

export const ValidationResults: React.FC<ValidationResultsProps> = ({
  validation,
  isLoading = false,
  onRetry,
}) => {
  const [expandedRows, setExpandedRows] = React.useState<Set<number>>(new Set());

  if (!validation) {
    return null;
  }

  const toggleRowExpand = (row: number) => {
    const newExpanded = new Set(expandedRows);
    if (newExpanded.has(row)) {
      newExpanded.delete(row);
    } else {
      newExpanded.add(row);
    }
    setExpandedRows(newExpanded);
  };

  const getSeverityIcon = (errorType: string) => {
    if (errorType.includes('ERROR') || errorType.includes('INVALID')) {
      return <ErrorIcon sx={{ color: 'error.main' }} />;
    }
    return <WarningIcon sx={{ color: 'warning.main' }} />;
  };

  const getSeverityColor = (errorType: string): 'error' | 'warning' | 'info' => {
    if (errorType.includes('ERROR') || errorType.includes('INVALID')) {
      return 'error';
    }
    return 'warning';
  };

  return (
    <Box sx={{ mt: 3, mb: 3 }}>
      {/* Estado General */}
      <Card sx={{ mb: 2 }}>
        <CardContent>
          <Box sx={{ display: 'flex', alignItems: 'center', mb: 2 }}>
            {validation.is_valid ? (
              <CheckCircleIcon sx={{ color: 'success.main', mr: 1, fontSize: 32 }} />
            ) : (
              <WarningIcon sx={{ color: 'warning.main', mr: 1, fontSize: 32 }} />
            )}
            <Typography variant="h6">
              {validation.is_valid
                ? '✅ Archivo validado correctamente'
                : '⚠️ Se encontraron problemas en el archivo'}
            </Typography>
          </Box>

          {/* Estadísticas */}
          {validation.stats && (
            <Box sx={{ mb: 2 }}>
              <Typography variant="subtitle2" sx={{ mb: 1 }}>
                Estadísticas:
              </Typography>
              <Box sx={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(150px, 1fr))', gap: 2 }}>
                {validation.stats.total_rows !== undefined && (
                  <Paper sx={{ p: 1.5, textAlign: 'center' }}>
                    <Typography variant="body2" color="textSecondary">
                      Total de filas
                    </Typography>
                    <Typography variant="h6">{validation.stats.total_rows}</Typography>
                  </Paper>
                )}
                {validation.stats.valid_rows !== undefined && (
                  <Paper sx={{ p: 1.5, textAlign: 'center', bgcolor: 'success.light' }}>
                    <Typography variant="body2" color="textSecondary">
                      Filas válidas
                    </Typography>
                    <Typography variant="h6" sx={{ color: 'success.dark' }}>
                      {validation.stats.valid_rows}
                    </Typography>
                  </Paper>
                )}
                {validation.stats.invalid_rows !== undefined && (
                  <Paper sx={{ p: 1.5, textAlign: 'center', bgcolor: 'warning.light' }}>
                    <Typography variant="body2" color="textSecondary">
                      Filas con problemas
                    </Typography>
                    <Typography variant="h6" sx={{ color: 'warning.dark' }}>
                      {validation.stats.invalid_rows}
                    </Typography>
                  </Paper>
                )}
              </Box>
            </Box>
          )}

          {/* Análisis por Columna */}
          {validation.stats.column_analysis && (
            <Box sx={{ mb: 2 }}>
              <Typography variant="subtitle2" sx={{ mb: 1 }}>
                Análisis por Columna:
              </Typography>
              <Box sx={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))', gap: 2 }}>
                {Object.entries(validation.stats.column_analysis).map(([colName, analysis]) => {
                  const successRate = analysis.success_rate || 0;
                  const color = successRate >= 90 ? 'success' : successRate >= 70 ? 'warning' : 'error';

                  return (
                    <Paper key={colName} sx={{ p: 2 }}>
                      <Typography variant="subtitle2" sx={{ mb: 1 }}>
                        {colName}
                      </Typography>
                      <Box sx={{ mb: 1 }}>
                        <LinearProgress
                          variant="determinate"
                          value={successRate}
                          color={color}
                          sx={{ mb: 0.5 }}
                        />
                        <Typography variant="caption">{successRate.toFixed(1)}% éxito</Typography>
                      </Box>
                      {analysis.issues_count > 0 && (
                        <Chip
                          label={`${analysis.issues_count} problemas`}
                          size="small"
                          color={color}
                          variant="outlined"
                        />
                      )}
                    </Paper>
                  );
                })}
              </Box>
            </Box>
          )}

          {/* Format Hints */}
          {Object.keys(validation.format_hints).length > 0 && (
            <Box sx={{ mb: 2 }}>
              <Typography variant="subtitle2" sx={{ mb: 1 }}>
                💡 Formatos Detectados:
              </Typography>
              <Box sx={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))', gap: 1 }}>
                {Object.entries(validation.format_hints).map(([fieldName, hint]) => (
                  <Alert key={fieldName} severity="info" sx={{ py: 0.5 }}>
                    <Typography variant="caption">
                      <strong>{fieldName}:</strong> {hint.format_string}
                      {hint.is_ambiguous && ' (ambiguo)'}
                      {hint.confidence && ` (${(hint.confidence * 100).toFixed(0)}% confianza)`}
                    </Typography>
                  </Alert>
                ))}
              </Box>
            </Box>
          )}
        </CardContent>
      </Card>

      {/* Problemas Detallados */}
      {validation.issues && validation.issues.length > 0 && (
        <Card sx={{ mb: 2 }}>
          <CardContent>
            <Typography variant="h6" sx={{ mb: 2 }}>
              📋 Problemas Encontrados ({validation.issues.length})
            </Typography>

            <TableContainer component={Paper}>
              <Table size="small">
                <TableHead>
                  <TableRow sx={{ bgcolor: 'grey.100' }}>
                    <TableCell></TableCell>
                    <TableCell>Fila</TableCell>
                    <TableCell>Columna</TableCell>
                    <TableCell>Problema</TableCell>
                    <TableCell>Sugerencia</TableCell>
                  </TableRow>
                </TableHead>
                <TableBody>
                  {validation.issues.map((issue, idx) => (
                    <React.Fragment key={idx}>
                      <TableRow
                        sx={{
                          bgcolor: idx % 2 === 0 ? 'transparent' : 'grey.50',
                          '&:hover': { bgcolor: 'grey.100' },
                        }}
                      >
                        <TableCell>
                          <IconButton
                            size="small"
                            onClick={() => toggleRowExpand(idx)}
                          >
                            {expandedRows.has(idx) ? <ExpandLessIcon /> : <ExpandMoreIcon />}
                          </IconButton>
                        </TableCell>
                        <TableCell>{issue.row || '—'}</TableCell>
                        <TableCell>{issue.column || '—'}</TableCell>
                        <TableCell>
                          <Box sx={{ display: 'flex', alignItems: 'center', gap: 1 }}>
                            {getSeverityIcon(issue.error_type)}
                            <Typography variant="body2">{issue.message.substring(0, 50)}</Typography>
                          </Box>
                        </TableCell>
                        <TableCell>
                          <Typography variant="caption" color="info.main">
                            {issue.suggestion ? issue.suggestion.substring(0, 40) : '—'}
                          </Typography>
                        </TableCell>
                      </TableRow>

                      {/* Detalles expandibles */}
                      <TableRow>
                        <TableCell colSpan={5}>
                          <Collapse in={expandedRows.has(idx)} timeout="auto" unmountOnExit>
                            <Box sx={{ p: 2, bgcolor: 'grey.50' }}>
                              <Alert severity={getSeverityColor(issue.error_type)} sx={{ mb: 1 }}>
                                <AlertTitle>{issue.error_type}</AlertTitle>
                                {issue.message}
                              </Alert>
                              {issue.suggestion && (
                                <Alert severity="info">
                                  <AlertTitle>💡 Sugerencia</AlertTitle>
                                  {issue.suggestion}
                                </Alert>
                              )}
                              {issue.value && (
                                <Typography variant="caption" sx={{ mt: 1, display: 'block' }}>
                                  <strong>Valor:</strong> "{issue.value}"
                                </Typography>
                              )}
                            </Box>
                          </Collapse>
                        </TableCell>
                      </TableRow>
                    </React.Fragment>
                  ))}
                </TableBody>
              </Table>
            </TableContainer>
          </CardContent>
        </Card>
      )}

      {/* Acciones */}
      {!validation.is_valid && onRetry && (
        <Alert severity="warning">
          <AlertTitle>⚠️ Revisa los errores antes de continuar</AlertTitle>
          Corrige tu archivo CSV/XLSX y vuelve a cargar.
        </Alert>
      )}
    </Box>
  );
};

export default ValidationResults;
