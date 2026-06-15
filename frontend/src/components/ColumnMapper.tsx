import React, { useState } from 'react';
import {
  Box,
  Button,
  FormControl,
  InputLabel,
  MenuItem,
  Select,
  Alert,
  Typography,
} from '@mui/material';
import { ColumnMapping } from '../types';

interface ColumnMapperProps {
  headers: string[];
  preview: any[];
  onComplete: (mapping: ColumnMapping) => void;
  loading?: boolean;
  suggestedMapping?: Record<string, string>;
}

export default function ColumnMapper({
  headers,
  preview,
  onComplete,
  loading = false,
  suggestedMapping = {},
}: ColumnMapperProps) {
  const [mapping, setMapping] = useState<Partial<ColumnMapping>>(suggestedMapping);
  const [errors, setErrors] = useState<string[]>([]);

  const requiredFields = ['fecha', 'concepto', 'monto', 'categoria', 'tipo'];

  const handleMappingChange = (field: string, value: string | null) => {
    setMapping(prev => ({
      ...prev,
      [field]: value,
    }));
  };

  const validateMapping = () => {
    const newErrors: string[] = [];

    for (const field of requiredFields) {
      if (!mapping[field as keyof ColumnMapping]) {
        newErrors.push(`Campo requerido: ${field}`);
      }
    }

    if (newErrors.length > 0) {
      setErrors(newErrors);
      return false;
    }

    return true;
  };

  const handleComplete = () => {
    if (validateMapping()) {
      onComplete(mapping as ColumnMapping);
    }
  };

  return (
    <Box>
      {errors.length > 0 && (
        <Alert severity="error" sx={{ mb: 2 }}>
          {errors.join(', ')}
        </Alert>
      )}

      <Typography variant="subtitle2" sx={{ mb: 2, fontWeight: 'bold' }}>
        Mapea tus columnas (campos con * son requeridos)
      </Typography>

      <Box sx={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 2, mb: 3 }}>
        {requiredFields.map(field => (
          <FormControl fullWidth key={field}>
            <InputLabel>{field} *</InputLabel>
            <Select
              value={mapping[field as keyof ColumnMapping] || ''}
              label={field + ' *'}
              onChange={(e) => handleMappingChange(field, e.target.value || null)}
            >
              {headers.map(h => (
                <MenuItem key={h} value={h}>{h}</MenuItem>
              ))}
            </Select>
          </FormControl>
        ))}
      </Box>

      <Button
        variant="contained"
        onClick={handleComplete}
        disabled={loading}
        fullWidth
      >
        Siguiente →
      </Button>
    </Box>
  );
}
