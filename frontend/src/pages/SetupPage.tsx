import React, { useState } from 'react';
import {
  Container,
  Box,
  Paper,
  Button,
  Typography,
  Alert,
  CircularProgress,
} from '@mui/material';
import CloudUploadIcon from '@mui/icons-material/CloudUpload';
import DatabaseIcon from '@mui/icons-material/Storage';
import { setupAPI } from '../services/api';
import { ColumnMapping, ParseFileResponse } from '../types';
import FileUpload from '../components/FileUpload';
import ColumnMapper from '../components/ColumnMapper';
import ValidationResults from '../components/ValidationResults';

interface SetupPageProps {
  onSetupComplete: () => void;
}

type SetupStep = 'select' | 'upload' | 'validation' | 'mapping' | 'processing' | 'complete';

export default function SetupPage({ onSetupComplete }: SetupPageProps) {
  const [step, setStep] = useState<SetupStep>('select');
  const [uploadType, setUploadType] = useState<'sql' | 'csv' | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [fileData, setFileData] = useState<ParseFileResponse | null>(null);
  const [selectedFile, setSelectedFile] = useState<File | null>(null);
  const [importResults, setImportResults] = useState<any>(null);

  const handleSelectUploadType = (type: 'sql' | 'csv') => {
    setUploadType(type);
    setStep('upload');
  };

  const handleFileSelected = async (file: File) => {
    setSelectedFile(file);
    setLoading(true);
    setError(null);

    try {
      if (uploadType === 'sql') {
        const result = await setupAPI.validateSQL(file);
        if (!result.valid) {
          setError(`SQL no válido: ${result.errors.join(', ')}`);
          return;
        }
        setImportResults(result);
        setStep('processing');
        // Simular importación rápida
        setTimeout(() => {
          setImportResults({ records_imported: result.record_count, database_path: 'data/expense.db' });
          setStep('complete');
          setTimeout(() => onSetupComplete(), 2000);
        }, 1500);
      } else {
        const data = await setupAPI.parseFile(file);
        setFileData(data);
        setStep('validation');  // Ir a validación primero
      }
    } catch (err: any) {
      setError(err.response?.data?.detail || 'Error procesando archivo');
    } finally {
      setLoading(false);
    }
  };

  const handleMappingComplete = async (mapping: ColumnMapping) => {
    if (!selectedFile) return;

    setLoading(true);
    setStep('processing');
    setError(null);

    try {
      const result = await setupAPI.createDatabase(selectedFile, mapping);

      if (result.success) {
        setImportResults(result);
        setStep('complete');
        setTimeout(() => onSetupComplete(), 2000);
      } else {
        setError(`Error: ${result.errors.join(', ')}`);
        setStep('mapping');
      }
    } catch (err: any) {
      setError(err.response?.data?.detail || 'Error creando BD');
      setStep('mapping');
    } finally {
      setLoading(false);
    }
  };

  const handleReset = () => {
    setStep('select');
    setUploadType(null);
    setFileData(null);
    setSelectedFile(null);
    setError(null);
  };

  return (
    <Container maxWidth="sm">
      <Box sx={{ py: 4 }}>
        <Paper elevation={3} sx={{ p: 4 }}>
          <Typography variant="h4" gutterBottom sx={{ textAlign: 'center', mb: 4 }}>
            📊 Expense Categorizer
          </Typography>

          {step === 'select' && (
            <Box>
              <Typography variant="h6" gutterBottom>
                ¿De dónde cargar tus datos?
              </Typography>

              <Box sx={{ display: 'flex', gap: 2, flexDirection: 'column' }}>
                <Button
                  variant="contained"
                  size="large"
                  startIcon={<DatabaseIcon />}
                  onClick={() => handleSelectUploadType('sql')}
                >
                  📄 Sube SQL Existente
                </Button>
                <Button
                  variant="outlined"
                  size="large"
                  startIcon={<CloudUploadIcon />}
                  onClick={() => handleSelectUploadType('csv')}
                >
                  📊 Sube CSV o XLSX
                </Button>
              </Box>
            </Box>
          )}

          {step === 'upload' && (
            <Box>
              <Typography variant="h6" gutterBottom>
                {uploadType === 'sql' ? 'Sube archivo SQL' : 'Sube archivo CSV o XLSX'}
              </Typography>

              {error && <Alert severity="error" sx={{ mb: 2 }}>{error}</Alert>}

              <FileUpload 
                onFileSelected={handleFileSelected}
                accept={uploadType === 'sql' ? '.db' : '.csv,.xlsx'}
                loading={loading}
              />

              <Button
                variant="text"
                sx={{ mt: 2 }}
                onClick={handleReset}
              >
                ← Volver
              </Button>
            </Box>
          )}

          {step === 'validation' && fileData && (
            <Box>
              <Typography variant="h6" gutterBottom>
                ✓ Validación de Datos
              </Typography>

              <ValidationResults 
                validation={fileData.validation_result}
                onRetry={handleReset}
              />

              <Box sx={{ display: 'flex', gap: 2, mt: 3 }}>
                <Button
                  variant="outlined"
                  onClick={() => setStep('upload')}
                >
                  ← Cargar otro archivo
                </Button>
                <Button
                  variant="contained"
                  onClick={() => setStep('mapping')}
                  disabled={fileData.validation_result && !fileData.validation_result.is_valid}
                >
                  Continuar →
                </Button>
              </Box>
            </Box>
          )}

          {step === 'mapping' && fileData && (
            <Box>
              <Typography variant="h6" gutterBottom>
                Mapea tus columnas
              </Typography>

              {error && <Alert severity="error" sx={{ mb: 2 }}>{error}</Alert>}

              <ColumnMapper
                headers={fileData.headers}
                preview={fileData.preview}
                onComplete={handleMappingComplete}
                loading={loading}
                suggestedMapping={fileData.suggested_mapping}
              />
            </Box>
          )}

          {step === 'processing' && (
            <Box sx={{ textAlign: 'center' }}>
              <CircularProgress sx={{ mb: 2 }} />
              <Typography variant="body1">
                Importando datos...
              </Typography>
            </Box>
          )}

          {step === 'complete' && (
            <Box sx={{ textAlign: 'center' }}>
              <Typography variant="h5" sx={{ mb: 2, color: 'success.main' }}>
                ✅ ¡Configuración Completada!
              </Typography>
              {importResults && (
                <Box sx={{ mb: 3 }}>
                  <Typography>
                    📊 {importResults.records_imported || 0} transacciones importadas
                  </Typography>
                </Box>
              )}
              <Typography variant="body2" color="textSecondary">
                Cargando dashboard...
              </Typography>
            </Box>
          )}
        </Paper>
      </Box>
    </Container>
  );
}
