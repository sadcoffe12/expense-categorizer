import React, { useRef } from 'react';
import {
  Box,
  Button,
  Typography,
  CircularProgress,
} from '@mui/material';
import CloudUploadIcon from '@mui/icons-material/CloudUpload';

interface FileUploadProps {
  onFileSelected: (file: File) => void;
  accept?: string;
  loading?: boolean;
}

export default function FileUpload({
  onFileSelected,
  accept = '.csv,.xlsx',
  loading = false,
}: FileUploadProps) {
  const inputRef = useRef<HTMLInputElement>(null);

  const handleClick = () => {
    inputRef.current?.click();
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      onFileSelected(file);
    }
  };

  return (
    <Box
      onClick={handleClick}
      sx={{
        border: '2px dashed #ccc',
        borderRadius: 2,
        p: 3,
        textAlign: 'center',
        cursor: 'pointer',
        transition: 'all 0.3s',
        '&:hover': {
          borderColor: '#1976d2',
          backgroundColor: '#f5f5f5',
        },
      }}
    >
      <input
        ref={inputRef}
        type="file"
        accept={accept}
        onChange={handleFileChange}
        style={{ display: 'none' }}
        disabled={loading}
      />

      {loading ? (
        <>
          <CircularProgress sx={{ mb: 1 }} />
          <Typography variant="body2">Procesando archivo...</Typography>
        </>
      ) : (
        <>
          <CloudUploadIcon sx={{ fontSize: 48, mb: 1, color: '#1976d2' }} />
          <Typography variant="body1" sx={{ fontWeight: 'bold' }}>
            Arrastra tu archivo aquí
          </Typography>
          <Typography variant="body2" color="textSecondary">
            o haz clic para seleccionar
          </Typography>
          <Typography variant="caption" color="textSecondary" sx={{ mt: 1 }}>
            Formatos soportados: {accept === '.db' ? 'SQL' : 'CSV, XLSX'}
          </Typography>
        </>
      )}
    </Box>
  );
}
