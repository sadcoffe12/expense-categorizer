import React, { useState, useEffect } from 'react';
import {
  Container,
  Box,
  Paper,
  Table,
  TableBody,
  TableCell,
  TableContainer,
  TableHead,
  TableRow,
  Button,
  Dialog,
  DialogTitle,
  DialogContent,
  DialogActions,
  TextField,
  Select,
  MenuItem,
  FormControl,
  InputLabel,
  CircularProgress,
  Alert,
  Chip,
  Grid,
} from '@mui/material';
import DeleteIcon from '@mui/icons-material/Delete';
import EditIcon from '@mui/icons-material/Edit';
import apiClient from '../services/api';
import { Expense } from '../types';

export default function ExpensesPage() {
  const [expenses, setExpenses] = useState<Expense[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [editingId, setEditingId] = useState<number | null>(null);
  const [openDialog, setOpenDialog] = useState(false);
  const [filterCategory, setFilterCategory] = useState('');
  const [editFormData, setEditFormData] = useState({
    description: '',
    category_id: 0,
    amount: 0,
    type: '',
    location: '',
    notes: '',
  });

  useEffect(() => {
    loadExpenses();
  }, [filterCategory]);

  const loadExpenses = async () => {
    try {
      setLoading(true);
      const params: any = { limit: 500 };
      if (filterCategory) {
        params.category_id = filterCategory;
      }
      const response = await apiClient.get('/expenses/', { params });
      setExpenses(response.data);
    } catch (err: any) {
      setError(err.response?.data?.detail || 'Error cargando gastos');
    } finally {
      setLoading(false);
    }
  };

  const handleEdit = (expense: Expense) => {
    setEditingId(expense.id);
    setEditFormData({
      description: expense.description,
      category_id: expense.category_id,
      amount: expense.amount,
      type: expense.type,
      location: expense.location || '',
      notes: expense.notes || '',
    });
    setOpenDialog(true);
  };

  const handleSave = async () => {
    if (!editingId) return;

    try {
      const updateData: any = {};
      if (editFormData.description !== expenses.find(e => e.id === editingId)?.description) {
        updateData.description = editFormData.description;
      }
      if (editFormData.category_id !== expenses.find(e => e.id === editingId)?.category_id) {
        updateData.category_id = editFormData.category_id;
      }
      if (editFormData.amount !== expenses.find(e => e.id === editingId)?.amount) {
        updateData.amount = editFormData.amount;
      }
      if (editFormData.type !== expenses.find(e => e.id === editingId)?.type) {
        updateData.type_ = editFormData.type;
      }

      await apiClient.put(`/expenses/${editingId}`, updateData);
      setOpenDialog(false);
      await loadExpenses();
    } catch (err: any) {
      setError(err.response?.data?.detail || 'Error actualizando gasto');
    }
  };

  const handleDelete = async (id: number) => {
    if (!window.confirm('¿Eliminar este gasto?')) return;

    try {
      await apiClient.delete(`/expenses/${id}`);
      await loadExpenses();
    } catch (err: any) {
      setError(err.response?.data?.detail || 'Error eliminando gasto');
    }
  };

  if (loading) {
    return (
      <Container maxWidth="lg" sx={{ py: 4, textAlign: 'center' }}>
        <CircularProgress />
      </Container>
    );
  }

  return (
    <Container maxWidth="lg" sx={{ py: 4 }}>
      {error && <Alert severity="error" sx={{ mb: 2 }}>{error}</Alert>}

      <Box sx={{ mb: 3 }}>
        <FormControl sx={{ minWidth: 200 }}>
          <InputLabel>Filtrar por categoría</InputLabel>
          <Select
            value={filterCategory}
            label="Filtrar por categoría"
            onChange={(e) => setFilterCategory(e.target.value)}
          >
            <MenuItem value="">Todas</MenuItem>
            {[...new Set(expenses.map(e => e.category))].map((cat) => (
              <MenuItem key={cat} value={expenses.find(e => e.category === cat)?.category_id}>
                {cat}
              </MenuItem>
            ))}
          </Select>
        </FormControl>
      </Box>

      <TableContainer component={Paper}>
        <Table>
          <TableHead sx={{ backgroundColor: '#f5f5f5' }}>
            <TableRow>
              <TableCell sx={{ fontWeight: 'bold' }}>Fecha</TableCell>
              <TableCell sx={{ fontWeight: 'bold' }}>Descripción</TableCell>
              <TableCell sx={{ fontWeight: 'bold' }} align="right">
                Monto
              </TableCell>
              <TableCell sx={{ fontWeight: 'bold' }}>Categoría</TableCell>
              <TableCell sx={{ fontWeight: 'bold' }}>Tipo</TableCell>
              <TableCell sx={{ fontWeight: 'bold' }} align="center">
                Acciones
              </TableCell>
            </TableRow>
          </TableHead>
          <TableBody>
            {expenses.map((expense) => (
              <TableRow key={expense.id} hover>
                <TableCell>{expense.date}</TableCell>
                <TableCell>{expense.description}</TableCell>
                <TableCell align="right" sx={{ fontWeight: 'bold' }}>
                  ${expense.amount.toFixed(2)}
                </TableCell>
                <TableCell>
                  <Chip label={expense.category} size="small" variant="outlined" />
                </TableCell>
                <TableCell>
                  <Chip
                    label={expense.type}
                    size="small"
                    color={expense.type === 'Gasto' ? 'error' : 'success'}
                    variant="filled"
                  />
                </TableCell>
                <TableCell align="center">
                  <Button
                    size="small"
                    startIcon={<EditIcon />}
                    onClick={() => handleEdit(expense)}
                  >
                    Editar
                  </Button>
                  <Button
                    size="small"
                    color="error"
                    startIcon={<DeleteIcon />}
                    onClick={() => handleDelete(expense.id)}
                  >
                    Eliminar
                  </Button>
                </TableCell>
              </TableRow>
            ))}
          </TableBody>
        </Table>
      </TableContainer>

      {/* Edit Dialog */}
      <Dialog open={openDialog} onClose={() => setOpenDialog(false)} maxWidth="sm" fullWidth>
        <DialogTitle>Editar Gasto</DialogTitle>
        <DialogContent sx={{ pt: 2 }}>
          <Grid container spacing={2}>
            <Grid item xs={12}>
              <TextField
                label="Descripción"
                fullWidth
                value={editFormData.description}
                onChange={(e) =>
                  setEditFormData({ ...editFormData, description: e.target.value })
                }
              />
            </Grid>
            <Grid item xs={12}>
              <TextField
                label="Monto"
                type="number"
                fullWidth
                value={editFormData.amount}
                onChange={(e) =>
                  setEditFormData({
                    ...editFormData,
                    amount: parseFloat(e.target.value),
                  })
                }
              />
            </Grid>
            <Grid item xs={12}>
              <TextField
                label="Tipo"
                fullWidth
                value={editFormData.type}
                onChange={(e) => setEditFormData({ ...editFormData, type: e.target.value })}
              />
            </Grid>
            <Grid item xs={12}>
              <TextField
                label="Ubicación"
                fullWidth
                value={editFormData.location}
                onChange={(e) =>
                  setEditFormData({ ...editFormData, location: e.target.value })
                }
              />
            </Grid>
            <Grid item xs={12}>
              <TextField
                label="Notas"
                fullWidth
                multiline
                rows={3}
                value={editFormData.notes}
                onChange={(e) =>
                  setEditFormData({ ...editFormData, notes: e.target.value })
                }
              />
            </Grid>
          </Grid>
        </DialogContent>
        <DialogActions>
          <Button onClick={() => setOpenDialog(false)}>Cancelar</Button>
          <Button onClick={handleSave} variant="contained">
            Guardar
          </Button>
        </DialogActions>
      </Dialog>
    </Container>
  );
}
