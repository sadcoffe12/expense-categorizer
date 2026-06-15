import React, { useState, useEffect } from 'react';
import {
  Container,
  Box,
  Paper,
  Typography,
  TextField,
  Button,
  Select,
  MenuItem,
  FormControl,
  InputLabel,
  Dialog,
  DialogTitle,
  DialogContent,
  DialogActions,
  Table,
  TableBody,
  TableCell,
  TableContainer,
  TableHead,
  TableRow,
  Chip,
  CircularProgress,
  Alert,
  Grid,
  Card,
  CardContent,
  Tab,
  Tabs,
  IconButton,
  Tooltip
} from '@mui/material';
import DeleteIcon from '@mui/icons-material/Delete';
import EditIcon from '@mui/icons-material/Edit';
import AddIcon from '@mui/icons-material/Add';
import CheckCircleIcon from '@mui/icons-material/CheckCircle';
import WarningIcon from '@mui/icons-material/Warning';

import apiClient from '../services/api';
import { Budget, Category } from '../types';

interface TabPanelProps {
  children?: React.ReactNode;
  index: number;
  value: number;
}

function TabPanel(props: TabPanelProps) {
  const { children, value, index, ...other } = props;
  return (
    <div role="tabpanel" hidden={value !== index} {...other}>
      {value === index && <Box sx={{ p: 3 }}>{children}</Box>}
    </div>
  );
}

export default function SettingsPage() {
  const [tabValue, setTabValue] = useState(0);
  const [categories, setCategories] = useState<Category[]>([]);
  const [budgets, setBudgets] = useState<any[]>([]);
  const [alerts, setAlerts] = useState<any[]>([]);
  const [alertSummary, setAlertSummary] = useState<any>(null);
  
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [success, setSuccess] = useState('');
  
  // Budget dialog state
  const [budgetDialogOpen, setBudgetDialogOpen] = useState(false);
  const [editingBudgetId, setEditingBudgetId] = useState<number | null>(null);
  const [budgetForm, setBudgetForm] = useState({
    category_id: '',
    amount: '',
    period: 'month'
  });

  useEffect(() => {
    loadData();
  }, []);

  const loadData = async () => {
    try {
      setLoading(true);
      setError('');
      
      // Load categories
      const categoriesData = await fetch('/api/setup/load-database').then(r => r.json());
      if (categoriesData.categories) {
        setCategories(categoriesData.categories);
      }
      
      // Load budgets
      const budgetsData = await fetch('/api/budgets/').then(r => r.json());
      setBudgets(budgetsData);
      
      // Load alerts
      const alertsData = await fetch('/api/alerts/').then(r => r.json());
      setAlerts(alertsData);
      
      // Load alert summary
      const summaryData = await fetch('/api/alerts/summary').then(r => r.json());
      setAlertSummary(summaryData);
    } catch (err: any) {
      setError(err.message || 'Error loading data');
    } finally {
      setLoading(false);
    }
  };

  const handleCreateBudget = async () => {
    try {
      if (!budgetForm.category_id || !budgetForm.amount) {
        setError('Please fill in all required fields');
        return;
      }

      setLoading(true);
      setError('');

      const payload = {
        category_id: parseInt(budgetForm.category_id as string),
        amount: parseFloat(budgetForm.amount as string),
        period: budgetForm.period,
        start_date: null,
        end_date: null
      };

      const response = await fetch('/api/budgets/', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });

      if (!response.ok) throw new Error('Failed to create budget');

      setSuccess('Budget created successfully!');
      setBudgetDialogOpen(false);
      setBudgetForm({ category_id: '', amount: '', period: 'month' });
      loadData();
    } catch (err: any) {
      setError(err.message || 'Error creating budget');
    } finally {
      setLoading(false);
    }
  };

  const handleDeleteBudget = async (budgetId: number) => {
    if (!window.confirm('Are you sure you want to delete this budget?')) return;

    try {
      setLoading(true);
      const response = await fetch(`/api/budgets/${budgetId}`, { method: 'DELETE' });
      if (!response.ok) throw new Error('Failed to delete budget');
      
      setSuccess('Budget deleted successfully!');
      loadData();
    } catch (err: any) {
      setError(err.message || 'Error deleting budget');
    } finally {
      setLoading(false);
    }
  };

  const handleAcknowledgeAlert = async (alertId: number) => {
    try {
      const response = await fetch(`/api/alerts/${alertId}/acknowledge`, { method: 'PUT' });
      if (!response.ok) throw new Error('Failed to acknowledge alert');
      
      loadData();
    } catch (err: any) {
      setError(err.message || 'Error acknowledging alert');
    }
  };

  const getCategoryName = (categoryId: number) => {
    const cat = categories.find(c => c.id === categoryId);
    return cat ? cat.name : 'Unknown';
  };

  return (
    <Container maxWidth="lg" sx={{ py: 4 }}>
      <Box sx={{ mb: 4 }}>
        <Typography variant="h4" sx={{ mb: 2 }}>Settings & Management</Typography>
        
        {error && <Alert severity="error" sx={{ mb: 2 }}>{error}</Alert>}
        {success && <Alert severity="success" sx={{ mb: 2 }}>{success}</Alert>}
      </Box>

      {loading && !budgets.length && <CircularProgress />}

      <Paper sx={{ mb: 4 }}>
        <Tabs value={tabValue} onChange={(e, newValue) => setTabValue(newValue)}>
          <Tab label="Budgets" />
          <Tab label="Alerts" />
          <Tab label="Rules & Learning" />
        </Tabs>

        {/* Budgets Tab */}
        <TabPanel value={tabValue} index={0}>
          <Box sx={{ mb: 3 }}>
            <Button
              variant="contained"
              startIcon={<AddIcon />}
              onClick={() => setBudgetDialogOpen(true)}
            >
              New Budget
            </Button>
          </Box>

          <TableContainer>
            <Table>
              <TableHead>
                <TableRow sx={{ backgroundColor: '#f5f5f5' }}>
                  <TableCell>Category</TableCell>
                  <TableCell align="right">Budget Amount</TableCell>
                  <TableCell align="right">Spent</TableCell>
                  <TableCell align="right">Remaining</TableCell>
                  <TableCell>Progress</TableCell>
                  <TableCell>Period</TableCell>
                  <TableCell align="center">Actions</TableCell>
                </TableRow>
              </TableHead>
              <TableBody>
                {budgets.map((budget) => (
                  <TableRow key={budget.id} hover>
                    <TableCell>{budget.category}</TableCell>
                    <TableCell align="right">${budget.amount?.toFixed(2)}</TableCell>
                    <TableCell align="right">
                      <span style={{ color: budget.spent > budget.amount ? '#d32f2f' : '#388e3c' }}>
                        ${budget.spent?.toFixed(2)}
                      </span>
                    </TableCell>
                    <TableCell align="right">
                      ${(budget.amount - (budget.spent || 0)).toFixed(2)}
                    </TableCell>
                    <TableCell>
                      <div style={{ width: '100px', height: '6px', backgroundColor: '#e0e0e0', borderRadius: '3px', overflow: 'hidden' }}>
                        <div
                          style={{
                            height: '100%',
                            width: `${Math.min((budget.spent || 0) / budget.amount * 100, 100)}%`,
                            backgroundColor: budget.spent > budget.amount ? '#d32f2f' : '#388e3c',
                            transition: 'width 0.3s'
                          }}
                        />
                      </div>
                      <Typography variant="caption">{budget.percentage?.toFixed(0)}%</Typography>
                    </TableCell>
                    <TableCell>{budget.period}</TableCell>
                    <TableCell align="center">
                      <Tooltip title="Delete">
                        <IconButton
                          size="small"
                          onClick={() => handleDeleteBudget(budget.id)}
                          color="error"
                        >
                          <DeleteIcon fontSize="small" />
                        </IconButton>
                      </Tooltip>
                    </TableCell>
                  </TableRow>
                ))}
              </TableBody>
            </Table>
          </TableContainer>

          {budgets.length === 0 && <Typography sx={{ p: 3, textAlign: 'center', color: '#999' }}>No budgets set yet</Typography>}
        </TabPanel>

        {/* Alerts Tab */}
        <TabPanel value={tabValue} index={1}>
          {alertSummary && (
            <Box sx={{ mb: 3 }}>
              <Grid container spacing={2}>
                <Grid item xs={12} sm={6} md={4}>
                  <Card>
                    <CardContent>
                      <Typography color="textSecondary" gutterBottom>Unacknowledged Alerts</Typography>
                      <Typography variant="h4">{alertSummary.total_unacknowledged}</Typography>
                    </CardContent>
                  </Card>
                </Grid>
                {Object.entries(alertSummary.by_type || {}).map(([type, count]: any) => (
                  <Grid item xs={12} sm={6} md={4} key={type}>
                    <Card>
                      <CardContent>
                        <Typography color="textSecondary" gutterBottom>{type}</Typography>
                        <Typography variant="h4">{count}</Typography>
                      </CardContent>
                    </Card>
                  </Grid>
                ))}
              </Grid>
            </Box>
          )}

          <TableContainer>
            <Table>
              <TableHead>
                <TableRow sx={{ backgroundColor: '#f5f5f5' }}>
                  <TableCell>Type</TableCell>
                  <TableCell>Message</TableCell>
                  <TableCell>Status</TableCell>
                  <TableCell>Created</TableCell>
                  <TableCell align="center">Actions</TableCell>
                </TableRow>
              </TableHead>
              <TableBody>
                {alerts.map((alert) => (
                  <TableRow key={alert.id} hover>
                    <TableCell>
                      <Chip label={alert.alert_type} size="small" />
                    </TableCell>
                    <TableCell>{alert.message}</TableCell>
                    <TableCell>
                      {alert.acknowledged ? (
                        <Chip icon={<CheckCircleIcon />} label="Acknowledged" color="success" size="small" />
                      ) : (
                        <Chip icon={<WarningIcon />} label="Pending" color="warning" size="small" />
                      )}
                    </TableCell>
                    <TableCell>{new Date(alert.created_at).toLocaleDateString()}</TableCell>
                    <TableCell align="center">
                      {!alert.acknowledged && (
                        <Button
                          size="small"
                          onClick={() => handleAcknowledgeAlert(alert.id)}
                        >
                          Acknowledge
                        </Button>
                      )}
                    </TableCell>
                  </TableRow>
                ))}
              </TableBody>
            </Table>
          </TableContainer>

          {alerts.length === 0 && <Typography sx={{ p: 3, textAlign: 'center', color: '#999' }}>No alerts</Typography>}
        </TabPanel>

        {/* Rules & Learning Tab */}
        <TabPanel value={tabValue} index={2}>
          <Alert severity="info">
            <Typography variant="body2">
              <strong>Rules Learning:</strong> The system learns from your categorization decisions. 
              When you accept a suggestion, it becomes part of the automatic categorization rules.
              Common patterns are automatically turned into rules with high confidence scores.
            </Typography>
          </Alert>

          <Box sx={{ mt: 3, p: 2, backgroundColor: '#f5f5f5', borderRadius: 1 }}>
            <Typography variant="h6" gutterBottom>How It Works</Typography>
            <Typography component="div" variant="body2" sx={{ lineHeight: 1.8 }}>
              <ol>
                <li><strong>Normalize:</strong> Text is cleaned to remove accents and noise</li>
                <li><strong>Match:</strong> Keywords are compared against existing rules</li>
                <li><strong>Score:</strong> Confidence calculated based on similarity (0-1)</li>
                <li><strong>Learn:</strong> Your feedback helps improve future suggestions</li>
                <li><strong>Suggest:</strong> New rules are recommended when patterns emerge</li>
              </ol>
            </Typography>
          </Box>

          <Box sx={{ mt: 3 }}>
            <Button variant="outlined">
              Refresh Rule Suggestions
            </Button>
          </Box>
        </TabPanel>
      </Paper>

      {/* Create Budget Dialog */}
      <Dialog open={budgetDialogOpen} onClose={() => setBudgetDialogOpen(false)}>
        <DialogTitle>Create New Budget</DialogTitle>
        <DialogContent sx={{ minWidth: 400 }}>
          <Box sx={{ pt: 2, display: 'flex', flexDirection: 'column', gap: 2 }}>
            <FormControl fullWidth>
              <InputLabel>Category</InputLabel>
              <Select
                value={budgetForm.category_id}
                onChange={(e) => setBudgetForm({ ...budgetForm, category_id: e.target.value })}
                label="Category"
              >
                {categories.map((cat) => (
                  <MenuItem key={cat.id} value={cat.id}>
                    {cat.name}
                  </MenuItem>
                ))}
              </Select>
            </FormControl>

            <TextField
              fullWidth
              type="number"
              label="Budget Amount"
              value={budgetForm.amount}
              onChange={(e) => setBudgetForm({ ...budgetForm, amount: e.target.value })}
              inputProps={{ step: '0.01', min: '0' }}
            />

            <FormControl fullWidth>
              <InputLabel>Period</InputLabel>
              <Select
                value={budgetForm.period}
                onChange={(e) => setBudgetForm({ ...budgetForm, period: e.target.value })}
                label="Period"
              >
                <MenuItem value="month">Monthly</MenuItem>
                <MenuItem value="year">Yearly</MenuItem>
                <MenuItem value="custom">Custom</MenuItem>
              </Select>
            </FormControl>
          </Box>
        </DialogContent>
        <DialogActions>
          <Button onClick={() => setBudgetDialogOpen(false)}>Cancel</Button>
          <Button onClick={handleCreateBudget} variant="contained">Create</Button>
        </DialogActions>
      </Dialog>
    </Container>
  );
}
