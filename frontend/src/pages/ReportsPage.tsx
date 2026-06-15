import React, { useState, useEffect } from 'react';
import {
  Container,
  Box,
  Paper,
  Typography,
  Button,
  Select,
  MenuItem,
  FormControl,
  InputLabel,
  Grid,
  Card,
  CardContent,
  CircularProgress,
  Alert,
  Table,
  TableBody,
  TableCell,
  TableContainer,
  TableHead,
  TableRow,
  Chip,
  Dialog,
  DialogTitle,
  DialogContent,
  DialogActions,
  Tab,
  Tabs,
  LinearProgress,
  TextField
} from '@mui/material';
import {
  LineChart,
  Line,
  BarChart,
  Bar,
  ResponsiveContainer,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  Legend
} from 'recharts';
import FileDownloadIcon from '@mui/icons-material/FileDownload';
import TrendingUpIcon from '@mui/icons-material/TrendingUp';
import WarningIcon from '@mui/icons-material/Warning';
import SaveIcon from '@mui/icons-material/Save';

interface TabPanelProps {
  children?: React.ReactNode;
  index: number;
  value: number;
}

function TabPanel(props: TabPanelProps) {
  const { children, value, index, ...other } = props;
  return (
    <div role="tabpanel" hidden={value !== index} {...other}>
      {value === index && <Box sx={{ p: 2 }}>{children}</Box>}
    </div>
  );
}

export default function ReportsPage() {
  const [tabValue, setTabValue] = useState(0);
  const [anomalies, setAnomalies] = useState<any>(null);
  const [patterns, setPatterns] = useState<any>(null);
  const [forecast, setForecast] = useState<any>(null);
  
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [success, setSuccess] = useState('');
  
  const [anomalyDays, setAnomalyDays] = useState(30);
  const [forecastMonths, setForecastMonths] = useState(3);
  const [saveViewDialogOpen, setSaveViewDialogOpen] = useState(false);
  const [viewName, setViewName] = useState('');

  useEffect(() => {
    loadReportData();
  }, [anomalyDays, forecastMonths]);

  const loadReportData = async () => {
    try {
      setLoading(true);
      setError('');

      // Load anomalies
      const anomaliesRes = await fetch(`/api/advanced-analytics/anomalies?days=${anomalyDays}`);
      if (anomaliesRes.ok) {
        setAnomalies(await anomaliesRes.json());
      }

      // Load patterns
      const patternsRes = await fetch('/api/advanced-analytics/spending-patterns?months=3');
      if (patternsRes.ok) {
        setPatterns(await patternsRes.json());
      }

      // Load forecast
      const forecastRes = await fetch(`/api/advanced-analytics/forecasting?months_ahead=${forecastMonths}`);
      if (forecastRes.ok) {
        setForecast(await forecastRes.json());
      }
    } catch (err: any) {
      setError(err.message || 'Error loading reports');
    } finally {
      setLoading(false);
    }
  };

  const handleExportCSV = async () => {
    try {
      const response = await fetch('/api/exports/expenses-csv');
      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `expenses_${new Date().toISOString().split('T')[0]}.csv`;
      a.click();
      setSuccess('Export successful!');
    } catch (err: any) {
      setError('Export failed');
    }
  };

  const handleSaveView = async () => {
    try {
      const response = await fetch('/api/views/', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          name: viewName,
          filters: {},
          layout: { activeTab: tabValue }
        })
      });

      if (response.ok) {
        setSuccess('View saved successfully!');
        setSaveViewDialogOpen(false);
        setViewName('');
      }
    } catch (err: any) {
      setError('Failed to save view');
    }
  };

  return (
    <Container maxWidth="lg" sx={{ py: 4 }}>
      <Box sx={{ mb: 3, display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
        <Typography variant="h4">Advanced Reports & Analytics</Typography>
        <Box sx={{ display: 'flex', gap: 1 }}>
          <Button
            variant="outlined"
            startIcon={<FileDownloadIcon />}
            onClick={handleExportCSV}
          >
            Export CSV
          </Button>
          <Button
            variant="outlined"
            startIcon={<SaveIcon />}
            onClick={() => setSaveViewDialogOpen(true)}
          >
            Save View
          </Button>
        </Box>
      </Box>

      {error && <Alert severity="error" sx={{ mb: 2 }}>{error}</Alert>}
      {success && <Alert severity="success" sx={{ mb: 2 }}>{success}</Alert>}

      {loading ? (
        <CircularProgress />
      ) : (
        <Paper>
          <Tabs value={tabValue} onChange={(e, newValue) => setTabValue(newValue)}>
            <Tab label="Anomalies" icon={<WarningIcon />} />
            <Tab label="Patterns" icon={<TrendingUpIcon />} />
            <Tab label="Forecasting" />
          </Tabs>

          {/* Anomalies Tab */}
          <TabPanel value={tabValue} index={0}>
            <Box sx={{ mb: 3 }}>
              <FormControl sx={{ minWidth: 200 }}>
                <InputLabel>Analyze Last</InputLabel>
                <Select
                  value={anomalyDays}
                  onChange={(e) => setAnomalyDays(e.target.value as any)}
                  label="Analyze Last"
                >
                  <MenuItem value={7}>7 Days</MenuItem>
                  <MenuItem value={30}>30 Days</MenuItem>
                  <MenuItem value={60}>60 Days</MenuItem>
                  <MenuItem value={90}>90 Days</MenuItem>
                </Select>
              </FormControl>
            </Box>

            {anomalies && (
              <>
                <Grid container spacing={2} sx={{ mb: 3 }}>
                  <Grid item xs={12} sm={6} md={3}>
                    <Card>
                      <CardContent>
                        <Typography color="textSecondary" gutterBottom>Total Expenses</Typography>
                        <Typography variant="h5">{anomalies.analysis?.total_expenses}</Typography>
                      </CardContent>
                    </Card>
                  </Grid>
                  <Grid item xs={12} sm={6} md={3}>
                    <Card>
                      <CardContent>
                        <Typography color="textSecondary" gutterBottom>Anomalies Found</Typography>
                        <Typography variant="h5" sx={{ color: '#d32f2f' }}>
                          {anomalies.anomalies?.length}
                        </Typography>
                      </CardContent>
                    </Card>
                  </Grid>
                </Grid>

                {anomalies.anomalies?.length > 0 ? (
                  <TableContainer>
                    <Table>
                      <TableHead>
                        <TableRow sx={{ backgroundColor: '#f5f5f5' }}>
                          <TableCell>Category</TableCell>
                          <TableCell align="right">Amount</TableCell>
                          <TableCell align="right">Average</TableCell>
                          <TableCell align="right">Z-Score</TableCell>
                          <TableCell>Severity</TableCell>
                        </TableRow>
                      </TableHead>
                      <TableBody>
                        {anomalies.anomalies?.map((anomaly: any, idx: number) => (
                          <TableRow key={idx} hover>
                            <TableCell>{anomaly.category}</TableCell>
                            <TableCell align="right">${anomaly.amount?.toFixed(2)}</TableCell>
                            <TableCell align="right">${anomaly.mean?.toFixed(2)}</TableCell>
                            <TableCell align="right">{anomaly.z_score?.toFixed(2)}</TableCell>
                            <TableCell>
                              <Chip
                                label={anomaly.severity}
                                color={anomaly.severity === 'high' ? 'error' : 'warning'}
                                size="small"
                              />
                            </TableCell>
                          </TableRow>
                        ))}
                      </TableBody>
                    </Table>
                  </TableContainer>
                ) : (
                  <Alert severity="success">No anomalies detected! Spending is within normal patterns.</Alert>
                )}
              </>
            )}
          </TabPanel>

          {/* Patterns Tab */}
          <TabPanel value={tabValue} index={1}>
            {patterns && (
              <>
                <Grid container spacing={2} sx={{ mb: 3 }}>
                  <Grid item xs={12} sm={6} md={3}>
                    <Card>
                      <CardContent>
                        <Typography color="textSecondary" gutterBottom>Total Spending</Typography>
                        <Typography variant="h5">${patterns.total_spending?.toFixed(2)}</Typography>
                      </CardContent>
                    </Card>
                  </Grid>
                  <Grid item xs={12} sm={6} md={3}>
                    <Card>
                      <CardContent>
                        <Typography color="textSecondary" gutterBottom>Transactions</Typography>
                        <Typography variant="h5">{patterns.total_transactions}</Typography>
                      </CardContent>
                    </Card>
                  </Grid>
                </Grid>

                <Typography variant="h6" sx={{ mb: 2 }}>Spending by Day of Week</Typography>
                <Box sx={{ mb: 3 }}>
                  {Object.entries(patterns.by_day_of_week || {}).map(([day, data]: any) => (
                    <Box key={day} sx={{ mb: 2 }}>
                      <Box sx={{ display: 'flex', justifyContent: 'space-between', mb: 0.5 }}>
                        <Typography variant="body2">{day}</Typography>
                        <Typography variant="body2">${data.average?.toFixed(2)}</Typography>
                      </Box>
                      <LinearProgress variant="determinate" value={Math.min((data.average / 100) * 100, 100)} />
                    </Box>
                  ))}
                </Box>
              </>
            )}
          </TabPanel>

          {/* Forecasting Tab */}
          <TabPanel value={tabValue} index={2}>
            <Box sx={{ mb: 3 }}>
              <FormControl sx={{ minWidth: 200 }}>
                <InputLabel>Forecast</InputLabel>
                <Select
                  value={forecastMonths}
                  onChange={(e) => setForecastMonths(e.target.value as any)}
                  label="Forecast"
                >
                  <MenuItem value={1}>1 Month</MenuItem>
                  <MenuItem value={3}>3 Months</MenuItem>
                  <MenuItem value={6}>6 Months</MenuItem>
                </Select>
              </FormControl>
            </Box>

            {forecast && (
              <>
                <Grid container spacing={2} sx={{ mb: 3 }}>
                  <Grid item xs={12} sm={6} md={3}>
                    <Card>
                      <CardContent>
                        <Typography color="textSecondary" gutterBottom>Average Monthly</Typography>
                        <Typography variant="h5">${forecast.average_monthly_spending?.toFixed(2)}</Typography>
                      </CardContent>
                    </Card>
                  </Grid>
                  <Grid item xs={12} sm={6} md={3}>
                    <Card>
                      <CardContent>
                        <Typography color="textSecondary" gutterBottom>Confidence</Typography>
                        <Typography variant="h5">{(forecast.confidence * 100).toFixed(0)}%</Typography>
                      </CardContent>
                    </Card>
                  </Grid>
                </Grid>

                <ResponsiveContainer width="100%" height={300}>
                  <BarChart data={forecast.forecast || []}>
                    <CartesianGrid strokeDasharray="3 3" />
                    <XAxis dataKey="month" />
                    <YAxis />
                    <Tooltip />
                    <Legend />
                    <Bar dataKey="predicted_spending" fill="#1976d2" name="Predicted" />
                    <Bar dataKey="lower_bound" fill="#90caf9" name="Lower Bound" />
                    <Bar dataKey="upper_bound" fill="#64b5f6" name="Upper Bound" />
                  </BarChart>
                </ResponsiveContainer>

                <TableContainer sx={{ mt: 3 }}>
                  <Table>
                    <TableHead>
                      <TableRow sx={{ backgroundColor: '#f5f5f5' }}>
                        <TableCell>Month</TableCell>
                        <TableCell align="right">Predicted</TableCell>
                        <TableCell align="right">Lower Bound</TableCell>
                        <TableCell align="right">Upper Bound</TableCell>
                      </TableRow>
                    </TableHead>
                    <TableBody>
                      {forecast.forecast?.map((item: any, idx: number) => (
                        <TableRow key={idx}>
                          <TableCell>{item.month}</TableCell>
                          <TableCell align="right">${item.predicted_spending?.toFixed(2)}</TableCell>
                          <TableCell align="right">${item.lower_bound?.toFixed(2)}</TableCell>
                          <TableCell align="right">${item.upper_bound?.toFixed(2)}</TableCell>
                        </TableRow>
                      ))}
                    </TableBody>
                  </Table>
                </TableContainer>
              </>
            )}
          </TabPanel>
        </Paper>
      )}

      {/* Save View Dialog */}
      <Dialog open={saveViewDialogOpen} onClose={() => setSaveViewDialogOpen(false)}>
        <DialogTitle>Save Current View</DialogTitle>
        <DialogContent>
          <Box sx={{ pt: 2 }}>
            <TextField
              fullWidth
              label="View Name"
              value={viewName}
              onChange={(e) => setViewName(e.target.value)}
              placeholder="e.g., Monthly Review"
            />
          </Box>
        </DialogContent>
        <DialogActions>
          <Button onClick={() => setSaveViewDialogOpen(false)}>Cancel</Button>
          <Button onClick={handleSaveView} variant="contained">Save</Button>
        </DialogActions>
      </Dialog>
    </Container>
  );
}
