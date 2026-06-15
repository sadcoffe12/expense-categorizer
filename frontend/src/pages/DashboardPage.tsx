import React, { useState, useEffect } from 'react';
import {
  Container,
  Box,
  Grid,
  Paper,
  Typography,
  Card,
  CardContent,
  CircularProgress,
  Alert,
} from '@mui/material';
import {
  PieChart,
  Pie,
  Cell,
  LineChart,
  Line,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  Legend,
  ResponsiveContainer,
  BarChart,
  Bar,
} from 'recharts';
import apiClient from '../services/api';
import { ConfigStatus } from '../types';

const COLORS = ['#8884d8', '#82ca9d', '#ffc658', '#ff7c7c', '#8dd1e1', '#d084d0'];

interface SummaryData {
  total: number;
  count: number;
  average: number;
  min: number;
  max: number;
  median: number;
  by_type: Record<string, any>;
  by_category: Record<string, any>;
}

interface TrendData {
  period: string;
  data: Record<string, { total: number; count: number; average: number }>;
}

export default function DashboardPage() {
  const [summary, setSummary] = useState<SummaryData | null>(null);
  const [trends, setTrends] = useState<TrendData | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    loadDashboardData();
  }, []);

  const loadDashboardData = async () => {
    try {
      setLoading(true);
      setError(null);

      // Get last 30 days summary
      const thirtyDaysAgo = new Date();
      thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);

      const [summaryRes, trendsRes] = await Promise.all([
        apiClient.get('/analytics/summary', {
          params: { date_from: thirtyDaysAgo.toISOString().split('T')[0] },
        }),
        apiClient.get('/analytics/trends', {
          params: { period: 'daily', months: 1 },
        }),
      ]);

      setSummary(summaryRes.data);
      setTrends(trendsRes.data);
    } catch (err: any) {
      setError(err.response?.data?.detail || 'Error cargando dashboard');
    } finally {
      setLoading(false);
    }
  };

  if (loading) {
    return (
      <Container maxWidth="lg" sx={{ py: 4, textAlign: 'center' }}>
        <CircularProgress />
        <Typography variant="body2" sx={{ mt: 2 }}>
          Cargando dashboard...
        </Typography>
      </Container>
    );
  }

  if (error) {
    return (
      <Container maxWidth="lg" sx={{ py: 4 }}>
        <Alert severity="error">{error}</Alert>
      </Container>
    );
  }

  // Summary Cards
  const summaryCards = summary
    ? [
        { label: 'Total', value: `$${summary.total.toFixed(2)}`, color: '#1976d2' },
        {
          label: 'Promedio',
          value: `$${summary.average.toFixed(2)}`,
          color: '#388e3c',
        },
        { label: 'Cantidad', value: summary.count, color: '#f57c00' },
        { label: 'Mediana', value: `$${summary.median.toFixed(2)}`, color: '#c2185b' },
      ]
    : [];

  // Category data for pie chart
  const categoryData = summary
    ? Object.entries(summary.by_category).map(([name, data]: any) => ({
        name,
        value: data.total,
      }))
    : [];

  // Trend data
  const trendData = trends
    ? Object.entries(trends.data)
        .sort(([dateA], [dateB]) => dateA.localeCompare(dateB))
        .map(([date, data]: any) => ({
          date,
          total: data.total,
        }))
    : [];

  return (
    <Container maxWidth="lg">
      <Box sx={{ py: 4 }}>
        <Typography variant="h4" gutterBottom sx={{ mb: 4, fontWeight: 'bold' }}>
          📊 Dashboard de Gastos
        </Typography>

        {/* Summary Cards */}
        <Grid container spacing={2} sx={{ mb: 4 }}>
          {summaryCards.map((card, idx) => (
            <Grid item xs={12} sm={6} md={3} key={idx}>
              <Card>
                <CardContent sx={{ backgroundColor: '#f5f5f5' }}>
                  <Typography color="textSecondary" gutterBottom>
                    {card.label}
                  </Typography>
                  <Typography variant="h5" sx={{ color: card.color, fontWeight: 'bold' }}>
                    {card.value}
                  </Typography>
                </CardContent>
              </Card>
            </Grid>
          ))}
        </Grid>

        {/* Charts */}
        <Grid container spacing={3}>
          {/* Line Chart - Trends */}
          <Grid item xs={12} md={6}>
            <Paper elevation={2} sx={{ p: 2 }}>
              <Typography variant="h6" gutterBottom sx={{ fontWeight: 'bold' }}>
                📈 Tendencia de Gastos
              </Typography>
              {trendData.length > 0 ? (
                <ResponsiveContainer width="100%" height={300}>
                  <LineChart data={trendData}>
                    <CartesianGrid strokeDasharray="3 3" />
                    <XAxis
                      dataKey="date"
                      angle={-45}
                      textAnchor="end"
                      height={80}
                    />
                    <YAxis />
                    <Tooltip />
                    <Line
                      type="monotone"
                      dataKey="total"
                      stroke="#1976d2"
                      strokeWidth={2}
                      dot={{ fill: '#1976d2' }}
                    />
                  </LineChart>
                </ResponsiveContainer>
              ) : (
                <Typography color="textSecondary">Sin datos</Typography>
              )}
            </Paper>
          </Grid>

          {/* Pie Chart - Categories */}
          <Grid item xs={12} md={6}>
            <Paper elevation={2} sx={{ p: 2 }}>
              <Typography variant="h6" gutterBottom sx={{ fontWeight: 'bold' }}>
                🥧 Por Categoría
              </Typography>
              {categoryData.length > 0 ? (
                <ResponsiveContainer width="100%" height={300}>
                  <PieChart>
                    <Pie
                      data={categoryData}
                      cx="50%"
                      cy="50%"
                      labelLine={false}
                      label={({ name, percent }) =>
                        `${name}: ${(percent * 100).toFixed(0)}%`
                      }
                      outerRadius={100}
                      fill="#8884d8"
                      dataKey="value"
                    >
                      {categoryData.map((entry, index) => (
                        <Cell
                          key={`cell-${index}`}
                          fill={COLORS[index % COLORS.length]}
                        />
                      ))}
                    </Pie>
                    <Tooltip formatter={(value: any) => `$${(typeof value === 'number' ? value : parseFloat(value)).toFixed(2)}`} />
                  </PieChart>
                </ResponsiveContainer>
              ) : (
                <Typography color="textSecondary">Sin datos</Typography>
              )}
            </Paper>
          </Grid>

          {/* Type Summary */}
          <Grid item xs={12}>
            <Paper elevation={2} sx={{ p: 2 }}>
              <Typography variant="h6" gutterBottom sx={{ fontWeight: 'bold' }}>
                💰 Por Tipo
              </Typography>
              {summary && Object.keys(summary.by_type).length > 0 ? (
                <Box sx={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))', gap: 2 }}>
                  {Object.entries(summary.by_type).map(([type, data]: any) => (
                    <Card key={type}>
                      <CardContent>
                        <Typography variant="body2" color="textSecondary">
                          {type}
                        </Typography>
                        <Typography variant="h6">
                          ${data.total.toFixed(2)}
                        </Typography>
                        <Typography variant="body2" color="textSecondary">
                          {data.count} transacciones
                        </Typography>
                      </CardContent>
                    </Card>
                  ))}
                </Box>
              ) : (
                <Typography color="textSecondary">Sin datos</Typography>
              )}
            </Paper>
          </Grid>
        </Grid>
      </Box>
    </Container>
  );
}
