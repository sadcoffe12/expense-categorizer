import React, { useState, useEffect } from 'react';
import { BrowserRouter, Routes, Route, Navigate, Link } from 'react-router-dom';
import { ThemeProvider, createTheme } from '@mui/material/styles';
import CssBaseline from '@mui/material/CssBaseline';
import {
  AppBar,
  Toolbar,
  Typography,
  Button,
  Container,
  Box,
  CircularProgress,
} from '@mui/material';
import SetupPage from './pages/SetupPage';
import DashboardPage from './pages/DashboardPage';
import ExpensesPage from './pages/ExpensesPage';
import SettingsPage from './pages/SettingsPage';
import ReportsPage from './pages/ReportsPage';
import { configAPI } from './services/api';

const theme = createTheme({
  palette: {
    primary: {
      main: '#1976d2',
    },
    secondary: {
      main: '#dc004e',
    },
  },
});

function AppLayout({ children }: { children: React.ReactNode }) {
  return (
    <Box>
      <AppBar position="static">
        <Toolbar>
          <Typography variant="h6" sx={{ flexGrow: 1, fontWeight: 'bold' }}>
            💰 Expense Categorizer
          </Typography>
          <Button color="inherit" component={Link} to="/">
            Dashboard
          </Button>
          <Button color="inherit" component={Link} to="/expenses">
            Gastos
          </Button>
          <Button color="inherit" component={Link} to="/reports">
            Reports
          </Button>
          <Button color="inherit" component={Link} to="/settings">
            Settings
          </Button>
        </Toolbar>
      </AppBar>
      {children}
    </Box>
  );
}

function App() {
  const [isConfigured, setIsConfigured] = useState<boolean | null>(null);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    const checkConfiguration = async () => {
      try {
        const status = await configAPI.getStatus();
        setIsConfigured(status.configured);
      } catch (error) {
        console.error('Error checking configuration:', error);
        setIsConfigured(false);
      } finally {
        setLoading(false);
      }
    };

    checkConfiguration();
  }, []);

  if (loading) {
    return (
      <Container maxWidth="lg" sx={{ py: 4, textAlign: 'center' }}>
        <CircularProgress />
        <Typography variant="body2" sx={{ mt: 2 }}>
          Inicializando...
        </Typography>
      </Container>
    );
  }

  return (
    <ThemeProvider theme={theme}>
      <CssBaseline />
      <BrowserRouter>
        {!isConfigured ? (
          <SetupPage
            onSetupComplete={() => {
              setIsConfigured(true);
            }}
          />
        ) : (
          <AppLayout>
            <Routes>
              <Route path="/" element={<DashboardPage />} />
              <Route path="/expenses" element={<ExpensesPage />} />
              <Route path="/reports" element={<ReportsPage />} />
              <Route path="/settings" element={<SettingsPage />} />
              <Route path="*" element={<Navigate to="/" />} />
            </Routes>
          </AppLayout>
        )}
      </BrowserRouter>
    </ThemeProvider>
  );
}

export default App;
