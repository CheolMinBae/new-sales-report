'use client';

import { useState } from 'react';
import {
  Box,
  Container,
  Typography,
  Grid,
  Card,
  CardContent,
  Button,
  AppBar,
  Toolbar,
  Drawer,
  List,
  ListItem,
  ListItemIcon,
  ListItemText,
  ListItemButton,
  Divider,
  ThemeProvider,
  createTheme,
  CssBaseline,
  FormControl,
  InputLabel,
  Select,
  MenuItem
} from '@mui/material';
import {
  Dashboard,
  Receipt,
  TrendingUp,
  TrendingDown,
  AttachMoney,
  AccountBalance
} from '@mui/icons-material';
import { kpiData } from '@/data/sampleData';
import RevenueChart from '@/components/charts/RevenueChart';
import ExpenseChart from '@/components/charts/ExpenseChart';
import MonthlyStatsChart from '@/components/charts/MonthlyStatsChart';
import ProfitStructureChart from '@/components/charts/ProfitStructureChart';
import WeeklySalesChart from '@/components/charts/WeeklySalesChart';
import MonthlyReportManagement from '@/components/MonthlyReportManagement';

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

const drawerWidth = 240;

interface KPICardProps {
  title: string;
  value: string;
  change: number;
  icon: React.ReactNode;
  color: 'primary' | 'secondary' | 'success' | 'warning';
}

function KPICard({ title, value, change, icon, color }: KPICardProps) {
  const isPositive = change >= 0;
  
  return (
    <Card sx={{ height: '100%' }}>
      <CardContent>
        <Box display="flex" alignItems="center" justifyContent="space-between">
          <Box>
            <Typography color="textSecondary" variant="h6" component="div">
              {title}
            </Typography>
            <Typography variant="h4" component="div" fontWeight="bold">
              {value}
            </Typography>
            <Box display="flex" alignItems="center" mt={1}>
              {isPositive ? (
                <TrendingUp color="success" fontSize="small" />
              ) : (
                <TrendingDown color="error" fontSize="small" />
              )}
              <Typography 
                variant="body2" 
                color={isPositive ? 'success.main' : 'error.main'}
                ml={0.5}
              >
                {isPositive ? '+' : ''}{change.toFixed(1)}%
              </Typography>
            </Box>
          </Box>
          <Box color={`${color}.main`}>
            {icon}
          </Box>
        </Box>
      </CardContent>
    </Card>
  );
}

export default function FinanceDashboard() {
  const [selectedMenu, setSelectedMenu] = useState('dashboard');
  const [selectedYear, setSelectedYear] = useState(2024);

  const formatCurrency = (amount: number) => {
    return new Intl.NumberFormat('ko-KR', {
      style: 'currency',
      currency: 'KRW',
      minimumFractionDigits: 0,
      maximumFractionDigits: 0
    }).format(amount);
  };

  const menuItems = [
    { id: 'dashboard', label: '전체 누적 대시보드', icon: <Dashboard /> },
    { id: 'monthly', label: '월별 레포트 관리', icon: <Receipt /> },
  ];

  return (
    <ThemeProvider theme={theme}>
      <CssBaseline />
      <Box sx={{ display: 'flex' }}>
        {/* AppBar */}
        <AppBar 
          position="fixed" 
          sx={{ 
            width: `calc(100% - ${drawerWidth}px)`, 
            ml: `${drawerWidth}px` 
          }}
        >
          <Toolbar>
            <Typography variant="h6" noWrap component="div">
              재무 리포트 대시보드
            </Typography>
          </Toolbar>
        </AppBar>

        {/* Sidebar */}
        <Drawer
          sx={{
            width: drawerWidth,
            flexShrink: 0,
            '& .MuiDrawer-paper': {
              width: drawerWidth,
              boxSizing: 'border-box',
            },
          }}
          variant="permanent"
          anchor="left"
        >
          <Toolbar>
            <Typography variant="h6" noWrap component="div" fontWeight="bold">
              Finance
            </Typography>
          </Toolbar>
          <Divider />
          <List>
            {menuItems.map((item) => (
              <ListItem key={item.id} disablePadding>
                <ListItemButton
                  selected={selectedMenu === item.id}
                  onClick={() => setSelectedMenu(item.id)}
                >
                  <ListItemIcon>
                    {item.icon}
                  </ListItemIcon>
                  <ListItemText primary={item.label} />
                </ListItemButton>
              </ListItem>
            ))}
          </List>
        </Drawer>

        {/* Main Content */}
        <Box
          component="main"
          sx={{
            flexGrow: 1,
            bgcolor: 'background.default',
            p: 3,
            mt: 8
          }}
        >
          {selectedMenu === 'dashboard' && (
            <Container maxWidth="xl">
              <Typography variant="h4" gutterBottom fontWeight="bold">
                전체 누적 대시보드
              </Typography>
              
              {/* 년도 선택 */}
              <Grid container spacing={3} mb={4}>
                <Grid item xs={12} md={3}>
                  <FormControl fullWidth>
                    <InputLabel>년도 선택</InputLabel>
                    <Select
                      value={selectedYear}
                      label="년도 선택"
                      onChange={(e) => setSelectedYear(Number(e.target.value))}
                    >
                      {Array.from({ length: 6 }, (_, i) => 2020 + i).map((year) => (
                        <MenuItem key={year} value={year}>
                          {year}년
                        </MenuItem>
                      ))}
                    </Select>
                  </FormControl>
                </Grid>
              </Grid>
              
              {/* KPI Cards */}
              <Grid container spacing={3} mb={4}>
                <Grid item xs={12} sm={6} md={3}>
                  <KPICard
                    title="총 매출"
                    value={formatCurrency(kpiData.totalRevenue)}
                    change={kpiData.revenueChange}
                    icon={<AttachMoney fontSize="large" />}
                    color="primary"
                  />
                </Grid>
                <Grid item xs={12} sm={6} md={3}>
                  <KPICard
                    title="총 매입"
                    value={formatCurrency(kpiData.totalExpense)}
                    change={kpiData.expenseChange}
                    icon={<Receipt fontSize="large" />}
                    color="secondary"
                  />
                </Grid>
                <Grid item xs={12} sm={6} md={3}>
                  <KPICard
                    title="현금 잔고"
                    value={formatCurrency(kpiData.currentCashBalance)}
                    change={kpiData.cashBalanceChange}
                    icon={<AccountBalance fontSize="large" />}
                    color="success"
                  />
                </Grid>
                <Grid item xs={12} sm={6} md={3}>
                  <KPICard
                    title="순이익률"
                    value={`${kpiData.profitMargin.toFixed(1)}%`}
                    change={kpiData.profitMarginChange}
                    icon={<TrendingUp fontSize="large" />}
                    color="warning"
                  />
                </Grid>
              </Grid>

              {/* Charts Section */}
              <Grid container spacing={3}>
                <Grid item xs={12} md={8}>
                  <Card>
                    <CardContent>
                      <Typography variant="h6" gutterBottom>
                        월별 매출 현황
                      </Typography>
                      <RevenueChart />
                    </CardContent>
                  </Card>
                </Grid>
                <Grid item xs={12} md={4}>
                  <Card>
                    <CardContent>
                      <Typography variant="h6" gutterBottom>
                        카테고리별 지출 (3월)
                      </Typography>
                      <ExpenseChart />
                    </CardContent>
                  </Card>
                </Grid>

                {/* 두 번째 줄 차트들 */}
                <Grid item xs={12} md={6}>
                  <Card>
                    <CardContent>
                      <Typography variant="h6" gutterBottom>
                        월간 통계 (Area Chart)
                      </Typography>
                      <MonthlyStatsChart />
                    </CardContent>
                  </Card>
                </Grid>
                <Grid item xs={12} md={6}>
                  <Card>
                    <CardContent>
                      <Typography variant="h6" gutterBottom>
                        주간 매출 현황
                      </Typography>
                      <WeeklySalesChart />
                    </CardContent>
                  </Card>
                </Grid>

                {/* 세 번째 줄 - 수익 구조 분석 */}
                <Grid item xs={12}>
                  <Card>
                    <CardContent>
                      <Typography variant="h6" gutterBottom>
                        수익 구조 분석
                      </Typography>
                      <ProfitStructureChart />
                    </CardContent>
                  </Card>
                </Grid>
              </Grid>
            </Container>
          )}

          {selectedMenu === 'monthly' && (
            <MonthlyReportManagement />
          )}
        </Box>
      </Box>
    </ThemeProvider>
  );
} 