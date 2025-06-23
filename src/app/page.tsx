'use client';

import { useState, useEffect } from 'react';
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
  MenuItem,
  CircularProgress
} from '@mui/material';
import {
  Dashboard,
  Receipt,
  TrendingUp,
  TrendingDown,
  AttachMoney,
  AccountBalance
} from '@mui/icons-material';
import RevenueChart from '@/components/charts/RevenueChart';
import ExpenseChart from '@/components/charts/ExpenseChart';
import MonthlyStatsChart from '@/components/charts/MonthlyStatsChart';
import ProfitStructureChart from '@/components/charts/ProfitStructureChart';
import WeeklySalesChart from '@/components/charts/WeeklySalesChart';
import CashFlowChart from '@/components/charts/CashFlowChart';
import FixedVariableGauge from '@/components/charts/FixedVariableGauge';
import WaterfallChart from '@/components/charts/WaterfallChart';
import MonthlyDetailTable from '@/components/MonthlyDetailTable';
import MonthlyReportManagement from '@/components/MonthlyReportManagement';
import { IncomeBreakdownChart } from '@/components/charts/IncomeBreakdownChart';

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

interface DashboardData {
  year: number;
  kpi: {
    totalRevenue: number;
    totalExpense: number;
    totalNetIncome: number;
    avgProfitMargin: number;
    currentCashBalance: number;
  };
  chartData: any;
  expenseData: any[];
  weeklySalesData: any[];
  monthlyReports: any[];
}

export default function FinanceDashboard() {
  const [selectedMenu, setSelectedMenu] = useState('dashboard');
  const [selectedYear, setSelectedYear] = useState(2024);
  const [dashboardData, setDashboardData] = useState<DashboardData | null>(null);
  const [loading, setLoading] = useState(false);

  const formatCurrency = (amount: number) => {
    return new Intl.NumberFormat('ko-KR', {
      style: 'currency',
      currency: 'KRW',
      minimumFractionDigits: 0,
      maximumFractionDigits: 0
    }).format(amount);
  };

  // API에서 대시보드 데이터 가져오기
  const fetchDashboardData = async (year: number) => {
    setLoading(true);
    try {
      const response = await fetch(`/api/dashboard?year=${year}`);
      const result = await response.json();
      
      if (result.success) {
        setDashboardData(result.data);
      } else {
        console.error('대시보드 데이터 로드 실패:', result.error);
      }
    } catch (error) {
      console.error('API 호출 오류:', error);
    } finally {
      setLoading(false);
    }
  };

  // 년도 변경 시 데이터 다시 로드
  useEffect(() => {
    if (selectedMenu === 'dashboard') {
      fetchDashboardData(selectedYear);
    }
  }, [selectedYear, selectedMenu]);

  // 초기 로드
  useEffect(() => {
    if (selectedMenu === 'dashboard') {
      fetchDashboardData(selectedYear);
    }
  }, []);

  const menuItems = [
    { id: 'dashboard', label: '전체 누적 대시보드', icon: <Dashboard /> },
    { id: 'monthly', label: '월별 레포트 관리', icon: <Receipt /> },
  ];

  // 차트 데이터 변환 함수 수정
  const getChartData = () => {
    return dashboardData?.monthlyReports.map(data => ({
      period: `${data.year}-${data.month < 10 ? '0' + data.month : data.month}`,
      revenue: data.totalRevenue,
      expense: data.totalExpense,
      netIncome: data.netIncome,
      cashSales: data.salesRevenue - data.creditSales, // 현금매출 = 총매출 - 외상매출
      creditSales: data.creditSales, // 외상매출
      otherIncome: data.otherIncome // 기타수입
    })) || [];
  };

  const getIncomeBreakdownData = () => {
    return dashboardData?.monthlyReports.map(data => ({
      month: `${data.month}월`,
      cashSales: data.salesRevenue - (data.creditSales || 0), // 현금매출 = 총매출 - 외상매출
      creditSales: data.creditSales || 0, // 외상매출
      otherIncome: data.otherIncome || 0 // 기타수입
    })) || [];
  };

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
              
              {/* Loading Indicator */}
              {loading && (
                <Box display="flex" justifyContent="center" my={4}>
                  <CircularProgress />
                </Box>
              )}

              {/* KPI Cards */}
              {dashboardData && !loading && (
                <Grid container spacing={3} mb={4}>
                  <Grid item xs={12} sm={6} md={3}>
                    <KPICard
                      title="총 매출"
                      value={formatCurrency(dashboardData.kpi.totalRevenue)}
                      change={5.2} // 임시 값 (나중에 이전 년도 대비 계산 로직 추가)
                      icon={<AttachMoney fontSize="large" />}
                      color="primary"
                    />
                  </Grid>
                  <Grid item xs={12} sm={6} md={3}>
                    <KPICard
                      title="총 매입"
                      value={formatCurrency(dashboardData.kpi.totalExpense)}
                      change={-2.1} // 임시 값
                      icon={<Receipt fontSize="large" />}
                      color="secondary"
                    />
                  </Grid>
                  <Grid item xs={12} sm={6} md={3}>
                    <KPICard
                      title="현금 잔고"
                      value={formatCurrency(dashboardData.kpi.currentCashBalance)}
                      change={8.7} // 임시 값
                      icon={<AccountBalance fontSize="large" />}
                      color="success"
                    />
                  </Grid>
                  <Grid item xs={12} sm={6} md={3}>
                    <KPICard
                      title="순이익률"
                      value={`${dashboardData.kpi.avgProfitMargin.toFixed(1)}%`}
                      change={1.5} // 임시 값
                      icon={<TrendingUp fontSize="large" />}
                      color="warning"
                    />
                  </Grid>
                </Grid>
              )}

              {/* Charts Section */}
              <Grid container spacing={3}>
                {/* 첫 번째 줄 - 현금흐름 & 고정비/유동비 게이지 */}
                <Grid item xs={12} md={8}>
                  <Card>
                    <CardContent>
                      <Typography variant="h6" gutterBottom>
                        📊 6개월간 현금흐름 현황
                      </Typography>
                      <CashFlowChart />
                    </CardContent>
                  </Card>
                </Grid>
                <Grid item xs={12} md={4}>
                  <Card>
                    <CardContent>
                      <Typography variant="h6" gutterBottom>
                        ⚖️ 고정비/유동비 게이지
                      </Typography>
                      <FixedVariableGauge />
                    </CardContent>
                  </Card>
                </Grid>

                {/* 두 번째 줄 - 폭포차트 */}
                <Grid item xs={12}>
                  <Card>
                    <CardContent>
                      <Typography variant="h6" gutterBottom>
                        💰 매출-매입-순이익 구조 분석 (폭포차트)
                      </Typography>
                      <WaterfallChart />
                    </CardContent>
                  </Card>
                </Grid>

                {/* 세 번째 줄 - 기존 차트들 */}
                {dashboardData && !loading && (
                  <>
                    <Grid item xs={12} md={8}>
                      <Card>
                        <CardContent>
                          <Typography variant="h6" gutterBottom>
                            {selectedYear}년 월별 매출 현황
                          </Typography>
                          <RevenueChart data={dashboardData.chartData.revenueData} />
                        </CardContent>
                      </Card>
                    </Grid>
                    <Grid item xs={12} md={4}>
                      <Card>
                        <CardContent>
                          <Typography variant="h6" gutterBottom>
                            {selectedYear}년 카테고리별 지출
                          </Typography>
                          <ExpenseChart data={dashboardData.expenseData} />
                        </CardContent>
                      </Card>
                    </Grid>
                  </>
                )}

                {/* 네 번째 줄 - 추가 차트들 */}
                {dashboardData && !loading && (
                  <>
                    <Grid item xs={12} md={6}>
                      <Card>
                        <CardContent>
                          <Typography variant="h6" gutterBottom>
                            {selectedYear}년 월간 통계
                          </Typography>
                          <MonthlyStatsChart data={dashboardData.chartData.monthlyStats} />
                        </CardContent>
                      </Card>
                    </Grid>
                  </>
                )}
                {dashboardData && !loading && (
                  <Grid item xs={12} md={6}>
                    <Card>
                      <CardContent>
                        <Typography variant="h6" gutterBottom>
                          {selectedYear}년 주간 매출 현황
                        </Typography>
                        <WeeklySalesChart />
                      </CardContent>
                    </Card>
                  </Grid>
                )}

                {/* 다섯 번째 줄 - 수익 구조 분석 */}
                {dashboardData && !loading && (
                  <Grid item xs={12}>
                    <Card>
                      <CardContent>
                        <Typography variant="h6" gutterBottom>
                          {selectedYear}년 수익 구조 분석
                        </Typography>
                        <ProfitStructureChart data={dashboardData.chartData.profitStructure} />
                      </CardContent>
                    </Card>
                  </Grid>
                )}

                {/* 여섯 번째 줄 - 월별 상세 테이블 */}
                {dashboardData && !loading && (
                  <Grid item xs={12}>
                    <Card>
                      <CardContent>
                        <Typography variant="h6" gutterBottom>
                          {selectedYear}년 월별 상세 데이터
                        </Typography>
                        <MonthlyDetailTable data={dashboardData.monthlyReports} />
                      </CardContent>
                    </Card>
                  </Grid>
                )}

                {/* 수입 구성 분석 차트 추가 */}
                <Grid item xs={12} md={6}>
                  <Card>
                    <CardContent>
                      <IncomeBreakdownChart data={getIncomeBreakdownData()} />
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