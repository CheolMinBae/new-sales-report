'use client';

import { useState, useEffect } from 'react';
import {
  Box,
  Container,
  Typography,
  Grid,
  Card,
  CardContent,
  Select,
  MenuItem,
  FormControl,
  InputLabel,
  Chip,
  Button,
  Modal,
  TextField,
  Alert,
  Divider,
  CircularProgress,
  Snackbar
} from '@mui/material';
import {
  CheckCircle,
  Cancel,
  Pending,
  TrendingUp,
  TrendingDown,
  AttachMoney,
  Receipt,
  AccountBalance
} from '@mui/icons-material';
import { MonthlyFinanceData, ApprovalStatus } from '@/types/finance';
import { getAllReports, updateReportApproval } from '@/services/api';
import RevenueChart from '@/components/charts/RevenueChart';
import ExpenseChart from '@/components/charts/ExpenseChart';
import WeeklySalesChart from '@/components/charts/WeeklySalesChart';
import ProfitStructureChart from '@/components/charts/ProfitStructureChart';

interface ApprovalModalProps {
  open: boolean;
  onClose: () => void;
  onConfirm: (memo: string) => void;
  action: 'approve' | 'reject';
  monthData: MonthlyFinanceData | null;
}

function ApprovalModal({ open, onClose, onConfirm, action, monthData }: ApprovalModalProps) {
  const [memo, setMemo] = useState('');

  const handleConfirm = () => {
    onConfirm(memo);
    setMemo('');
    onClose();
  };

  return (
    <Modal open={open} onClose={onClose}>
      <Box
        sx={{
          position: 'absolute',
          top: '50%',
          left: '50%',
          transform: 'translate(-50%, -50%)',
          width: 400,
          bgcolor: 'background.paper',
          borderRadius: 2,
          boxShadow: 24,
          p: 4,
        }}
      >
        <Typography variant="h6" component="h2" mb={2}>
          {action === 'approve' ? '승인' : '반려'} 확인
        </Typography>
        
        <Alert 
          severity={action === 'approve' ? 'info' : 'warning'} 
          sx={{ mb: 2 }}
        >
          {monthData && `${monthData.year}년 ${monthData.month}월 리포트를 ${action === 'approve' ? '승인' : '반려'}하시겠습니까?`}
        </Alert>

        <TextField
          fullWidth
          multiline
          rows={4}
          label="메모 (선택사항)"
          value={memo}
          onChange={(e) => setMemo(e.target.value)}
          sx={{ mb: 3 }}
        />

        <Box sx={{ display: 'flex', gap: 2, justifyContent: 'flex-end' }}>
          <Button variant="outlined" onClick={onClose}>
            취소
          </Button>
          <Button
            variant="contained"
            color={action === 'approve' ? 'primary' : 'error'}
            onClick={handleConfirm}
          >
            {action === 'approve' ? '승인' : '반려'}
          </Button>
        </Box>
      </Box>
    </Modal>
  );
}

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
                전월 대비 {isPositive ? '+' : ''}{change.toFixed(1)}%
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

export default function MonthlyReportManagement() {
  const [selectedYear, setSelectedYear] = useState(2024);
  const [selectedMonth, setSelectedMonth] = useState(6);
  const [modalOpen, setModalOpen] = useState(false);
  const [modalAction, setModalAction] = useState<'approve' | 'reject'>('approve');
  const [monthlyData, setMonthlyData] = useState<MonthlyFinanceData[]>([]);
  const [loading, setLoading] = useState(true);
  const [dataRefreshing, setDataRefreshing] = useState(false);
  const [snackbar, setSnackbar] = useState({ open: false, message: '', severity: 'success' as 'success' | 'error' });

  // 초기 데이터 로드
  useEffect(() => {
    const loadInitialData = async () => {
      try {
        setLoading(true);
        const data = await getAllReports();
        setMonthlyData(data);
      } catch (error) {
        console.error('초기 데이터 로드 실패:', error);
        setSnackbar({
          open: true,
          message: '데이터를 불러오는데 실패했습니다.',
          severity: 'error'
        });
      } finally {
        setLoading(false);
      }
    };

    loadInitialData();
  }, []); // 처음 마운트 시에만 실행

  // 년도/월 변경 시 데이터 새로고침
  useEffect(() => {
    const refreshData = async () => {
      try {
        setDataRefreshing(true);
        
        // 전체 데이터 및 선택된 년도/월 데이터 병렬 로드
        const [allData, specificData] = await Promise.all([
          getAllReports(),
          // 특정 년도/월 데이터도 시도 (존재하지 않아도 에러 없이 처리)
          fetch(`/api/reports/${selectedMonth}?year=${selectedYear}`)
            .then(res => res.ok ? res.json() : null)
            .catch(() => null)
        ]);
        
        setMonthlyData(allData);
        
        // 선택된 년도/월 데이터 로드 성공 시 콘솔에 표시
        if (specificData) {
          console.log(`✅ ${selectedYear}년 ${selectedMonth}월 데이터 로드 완료:`, specificData);
          setSnackbar({
            open: true,
            message: `${selectedYear}년 ${selectedMonth}월 데이터를 새로고침했습니다.`,
            severity: 'success'
          });
        } else {
          console.log(`ℹ️ ${selectedYear}년 ${selectedMonth}월 데이터 없음`);
        }
        
      } catch (error) {
        console.error('데이터 새로고침 실패:', error);
        setSnackbar({
          open: true,
          message: '데이터 새로고침에 실패했습니다.',
          severity: 'error'
        });
      } finally {
        setDataRefreshing(false);
      }
    };

    // 초기 로딩이 완료된 후에만 새로고침 실행
    if (!loading) {
      refreshData();
    }
  }, [selectedYear, selectedMonth, loading]); // 년도/월 변경 시 데이터 새로고침

  const formatCurrency = (amount: number) => {
    return new Intl.NumberFormat('ko-KR', {
      style: 'currency',
      currency: 'KRW',
      minimumFractionDigits: 0,
      maximumFractionDigits: 0
    }).format(amount);
  };

  // 현재 선택된 월의 데이터
  const currentData = monthlyData.find(data => 
    data.year === selectedYear && data.month === selectedMonth
  );

  // 이전 월 데이터 (변화율 계산용)
  const previousData = monthlyData.find(data => {
    const prevMonth = selectedMonth === 1 ? 12 : selectedMonth - 1;
    const prevYear = selectedMonth === 1 ? selectedYear - 1 : selectedYear;
    return data.year === prevYear && data.month === prevMonth;
  });

  const getStatusIcon = (status: ApprovalStatus) => {
    switch (status) {
      case 'approved':
        return <CheckCircle color="success" />;
      case 'rejected':
        return <Cancel color="error" />;
      default:
        return <Pending color="warning" />;
    }
  };

  const getStatusColor = (status: ApprovalStatus) => {
    switch (status) {
      case 'approved':
        return 'success';
      case 'rejected':
        return 'error';
      default:
        return 'warning';
    }
  };

  const getStatusText = (status: ApprovalStatus) => {
    switch (status) {
      case 'approved':
        return '승인완료';
      case 'rejected':
        return '반려';
      default:
        return '승인대기';
    }
  };

  const calculateChange = (current: number, previous: number) => {
    if (!previous) return 0;
    return ((current - previous) / previous) * 100;
  };

  const handleApproval = (action: 'approve' | 'reject') => {
    setModalAction(action);
    setModalOpen(true);
  };

  const handleConfirmApproval = async (memo: string) => {
    if (!currentData) return;

    try {
      const approvalStatus = modalAction === 'approve' ? 'approved' : 'rejected';
      const updatedData = await updateReportApproval(
        currentData.month,
        approvalStatus,
        memo,
        '관리자'
      );

      // 로컬 상태 업데이트
      setMonthlyData(prev => 
        prev.map(item => 
          item.month === currentData.month && item.year === currentData.year 
            ? updatedData 
            : item
        )
      );

      setSnackbar({
        open: true,
        message: `${currentData.month}월 레포트가 ${modalAction === 'approve' ? '승인' : '반려'}되었습니다.`,
        severity: 'success'
      });

    } catch (error) {
      console.error('승인/반려 처리 실패:', error);
      setSnackbar({
        open: true,
        message: '처리 중 오류가 발생했습니다.',
        severity: 'error'
      });
    }
  };

  if (loading) {
    return (
      <Container maxWidth="xl">
        <Box display="flex" justifyContent="center" alignItems="center" minHeight="200px">
          <CircularProgress />
        </Box>
      </Container>
    );
  }

  if (!currentData) {
    const year = selectedYear;
    const month = selectedMonth;
    return (
      <Container maxWidth="xl">
        <Typography variant="h4" gutterBottom fontWeight="bold">
          월별 레포트 관리
        </Typography>
        
        {/* 년도/월 선택 */}
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
          <Grid item xs={12} md={3}>
            <FormControl fullWidth>
              <InputLabel>월 선택</InputLabel>
              <Select
                value={selectedMonth}
                label="월 선택"
                onChange={(e) => setSelectedMonth(Number(e.target.value))}
              >
                {Array.from({ length: 12 }, (_, i) => i + 1).map((month) => {
                  const hasData = monthlyData.some(data => 
                    data.year === selectedYear && data.month === month
                  );
                  return (
                    <MenuItem 
                      key={month}
                      value={month}
                    >
                      {month}월 {hasData ? '📊' : '📝'}
                    </MenuItem>
                  );
                })}
              </Select>
            </FormControl>
          </Grid>
        </Grid>

        <Alert severity="info">
          {year}년 {month}월의 데이터가 아직 전송되지 않았습니다. 
          <br />
          엑셀 VBA에서 데이터를 전송하거나, 아래 버튼으로 테스트 데이터를 생성하세요.
        </Alert>
        
        <Box mt={3}>
          <Button
            variant="contained"
            color="primary"
            onClick={async () => {
              try {
                const response = await fetch('/api/reports/submit', {
                  method: 'POST',
                  headers: { 'Content-Type': 'application/json' },
                  body: JSON.stringify({
                    year,
                    month,
                    salesRevenue: 10000000,
                    otherIncome: 500000,
                    rentExpense: 2000000,
                    laborExpense: 3000000,
                    materialExpense: 1500000,
                    operatingExpense: 1000000,
                    otherExpense: 500000,
                    cashBalance: 50000000,
                    submittedBy: 'Web Dashboard'
                  })
                });
                
                if (response.ok) {
                  setSnackbar({
                    open: true,
                    message: `${month}월 테스트 데이터가 생성되었습니다.`,
                    severity: 'success'
                  });
                  // 데이터 새로고침
                  const data = await getAllReports();
                  setMonthlyData(data);
                } else {
                  throw new Error('데이터 생성 실패');
                }
              } catch (error) {
                setSnackbar({
                  open: true,
                  message: '테스트 데이터 생성에 실패했습니다.',
                  severity: 'error'
                });
              }
            }}
          >
            📝 {month}월 테스트 데이터 생성
          </Button>
        </Box>
      </Container>
    );
  }

  const changes = previousData ? {
    revenue: calculateChange(currentData.totalRevenue, previousData.totalRevenue),
    expense: calculateChange(currentData.totalExpense, previousData.totalExpense),
    cash: calculateChange(currentData.cashBalance, previousData.cashBalance),
    profit: calculateChange(currentData.profitMargin, previousData.profitMargin)
  } : { revenue: 0, expense: 0, cash: 0, profit: 0 };

  return (
    <Container maxWidth="xl">
      <Typography variant="h4" gutterBottom fontWeight="bold">
        월별 레포트 관리
      </Typography>

      {/* 년도/월 선택 및 승인 상태 */}
      <Grid container spacing={3} mb={4}>
        <Grid item xs={12} md={2}>
          <FormControl fullWidth>
            <InputLabel>년도 선택</InputLabel>
            <Select
              value={selectedYear}
              label="년도 선택"
              onChange={(e) => setSelectedYear(Number(e.target.value))}
              disabled={dataRefreshing}
            >
              {Array.from({ length: 6 }, (_, i) => 2020 + i).map((year) => (
                <MenuItem key={year} value={year}>
                  {year}년
                </MenuItem>
              ))}
            </Select>
          </FormControl>
        </Grid>
        <Grid item xs={12} md={2}>
          <FormControl fullWidth>
            <InputLabel>월 선택</InputLabel>
            <Select
              value={selectedMonth}
              label="월 선택"
              onChange={(e) => setSelectedMonth(Number(e.target.value))}
              disabled={dataRefreshing}
            >
              {Array.from({ length: 12 }, (_, i) => i + 1).map((month) => {
                const hasData = monthlyData.some(data => 
                  data.year === selectedYear && data.month === month
                );
                return (
                  <MenuItem 
                    key={month}
                    value={month}
                  >
                    {month}월 {hasData ? '📊' : '📝'}
                  </MenuItem>
                );
              })}
            </Select>
          </FormControl>
        </Grid>
        {dataRefreshing && (
          <Grid item xs={12} md={2}>
            <Box display="flex" alignItems="center" gap={1}>
              <CircularProgress size={20} />
              <Typography variant="body2" color="primary">
                데이터 새로고침 중...
              </Typography>
            </Box>
          </Grid>
        )}

        <Grid item xs={12} md={4}>
          <Box display="flex" alignItems="center" gap={2} height="100%">
            <Typography variant="h6">승인 상태:</Typography>
            <Chip
              icon={getStatusIcon(currentData.approvalStatus)}
              label={getStatusText(currentData.approvalStatus)}
              color={getStatusColor(currentData.approvalStatus)}
              variant="outlined"
            />
          </Box>
        </Grid>

        <Grid item xs={12} md={4}>
          {currentData.approvalStatus === 'pending' && (
            <Box display="flex" gap={2} height="100%">
              <Button
                variant="contained"
                color="primary"
                startIcon={<CheckCircle />}
                onClick={() => handleApproval('approve')}
              >
                승인
              </Button>
              <Button
                variant="contained"
                color="error"
                startIcon={<Cancel />}
                onClick={() => handleApproval('reject')}
              >
                반려
              </Button>
            </Box>
          )}
        </Grid>
      </Grid>

      {/* 이달의 KPI */}
      <Typography variant="h5" gutterBottom fontWeight="bold" mb={2}>
        이달의 주요 지표
      </Typography>
      
      <Grid container spacing={3} mb={4}>
        <Grid item xs={12} sm={6} md={3}>
          <KPICard
            title="이달의 매출"
            value={formatCurrency(currentData.totalRevenue)}
            change={changes.revenue}
            icon={<AttachMoney fontSize="large" />}
            color="primary"
          />
        </Grid>
        <Grid item xs={12} sm={6} md={3}>
          <KPICard
            title="이달의 매입"
            value={formatCurrency(currentData.totalExpense)}
            change={changes.expense}
            icon={<Receipt fontSize="large" />}
            color="secondary"
          />
        </Grid>
        <Grid item xs={12} sm={6} md={3}>
          <KPICard
            title="현금 잔고"
            value={formatCurrency(currentData.cashBalance)}
            change={changes.cash}
            icon={<AccountBalance fontSize="large" />}
            color="success"
          />
        </Grid>
        <Grid item xs={12} sm={6} md={3}>
          <KPICard
            title="순이익률"
            value={`${currentData.profitMargin.toFixed(1)}%`}
            change={changes.profit}
            icon={<TrendingUp fontSize="large" />}
            color="warning"
          />
        </Grid>
      </Grid>

      <Divider sx={{ my: 4 }} />

      {/* 차트 분석 */}
      <Typography variant="h5" gutterBottom fontWeight="bold" mb={2}>
        월별 분석 차트
      </Typography>

      <Grid container spacing={3}>
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
        
        <Grid item xs={12} md={6}>
          <Card>
            <CardContent>
              <Typography variant="h6" gutterBottom>
                카테고리별 지출
              </Typography>
              <ExpenseChart />
            </CardContent>
          </Card>
        </Grid>

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

      {/* 승인/반려 모달 */}
      <ApprovalModal
        open={modalOpen}
        onClose={() => setModalOpen(false)}
        onConfirm={handleConfirmApproval}
        action={modalAction}
        monthData={currentData}
      />

      {/* 알림 스낵바 */}
      <Snackbar
        open={snackbar.open}
        autoHideDuration={6000}
        onClose={() => setSnackbar(prev => ({ ...prev, open: false }))}
      >
        <Alert 
          onClose={() => setSnackbar(prev => ({ ...prev, open: false }))} 
          severity={snackbar.severity}
          sx={{ width: '100%' }}
        >
          {snackbar.message}
        </Alert>
      </Snackbar>
    </Container>
  );
} 