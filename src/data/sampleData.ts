import { MonthlyFinanceData, DailyTransaction, KPIData } from '@/types/finance';

// 전역 데이터 저장소 (메모리)
let globalMonthlyReports: MonthlyFinanceData[] = [
  {
    year: 2024,
    month: 1,
    salesRevenue: 244737761,
    otherIncome: 0,
    totalRevenue: 244737761,
    rentExpense: 34286046,
    laborExpense: 34746632,
    materialExpense: 30955171,
    operatingExpense: 32084635,
    otherExpense: 30734362,
    totalExpense: 162806846,
    netIncome: 81930915,
    profitMargin: 33.5,
    cashBalance: 6356416405,
    cashFlowChange: 81930915,
    approvalStatus: 'approved',
    approvedAt: '2024-02-05T09:00:00Z',
    approvedBy: '관리자',
    memo: '1월 실적 승인완료',
    createdAt: '2024-01-31T18:00:00Z',
    updatedAt: '2024-02-05T09:00:00Z'
  },
  {
    year: 2024,
    month: 2,
    salesRevenue: 220000000,
    otherIncome: 5000000,
    totalRevenue: 225000000,
    rentExpense: 30000000,
    laborExpense: 35000000,
    materialExpense: 28000000,
    operatingExpense: 30000000,
    otherExpense: 25000000,
    totalExpense: 148000000,
    netIncome: 77000000,
    profitMargin: 34.2,
    cashBalance: 6433416405,
    cashFlowChange: 77000000,
    approvalStatus: 'approved',
    approvedAt: '2024-03-05T09:00:00Z',
    approvedBy: '관리자',
    memo: '2월 실적 승인완료',
    createdAt: '2024-02-29T18:00:00Z',
    updatedAt: '2024-03-05T09:00:00Z'
  },
  {
    year: 2024,
    month: 3,
    salesRevenue: 280000000,
    otherIncome: 3000000,
    totalRevenue: 283000000,
    rentExpense: 32000000,
    laborExpense: 38000000,
    materialExpense: 35000000,
    operatingExpense: 33000000,
    otherExpense: 28000000,
    totalExpense: 166000000,
    netIncome: 117000000,
    profitMargin: 41.3,
    cashBalance: 6550416405,
    cashFlowChange: 117000000,
    approvalStatus: 'pending',
    memo: '',
    createdAt: '2024-03-31T18:00:00Z',
    updatedAt: '2024-03-31T18:00:00Z'
  }
];

// 데이터 관리 함수들
export const getMonthlyReports = (): MonthlyFinanceData[] => {
  return [...globalMonthlyReports]; // 복사본 반환
};

export const getMonthlyReport = (year: number, month: number): MonthlyFinanceData | undefined => {
  return globalMonthlyReports.find(r => r.year === year && r.month === month);
};

export const addMonthlyReport = (report: MonthlyFinanceData): MonthlyFinanceData => {
  const existingIndex = globalMonthlyReports.findIndex(
    r => r.year === report.year && r.month === report.month
  );

  if (existingIndex >= 0) {
    // 기존 데이터 업데이트
    globalMonthlyReports[existingIndex] = report;
    return report;
  } else {
    // 새 데이터 추가
    globalMonthlyReports.push(report);
    return report;
  }
};

export const updateMonthlyReport = (year: number, month: number, updates: Partial<MonthlyFinanceData>): MonthlyFinanceData | null => {
  const index = globalMonthlyReports.findIndex(r => r.year === year && r.month === month);
  
  if (index >= 0) {
    globalMonthlyReports[index] = { ...globalMonthlyReports[index], ...updates };
    return globalMonthlyReports[index];
  }
  
  return null;
};

// 2024년 월별 샘플 데이터 (하위 호환성)
export const monthlyData: MonthlyFinanceData[] = globalMonthlyReports;

// monthlyReports로도 내보내기 (하위 호환성)
export const monthlyReports = globalMonthlyReports;

// 샘플 일별 거래 데이터 (3월)
export const dailyTransactions: DailyTransaction[] = [
  {
    id: '20240301-001',
    date: '2024-03-01',
    description: '매출입금',
    amount: 4200000,
    type: 'income',
    category: '매출',
    detail: '고객 A 서비스 대금',
    customer: '고객 A'
  },
  {
    id: '20240301-002',
    date: '2024-03-01',
    description: '임대료 지출',
    amount: -1500000,
    type: 'expense',
    category: '임대료',
    detail: '사무실 임대료',
    vendor: '임대업체'
  },
  {
    id: '20240302-001',
    date: '2024-03-02',
    description: '매출입금',
    amount: 3800000,
    type: 'income',
    category: '매출',
    detail: '고객 B 서비스 대금',
    customer: '고객 B'
  },
  {
    id: '20240302-002',
    date: '2024-03-02',
    description: '인건비 지출',
    amount: -2200000,
    type: 'expense',
    category: '인건비',
    detail: '직원 급여',
    vendor: '급여'
  },
  {
    id: '20240303-001',
    date: '2024-03-03',
    description: '재료비 지출',
    amount: -800000,
    type: 'expense',
    category: '재료비',
    detail: '원자재 구매',
    vendor: '자재업체'
  }
];

// KPI 샘플 데이터
export const kpiData: KPIData = {
  totalRevenue: 749737761,
  totalExpense: 476806846,
  currentCashBalance: 6550416405,
  profitMargin: 36.4,
  revenueChange: 25.6,
  expenseChange: 12.1,
  cashBalanceChange: 3.1,
  profitMarginChange: 20.7
};

// 차트용 데이터 변환 함수
export const getChartData = () => {
  return monthlyData.map(data => ({
    period: `${data.year}-${data.month < 10 ? '0' + data.month : data.month}`,
    revenue: data.totalRevenue,
    expense: data.totalExpense,
    netIncome: data.netIncome
  }));
};

// 카테고리별 지출 데이터
export const getExpenseByCategory = (month: number, year: number = 2024) => {
  const data = monthlyData.find(d => d.year === year && d.month === month);
  if (!data) return [];

  const total = data.totalExpense;
  return [
    { category: '임대료', amount: data.rentExpense, percentage: (data.rentExpense / total) * 100 },
    { category: '인건비', amount: data.laborExpense, percentage: (data.laborExpense / total) * 100 },
    { category: '재료비', amount: data.materialExpense, percentage: (data.materialExpense / total) * 100 },
    { category: '운영비', amount: data.operatingExpense, percentage: (data.operatingExpense / total) * 100 },
    { category: '기타', amount: data.otherExpense, percentage: (data.otherExpense / total) * 100 }
  ];
}; 