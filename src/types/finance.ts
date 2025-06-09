// 승인 상태
export type ApprovalStatus = 'pending' | 'approved' | 'rejected';

// 거래 카테고리
export type TransactionCategory = 
  | '임대료' 
  | '인건비' 
  | '재료비' 
  | '운영비' 
  | '기타';

// 월별 재무 데이터
export interface MonthlyFinanceData {
  year: number;
  month: number;
  // 매출 관련
  salesRevenue: number;
  otherIncome: number;
  totalRevenue: number;
  
  // 비용 관련
  rentExpense: number;
  laborExpense: number;
  materialExpense: number;
  operatingExpense: number;
  otherExpense: number;
  totalExpense: number;
  
  // 수익성 지표
  netIncome: number;
  profitMargin: number;
  
  // 현금 관련
  cashBalance: number;
  cashFlowChange: number;
  
  // 승인 관련
  approvalStatus: ApprovalStatus;
  approvedAt?: string;
  approvedBy?: string;
  memo?: string;
  
  // 메타데이터
  createdAt: string;
  updatedAt: string;
}

// 일별 거래 내역
export interface DailyTransaction {
  id: string;
  date: string;
  description: string;
  amount: number;
  type: 'income' | 'expense';
  category: TransactionCategory | '매출';
  detail: string;
  customer?: string;
  vendor?: string;
}

// 카테고리별 지출 요약
export interface ExpenseByCategory {
  category: TransactionCategory;
  amount: number;
  percentage: number;
}

// KPI 데이터
export interface KPIData {
  totalRevenue: number;
  totalExpense: number;
  currentCashBalance: number;
  profitMargin: number;
  
  // 전월 대비 변화율
  revenueChange: number;
  expenseChange: number;
  cashBalanceChange: number;
  profitMarginChange: number;
}

// 차트 데이터
export interface ChartDataPoint {
  period: string;
  revenue: number;
  expense: number;
  netIncome: number;
}

// 현금 흐름 데이터
export interface CashFlowData {
  year: number;
  openingBalance: number;
  
  // 매출 관련
  salesCollection: number;
  otherInflows: number;
  
  // 고정비
  fixedCosts: {
    financialProducts: number;
    salary: number;
    insurance: number;
    tax: number;
    fees: number;
    interestExpense: number;
    rent: number;
    utilities: number;
    communication: number;
  };
  
  // 유동비
  variableCosts: {
    assets: number;
    advertising: number;
    education: number;
    additionalSalary: number;
    donations: number;
    others: number;
    shortTermLoans: number;
    supplies: number;
    deposits: number;
    guarantees: number;
    officeSupplies: number;
    consumables: number;
    testFees: number;
    travel: number;
    savings: number;
    payables: number;
    transportation: number;
    associationFees: number;
  };
  
  totalFixedCosts: number;
  totalVariableCosts: number;
  totalExpenses: number;
  
  monthlyBalance: number;
  closingBalance: number;
} 