import { MonthlyFinanceData } from '@/types/finance';

const API_BASE = '/api';

// API 응답 타입 정의
interface ApiResponse<T> {
  success: boolean;
  data?: T;
  error?: string;
  message?: string;
}

// 전체 월별 레포트 조회
export const getAllReports = async (): Promise<MonthlyFinanceData[]> => {
  try {
    const response = await fetch(`${API_BASE}/reports`);
    const result: ApiResponse<MonthlyFinanceData[]> = await response.json();
    
    if (!result.success) {
      throw new Error(result.error || '데이터를 불러오는데 실패했습니다.');
    }
    
    return result.data || [];
  } catch (error) {
    console.error('getAllReports error:', error);
    throw error;
  }
};

// 특정 월 레포트 조회
export const getReport = async (month: number, year: number = 2024): Promise<MonthlyFinanceData> => {
  try {
    const response = await fetch(`${API_BASE}/reports/${month}?year=${year}`);
    const result: ApiResponse<MonthlyFinanceData> = await response.json();
    
    if (!result.success) {
      throw new Error(result.error || '데이터를 불러오는데 실패했습니다.');
    }
    
    if (!result.data) {
      throw new Error('데이터가 없습니다.');
    }
    
    return result.data;
  } catch (error) {
    console.error('getReport error:', error);
    throw error;
  }
};

// 레포트 승인/반려 업데이트
export const updateReportApproval = async (
  month: number,
  approvalStatus: 'approved' | 'rejected' | 'pending',
  memo: string = '',
  approvedBy: string = '관리자'
): Promise<MonthlyFinanceData> => {
  try {
    const response = await fetch(`${API_BASE}/reports/${month}`, {
      method: 'PUT',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        approvalStatus,
        memo,
        approvedBy
      }),
    });
    
    const result: ApiResponse<MonthlyFinanceData> = await response.json();
    
    if (!result.success) {
      throw new Error(result.error || '상태 업데이트에 실패했습니다.');
    }
    
    if (!result.data) {
      throw new Error('업데이트된 데이터가 없습니다.');
    }
    
    return result.data;
  } catch (error) {
    console.error('updateReportApproval error:', error);
    throw error;
  }
};

// 대시보드 데이터 조회
export const getDashboardData = async () => {
  try {
    const response = await fetch(`${API_BASE}/dashboard`);
    const result = await response.json();
    
    if (!result.success) {
      throw new Error(result.error || '대시보드 데이터를 불러오는데 실패했습니다.');
    }
    
    return result.data;
  } catch (error) {
    console.error('getDashboardData error:', error);
    throw error;
  }
};

// 엑셀 VBA 연동 - 승인/반려 정보 전송
export const sendExcelApproval = async (
  month: number,
  year: number,
  approvalStatus: 'approved' | 'rejected' | 'pending',
  memo: string = '',
  approvedBy: string = '',
  excelVersion: string = ''
) => {
  try {
    const response = await fetch(`${API_BASE}/excel`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        month,
        year,
        approvalStatus,
        memo,
        approvedBy,
        excelVersion
      }),
    });
    
    const result = await response.json();
    
    if (!result.success) {
      throw new Error(result.error || '엑셀 연동에 실패했습니다.');
    }
    
    return result.data;
  } catch (error) {
    console.error('sendExcelApproval error:', error);
    throw error;
  }
};

// 엑셀에서 승인 상태 조회
export const getExcelApprovalStatus = async (month: number, year: number = 2024) => {
  try {
    const response = await fetch(`${API_BASE}/excel?month=${month}&year=${year}`);
    const result = await response.json();
    
    if (!result.success) {
      throw new Error(result.error || '상태 조회에 실패했습니다.');
    }
    
    return result.data;
  } catch (error) {
    console.error('getExcelApprovalStatus error:', error);
    throw error;
  }
}; 