import { NextRequest, NextResponse } from 'next/server';
import { monthlyReports, kpiData, getChartData, getExpenseByCategory } from '@/data/sampleData';

// GET /api/dashboard - 대시보드 전체 데이터 조회
export async function GET() {
  try {
    // 전체 KPI 계산
    const totalRevenue = monthlyReports.reduce((sum, report) => sum + report.totalRevenue, 0);
    const totalExpense = monthlyReports.reduce((sum, report) => sum + report.totalExpense, 0);
    const totalNetIncome = monthlyReports.reduce((sum, report) => sum + report.netIncome, 0);
    const avgProfitMargin = monthlyReports.reduce((sum, report) => sum + report.profitMargin, 0) / monthlyReports.length;
    const currentCashBalance = monthlyReports[monthlyReports.length - 1]?.cashBalance || 0;

    // 차트 데이터
    const chartData = getChartData();
    
    // 최근 3개월 카테고리별 지출 합계
    const recentExpenses = monthlyReports.slice(-3).reduce((acc, report) => {
      acc.rent += report.rentExpense;
      acc.labor += report.laborExpense;
      acc.material += report.materialExpense;
      acc.operating += report.operatingExpense;
      acc.other += report.otherExpense;
      return acc;
    }, { rent: 0, labor: 0, material: 0, operating: 0, other: 0 });

    const expenseData = [
      { category: '임대료', amount: recentExpenses.rent },
      { category: '인건비', amount: recentExpenses.labor },
      { category: '재료비', amount: recentExpenses.material },
      { category: '운영비', amount: recentExpenses.operating },
      { category: '기타', amount: recentExpenses.other }
    ];

    // 주간 매출 시뮬레이션 (3월 기준)
    const weeklySalesData = [
      { week: '1주차', revenue: 70000000, target: 65000000 },
      { week: '2주차', revenue: 75000000, target: 70000000 },
      { week: '3주차', revenue: 68000000, target: 70000000 },
      { week: '4주차', revenue: 67000000, target: 65000000 }
    ];

    return NextResponse.json({
      success: true,
      data: {
        kpi: {
          totalRevenue,
          totalExpense,
          totalNetIncome,
          avgProfitMargin,
          currentCashBalance
        },
        chartData,
        expenseData,
        weeklySalesData,
        monthlyReports: monthlyReports.map(report => ({
          ...report,
          // 민감한 정보는 제외하고 반환
          memo: report.approvalStatus === 'approved' ? report.memo : undefined
        }))
      }
    });

  } catch (error) {
    return NextResponse.json(
      { success: false, error: '대시보드 데이터를 불러오는데 실패했습니다.' },
      { status: 500 }
    );
  }
} 