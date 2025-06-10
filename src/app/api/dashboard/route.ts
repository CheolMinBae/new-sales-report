import { NextRequest, NextResponse } from 'next/server';
import { getBulkFinanceData, type YearlyFinanceData, type BulkDataRequest } from '@/lib/dataStore';

// GET /api/dashboard - 대시보드 전체 데이터 조회
export async function GET(request: NextRequest) {
  try {
    const { searchParams } = new URL(request.url);
    const yearParam = searchParams.get('year');
    const currentYear = yearParam ? parseInt(yearParam) : new Date().getFullYear();
    
    // 실제 전송된 bulk data 가져오기
    const bulkDataList = await getBulkFinanceData();
    
    // 가장 최근에 전송된 데이터 사용
    const latestBulkData = bulkDataList.length > 0 ? bulkDataList[bulkDataList.length - 1] : null;
    
    if (!latestBulkData) {
      return NextResponse.json({
        success: true,
        data: {
          year: currentYear,
          kpi: {
            totalRevenue: 0,
            totalExpense: 0,
            totalNetIncome: 0,
            avgProfitMargin: 0,
            currentCashBalance: 0
          },
          chartData: {
            monthlyStats: [],
            revenueData: [],
            profitStructure: []
          },
          expenseData: [],
          weeklySalesData: [],
          monthlyReports: []
        }
      });
    }
    
    // 요청된 년도의 데이터 찾기
    const yearData = latestBulkData.yearlyData.find(data => data.year === currentYear);
    
    if (!yearData) {
      return NextResponse.json({
        success: true,
        data: {
          year: currentYear,
          kpi: {
            totalRevenue: 0,
            totalExpense: 0,
            totalNetIncome: 0,
            avgProfitMargin: 0,
            currentCashBalance: 0
          },
          chartData: {
            monthlyStats: [],
            revenueData: [],
            profitStructure: []
          },
          expenseData: [],
          weeklySalesData: [],
          monthlyReports: []
        }
      });
    }
    
    // 월별 데이터를 배열로 변환
    const monthlyReports = Object.entries(yearData.monthlyData).map(([monthName, data]) => {
      const totalRevenue = data.salesRevenue + data.otherIncome;
      const totalExpense = data.rentExpense + data.laborExpense + data.materialExpense + data.operatingExpense + data.otherExpense;
      const netIncome = totalRevenue - totalExpense;
      const profitMargin = totalRevenue > 0 ? (netIncome / totalRevenue) * 100 : 0;
      
      return {
        month: monthName,
        year: currentYear,
        salesRevenue: data.salesRevenue,
        otherIncome: data.otherIncome,
        totalRevenue,
        rentExpense: data.rentExpense,
        laborExpense: data.laborExpense,
        materialExpense: data.materialExpense,
        operatingExpense: data.operatingExpense,
        otherExpense: data.otherExpense,
        totalExpense,
        netIncome,
        profitMargin,
        cashBalance: data.cashBalance
      };
    });
    
    // 년도별 KPI 계산
    const totalRevenue = monthlyReports.reduce((sum, report) => sum + report.totalRevenue, 0);
    const totalExpense = monthlyReports.reduce((sum, report) => sum + report.totalExpense, 0);
    const totalNetIncome = monthlyReports.reduce((sum, report) => sum + report.netIncome, 0);
    const avgProfitMargin = monthlyReports.length > 0 ? monthlyReports.reduce((sum, report) => sum + report.profitMargin, 0) / monthlyReports.length : 0;
    const currentCashBalance = monthlyReports.length > 0 ? monthlyReports[monthlyReports.length - 1]?.cashBalance || 0 : 0;

    // 년도별 차트 데이터 생성
    const chartData = {
      monthlyStats: monthlyReports.map(report => ({
        month: report.month,
        revenue: report.totalRevenue,
        expense: report.totalExpense,
        netIncome: report.netIncome
      })),
      revenueData: monthlyReports.map(report => ({
        month: report.month,
        revenue: report.salesRevenue,
        expense: report.totalExpense,
        netIncome: report.netIncome
      })),
      profitStructure: monthlyReports.map(report => ({
        month: report.month,
        profit: Math.max(0, report.netIncome),
        loss: Math.min(0, report.netIncome)
      }))
    };
    
    // 해당 년도 카테고리별 지출 합계
    const yearlyExpenses = monthlyReports.reduce((acc, report) => {
      acc.rent += report.rentExpense;
      acc.labor += report.laborExpense;
      acc.material += report.materialExpense;
      acc.operating += report.operatingExpense;
      acc.other += report.otherExpense;
      return acc;
    }, { rent: 0, labor: 0, material: 0, operating: 0, other: 0 });

    const expenseData = [
      { category: '임대료', amount: yearlyExpenses.rent },
      { category: '인건비', amount: yearlyExpenses.labor },
      { category: '재료비', amount: yearlyExpenses.material },
      { category: '운영비', amount: yearlyExpenses.operating },
      { category: '기타', amount: yearlyExpenses.other }
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
        year: currentYear,
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
        monthlyReports: monthlyReports
      }
    });

  } catch (error) {
    return NextResponse.json(
      { success: false, error: '대시보드 데이터를 불러오는데 실패했습니다.' },
      { status: 500 }
    );
  }
} 