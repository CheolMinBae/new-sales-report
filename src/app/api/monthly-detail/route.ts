import { NextRequest, NextResponse } from 'next/server';
import { getMonthlyDetailTableData } from '@/data/sampleData';

// GET /api/monthly-detail - 월별 상세 테이블 데이터 조회
export async function GET() {
  try {
    const monthlyDetailData = getMonthlyDetailTableData();
    
    // 요약 정보 계산
    const summary = monthlyDetailData.reduce(
      (acc, curr) => ({
        totalRevenue: acc.totalRevenue + curr.revenue,
        totalExpense: acc.totalExpense + curr.expense,
        totalNetIncome: acc.totalNetIncome + curr.netIncome,
        averageProfitMargin: acc.averageProfitMargin // 나중에 계산
      }),
      { totalRevenue: 0, totalExpense: 0, totalNetIncome: 0, averageProfitMargin: 0 }
    );

    // 평균 순이익률 계산
    summary.averageProfitMargin = summary.totalRevenue > 0 
      ? (summary.totalNetIncome / summary.totalRevenue) * 100
      : 0;

    return NextResponse.json({
      success: true,
      data: {
        monthlyData: monthlyDetailData,
        summary: {
          totalRevenue: summary.totalRevenue,
          totalExpense: summary.totalExpense,
          totalNetIncome: summary.totalNetIncome,
          averageProfitMargin: Math.round(summary.averageProfitMargin * 10) / 10,
          dataCount: monthlyDetailData.length
        }
      },
      message: '월별 상세 데이터 조회 성공'
    });
  } catch (error) {
    console.error('월별 상세 데이터 조회 오류:', error);
    return NextResponse.json(
      { 
        success: false, 
        message: '월별 상세 데이터 조회 실패',
        error: error instanceof Error ? error.message : '알 수 없는 오류'
      },
      { status: 500 }
    );
  }
}

// POST /api/monthly-detail - 월별 상세 데이터 업데이트
export async function POST(request: NextRequest) {
  try {
    const body = await request.json();
    const { month, year, revenue, expense, netIncome } = body;

    // 데이터 유효성 검사
    if (!month || !year || revenue === undefined || expense === undefined) {
      return NextResponse.json(
        { 
          success: false, 
          message: '필수 데이터가 누락되었습니다. (month, year, revenue, expense 필요)' 
        },
        { status: 400 }
      );
    }

    // 순이익 자동 계산 (제공되지 않은 경우)
    const calculatedNetIncome = netIncome !== undefined ? netIncome : revenue - expense;

    // 여기서는 메모리에 저장하는 시뮬레이션
    // 실제 환경에서는 데이터베이스에 저장
    console.log('월별 상세 데이터 저장:', {
      year,
      month,
      revenue,
      expense,
      netIncome: calculatedNetIncome,
      status: calculatedNetIncome >= 0 ? '흑자' : '적자',
      updatedAt: new Date().toISOString()
    });

    return NextResponse.json({
      success: true,
      data: {
        year,
        month,
        revenue,
        expense,
        netIncome: calculatedNetIncome,
        status: calculatedNetIncome >= 0 ? '흑자' : '적자',
        updatedAt: new Date().toISOString()
      },
      message: '월별 상세 데이터 저장 성공'
    });

  } catch (error) {
    console.error('월별 상세 데이터 저장 오류:', error);
    return NextResponse.json(
      { 
        success: false, 
        message: '월별 상세 데이터 저장 실패',
        error: error instanceof Error ? error.message : '알 수 없는 오류'
      },
      { status: 500 }
    );
  }
} 