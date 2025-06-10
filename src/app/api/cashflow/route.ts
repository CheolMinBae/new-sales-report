import { NextRequest, NextResponse } from 'next/server';
import { getCashFlowData } from '@/data/sampleData';

// GET /api/cashflow - 현금흐름 데이터 조회
export async function GET() {
  try {
    const cashFlowData = getCashFlowData();
    
    return NextResponse.json({
      success: true,
      data: cashFlowData,
      message: '현금흐름 데이터 조회 성공'
    });
  } catch (error) {
    console.error('현금흐름 데이터 조회 오류:', error);
    return NextResponse.json(
      { 
        success: false, 
        message: '현금흐름 데이터 조회 실패',
        error: error instanceof Error ? error.message : '알 수 없는 오류'
      },
      { status: 500 }
    );
  }
}

// POST /api/cashflow - 현금흐름 데이터 저장
export async function POST(request: NextRequest) {
  try {
    const body = await request.json();
    const { cashInflow, cashOutflow, netCashFlow, month, year } = body;

    // 데이터 유효성 검사
    if (!cashInflow || !cashOutflow || !month || !year) {
      return NextResponse.json(
        { 
          success: false, 
          message: '필수 데이터가 누락되었습니다. (cashInflow, cashOutflow, month, year 필요)' 
        },
        { status: 400 }
      );
    }

    // 여기서는 메모리에 저장하는 시뮬레이션
    // 실제 환경에서는 데이터베이스에 저장
    console.log('현금흐름 데이터 저장:', {
      year,
      month,
      cashInflow,
      cashOutflow,
      netCashFlow,
      savedAt: new Date().toISOString()
    });

    return NextResponse.json({
      success: true,
      data: {
        year,
        month,
        cashInflow,
        cashOutflow,
        netCashFlow,
        savedAt: new Date().toISOString()
      },
      message: '현금흐름 데이터 저장 성공'
    });

  } catch (error) {
    console.error('현금흐름 데이터 저장 오류:', error);
    return NextResponse.json(
      { 
        success: false, 
        message: '현금흐름 데이터 저장 실패',
        error: error instanceof Error ? error.message : '알 수 없는 오류'
      },
      { status: 500 }
    );
  }
} 