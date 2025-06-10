import { NextRequest, NextResponse } from 'next/server';
import { getFixedVariableData } from '@/data/sampleData';

// GET /api/fixed-variable - 고정비/유동비 데이터 조회
export async function GET() {
  try {
    const fixedVariableData = getFixedVariableData();
    
    return NextResponse.json({
      success: true,
      data: fixedVariableData,
      message: '고정비/유동비 데이터 조회 성공'
    });
  } catch (error) {
    console.error('고정비/유동비 데이터 조회 오류:', error);
    return NextResponse.json(
      { 
        success: false, 
        message: '고정비/유동비 데이터 조회 실패',
        error: error instanceof Error ? error.message : '알 수 없는 오류'
      },
      { status: 500 }
    );
  }
}

// POST /api/fixed-variable - 고정비/유동비 데이터 저장
export async function POST(request: NextRequest) {
  try {
    const body = await request.json();
    const { fixedCost, variableCost, fixedRatio, variableRatio, month, year } = body;

    // 데이터 유효성 검사
    if (!fixedCost || !variableCost || !month || !year) {
      return NextResponse.json(
        { 
          success: false, 
          message: '필수 데이터가 누락되었습니다. (fixedCost, variableCost, month, year 필요)' 
        },
        { status: 400 }
      );
    }

    // 비율 자동 계산 (클라이언트에서 보내지 않은 경우)
    const totalCost = fixedCost + variableCost;
    const calculatedFixedRatio = fixedRatio || (fixedCost / totalCost) * 100;
    const calculatedVariableRatio = variableRatio || (variableCost / totalCost) * 100;

    // 여기서는 메모리에 저장하는 시뮬레이션
    // 실제 환경에서는 데이터베이스에 저장
    console.log('고정비/유동비 데이터 저장:', {
      year,
      month,
      fixedCost,
      variableCost,
      fixedRatio: Math.round(calculatedFixedRatio * 10) / 10,
      variableRatio: Math.round(calculatedVariableRatio * 10) / 10,
      savedAt: new Date().toISOString()
    });

    return NextResponse.json({
      success: true,
      data: {
        year,
        month,
        fixedCost,
        variableCost,
        fixedRatio: Math.round(calculatedFixedRatio * 10) / 10,
        variableRatio: Math.round(calculatedVariableRatio * 10) / 10,
        savedAt: new Date().toISOString()
      },
      message: '고정비/유동비 데이터 저장 성공'
    });

  } catch (error) {
    console.error('고정비/유동비 데이터 저장 오류:', error);
    return NextResponse.json(
      { 
        success: false, 
        message: '고정비/유동비 데이터 저장 실패',
        error: error instanceof Error ? error.message : '알 수 없는 오류'
      },
      { status: 500 }
    );
  }
} 