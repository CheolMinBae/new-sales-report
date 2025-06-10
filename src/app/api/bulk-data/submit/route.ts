import { NextRequest, NextResponse } from 'next/server';
import { addBulkFinanceData, getBulkFinanceData, getAllBulkDataSummary, type BulkDataRequest } from '@/lib/dataStore';

export async function POST(request: NextRequest) {
  try {
    const body: BulkDataRequest = await request.json();
    
    // 데이터 유효성 검증
    if (!body.yearlyData || !Array.isArray(body.yearlyData)) {
      return NextResponse.json(
        { 
          success: false, 
          error: 'yearlyData is required and must be an array' 
        },
        { status: 400 }
      );
    }

    // 제출 시간 추가
    const bulkData: BulkDataRequest = {
      ...body,
      submittedAt: new Date().toISOString()
    };

    // 공유 저장소에 데이터 저장
    await addBulkFinanceData(bulkData);

    // 성공 응답
    return NextResponse.json({
      success: true,
      message: `${body.yearlyData.length}개 년도의 데이터가 성공적으로 저장되었습니다.`,
      data: {
        totalYears: body.yearlyData.length,
        years: body.yearlyData.map(data => data.year),
        submittedBy: body.submittedBy,
        submittedAt: bulkData.submittedAt,
        sheetName: body.sheetName
      }
    });

  } catch (error) {
    console.error('Bulk data submission error:', error);
    return NextResponse.json(
      {
        success: false,
        error: 'Internal server error',
        details: error instanceof Error ? error.message : 'Unknown error'
      },
      { status: 500 }
    );
  }
}

export async function GET(request: NextRequest) {
  try {
    const { searchParams } = new URL(request.url);
    const submittedBy = searchParams.get('submittedBy');

    let responseData = await getBulkFinanceData();

    // 특정 제출자의 데이터만 조회
    if (submittedBy) {
      responseData = responseData.filter(
        data => data.submittedBy === submittedBy
      );
    }

    return NextResponse.json({
      success: true,
      totalRecords: responseData.length,
      data: responseData,
      summary: await getAllBulkDataSummary() // 디버깅용 요약 정보 추가
    });

  } catch (error) {
    console.error('Bulk data retrieval error:', error);
    return NextResponse.json(
      {
        success: false,
        error: 'Internal server error'
      },
      { status: 500 }
    );
  }
} 