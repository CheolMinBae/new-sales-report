import { NextRequest, NextResponse } from 'next/server';
import { monthlyReports } from '@/data/sampleData';

// POST /api/excel - 엑셀 VBA에서 승인/반려 정보 수신
export async function POST(request: NextRequest) {
  try {
    const body = await request.json();
    const { month, year, approvalStatus, memo, approvedBy, excelVersion } = body;

    // 유효성 검사
    if (!month || !year || !approvalStatus) {
      return NextResponse.json(
        { success: false, error: '필수 파라미터가 누락되었습니다.' },
        { status: 400 }
      );
    }

    if (!['approved', 'rejected', 'pending'].includes(approvalStatus)) {
      return NextResponse.json(
        { success: false, error: '유효하지 않은 승인 상태입니다.' },
        { status: 400 }
      );
    }

    // 해당 월 데이터 찾기
    const reportIndex = monthlyReports.findIndex(
      r => r.month === parseInt(month) && r.year === parseInt(year)
    );

    if (reportIndex === -1) {
      return NextResponse.json(
        { success: false, error: '해당 월의 데이터를 찾을 수 없습니다.' },
        { status: 404 }
      );
    }

    // 데이터 업데이트
    monthlyReports[reportIndex] = {
      ...monthlyReports[reportIndex],
      approvalStatus,
      memo: memo || '',
      approvedBy: approvedBy || 'Excel VBA',
      approvedAt: approvalStatus !== 'pending' ? new Date().toISOString() : undefined,
      updatedAt: new Date().toISOString()
    };

    // 로그 기록 (실제 운영에서는 데이터베이스 로그 테이블에 저장)
    console.log(`Excel VBA 연동: ${year}년 ${month}월 레포트 ${approvalStatus} by ${approvedBy}`);

    return NextResponse.json({
      success: true,
      data: {
        month: parseInt(month),
        year: parseInt(year),
        approvalStatus,
        message: `${year}년 ${month}월 레포트가 성공적으로 ${approvalStatus === 'approved' ? '승인' : '반려'}되었습니다.`,
        timestamp: new Date().toISOString(),
        excelVersion
      }
    });

  } catch (error) {
    console.error('Excel API Error:', error);
    return NextResponse.json(
      { success: false, error: '엑셀 연동 처리 중 오류가 발생했습니다.' },
      { status: 500 }
    );
  }
}

// GET /api/excel - 엑셀에서 현재 승인 상태 조회
export async function GET(request: NextRequest) {
  try {
    const url = new URL(request.url);
    const month = url.searchParams.get('month');
    const year = url.searchParams.get('year') || '2024';

    if (!month) {
      return NextResponse.json(
        { success: false, error: 'month 파라미터가 필요합니다.' },
        { status: 400 }
      );
    }

    const report = monthlyReports.find(
      r => r.month === parseInt(month) && r.year === parseInt(year)
    );

    if (!report) {
      return NextResponse.json(
        { success: false, error: '해당 월의 데이터를 찾을 수 없습니다.' },
        { status: 404 }
      );
    }

    // 엑셀에서 필요한 정보만 반환
    return NextResponse.json({
      success: true,
      data: {
        month: report.month,
        year: report.year,
        approvalStatus: report.approvalStatus,
        approvedBy: report.approvedBy,
        approvedAt: report.approvedAt,
        memo: report.memo,
        canEdit: report.approvalStatus === 'pending'
      }
    });

  } catch (error) {
    return NextResponse.json(
      { success: false, error: '데이터 조회 중 오류가 발생했습니다.' },
      { status: 500 }
    );
  }
} 