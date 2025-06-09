import { NextRequest, NextResponse } from 'next/server';

// GET /api/test - VBA 연동 테스트용 API
export async function GET(request: NextRequest) {
  try {
    const currentTime = new Date().toISOString();
    const userAgent = request.headers.get('user-agent') || 'Unknown';
    const url = new URL(request.url);
    const testParam = url.searchParams.get('message') || 'Hello from API!';

    return NextResponse.json({
      success: true,
      message: "API 연결이 성공적으로 작동하고 있습니다!",
      data: {
        timestamp: currentTime,
        userAgent: userAgent,
        testMessage: testParam,
        server: "Next.js API Routes",
        status: "정상 작동 중"
      },
      korean: {
        title: "VBA 연동 테스트",
        description: "이 API는 엑셀 VBA와의 연동을 테스트하기 위한 것입니다.",
        instructions: [
          "1. VBA에서 WinHTTP로 이 API 호출",
          "2. JSON 응답 확인",
          "3. success: true 값 검증",
          "4. 메인 API 연동 진행"
        ]
      }
    });

  } catch (error) {
    return NextResponse.json(
      { 
        success: false, 
        error: '테스트 API 처리 중 오류가 발생했습니다.',
        timestamp: new Date().toISOString()
      },
      { status: 500 }
    );
  }
} 