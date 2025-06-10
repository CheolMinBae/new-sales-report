import { NextRequest, NextResponse } from 'next/server';

export async function GET(request: NextRequest) {
  try {
    const { searchParams } = new URL(request.url);
    const message = searchParams.get('message') || 'Hello from Bulk Data API!';

    return NextResponse.json({
      success: true,
      message: `✅ Bulk Data API 연결 성공! ${message}`,
      timestamp: new Date().toISOString(),
      endpoint: '/api/bulk-data/test',
      status: 'running'
    });

  } catch (error) {
    console.error('Bulk data test error:', error);
    return NextResponse.json(
      {
        success: false,
        error: 'Test endpoint error',
        details: error instanceof Error ? error.message : 'Unknown error'
      },
      { status: 500 }
    );
  }
}

export async function POST(request: NextRequest) {
  try {
    const body = await request.json();
    
    return NextResponse.json({
      success: true,
      message: '✅ Bulk Data API POST 테스트 성공!',
      receivedData: body,
      timestamp: new Date().toISOString(),
      endpoint: '/api/bulk-data/test'
    });

  } catch (error) {
    console.error('Bulk data POST test error:', error);
    return NextResponse.json(
      {
        success: false,
        error: 'POST test endpoint error',
        details: error instanceof Error ? error.message : 'Unknown error'
      },
      { status: 500 }
    );
  }
} 