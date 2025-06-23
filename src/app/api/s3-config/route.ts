import { NextRequest, NextResponse } from 'next/server';
import { getS3ConfigInfo } from '@/lib/dataStore';
import { S3Client, ListObjectsV2Command, PutObjectCommand, GetObjectCommand } from '@aws-sdk/client-s3';

// S3 설정 가져오기
function getS3Config() {
  return {
    bucketName: process.env.S3_BUCKET_NAME || 'sales-report-data',
    region: process.env.S3_REGION || 'ap-northeast-2',
    accessKeyId: process.env.S3_ACCESS_KEY_ID,
    secretAccessKey: process.env.S3_SECRET_ACCESS_KEY,
  };
}

// S3 클라이언트 생성
function createS3Client() {
  const config = getS3Config();
  return new S3Client({
    region: config.region,
    credentials: {
      accessKeyId: config.accessKeyId || '',
      secretAccessKey: config.secretAccessKey || '',
    },
  });
}

// GET - S3 설정 정보 조회
export async function GET() {
  try {
    const s3Config = await getS3ConfigInfo();
    
    return NextResponse.json({
      success: true,
      s3Config: s3Config,
      timestamp: new Date().toISOString(),
    });
  } catch (error) {
    console.error('S3 config GET error:', error);
    return NextResponse.json({
      success: false,
      error: 'S3 config retrieval failed',
      details: error instanceof Error ? error.message : 'Unknown error',
    }, { status: 500 });
  }
}

// POST - S3 연결 테스트
export async function POST(request: NextRequest) {
  try {
    const body = await request.json();
    const { action } = body;
    
    const config = getS3Config();
    const s3Client = createS3Client();
    
    switch (action) {
      case 'test-connection':
        // S3 버킷 접근 테스트
        const listCommand = new ListObjectsV2Command({
          Bucket: config.bucketName,
          MaxKeys: 1,
        });
        
        await s3Client.send(listCommand);
        
        return NextResponse.json({
          success: true,
          message: `S3 버킷 ${config.bucketName}에 성공적으로 연결되었습니다.`,
          bucketName: config.bucketName,
          region: config.region,
          timestamp: new Date().toISOString(),
        });
        
      case 'test-write':
        // S3 쓰기 테스트
        const testData = {
          test: true,
          timestamp: new Date().toISOString(),
          message: 'S3 쓰기 테스트',
        };
        
        const putCommand = new PutObjectCommand({
          Bucket: config.bucketName,
          Key: 'test-connection.json',
          Body: JSON.stringify(testData, null, 2),
          ContentType: 'application/json',
        });
        
        await s3Client.send(putCommand);
        
        return NextResponse.json({
          success: true,
          message: 'S3 쓰기 테스트가 성공했습니다.',
          testFile: 'test-connection.json',
          timestamp: new Date().toISOString(),
        });
        
      case 'test-read':
        // S3 읽기 테스트
        const getCommand = new GetObjectCommand({
          Bucket: config.bucketName,
          Key: 'test-connection.json',
        });
        
        const response = await s3Client.send(getCommand);
        const data = await response.Body?.transformToString();
        
        return NextResponse.json({
          success: true,
          message: 'S3 읽기 테스트가 성공했습니다.',
          testData: data ? JSON.parse(data) : null,
          timestamp: new Date().toISOString(),
        });
        
      default:
        return NextResponse.json({
          success: false,
          error: 'Invalid action',
          validActions: ['test-connection', 'test-write', 'test-read'],
        }, { status: 400 });
    }
  } catch (error: any) {
    console.error('S3 config POST error:', error);
    
    let errorMessage = 'S3 작업 실패';
    if (error.name === 'NoSuchBucket') {
      errorMessage = `S3 버킷 ${getS3Config().bucketName}가 존재하지 않습니다.`;
    } else if (error.name === 'AccessDenied') {
      errorMessage = 'S3 접근 권한이 없습니다. 액세스 키와 시크릿 키를 확인하세요.';
    } else if (error.name === 'InvalidAccessKeyId') {
      errorMessage = '유효하지 않은 액세스 키입니다.';
    } else if (error.name === 'SignatureDoesNotMatch') {
      errorMessage = '시크릿 키가 올바르지 않습니다.';
    }
    
    return NextResponse.json({
      success: false,
      error: errorMessage,
      details: error.message,
      errorCode: error.name,
      timestamp: new Date().toISOString(),
    }, { status: 500 });
  }
} 