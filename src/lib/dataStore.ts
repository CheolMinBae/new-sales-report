import { S3Client, GetObjectCommand, PutObjectCommand, ListObjectsV2Command } from '@aws-sdk/client-s3';
import { getSignedUrl } from '@aws-sdk/s3-request-presigner';

// 공유 데이터 저장소
interface YearlyFinanceData {
  year: number;
  monthlyData: {
    [month: string]: {
      salesRevenue: number;
      otherIncome: number;
      rentExpense: number;
      laborExpense: number;
      materialExpense: number;
      operatingExpense: number;
      otherExpense: number;
      cashBalance: number;
    };
  };
}

interface BulkDataRequest {
  yearlyData: YearlyFinanceData[];
  submittedBy: string;
  submittedAt: string;
  sheetName: string;
}

// S3 설정
const S3_BUCKET_NAME = process.env.S3_BUCKET_NAME || 'sales-report-data';
const S3_REGION = process.env.S3_REGION || 'ap-northeast-2';
const S3_ACCESS_KEY_ID = process.env.S3_ACCESS_KEY_ID;
const S3_SECRET_ACCESS_KEY = process.env.S3_SECRET_ACCESS_KEY;
const S3_DATA_FILE_KEY = 'bulk-finance-data.json';

// S3 클라이언트 생성
const s3Client = new S3Client({
  region: S3_REGION,
  credentials: {
    accessKeyId: S3_ACCESS_KEY_ID || '',
    secretAccessKey: S3_SECRET_ACCESS_KEY || '',
  },
});

// S3에서 데이터 읽기
async function readDataFromS3(): Promise<BulkDataRequest[]> {
  try {
    const command = new GetObjectCommand({
      Bucket: S3_BUCKET_NAME,
      Key: S3_DATA_FILE_KEY,
    });

    const response = await s3Client.send(command);
    const data = await response.Body?.transformToString();
    
    if (data) {
      console.log('[DataStore] S3에서 데이터 읽기 성공');
      return JSON.parse(data);
    } else {
      console.log('[DataStore] S3에 데이터가 없음, 빈 배열 반환');
      return [];
    }
  } catch (error: any) {
    if (error.name === 'NoSuchKey') {
      console.log('[DataStore] S3에 파일이 없음, 빈 배열 반환');
      return [];
    }
    console.error('[DataStore] S3에서 데이터 읽기 실패:', error);
    throw error;
  }
}

// S3에 데이터 쓰기
async function writeDataToS3(data: BulkDataRequest[]): Promise<void> {
  try {
    const command = new PutObjectCommand({
      Bucket: S3_BUCKET_NAME,
      Key: S3_DATA_FILE_KEY,
      Body: JSON.stringify(data, null, 2),
      ContentType: 'application/json',
    });

    await s3Client.send(command);
    console.log('[DataStore] S3에 데이터 저장 완료');
  } catch (error) {
    console.error('[DataStore] S3에 데이터 저장 실패:', error);
    throw error;
  }
}

// S3 버킷 존재 확인 및 생성 (필요시)
async function ensureS3BucketExists(): Promise<void> {
  try {
    const command = new ListObjectsV2Command({
      Bucket: S3_BUCKET_NAME,
      MaxKeys: 1,
    });
    await s3Client.send(command);
    console.log(`[DataStore] S3 버킷 ${S3_BUCKET_NAME} 접근 가능`);
  } catch (error: any) {
    if (error.name === 'NoSuchBucket') {
      console.error(`[DataStore] S3 버킷 ${S3_BUCKET_NAME}가 존재하지 않습니다.`);
      console.error('AWS 콘솔에서 버킷을 생성하거나 환경 변수를 확인하세요.');
    }
    throw error;
  }
}

// 데이터 저장 함수
export async function addBulkFinanceData(data: BulkDataRequest): Promise<void> {
  try {
    // S3 버킷 확인
    await ensureS3BucketExists();
    
    // 기존 데이터 읽기
    const existingData = await readDataFromS3();
    
    // 기존 데이터에서 같은 제출자의 데이터 제거 (업데이트 개념)
    const filteredData = existingData.filter(
      existing => existing.submittedBy !== data.submittedBy
    );
    
    // 새 데이터 추가
    filteredData.push(data);
    
    // S3에 저장
    await writeDataToS3(filteredData);
    
    console.log(`[DataStore] S3 데이터 저장 완료. 총 ${filteredData.length}개 항목`);
    console.log(`[DataStore] 저장된 데이터:`, JSON.stringify(data, null, 2));
  } catch (error) {
    console.error('[DataStore] S3 데이터 저장 오류:', error);
    throw error;
  }
}

// 데이터 조회 함수
export async function getBulkFinanceData(): Promise<BulkDataRequest[]> {
  try {
    const data = await readDataFromS3();
    console.log(`[DataStore] S3 데이터 조회. 총 ${data.length}개 항목`);
    return data;
  } catch (error) {
    console.error('[DataStore] S3 데이터 조회 오류:', error);
    return [];
  }
}

// 모든 데이터 조회 (디버깅용)
export async function getAllBulkDataSummary(): Promise<any[]> {
  try {
    const data = await readDataFromS3();
    return data.map(record => ({
      submittedBy: record.submittedBy,
      submittedAt: record.submittedAt,
      sheetName: record.sheetName,
      totalYears: record.yearlyData.length,
      years: record.yearlyData.map(data => data.year),
      totalMonths: record.yearlyData.reduce((sum, year) => 
        sum + Object.keys(year.monthlyData).length, 0
      )
    }));
  } catch (error) {
    console.error('[DataStore] S3 요약 데이터 조회 오류:', error);
    return [];
  }
}

// S3 설정 정보 조회 (디버깅용)
export async function getS3ConfigInfo(): Promise<any> {
  return {
    bucketName: S3_BUCKET_NAME,
    region: S3_REGION,
    fileKey: S3_DATA_FILE_KEY,
    hasCredentials: !!(S3_ACCESS_KEY_ID && S3_SECRET_ACCESS_KEY),
    accessKeyId: S3_ACCESS_KEY_ID ? `${S3_ACCESS_KEY_ID.substring(0, 8)}...` : 'Not set',
  };
}

export type { YearlyFinanceData, BulkDataRequest }; 