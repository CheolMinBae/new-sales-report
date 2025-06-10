import { promises as fs } from 'fs';
import path from 'path';

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

// 데이터 파일 경로
const DATA_FILE_PATH = path.join(process.cwd(), 'data', 'bulk-finance-data.json');

// 데이터 디렉토리 생성
async function ensureDataDirectory(): Promise<void> {
  const dataDir = path.dirname(DATA_FILE_PATH);
  try {
    await fs.access(dataDir);
  } catch {
    await fs.mkdir(dataDir, { recursive: true });
  }
}

// 파일에서 데이터 읽기
async function readDataFromFile(): Promise<BulkDataRequest[]> {
  try {
    await ensureDataDirectory();
    const data = await fs.readFile(DATA_FILE_PATH, 'utf-8');
    return JSON.parse(data);
  } catch (error) {
    console.log('[DataStore] 파일이 없거나 읽기 실패, 빈 배열 반환');
    return [];
  }
}

// 파일에 데이터 쓰기
async function writeDataToFile(data: BulkDataRequest[]): Promise<void> {
  try {
    await ensureDataDirectory();
    await fs.writeFile(DATA_FILE_PATH, JSON.stringify(data, null, 2), 'utf-8');
    console.log('[DataStore] 파일에 데이터 저장 완료');
  } catch (error) {
    console.error('[DataStore] 파일 저장 실패:', error);
    throw error;
  }
}

// 데이터 저장 함수
export async function addBulkFinanceData(data: BulkDataRequest): Promise<void> {
  try {
    // 기존 데이터 읽기
    const existingData = await readDataFromFile();
    
    // 기존 데이터에서 같은 제출자의 데이터 제거 (업데이트 개념)
    const filteredData = existingData.filter(
      existing => existing.submittedBy !== data.submittedBy
    );
    
    // 새 데이터 추가
    filteredData.push(data);
    
    // 파일에 저장
    await writeDataToFile(filteredData);
    
    console.log(`[DataStore] 데이터 저장 완료. 총 ${filteredData.length}개 항목`);
    console.log(`[DataStore] 저장된 데이터:`, JSON.stringify(data, null, 2));
  } catch (error) {
    console.error('[DataStore] 데이터 저장 오류:', error);
    throw error;
  }
}

// 데이터 조회 함수
export async function getBulkFinanceData(): Promise<BulkDataRequest[]> {
  try {
    const data = await readDataFromFile();
    console.log(`[DataStore] 데이터 조회. 총 ${data.length}개 항목`);
    return data;
  } catch (error) {
    console.error('[DataStore] 데이터 조회 오류:', error);
    return [];
  }
}

// 모든 데이터 조회 (디버깅용)
export async function getAllBulkDataSummary(): Promise<any[]> {
  try {
    const data = await readDataFromFile();
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
    console.error('[DataStore] 요약 데이터 조회 오류:', error);
    return [];
  }
}

export type { YearlyFinanceData, BulkDataRequest }; 