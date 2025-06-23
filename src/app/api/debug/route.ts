import { NextResponse } from 'next/server';
import { getBulkFinanceData, getAllBulkDataSummary, getS3ConfigInfo } from '@/lib/dataStore';

export async function GET() {
  try {
    const bulkData = await getBulkFinanceData();
    const summary = await getAllBulkDataSummary();
    const s3Config = await getS3ConfigInfo();
    
    return NextResponse.json({
      success: true,
      timestamp: new Date().toISOString(),
      s3Config: s3Config,
      totalRecords: bulkData.length,
      summary: summary,
      fullData: bulkData.map((record, index) => ({
        index,
        submittedBy: record.submittedBy,
        submittedAt: record.submittedAt,
        sheetName: record.sheetName,
        yearCount: record.yearlyData.length,
        years: record.yearlyData.map(y => y.year),
        firstYearSample: record.yearlyData[0] ? {
          year: record.yearlyData[0].year,
          monthsCount: Object.keys(record.yearlyData[0].monthlyData).length,
          months: Object.keys(record.yearlyData[0].monthlyData),
          firstMonthSample: Object.entries(record.yearlyData[0].monthlyData)[0]
        } : null
      }))
    });
  } catch (error) {
    console.error('Debug endpoint error:', error);
    return NextResponse.json({
      success: false,
      error: 'Debug endpoint error',
      details: error instanceof Error ? error.message : 'Unknown error',
      timestamp: new Date().toISOString()
    }, { status: 500 });
  }
} 