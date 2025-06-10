import { NextResponse } from 'next/server';
import { getBulkFinanceData, getAllBulkDataSummary } from '@/lib/dataStore';

export async function GET() {
  const bulkData = await getBulkFinanceData();
  const summary = await getAllBulkDataSummary();
  
  return NextResponse.json({
    success: true,
    timestamp: new Date().toISOString(),
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
} 