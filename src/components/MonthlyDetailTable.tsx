'use client';

import React from 'react';
import {
  Table,
  TableBody,
  TableCell,
  TableContainer,
  TableHead,
  TableRow,
  Paper,
  Typography,
  Chip,
  Box
} from '@mui/material';
const formatCurrency = (amount: number) => {
  return new Intl.NumberFormat('ko-KR', {
    style: 'currency',
    currency: 'KRW',
    minimumFractionDigits: 0,
    maximumFractionDigits: 0
  }).format(amount);
};

const getStatusColor = (status: string) => {
  return status === '흑자' ? 'success' : 'error';
};

const getNetIncomeColor = (amount: number) => {
  return amount >= 0 ? '#4caf50' : '#f44336';
};

interface MonthlyDetailTableProps {
  data?: any[];
}

export default function MonthlyDetailTable({ data = [] }: MonthlyDetailTableProps) {
  // 데이터를 테이블 형식으로 변환
  const monthlyDetailData = data.map((item, index) => ({
    month: `${index + 1}월`,
    revenue: item.totalRevenue || 0,
    expense: item.totalExpense || 0,
    netIncome: item.netIncome || 0,
    status: (item.netIncome || 0) >= 0 ? '흑자' : '적자',
    cumulativeRevenue: data.slice(0, index + 1).reduce((sum, d) => sum + (d.totalRevenue || 0), 0),
    cumulativeExpense: data.slice(0, index + 1).reduce((sum, d) => sum + (d.totalExpense || 0), 0),
    cumulativeNet: data.slice(0, index + 1).reduce((sum, d) => sum + (d.netIncome || 0), 0)
  }));
  
  return (
    <Box>
      <Typography variant="h6" gutterBottom fontWeight="bold">
        월별 상세 실적 및 누계
      </Typography>
      
      <TableContainer component={Paper} elevation={2}>
        <Table>
          <TableHead>
            <TableRow sx={{ backgroundColor: '#f5f5f5' }}>
              <TableCell align="center" sx={{ fontWeight: 'bold' }}>월</TableCell>
              <TableCell align="right" sx={{ fontWeight: 'bold' }}>매출</TableCell>
              <TableCell align="right" sx={{ fontWeight: 'bold' }}>매입</TableCell>
              <TableCell align="right" sx={{ fontWeight: 'bold' }}>순이익</TableCell>
              <TableCell align="center" sx={{ fontWeight: 'bold' }}>상태</TableCell>
              <TableCell align="right" sx={{ fontWeight: 'bold', backgroundColor: '#e3f2fd' }}>
                누계 매출
              </TableCell>
              <TableCell align="right" sx={{ fontWeight: 'bold', backgroundColor: '#e3f2fd' }}>
                누계 매입
              </TableCell>
              <TableCell align="right" sx={{ fontWeight: 'bold', backgroundColor: '#e3f2fd' }}>
                누계 순이익
              </TableCell>
            </TableRow>
          </TableHead>
          <TableBody>
            {monthlyDetailData.map((row: any) => (
              <TableRow key={row.month} hover>
                <TableCell align="center" sx={{ fontWeight: 'bold' }}>
                  {row.month}
                </TableCell>
                <TableCell align="right">
                  {formatCurrency(row.revenue)}
                </TableCell>
                <TableCell align="right">
                  {formatCurrency(row.expense)}
                </TableCell>
                <TableCell 
                  align="right" 
                  sx={{ 
                    color: getNetIncomeColor(row.netIncome),
                    fontWeight: 'bold'
                  }}
                >
                  {formatCurrency(row.netIncome)}
                </TableCell>
                <TableCell align="center">
                  <Chip 
                    label={row.status} 
                    size="small" 
                    color={getStatusColor(row.status)}
                  />
                </TableCell>
                <TableCell 
                  align="right" 
                  sx={{ backgroundColor: '#f8f9fa', fontWeight: 'bold' }}
                >
                  {formatCurrency(row.cumulativeRevenue)}
                </TableCell>
                <TableCell 
                  align="right" 
                  sx={{ backgroundColor: '#f8f9fa', fontWeight: 'bold' }}
                >
                  {formatCurrency(row.cumulativeExpense)}
                </TableCell>
                <TableCell 
                  align="right" 
                  sx={{ 
                    backgroundColor: '#f8f9fa',
                    color: getNetIncomeColor(row.cumulativeNet),
                    fontWeight: 'bold'
                  }}
                >
                  {formatCurrency(row.cumulativeNet)}
                </TableCell>
              </TableRow>
            ))}
          </TableBody>
        </Table>
      </TableContainer>
      
      {/* 요약 정보 */}
      {monthlyDetailData.length > 0 && (
        <Box mt={3} p={2} bgcolor="grey.50" borderRadius={1}>
          <Typography variant="body2" color="textSecondary">
            <strong>📊 연간 요약:</strong>
          </Typography>
          <Typography variant="body2" mt={1}>
            • 총 매출: {formatCurrency(monthlyDetailData[monthlyDetailData.length - 1]?.cumulativeRevenue || 0)} | 
            총 매입: {formatCurrency(monthlyDetailData[monthlyDetailData.length - 1]?.cumulativeExpense || 0)}
          </Typography>
          <Typography variant="body2">
            • 최종 순이익: <span style={{ 
              color: getNetIncomeColor(monthlyDetailData[monthlyDetailData.length - 1]?.cumulativeNet || 0), 
              fontWeight: 'bold' 
            }}>
              {formatCurrency(monthlyDetailData[monthlyDetailData.length - 1]?.cumulativeNet || 0)}
            </span>
          </Typography>
          <Typography variant="body2">
            • 평균 월 순이익률: <span style={{ fontWeight: 'bold' }}>
              {monthlyDetailData.length > 0 ? 
                ((monthlyDetailData[monthlyDetailData.length - 1]?.cumulativeNet || 0) / 
                 (monthlyDetailData[monthlyDetailData.length - 1]?.cumulativeRevenue || 1) * 100).toFixed(1) : 0}%
            </span>
          </Typography>
        </Box>
      )}
    </Box>
  );
} 