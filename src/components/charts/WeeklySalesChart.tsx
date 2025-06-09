'use client';

import {
  LineChart,
  Line,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  Legend,
  ResponsiveContainer,
  ReferenceLine
} from 'recharts';
import { Box, Typography, useTheme } from '@mui/material';

export default function WeeklySalesChart() {
  const theme = useTheme();
  
  // 주간 매출 샘플 데이터 (3월 4주간)
  const data = [
    {
      week: '3월 1주',
      daily: [18000000, 22000000, 25000000, 19000000, 24000000, 28000000, 15000000],
      average: 21571429,
      total: 151000000
    },
    {
      week: '3월 2주', 
      daily: [20000000, 24000000, 27000000, 21000000, 26000000, 30000000, 17000000],
      average: 23571429,
      total: 165000000
    },
    {
      week: '3월 3주',
      daily: [22000000, 26000000, 29000000, 23000000, 28000000, 32000000, 19000000],
      average: 25571429,
      total: 179000000
    },
    {
      week: '3월 4주',
      daily: [25000000, 29000000, 32000000, 26000000, 31000000, 35000000, 22000000],
      average: 28571429,
      total: 200000000
    }
  ];

  // 일별 데이터로 변환 (선 그래프용)
  const weeklyTrendData = data.map(item => ({
    week: item.week,
    평균매출: item.average,
    주간총매출: item.total,
    목표매출: 25000000 // 목표 평균 매출
  }));

  const formatCurrency = (value: number) => {
    return new Intl.NumberFormat('ko-KR', {
      style: 'currency',
      currency: 'KRW',
      notation: 'compact',
      maximumFractionDigits: 1
    }).format(value);
  };

  const CustomTooltip = ({ active, payload, label }: any) => {
    if (active && payload && payload.length) {
      return (
        <Box
          sx={{
            backgroundColor: 'background.paper',
            border: 1,
            borderColor: 'divider',
            borderRadius: 1,
            p: 2,
            boxShadow: 1
          }}
        >
          <Typography variant="body2" fontWeight="bold" mb={1}>
            {label}
          </Typography>
          {payload.map((entry: any, index: number) => (
            <Typography 
              key={index}
              variant="body2" 
              sx={{ color: entry.color }}
            >
              {entry.name}: {formatCurrency(entry.value)}
            </Typography>
          ))}
        </Box>
      );
    }
    return null;
  };

  const CustomDot = ({ cx, cy, payload }: any) => {
    const isAboveTarget = payload.평균매출 > payload.목표매출;
    return (
      <circle
        cx={cx}
        cy={cy}
        r={4}
        fill={isAboveTarget ? theme.palette.success.main : theme.palette.warning.main}
        stroke={isAboveTarget ? theme.palette.success.main : theme.palette.warning.main}
        strokeWidth={2}
      />
    );
  };

  return (
    <Box sx={{ width: '100%', height: 300 }}>
      <ResponsiveContainer width="100%" height="100%">
        <LineChart
          data={weeklyTrendData}
          margin={{
            top: 20,
            right: 30,
            left: 20,
            bottom: 5,
          }}
        >
          <CartesianGrid strokeDasharray="3 3" stroke={theme.palette.divider} />
          <XAxis 
            dataKey="week" 
            tick={{ fontSize: 12 }}
            stroke={theme.palette.text.secondary}
          />
          <YAxis 
            tick={{ fontSize: 12 }}
            stroke={theme.palette.text.secondary}
            tickFormatter={formatCurrency}
          />
          <Tooltip content={<CustomTooltip />} />
          <Legend />
          
          {/* 목표 매출 기준선 */}
          <ReferenceLine 
            y={25000000} 
            stroke={theme.palette.warning.main}
            strokeDasharray="8 8"
            label={{ value: "목표", position: "insideTopRight" }}
          />
          
          {/* 평균 매출 라인 */}
          <Line
            type="monotone"
            dataKey="평균매출"
            stroke={theme.palette.primary.main}
            strokeWidth={3}
            name="일평균 매출"
            dot={<CustomDot />}
            activeDot={{ r: 6, fill: theme.palette.primary.main }}
          />
          
          {/* 주간 총매출 라인 */}
          <Line
            type="monotone"
            dataKey="주간총매출"
            stroke={theme.palette.secondary.main}
            strokeWidth={2}
            name="주간 총매출"
            strokeDasharray="5 5"
            dot={{ fill: theme.palette.secondary.main, r: 3 }}
          />
        </LineChart>
      </ResponsiveContainer>
    </Box>
  );
} 