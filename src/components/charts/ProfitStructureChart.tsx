'use client';

import {
  BarChart,
  Bar,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  Legend,
  ResponsiveContainer
} from 'recharts';
import { Box, Typography, useTheme } from '@mui/material';

interface ProfitStructureChartProps {
  data?: any[];
}

export default function ProfitStructureChart({ data = [] }: ProfitStructureChartProps) {
  const theme = useTheme();
  
  // 데이터를 수익 구조 분석에 맞게 변환
  const chartData = data.map(item => ({
    month: item.month,
    profit: item.profit || 0,
    loss: item.loss || 0
  }));

  const formatCurrency = (value: number) => {
    return new Intl.NumberFormat('ko-KR', {
      style: 'currency',
      currency: 'KRW',
      notation: 'compact',
      maximumFractionDigits: 1
    }).format(Math.abs(value));
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
            boxShadow: 1,
            minWidth: 200
          }}
        >
          <Typography variant="body2" fontWeight="bold" mb={1}>
            {label}
          </Typography>
          {payload
            .sort((a: any, b: any) => Math.abs(b.value) - Math.abs(a.value))
            .map((entry: any, index: number) => (
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

  return (
    <Box sx={{ width: '100%', height: 350 }}>
      <ResponsiveContainer width="100%" height="100%">
        <BarChart
          data={chartData}
          margin={{
            top: 20,
            right: 30,
            left: 20,
            bottom: 5,
          }}
          stackOffset="sign"
        >
          <CartesianGrid strokeDasharray="3 3" stroke={theme.palette.divider} />
          <XAxis 
            dataKey="month" 
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
          
          {/* 수익 */}
          <Bar 
            dataKey="profit" 
            stackId="result"
            fill={theme.palette.success.main}
            name="수익"
          />
          
          {/* 손실 */}
          <Bar 
            dataKey="loss" 
            stackId="result"
            fill={theme.palette.error.main}
            name="손실"
          />
        </BarChart>
      </ResponsiveContainer>
    </Box>
  );
} 