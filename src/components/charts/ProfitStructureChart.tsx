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
import { monthlyData } from '@/data/sampleData';

export default function ProfitStructureChart() {
  const theme = useTheme();
  
  // 데이터를 수익 구조 분석에 맞게 변환
  const data = monthlyData.map(item => ({
    period: `${item.year}-${item.month < 10 ? '0' + item.month : item.month}`,
    총매출: item.totalRevenue,
    임대료: -item.rentExpense,
    인건비: -item.laborExpense,
    재료비: -item.materialExpense,
    운영비: -item.operatingExpense,
    기타비용: -item.otherExpense,
    순이익: item.netIncome
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
          data={data}
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
            dataKey="period" 
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
            dataKey="총매출" 
            stackId="profit"
            fill={theme.palette.primary.main}
            name="총매출"
          />
          
          {/* 비용들 */}
          <Bar 
            dataKey="임대료" 
            stackId="cost"
            fill="#FF6B6B"
            name="임대료"
          />
          <Bar 
            dataKey="인건비" 
            stackId="cost"
            fill="#4ECDC4"
            name="인건비"
          />
          <Bar 
            dataKey="재료비" 
            stackId="cost"
            fill="#45B7D1"
            name="재료비"
          />
          <Bar 
            dataKey="운영비" 
            stackId="cost"
            fill="#FFA07A"
            name="운영비"
          />
          <Bar 
            dataKey="기타비용" 
            stackId="cost"
            fill="#DDA0DD"
            name="기타비용"
          />
          
          {/* 순이익 */}
          <Bar 
            dataKey="순이익" 
            stackId="result"
            fill={theme.palette.success.main}
            name="순이익"
          />
        </BarChart>
      </ResponsiveContainer>
    </Box>
  );
} 