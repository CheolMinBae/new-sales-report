'use client';

import {
  ComposedChart,
  Line,
  Bar,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  Legend,
  ResponsiveContainer
} from 'recharts';
import { Box, Typography, useTheme } from '@mui/material';
import { getChartData } from '@/data/sampleData';

export default function RevenueChart() {
  const theme = useTheme();
  const data = getChartData();

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

  return (
    <Box sx={{ width: '100%', height: 300 }}>
      <ResponsiveContainer width="100%" height="100%">
        <ComposedChart
          data={data}
          margin={{
            top: 20,
            right: 30,
            left: 20,
            bottom: 5,
          }}
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
          <Legend 
            wrapperStyle={{ paddingTop: '10px' }}
          />
          <Bar 
            dataKey="revenue" 
            fill={theme.palette.primary.main}
            name="매출"
            radius={[4, 4, 0, 0]}
          />
          <Bar 
            dataKey="expense" 
            fill={theme.palette.secondary.main}
            name="매입"
            radius={[4, 4, 0, 0]}
          />
          <Line 
            type="monotone" 
            dataKey="netIncome" 
            stroke={theme.palette.success.main}
            strokeWidth={3}
            name="순이익"
            dot={{ fill: theme.palette.success.main, r: 4 }}
          />
        </ComposedChart>
      </ResponsiveContainer>
    </Box>
  );
} 