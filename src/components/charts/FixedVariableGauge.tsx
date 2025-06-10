'use client';

import React from 'react';
import { PieChart, Pie, Cell, ResponsiveContainer, Legend } from 'recharts';
import { Box, Typography, Grid, Paper } from '@mui/material';
import { getFixedVariableData } from '@/data/sampleData';

const RADIAN = Math.PI / 180;

const CustomizedLabel = ({ cx, cy, midAngle, innerRadius, outerRadius, percent }: any) => {
  const radius = innerRadius + (outerRadius - innerRadius) * 0.5;
  const x = cx + radius * Math.cos(-midAngle * RADIAN);
  const y = cy + radius * Math.sin(-midAngle * RADIAN);

  return (
    <text 
      x={x} 
      y={y} 
      fill="white" 
      textAnchor={x > cx ? 'start' : 'end'} 
      dominantBaseline="central"
      fontSize="14"
      fontWeight="bold"
    >
      {`${(percent * 100).toFixed(0)}%`}
    </text>
  );
};

export default function FixedVariableGauge() {
  const fixedVariableData = getFixedVariableData();
  const fixedRatio = fixedVariableData[0].value;
  const variableRatio = fixedVariableData[1].value;
  
  // 재무 건전성 평가
  const getHealthStatus = (fixedRatio: number) => {
    if (fixedRatio <= 30) return { status: '우수', color: '#4caf50' };
    if (fixedRatio <= 40) return { status: '양호', color: '#ff9800' };
    return { status: '주의', color: '#f44336' };
  };

  const healthStatus = getHealthStatus(fixedRatio);

  return (
    <Box>
      <ResponsiveContainer width="100%" height={280}>
        <PieChart>
          <Pie
            data={fixedVariableData}
            cx="50%"
            cy="50%"
            labelLine={false}
            label={CustomizedLabel}
            outerRadius={100}
            innerRadius={40}
            fill="#8884d8"
            dataKey="value"
          >
            {fixedVariableData.map((entry: any, index: number) => (
              <Cell key={`cell-${index}`} fill={entry.color} />
            ))}
          </Pie>
          <Legend 
            verticalAlign="bottom" 
            height={36}
            formatter={(value, entry) => (
              <span style={{ color: entry.color, fontWeight: 'bold' }}>
                {value}
              </span>
            )}
          />
        </PieChart>
      </ResponsiveContainer>
      
      {/* 건전성 지표 */}
      <Box mt={2}>
        <Grid container spacing={2}>
          <Grid item xs={6}>
            <Paper elevation={1} sx={{ p: 2, textAlign: 'center' }}>
              <Typography variant="body2" color="textSecondary">
                재무 건전성
              </Typography>
              <Typography 
                variant="h6" 
                fontWeight="bold"
                sx={{ color: healthStatus.color }}
              >
                {healthStatus.status}
              </Typography>
            </Paper>
          </Grid>
          <Grid item xs={6}>
            <Paper elevation={1} sx={{ p: 2, textAlign: 'center' }}>
              <Typography variant="body2" color="textSecondary">
                고정비 비율
              </Typography>
              <Typography 
                variant="h6" 
                fontWeight="bold"
                sx={{ color: healthStatus.color }}
              >
                {fixedRatio}%
              </Typography>
            </Paper>
          </Grid>
        </Grid>
        
        <Box mt={2} p={2} bgcolor="grey.50" borderRadius={1}>
          <Typography variant="body2" color="textSecondary">
            💡 <strong>분석:</strong> 고정비 비율이 {fixedRatio}%로 
            {fixedRatio <= 30 ? ' 안정적인 수준입니다.' : 
             fixedRatio <= 40 ? ' 관리 가능한 수준입니다.' : 
             ' 개선이 필요합니다.'}
          </Typography>
        </Box>
      </Box>
    </Box>
  );
} 