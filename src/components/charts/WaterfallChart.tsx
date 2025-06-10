'use client';

import React from 'react';
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, Cell } from 'recharts';
import { getWaterfallData } from '@/data/sampleData';

const formatCurrency = (value: number) => {
  return new Intl.NumberFormat('ko-KR', {
    style: 'currency',
    currency: 'KRW',
    notation: 'compact',
    maximumFractionDigits: 0
  }).format(Math.abs(value));
};

const getBarColor = (type: string) => {
  switch (type) {
    case 'positive': return '#4caf50';
    case 'negative': return '#f44336';
    case 'total': return '#2196f3';
    default: return '#757575';
  }
};

const CustomTooltip = ({ active, payload, label }: any) => {
  if (active && payload && payload.length) {
    const data = payload[0].payload;
    return (
      <div style={{
        backgroundColor: 'white',
        padding: '12px',
        border: '1px solid #ccc',
        borderRadius: '4px',
        boxShadow: '0 2px 8px rgba(0,0,0,0.1)'
      }}>
        <p style={{ margin: '4px 0', fontWeight: 'bold' }}>{label}</p>
        <p style={{ margin: '4px 0', color: getBarColor(data.type) }}>
          금액: {data.value >= 0 ? '+' : ''}{formatCurrency(data.value)}
        </p>
        <p style={{ margin: '4px 0', color: '#666' }}>
          누적: {formatCurrency(data.cumulative)}
        </p>
      </div>
    );
  }
  return null;
};

export default function WaterfallChart() {
  const waterfallData = getWaterfallData();
  
  return (
    <ResponsiveContainer width="100%" height={400}>
      <BarChart
        data={waterfallData}
        margin={{
          top: 20,
          right: 30,
          left: 20,
          bottom: 5,
        }}
      >
        <CartesianGrid strokeDasharray="3 3" />
        <XAxis 
          dataKey="name" 
          angle={-45}
          textAnchor="end"
          height={80}
          interval={0}
        />
        <YAxis tickFormatter={formatCurrency} />
        <Tooltip content={<CustomTooltip />} />
        <Bar dataKey="value" radius={[4, 4, 0, 0]}>
          {waterfallData.map((entry: any, index: number) => (
            <Cell key={`cell-${index}`} fill={getBarColor(entry.type)} />
          ))}
        </Bar>
      </BarChart>
    </ResponsiveContainer>
  );
} 