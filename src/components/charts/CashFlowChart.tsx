'use client';

import React from 'react';
import {
  LineChart,
  Line,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  ResponsiveContainer,
  Legend
} from 'recharts';
import { getCashFlowData } from '@/data/sampleData';

const formatCurrency = (value: number) => {
  return new Intl.NumberFormat('ko-KR', {
    style: 'currency',
    currency: 'KRW',
    notation: 'compact',
    maximumFractionDigits: 0
  }).format(value);
};

const CustomTooltip = ({ active, payload, label }: any) => {
  if (active && payload && payload.length) {
    const data = payload[0]?.payload;
    return (
      <div style={{
        backgroundColor: 'white',
        padding: '12px',
        border: '1px solid #ccc',
        borderRadius: '4px',
        boxShadow: '0 2px 8px rgba(0,0,0,0.1)'
      }}>
        <p style={{ margin: '4px 0', fontWeight: 'bold' }}>{`${label}`}</p>
        <p style={{ margin: '4px 0', color: '#2196f3' }}>
          💰 현금 유입: {formatCurrency(data?.inflow || 0)}
        </p>
        <p style={{ margin: '4px 0', color: '#ff9800' }}>
          💸 현금 유출: {formatCurrency(data?.outflow || 0)}
        </p>
        <p style={{ 
          margin: '4px 0', 
          color: data?.netFlow >= 0 ? '#4caf50' : '#f44336',
          fontWeight: 'bold'
        }}>
          📊 순현금흐름: {formatCurrency(data?.netFlow || 0)}
        </p>
      </div>
    );
  }
  return null;
};

export default function CashFlowChart() {
  const cashFlowData = getCashFlowData();
  
  return (
    <ResponsiveContainer width="100%" height={300}>
      <LineChart
        data={cashFlowData}
        margin={{
          top: 20,
          right: 30,
          left: 20,
          bottom: 5,
        }}
      >
        <CartesianGrid strokeDasharray="3 3" />
        <XAxis dataKey="month" />
        <YAxis tickFormatter={formatCurrency} />
        <Tooltip content={<CustomTooltip />} />
        <Legend />
        
        {/* 현금 유입 라인 */}
        <Line
          type="monotone"
          dataKey="inflow"
          stroke="#2196f3"
          strokeWidth={3}
          name="💰 현금 유입"
          dot={{ fill: '#2196f3', strokeWidth: 2, r: 6 }}
          activeDot={{ r: 8, stroke: '#2196f3', strokeWidth: 2 }}
        />
        
        {/* 현금 유출 라인 */}
        <Line
          type="monotone"
          dataKey="outflow"
          stroke="#ff9800"
          strokeWidth={3}
          name="💸 현금 유출"
          dot={{ fill: '#ff9800', strokeWidth: 2, r: 6 }}
          activeDot={{ r: 8, stroke: '#ff9800', strokeWidth: 2 }}
        />
        
        {/* 순현금흐름 라인 */}
        <Line
          type="monotone"
          dataKey="netFlow"
          stroke="#4caf50"
          strokeWidth={4}
          name="📊 순현금흐름"
          dot={{ fill: '#4caf50', strokeWidth: 2, r: 7 }}
          activeDot={{ r: 9, stroke: '#4caf50', strokeWidth: 2 }}
          strokeDasharray="5 5"
        />
      </LineChart>
    </ResponsiveContainer>
  );
} 