'use client';

import { useState } from 'react';
import { Container, Typography, Button, Card, CardContent, Box, Alert } from '@mui/material';

export default function TestBulkPage() {
  const [result, setResult] = useState<string>('');
  const [loading, setLoading] = useState(false);

  const testData = {
    yearlyData: [
      {
        year: 2024,
        monthlyData: {
          "1월": {
            salesRevenue: 50000000,
            otherIncome: 5000000,
            rentExpense: 10000000,
            laborExpense: 15000000,
            materialExpense: 8000000,
            operatingExpense: 12000000,
            otherExpense: 3000000,
            cashBalance: 20000000
          },
          "2월": {
            salesRevenue: 55000000,
            otherIncome: 4000000,
            rentExpense: 10000000,
            laborExpense: 16000000,
            materialExpense: 9000000,
            operatingExpense: 11000000,
            otherExpense: 2000000,
            cashBalance: 25000000
          },
          "3월": {
            salesRevenue: 60000000,
            otherIncome: 6000000,
            rentExpense: 10000000,
            laborExpense: 17000000,
            materialExpense: 10000000,
            operatingExpense: 13000000,
            otherExpense: 4000000,
            cashBalance: 30000000
          }
        }
      },
      {
        year: 2025,
        monthlyData: {
          "1월": {
            salesRevenue: 65000000,
            otherIncome: 7000000,
            rentExpense: 11000000,
            laborExpense: 18000000,
            materialExpense: 11000000,
            operatingExpense: 14000000,
            otherExpense: 5000000,
            cashBalance: 35000000
          }
        }
      }
    ],
    submittedBy: "VBA_테스트사용자",
    sheetName: "20~25년_정리표"
  };

  const sendBulkData = async () => {
    setLoading(true);
    try {
      const response = await fetch('/api/bulk-data/submit', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify(testData),
      });

      const data = await response.json();
      setResult(JSON.stringify(data, null, 2));
    } catch (error) {
      setResult(`오류: ${error}`);
    } finally {
      setLoading(false);
    }
  };

  const testDashboard2024 = async () => {
    setLoading(true);
    try {
      const response = await fetch('/api/dashboard?year=2024');
      const data = await response.json();
      setResult(JSON.stringify(data, null, 2));
    } catch (error) {
      setResult(`오류: ${error}`);
    } finally {
      setLoading(false);
    }
  };

  const testDashboard2025 = async () => {
    setLoading(true);
    try {
      const response = await fetch('/api/dashboard?year=2025');
      const data = await response.json();
      setResult(JSON.stringify(data, null, 2));
    } catch (error) {
      setResult(`오류: ${error}`);
    } finally {
      setLoading(false);
    }
  };

  const testDebugData = async () => {
    setLoading(true);
    try {
      const response = await fetch('/api/debug');
      const data = await response.json();
      setResult(JSON.stringify(data, null, 2));
    } catch (error) {
      setResult(`오류: ${error}`);
    } finally {
      setLoading(false);
    }
  };

  return (
    <Container maxWidth="lg" sx={{ py: 4 }}>
      <Typography variant="h4" gutterBottom>
        📊 VBA 전송 시뮬레이션 테스트
      </Typography>
      
      <Alert severity="info" sx={{ mb: 3 }}>
        이 페이지는 VBA에서 전체년도 데이터를 전송하는 것을 시뮬레이션하고, 
        대시보드가 전송된 실제 데이터를 사용하는지 테스트합니다.
      </Alert>

      <Box sx={{ mb: 3, display: 'flex', gap: 2, flexWrap: 'wrap' }}>
        <Button 
          variant="contained" 
          color="primary" 
          onClick={sendBulkData}
          disabled={loading}
        >
          📤 VBA 데이터 전송 시뮬레이션
        </Button>
        
        <Button 
          variant="outlined" 
          color="secondary" 
          onClick={testDashboard2024}
          disabled={loading}
        >
          📊 2024년 대시보드 테스트
        </Button>
        
        <Button 
          variant="outlined" 
          color="secondary" 
          onClick={testDashboard2025}
          disabled={loading}
        >
          📊 2025년 대시보드 테스트
        </Button>
        
        <Button 
          variant="outlined" 
          color="warning" 
          onClick={testDebugData}
          disabled={loading}
        >
          🔍 저장된 데이터 확인
        </Button>
      </Box>

      <Card>
        <CardContent>
          <Typography variant="h6" gutterBottom>
            📋 테스트 결과:
          </Typography>
          <Box 
            component="pre" 
            sx={{ 
              backgroundColor: '#f5f5f5', 
              padding: 2, 
              borderRadius: 1, 
              overflow: 'auto',
              maxHeight: '500px',
              fontSize: '0.875rem'
            }}
          >
            {loading ? '로딩 중...' : result || '버튼을 클릭하여 테스트를 시작하세요.'}
          </Box>
        </CardContent>
      </Card>

      <Card sx={{ mt: 3 }}>
        <CardContent>
          <Typography variant="h6" gutterBottom>
            📋 전송될 테스트 데이터:
          </Typography>
          <Box 
            component="pre" 
            sx={{ 
              backgroundColor: '#f0f8ff', 
              padding: 2, 
              borderRadius: 1, 
              overflow: 'auto',
              maxHeight: '400px',
              fontSize: '0.875rem'
            }}
          >
            {JSON.stringify(testData, null, 2)}
          </Box>
        </CardContent>
      </Card>
    </Container>
  );
} 