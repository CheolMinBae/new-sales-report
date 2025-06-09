'use client';

import { useState } from 'react';
import {
  Container,
  Typography,
  Button,
  Box,
  Paper,
  TextField,
  Select,
  MenuItem,
  FormControl,
  InputLabel,
  Alert,
  Divider
} from '@mui/material';
import { getAllReports, getReport, updateReportApproval, getDashboardData, sendExcelApproval, getExcelApprovalStatus } from '@/services/api';

export default function ApiTestPage() {
  const [results, setResults] = useState<any>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [selectedMonth, setSelectedMonth] = useState(3);
  const [memo, setMemo] = useState('테스트 메모');

  const handleApiCall = async (apiFunction: () => Promise<any>, description: string) => {
    setLoading(true);
    setError(null);
    try {
      const result = await apiFunction();
      setResults({
        description,
        data: result,
        timestamp: new Date().toISOString()
      });
    } catch (err: any) {
      setError(err.message || '알 수 없는 오류가 발생했습니다.');
    } finally {
      setLoading(false);
    }
  };

  return (
    <Container maxWidth="lg">
      <Typography variant="h4" gutterBottom>
        API 엔드포인트 테스트
      </Typography>

      <Box display="flex" flexDirection="column" gap={2}>
        {/* API 연결 테스트 */}
        <Button
          variant="contained"
          color="success"
          onClick={() => handleApiCall(
            () => fetch('/api/test?message=웹에서 테스트').then(res => res.json()),
            'GET /api/test - VBA 연동 테스트 API'
          )}
          disabled={loading}
        >
          🔧 API 연결 테스트 (VBA용)
        </Button>

        <Divider />

        {/* 전체 레포트 조회 */}
        <Button
          variant="contained"
          onClick={() => handleApiCall(getAllReports, 'GET /api/reports - 전체 레포트 조회')}
          disabled={loading}
        >
          전체 레포트 조회
        </Button>

        {/* 특정 월 레포트 조회 */}
        <Box display="flex" gap={2} alignItems="center">
          <FormControl size="small">
            <InputLabel>월</InputLabel>
            <Select
              value={selectedMonth}
              label="월"
              onChange={(e) => setSelectedMonth(e.target.value as number)}
            >
              <MenuItem value={1}>1월</MenuItem>
              <MenuItem value={2}>2월</MenuItem>
              <MenuItem value={3}>3월</MenuItem>
            </Select>
          </FormControl>
          <Button
            variant="contained"
            onClick={() => handleApiCall(
              () => getReport(selectedMonth),
              `GET /api/reports/${selectedMonth} - ${selectedMonth}월 레포트 조회`
            )}
            disabled={loading}
          >
            특정 월 레포트 조회
          </Button>
        </Box>

        {/* 승인 처리 */}
        <Box display="flex" gap={2} alignItems="center">
          <TextField
            size="small"
            label="메모"
            value={memo}
            onChange={(e) => setMemo(e.target.value)}
          />
          <Button
            variant="contained"
            color="primary"
            onClick={() => handleApiCall(
              () => updateReportApproval(selectedMonth, 'approved', memo, '테스터'),
              `PUT /api/reports/${selectedMonth} - ${selectedMonth}월 레포트 승인`
            )}
            disabled={loading}
          >
            승인 처리
          </Button>
          <Button
            variant="contained"
            color="error"
            onClick={() => handleApiCall(
              () => updateReportApproval(selectedMonth, 'rejected', memo, '테스터'),
              `PUT /api/reports/${selectedMonth} - ${selectedMonth}월 레포트 반려`
            )}
            disabled={loading}
          >
            반려 처리
          </Button>
        </Box>

        {/* 대시보드 데이터 조회 */}
        <Button
          variant="contained"
          onClick={() => handleApiCall(getDashboardData, 'GET /api/dashboard - 대시보드 데이터 조회')}
          disabled={loading}
        >
          대시보드 데이터 조회
        </Button>

        {/* 재무 데이터 전송 테스트 */}
        <Button
          variant="outlined"
          color="primary"
          onClick={() => handleApiCall(
            () => fetch('/api/reports/submit', {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({
                year: 2024,
                month: selectedMonth,
                salesRevenue: 10000000,
                otherIncome: 500000,
                rentExpense: 2000000,
                laborExpense: 3000000,
                materialExpense: 1500000,
                operatingExpense: 1000000,
                otherExpense: 500000,
                cashBalance: 50000000,
                submittedBy: '테스터'
              })
            }).then(res => res.json()),
            `POST /api/reports/submit - ${selectedMonth}월 재무 데이터 전송`
          )}
          disabled={loading}
        >
          📤 재무 데이터 전송 테스트
        </Button>

        {/* 6월 데이터 전송 테스트 */}
        <Button
          variant="outlined"
          color="secondary"
          onClick={() => handleApiCall(
            () => fetch('/api/reports/submit', {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({
                year: 2024,
                month: 6,
                salesRevenue: 15000000,
                otherIncome: 800000,
                rentExpense: 2500000,
                laborExpense: 3500000,
                materialExpense: 2000000,
                operatingExpense: 1200000,
                otherExpense: 600000,
                cashBalance: 60000000,
                submittedBy: '6월 테스터'
              })
            }).then(res => res.json()),
            'POST /api/reports/submit - 6월 데이터 전송 테스트'
          )}
          disabled={loading}
        >
          📅 6월 데이터 전송 테스트
        </Button>

        <Divider />

        {/* 엑셀 VBA 연동 테스트 */}
        <Typography variant="h6">엑셀 VBA 연동 테스트</Typography>
        
        <Button
          variant="outlined"
          onClick={() => handleApiCall(
            () => sendExcelApproval(selectedMonth, 2024, 'approved', '엑셀에서 승인', 'Excel VBA', 'v1.0'),
            `POST /api/excel - 엑셀에서 ${selectedMonth}월 승인 전송`
          )}
          disabled={loading}
        >
          엑셀 승인 전송
        </Button>

        <Button
          variant="outlined"
          onClick={() => handleApiCall(
            () => getExcelApprovalStatus(selectedMonth),
            `GET /api/excel?month=${selectedMonth} - 엑셀 승인 상태 조회`
          )}
          disabled={loading}
        >
          엑셀 승인 상태 조회
        </Button>
      </Box>

      {/* 결과 표시 */}
      {error && (
        <Alert severity="error" sx={{ mt: 3 }}>
          {error}
        </Alert>
      )}

      {results && (
        <Paper sx={{ mt: 3, p: 2 }}>
          <Typography variant="h6" gutterBottom>
            {results.description}
          </Typography>
          <Typography variant="caption" color="textSecondary">
            {results.timestamp}
          </Typography>
          <Box
            component="pre"
            sx={{
              mt: 2,
              p: 2,
              bgcolor: 'grey.100',
              borderRadius: 1,
              overflow: 'auto',
              fontSize: '0.875rem'
            }}
          >
            {JSON.stringify(results.data, null, 2)}
          </Box>
        </Paper>
      )}

      {loading && (
        <Box display="flex" justifyContent="center" mt={3}>
          <Typography>API 호출 중...</Typography>
        </Box>
      )}
    </Container>
  );
} 