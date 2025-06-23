'use client';

import { useState } from 'react';
import { 
  Container, 
  Typography, 
  Button, 
  Card, 
  CardContent, 
  Box, 
  Alert, 
  TextField,
  Divider,
  Chip
} from '@mui/material';

export default function S3TestPage() {
  const [result, setResult] = useState<string>('');
  const [loading, setLoading] = useState(false);
  const [s3Config, setS3Config] = useState<any>(null);

  // S3 설정 정보 조회
  const getS3Config = async () => {
    setLoading(true);
    try {
      const response = await fetch('/api/s3-config');
      const data = await response.json();
      
      if (data.success) {
        setS3Config(data.s3Config);
        setResult(JSON.stringify(data, null, 2));
      } else {
        setResult(`오류: ${data.error}`);
      }
    } catch (error) {
      setResult(`오류: ${error}`);
    } finally {
      setLoading(false);
    }
  };

  // S3 연결 테스트
  const testS3Connection = async (action: string) => {
    setLoading(true);
    try {
      const response = await fetch('/api/s3-config', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({ action }),
      });

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
        🔧 S3 설정 및 연결 테스트
      </Typography>
      
      <Alert severity="info" sx={{ mb: 3 }}>
        이 페이지에서 S3 설정을 확인하고 연결을 테스트할 수 있습니다.
        환경 변수 S3_BUCKET_NAME, S3_REGION, S3_ACCESS_KEY_ID, S3_SECRET_ACCESS_KEY가 설정되어 있어야 합니다.
      </Alert>

      <Box sx={{ mb: 3, display: 'flex', gap: 2, flexWrap: 'wrap' }}>
        <Button 
          variant="contained" 
          color="primary" 
          onClick={getS3Config}
          disabled={loading}
        >
          🔍 S3 설정 정보 조회
        </Button>
        
        <Button 
          variant="outlined" 
          color="secondary" 
          onClick={() => testS3Connection('test-connection')}
          disabled={loading}
        >
          🔗 S3 연결 테스트
        </Button>
        
        <Button 
          variant="outlined" 
          color="warning" 
          onClick={() => testS3Connection('test-write')}
          disabled={loading}
        >
          ✏️ S3 쓰기 테스트
        </Button>
        
        <Button 
          variant="outlined" 
          color="info" 
          onClick={() => testS3Connection('test-read')}
          disabled={loading}
        >
          📖 S3 읽기 테스트
        </Button>
      </Box>

      {/* S3 설정 정보 표시 */}
      {s3Config && (
        <Card sx={{ mb: 3 }}>
          <CardContent>
            <Typography variant="h6" gutterBottom>
              📋 S3 설정 정보
            </Typography>
            <Box sx={{ display: 'flex', gap: 1, flexWrap: 'wrap', mb: 2 }}>
              <Chip 
                label={`버킷: ${s3Config.bucketName}`} 
                color="primary" 
                variant="outlined" 
              />
              <Chip 
                label={`리전: ${s3Config.region}`} 
                color="secondary" 
                variant="outlined" 
              />
              <Chip 
                label={`파일: ${s3Config.fileKey}`} 
                color="info" 
                variant="outlined" 
              />
              <Chip 
                label={s3Config.hasCredentials ? "✅ 인증정보 있음" : "❌ 인증정보 없음"} 
                color={s3Config.hasCredentials ? "success" : "error"} 
                variant="outlined" 
              />
            </Box>
            <Typography variant="body2" color="text.secondary">
              액세스 키: {s3Config.accessKeyId}
            </Typography>
          </CardContent>
        </Card>
      )}

      <Divider sx={{ my: 3 }} />

      {/* 환경 변수 설정 가이드 */}
      <Card sx={{ mb: 3 }}>
        <CardContent>
          <Typography variant="h6" gutterBottom>
            ⚙️ 환경 변수 설정 가이드
          </Typography>
          <Typography variant="body2" paragraph>
            프로젝트 루트에 <code>.env.local</code> 파일을 생성하고 다음 내용을 추가하세요:
          </Typography>
          <Box component="pre" sx={{ 
            backgroundColor: '#f5f5f5', 
            p: 2, 
            borderRadius: 1, 
            overflow: 'auto',
            fontSize: '0.875rem'
          }}>
{`# S3 설정
S3_BUCKET_NAME=your-bucket-name
S3_REGION=ap-northeast-2
S3_ACCESS_KEY_ID=your_access_key_id
S3_SECRET_ACCESS_KEY=your_secret_access_key`}
          </Box>
          <Alert severity="warning" sx={{ mt: 2 }}>
            <strong>주의:</strong> .env.local 파일은 .gitignore에 포함되어 있어야 하며, 
            실제 액세스 키와 시크릿 키는 절대 공개 저장소에 커밋하지 마세요.
          </Alert>
        </CardContent>
      </Card>

      {/* 결과 표시 */}
      <Card>
        <CardContent>
          <Typography variant="h6" gutterBottom>
            📊 테스트 결과
          </Typography>
          <TextField
            fullWidth
            multiline
            rows={10}
            value={result}
            onChange={(e) => setResult(e.target.value)}
            variant="outlined"
            placeholder="테스트 결과가 여기에 표시됩니다..."
            InputProps={{
              readOnly: true,
            }}
          />
        </CardContent>
      </Card>
    </Container>
  );
} 