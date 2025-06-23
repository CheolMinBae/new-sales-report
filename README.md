# 재무 리포트 대시보드

월별로 제출되는 재무 내용을 확인하고 해당 내용을 승인/반려하는 재무 관리 시스템

## 주요 기능

### 1. 전체 누적 대시보드
- **핵심 성과 지표**: 총 매출, 총 매입, 현금 잔고, 순이익률
- **차트 분석**: 월별 매출현황, 비용구조 분석, 월간 통계, 수익/손실 분석
- **카테고리별 지출**: 임대료, 인건비, 재료비, 운영비, 기타
- **월별 상세 현황**: 테이블 형태의 상세 데이터

### 2. 월별 레포트 관리
- **월 선택**: 월별 데이터 선택 기능
- **승인 시스템**: 승인/반려 상태 관리
- **메모 기능**: 승인/반려 시 메모 작성
- **이전 달 대비**: +/- 변화율 표시
- **주간 분석**: 주간 매출현황, 비용구조, 통계

### 3. Excel VBA 연동
- **자동 데이터 수집**: 시트에서 자동으로 데이터 수집
- **전체 년도 일괄 전송**: 2020~2025년 데이터 일괄 처리
- **실시간 상태 확인**: 승인 상태 실시간 조회
- **오류 처리**: 상세한 오류 메시지 및 디버깅

### 4. AWS S3 데이터 저장
- **클라우드 저장**: AWS S3를 통한 안전한 데이터 저장
- **확장성**: 대용량 데이터 처리 지원
- **백업**: 자동 데이터 백업 및 복구
- **보안**: IAM 기반 접근 제어

## 기술 스택

- **Frontend**: Next.js 14 + TypeScript
- **UI Library**: Material-UI (MUI)
- **Charts**: Recharts / MUI X Charts
- **데이터 저장**: JSON 파일 + AWS S3
- **API**: Next.js API Routes
- **엑셀 연동**: VBA + REST API
- **클라우드**: AWS S3, AWS SDK

## 설치 및 실행

### 사전 요구사항
- Node.js 18+ 설치 필요
- pnpm 패키지 매니저 (권장)
- AWS 계정 (S3 사용 시)

### 설치
```bash
# 의존성 설치
pnpm install

# AWS SDK 설치 (S3 사용 시)
pnpm add @aws-sdk/client-s3 @aws-sdk/s3-request-presigner
```

### 환경 변수 설정

프로젝트 루트에 `.env.local` 파일을 생성하고 다음 내용을 추가하세요:

```env
# API 설정
NEXT_PUBLIC_API_URL=http://turfintra.com:3002

# S3 설정 (선택사항 - S3 사용 시)
S3_BUCKET_NAME=your-bucket-name
S3_REGION=ap-northeast-2
S3_ACCESS_KEY_ID=your_access_key_id
S3_SECRET_ACCESS_KEY=your_secret_access_key
```

### S3 설정 (선택사항)

S3를 사용하여 데이터를 저장하려면:

1. **AWS S3 버킷 생성**
   - AWS 콘솔에서 새 S3 버킷 생성
   - 버킷 이름을 `.env.local`의 `S3_BUCKET_NAME`에 설정

2. **IAM 사용자 생성**
   - S3 접근 권한을 가진 IAM 사용자 생성
   - 액세스 키와 시크릿 키를 `.env.local`에 설정

3. **S3 연결 테스트**
   - `/s3-test` 페이지에서 연결 테스트
   - 연결, 쓰기, 읽기 테스트 수행

### 개발 서버 실행
```bash
pnpm dev
```

브라우저에서 [http://localhost:3000](http://localhost:3000)로 접속

### 프로덕션 빌드
```bash
pnpm build
pnpm start
```

## 프로젝트 구조

```
src/
├── app/                    # Next.js 13+ App Router
│   ├── layout.tsx         # 루트 레이아웃
│   ├── page.tsx          # 메인 대시보드
│   ├── globals.css       # 글로벌 CSS
│   ├── api/              # API 엔드포인트
│   │   ├── bulk-data/    # 대용량 데이터 처리
│   │   ├── dashboard/    # 대시보드 데이터
│   │   ├── excel/        # Excel VBA 연동
│   │   ├── reports/      # 재무 리포트
│   │   └── s3-config/    # S3 설정 관리
│   ├── api-test/         # API 테스트 페이지
│   ├── test-bulk/        # 대용량 데이터 테스트
│   └── s3-test/          # S3 설정 테스트
├── components/           # 재사용 가능한 컴포넌트
├── lib/                 # 유틸리티 함수 (dataStore.ts)
├── types/               # TypeScript 타입 정의
└── utils/               # 유틸리티 함수
```

## 데이터 구조

### 월별 재무 데이터
- 매출 관련: 매출액, 기타수입
- 비용 관련: 임대료, 인건비, 재료비, 운영비, 기타
- 수익성 지표: 순이익, 순이익률
- 현금 관련: 현금잔고, 현금흐름 변화
- 승인 관련: 승인상태, 승인일시, 승인자, 메모

### 대용량 데이터 (Bulk Data)
```json
{
  "yearlyData": [
    {
      "year": 2024,
      "monthlyData": {
        "1월": {
          "salesRevenue": 50000000,
          "otherIncome": 5000000,
          "rentExpense": 10000000,
          "laborExpense": 15000000,
          "materialExpense": 8000000,
          "operatingExpense": 12000000,
          "otherExpense": 3000000,
          "cashBalance": 20000000
        }
      }
    }
  ],
  "submittedBy": "사용자명",
  "sheetName": "시트명",
  "submittedAt": "2024-01-01 12:00:00"
}
```

## API 엔드포인트

### 재무 데이터
- `GET /api/finance/monthly` - 월별 데이터 조회
- `POST /api/finance/monthly` - 월별 데이터 생성
- `PUT /api/finance/monthly/[id]` - 월별 데이터 수정
- `POST /api/finance/approve` - 승인/반려 처리

### 대용량 데이터
- `POST /api/bulk-data/submit` - 전체 년도 데이터 전송
- `GET /api/bulk-data` - 대용량 데이터 조회
- `GET /api/bulk-data/test` - 대용량 데이터 테스트

### Excel VBA 연동
- `POST /api/excel` - 승인/반려 처리
- `GET /api/excel` - 승인 상태 조회

### S3 관리
- `GET /api/s3-config` - S3 설정 정보
- `POST /api/s3-config` - S3 연결 테스트

### 디버깅
- `GET /api/debug` - 전체 데이터 및 설정 정보

## 엑셀 VBA 연동

### 주요 기능
- **자동 데이터 수집**: 시트에서 자동으로 데이터 수집
- **전체 년도 일괄 전송**: 2020~2025년 데이터 일괄 처리
- **실시간 상태 확인**: 승인 상태 실시간 조회
- **오류 처리**: 상세한 오류 메시지 및 디버깅

### VBA 코드 사용법
1. `VBA-통합버전.vba` 파일을 Excel에 import
2. 매크로 보안 설정 활성화
3. 데이터 입력 후 버튼 클릭으로 전송

### 데이터 전송 형식
```json
{
  "year": 2024,
  "month": 3,
  "salesRevenue": 280000000,
  "otherIncome": 3000000,
  "expenses": {
    "rent": 32000000,
    "labor": 38000000,
    "materials": 35000000,
    "operating": 33000000,
    "others": 28000000
  }
}
```

## 테스트 페이지

- `/api-test` - API 엔드포인트 테스트
- `/test-bulk` - 대용량 데이터 전송 테스트
- `/s3-test` - S3 설정 및 연결 테스트

## 문제 해결

### S3 연결 오류
1. 환경 변수 확인
2. IAM 권한 확인
3. 버킷 존재 여부 확인
4. `/s3-test` 페이지에서 연결 테스트

### VBA 전송 오류
1. API 서버 상태 확인
2. 네트워크 연결 확인
3. 데이터 형식 검증
4. 타임아웃 설정 확인

### 전체 년도 전송 오류
1. 시트 구조 확인
2. 데이터 수집 로직 검증
3. JSON 형식 검증
4. 서버 로그 확인

## 라이선스

MIT License 