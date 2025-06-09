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

## 기술 스택

- **Frontend**: Next.js 14 + TypeScript
- **UI Library**: Material-UI (MUI)
- **Charts**: Recharts / MUI X Charts
- **데이터 저장**: JSON 파일
- **API**: Next.js API Routes
- **엑셀 연동**: VBA + REST API

## 설치 및 실행

### 사전 요구사항
- Node.js 18+ 설치 필요
- npm 또는 yarn 패키지 매니저

### 설치
```bash
# 의존성 설치
npm install

# 또는 yarn 사용시
yarn install
```

### 개발 서버 실행
```bash
npm run dev

# 또는 yarn 사용시
yarn dev
```

브라우저에서 [http://localhost:3000](http://localhost:3000)로 접속

### 프로덕션 빌드
```bash
npm run build
npm start
```

## 프로젝트 구조

```
src/
├── app/                    # Next.js 13+ App Router
│   ├── layout.tsx         # 루트 레이아웃
│   ├── page.tsx          # 메인 대시보드
│   ├── globals.css       # 글로벌 CSS
│   └── api/              # API 엔드포인트
├── components/           # 재사용 가능한 컴포넌트
├── data/                # 샘플 데이터 및 JSON 파일
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

### 일별 거래 내역
- 기본 정보: 날짜, 설명, 금액, 유형
- 분류: 카테고리, 상세내용
- 관련자: 고객, 업체

## API 엔드포인트

### 재무 데이터
- `GET /api/finance/monthly` - 월별 데이터 조회
- `POST /api/finance/monthly` - 월별 데이터 생성
- `PUT /api/finance/monthly/[id]` - 월별 데이터 수정
- `POST /api/finance/approve` - 승인/반려 처리

### 거래 내역
- `GET /api/transactions` - 거래 내역 조회
- `POST /api/transactions` - 거래 내역 생성

## 엑셀 VBA 연동

### VBA 코드 예시
```vb
' API 호출을 위한 VBA 함수
Function SubmitMonthlyData()
    ' 월별 데이터를 JSON 형태로 변환하여 API 전송
    ' ...
End Function

Function GetApprovalStatus()
    ' 승인 상태 조회
    ' ...
End Function
```

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

## 라이선스

MIT License 