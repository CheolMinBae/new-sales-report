# 🚀 AWS Docker 배포 가이드

## 📋 사전 준비사항

### 1. AWS CLI 설치 및 설정
```bash
# AWS CLI 설치 (Windows)
curl "https://awscli.amazonaws.com/AWSCLIV2.msi" -o "AWSCLIV2.msi"
start AWSCLIV2.msi

# AWS 계정 설정
aws configure
```

### 2. Docker 설치
- [Docker Desktop](https://www.docker.com/products/docker-desktop/) 설치

### 3. 필요한 정보
- AWS 계정 ID
- AWS 액세스 키
- AWS 시크릿 키

## 🔧 로컬 테스트

### 1. Docker로 로컬 빌드 테스트
```bash
# 이미지 빌드
docker build -t sales-report-dashboard .

# 컨테이너 실행
docker run -p 3000:3000 -v ./data:/app/data sales-report-dashboard
```

### 2. docker-compose로 테스트
```bash
# 빌드 및 실행
docker-compose up --build

# 백그라운드 실행
docker-compose up -d --build

# 중지
docker-compose down
```

## 🌊 AWS 배포 (ECS Fargate)

### 1. 스크립트 설정
`deploy-aws.sh` 파일에서 다음 값을 수정:
```bash
AWS_ACCOUNT_ID="123456789012"  # 실제 AWS 계정 ID
```

### 2. 스크립트 실행 권한 부여
```bash
chmod +x deploy-aws.sh
```

### 3. 배포 스크립트 실행
```bash
./deploy-aws.sh
```

### 4. ECS 태스크 정의 설정
`ecs-task-definition.json` 파일에서 다음 값들을 수정:
- `YOUR_AWS_ACCOUNT_ID` → 실제 AWS 계정 ID
- `fs-xxxxxxxxx` → 실제 EFS 파일시스템 ID (선택사항)
- `fsap-xxxxxxxxx` → 실제 EFS 액세스 포인트 ID (선택사항)

### 5. ECS 서비스 생성
AWS 콘솔에서:
1. ECS → 클러스터 → sales-report-cluster
2. 서비스 탭 → 생성
3. 태스크 정의: sales-report-dashboard
4. 서비스명: sales-report-service
5. 원하는 작업 수: 1
6. 로드 밸런서 설정 (선택사항)

## 🔗 ALB (Application Load Balancer) 설정

### 1. ALB 생성
```bash
# VPC와 서브넷 정보 확인
aws ec2 describe-vpcs
aws ec2 describe-subnets

# ALB 생성
aws elbv2 create-load-balancer \
  --name sales-report-alb \
  --subnets subnet-xxxxxxxx subnet-yyyyyyyy \
  --security-groups sg-xxxxxxxx
```

### 2. 타겟 그룹 생성
```bash
aws elbv2 create-target-group \
  --name sales-report-targets \
  --protocol HTTP \
  --port 3000 \
  --vpc-id vpc-xxxxxxxx \
  --target-type ip \
  --health-check-path /api/debug
```

## 📊 모니터링 설정

### 1. CloudWatch 로그 그룹 생성
```bash
aws logs create-log-group --log-group-name /ecs/sales-report-dashboard
```

### 2. CloudWatch 대시보드 생성
AWS 콘솔에서 CloudWatch → 대시보드 → 생성

## 🔄 배포 업데이트

### 새 버전 배포
```bash
# 1. 코드 변경 후
# 2. 스크립트 재실행
./deploy-aws.sh

# 3. ECS 서비스 업데이트
aws ecs update-service \
  --cluster sales-report-cluster \
  --service sales-report-service \
  --force-new-deployment
```

## 📝 중요한 고려사항

### 1. 데이터 영속성
- **EFS (Elastic File System)** 사용 권장
- 컨테이너 재시작 시에도 VBA 전송 데이터 유지
- 다중 AZ 지원으로 고가용성 확보

### 2. 보안
- ECS 태스크 역할에 최소 권한 부여
- VPC 내부 통신으로 보안 강화
- HTTPS 사용 (ALB + SSL 인증서)

### 3. 비용 최적화
- Fargate Spot 인스턴스 사용 고려
- 사용하지 않는 시간대에 스케일링 조정
- CloudWatch 로그 보존 기간 설정

## 🆘 문제 해결

### 일반적인 문제들

#### 1. Docker 빌드 실패
```bash
# 캐시 없이 재빌드
docker build --no-cache -t sales-report-dashboard .
```

#### 2. ECS 태스크 시작 실패
```bash
# 로그 확인
aws logs get-log-events \
  --log-group-name /ecs/sales-report-dashboard \
  --log-stream-name ecs/sales-report-dashboard/TASK_ID
```

#### 3. VBA 연결 실패
- ECS 서비스의 퍼블릭 IP 확인
- 보안 그룹에서 3000 포트 인바운드 허용
- VBA 코드의 API_BASE_URL 업데이트

## 📞 지원

문제가 발생하면:
1. CloudWatch 로그 확인
2. ECS 태스크 상태 확인
3. 네트워크 연결 테스트
4. VBA 코드의 URL 확인

## 💰 예상 비용 (월간)

- **ECS Fargate**: 약 $15-30 (CPU 0.5, 메모리 1GB 기준)
- **ALB**: 약 $20-25
- **EFS**: 약 $3-10 (데이터량에 따라)
- **CloudWatch**: 약 $2-5

**총 예상 비용**: 약 $40-70/월 