#!/bin/bash

# AWS ECS 배포 스크립트

# 설정 변수들
AWS_REGION="ap-northeast-2"  # 서울 리전
AWS_ACCOUNT_ID="YOUR_AWS_ACCOUNT_ID"  # AWS 계정 ID를 입력하세요
ECR_REPO_NAME="sales-report-dashboard"
IMAGE_TAG="latest"
CLUSTER_NAME="sales-report-cluster"
SERVICE_NAME="sales-report-service"

echo "🚀 AWS ECS 배포 시작..."

# 1. AWS CLI 로그인 확인
echo "1. AWS CLI 인증 확인 중..."
aws sts get-caller-identity

# 2. ECR 레포지토리 생성 (없으면)
echo "2. ECR 레포지토리 생성 중..."
aws ecr describe-repositories --repository-names $ECR_REPO_NAME --region $AWS_REGION || \
aws ecr create-repository --repository-name $ECR_REPO_NAME --region $AWS_REGION

# 3. ECR 로그인
echo "3. ECR 로그인 중..."
aws ecr get-login-password --region $AWS_REGION | docker login --username AWS --password-stdin $AWS_ACCOUNT_ID.dkr.ecr.$AWS_REGION.amazonaws.com

# 4. Docker 이미지 빌드
echo "4. Docker 이미지 빌드 중..."
docker build -t $ECR_REPO_NAME:$IMAGE_TAG .

# 5. Docker 이미지 태깅
echo "5. Docker 이미지 태깅 중..."
docker tag $ECR_REPO_NAME:$IMAGE_TAG $AWS_ACCOUNT_ID.dkr.ecr.$AWS_REGION.amazonaws.com/$ECR_REPO_NAME:$IMAGE_TAG

# 6. ECR에 푸시
echo "6. ECR에 이미지 푸시 중..."
docker push $AWS_ACCOUNT_ID.dkr.ecr.$AWS_REGION.amazonaws.com/$ECR_REPO_NAME:$IMAGE_TAG

echo "✅ Docker 이미지가 ECR에 성공적으로 푸시되었습니다!"
echo "📝 이미지 URI: $AWS_ACCOUNT_ID.dkr.ecr.$AWS_REGION.amazonaws.com/$ECR_REPO_NAME:$IMAGE_TAG"

echo "🔧 다음 단계:"
echo "1. ECS 클러스터 생성"
echo "2. 태스크 정의 생성"
echo "3. ECS 서비스 생성"
echo "4. ALB (Application Load Balancer) 설정"

# ECS 클러스터 생성 (선택적)
read -p "ECS 클러스터를 생성하시겠습니까? (y/n): " -n 1 -r
echo
if [[ $REPLY =~ ^[Yy]$ ]]
then
    echo "7. ECS 클러스터 생성 중..."
    aws ecs create-cluster --cluster-name $CLUSTER_NAME --region $AWS_REGION
    echo "✅ ECS 클러스터 '$CLUSTER_NAME'이 생성되었습니다!"
fi 