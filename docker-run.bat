@echo off
chcp 65001 >nul

echo 🚀 Sales Report Dashboard Docker 실행 스크립트 (Windows)
echo ==================================================

REM 환경 변수 확인
if "%S3_ACCESS_KEY_ID%"=="" (
    echo ⚠️  S3_ACCESS_KEY_ID 환경 변수가 설정되지 않았습니다.
    echo 다음 환경 변수를 설정하세요:
    echo   - S3_ACCESS_KEY_ID
    echo   - S3_SECRET_ACCESS_KEY
    echo.
    echo 예시:
    echo   set S3_ACCESS_KEY_ID=your_access_key
    echo   set S3_SECRET_ACCESS_KEY=your_secret_key
    echo.
)

if "%S3_SECRET_ACCESS_KEY%"=="" (
    echo ⚠️  S3_SECRET_ACCESS_KEY 환경 변수가 설정되지 않았습니다.
    echo.
)

REM 기본 환경 변수 설정
if "%S3_BUCKET_NAME%"=="" set S3_BUCKET_NAME=sales-report-data
if "%S3_REGION%"=="" set S3_REGION=ap-northeast-2
if "%NODE_ENV%"=="" set NODE_ENV=production
if "%PORT%"=="" set PORT=3000

echo ✅ 환경 변수 설정:
echo   S3_BUCKET_NAME: %S3_BUCKET_NAME%
echo   S3_REGION: %S3_REGION%
echo   NODE_ENV: %NODE_ENV%
echo   PORT: %PORT%
echo.

REM Docker 이미지 빌드
echo 🔨 Docker 이미지 빌드 중...
docker build -t sales-report-dashboard .

if %ERRORLEVEL% EQU 0 (
    echo ✅ 이미지 빌드 완료
) else (
    echo ❌ 이미지 빌드 실패
    pause
    exit /b 1
)

REM 기존 컨테이너 중지 및 제거
echo 🛑 기존 컨테이너 정리 중...
docker stop sales-report-container 2>nul
docker rm sales-report-container 2>nul

REM 새 컨테이너 실행
echo 🚀 컨테이너 실행 중...
docker run -d ^
    --name sales-report-container ^
    -p 3000:3000 ^
    -e NODE_ENV=%NODE_ENV% ^
    -e PORT=%PORT% ^
    -e NEXT_PUBLIC_API_URL=http://turfintra.com:3002 ^
    -e S3_BUCKET_NAME=%S3_BUCKET_NAME% ^
    -e S3_REGION=%S3_REGION% ^
    -e S3_ACCESS_KEY_ID=%S3_ACCESS_KEY_ID% ^
    -e S3_SECRET_ACCESS_KEY=%S3_SECRET_ACCESS_KEY% ^
    --restart unless-stopped ^
    sales-report-dashboard

if %ERRORLEVEL% EQU 0 (
    echo ✅ 컨테이너 실행 완료
    echo.
    echo 🌐 애플리케이션 접속:
    echo   http://localhost:3000
    echo.
    echo 📊 테스트 페이지:
    echo   http://localhost:3000/s3-test
    echo   http://localhost:3000/api-test
    echo.
    echo 📝 로그 확인:
    echo   docker logs sales-report-container
    echo.
    echo 🛑 컨테이너 중지:
    echo   docker stop sales-report-container
) else (
    echo ❌ 컨테이너 실행 실패
    pause
    exit /b 1
)

pause 