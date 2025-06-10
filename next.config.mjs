/** @type {import('next').NextConfig} */
const nextConfig = {
  // Docker 배포를 위한 standalone 모드 활성화
  output: 'standalone',
  
  // 이미지 최적화 설정
  images: {
    unoptimized: true
  },
  
  // 환경 변수 설정
  env: {
    CUSTOM_KEY: process.env.CUSTOM_KEY,
  }
};

export default nextConfig; 