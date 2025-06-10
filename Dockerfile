# Node.js 18 Alpine 이미지 사용
FROM node:18-alpine AS base

# 의존성 설치를 위한 단계
FROM base AS deps
RUN apk add --no-cache libc6-compat
WORKDIR /app

# 패키지 파일들을 복사하고 의존성 설치
COPY package.json pnpm-lock.yaml* ./
RUN corepack enable pnpm && pnpm i --frozen-lockfile

# 소스 코드 빌드를 위한 단계
FROM base AS builder
WORKDIR /app
COPY --from=deps /app/node_modules ./node_modules
COPY . .

# Next.js 텔레메트리 비활성화
ENV NEXT_TELEMETRY_DISABLED 1

# pnpm 활성화 및 빌드
RUN corepack enable pnpm && pnpm build

# 운영 이미지, 소스 코드를 복사하고 Next.js 시작
FROM base AS runner
WORKDIR /app

ENV NODE_ENV production
ENV NEXT_TELEMETRY_DISABLED 1

# 사용자 생성
RUN addgroup --system --gid 1001 nodejs
RUN adduser --system --uid 1001 nextjs

# 빌드된 파일들 복사
COPY --from=builder /app/public ./public

# 빌드 결과물 복사
COPY --from=builder --chown=nextjs:nodejs /app/.next/standalone ./
COPY --from=builder --chown=nextjs:nodejs /app/.next/static ./.next/static

# 데이터 디렉토리 생성 및 권한 설정
RUN mkdir -p ./data && chown nextjs:nodejs ./data

USER nextjs

EXPOSE 3000

ENV PORT 3000
ENV HOSTNAME "0.0.0.0"

# 서버 시작
CMD ["node", "server.js"] 