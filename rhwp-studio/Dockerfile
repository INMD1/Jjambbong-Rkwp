FROM node:22-alpine

WORKDIR /app

# 1. 먼저 package.json 만 복사하여 의존성 설치
COPY package.json package-lock.json* ./
RUN npm install

# 2. 그 다음 전체 소스 복사
COPY . .

# 포트 노출
EXPOSE 7700

# Vite dev 서버 실행 (호스트 노출)
CMD ["npm", "run", "dev", "--", "--host", "0.0.0.0"]
