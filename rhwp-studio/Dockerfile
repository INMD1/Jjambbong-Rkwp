FROM node:22-alpine

WORKDIR /app

# 의존성 복사 및 설치
COPY package*.json ./
RUN npm install

# 소스 코드 복사
COPY . .

# 포트 노출
EXPOSE 7700

# Vite dev 서버 실행 (호스트 노출)
CMD ["npm", "run", "dev", "--", "--host", "0.0.0.0"]
