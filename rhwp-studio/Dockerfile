FROM node:22-alpine

WORKDIR /app

# 전체 소스 복사 (package.json 포함)
COPY . .

# 의존성 설치
RUN npm install

# 포트 노출
EXPOSE 7700

# Vite dev 서버 실행 (호스트 노출)
CMD ["npm", "run", "dev", "--", "--host", "0.0.0.0"]
