FROM node:22-alpine

WORKDIR /app

# 1. 패키지 파일 복사
COPY package.json package-lock.json* ./

# 2. hwpkit-extension 이 외부 경로에 있어 오류가 발생할 수 있으므로, 
# 우선 더미 디렉토리를 만들어 npm install 이 실패하지 않도록 함
RUN mkdir -p ../hwpkit-extension && touch ../hwpkit-extension/package.json && echo '{"name":"hwpkit-extension","version":"0.0.0"}' > ../hwpkit-extension/package.json

# 3. 의존성 설치 (devDependencies 포함)
RUN npm install

# 4. 전체 소스 복사
COPY . .

# 포트 노출
EXPOSE 7700

# Vite dev 서버 실행 (npx를 통해 경로 문제 방지 및 호스트 노출)
CMD ["npx", "vite", "--host", "0.0.0.0"]

