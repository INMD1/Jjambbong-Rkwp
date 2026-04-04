# @rhwp/core

**알(R), 모두의 한글** — 브라우저에서 HWP 파일을 열어보세요

[![npm](https://img.shields.io/npm/v/@rhwp/core)](https://www.npmjs.com/package/@rhwp/core)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

Rust + WebAssembly 기반 HWP/HWPX 파서 & 렌더러입니다.
HWP 파일을 파싱하고 SVG로 렌더링하는 저수준 API를 제공합니다.

> 편집 기능(메뉴, 툴바, 서식)이 필요하면 **[@rhwp/editor](https://www.npmjs.com/package/@rhwp/editor)** 를 사용하세요.
> 3줄이면 완전한 HWP 에디터를 임베드할 수 있습니다.

| 패키지 | 용도 |
|--------|------|
| **@rhwp/core** (이 패키지) | WASM 파서/렌더러 — 직접 API 호출 |
| **@rhwp/editor** | 완전한 에디터 UI — iframe 임베드 |

## 빠른 시작 — 처음부터 따라하기

### 1. 프로젝트 생성

```bash
mkdir my-hwp-viewer
cd my-hwp-viewer
npm init -y
npm install @rhwp/core
npm install vite --save-dev
```

### 2. WASM 파일 복사

`@rhwp/core`에 포함된 WASM 바이너리를 웹 서버가 제공할 수 있는 위치에 복사합니다.

```bash
mkdir public
cp node_modules/@rhwp/core/rhwp_bg.wasm public/
```

### 3. HTML 작성 — `index.html`

```html
<!DOCTYPE html>
<html lang="ko">
<head>
  <meta charset="UTF-8" />
  <title>HWP 뷰어</title>
</head>
<body>
  <h1>HWP 뷰어</h1>
  <input type="file" id="file-input" accept=".hwp,.hwpx" />
  <p id="status">파일을 선택해주세요.</p>
  <div id="viewer"></div>
  <script type="module" src="/main.js"></script>
</body>
</html>
```

### 4. JavaScript 작성 — `main.js`

```javascript
import init, { HwpDocument } from '@rhwp/core';

// ① 텍스트 폭 측정 함수 등록 (필수 — 아래 "왜 필요한가?" 참고)
let ctx = null;
let lastFont = '';
globalThis.measureTextWidth = (font, text) => {
  if (!ctx) ctx = document.createElement('canvas').getContext('2d');
  if (font !== lastFont) { ctx.font = font; lastFont = font; }
  return ctx.measureText(text).width;
};

// ② WASM 초기화
await init({ module_or_path: '/rhwp_bg.wasm' });
document.getElementById('status').textContent = '준비 완료!';

// ③ 파일 선택 시 렌더링
document.getElementById('file-input').addEventListener('change', async (e) => {
  const file = e.target.files[0];
  if (!file) return;

  const buffer = new Uint8Array(await file.arrayBuffer());
  const doc = new HwpDocument(buffer);

  // SVG로 첫 페이지 렌더링
  const svg = doc.renderPageSvg(0);
  document.getElementById('viewer').innerHTML = svg;
  document.getElementById('status').textContent =
    `${file.name} — ${doc.pageCount()}페이지`;
});
```

### 5. 실행

```bash
npx vite --port 3000
```

브라우저에서 `http://localhost:3000` 을 열고 HWP 파일을 선택하면 렌더링됩니다.

## API

### 초기화

```javascript
import init, { HwpDocument } from '@rhwp/core';
await init({ module_or_path: '/rhwp_bg.wasm' });
```

### 문서 로드

```javascript
// 파일 입력에서 로드
const doc = new HwpDocument(new Uint8Array(buffer));

// fetch로 로드
const resp = await fetch('/sample.hwp');
const doc = new HwpDocument(new Uint8Array(await resp.arrayBuffer()));
```

### SVG 렌더링

```javascript
const svg = doc.renderPageSvg(0);   // 첫 페이지
document.getElementById('viewer').innerHTML = svg;
```

### 페이지 네비게이션

```javascript
const total = doc.pageCount();
for (let i = 0; i < total; i++) {
  const svg = doc.renderPageSvg(i);
  // ...
}
```

## 필수 설정: measureTextWidth

WASM 내부에서 텍스트 레이아웃(줄바꿈, 정렬 등)을 계산할 때
브라우저의 Canvas 텍스트 측정 API가 필요합니다.
**WASM 초기화 전에 반드시 등록**해야 합니다.

```javascript
let ctx = null;
let lastFont = '';
globalThis.measureTextWidth = (font, text) => {
  if (!ctx) ctx = document.createElement('canvas').getContext('2d');
  if (font !== lastFont) { ctx.font = font; lastFont = font; }
  return ctx.measureText(text).width;
};
```

### 왜 필요한가?

HWP 문서의 텍스트 배치(줄바꿈 위치, 양쪽 정렬 간격)를 정확하게 계산하려면
각 글자의 실제 렌더링 폭을 알아야 합니다.
WASM 내부에는 브라우저 폰트에 접근할 수 없으므로,
JavaScript의 `Canvas.measureText()`를 콜백으로 호출합니다.

## 지원 기능

- **HWP 5.0** (바이너리) + **HWPX** (XML) 파싱
- 문단, 표, 수식, 이미지, 차트, 도형 렌더링
- 페이지네이션 (다단, 표 행 분할)
- 그림 자르기(crop), 이미지 테두리선
- SVG 출력
- 머리말/꼬리말/바탕쪽/각주/미주

## 링크

- **[온라인 데모](https://edwardkim.github.io/rhwp/)**
- **[GitHub](https://github.com/edwardkim/rhwp)**
- **[@rhwp/editor](https://www.npmjs.com/package/@rhwp/editor)** — 에디터 UI 임베드
- **[VS Code 확장](https://marketplace.visualstudio.com/items?itemName=edwardkim.rhwp-vscode)**

## Notice

본 제품은 한글과컴퓨터의 한글 문서 파일(.hwp) 공개 문서를 참고하여 개발하였습니다.

## License

MIT
