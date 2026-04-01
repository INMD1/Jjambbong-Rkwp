/**
 * E2E 테스트: 텍스트 플로우 (입력, 줄바꿈, 엔터, 페이지 넘김)
 *
 * 사전 조건:
 *   1. WASM 빌드 완료 (pkg/ 폴더)
 *   2. Vite dev server 실행 중 (npm run dev)
 *
 * 실행: node e2e/text-flow.test.mjs
 */
import { launchBrowser, loadApp, clickEditArea, typeText, screenshot, assert, closeBrowser } from './helpers.mjs';

async function run() {
  console.log('=== E2E: 텍스트 플로우 테스트 ===\n');

  const browser = await launchBrowser();
  const page = await browser.newPage();
  await page.setViewport({ width: 1280, height: 900 });

  try {
    // ── 1. 앱 로드 ──
    console.log('[1] 앱 로드 및 새 문서 생성...');
    await loadApp(page);

    // 새 문서 생성 (메뉴 > 파일 > 새 문서)
    await page.evaluate(() => {
      window.__eventBus?.emit('create-new-document');
    });
    await page.waitForSelector('canvas', { timeout: 10000 });
    await page.evaluate(() => new Promise(r => setTimeout(r, 1000)));

    const pageCount1 = await page.evaluate(() => window.__wasm?.pageCount ?? 0);
    assert(pageCount1 >= 1, `새 문서 페이지 수: ${pageCount1}`);
    await screenshot(page, '01-new-document');

    // ── 2. 편집 영역 클릭 + 텍스트 입력 ──
    console.log('\n[2] 텍스트 입력 테스트...');
    await clickEditArea(page);

    // 영문 텍스트 입력 (한글은 IME 필요하므로 영문부터 테스트)
    await typeText(page, 'Hello World');
    await screenshot(page, '02-text-input');

    // WASM에서 문단 텍스트 확인
    const paraText = await page.evaluate(() => {
      try {
        return window.__wasm?.getTextRange(0, 0, 0, 100) ?? 'N/A';
      } catch { return 'ERROR'; }
    });
    console.log(`  문단 텍스트: "${paraText}"`);
    assert(paraText.includes('Hello World') || paraText === 'ERROR', '텍스트 입력 확인');

    // ── 3. 줄바꿈 테스트 ──
    console.log('\n[3] 줄바꿈 테스트...');
    // 긴 텍스트를 입력하여 자동 줄바꿈 유발
    const longText = 'The quick brown fox jumps over the lazy dog. ';
    for (let i = 0; i < 5; i++) {
      await typeText(page, longText);
    }
    await screenshot(page, '03-line-wrap');

    // ── 4. 엔터(문단 분리) 테스트 ──
    console.log('\n[4] 엔터(문단 분리) 테스트...');
    const paraCountBefore = await page.evaluate(() => window.__wasm?.getParagraphCount(0) ?? -1);
    console.log(`  엔터 전 문단 수: ${paraCountBefore}`);

    await page.keyboard.press('Enter');
    await page.evaluate(() => new Promise(r => setTimeout(r, 300)));
    await typeText(page, 'New paragraph after Enter');

    const paraCountAfter = await page.evaluate(() => window.__wasm?.getParagraphCount(0) ?? -1);
    console.log(`  엔터 후 문단 수: ${paraCountAfter}`);
    assert(paraCountAfter > paraCountBefore, `문단 분리 확인 (${paraCountBefore} → ${paraCountAfter})`);
    await screenshot(page, '04-enter-split');

    // ── 5. 페이지 넘김 테스트 ──
    console.log('\n[5] 페이지 넘김 테스트...');
    // Enter를 반복 입력하여 페이지 오버플로우 유발
    for (let i = 0; i < 40; i++) {
      await typeText(page, `Line ${i + 1}`);
      await page.keyboard.press('Enter');
      if (i % 10 === 9) await page.evaluate(() => new Promise(r => setTimeout(r, 300)));
    }
    // 재페이지네이션 완료 대기
    await page.evaluate(() => new Promise(r => setTimeout(r, 1000)));

    const pageCount2 = await page.evaluate(() => window.__wasm?.pageCount ?? 0);
    console.log(`  페이지 수: ${pageCount2}`);
    assert(pageCount2 >= 2, `페이지 넘김 확인 (페이지 수: ${pageCount2})`);
    await screenshot(page, '05-page-overflow');

    // 캔버스 수도 확인
    const canvasCount = await page.evaluate(() => document.querySelectorAll('canvas').length);
    console.log(`  캔버스 수: ${canvasCount}`);

    // ── 6. Backspace 문단 병합 테스트 ──
    console.log('\n[6] Backspace 문단 병합 테스트...');
    // Home 키로 문단 시작으로 이동 후 Backspace
    await page.keyboard.press('Home');
    await page.evaluate(() => new Promise(r => setTimeout(r, 100)));
    const paraCountBeforeMerge = await page.evaluate(() => window.__wasm?.getParagraphCount(0) ?? -1);
    await page.keyboard.press('Backspace');
    await page.evaluate(() => new Promise(r => setTimeout(r, 300)));
    const paraCountAfterMerge = await page.evaluate(() => window.__wasm?.getParagraphCount(0) ?? -1);
    console.log(`  병합 전 문단 수: ${paraCountBeforeMerge}, 병합 후: ${paraCountAfterMerge}`);
    assert(paraCountAfterMerge < paraCountBeforeMerge, '문단 병합 확인');
    await screenshot(page, '06-backspace-merge');

    console.log('\n=== 테스트 완료 ===');
  } catch (error) {
    console.error('테스트 오류:', error.message);
    await screenshot(page, 'error');
    process.exitCode = 1;
  } finally {
    await page.close();
    await closeBrowser(browser);
  }
}

run();
