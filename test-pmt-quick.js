import { chromium } from 'playwright';

(async () => {
  const browser = await chromium.launch({ headless: true });
  const page = await browser.newPage();
  
  // コンソールログをキャプチャ
  const logs = [];
  page.on('console', msg => {
    const text = msg.text();
    logs.push(`[${msg.type()}] ${text}`);
    // PMT関連のログのみ表示
    if (text.includes('PMT') || text.includes('matchFormula')) {
      console.log(`[${msg.type()}] ${text}`);
    }
  });
  
  await page.goto('http://localhost:5173/demo');
  await page.click('text=個別関数デモ');
  await page.waitForLoadState('networkidle');
  
  // 財務カテゴリを選択
  const categoryButton = page.locator('button').filter({ hasText: '07. 財務' });
  await categoryButton.first().click();
  await page.waitForTimeout(500);
  
  // PMT関数を選択
  const buttons = await page.locator('button').all();
  for (const button of buttons) {
    const text = await button.textContent();
    if (text && text.includes('PMT') && text.includes('ローン定期支払額')) {
      await button.click();
      await page.waitForTimeout(2000);
      break;
    }
  }
  
  // D2の値を取得（PMTの結果セル）
  const d2Cell = page.locator('table tbody tr').nth(2).locator('td').nth(3);
  const d2Value = await d2Cell.textContent();
  console.log('\nD2 value:', d2Value);
  
  // すべてのログを出力
  console.log('\n--- All console logs ---');
  logs.forEach(log => console.log(log));
  
  await browser.close();
})();