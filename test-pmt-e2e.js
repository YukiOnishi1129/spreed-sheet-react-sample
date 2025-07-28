import { chromium } from 'playwright';

(async () => {
  const browser = await chromium.launch({ headless: false });
  const page = await browser.newPage();
  
  // コンソールログをキャプチャ
  page.on('console', msg => {
    console.log(`[${msg.type()}] ${msg.text()}`);
  });
  
  await page.goto('http://localhost:5173/');
  
  // セルA1に100000を入力
  await page.click('td[data-testid="cell-0-0"]');
  await page.keyboard.type('100000');
  await page.keyboard.press('Enter');
  
  // セルB1に0.005を入力
  await page.click('td[data-testid="cell-0-1"]');
  await page.keyboard.type('0.005');
  await page.keyboard.press('Enter');
  
  // セルC1に12を入力
  await page.click('td[data-testid="cell-0-2"]');
  await page.keyboard.type('12');
  await page.keyboard.press('Enter');
  
  // セルD1にPMT関数を入力
  await page.click('td[data-testid="cell-0-3"]');
  await page.keyboard.type('=PMT(B1,C1,-A1)');
  await page.keyboard.press('Enter');
  
  // 少し待つ
  await page.waitForTimeout(3000);
  
  // D1の値を取得
  const d1Value = await page.locator('td[data-testid="cell-0-3"] input').inputValue();
  console.log('D1 value:', d1Value);
  
  // ブラウザを開いたままにする
  await page.waitForTimeout(10000);
  
  await browser.close();
})();