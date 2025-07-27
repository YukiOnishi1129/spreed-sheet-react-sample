import { test, expect } from '@playwright/test';

test.describe('基本的な関数テスト', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/demo');
    await page.click('text=個別関数デモ');
    await page.waitForLoadState('networkidle');
  });

  test('SUM関数が正しく計算される', async ({ page }) => {
    // モーダルを開く
    await page.click('text=関数を選択');
    
    // 数学カテゴリーを選択
    await page.click('text=01. 数学・三角');
    
    // SUM関数を選択
    await page.click('button:has-text("SUM")');
    
    // スプレッドシートが更新されるまで待機
    await page.waitForTimeout(2000);
    
    // 選択された関数が表示される
    await expect(page.locator('text=選択中: SUM')).toBeVisible();
    
    // 期待値が表示されていることを確認
    await expect(page.locator('text=期待値: E2=100')).toBeVisible();
  });

  test('複数の関数を連続で選択できる', async ({ page }) => {
    const functions = ['SUM', 'AVERAGE', 'MAX', 'MIN'];
    
    for (const func of functions) {
      // モーダルを開く
      await page.click('text=関数を選択');
      
      // すべてのカテゴリーを表示
      await page.click('text=すべて');
      
      // 関数を選択
      await page.click(`button:has-text("${func}")`);
      
      // 選択された関数が表示される
      await expect(page.locator(`text=選択中: ${func}`)).toBeVisible();
      
      await page.waitForTimeout(500);
    }
  });
});