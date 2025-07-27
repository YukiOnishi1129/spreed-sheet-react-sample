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
    
    // モーダルが開くのを待つ（固定背景が表示される）
    await page.waitForSelector('.fixed.inset-0.bg-black', { state: 'visible' });
    
    // 数学カテゴリーを選択
    await page.click('button:has-text("01. 数学")');
    
    // カテゴリーが変更されるのを待つ
    await page.waitForTimeout(500);
    
    // SUM関数を選択（より具体的なセレクタ）
    const sumButton = page.locator('button').filter({ hasText: 'SUM' }).filter({ hasText: '数値の合計を計算' });
    await sumButton.click();
    
    // スプレッドシートが更新されるまで待機
    await page.waitForTimeout(2000);
    
    // 選択された関数が表示される
    await expect(page.locator('text=選択中: SUM')).toBeVisible();
    
    // 期待値が表示されていることを確認
    await expect(page.locator('text=期待値: E2=100')).toBeVisible();
    
    // デバッグ用にスクリーンショットを取る
    await page.screenshot({ path: 'test-results/debug-sum-test.png', fullPage: true });
    
    // セルE2の実際の値を確認
    // 100という値を含むセルを探す（完全一致）
    const cellWith100 = page.locator('td').filter({ hasText: /^100$/ });
    await expect(cellWith100).toBeVisible();
  });

  test('複数の関数を連続で選択できる', async ({ page }) => {
    const functions = [
      { name: 'SUM', description: '数値の合計を計算' },
      { name: 'AVERAGE', description: '平均値を計算' },
      { name: 'MAX', description: '最大値を返す' },
      { name: 'MIN', description: '最小値を返す' }
    ];
    
    for (const func of functions) {
      // モーダルを開く
      await page.click('text=関数を選択');
      
      // モーダルが開くのを待つ
      await page.waitForSelector('.fixed.inset-0.bg-black', { state: 'visible' });
      
      // すべてのカテゴリーを表示
      await page.click('text=すべて');
      
      // 関数を選択（より具体的なセレクタ）
      const funcButton = page.locator('button').filter({ 
        hasText: func.name 
      }).filter({ 
        hasText: func.description 
      });
      await funcButton.click();
      
      // 選択された関数が表示される
      await expect(page.locator(`text=選択中: ${func.name}`)).toBeVisible();
      
      // モーダルが閉じるのを待つ
      await page.waitForSelector('.fixed.inset-0.bg-black', { state: 'hidden' });
      await page.waitForTimeout(500);
    }
  });
});