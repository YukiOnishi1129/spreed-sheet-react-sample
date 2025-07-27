import { test, expect } from '@playwright/test';
import { allIndividualFunctionTests } from '../src/data/individualFunctionTests';

test.describe('個別関数の計算結果検証', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/demo');
    
    // 個別関数デモモードに切り替え
    await page.click('text=個別関数デモ');
    
    // ページが完全に読み込まれるまで待機
    await page.waitForLoadState('networkidle');
  });

  // すべての関数をテスト
  allIndividualFunctionTests.forEach((functionTest) => {
    test(`${functionTest.name} - ${functionTest.description}`, async ({ page }) => {
      // モーダルを開く
      await page.click('text=関数を選択');
      
      // カテゴリーを選択
      const categoryButton = page.getByRole('button', { name: new RegExp(functionTest.category) });
      await categoryButton.click();
      
      // 関数を選択
      const functionButton = page.getByRole('button', { name: new RegExp(`${functionTest.name}.*${functionTest.description}`) });
      await functionButton.click();
      
      // スプレッドシートが更新されるまで待機
      await page.waitForTimeout(1000);
      
      // 期待値がある場合はセルの値を検証
      if (functionTest.expectedValues) {
        for (const [cellAddress, expectedValue] of Object.entries(functionTest.expectedValues)) {
          // セルアドレスから行と列を計算 (例: 'B2' -> row: 1, col: 1)
          const col = cellAddress.charCodeAt(0) - 65; // A=0, B=1, ...
          const row = parseInt(cellAddress.substring(1)) - 1; // 1-indexed to 0-indexed
          
          // react-spreadsheetのセル要素を取得
          const cellSelector = `.cell-viewer:nth-child(${row + 1}) .cell:nth-child(${col + 1})`;
          const cellElement = page.locator(cellSelector);
          
          // セルの値を取得
          const cellValue = await cellElement.textContent();
          
          // 数値の場合は小数点を考慮した比較
          if (typeof expectedValue === 'number') {
            const actualNumber = parseFloat(cellValue || '0');
            expect(actualNumber).toBeCloseTo(expectedValue, 2);
          } else {
            expect(cellValue).toBe(String(expectedValue));
          }
        }
      }
    });
  });

  // カテゴリー別の関数数を確認
  test('カテゴリー別の関数数が正しい', async ({ page }) => {
    await page.click('text=関数を選択');
    
    // 各カテゴリーの関数数を確認
    const categories = await page.locator('button[class*="rounded-lg"]').filter({ hasText: /^\d{2}\./ }).all();
    
    for (const category of categories) {
      const text = await category.textContent();
      const match = text?.match(/\((\d+)\)/);
      if (match) {
        const count = parseInt(match[1]);
        expect(count).toBeGreaterThan(0);
      }
    }
  });

  // パフォーマンステスト
  test('大量の関数でもパフォーマンスが保たれる', async ({ page }) => {
    const startTime = Date.now();
    
    // モーダルを開く
    await page.click('text=関数を選択');
    
    // すべてのカテゴリーを表示
    await page.click('text=すべて');
    
    // モーダルが表示されるまでの時間を測定
    await page.waitForSelector('text=512 個の関数');
    
    const loadTime = Date.now() - startTime;
    expect(loadTime).toBeLessThan(3000); // 3秒以内
  });
});

// 特定の複雑な関数のテスト
test.describe('複雑な関数の詳細テスト', () => {
  test('VLOOKUP関数が正しく動作する', async ({ page }) => {
    await page.goto('/demo');
    await page.click('text=個別関数デモ');
    await page.click('text=関数を選択');
    
    // 検索カテゴリーを選択
    await page.click('text=06. 検索');
    
    // VLOOKUP関数を選択
    await page.click('button:has-text("VLOOKUP")');
    
    // 特定のセルの値を確認
    const resultCell = page.locator('.cell-viewer:nth-child(3) .cell:nth-child(3)'); // C3
    await expect(resultCell).toHaveText('20');
  });

  test('日付関数が正しく動作する', async ({ page }) => {
    await page.goto('/demo');
    await page.click('text=個別関数デモ');
    await page.click('text=関数を選択');
    
    // 日付カテゴリーを選択
    await page.click('text=04. 日付');
    
    // DATEDIF関数を選択
    await page.click('button:has-text("DATEDIF")');
    
    // 結果を確認
    await page.waitForTimeout(1000);
    
    // セルの値が数値であることを確認
    const resultCell = page.locator('.cell-viewer:nth-child(2) .cell:nth-child(4)'); // D2
    const value = await resultCell.textContent();
    expect(parseInt(value || '0')).toBeGreaterThan(0);
  });
});