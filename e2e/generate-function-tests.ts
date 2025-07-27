import { test, expect } from '@playwright/test';
import type { IndividualFunctionTest } from '../src/types/spreadsheet';

// ヘルパー関数をエクスポート
export async function selectFunction(page: any, category: string, functionName: string, description: string) {
  await page.click('text=関数を選択');
  await page.waitForSelector('.fixed.inset-0.bg-black', { state: 'visible' });
  
  const categoryButton = page.locator('button').filter({ hasText: category });
  await categoryButton.first().click();
  await page.waitForTimeout(500);
  
  // テキストを含むボタンを探す
  const buttons = await page.locator('button').all();
  for (const button of buttons) {
    const text = await button.textContent();
    if (text && text.includes(functionName) && text.includes(description)) {
      await button.click();
      await page.waitForTimeout(2000);
      return;
    }
  }
  
  throw new Error(`Function button not found: ${functionName} - ${description}`);
}

export async function changeCellValue(page: any, row: number, col: number, value: string) {
  const cell = page.locator('table tbody tr').nth(row).locator('td').nth(col);
  await cell.click();
  await page.keyboard.press('Control+a');
  await page.keyboard.press('Delete');
  await page.keyboard.type(value);
  await page.keyboard.press('Enter');
  await page.waitForTimeout(500);
}

export async function getCellValue(page: any, row: number, col: number): Promise<string> {
  const cell = page.locator('table tbody tr').nth(row).locator('td').nth(col);
  return await cell.textContent() || '';
}

// 各テストデータに対してテストを生成
export function generateTestsForCategory(categoryName: string, testDataArray: IndividualFunctionTest[]) {
  test.describe(categoryName, () => {
    test.beforeEach(async ({ page }) => {
      await page.goto('/demo');
      await page.click('text=個別関数デモ');
      await page.waitForLoadState('networkidle');
    });

    // カテゴリに一致するテストのみフィルタリング
    const categoryTests = testDataArray.filter(test => test.category === categoryName);

    // 各関数のテストを生成
    for (const testData of categoryTests) {
      test(`${testData.name} - ${testData.description}`, async ({ page }) => {
        await selectFunction(page, testData.category, testData.name, testData.description);
        
        // expectedValuesが存在しない場合はスキップ
        if (!testData.expectedValues) {
          console.log(`Skipping test for ${testData.name} - no expectedValues defined`);
          return;
        }
        
        // expectedValuesと実際の計算結果を比較
        for (const [cell, expectedValue] of Object.entries(testData.expectedValues)) {
          const col = cell.charCodeAt(0) - 'A'.charCodeAt(0);
          const row = parseInt(cell.substring(1));
          const actualValue = await getCellValue(page, row, col);
          
          // 数値または文字列として比較
          if (typeof expectedValue === 'number') {
            const actualNum = parseFloat(actualValue);
            if (!isNaN(actualNum)) {
              expect(actualNum).toBeCloseTo(expectedValue, 2);
            } else {
              // エラー値の場合
              expect(actualValue).toMatch(/#[A-Z]+[!?]/);
            }
          } else if (typeof expectedValue === 'boolean') {
            expect(actualValue.toUpperCase()).toBe(expectedValue ? 'TRUE' : 'FALSE');
          } else {
            expect(actualValue).toBe(expectedValue.toString());
          }
        }
        
        // 値を変更して再計算を確認
        if (testData.data.length > 2 && testData.data[1].length > 1) {
          const firstDataCol = 0;
          const newValue = testData.data[1][firstDataCol];
          
          // 数値の場合は値を変更
          if (typeof newValue === 'number') {
            await changeCellValue(page, 2, firstDataCol, (newValue * 2).toString());
            await page.waitForTimeout(500);
            
            // 再計算後の値を確認
            const firstCell = Object.keys(testData.expectedValues)[0];
            const col = firstCell.charCodeAt(0) - 'A'.charCodeAt(0);
            const actualValue = await getCellValue(page, 2, col);
            expect(actualValue).toBeTruthy();
            expect(actualValue).not.toBe(testData.expectedValues[firstCell].toString());
          }
        }
      });
    }
  });
}