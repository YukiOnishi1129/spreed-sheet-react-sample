import { test, expect } from '@playwright/test';

test.describe('個別関数デモ - エッジケースと高度な機能', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/demo');
    await page.click('text=個別関数デモ');
    await page.waitForLoadState('networkidle');
  });

  test.describe('エラーハンドリング', () => {
    test('ゼロ除算のエラー確認', async ({ page }) => {
      // モーダルを開く
      await page.click('text=関数を選択');
      await page.waitForSelector('.fixed.inset-0.bg-black', { state: 'visible' });
      
      // 基本演算子カテゴリを選択
      await page.click('button:has-text("00. 基本演算子")');
      
      // DIVIDE_OPERATOR を選択
      const divButton = page.locator('button').filter({ 
        hasText: 'DIVIDE_OPERATOR' 
      }).filter({ 
        hasText: '除算演算子（/）' 
      });
      await divButton.click();
      
      await page.waitForTimeout(2000);
      
      // B2セルの値を0に変更（ゼロ除算）
      const b2Cell = page.locator('table tbody tr').nth(2).locator('td').nth(1);
      await b2Cell.click();
      await page.keyboard.type('0');
      await page.keyboard.press('Enter');
      
      // エラーまたは特殊な値が表示されることを確認
      await page.waitForTimeout(1000);
      const c2Cell = page.locator('table tbody tr').nth(2).locator('td').nth(2);
      const cellText = await c2Cell.textContent();
      
      // #DIV/0! または Infinity が表示されることを確認
      expect(cellText).toMatch(/(#DIV\/0!|Infinity|∞)/);
    });

    test('無効な引数でのエラー確認', async ({ page }) => {
      // モーダルを開く
      await page.click('text=関数を選択');
      await page.waitForSelector('.fixed.inset-0.bg-black', { state: 'visible' });
      
      // 数学カテゴリを選択
      await page.click('button').filter({ hasText: '01. 数学' }).first();
      
      // SQRT関数を選択
      const sqrtButton = page.locator('button').filter({ 
        hasText: 'SQRT' 
      }).filter({ 
        hasText: '平方根を計算' 
      });
      await sqrtButton.click();
      
      await page.waitForTimeout(2000);
      
      // A2セルの値を負の数に変更
      const a2Cell = page.locator('table tbody tr').nth(2).locator('td').nth(0);
      await a2Cell.click();
      await page.keyboard.type('-16');
      await page.keyboard.press('Enter');
      
      // エラーが表示されることを確認
      await page.waitForTimeout(1000);
      const b2Cell = page.locator('table tbody tr').nth(2).locator('td').nth(1);
      const cellText = await b2Cell.textContent();
      
      // #NUM! または NaN が表示されることを確認
      expect(cellText).toMatch(/(#NUM!|NaN)/);
    });
  });

  test.describe('境界値テスト', () => {
    test('非常に大きな数値の計算', async ({ page }) => {
      // モーダルを開く
      await page.click('text=関数を選択');
      await page.waitForSelector('.fixed.inset-0.bg-black', { state: 'visible' });
      
      // 数学カテゴリを選択
      await page.click('button').filter({ hasText: '01. 数学' }).first();
      
      // POWER関数を選択
      const powerButton = page.locator('button').filter({ 
        hasText: 'POWER' 
      }).filter({ 
        hasText: 'べき乗を計算' 
      });
      await powerButton.click();
      
      await page.waitForTimeout(2000);
      
      // 大きな指数を設定
      const b2Cell = page.locator('table tbody tr').nth(2).locator('td').nth(1);
      await b2Cell.click();
      await page.keyboard.type('100');
      await page.keyboard.press('Enter');
      
      // 結果が正しく表示されることを確認
      await page.waitForTimeout(1000);
      const c2Cell = page.locator('table tbody tr').nth(2).locator('td').nth(2);
      const cellText = await c2Cell.textContent();
      
      // 科学的記数法または非常に大きな数が表示されることを確認
      expect(cellText).toBeTruthy();
    });

    test('小数点以下の精度確認', async ({ page }) => {
      // モーダルを開く
      await page.click('text=関数を選択');
      await page.waitForSelector('.fixed.inset-0.bg-black', { state: 'visible' });
      
      // 基本演算子カテゴリを選択
      await page.click('button:has-text("00. 基本演算子")');
      
      // DIVIDE_OPERATOR を選択
      const divButton = page.locator('button').filter({ 
        hasText: 'DIVIDE_OPERATOR' 
      }).filter({ 
        hasText: '除算演算子（/）' 
      });
      await divButton.click();
      
      await page.waitForTimeout(2000);
      
      // 循環小数を生成する値に変更（1/3）
      const a2Cell = page.locator('table tbody tr').nth(2).locator('td').nth(0);
      await a2Cell.click();
      await page.keyboard.type('1');
      await page.keyboard.press('Enter');
      
      const b2Cell = page.locator('table tbody tr').nth(2).locator('td').nth(1);
      await b2Cell.click();
      await page.keyboard.type('3');
      await page.keyboard.press('Enter');
      
      // 結果の精度を確認
      await page.waitForTimeout(1000);
      const c2Cell = page.locator('table tbody tr').nth(2).locator('td').nth(2);
      const cellText = await c2Cell.textContent();
      
      // 0.333... の形式で表示されることを確認
      expect(parseFloat(cellText || '0')).toBeCloseTo(0.333333, 5);
    });
  });

  test.describe('複雑な数式の組み合わせ', () => {
    test('ネストされた関数の動作確認', async ({ page }) => {
      // モーダルを開く
      await page.click('text=関数を選択');
      await page.waitForSelector('.fixed.inset-0.bg-black', { state: 'visible' });
      
      // 数学カテゴリを選択
      await page.click('button').filter({ hasText: '01. 数学' }).first();
      
      // ROUND関数を選択
      const roundButton = page.locator('button').filter({ 
        hasText: 'ROUND' 
      }).filter({ 
        hasText: '指定桁数で四捨五入' 
      });
      await roundButton.click();
      
      await page.waitForTimeout(2000);
      
      // セルに複雑な式を入力
      const a2Cell = page.locator('table tbody tr').nth(2).locator('td').nth(0);
      await a2Cell.click();
      await page.keyboard.type('=SQRT(2)*PI()');
      await page.keyboard.press('Enter');
      
      // 結果が計算されることを確認
      await page.waitForTimeout(1000);
      const c2Cell = page.locator('table tbody tr').nth(2).locator('td').nth(2);
      const cellText = await c2Cell.textContent();
      
      // 値が数値であることを確認
      expect(parseFloat(cellText || '0')).toBeGreaterThan(0);
    });
  });

  test.describe('配列関数のテスト', () => {
    test('範囲指定での計算', async ({ page }) => {
      // モーダルを開く
      await page.click('text=関数を選択');
      await page.waitForSelector('.fixed.inset-0.bg-black', { state: 'visible' });
      
      // 数学カテゴリを選択
      await page.click('button').filter({ hasText: '01. 数学' }).first();
      
      // SUM関数を選択
      const sumButton = page.locator('button').filter({ 
        hasText: 'SUM' 
      }).filter({ 
        hasText: '数値の合計を計算' 
      });
      await sumButton.click();
      
      await page.waitForTimeout(2000);
      
      // 複数のセルの値を一度に変更
      const cells = ['A2', 'B2', 'C2', 'D2'];
      const newValues = [5, 10, 15, 20];
      
      for (let i = 0; i < cells.length; i++) {
        const cell = page.locator('table tbody tr').nth(2).locator('td').nth(i);
        await cell.click();
        await page.keyboard.type(newValues[i].toString());
        await page.keyboard.press('Enter');
      }
      
      // 合計が更新されることを確認
      await page.waitForTimeout(1000);
      const e2Cell = page.locator('table tbody tr').nth(2).locator('td').nth(4);
      await expect(e2Cell).toHaveText('50');
    });
  });

  test.describe('特殊文字と国際化', () => {
    test('日本語を含むテキスト関数', async ({ page }) => {
      // モーダルを開く
      await page.click('text=関数を選択');
      await page.waitForSelector('.fixed.inset-0.bg-black', { state: 'visible' });
      
      // テキストカテゴリを選択
      await page.click('button:has-text("03. テキスト")');
      
      // CONCATENATE関数を選択
      const concatButton = page.locator('button').filter({ 
        hasText: 'CONCATENATE' 
      }).filter({ 
        hasText: '文字列を結合' 
      });
      await concatButton.click();
      
      await page.waitForTimeout(2000);
      
      // 日本語文字列を入力
      const a2Cell = page.locator('table tbody tr').nth(2).locator('td').nth(0);
      await a2Cell.click();
      await page.keyboard.type('こんにちは');
      await page.keyboard.press('Enter');
      
      const b2Cell = page.locator('table tbody tr').nth(2).locator('td').nth(1);
      await b2Cell.click();
      await page.keyboard.type('世界');
      await page.keyboard.press('Enter');
      
      // 結合結果を確認
      await page.waitForTimeout(1000);
      const c2Cell = page.locator('table tbody tr').nth(2).locator('td').nth(2);
      const cellText = await c2Cell.textContent();
      
      expect(cellText).toContain('こんにちは');
      expect(cellText).toContain('世界');
    });
  });

  test.describe('パフォーマンステスト', () => {
    test('大量のセル変更での再計算速度', async ({ page }) => {
      // モーダルを開く
      await page.click('text=関数を選択');
      await page.waitForSelector('.fixed.inset-0.bg-black', { state: 'visible' });
      
      // 統計カテゴリを選択
      await page.click('button:has-text("02. 統計")');
      
      // AVERAGE関数を選択
      const avgButton = page.locator('button').filter({ 
        hasText: 'AVERAGE' 
      }).filter({ 
        hasText: '平均値を計算' 
      });
      await avgButton.click();
      
      await page.waitForTimeout(2000);
      
      const startTime = Date.now();
      
      // 複数のセルを連続で変更
      for (let i = 0; i < 4; i++) {
        const cell = page.locator('table tbody tr').nth(2).locator('td').nth(i);
        await cell.click();
        await page.keyboard.type((i * 10).toString());
        await page.keyboard.press('Enter');
      }
      
      // 最終的な結果を確認
      await page.waitForTimeout(500);
      const e2Cell = page.locator('table tbody tr').nth(2).locator('td').nth(4);
      const cellText = await e2Cell.textContent();
      
      const endTime = Date.now();
      
      // 再計算が2秒以内に完了することを確認
      expect(endTime - startTime).toBeLessThan(2000);
      
      // 平均値が正しいことを確認（(0+10+20+30)/4 = 15）
      expect(parseFloat(cellText || '0')).toBe(15);
    });
  });

  test('セルの参照変更による連鎖的な再計算', async ({ page }) => {
    // モーダルを開く
    await page.click('text=関数を選択');
    await page.waitForSelector('.fixed.inset-0.bg-black', { state: 'visible' });
    
    // 基本演算子カテゴリを選択
    await page.click('button:has-text("00. 基本演算子")');
    
    // ADD_OPERATOR を選択
    const addButton = page.locator('button').filter({ 
      hasText: 'ADD_OPERATOR' 
    }).filter({ 
      hasText: '加算演算子（+）' 
    });
    await addButton.click();
    
    await page.waitForTimeout(2000);
    
    // C3セルに C2の値を参照する式を入力
    const c3Cell = page.locator('table tbody tr').nth(3).locator('td').nth(2);
    await c3Cell.click();
    await page.keyboard.type('=C2*2');
    await page.keyboard.press('Enter');
    
    // A2の値を変更
    const a2Cell = page.locator('table tbody tr').nth(2).locator('td').nth(0);
    await a2Cell.click();
    await page.keyboard.type('15');
    await page.keyboard.press('Enter');
    
    await page.waitForTimeout(1000);
    
    // C2が更新される（35）
    const c2Cell = page.locator('table tbody tr').nth(2).locator('td').nth(2);
    await expect(c2Cell).toHaveText('35');
    
    // C3も連鎖的に更新される（70）
    await expect(c3Cell).toHaveText('70');
  });
});