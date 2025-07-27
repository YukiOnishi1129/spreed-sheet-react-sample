import { test, expect } from '@playwright/test';

test.describe('個別関数デモ - 包括的テスト', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/demo');
    await page.click('text=個別関数デモ');
    await page.waitForLoadState('networkidle');
  });

  test.describe('基本演算子', () => {
    test('加算演算子（+）の動作確認と値変更による再計算', async ({ page }) => {
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
      
      // スプレッドシートが更新されるまで待機
      await page.waitForTimeout(2000);
      
      // 初期値の確認
      const c2Cell = page.locator('table tbody tr').nth(2).locator('td').nth(2);
      await expect(c2Cell).toHaveText('30');
      
      // A2セルの値を変更（10 → 15）
      const a2Cell = page.locator('table tbody tr').nth(2).locator('td').nth(0);
      await a2Cell.click();
      await page.keyboard.type('15');
      await page.keyboard.press('Enter');
      
      // 再計算を待つ
      await page.waitForTimeout(1000);
      
      // 計算結果が更新されたことを確認（30 → 35）
      await expect(c2Cell).toHaveText('35');
    });

    test('全ての基本演算子のテスト', async ({ page }) => {
      const operators = [
        { name: 'SUBTRACT_OPERATOR', description: '減算演算子（-）', cell: 'C2', initial: '20', newValue: '25' },
        { name: 'MULTIPLY_OPERATOR', description: '乗算演算子（*）', cell: 'C2', initial: '30', newValue: '50' },
        { name: 'DIVIDE_OPERATOR', description: '除算演算子（/）', cell: 'C2', initial: '5', newValue: '2' },
      ];

      for (const op of operators) {
        // モーダルを開く
        await page.click('text=関数を選択');
        await page.waitForSelector('.fixed.inset-0.bg-black', { state: 'visible' });
        
        // 基本演算子カテゴリを選択
        await page.click('button:has-text("00. 基本演算子")');
        
        // 演算子を選択
        const opButton = page.locator('button').filter({ 
          hasText: op.name 
        }).filter({ 
          hasText: op.description 
        });
        await opButton.click();
        
        // 初期値を確認
        await page.waitForTimeout(2000);
        const resultCell = page.locator('table tbody tr').nth(2).locator('td').nth(2);
        await expect(resultCell).toHaveText(op.initial);
        
        // モーダルを閉じる
        await page.waitForSelector('.fixed.inset-0.bg-black', { state: 'hidden' });
      }
    });
  });

  test.describe('数学関数', () => {
    test('SUM関数の複数セル変更による再計算', async ({ page }) => {
      // モーダルを開く
      await page.click('text=関数を選択');
      await page.waitForSelector('.fixed.inset-0.bg-black', { state: 'visible' });
      
      // 数学カテゴリを選択
      await page.locator('button').filter({ hasText: '01. 数学' }).first().click();
      
      // SUM関数を選択
      const sumButton = page.locator('button').filter({ 
        hasText: 'SUM' 
      }).filter({ 
        hasText: '数値の合計を計算' 
      });
      await sumButton.click();
      
      // 初期値を確認（E2 = 100）
      await page.waitForTimeout(2000);
      const e2Cell = page.locator('table tbody tr').nth(2).locator('td').nth(4);
      await expect(e2Cell).toHaveText('100');
      
      // A2セルの値を変更（10 → 15）
      const a2Cell = page.locator('table tbody tr').nth(2).locator('td').nth(0);
      await a2Cell.click();
      await page.keyboard.type('15');
      await page.keyboard.press('Enter');
      
      // B2セルの値を変更（20 → 25）
      const b2Cell = page.locator('table tbody tr').nth(2).locator('td').nth(1);
      await b2Cell.click();
      await page.keyboard.type('25');
      await page.keyboard.press('Enter');
      
      // 再計算を待つ
      await page.waitForTimeout(1000);
      
      // 計算結果が更新されたことを確認（100 → 110）
      await expect(e2Cell).toHaveText('110');
    });

    test('ROUND関数の桁数変更による動作確認', async ({ page }) => {
      // モーダルを開く
      await page.click('text=関数を選択');
      await page.waitForSelector('.fixed.inset-0.bg-black', { state: 'visible' });
      
      // 数学カテゴリを選択
      await page.locator('button').filter({ hasText: '01. 数学' }).first().click();
      
      // ROUND関数を選択
      const roundButton = page.locator('button').filter({ 
        hasText: 'ROUND' 
      }).filter({ 
        hasText: '指定桁数で四捨五入' 
      });
      await roundButton.click();
      
      // 初期値を確認（C2 = 3.14）
      await page.waitForTimeout(2000);
      const c2Cell = page.locator('table tbody tr').nth(2).locator('td').nth(2);
      await expect(c2Cell).toHaveText('3.14');
      
      // B2セルの桁数を変更（2 → 3）
      const b2Cell = page.locator('table tbody tr').nth(2).locator('td').nth(1);
      await b2Cell.click();
      await page.keyboard.type('3');
      await page.keyboard.press('Enter');
      
      // 再計算を待つ
      await page.waitForTimeout(1000);
      
      // 計算結果が更新されたことを確認（3.14 → 3.142）
      await expect(c2Cell).toHaveText('3.142');
    });
  });

  test.describe('統計関数', () => {
    test('AVERAGE関数の値変更による平均値の再計算', async ({ page }) => {
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
      
      // 初期値を確認（E2 = 25）
      await page.waitForTimeout(2000);
      const e2Cell = page.locator('table tbody tr').nth(2).locator('td').nth(4);
      await expect(e2Cell).toHaveText('25');
      
      // D2セルの値を変更（40 → 60）
      const d2Cell = page.locator('table tbody tr').nth(2).locator('td').nth(3);
      await d2Cell.click();
      await page.keyboard.type('60');
      await page.keyboard.press('Enter');
      
      // 再計算を待つ
      await page.waitForTimeout(1000);
      
      // 計算結果が更新されたことを確認（25 → 30）
      await expect(e2Cell).toHaveText('30');
    });

    test('MAX/MIN関数の動的な値変更', async ({ page }) => {
      // MAX関数のテスト
      await page.click('text=関数を選択');
      await page.waitForSelector('.fixed.inset-0.bg-black', { state: 'visible' });
      await page.click('button:has-text("02. 統計")');
      
      const maxButton = page.locator('button').filter({ 
        hasText: 'MAX' 
      }).filter({ 
        hasText: '最大値を返す' 
      });
      await maxButton.click();
      
      await page.waitForTimeout(2000);
      const e2Cell = page.locator('table tbody tr').nth(2).locator('td').nth(4);
      await expect(e2Cell).toHaveText('40');
      
      // A2セルの値を最大値に変更（10 → 50）
      const a2Cell = page.locator('table tbody tr').nth(2).locator('td').nth(0);
      await a2Cell.click();
      await page.keyboard.type('50');
      await page.keyboard.press('Enter');
      
      await page.waitForTimeout(1000);
      await expect(e2Cell).toHaveText('50');
    });
  });

  test.describe('テキスト関数', () => {
    test('CONCATENATE関数の文字列結合と変更', async ({ page }) => {
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
      
      // A2セルの文字列を変更
      const a2Cell = page.locator('table tbody tr').nth(2).locator('td').nth(0);
      await a2Cell.click();
      await page.keyboard.type('Test');
      await page.keyboard.press('Enter');
      
      // 結果セルが更新されることを確認
      await page.waitForTimeout(1000);
      const c2Cell = page.locator('table tbody tr').nth(2).locator('td').nth(2);
      const cellText = await c2Cell.textContent();
      expect(cellText).toContain('Test');
    });
  });

  test.describe('日付関数', () => {
    test('TODAY関数の動作確認', async ({ page }) => {
      // モーダルを開く
      await page.click('text=関数を選択');
      await page.waitForSelector('.fixed.inset-0.bg-black', { state: 'visible' });
      
      // 日付カテゴリを選択
      await page.click('button:has-text("04. 日付")');
      
      // TODAY関数を選択
      const todayButton = page.locator('button').filter({ 
        hasText: 'TODAY' 
      }).filter({ 
        hasText: '今日の日付' 
      });
      await todayButton.click();
      
      await page.waitForTimeout(2000);
      
      // TODAY関数の結果が日付形式であることを確認
      const a2Cell = page.locator('table tbody tr').nth(2).locator('td').nth(0);
      const cellText = await a2Cell.textContent();
      
      // 日付パターンにマッチすることを確認
      expect(cellText).toMatch(/^\d{1,2}\/\d{1,2}\/\d{4}$/);
    });
  });

  test.describe('論理関数', () => {
    test('IF関数の条件変更による結果の切り替え', async ({ page }) => {
      // モーダルを開く
      await page.click('text=関数を選択');
      await page.waitForSelector('.fixed.inset-0.bg-black', { state: 'visible' });
      
      // 論理カテゴリを選択
      await page.click('button:has-text("05. 論理")');
      
      // IF関数を選択
      const ifButton = page.locator('button').filter({ 
        hasText: 'IF' 
      }).filter({ 
        hasText: '条件分岐' 
      });
      await ifButton.click();
      
      await page.waitForTimeout(2000);
      
      // A2セルの値を変更して条件を切り替える
      const a2Cell = page.locator('table tbody tr').nth(2).locator('td').nth(0);
      await a2Cell.click();
      await page.keyboard.type('0'); // 条件をFALSEにする
      await page.keyboard.press('Enter');
      
      // 結果が変わることを確認
      await page.waitForTimeout(1000);
      const b2Cell = page.locator('table tbody tr').nth(2).locator('td').nth(1);
      await expect(b2Cell).toHaveText('FALSE');
    });
  });

  test.describe('複雑な計算の連鎖', () => {
    test('複数のセルが連動して再計算される', async ({ page }) => {
      // モーダルを開く
      await page.click('text=関数を選択');
      await page.waitForSelector('.fixed.inset-0.bg-black', { state: 'visible' });
      
      // 数学カテゴリを選択
      await page.locator('button').filter({ hasText: '01. 数学' }).first().click();
      
      // POWER関数を選択（べき乗計算）
      const powerButton = page.locator('button').filter({ 
        hasText: 'POWER' 
      }).filter({ 
        hasText: 'べき乗を計算' 
      });
      await powerButton.click();
      
      await page.waitForTimeout(2000);
      
      // 初期値を確認（C2 = 8, C3 = 25）
      const c2Cell = page.locator('table tbody tr').nth(2).locator('td').nth(2);
      const c3Cell = page.locator('table tbody tr').nth(3).locator('td').nth(2);
      await expect(c2Cell).toHaveText('8');
      await expect(c3Cell).toHaveText('25');
      
      // 底の値を変更
      const a2Cell = page.locator('table tbody tr').nth(2).locator('td').nth(0);
      await a2Cell.click();
      await page.keyboard.type('3');
      await page.keyboard.press('Enter');
      
      // 指数の値を変更
      const b2Cell = page.locator('table tbody tr').nth(2).locator('td').nth(1);
      await b2Cell.click();
      await page.keyboard.type('4');
      await page.keyboard.press('Enter');
      
      await page.waitForTimeout(1000);
      
      // 結果が更新されることを確認（3^4 = 81）
      await expect(c2Cell).toHaveText('81');
    });
  });

  test('カテゴリ切り替えのパフォーマンステスト', async ({ page }) => {
    const categories = [
      '00. 基本演算子',
      '01. 数学',
      '02. 統計',
      '03. テキスト',
      '04. 日付',
      '05. 論理'
    ];

    for (const category of categories) {
      await page.click('text=関数を選択');
      await page.waitForSelector('.fixed.inset-0.bg-black', { state: 'visible' });
      
      const startTime = Date.now();
      await page.click(`button:has-text("${category}")`);
      
      // カテゴリの関数が表示されるまで待つ
      await page.waitForSelector('button.p-4.text-left', { state: 'visible' });
      const endTime = Date.now();
      
      // カテゴリ切り替えが1秒以内に完了することを確認
      expect(endTime - startTime).toBeLessThan(1000);
      
      // モーダルを閉じる
      await page.keyboard.press('Escape');
      await page.waitForSelector('.fixed.inset-0.bg-black', { state: 'hidden' });
    }
  });
});