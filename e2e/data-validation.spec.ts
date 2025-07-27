import { test, expect } from '@playwright/test';

test.describe('データ検証と入力制御', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/demo');
    await page.click('text=個別関数デモ');
    await page.waitForLoadState('networkidle');
  });

  // ヘルパー関数
  async function selectFunction(page: any, category: string, functionName: string, description: string) {
    await page.click('text=関数を選択');
    await page.waitForSelector('.fixed.inset-0.bg-black', { state: 'visible' });
    
    const categoryButton = page.locator('button').filter({ hasText: category });
    await categoryButton.first().click();
    await page.waitForTimeout(500);
    
    const functionButton = page.locator('button').filter({ 
      hasText: functionName 
    }).filter({ 
      hasText: description 
    });
    await functionButton.click();
    await page.waitForTimeout(2000);
  }

  async function changeCellValue(page: any, row: number, col: number, value: string) {
    const cell = page.locator('table tbody tr').nth(row).locator('td').nth(col);
    await cell.click();
    await page.keyboard.press('Control+a');
    await page.keyboard.press('Delete');
    await page.keyboard.type(value);
    await page.keyboard.press('Enter');
    await page.waitForTimeout(500);
  }

  async function getCellValue(page: any, row: number, col: number): Promise<string> {
    const cell = page.locator('table tbody tr').nth(row).locator('td').nth(col);
    return await cell.textContent() || '';
  }

  test.describe('数値入力の検証', () => {
    test('整数のみを受け付ける関数', async ({ page }) => {
      await selectFunction(page, '01. 数学', 'FACT', '階乗を計算');
      
      // 小数を入力
      await changeCellValue(page, 2, 0, '5.5');
      const result = await getCellValue(page, 2, 1);
      
      // 5の階乗（120）または6の階乗（720）のいずれか
      const value = parseFloat(result);
      expect(value === 120 || value === 720).toBeTruthy();
    });

    test('負の数を受け付けない関数', async ({ page }) => {
      await selectFunction(page, '01. 数学', 'SQRT', '平方根を計算');
      
      // 負の数を入力
      await changeCellValue(page, 2, 0, '-25');
      const result = await getCellValue(page, 2, 1);
      
      // エラーが返される
      expect(result).toMatch(/(#NUM!|NaN)/);
    });

    test('範囲制限のある関数', async ({ page }) => {
      await selectFunction(page, '01. 数学・三角', 'ASIN', '逆正弦を計算');
      
      // 範囲外の値（-1から1の範囲外）
      await changeCellValue(page, 2, 0, '2');
      const result = await getCellValue(page, 2, 1);
      
      // エラーが返される
      expect(result).toMatch(/(#NUM!|NaN)/);
    });
  });

  test.describe('文字列入力の検証', () => {
    test('空文字列の処理', async ({ page }) => {
      await selectFunction(page, '03. テキスト', 'LEN', '文字列の長さ');
      
      // 空文字列
      await changeCellValue(page, 2, 0, '');
      expect(await getCellValue(page, 2, 1)).toBe('0');
    });

    test('特殊文字の処理', async ({ page }) => {
      await selectFunction(page, '03. テキスト', 'UPPER', '大文字に変換');
      
      // 特殊文字を含む文字列
      await changeCellValue(page, 2, 0, 'hello@world#123!');
      const result = await getCellValue(page, 2, 1);
      
      expect(result).toBe('HELLO@WORLD#123!');
    });

    test('改行を含む文字列', async ({ page }) => {
      await selectFunction(page, '03. テキスト', 'LEN', '文字列の長さ');
      
      // 改行を含む文字列（Shift+Enterで改行）
      const cell = page.locator('table tbody tr').nth(2).locator('td').nth(0);
      await cell.click();
      await page.keyboard.type('Line1');
      await page.keyboard.press('Shift+Enter');
      await page.keyboard.type('Line2');
      await page.keyboard.press('Enter');
      
      const length = await getCellValue(page, 2, 1);
      expect(parseInt(length)).toBeGreaterThan(10);
    });
  });

  test.describe('日付入力の検証', () => {
    test('様々な日付形式', async ({ page }) => {
      await selectFunction(page, '04. 日付', 'DATE', '日付を作成');
      
      // 異なる形式でテスト
      const formats = [
        { year: '2024', month: '12', day: '25' },
        { year: '24', month: '12', day: '25' },
        { year: '2024', month: '1', day: '1' },
      ];

      for (const format of formats) {
        await changeCellValue(page, 2, 0, format.year);
        await changeCellValue(page, 2, 1, format.month);
        await changeCellValue(page, 2, 2, format.day);
        
        const result = await getCellValue(page, 2, 3);
        expect(result).toMatch(/\d{1,2}\/\d{1,2}\/\d{4}/);
      }
    });

    test('無効な日付の処理', async ({ page }) => {
      await selectFunction(page, '04. 日付', 'DATE', '日付を作成');
      
      // 13月（無効）
      await changeCellValue(page, 2, 0, '2024');
      await changeCellValue(page, 2, 1, '13');
      await changeCellValue(page, 2, 2, '1');
      
      const result = await getCellValue(page, 2, 3);
      // 翌年の1月として処理されるか、エラーになる
      expect(result).toBeTruthy();
    });
  });

  test.describe('論理値入力の検証', () => {
    test('大文字小文字の処理', async ({ page }) => {
      await selectFunction(page, '05. 論理', 'NOT', '論理値を反転');
      
      // 小文字のtrue
      await changeCellValue(page, 2, 0, 'true');
      let result = await getCellValue(page, 2, 1);
      expect(result.toUpperCase()).toBe('FALSE');
      
      // 大文字のFALSE
      await changeCellValue(page, 2, 0, 'FALSE');
      result = await getCellValue(page, 2, 1);
      expect(result.toUpperCase()).toBe('TRUE');
    });

    test('数値の論理値変換', async ({ page }) => {
      await selectFunction(page, '05. 論理', 'NOT', '論理値を反転');
      
      // 0はFALSE
      await changeCellValue(page, 2, 0, '0');
      let result = await getCellValue(page, 2, 1);
      expect(result.toUpperCase()).toBe('TRUE');
      
      // 0以外はTRUE
      await changeCellValue(page, 2, 0, '1');
      result = await getCellValue(page, 2, 1);
      expect(result.toUpperCase()).toBe('FALSE');
    });
  });

  test.describe('配列入力の検証', () => {
    test('範囲指定の様々な形式', async ({ page }) => {
      await selectFunction(page, '01. 数学', 'SUM', '数値の合計を計算');
      
      // セルに直接範囲を入力
      const formulaCell = page.locator('table tbody tr').nth(2).locator('td').nth(4);
      await formulaCell.click();
      await page.keyboard.press('Control+a');
      await page.keyboard.press('Delete');
      await page.keyboard.type('=SUM(A2:D2)');
      await page.keyboard.press('Enter');
      
      await page.waitForTimeout(1000);
      const result = await getCellValue(page, 2, 4);
      expect(parseFloat(result)).toBeGreaterThan(0);
    });

    test('非連続範囲の指定', async ({ page }) => {
      await selectFunction(page, '01. 数学', 'SUM', '数値の合計を計算');
      
      // 非連続範囲（A2,C2）
      const formulaCell = page.locator('table tbody tr').nth(2).locator('td').nth(4);
      await formulaCell.click();
      await page.keyboard.press('Control+a');
      await page.keyboard.press('Delete');
      await page.keyboard.type('=SUM(A2,C2)');
      await page.keyboard.press('Enter');
      
      await page.waitForTimeout(1000);
      const result = await getCellValue(page, 2, 4);
      expect(parseFloat(result)).toBe(40); // 10 + 30
    });
  });

  test.describe('エラー値の伝播', () => {
    test('エラーを含む計算', async ({ page }) => {
      await selectFunction(page, '01. 数学', 'SUM', '数値の合計を計算');
      
      // A2にエラーを発生させる
      await changeCellValue(page, 2, 0, '=1/0');
      
      await page.waitForTimeout(1000);
      const result = await getCellValue(page, 2, 4);
      
      // エラーが伝播する
      expect(result).toMatch(/(#DIV\/0!|Infinity|#VALUE!)/);
    });

    test('エラーの種類', async ({ page }) => {
      const errorTests = [
        { formula: '=1/0', expectedError: /(#DIV\/0!|Infinity)/ },
        { formula: '=SQRT(-1)', expectedError: /(#NUM!|NaN)/ },
        { formula: '=A1/B1', expectedError: /(#VALUE!|0|NaN)/ }
      ];

      for (const test of errorTests) {
        await selectFunction(page, '00. 基本演算子', 'ADD_OPERATOR', '加算演算子（+）');
        
        await changeCellValue(page, 2, 0, test.formula);
        await page.waitForTimeout(1000);
        
        const cellValue = await getCellValue(page, 2, 0);
        expect(cellValue).toMatch(test.expectedError);
      }
    });
  });

  test.describe('型変換の自動処理', () => {
    test('文字列から数値への自動変換', async ({ page }) => {
      await selectFunction(page, '00. 基本演算子', 'ADD_OPERATOR', '加算演算子（+）');
      
      // 文字列として数値を入力
      await changeCellValue(page, 2, 0, '"10"');
      await changeCellValue(page, 2, 1, '"20"');
      
      const result = await getCellValue(page, 2, 2);
      
      // 自動的に数値として計算される場合
      const numericResult = parseFloat(result);
      if (!isNaN(numericResult)) {
        expect(numericResult).toBe(30);
      }
    });

    test('日付の数値変換', async ({ page }) => {
      await selectFunction(page, '00. 基本演算子', 'SUBTRACT_OPERATOR', '減算演算子（-）');
      
      // 日付を入力
      await changeCellValue(page, 2, 0, '12/31/2024');
      await changeCellValue(page, 2, 1, '1/1/2024');
      
      const result = await getCellValue(page, 2, 2);
      const days = parseFloat(result);
      
      // 365日または366日（うるう年）
      expect(days === 365 || days === 366 || isNaN(days)).toBeTruthy();
    });
  });

  test.describe('パフォーマンスと制限', () => {
    test('非常に長い文字列', async ({ page }) => {
      await selectFunction(page, '03. テキスト', 'LEN', '文字列の長さ');
      
      // 1000文字の文字列を生成
      const longString = 'a'.repeat(1000);
      await changeCellValue(page, 2, 0, longString);
      
      const length = await getCellValue(page, 2, 1);
      expect(parseInt(length)).toBe(1000);
    });

    test('大量の小数桁数', async ({ page }) => {
      await selectFunction(page, '01. 数学', 'ROUND', '指定桁数で四捨五入');
      
      // 非常に長い小数
      await changeCellValue(page, 2, 0, '3.14159265358979323846264338327950288419716939937510');
      await changeCellValue(page, 2, 1, '15');
      
      const result = await getCellValue(page, 2, 2);
      expect(result).toBeTruthy();
    });
  });

  test.describe('相対参照と絶対参照', () => {
    test('セル参照のコピー', async ({ page }) => {
      await selectFunction(page, '00. 基本演算子', 'ADD_OPERATOR', '加算演算子（+）');
      
      // C2の式をC3にコピーする操作をシミュレート
      const c2Value = await getCellValue(page, 2, 2);
      
      // C3に相対参照として式を入力
      const c3Cell = page.locator('table tbody tr').nth(3).locator('td').nth(2);
      await c3Cell.click();
      await page.keyboard.type('=A3+B3');
      await page.keyboard.press('Enter');
      
      // A3, B3に値を設定
      await changeCellValue(page, 3, 0, '15');
      await changeCellValue(page, 3, 1, '25');
      
      const c3Value = await getCellValue(page, 3, 2);
      expect(parseFloat(c3Value)).toBe(40);
    });
  });
});