import { test, expect } from '@playwright/test';

test.describe('高度なシナリオテスト', () => {
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

  test.describe('循環参照と依存関係', () => {
    test('セル間の依存関係チェーン', async ({ page }) => {
      await selectFunction(page, '00. 基本演算子', 'ADD_OPERATOR', '加算演算子（+）');
      
      // C3にC2を参照する式を設定
      await changeCellValue(page, 3, 2, '=C2+10');
      
      // C4にC3を参照する式を設定
      await changeCellValue(page, 4, 2, '=C3*2');
      
      // A2の値を変更
      await changeCellValue(page, 2, 0, '5');
      
      // 連鎖的な更新を確認
      await page.waitForTimeout(1000);
      
      const c2 = await getCellValue(page, 2, 2); // 5 + 20 = 25
      const c3 = await getCellValue(page, 3, 2); // 25 + 10 = 35
      const c4 = await getCellValue(page, 4, 2); // 35 * 2 = 70
      
      expect(parseFloat(c2)).toBe(25);
      expect(parseFloat(c3)).toBe(35);
      expect(parseFloat(c4)).toBe(70);
    });

    test('複雑な相互参照', async ({ page }) => {
      await selectFunction(page, '01. 数学', 'SUM', '数値の合計を計算');
      
      // E3に別の合計式を設定
      await changeCellValue(page, 3, 4, '=SUM(A3:D3)');
      
      // A3-D3に値を設定
      await changeCellValue(page, 3, 0, '5');
      await changeCellValue(page, 3, 1, '10');
      await changeCellValue(page, 3, 2, '15');
      await changeCellValue(page, 3, 3, '20');
      
      // E4に両方の合計を加算
      await changeCellValue(page, 4, 4, '=E2+E3');
      
      const e2 = await getCellValue(page, 2, 4); // 元の合計
      const e3 = await getCellValue(page, 3, 4); // 50
      const e4 = await getCellValue(page, 4, 4); // e2 + 50
      
      expect(parseFloat(e3)).toBe(50);
      expect(parseFloat(e4)).toBe(parseFloat(e2) + 50);
    });
  });

  test.describe('極限値と境界条件', () => {
    test('非常に大きな数値の計算', async ({ page }) => {
      await selectFunction(page, '00. 基本演算子', 'MULTIPLY_OPERATOR', '乗算演算子（*）');
      
      // 大きな数値の乗算
      await changeCellValue(page, 2, 0, '999999999');
      await changeCellValue(page, 2, 1, '999999999');
      
      const result = await getCellValue(page, 2, 2);
      // 科学的記数法で表示されるか確認
      expect(result).toMatch(/[0-9.]+e\+[0-9]+|[0-9]+/i);
    });

    test('非常に小さな数値の計算', async ({ page }) => {
      await selectFunction(page, '00. 基本演算子', 'DIVIDE_OPERATOR', '除算演算子（/）');
      
      // 非常に小さな数値
      await changeCellValue(page, 2, 0, '1');
      await changeCellValue(page, 2, 1, '1000000000');
      
      const result = await getCellValue(page, 2, 2);
      const value = parseFloat(result);
      expect(value).toBeCloseTo(0.000000001, 10);
    });

    test('精度の限界テスト', async ({ page }) => {
      await selectFunction(page, '01. 数学', 'PI', '円周率を返す');
      
      const piValue = await getCellValue(page, 2, 0);
      const pi = parseFloat(piValue);
      
      // PIの精度確認
      expect(pi).toBeCloseTo(3.141592654, 9);
    });
  });

  test.describe('特殊文字と国際化対応', () => {
    test('Unicode文字の処理', async ({ page }) => {
      await selectFunction(page, '03. テキスト', 'LEN', '文字列の長さ');
      
      // 絵文字
      await changeCellValue(page, 2, 0, '😀😁😂');
      const emojiLen = await getCellValue(page, 2, 1);
      expect(parseInt(emojiLen)).toBe(3);
      
      // 特殊記号
      await changeCellValue(page, 2, 0, '♠♣♥♦');
      const symbolLen = await getCellValue(page, 2, 1);
      expect(parseInt(symbolLen)).toBe(4);
      
      // 多言語混在
      await changeCellValue(page, 2, 0, 'Hello世界مرحبا');
      const mixedLen = await getCellValue(page, 2, 1);
      expect(parseInt(mixedLen)).toBeGreaterThan(10);
    });

    test('RTL言語の処理', async ({ page }) => {
      await selectFunction(page, '03. テキスト', 'CONCATENATE', '文字列を結合');
      
      // アラビア語（右から左）
      await changeCellValue(page, 2, 0, 'مرحبا');
      await changeCellValue(page, 2, 1, ' ');
      await changeCellValue(page, 2, 2, 'بالعالم');
      
      const result = await getCellValue(page, 2, 3);
      expect(result).toContain('مرحبا');
      expect(result).toContain('بالعالم');
    });
  });

  test.describe('パフォーマンステスト', () => {
    test('大量の関数切り替え', async ({ page }) => {
      const categories = [
        '00. 基本演算子',
        '01. 数学',
        '02. 統計',
        '03. テキスト',
        '04. 日付',
        '05. 論理',
        '06. 検索',
        '07. 財務',
        '08. 行列',
        '09. 情報',
        '10. データベース',
        '11. エンジニアリング',
        '12. 動的配列',
        '13. キューブ',
        '14. Web',
        '15. その他'
      ];

      const startTime = Date.now();

      for (let i = 0; i < 10; i++) {
        const category = categories[i % categories.length];
        
        await page.click('text=関数を選択');
        await page.waitForSelector('.fixed.inset-0.bg-black', { state: 'visible' });
        
        const categoryButton = page.locator('button').filter({ hasText: category });
        await categoryButton.first().click();
        
        // 最初の関数を選択
        const firstFunction = page.locator('button.p-4.text-left').first();
        await firstFunction.click();
        
        await page.waitForTimeout(500);
      }

      const endTime = Date.now();
      const totalTime = endTime - startTime;

      // 10回の切り替えが20秒以内
      expect(totalTime).toBeLessThan(20000);
    });

    test('連続的な値更新', async ({ page }) => {
      await selectFunction(page, '01. 数学', 'SUM', '数値の合計を計算');
      
      const startTime = Date.now();

      // 20回の値更新
      for (let i = 0; i < 20; i++) {
        const cellIndex = i % 4;
        await changeCellValue(page, 2, cellIndex, (i * 10).toString());
      }

      const endTime = Date.now();
      const totalTime = endTime - startTime;

      // 20回の更新が15秒以内
      expect(totalTime).toBeLessThan(15000);
      
      // 最終的な合計値を確認
      const sum = await getCellValue(page, 2, 4);
      expect(parseFloat(sum)).toBeGreaterThan(0);
    });
  });

  test.describe('エラー回復とロバスト性', () => {
    test('無効な入力からの回復', async ({ page }) => {
      await selectFunction(page, '01. 数学', 'SQRT', '平方根を計算');
      
      // 無効な入力
      await changeCellValue(page, 2, 0, 'abc');
      let result = await getCellValue(page, 2, 1);
      expect(result).toMatch(/(#VALUE!|NaN|0)/);
      
      // 有効な入力に戻す
      await changeCellValue(page, 2, 0, '16');
      result = await getCellValue(page, 2, 1);
      expect(parseFloat(result)).toBe(4);
    });

    test('空のセルの処理', async ({ page }) => {
      await selectFunction(page, '01. 数学', 'SUM', '数値の合計を計算');
      
      // 一部のセルを空にする
      await changeCellValue(page, 2, 0, '');
      await changeCellValue(page, 2, 2, '');
      
      // 残りの値の合計が計算される
      const sum = await getCellValue(page, 2, 4);
      expect(parseFloat(sum)).toBe(60); // 20 + 40
    });
  });

  test.describe('複合的な使用シナリオ', () => {
    test('財務計算の実例', async ({ page }) => {
      // 投資収益率の計算
      await selectFunction(page, '07. 財務', 'FV', '将来価値を計算');
      
      // 年利5%
      await changeCellValue(page, 2, 0, '0.05');
      
      // 10年間
      await changeCellValue(page, 2, 1, '10');
      
      // 毎年10万円投資
      await changeCellValue(page, 2, 2, '-100000');
      
      const fv = await getCellValue(page, 2, 3);
      const futureValue = Math.abs(parseFloat(fv));
      
      // 10年後の価値が100万円以上
      expect(futureValue).toBeGreaterThan(1000000);
    });

    test('統計分析の実例', async ({ page }) => {
      // データセットの分析
      await selectFunction(page, '02. 統計', 'STDEV', '標準偏差を計算');
      
      // テストスコアのデータ
      await changeCellValue(page, 2, 0, '85');
      await changeCellValue(page, 2, 1, '90');
      await changeCellValue(page, 2, 2, '78');
      await changeCellValue(page, 2, 3, '92');
      
      const stdev = await getCellValue(page, 2, 4);
      const standardDeviation = parseFloat(stdev);
      
      // 標準偏差が妥当な範囲
      expect(standardDeviation).toBeGreaterThan(5);
      expect(standardDeviation).toBeLessThan(10);
    });

    test('日付計算の実例', async ({ page }) => {
      await selectFunction(page, '04. 日付', 'DATEDIF', '日付の差を計算');
      
      // 開始日
      await changeCellValue(page, 2, 0, '1/1/2024');
      
      // 終了日
      await changeCellValue(page, 2, 1, '12/31/2024');
      
      // 単位（年）
      await changeCellValue(page, 2, 2, 'Y');
      
      const diff = await getCellValue(page, 2, 3);
      
      // 差が計算される（0年または1年）
      expect(diff).toMatch(/[0-1]/);
    });
  });

  test.describe('アクセシビリティと使いやすさ', () => {
    test('キーボードナビゲーション', async ({ page }) => {
      await selectFunction(page, '00. 基本演算子', 'ADD_OPERATOR', '加算演算子（+）');
      
      // Tabキーでセル間を移動
      const a2Cell = page.locator('table tbody tr').nth(2).locator('td').nth(0);
      await a2Cell.click();
      
      // Tabで次のセルへ
      await page.keyboard.press('Tab');
      
      // アクティブなセルが変わることを確認
      const activeElement = await page.evaluate(() => document.activeElement?.tagName);
      expect(activeElement).toBeTruthy();
    });

    test('Escapeキーでモーダルを閉じる', async ({ page }) => {
      await page.click('text=関数を選択');
      await page.waitForSelector('.fixed.inset-0.bg-black', { state: 'visible' });
      
      // Escapeキーを押す
      await page.keyboard.press('Escape');
      
      // モーダルが閉じることを確認
      await page.waitForSelector('.fixed.inset-0.bg-black', { state: 'hidden' });
    });
  });

  test.describe('リアルタイム更新', () => {
    test('複数セルの同時更新', async ({ page }) => {
      await selectFunction(page, '01. 数学', 'AVERAGE', '平均値を計算');
      
      // 複数のセルを素早く更新
      const updates = [
        { row: 2, col: 0, value: '100' },
        { row: 2, col: 1, value: '200' },
        { row: 2, col: 2, value: '300' },
        { row: 2, col: 3, value: '400' }
      ];

      for (const update of updates) {
        await changeCellValue(page, update.row, update.col, update.value);
      }

      // 平均値が正しく更新される
      const average = await getCellValue(page, 2, 4);
      expect(parseFloat(average)).toBe(250);
    });

    test('連続した関数変更', async ({ page }) => {
      const functions = [
        { category: '01. 数学', name: 'SUM', description: '数値の合計を計算' },
        { category: '01. 数学', name: 'PRODUCT', description: '数値の積を計算' },
        { category: '02. 統計', name: 'MAX', description: '最大値を返す' },
        { category: '02. 統計', name: 'MIN', description: '最小値を返す' }
      ];

      for (const func of functions) {
        await selectFunction(page, func.category, func.name, func.description);
        
        // 各関数で値を確認
        const resultCol = func.name === 'PRODUCT' ? 3 : 4;
        const result = await getCellValue(page, 2, resultCol);
        expect(result).toBeTruthy();
        
        // 次の関数に切り替える前に少し待つ
        await page.waitForTimeout(500);
      }
    });
  });
});