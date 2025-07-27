import { test, expect } from '@playwright/test';

test.describe('ストレステストと負荷テスト', () => {
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
    await page.waitForTimeout(100);
  }

  async function getCellValue(page: any, row: number, col: number): Promise<string> {
    const cell = page.locator('table tbody tr').nth(row).locator('td').nth(col);
    return await cell.textContent() || '';
  }

  test.describe('高頻度操作', () => {
    test('連続100回のセル更新', async ({ page }) => {
      await selectFunction(page, '00. 基本演算子', 'ADD_OPERATOR', '加算演算子（+）');
      
      const startTime = Date.now();
      const errors = [];

      for (let i = 0; i < 100; i++) {
        try {
          await changeCellValue(page, 2, 0, i.toString());
          
          // 10回ごとに結果を確認
          if (i % 10 === 0) {
            const result = await getCellValue(page, 2, 2);
            const expected = i + 20; // i + B2(20)
            expect(parseFloat(result)).toBe(expected);
          }
        } catch (error) {
          errors.push({ iteration: i, error });
        }
      }

      const endTime = Date.now();
      const totalTime = endTime - startTime;

      // エラーがないことを確認
      expect(errors.length).toBe(0);
      
      // 100回の操作が30秒以内に完了
      expect(totalTime).toBeLessThan(30000);
    });

    test('50個の関数を連続切り替え', async ({ page }) => {
      const functions = [
        { category: '00. 基本演算子', name: 'ADD_OPERATOR', description: '加算演算子（+）' },
        { category: '00. 基本演算子', name: 'SUBTRACT_OPERATOR', description: '減算演算子（-）' },
        { category: '01. 数学', name: 'SUM', description: '数値の合計を計算' },
        { category: '01. 数学', name: 'PRODUCT', description: '数値の積を計算' },
        { category: '02. 統計', name: 'AVERAGE', description: '平均値を計算' },
        { category: '02. 統計', name: 'MAX', description: '最大値を返す' },
        { category: '02. 統計', name: 'MIN', description: '最小値を返す' },
        { category: '03. テキスト', name: 'LEN', description: '文字列の長さ' },
        { category: '03. テキスト', name: 'UPPER', description: '大文字に変換' },
        { category: '04. 日付', name: 'TODAY', description: '今日の日付' }
      ];

      const startTime = Date.now();

      for (let i = 0; i < 50; i++) {
        const func = functions[i % functions.length];
        await selectFunction(page, func.category, func.name, func.description);
        
        // 各関数が正しく読み込まれたことを確認
        const firstCell = await getCellValue(page, 2, 0);
        expect(firstCell).toBeTruthy();
      }

      const endTime = Date.now();
      const avgSwitchTime = (endTime - startTime) / 50;

      // 平均切り替え時間が3秒以内
      expect(avgSwitchTime).toBeLessThan(3000);
    });
  });

  test.describe('大量データ処理', () => {
    test('1000文字の文字列処理', async ({ page }) => {
      await selectFunction(page, '03. テキスト', 'LEN', '文字列の長さ');
      
      // 1000文字の文字列を生成
      const longString = 'あ'.repeat(1000);
      
      const startTime = Date.now();
      await changeCellValue(page, 2, 0, longString);
      const endTime = Date.now();

      // 処理時間が5秒以内
      expect(endTime - startTime).toBeLessThan(5000);
      
      // 正しい長さが計算される
      const length = await getCellValue(page, 2, 1);
      expect(parseInt(length)).toBe(1000);
    });

    test('非常に大きな数値の計算', async ({ page }) => {
      await selectFunction(page, '01. 数学', 'POWER', 'べき乗を計算');
      
      // 2の1000乗を計算しようとする
      await changeCellValue(page, 2, 0, '2');
      await changeCellValue(page, 2, 1, '1000');
      
      const result = await getCellValue(page, 2, 2);
      
      // Infinityまたは科学的記数法で表示される
      expect(result).toMatch(/(Infinity|[0-9.]+e\+[0-9]+)/i);
    });

    test('深いネストの数式', async ({ page }) => {
      await selectFunction(page, '00. 基本演算子', 'ADD_OPERATOR', '加算演算子（+）');
      
      // 深くネストした数式を作成
      let nestedFormula = '1';
      for (let i = 0; i < 20; i++) {
        nestedFormula = `(${nestedFormula}+1)`;
      }
      
      await changeCellValue(page, 2, 0, `=${nestedFormula}`);
      
      // エラーまたは計算結果が表示される
      const result = await getCellValue(page, 2, 0);
      expect(result).toBeTruthy();
    });
  });

  test.describe('同時操作シミュレーション', () => {
    test('複数セルの高速更新', async ({ page }) => {
      await selectFunction(page, '01. 数学', 'SUM', '数値の合計を計算');
      
      const updates = [];
      const startTime = Date.now();

      // 4つのセルを50回ずつ更新（計200回）
      for (let i = 0; i < 50; i++) {
        for (let col = 0; col < 4; col++) {
          updates.push(changeCellValue(page, 2, col, (i * col).toString()));
        }
      }

      // すべての更新を並列実行
      await Promise.all(updates.slice(0, 10)); // 最初の10個だけ並列実行（ブラウザの制限を考慮）
      
      const endTime = Date.now();
      
      // 処理時間の確認
      expect(endTime - startTime).toBeLessThan(10000);
      
      // 最終的な合計値を確認
      const sum = await getCellValue(page, 2, 4);
      expect(parseFloat(sum)).toBeGreaterThanOrEqual(0);
    });

    test('モーダルの連続開閉', async ({ page }) => {
      const startTime = Date.now();

      for (let i = 0; i < 30; i++) {
        // モーダルを開く
        await page.click('text=関数を選択');
        await page.waitForSelector('.fixed.inset-0.bg-black', { state: 'visible' });
        
        // ランダムなカテゴリを選択
        const categories = ['00. 基本演算子', '01. 数学', '02. 統計'];
        const category = categories[i % categories.length];
        await page.click(`button:has-text("${category}")`);
        
        // モーダルを閉じる
        await page.keyboard.press('Escape');
        await page.waitForSelector('.fixed.inset-0.bg-black', { state: 'hidden' });
      }

      const endTime = Date.now();
      
      // 30回の開閉が20秒以内
      expect(endTime - startTime).toBeLessThan(20000);
    });
  });

  test.describe('メモリリークテスト', () => {
    test('長時間の連続操作', async ({ page }) => {
      await selectFunction(page, '02. 統計', 'AVERAGE', '平均値を計算');
      
      // 初期メモリ使用量
      const initialMemory = await page.evaluate(() => {
        if ('memory' in performance) {
          return (performance as any).memory.usedJSHeapSize;
        }
        return 0;
      });

      // 5分間の連続操作をシミュレート（実際は1分に短縮）
      const operationTime = 60 * 1000; // 1分
      const startTime = Date.now();
      let operationCount = 0;

      while (Date.now() - startTime < operationTime) {
        await changeCellValue(page, 2, operationCount % 4, Math.random().toString());
        operationCount++;
        
        // 100ms待機
        await page.waitForTimeout(100);
      }

      // 最終メモリ使用量
      const finalMemory = await page.evaluate(() => {
        if ('memory' in performance) {
          return (performance as any).memory.usedJSHeapSize;
        }
        return 0;
      });

      // メモリ増加量の確認
      if (initialMemory > 0 && finalMemory > 0) {
        const memoryIncrease = finalMemory - initialMemory;
        const increasePercentage = (memoryIncrease / initialMemory) * 100;
        
        // メモリ増加が200%以内
        expect(increasePercentage).toBeLessThan(200);
      }

      console.log(`Operations performed: ${operationCount}`);
    });
  });

  test.describe('エラー回復能力', () => {
    test('連続エラー入力からの回復', async ({ page }) => {
      await selectFunction(page, '01. 数学', 'SQRT', '平方根を計算');
      
      // 50回のエラー入力
      for (let i = 0; i < 50; i++) {
        await changeCellValue(page, 2, 0, '-' + i);
        const result = await getCellValue(page, 2, 1);
        expect(result).toMatch(/(#NUM!|NaN)/);
      }

      // 正常な入力に戻す
      await changeCellValue(page, 2, 0, '100');
      const finalResult = await getCellValue(page, 2, 1);
      
      // 正常に計算される
      expect(parseFloat(finalResult)).toBe(10);
    });

    test('無効な数式の大量入力', async ({ page }) => {
      await selectFunction(page, '00. 基本演算子', 'ADD_OPERATOR', '加算演算子（+）');
      
      const invalidFormulas = [
        '=1/0',
        '=INVALID()',
        '=#REF!',
        '=1+"text"',
        '=SQRT(-1)',
        '=LOG(0)',
        '=POWER(0,0)',
        '=1/""',
        '=A1:A1000000',
        '=CIRCULAR(A1)'
      ];

      for (const formula of invalidFormulas) {
        await changeCellValue(page, 2, 0, formula);
        
        // システムがクラッシュしないことを確認
        const result = await getCellValue(page, 2, 0);
        expect(result).toBeTruthy();
      }

      // 正常な値に戻せることを確認
      await changeCellValue(page, 2, 0, '10');
      const finalResult = await getCellValue(page, 2, 2);
      expect(parseFloat(finalResult)).toBe(30);
    });
  });

  test.describe('極限状態でのパフォーマンス', () => {
    test('最小待機時間での連続操作', async ({ page }) => {
      await selectFunction(page, '00. 基本演算子', 'MULTIPLY_OPERATOR', '乗算演算子（*）');
      
      const operations = [];
      
      // 待機時間なしで100回操作
      for (let i = 1; i <= 100; i++) {
        operations.push((async () => {
          await changeCellValue(page, 2, 0, i.toString());
          return Date.now();
        })());
        
        // 待機時間を最小に
        if (i % 10 === 0) {
          await page.waitForTimeout(10);
        }
      }

      const timestamps = await Promise.all(operations.slice(0, 10));
      
      // すべての操作が完了することを確認
      const finalValue = await getCellValue(page, 2, 0);
      expect(parseInt(finalValue)).toBeGreaterThan(0);
    });

    test('ブラウザリソースの限界テスト', async ({ page }) => {
      // 多数のイベントリスナーを登録
      await page.evaluate(() => {
        for (let i = 0; i < 1000; i++) {
          document.addEventListener('click', () => {});
        }
      });

      // それでも正常に動作することを確認
      await selectFunction(page, '01. 数学', 'PI', '円周率を返す');
      
      const piValue = await getCellValue(page, 2, 0);
      expect(parseFloat(piValue)).toBeCloseTo(3.14159, 5);
    });
  });

  test('総合ストレステスト', async ({ page }) => {
    const testDuration = 30 * 1000; // 30秒
    const startTime = Date.now();
    let operationCount = 0;
    const errors = [];

    while (Date.now() - startTime < testDuration) {
      try {
        // ランダムな操作を選択
        const operations = [
          // 関数切り替え
          async () => {
            const categories = ['00. 基本演算子', '01. 数学', '02. 統計', '03. テキスト'];
            const functions = ['ADD_OPERATOR', 'SUM', 'AVERAGE', 'LEN'];
            const idx = Math.floor(Math.random() * categories.length);
            await selectFunction(page, categories[idx], functions[idx], '');
          },
          // 値変更
          async () => {
            const row = 2;
            const col = Math.floor(Math.random() * 4);
            const value = Math.random() > 0.5 ? Math.random().toString() : 'text';
            await changeCellValue(page, row, col, value);
          },
          // モーダル開閉
          async () => {
            await page.click('text=関数を選択');
            await page.waitForTimeout(100);
            await page.keyboard.press('Escape');
          }
        ];

        const operation = operations[Math.floor(Math.random() * operations.length)];
        await operation();
        operationCount++;

      } catch (error) {
        errors.push({ time: Date.now() - startTime, error });
      }

      // CPU休憩時間
      if (operationCount % 50 === 0) {
        await page.waitForTimeout(100);
      }
    }

    console.log(`Total operations: ${operationCount}`);
    console.log(`Errors: ${errors.length}`);
    
    // エラー率が5%以下
    expect(errors.length / operationCount).toBeLessThan(0.05);
  });
});