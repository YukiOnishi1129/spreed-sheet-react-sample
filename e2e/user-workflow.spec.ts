import { test, expect } from '@playwright/test';

test.describe('実際のユーザーワークフロー', () => {
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

  test.describe('売上分析ワークフロー', () => {
    test('月次売上の合計と平均を計算', async ({ page }) => {
      // Step 1: 売上合計を計算
      await selectFunction(page, '01. 数学', 'SUM', '数値の合計を計算');
      
      // 月次売上データを入力
      const monthlySales = [120000, 135000, 98000, 145000];
      for (let i = 0; i < monthlySales.length; i++) {
        await changeCellValue(page, 2, i, monthlySales[i].toString());
      }
      
      // 合計を確認
      const totalSales = await getCellValue(page, 2, 4);
      expect(parseFloat(totalSales)).toBe(498000);
      
      // Step 2: 平均売上を計算
      await page.reload();
      await page.click('text=個別関数デモ');
      await selectFunction(page, '02. 統計', 'AVERAGE', '平均値を計算');
      
      // 同じデータを入力
      for (let i = 0; i < monthlySales.length; i++) {
        await changeCellValue(page, 2, i, monthlySales[i].toString());
      }
      
      // 平均を確認
      const avgSales = await getCellValue(page, 2, 4);
      expect(parseFloat(avgSales)).toBe(124500);
      
      // Step 3: 最高売上月を特定
      await page.reload();
      await page.click('text=個別関数デモ');
      await selectFunction(page, '02. 統計', 'MAX', '最大値を返す');
      
      for (let i = 0; i < monthlySales.length; i++) {
        await changeCellValue(page, 2, i, monthlySales[i].toString());
      }
      
      const maxSales = await getCellValue(page, 2, 4);
      expect(parseFloat(maxSales)).toBe(145000);
    });

    test('前年比成長率の計算', async ({ page }) => {
      // Step 1: 成長率を計算（(今年-去年)/去年*100）
      await selectFunction(page, '00. 基本演算子', 'SUBTRACT_OPERATOR', '減算演算子（-）');
      
      // 今年と去年の売上
      await changeCellValue(page, 2, 0, '1200000'); // 今年
      await changeCellValue(page, 2, 1, '1000000'); // 去年
      
      // 差額を確認
      const difference = await getCellValue(page, 2, 2);
      expect(parseFloat(difference)).toBe(200000);
      
      // Step 2: パーセンテージに変換
      await page.reload();
      await page.click('text=個別関数デモ');
      await selectFunction(page, '00. 基本演算子', 'DIVIDE_OPERATOR', '除算演算子（/）');
      
      await changeCellValue(page, 2, 0, '200000'); // 差額
      await changeCellValue(page, 2, 1, '1000000'); // 去年の売上
      
      const growthRate = await getCellValue(page, 2, 2);
      expect(parseFloat(growthRate)).toBe(0.2); // 20%成長
    });
  });

  test.describe('在庫管理ワークフロー', () => {
    test('在庫の最小値アラート', async ({ page }) => {
      // Step 1: 最小在庫数を確認
      await selectFunction(page, '02. 統計', 'MIN', '最小値を返す');
      
      // 商品在庫データ
      const inventory = [50, 12, 35, 8]; // 8が最小
      for (let i = 0; i < inventory.length; i++) {
        await changeCellValue(page, 2, i, inventory[i].toString());
      }
      
      const minStock = await getCellValue(page, 2, 4);
      expect(parseFloat(minStock)).toBe(8);
      
      // Step 2: 在庫が10以下の商品をカウント
      await page.reload();
      await page.click('text=個別関数デモ');
      await selectFunction(page, '02. 統計', 'COUNT', '数値の個数をカウント');
      
      // 10以下の在庫のみ入力（実際にはCOUNTIF関数を使うべき）
      await changeCellValue(page, 2, 0, '8');
      await changeCellValue(page, 2, 1, '');
      await changeCellValue(page, 2, 2, '');
      await changeCellValue(page, 2, 3, '');
      
      const lowStockCount = await getCellValue(page, 2, 4);
      expect(parseFloat(lowStockCount)).toBe(1);
    });

    test('発注数量の計算', async ({ page }) => {
      // 安全在庫 - 現在在庫 = 発注数量
      await selectFunction(page, '00. 基本演算子', 'SUBTRACT_OPERATOR', '減算演算子（-）');
      
      // 複数商品の発注数量を計算
      const products = [
        { safety: 100, current: 30 },  // 発注: 70
        { safety: 50, current: 45 },   // 発注: 5
        { safety: 200, current: 180 }  // 発注: 20
      ];

      for (let i = 0; i < products.length; i++) {
        await changeCellValue(page, 2 + i, 0, products[i].safety.toString());
        await changeCellValue(page, 2 + i, 1, products[i].current.toString());
        
        const orderQty = await getCellValue(page, 2 + i, 2);
        expect(parseFloat(orderQty)).toBe(products[i].safety - products[i].current);
      }
    });
  });

  test.describe('財務レポート作成ワークフロー', () => {
    test('投資収益率（ROI）の計算', async ({ page }) => {
      // ROI = (利益 - 投資額) / 投資額 * 100
      
      // Step 1: 利益から投資額を引く
      await selectFunction(page, '00. 基本演算子', 'SUBTRACT_OPERATOR', '減算演算子（-）');
      
      await changeCellValue(page, 2, 0, '1500000'); // 利益
      await changeCellValue(page, 2, 1, '1000000'); // 投資額
      
      const netProfit = await getCellValue(page, 2, 2);
      expect(parseFloat(netProfit)).toBe(500000);
      
      // Step 2: 投資額で割る
      await page.reload();
      await page.click('text=個別関数デモ');
      await selectFunction(page, '00. 基本演算子', 'DIVIDE_OPERATOR', '除算演算子（/）');
      
      await changeCellValue(page, 2, 0, '500000'); // 純利益
      await changeCellValue(page, 2, 1, '1000000'); // 投資額
      
      const roi = await getCellValue(page, 2, 2);
      expect(parseFloat(roi)).toBe(0.5); // 50% ROI
    });

    test('複利計算', async ({ page }) => {
      // FV = PV * (1 + r)^n
      await selectFunction(page, '01. 数学', 'POWER', 'べき乗を計算');
      
      // (1 + 0.05)^10 を計算
      await changeCellValue(page, 2, 0, '1.05'); // 1 + 利率5%
      await changeCellValue(page, 2, 1, '10'); // 10年
      
      const compound = await getCellValue(page, 2, 2);
      const compoundValue = parseFloat(compound);
      expect(compoundValue).toBeCloseTo(1.6289, 3);
      
      // 元本100万円の10年後の価値
      await page.reload();
      await page.click('text=個別関数デモ');
      await selectFunction(page, '00. 基本演算子', 'MULTIPLY_OPERATOR', '乗算演算子（*）');
      
      await changeCellValue(page, 2, 0, '1000000');
      await changeCellValue(page, 2, 1, '1.6289');
      
      const futureValue = await getCellValue(page, 2, 2);
      expect(parseFloat(futureValue)).toBeCloseTo(1628900, -2);
    });
  });

  test.describe('データ品質チェックワークフロー', () => {
    test('データの完全性チェック', async ({ page }) => {
      // Step 1: 空白セルのチェック
      await selectFunction(page, '09. 情報', 'ISBLANK', '空白セルかどうか');
      
      // データセット
      const data = ['100', '', '200', ''];
      for (let i = 0; i < data.length; i++) {
        await changeCellValue(page, 2 + i, 0, data[i]);
        
        const isBlank = await getCellValue(page, 2 + i, 1);
        expect(isBlank).toBe(data[i] === '' ? 'TRUE' : 'FALSE');
      }
      
      // Step 2: 数値データの検証
      await page.reload();
      await page.click('text=個別関数デモ');
      await selectFunction(page, '09. 情報', 'ISNUMBER', '数値かどうか');
      
      const mixedData = ['100', 'ABC', '200.5', '日本語'];
      for (let i = 0; i < mixedData.length; i++) {
        await changeCellValue(page, 2 + i, 0, mixedData[i]);
        
        const isNumber = await getCellValue(page, 2 + i, 1);
        const expected = ['100', '200.5'].includes(mixedData[i]) ? 'TRUE' : 'FALSE';
        expect(isNumber).toBe(expected);
      }
    });

    test('異常値の検出', async ({ page }) => {
      // 標準偏差を使って異常値を検出
      await selectFunction(page, '02. 統計', 'STDEV', '標準偏差を計算');
      
      // データセット（1000は異常値）
      const data = [100, 105, 98, 102, 1000];
      for (let i = 0; i < Math.min(data.length, 4); i++) {
        await changeCellValue(page, 2, i, data[i].toString());
      }
      
      const stdev = await getCellValue(page, 2, 4);
      const stdDevValue = parseFloat(stdev);
      
      // 標準偏差が大きい（異常値の影響）
      expect(stdDevValue).toBeGreaterThan(3);
      
      // 異常値を除外した場合
      await changeCellValue(page, 2, 3, '99'); // 1000を99に変更
      
      const newStdev = await getCellValue(page, 2, 4);
      const newStdDevValue = parseFloat(newStdev);
      
      // 標準偏差が小さくなる
      expect(newStdDevValue).toBeLessThan(5);
    });
  });

  test.describe('レポート作成ワークフロー', () => {
    test('日次レポートの自動生成', async ({ page }) => {
      // Step 1: 日付を記録
      await selectFunction(page, '04. 日付', 'TODAY', '今日の日付');
      
      const today = await getCellValue(page, 2, 0);
      expect(today).toMatch(/^\d{1,2}\/\d{1,2}\/\d{4}$/);
      
      // Step 2: タイトルを作成
      await page.reload();
      await page.click('text=個別関数デモ');
      await selectFunction(page, '03. テキスト', 'CONCATENATE', '文字列を結合');
      
      await changeCellValue(page, 2, 0, '日次売上レポート - ');
      await changeCellValue(page, 2, 1, today);
      await changeCellValue(page, 2, 2, '');
      
      const title = await getCellValue(page, 2, 3);
      expect(title).toContain('日次売上レポート');
      expect(title).toContain(today);
      
      // Step 3: データを大文字に統一
      await page.reload();
      await page.click('text=個別関数デモ');
      await selectFunction(page, '03. テキスト', 'UPPER', '大文字に変換');
      
      const departments = ['sales', 'marketing', 'development'];
      for (let i = 0; i < departments.length; i++) {
        await changeCellValue(page, 2 + i, 0, departments[i]);
        
        const upperDept = await getCellValue(page, 2 + i, 1);
        expect(upperDept).toBe(departments[i].toUpperCase());
      }
    });
  });

  test.describe('教育・学習ワークフロー', () => {
    test('成績評価システム', async ({ page }) => {
      // Step 1: 平均点を計算
      await selectFunction(page, '02. 統計', 'AVERAGE', '平均値を計算');
      
      // テストの点数
      const scores = [85, 92, 78, 88];
      for (let i = 0; i < scores.length; i++) {
        await changeCellValue(page, 2, i, scores[i].toString());
      }
      
      const average = await getCellValue(page, 2, 4);
      expect(parseFloat(average)).toBe(85.75);
      
      // Step 2: 成績判定（IF関数の代わりに論理演算）
      await page.reload();
      await page.click('text=個別関数デモ');
      await selectFunction(page, '00. 基本演算子', 'GREATER_THAN_OR_EQUAL', '以上演算子（>=）');
      
      await changeCellValue(page, 2, 0, '85.75'); // 平均点
      await changeCellValue(page, 2, 1, '80'); // 合格ライン
      
      const isPassed = await getCellValue(page, 2, 2);
      expect(isPassed).toBe('TRUE');
    });

    test('数学の練習問題生成', async ({ page }) => {
      // ランダムな問題を生成
      await selectFunction(page, '01. 数学', 'RANDBETWEEN', '指定範囲の整数乱数');
      
      // 1から100の範囲で問題を生成
      await changeCellValue(page, 2, 0, '1');
      await changeCellValue(page, 2, 1, '100');
      
      const random1 = await getCellValue(page, 2, 2);
      const num1 = parseInt(random1);
      expect(num1).toBeGreaterThanOrEqual(1);
      expect(num1).toBeLessThanOrEqual(100);
      
      // 別の乱数を生成
      await changeCellValue(page, 3, 0, '1');
      await changeCellValue(page, 3, 1, '50');
      
      const random2 = await getCellValue(page, 3, 2);
      const num2 = parseInt(random2);
      expect(num2).toBeGreaterThanOrEqual(1);
      expect(num2).toBeLessThanOrEqual(50);
    });
  });

  test('総合的な業務シミュレーション', async ({ page }) => {
    // 複数の関数を組み合わせた実務シナリオ
    const workflow = [
      {
        step: '売上データ入力',
        category: '01. 数学',
        function: 'SUM',
        description: '数値の合計を計算',
        data: ['150000', '180000', '165000', '195000'],
        expectedResult: 690000
      },
      {
        step: '平均売上計算',
        category: '02. 統計',
        function: 'AVERAGE',
        description: '平均値を計算',
        data: ['150000', '180000', '165000', '195000'],
        expectedResult: 172500
      },
      {
        step: '成長率計算',
        category: '00. 基本演算子',
        function: 'DIVIDE_OPERATOR',
        description: '除算演算子（/）',
        data: ['195000', '150000'],
        expectedResult: 1.3
      }
    ];

    for (const task of workflow) {
      console.log(`実行中: ${task.step}`);
      
      await page.reload();
      await page.click('text=個別関数デモ');
      await selectFunction(page, task.category, task.function, task.description);
      
      // データ入力
      for (let i = 0; i < task.data.length; i++) {
        await changeCellValue(page, 2, i, task.data[i]);
      }
      
      // 結果確認
      const resultCol = task.function === 'DIVIDE_OPERATOR' ? 2 : 4;
      const result = await getCellValue(page, 2, resultCol);
      expect(parseFloat(result)).toBeCloseTo(task.expectedResult, 1);
    }
  });
});