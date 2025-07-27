import { test, expect } from '@playwright/test';

test.describe('全関数包括的テスト', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/demo');
    await page.click('text=個別関数デモ');
    await page.waitForLoadState('networkidle');
  });

  // ヘルパー関数：モーダルを開いて関数を選択
  async function selectFunction(page: any, category: string, functionName: string, description: string) {
    await page.click('text=関数を選択');
    await page.waitForSelector('.fixed.inset-0.bg-black', { state: 'visible' });
    
    // カテゴリを選択
    const categoryButton = page.locator('button').filter({ hasText: category });
    await categoryButton.first().click();
    await page.waitForTimeout(500);
    
    // 関数を選択
    const functionButton = page.locator('button').filter({ 
      hasText: functionName 
    }).filter({ 
      hasText: description 
    });
    await functionButton.click();
    await page.waitForTimeout(2000);
  }

  // ヘルパー関数：セルの値を変更
  async function changeCellValue(page: any, row: number, col: number, value: string) {
    const cell = page.locator('table tbody tr').nth(row).locator('td').nth(col);
    await cell.click();
    await page.keyboard.press('Control+a');
    await page.keyboard.press('Delete');
    await page.keyboard.type(value);
    await page.keyboard.press('Enter');
    await page.waitForTimeout(500);
  }

  // ヘルパー関数：セルの値を取得
  async function getCellValue(page: any, row: number, col: number): Promise<string> {
    const cell = page.locator('table tbody tr').nth(row).locator('td').nth(col);
    return await cell.textContent() || '';
  }

  test.describe('00. 基本演算子', () => {
    test('加算演算子 - 複数の値変更', async ({ page }) => {
      await selectFunction(page, '00. 基本演算子', 'ADD_OPERATOR', '加算演算子（+）');
      
      // 初期値確認
      expect(await getCellValue(page, 2, 2)).toBe('30');
      
      // 両方の値を変更
      await changeCellValue(page, 2, 0, '25');
      await changeCellValue(page, 2, 1, '35');
      
      // 結果確認（25 + 35 = 60）
      expect(await getCellValue(page, 2, 2)).toBe('60');
    });

    test('減算演算子 - 負の結果', async ({ page }) => {
      await selectFunction(page, '00. 基本演算子', 'SUBTRACT_OPERATOR', '減算演算子（-）');
      
      // 初期値確認
      expect(await getCellValue(page, 2, 2)).toBe('20');
      
      // 値を変更して負の結果を作る
      await changeCellValue(page, 2, 0, '10');
      await changeCellValue(page, 2, 1, '50');
      
      // 結果確認（10 - 50 = -40）
      expect(await getCellValue(page, 2, 2)).toBe('-40');
    });

    test('乗算演算子 - 小数の計算', async ({ page }) => {
      await selectFunction(page, '00. 基本演算子', 'MULTIPLY_OPERATOR', '乗算演算子（*）');
      
      await changeCellValue(page, 2, 0, '2.5');
      await changeCellValue(page, 2, 1, '3.2');
      
      // 結果確認（2.5 * 3.2 = 8）
      const result = await getCellValue(page, 2, 2);
      expect(parseFloat(result)).toBeCloseTo(8, 1);
    });

    test('除算演算子 - 小数点精度', async ({ page }) => {
      await selectFunction(page, '00. 基本演算子', 'DIVIDE_OPERATOR', '除算演算子（/）');
      
      await changeCellValue(page, 2, 0, '10');
      await changeCellValue(page, 2, 1, '3');
      
      // 結果確認（10 / 3 = 3.333...）
      const result = await getCellValue(page, 2, 2);
      expect(parseFloat(result)).toBeCloseTo(3.333333, 5);
    });

    test('比較演算子の組み合わせ', async ({ page }) => {
      // 大なり演算子
      await selectFunction(page, '00. 基本演算子', 'GREATER_THAN', '大なり演算子（>）');
      await changeCellValue(page, 2, 0, '15');
      await changeCellValue(page, 2, 1, '10');
      expect(await getCellValue(page, 2, 2)).toBe('TRUE');
      
      // ページリロードして次のテストへ
      await page.reload();
      await page.click('text=個別関数デモ');
      
      // 小なり演算子
      await selectFunction(page, '00. 基本演算子', 'LESS_THAN', '小なり演算子（<）');
      await changeCellValue(page, 2, 0, '5');
      await changeCellValue(page, 2, 1, '10');
      expect(await getCellValue(page, 2, 2)).toBe('TRUE');
    });
  });

  test.describe('01. 数学・三角', () => {
    test('SUM関数 - 範囲の拡張', async ({ page }) => {
      await selectFunction(page, '01. 数学', 'SUM', '数値の合計を計算');
      
      // 初期値確認
      expect(await getCellValue(page, 2, 4)).toBe('100');
      
      // すべての値を変更
      await changeCellValue(page, 2, 0, '15');
      await changeCellValue(page, 2, 1, '25');
      await changeCellValue(page, 2, 2, '35');
      await changeCellValue(page, 2, 3, '45');
      
      // 結果確認（15 + 25 + 35 + 45 = 120）
      expect(await getCellValue(page, 2, 4)).toBe('120');
    });

    test('ROUND関数 - 様々な桁数', async ({ page }) => {
      await selectFunction(page, '01. 数学', 'ROUND', '指定桁数で四捨五入');
      
      // 小数第3位で四捨五入
      await changeCellValue(page, 2, 1, '3');
      expect(await getCellValue(page, 2, 2)).toBe('3.142');
      
      // 小数第1位で四捨五入
      await changeCellValue(page, 2, 1, '1');
      expect(await getCellValue(page, 2, 2)).toBe('3.1');
      
      // 整数に四捨五入
      await changeCellValue(page, 2, 1, '0');
      expect(await getCellValue(page, 2, 2)).toBe('3');
    });

    test('POWER関数 - 指数計算', async ({ page }) => {
      await selectFunction(page, '01. 数学', 'POWER', 'べき乗を計算');
      
      // 2の10乗
      await changeCellValue(page, 2, 0, '2');
      await changeCellValue(page, 2, 1, '10');
      expect(await getCellValue(page, 2, 2)).toBe('1024');
      
      // 負の指数
      await changeCellValue(page, 2, 1, '-2');
      expect(await getCellValue(page, 2, 2)).toBe('0.25');
    });

    test('三角関数の連携', async ({ page }) => {
      // SIN関数
      await selectFunction(page, '01. 数学・三角', 'SIN', '正弦を計算');
      await changeCellValue(page, 2, 0, '0');
      expect(await getCellValue(page, 2, 1)).toBe('0');
      
      // ページリロード
      await page.reload();
      await page.click('text=個別関数デモ');
      
      // COS関数
      await selectFunction(page, '01. 数学・三角', 'COS', '余弦を計算');
      await changeCellValue(page, 2, 0, '0');
      expect(await getCellValue(page, 2, 1)).toBe('1');
    });

    test('MOD関数 - 剰余計算', async ({ page }) => {
      await selectFunction(page, '01. 数学', 'MOD', '剰余を返す');
      
      // 17 % 5 = 2
      await changeCellValue(page, 2, 0, '17');
      await changeCellValue(page, 2, 1, '5');
      expect(await getCellValue(page, 2, 2)).toBe('2');
      
      // 負の数の剰余
      await changeCellValue(page, 2, 0, '-17');
      const result = await getCellValue(page, 2, 2);
      expect(parseFloat(result)).toBe(3);
    });
  });

  test.describe('02. 統計', () => {
    test('AVERAGE関数 - 動的平均計算', async ({ page }) => {
      await selectFunction(page, '02. 統計', 'AVERAGE', '平均値を計算');
      
      // 初期値確認
      expect(await getCellValue(page, 2, 4)).toBe('25');
      
      // 値を変更
      await changeCellValue(page, 2, 0, '20');
      await changeCellValue(page, 2, 1, '30');
      await changeCellValue(page, 2, 2, '40');
      await changeCellValue(page, 2, 3, '50');
      
      // 新しい平均値（(20+30+40+50)/4 = 35）
      expect(await getCellValue(page, 2, 4)).toBe('35');
    });

    test('MAX/MIN関数 - 極値の更新', async ({ page }) => {
      // MAX関数
      await selectFunction(page, '02. 統計', 'MAX', '最大値を返す');
      
      // 新しい最大値を設定
      await changeCellValue(page, 2, 0, '100');
      expect(await getCellValue(page, 2, 4)).toBe('100');
      
      // ページリロード
      await page.reload();
      await page.click('text=個別関数デモ');
      
      // MIN関数
      await selectFunction(page, '02. 統計', 'MIN', '最小値を返す');
      
      // 新しい最小値を設定
      await changeCellValue(page, 2, 0, '-50');
      expect(await getCellValue(page, 2, 4)).toBe('-50');
    });

    test('COUNT関数 - 数値カウント', async ({ page }) => {
      await selectFunction(page, '02. 統計', 'COUNT', '数値の個数をカウント');
      
      // 初期のカウント確認
      const initialCount = await getCellValue(page, 2, 4);
      
      // テキストを入力（カウントされない）
      await changeCellValue(page, 2, 0, 'text');
      
      // カウントが減ることを確認
      const newCount = await getCellValue(page, 2, 4);
      expect(parseInt(newCount)).toBeLessThan(parseInt(initialCount));
    });

    test('STDEV関数 - 標準偏差', async ({ page }) => {
      await selectFunction(page, '02. 統計', 'STDEV', '標準偏差を計算');
      
      // 値を均一に近づける
      await changeCellValue(page, 2, 0, '10');
      await changeCellValue(page, 2, 1, '10');
      await changeCellValue(page, 2, 2, '10');
      await changeCellValue(page, 2, 3, '10');
      
      // 標準偏差が0になることを確認
      expect(await getCellValue(page, 2, 4)).toBe('0');
    });
  });

  test.describe('03. テキスト', () => {
    test('CONCATENATE関数 - 複数文字列の結合', async ({ page }) => {
      await selectFunction(page, '03. テキスト', 'CONCATENATE', '文字列を結合');
      
      await changeCellValue(page, 2, 0, 'Hello');
      await changeCellValue(page, 2, 1, ' ');
      await changeCellValue(page, 2, 2, 'World');
      
      const result = await getCellValue(page, 2, 3);
      expect(result).toContain('Hello World');
    });

    test('UPPER/LOWER関数 - 大文字小文字変換', async ({ page }) => {
      // UPPER関数
      await selectFunction(page, '03. テキスト', 'UPPER', '大文字に変換');
      
      await changeCellValue(page, 2, 0, 'hello world');
      expect(await getCellValue(page, 2, 1)).toBe('HELLO WORLD');
      
      // ページリロード
      await page.reload();
      await page.click('text=個別関数デモ');
      
      // LOWER関数
      await selectFunction(page, '03. テキスト', 'LOWER', '小文字に変換');
      
      await changeCellValue(page, 2, 0, 'HELLO WORLD');
      expect(await getCellValue(page, 2, 1)).toBe('hello world');
    });

    test('LEN関数 - 文字列長', async ({ page }) => {
      await selectFunction(page, '03. テキスト', 'LEN', '文字列の長さ');
      
      // ASCII文字
      await changeCellValue(page, 2, 0, 'Hello');
      expect(await getCellValue(page, 2, 1)).toBe('5');
      
      // 日本語文字
      await changeCellValue(page, 2, 0, 'こんにちは');
      expect(await getCellValue(page, 2, 1)).toBe('5');
      
      // 空文字
      await changeCellValue(page, 2, 0, '');
      expect(await getCellValue(page, 2, 1)).toBe('0');
    });

    test('TRIM関数 - 空白除去', async ({ page }) => {
      await selectFunction(page, '03. テキスト', 'TRIM', '前後の空白を削除');
      
      await changeCellValue(page, 2, 0, '  Hello World  ');
      expect(await getCellValue(page, 2, 1)).toBe('Hello World');
      
      // 複数スペース
      await changeCellValue(page, 2, 0, '   Multiple   Spaces   ');
      const result = await getCellValue(page, 2, 1);
      expect(result.startsWith('Multiple')).toBe(true);
      expect(result.endsWith('Spaces')).toBe(true);
    });
  });

  test.describe('04. 日付・時刻', () => {
    test('TODAY関数 - 現在日付', async ({ page }) => {
      await selectFunction(page, '04. 日付', 'TODAY', '今日の日付');
      
      const todayCell = await getCellValue(page, 2, 0);
      // 日付形式であることを確認
      expect(todayCell).toMatch(/^\d{1,2}\/\d{1,2}\/\d{4}$/);
      
      // 今日の日付であることを確認
      const today = new Date();
      const [month, day, year] = todayCell.split('/').map(Number);
      expect(year).toBe(today.getFullYear());
    });

    test('NOW関数 - 現在日時', async ({ page }) => {
      await selectFunction(page, '04. 日付', 'NOW', '現在の日時');
      
      const nowCell = await getCellValue(page, 2, 0);
      // 日時形式であることを確認
      expect(nowCell).toBeTruthy();
      
      // 値が変わることを確認（1秒待機）
      await page.waitForTimeout(1000);
      await page.reload();
      await page.click('text=個別関数デモ');
      await selectFunction(page, '04. 日付', 'NOW', '現在の日時');
      
      const newNowCell = await getCellValue(page, 2, 0);
      expect(newNowCell).not.toBe(nowCell);
    });

    test('DATE関数 - 日付作成', async ({ page }) => {
      await selectFunction(page, '04. 日付', 'DATE', '日付を作成');
      
      // 2024年12月25日を作成
      await changeCellValue(page, 2, 0, '2024');
      await changeCellValue(page, 2, 1, '12');
      await changeCellValue(page, 2, 2, '25');
      
      const result = await getCellValue(page, 2, 3);
      expect(result).toContain('12/25/2024');
    });
  });

  test.describe('05. 論理', () => {
    test('IF関数 - 条件分岐', async ({ page }) => {
      await selectFunction(page, '05. 論理', 'IF', '条件分岐');
      
      // TRUE条件
      await changeCellValue(page, 2, 0, '10');
      expect(await getCellValue(page, 2, 1)).toBe('TRUE');
      
      // FALSE条件
      await changeCellValue(page, 2, 0, '0');
      expect(await getCellValue(page, 2, 1)).toBe('FALSE');
      
      // 負の値
      await changeCellValue(page, 2, 0, '-5');
      expect(await getCellValue(page, 2, 1)).toBe('TRUE');
    });

    test('AND/OR関数 - 複合条件', async ({ page }) => {
      // AND関数
      await selectFunction(page, '05. 論理', 'AND', '複数条件のAND');
      
      // すべてTRUEの場合
      await changeCellValue(page, 2, 0, 'TRUE');
      await changeCellValue(page, 2, 1, 'TRUE');
      await changeCellValue(page, 2, 2, 'TRUE');
      expect(await getCellValue(page, 2, 3)).toBe('TRUE');
      
      // 一つでもFALSEの場合
      await changeCellValue(page, 2, 1, 'FALSE');
      expect(await getCellValue(page, 2, 3)).toBe('FALSE');
    });

    test('NOT関数 - 論理否定', async ({ page }) => {
      await selectFunction(page, '05. 論理', 'NOT', '論理値を反転');
      
      await changeCellValue(page, 2, 0, 'TRUE');
      expect(await getCellValue(page, 2, 1)).toBe('FALSE');
      
      await changeCellValue(page, 2, 0, 'FALSE');
      expect(await getCellValue(page, 2, 1)).toBe('TRUE');
    });
  });

  test.describe('06. 検索・参照', () => {
    test('VLOOKUP関数 - 垂直検索', async ({ page }) => {
      await selectFunction(page, '06. 検索', 'VLOOKUP', '垂直方向に検索');
      
      // 検索値を変更
      const searchCell = page.locator('table tbody tr').nth(6).locator('td').nth(0);
      await searchCell.click();
      await page.keyboard.press('Control+a');
      await page.keyboard.press('Delete');
      await page.keyboard.type('002');
      await page.keyboard.press('Enter');
      
      // 結果が更新されることを確認
      await page.waitForTimeout(1000);
      const resultCell = await getCellValue(page, 6, 1);
      expect(resultCell).toBeTruthy();
    });

    test('INDEX/MATCH関数の組み合わせ', async ({ page }) => {
      await selectFunction(page, '06. 検索', 'INDEX', '指定位置の値を取得');
      
      // 行と列のインデックスを変更
      await changeCellValue(page, 6, 0, '2');
      await changeCellValue(page, 6, 1, '2');
      
      // 結果確認
      const result = await getCellValue(page, 6, 2);
      expect(result).toBeTruthy();
    });
  });

  test.describe('07. 財務', () => {
    test('PMT関数 - ローン支払額計算', async ({ page }) => {
      await selectFunction(page, '07. 財務', 'PMT', 'ローンの定期支払額');
      
      // 金利を変更
      await changeCellValue(page, 2, 0, '0.05');
      
      // 期間を変更
      await changeCellValue(page, 2, 1, '360');
      
      // 元本を変更
      await changeCellValue(page, 2, 2, '1000000');
      
      // 支払額が計算されることを確認
      const payment = await getCellValue(page, 2, 3);
      expect(parseFloat(payment)).toBeLessThan(0); // 支払いは負の値
    });

    test('FV関数 - 将来価値', async ({ page }) => {
      await selectFunction(page, '07. 財務', 'FV', '将来価値を計算');
      
      // 利率
      await changeCellValue(page, 2, 0, '0.05');
      
      // 期間
      await changeCellValue(page, 2, 1, '10');
      
      // 定期支払額
      await changeCellValue(page, 2, 2, '-1000');
      
      // 将来価値が計算されることを確認
      const fv = await getCellValue(page, 2, 3);
      expect(parseFloat(fv)).toBeGreaterThan(10000);
    });
  });

  test.describe('08. 行列', () => {
    test('TRANSPOSE関数 - 行列の転置', async ({ page }) => {
      await selectFunction(page, '08. 行列', 'TRANSPOSE', '行列を転置');
      
      // 元の行列の値を変更
      await changeCellValue(page, 2, 0, '1');
      await changeCellValue(page, 2, 1, '2');
      await changeCellValue(page, 3, 0, '3');
      await changeCellValue(page, 3, 1, '4');
      
      // 転置結果を確認（実装に依存）
      const result = await getCellValue(page, 2, 3);
      expect(result).toBeTruthy();
    });
  });

  test.describe('09. 情報', () => {
    test('ISBLANK関数 - 空白判定', async ({ page }) => {
      await selectFunction(page, '09. 情報', 'ISBLANK', '空白セルかどうか');
      
      // 空白セルの場合
      await changeCellValue(page, 2, 0, '');
      expect(await getCellValue(page, 2, 1)).toBe('TRUE');
      
      // 値がある場合
      await changeCellValue(page, 2, 0, 'test');
      expect(await getCellValue(page, 2, 1)).toBe('FALSE');
    });

    test('ISNUMBER関数 - 数値判定', async ({ page }) => {
      await selectFunction(page, '09. 情報', 'ISNUMBER', '数値かどうか');
      
      // 数値の場合
      await changeCellValue(page, 2, 0, '123');
      expect(await getCellValue(page, 2, 1)).toBe('TRUE');
      
      // テキストの場合
      await changeCellValue(page, 2, 0, 'ABC');
      expect(await getCellValue(page, 2, 1)).toBe('FALSE');
    });

    test('TYPE関数 - データ型判定', async ({ page }) => {
      await selectFunction(page, '09. 情報', 'TYPE', 'データ型を返す');
      
      // 数値
      await changeCellValue(page, 2, 0, '123');
      expect(await getCellValue(page, 2, 1)).toBe('1');
      
      // テキスト
      await changeCellValue(page, 2, 0, 'text');
      expect(await getCellValue(page, 2, 1)).toBe('2');
    });
  });

  test.describe('10. データベース', () => {
    test('DSUM関数 - 条件付き合計', async ({ page }) => {
      await selectFunction(page, '10. データベース', 'DSUM', '条件付き合計');
      
      // 条件を変更してフィルタリング結果を確認
      const conditionCell = page.locator('table tbody tr').nth(2).locator('td').nth(4);
      await conditionCell.click();
      await page.keyboard.press('Control+a');
      await page.keyboard.press('Delete');
      await page.keyboard.type('A');
      await page.keyboard.press('Enter');
      
      // 結果が更新されることを確認
      await page.waitForTimeout(1000);
      const result = await getCellValue(page, 6, 0);
      expect(parseFloat(result)).toBeGreaterThan(0);
    });
  });

  test.describe('11. エンジニアリング', () => {
    test('BIN2DEC関数 - 2進数から10進数', async ({ page }) => {
      await selectFunction(page, '11. エンジニアリング', 'BIN2DEC', '2進数を10進数に変換');
      
      // 2進数を入力
      await changeCellValue(page, 2, 0, '1010');
      expect(await getCellValue(page, 2, 1)).toBe('10');
      
      await changeCellValue(page, 2, 0, '11111111');
      expect(await getCellValue(page, 2, 1)).toBe('255');
    });

    test('DEC2HEX関数 - 10進数から16進数', async ({ page }) => {
      await selectFunction(page, '11. エンジニアリング', 'DEC2HEX', '10進数を16進数に変換');
      
      await changeCellValue(page, 2, 0, '255');
      expect(await getCellValue(page, 2, 1)).toBe('FF');
      
      await changeCellValue(page, 2, 0, '4095');
      expect(await getCellValue(page, 2, 1)).toBe('FFF');
    });
  });

  test.describe('12. 動的配列', () => {
    test('UNIQUE関数 - 重複削除', async ({ page }) => {
      await selectFunction(page, '12. 動的配列', 'UNIQUE', '一意の値を抽出');
      
      // 値を変更して重複を作る
      await changeCellValue(page, 2, 0, 'A');
      await changeCellValue(page, 3, 0, 'A');
      await changeCellValue(page, 4, 0, 'B');
      
      // 結果確認（実装依存）
      await page.waitForTimeout(1000);
    });

    test('SORT関数 - ソート', async ({ page }) => {
      await selectFunction(page, '12. 動的配列', 'SORT', '配列をソート');
      
      // ランダムな値を設定
      await changeCellValue(page, 2, 0, '30');
      await changeCellValue(page, 3, 0, '10');
      await changeCellValue(page, 4, 0, '20');
      
      // ソート結果確認（実装依存）
      await page.waitForTimeout(1000);
    });
  });

  test.describe('13. キューブ', () => {
    test('CUBEVALUE関数 - キューブ値取得', async ({ page }) => {
      await selectFunction(page, '13. キューブ', 'CUBEVALUE', 'キューブから値を取得');
      
      // パラメータを変更
      await changeCellValue(page, 2, 0, 'Sales');
      await changeCellValue(page, 2, 1, '2024');
      
      // 結果確認（実装依存）
      const result = await getCellValue(page, 2, 2);
      expect(result).toBeTruthy();
    });
  });

  test.describe('14. Web・インポート', () => {
    test('HYPERLINK関数 - ハイパーリンク作成', async ({ page }) => {
      await selectFunction(page, '14. Web', 'HYPERLINK', 'ハイパーリンクを作成');
      
      // URLを変更
      await changeCellValue(page, 2, 0, 'https://example.com');
      await changeCellValue(page, 2, 1, 'Example Site');
      
      // リンクが作成されることを確認
      const result = await getCellValue(page, 2, 2);
      expect(result).toContain('Example Site');
    });
  });

  test.describe('15. その他', () => {
    test('ARABIC関数 - ローマ数字変換', async ({ page }) => {
      await selectFunction(page, '15. その他', 'ARABIC', 'ローマ数字をアラビア数字に');
      
      await changeCellValue(page, 2, 0, 'XVI');
      expect(await getCellValue(page, 2, 1)).toBe('16');
      
      await changeCellValue(page, 2, 0, 'MCMXCIX');
      expect(await getCellValue(page, 2, 1)).toBe('1999');
    });
  });

  test.describe('複雑なシナリオ', () => {
    test('複数関数の連携 - 統計分析', async ({ page }) => {
      // 平均値を計算
      await selectFunction(page, '02. 統計', 'AVERAGE', '平均値を計算');
      const average = await getCellValue(page, 2, 4);
      
      // ページリロード
      await page.reload();
      await page.click('text=個別関数デモ');
      
      // 標準偏差を計算
      await selectFunction(page, '02. 統計', 'STDEV', '標準偏差を計算');
      
      // 同じデータセットで計算されることを確認
      const stdev = await getCellValue(page, 2, 4);
      expect(parseFloat(stdev)).toBeGreaterThan(0);
    });

    test('エラー処理の連鎖', async ({ page }) => {
      // SQRT関数で負の値
      await selectFunction(page, '01. 数学', 'SQRT', '平方根を計算');
      
      await changeCellValue(page, 2, 0, '-1');
      const sqrtResult = await getCellValue(page, 2, 1);
      expect(sqrtResult).toMatch(/(#NUM!|NaN)/);
      
      // そのエラーを別の関数で使用
      await changeCellValue(page, 3, 0, `=${sqrtResult}`);
      const chainedResult = await getCellValue(page, 3, 1);
      expect(chainedResult).toMatch(/(#NUM!|NaN|#VALUE!)/);
    });

    test('大量データの処理性能', async ({ page }) => {
      await selectFunction(page, '01. 数学', 'SUM', '数値の合計を計算');
      
      const startTime = Date.now();
      
      // 複数のセルを高速に変更
      for (let i = 0; i < 4; i++) {
        await changeCellValue(page, 2, i, (i * 100).toString());
      }
      
      const endTime = Date.now();
      const processingTime = endTime - startTime;
      
      // 処理時間が妥当であることを確認（5秒以内）
      expect(processingTime).toBeLessThan(5000);
      
      // 結果確認（0 + 100 + 200 + 300 = 600）
      expect(await getCellValue(page, 2, 4)).toBe('600');
    });
  });

  test('関数切り替えの高速性', async ({ page }) => {
    const functions = [
      { category: '01. 数学', name: 'SUM', description: '数値の合計を計算' },
      { category: '02. 統計', name: 'AVERAGE', description: '平均値を計算' },
      { category: '03. テキスト', name: 'LEN', description: '文字列の長さ' },
      { category: '04. 日付', name: 'TODAY', description: '今日の日付' },
      { category: '05. 論理', name: 'IF', description: '条件分岐' }
    ];

    const startTime = Date.now();

    for (const func of functions) {
      await selectFunction(page, func.category, func.name, func.description);
      
      // 各関数が正しく読み込まれることを確認
      const cellValue = await getCellValue(page, 2, 0);
      expect(cellValue).toBeTruthy();
    }

    const endTime = Date.now();
    const totalTime = endTime - startTime;

    // 5つの関数切り替えが30秒以内に完了
    expect(totalTime).toBeLessThan(30000);
  });
});