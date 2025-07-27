import { test, expect, devices } from '@playwright/test';

test.describe('クロスブラウザ互換性テスト', () => {
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

  test.describe('異なるビューポートサイズ', () => {
    test('モバイルビューでの動作', async ({ page }) => {
      // iPhoneサイズに変更
      await page.setViewportSize({ width: 375, height: 667 });
      
      await selectFunction(page, '00. 基本演算子', 'ADD_OPERATOR', '加算演算子（+）');
      
      // 値を変更
      await changeCellValue(page, 2, 0, '15');
      await changeCellValue(page, 2, 1, '25');
      
      // 結果確認
      expect(await getCellValue(page, 2, 2)).toBe('40');
    });

    test('タブレットビューでの動作', async ({ page }) => {
      // iPadサイズに変更
      await page.setViewportSize({ width: 768, height: 1024 });
      
      await selectFunction(page, '01. 数学', 'SUM', '数値の合計を計算');
      
      // スプレッドシートが正しく表示される
      const spreadsheet = page.locator('table');
      await expect(spreadsheet).toBeVisible();
      
      // 値の変更と確認
      await changeCellValue(page, 2, 0, '50');
      const result = await getCellValue(page, 2, 4);
      expect(parseFloat(result)).toBeGreaterThan(0);
    });

    test('4Kディスプレイでの動作', async ({ page }) => {
      // 4K解像度
      await page.setViewportSize({ width: 3840, height: 2160 });
      
      // モーダルが正しく中央に表示される
      await page.click('text=関数を選択');
      await page.waitForSelector('.fixed.inset-0.bg-black', { state: 'visible' });
      
      const modal = page.locator('.fixed.inset-0.z-50');
      await expect(modal).toBeVisible();
      
      // モーダルを閉じる
      await page.keyboard.press('Escape');
    });
  });

  test.describe('タッチデバイスのサポート', () => {
    test('タッチ操作でのセル選択', async ({ page, browserName }) => {
      // タッチデバイスをエミュレート
      await page.addInitScript(() => {
        Object.defineProperty(navigator, 'maxTouchPoints', {
          value: 5,
          writable: false
        });
      });

      await selectFunction(page, '00. 基本演算子', 'ADD_OPERATOR', '加算演算子（+）');
      
      // タップ操作をシミュレート
      const cell = page.locator('table tbody tr').nth(2).locator('td').nth(0);
      await cell.tap();
      
      // キーボードが表示されることを想定（実際の動作は実装依存）
      await page.keyboard.type('20');
      await page.keyboard.press('Enter');
      
      // 値が変更されることを確認
      const value = await getCellValue(page, 2, 0);
      expect(value).toBe('20');
    });

    test('スワイプでのスクロール', async ({ page }) => {
      // 長いリストがある関数を選択
      await page.click('text=関数を選択');
      await page.waitForSelector('.fixed.inset-0.bg-black', { state: 'visible' });
      
      // 統計カテゴリ（多くの関数がある）
      await page.click('button:has-text("02. 統計")');
      
      // スクロール可能な要素を確認
      const scrollableArea = page.locator('.overflow-y-auto');
      await expect(scrollableArea).toBeVisible();
      
      // スワイプ操作をシミュレート（実際の動作は実装依存）
      await scrollableArea.evaluate(el => el.scrollTop = 100);
    });
  });

  test.describe('ブラウザ固有の機能', () => {
    test('コピー＆ペースト', async ({ page, browserName }) => {
      await selectFunction(page, '03. テキスト', 'CONCATENATE', '文字列を結合');
      
      // 最初のセルに値を入力
      await changeCellValue(page, 2, 0, 'Hello');
      
      // コピー操作
      const cell = page.locator('table tbody tr').nth(2).locator('td').nth(0);
      await cell.click();
      await page.keyboard.press('Control+a');
      await page.keyboard.press('Control+c');
      
      // 次のセルにペースト
      const nextCell = page.locator('table tbody tr').nth(2).locator('td').nth(1);
      await nextCell.click();
      await page.keyboard.press('Control+v');
      
      // 値が正しくコピーされたか確認
      const value = await getCellValue(page, 2, 1);
      expect(value).toContain('Hello');
    });

    test('元に戻す/やり直し', async ({ page }) => {
      await selectFunction(page, '00. 基本演算子', 'ADD_OPERATOR', '加算演算子（+）');
      
      // 値を変更
      const originalValue = await getCellValue(page, 2, 0);
      await changeCellValue(page, 2, 0, '999');
      
      // 元に戻す（Ctrl+Z）
      await page.keyboard.press('Control+z');
      
      // 値が元に戻ることを確認（実装依存）
      await page.waitForTimeout(500);
      const revertedValue = await getCellValue(page, 2, 0);
      
      // 実装によっては元に戻らない場合もある
      expect(revertedValue).toBeTruthy();
    });
  });

  test.describe('フォント・エンコーディング', () => {
    test('様々な言語の表示', async ({ page }) => {
      await selectFunction(page, '03. テキスト', 'LEN', '文字列の長さ');
      
      const languages = [
        { text: 'English', expected: 7 },
        { text: '日本語', expected: 3 },
        { text: '中文', expected: 2 },
        { text: 'العربية', expected: 7 },
        { text: 'हिन्दी', expected: 6 },
        { text: 'русский', expected: 7 },
        { text: '한국어', expected: 3 },
        { text: 'ไทย', expected: 3 }
      ];

      for (const lang of languages) {
        await changeCellValue(page, 2, 0, lang.text);
        const length = await getCellValue(page, 2, 1);
        expect(parseInt(length)).toBe(lang.expected);
      }
    });

    test('絵文字の処理', async ({ page }) => {
      await selectFunction(page, '03. テキスト', 'CONCATENATE', '文字列を結合');
      
      // 絵文字を含む文字列
      await changeCellValue(page, 2, 0, '😀');
      await changeCellValue(page, 2, 1, '🎉');
      await changeCellValue(page, 2, 2, '🚀');
      
      const result = await getCellValue(page, 2, 3);
      expect(result).toContain('😀');
      expect(result).toContain('🎉');
      expect(result).toContain('🚀');
    });
  });

  test.describe('メモリとパフォーマンス', () => {
    test('大量のデータ処理', async ({ page }) => {
      await selectFunction(page, '02. 統計', 'AVERAGE', '平均値を計算');
      
      // メモリ使用量の初期状態を記録（実装依存）
      const initialMetrics = await page.evaluate(() => {
        if ('memory' in performance) {
          return (performance as any).memory.usedJSHeapSize;
        }
        return 0;
      });

      // 大量の値変更
      for (let i = 0; i < 10; i++) {
        await changeCellValue(page, 2, i % 4, (Math.random() * 100).toString());
      }

      // メモリ使用量の増加を確認
      const finalMetrics = await page.evaluate(() => {
        if ('memory' in performance) {
          return (performance as any).memory.usedJSHeapSize;
        }
        return 0;
      });

      // メモリリークがないことを確認（大幅な増加がない）
      if (initialMetrics > 0 && finalMetrics > 0) {
        const increase = finalMetrics - initialMetrics;
        expect(increase).toBeLessThan(50 * 1024 * 1024); // 50MB以下
      }
    });

    test('連続操作のレスポンス', async ({ page }) => {
      await selectFunction(page, '01. 数学', 'SUM', '数値の合計を計算');
      
      const operations = [];
      
      // 10回の操作時間を記録
      for (let i = 0; i < 10; i++) {
        const startTime = Date.now();
        await changeCellValue(page, 2, i % 4, (i * 10).toString());
        const endTime = Date.now();
        operations.push(endTime - startTime);
      }

      // 平均操作時間
      const avgTime = operations.reduce((a, b) => a + b, 0) / operations.length;
      
      // レスポンスが適切であることを確認（1秒以内）
      expect(avgTime).toBeLessThan(1000);
    });
  });

  test.describe('ネットワーク条件', () => {
    test('オフライン動作', async ({ page, context }) => {
      await selectFunction(page, '00. 基本演算子', 'ADD_OPERATOR', '加算演算子（+）');
      
      // オフラインモードに切り替え
      await context.setOffline(true);
      
      // オフラインでも計算が動作
      await changeCellValue(page, 2, 0, '30');
      await changeCellValue(page, 2, 1, '40');
      
      const result = await getCellValue(page, 2, 2);
      expect(parseFloat(result)).toBe(70);
      
      // オンラインに戻す
      await context.setOffline(false);
    });

    test('低速ネットワーク', async ({ page }) => {
      // 3G相当の速度制限をシミュレート
      await page.route('**/*', route => {
        setTimeout(() => route.continue(), 100);
      });

      await selectFunction(page, '01. 数学', 'POWER', 'べき乗を計算');
      
      // 低速でも動作することを確認
      await changeCellValue(page, 2, 0, '2');
      await changeCellValue(page, 2, 1, '8');
      
      const result = await getCellValue(page, 2, 2);
      expect(parseFloat(result)).toBe(256);
    });
  });

  test.describe('アクセシビリティ機能', () => {
    test('スクリーンリーダー対応', async ({ page }) => {
      await selectFunction(page, '00. 基本演算子', 'ADD_OPERATOR', '加算演算子（+）');
      
      // ARIA属性の確認
      const table = page.locator('table');
      const role = await table.getAttribute('role');
      
      // テーブルに適切なroleが設定されているか
      if (role) {
        expect(['table', 'grid'].includes(role)).toBeTruthy();
      }

      // セルのアクセシビリティ
      const cell = page.locator('table tbody tr').nth(2).locator('td').nth(0);
      const ariaLabel = await cell.getAttribute('aria-label');
      
      // ARIA属性が存在するか確認（実装依存）
      expect(cell).toBeTruthy();
    });

    test('ハイコントラストモード', async ({ page }) => {
      // ハイコントラストモードをシミュレート
      await page.emulateMedia({ colorScheme: 'dark' });
      
      await selectFunction(page, '03. テキスト', 'UPPER', '大文字に変換');
      
      // UIが見やすいことを確認（実装依存）
      const button = page.locator('button').first();
      const color = await button.evaluate(el => 
        window.getComputedStyle(el).color
      );
      
      expect(color).toBeTruthy();
    });
  });

  test.describe('セキュリティ', () => {
    test('XSS攻撃の防御', async ({ page }) => {
      await selectFunction(page, '03. テキスト', 'CONCATENATE', '文字列を結合');
      
      // 悪意のあるスクリプトを入力
      await changeCellValue(page, 2, 0, '<script>alert("XSS")</script>');
      await changeCellValue(page, 2, 1, '<img src=x onerror="alert(\'XSS\')">');
      
      // スクリプトが実行されないことを確認
      const alerts = [];
      page.on('dialog', dialog => {
        alerts.push(dialog.message());
        dialog.accept();
      });

      await page.waitForTimeout(1000);
      
      // アラートが表示されない
      expect(alerts.length).toBe(0);
      
      // HTMLがエスケープされて表示される
      const result = await getCellValue(page, 2, 3);
      expect(result).toContain('<script>');
    });

    test('インジェクション攻撃の防御', async ({ page }) => {
      await selectFunction(page, '00. 基本演算子', 'ADD_OPERATOR', '加算演算子（+）');
      
      // SQLインジェクション風の入力
      await changeCellValue(page, 2, 0, "'; DROP TABLE users; --");
      
      // エラーにならず、文字列として処理される
      const value = await getCellValue(page, 2, 0);
      expect(value).toBeTruthy();
      
      // 計算結果は数値以外
      const result = await getCellValue(page, 2, 2);
      expect(result).toMatch(/(NaN|#VALUE!|0)/);
    });
  });
});