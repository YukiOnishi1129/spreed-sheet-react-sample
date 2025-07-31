import { describe, it, expect } from 'vitest';
import { DSUM } from '../databaseLogic';
import type { FormulaContext } from '../../shared/types';

describe('DSUM関数のデバッグ', () => {
  it('基本的なDSUM計算', () => {
    const context: FormulaContext = {
      data: [
        ['名前', '部署', '売上'],
        ['田中', '営業', 100],
        ['佐藤', '営業', 150],
        ['鈴木', '開発', 120],
        ['', '', ''],
        ['条件', '', ''],
        ['部署', '', ''],
        ['営業', '', '=DSUM(A1:C4,C1,A6:A7)']
      ]
    };

    // DSUM関数のパターンマッチをテスト
    const formula = 'DSUM(A1:C4,C1,A6:A7)';
    const matches = formula.match(DSUM.pattern);
    
    console.log('Formula:', formula);
    console.log('Pattern:', DSUM.pattern);
    console.log('Matches:', matches);
    
    if (matches) {
      const result = DSUM.calculate(matches, context);
      console.log('Result:', result);
      console.log('Expected:', 250);
      
      // データを確認
      console.log('\\nデータ確認:');
      console.log('営業部の売上データ:');
      console.log('田中(営業): 100');
      console.log('佐藤(営業): 150');
      console.log('合計期待値: 250');
    }
    
    expect(matches).toBeTruthy();
  });

  it('フィールド参照の解決をテスト', () => {
    const context: FormulaContext = {
      data: [
        ['名前', '部署', '売上'],
        ['田中', '営業', 100],
        ['佐藤', '営業', 150],
        ['鈴木', '開発', 120],
        ['', '', ''],
        ['条件', '', ''],
        ['部署', '', ''],
        ['営業', '', '']
      ]
    };

    // フィールドを直接指定
    const formula1 = 'DSUM(A1:C4,"売上",A6:A7)';
    const matches1 = formula1.match(DSUM.pattern);
    
    if (matches1) {
      const result = DSUM.calculate(matches1, context);
      console.log('\\nフィールド名指定の結果:', result);
    }

    // フィールドをインデックスで指定
    const formula2 = 'DSUM(A1:C4,3,A6:A7)';
    const matches2 = formula2.match(DSUM.pattern);
    
    if (matches2) {
      const result = DSUM.calculate(matches2, context);
      console.log('フィールドインデックス指定の結果:', result);
    }
  });
});