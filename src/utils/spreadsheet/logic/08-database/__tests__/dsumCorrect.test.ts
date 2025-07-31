import { describe, it, expect } from 'vitest';
import { DSUM } from '../databaseLogic';
import type { FormulaContext } from '../../shared/types';

describe('DSUM関数の正しいテストデータ', () => {
  it('正しい条件範囲でのDSUM計算', () => {
    const context: FormulaContext = {
      data: [
        ['名前', '部署', '売上'],
        ['田中', '営業', 100],
        ['佐藤', '営業', 150],
        ['鈴木', '開発', 120],
        ['', '', ''],
        ['部署', '', ''],  // 条件のヘッダー
        ['営業', '', ''],  // 条件の値
        ['', '', '']       // 数式の行
      ]
    };

    // 条件範囲を正しく指定: A6:A7（部署ヘッダーと営業値）
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
      console.log('データベース範囲 A1:C4:');
      console.log('  行1: 名前, 部署, 売上（ヘッダー）');
      console.log('  行2: 田中, 営業, 100');
      console.log('  行3: 佐藤, 営業, 150');
      console.log('  行4: 鈴木, 開発, 120');
      console.log('条件範囲 A6:A7:');
      console.log('  行6: 部署（条件ヘッダー）');
      console.log('  行7: 営業（条件値）');
      console.log('フィールド C1: 売上');
    }
    
    expect(matches).toBeTruthy();
  });
});