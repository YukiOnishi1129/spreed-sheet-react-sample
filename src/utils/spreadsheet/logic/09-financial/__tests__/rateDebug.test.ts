import { describe, it, expect } from 'vitest';
import { RATE } from '../financialLogic';
import type { FormulaContext } from '../../shared/types';

describe('RATE関数のデバッグ', () => {
  it('基本的なRATE計算', () => {
    const context: FormulaContext = {
      data: [
        ['期間', '支払額', '現在価値', '利率'],
        [12, -1000, 10000, '=RATE(A2,B2,C2)*12']
      ]
    };

    // RATE関数のパターンマッチをテスト
    const formula = 'RATE(A2,B2,C2)';
    const matches = formula.match(RATE.pattern);
    
    console.log('Formula:', formula);
    console.log('Pattern:', RATE.pattern);
    console.log('Matches:', matches);
    
    if (matches) {
      const result = RATE.calculate(matches, context);
      console.log('Result:', result);
      console.log('Result * 12:', typeof result === 'number' ? result * 12 : 'N/A');
      console.log('Expected (annual):', 0.0253);
      
      // セル参照の値を確認
      console.log('A2 (nper):', context.data[1][0]);
      console.log('B2 (pmt):', context.data[1][1]);
      console.log('C2 (pv):', context.data[1][2]);
      
      // 手動計算で検証
      // RATE(12, -1000, 10000) should return approximately 0.002108
      // Annual rate = 0.002108 * 12 ≈ 0.0253
    }
    
    expect(matches).toBeTruthy();
  });

  it('直接値でのRATE計算', () => {
    const context: FormulaContext = {
      data: []
    };

    // RATE(12, -1000, 10000)の直接計算
    const formula = 'RATE(12,-1000,10000)';
    const matches = formula.match(RATE.pattern);
    
    if (matches) {
      const result = RATE.calculate(matches, context);
      console.log('Direct calculation result:', result);
      console.log('Annual rate:', typeof result === 'number' ? result * 12 : 'N/A');
      
      // Excelでの期待値: RATE(12,-1000,10000) = 0.002107897
      // 年利 = 0.002107897 * 12 = 0.02529476
      console.log('Excel expected monthly rate:', 0.002107897);
      console.log('Excel expected annual rate:', 0.02529476);
    }
  });
});