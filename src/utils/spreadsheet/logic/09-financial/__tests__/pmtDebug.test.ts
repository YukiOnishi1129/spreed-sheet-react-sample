import { describe, it, expect } from 'vitest';
import { PMT } from '../financialLogic';
import type { FormulaContext } from '../../shared/types';

describe('PMT関数のデバッグ', () => {
  it('基本的なPMT計算', () => {
    const context: FormulaContext = {
      data: [
        ['元本', '月利率', '期間(月)', '月額支払額'],
        [100000, 0.005, 12, '=PMT(B2,C2,-A2)']
      ]
    };

    // PMT関数のパターンマッチをテスト
    const formula = '=PMT(B2,C2,-A2)';
    const matches = formula.match(PMT.pattern);
    
    console.log('Formula:', formula);
    console.log('Pattern:', PMT.pattern);
    console.log('Matches:', matches);
    
    if (matches) {
      const result = PMT.calculate(matches, context);
      console.log('Result:', result);
      console.log('Expected:', 8606.64);
      
      // セル参照の値を確認
      console.log('B2 (rate):', context.data[1][1]);
      console.log('C2 (nper):', context.data[1][2]);
      console.log('A2 (pv):', context.data[1][0]);
    }
    
    expect(matches).toBeTruthy();
  });

  it('セル参照の解決をテスト', () => {
    const context: FormulaContext = {
      data: [
        ['元本', '月利率', '期間(月)', '月額支払額'],
        [100000, 0.005, 12, '']
      ]
    };

    // PMT(0.005, 12, -100000)の直接計算
    const formula = 'PMT(0.005,12,-100000)';
    const matches = formula.match(PMT.pattern);
    
    if (matches) {
      const result = PMT.calculate(matches, context);
      console.log('Direct calculation result:', result);
      
      // Excel PMT formula: PMT = (PV × i) / (1 - (1 + i)^(-n))
      // Where: PV = -100000, i = 0.005, n = 12
      const rate = 0.005;
      const nper = 12;
      const pv = -100000;
      
      const pvFactor = (1 - Math.pow(1 + rate, -nper)) / rate;
      const expectedPmt = -(pv / pvFactor);
      
      console.log('Manual calculation:', expectedPmt);
    }
  });
});