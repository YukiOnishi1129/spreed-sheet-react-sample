import { describe, it, expect } from 'vitest';
import { ISO_CEILING } from '../mathLogic';
import type { FormulaContext } from '../../shared/types';

describe('ISO.CEILING関数のデバッグ', () => {
  it('正の数のISO.CEILING計算', () => {
    const context: FormulaContext = {
      data: []
    };

    const formula = 'ISO.CEILING(4.3,1)';
    const matches = formula.match(ISO_CEILING.pattern);
    
    if (matches) {
      const result = ISO_CEILING.calculate(matches, context);
      console.log('ISO.CEILING(4.3, 1) =', result);
      console.log('期待値: 5');
    }
  });

  it('負の数のISO.CEILING計算', () => {
    const context: FormulaContext = {
      data: []
    };

    const formula = 'ISO.CEILING(-4.3,1)';
    const matches = formula.match(ISO_CEILING.pattern);
    
    if (matches) {
      const result = ISO_CEILING.calculate(matches, context);
      console.log('ISO.CEILING(-4.3, 1) =', result);
      console.log('期待値: -4');
    }
  });

  it('負の数の詳細な計算過程', () => {
    const context: FormulaContext = {
      data: []
    };

    // ISO.CEILINGの負の数の動作を確認
    const testCases = [
      { num: -4.3, sig: 1, expected: -4 },
      { num: -4.3, sig: 2, expected: -4 },
      { num: -5.5, sig: 1, expected: -5 },
      { num: -5.5, sig: 2, expected: -4 }
    ];

    testCases.forEach(({ num, sig, expected }) => {
      const formula = `ISO.CEILING(${num},${sig})`;
      const matches = formula.match(ISO_CEILING.pattern);
      
      if (matches) {
        const result = ISO_CEILING.calculate(matches, context);
        console.log(`ISO.CEILING(${num}, ${sig}) = ${result}, 期待値: ${expected}`);
        
        // 手動計算
        const manual = Math.ceil(num / Math.abs(sig)) * Math.abs(sig);
        console.log(`  手動計算: Math.ceil(${num} / ${Math.abs(sig)}) * ${Math.abs(sig)} = ${manual}`);
      }
    });
  });
});