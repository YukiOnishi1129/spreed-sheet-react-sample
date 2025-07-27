import { SUMPRODUCT } from '../mathLogic';
import type { FormulaContext } from '../../shared/types';

describe('SUMPRODUCT関数のテスト', () => {
  it('2つの配列の積の和を正しく計算する', () => {
    const context: FormulaContext = {
      data: [
        ['配列1', '', '', '配列2', '', '', '結果'],
        [2, 3, 4, 5, 6, 7, '=SUMPRODUCT(A2:C2,D2:F2)']
      ],
      row: 1,
      col: 6
    };

    const matches = ['SUMPRODUCT(A2:C2,D2:F2)', 'A2:C2,D2:F2'];
    const result = SUMPRODUCT.calculate(matches as RegExpMatchArray, context);
    
    // 2*5 + 3*6 + 4*7 = 10 + 18 + 28 = 56
    expect(result).toBe(56);
  });

  it('単一の配列の合計を計算する', () => {
    const context: FormulaContext = {
      data: [
        ['値'],
        [10],
        [20],
        [30]
      ],
      row: 1,
      col: 1
    };

    const matches = ['SUMPRODUCT(A2:A4)', 'A2:A4'];
    const result = SUMPRODUCT.calculate(matches as RegExpMatchArray, context);
    
    // 10 + 20 + 30 = 60
    expect(result).toBe(60);
  });
});