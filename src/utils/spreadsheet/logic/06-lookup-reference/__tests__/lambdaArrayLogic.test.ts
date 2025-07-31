import { describe, it, expect } from 'vitest';
import { BYROW, BYCOL, MAP, REDUCE, SCAN, MAKEARRAY } from '../lambdaArrayLogic';
import type { FormulaContext } from '../../shared/types';
import { FormulaError } from '../../shared/types';

describe('Lambda配列関数', () => {
  describe('BYROW関数', () => {
    it('各行の合計を計算', () => {
      const context: FormulaContext = {
        data: [
          [{ value: 1 }, { value: 2 }, { value: 3 }],
          [{ value: 4 }, { value: 5 }, { value: 6 }],
          [{ value: 7 }, { value: 8 }, { value: 9 }]
        ],
        row: 0,
        col: 0
      };
      
      const matches = ['BYROW(A1:C3, LAMBDA(row, SUM(row)))', 'A1:C3', 'LAMBDA(row, SUM(row))'];
      const result = BYROW.calculate(matches as RegExpMatchArray, context);
      
      expect(result).toEqual([[6], [15], [24]]);
    });

    it('各行の平均を計算', () => {
      const context: FormulaContext = {
        data: [
          [{ value: 10 }, { value: 20 }, { value: 30 }],
          [{ value: 40 }, { value: 50 }, { value: 60 }]
        ],
        row: 0,
        col: 0
      };
      
      const matches = ['BYROW(A1:C2, LAMBDA(row, AVERAGE(row)))', 'A1:C2', 'LAMBDA(row, AVERAGE(row))'];
      const result = BYROW.calculate(matches as RegExpMatchArray, context);
      
      expect(result).toEqual([[20], [50]]);
    });

    it('各行の最大値を取得', () => {
      const context: FormulaContext = {
        data: [
          [{ value: 5 }, { value: 3 }, { value: 8 }],
          [{ value: 2 }, { value: 9 }, { value: 1 }]
        ],
        row: 0,
        col: 0
      };
      
      const matches = ['BYROW(A1:C2, LAMBDA(row, MAX(row)))', 'A1:C2', 'LAMBDA(row, MAX(row))'];
      const result = BYROW.calculate(matches as RegExpMatchArray, context);
      
      expect(result).toEqual([[8], [9]]);
    });

    it('無効な範囲参照でエラー', () => {
      const context: FormulaContext = {
        data: [],
        row: 0,
        col: 0
      };
      
      const matches = ['BYROW(INVALID, LAMBDA(row, SUM(row)))', 'INVALID', 'LAMBDA(row, SUM(row))'];
      const result = BYROW.calculate(matches as RegExpMatchArray, context);
      
      expect(result).toBe(FormulaError.VALUE);
    });
  });

  describe('BYCOL関数', () => {
    it('各列の合計を計算', () => {
      const context: FormulaContext = {
        data: [
          [{ value: 1 }, { value: 2 }, { value: 3 }],
          [{ value: 4 }, { value: 5 }, { value: 6 }],
          [{ value: 7 }, { value: 8 }, { value: 9 }]
        ],
        row: 0,
        col: 0
      };
      
      const matches = ['BYCOL(A1:C3, LAMBDA(column, SUM(column)))', 'A1:C3', 'LAMBDA(column, SUM(column))'];
      const result = BYCOL.calculate(matches as RegExpMatchArray, context);
      
      expect(result).toEqual([[12, 15, 18]]);
    });

    it('各列の平均を計算', () => {
      const context: FormulaContext = {
        data: [
          [{ value: 10 }, { value: 20 }],
          [{ value: 30 }, { value: 40 }],
          [{ value: 50 }, { value: 60 }]
        ],
        row: 0,
        col: 0
      };
      
      const matches = ['BYCOL(A1:B3, LAMBDA(column, AVERAGE(column)))', 'A1:B3', 'LAMBDA(column, AVERAGE(column))'];
      const result = BYCOL.calculate(matches as RegExpMatchArray, context);
      
      expect(result).toEqual([[30, 40]]);
    });

    it('各列の最小値を取得', () => {
      const context: FormulaContext = {
        data: [
          [{ value: 5 }, { value: 3 }],
          [{ value: 2 }, { value: 9 }],
          [{ value: 8 }, { value: 1 }]
        ],
        row: 0,
        col: 0
      };
      
      const matches = ['BYCOL(A1:B3, LAMBDA(column, MIN(column)))', 'A1:B3', 'LAMBDA(column, MIN(column))'];
      const result = BYCOL.calculate(matches as RegExpMatchArray, context);
      
      expect(result).toEqual([[2, 1]]);
    });
  });

  describe('MAP関数', () => {
    it('単一配列の各要素を2倍', () => {
      const context: FormulaContext = {
        data: [
          [{ value: 1 }, { value: 2 }, { value: 3 }],
          [{ value: 4 }, { value: 5 }, { value: 6 }]
        ],
        row: 0,
        col: 0
      };
      
      const matches = ['MAP(A1:C2, LAMBDA(a, a*2))', 'A1:C2, LAMBDA(a, a*2)'];
      const result = MAP.calculate(matches as RegExpMatchArray, context);
      
      expect(result).toEqual([
        [2, 4, 6],
        [8, 10, 12]
      ]);
    });

    it('2つの配列の対応する要素を加算', () => {
      const context: FormulaContext = {
        data: [
          [{ value: 1 }, { value: 2 }],
          [{ value: 3 }, { value: 4 }],
          [{ value: 10 }, { value: 20 }],
          [{ value: 30 }, { value: 40 }]
        ],
        row: 0,
        col: 0
      };
      
      const matches = ['MAP(A1:B2, A3:B4, LAMBDA(a, b, a+b))', 'A1:B2, A3:B4, LAMBDA(a, b, a+b)'];
      const result = MAP.calculate(matches as RegExpMatchArray, context);
      
      expect(result).toEqual([
        [11, 22],
        [33, 44]
      ]);
    });

    it('配列のサイズが一致しない場合エラー', () => {
      const context: FormulaContext = {
        data: [
          [{ value: 1 }, { value: 2 }],
          [{ value: 3 }, { value: 4 }, { value: 5 }]
        ],
        row: 0,
        col: 0
      };
      
      const matches = ['MAP(A1:B1, A2:C2, LAMBDA(a, b, a+b))', 'A1:B1, A2:C2, LAMBDA(a, b, a+b)'];
      const result = MAP.calculate(matches as RegExpMatchArray, context);
      
      expect(result).toBe(FormulaError.VALUE);
    });
  });

  describe('REDUCE関数', () => {
    it('配列の全要素を合計', () => {
      const context: FormulaContext = {
        data: [
          [{ value: 1 }, { value: 2 }, { value: 3 }],
          [{ value: 4 }, { value: 5 }, { value: 6 }]
        ],
        row: 0,
        col: 0
      };
      
      const matches = ['REDUCE(0, A1:C2, LAMBDA(accumulator, value, accumulator+value))', '0, A1:C2, LAMBDA(accumulator, value, accumulator+value)'];
      const result = REDUCE.calculate(matches as RegExpMatchArray, context);
      
      expect(result).toBe(21);
    });

    it('配列の全要素を乗算', () => {
      const context: FormulaContext = {
        data: [
          [{ value: 2 }, { value: 3 }],
          [{ value: 4 }, { value: 5 }]
        ],
        row: 0,
        col: 0
      };
      
      const matches = ['REDUCE(1, A1:B2, LAMBDA(accumulator, value, accumulator*value))', '1, A1:B2, LAMBDA(accumulator, value, accumulator*value)'];
      const result = REDUCE.calculate(matches as RegExpMatchArray, context);
      
      expect(result).toBe(120);
    });

    it('最大値を見つける', () => {
      const context: FormulaContext = {
        data: [
          [{ value: 5 }, { value: 3 }, { value: 8 }],
          [{ value: 2 }, { value: 9 }, { value: 1 }]
        ],
        row: 0,
        col: 0
      };
      
      const matches = ['REDUCE(0, A1:C2, LAMBDA(accumulator, value, MAX(accumulator, value)))', '0, A1:C2, LAMBDA(accumulator, value, MAX(accumulator, value))'];
      const result = REDUCE.calculate(matches as RegExpMatchArray, context);
      
      expect(result).toBe(9);
    });

    it('文字列の初期値を使用', () => {
      const context: FormulaContext = {
        data: [
          [{ value: 'Hello' }, { value: 'World' }]
        ],
        row: 0,
        col: 0
      };
      
      const matches = ['REDUCE("", A1:B1, LAMBDA(accumulator, value, accumulator&value))', '"", A1:B1, LAMBDA(accumulator, value, accumulator&value)'];
      const result = REDUCE.calculate(matches as RegExpMatchArray, context);
      
      expect(result).toBe('HelloWorld');
    });
  });

  describe('SCAN関数', () => {
    it('累積合計を計算', () => {
      const context: FormulaContext = {
        data: [
          [{ value: 1 }, { value: 2 }, { value: 3 }],
          [{ value: 4 }, { value: 5 }, { value: 6 }]
        ],
        row: 0,
        col: 0
      };
      
      const matches = ['SCAN(0, A1:C2, LAMBDA(accumulator, value, accumulator+value))', '0, A1:C2, LAMBDA(accumulator, value, accumulator+value)'];
      const result = SCAN.calculate(matches as RegExpMatchArray, context);
      
      expect(result).toEqual([
        [1, 3, 6],
        [10, 15, 21]
      ]);
    });

    it('累積乗算を計算', () => {
      const context: FormulaContext = {
        data: [
          [{ value: 2 }, { value: 3 }],
          [{ value: 4 }, { value: 5 }]
        ],
        row: 0,
        col: 0
      };
      
      const matches = ['SCAN(1, A1:B2, LAMBDA(accumulator, value, accumulator*value))', '1, A1:B2, LAMBDA(accumulator, value, accumulator*value)'];
      const result = SCAN.calculate(matches as RegExpMatchArray, context);
      
      expect(result).toEqual([
        [2, 6],
        [24, 120]
      ]);
    });

    it('最大値の累積を計算', () => {
      const context: FormulaContext = {
        data: [
          [{ value: 5 }, { value: 3 }, { value: 8 }],
          [{ value: 2 }, { value: 9 }, { value: 1 }]
        ],
        row: 0,
        col: 0
      };
      
      const matches = ['SCAN(0, A1:C2, LAMBDA(accumulator, value, MAX(accumulator, value)))', '0, A1:C2, LAMBDA(accumulator, value, MAX(accumulator, value))'];
      const result = SCAN.calculate(matches as RegExpMatchArray, context);
      
      expect(result).toEqual([
        [5, 5, 8],
        [8, 9, 9]
      ]);
    });
  });

  describe('MAKEARRAY関数', () => {
    it('行番号と列番号の積の配列を作成', () => {
      const context: FormulaContext = {
        data: [],
        row: 0,
        col: 0
      };
      
      const matches = ['MAKEARRAY(3, 4, LAMBDA(row, column, row*column))', '3, 4, LAMBDA(row, column, row*column)'];
      const result = MAKEARRAY.calculate(matches as RegExpMatchArray, context);
      
      expect(result).toEqual([
        [1, 2, 3, 4],
        [2, 4, 6, 8],
        [3, 6, 9, 12]
      ]);
    });

    it('行番号のみの配列を作成', () => {
      const context: FormulaContext = {
        data: [],
        row: 0,
        col: 0
      };
      
      const matches = ['MAKEARRAY(4, 3, LAMBDA(row, column, ROW()))', '4, 3, LAMBDA(row, column, ROW())'];
      const result = MAKEARRAY.calculate(matches as RegExpMatchArray, context);
      
      expect(result).toEqual([
        [1, 1, 1],
        [2, 2, 2],
        [3, 3, 3],
        [4, 4, 4]
      ]);
    });

    it('列番号のみの配列を作成', () => {
      const context: FormulaContext = {
        data: [],
        row: 0,
        col: 0
      };
      
      const matches = ['MAKEARRAY(2, 5, LAMBDA(row, column, COLUMN()))', '2, 5, LAMBDA(row, column, COLUMN())'];
      const result = MAKEARRAY.calculate(matches as RegExpMatchArray, context);
      
      expect(result).toEqual([
        [1, 2, 3, 4, 5],
        [1, 2, 3, 4, 5]
      ]);
    });

    it('セル参照から行数と列数を取得', () => {
      const context: FormulaContext = {
        data: [
          [{ value: 3 }, { value: 2 }]
        ],
        row: 0,
        col: 0
      };
      
      const matches = ['MAKEARRAY(A1, B1, LAMBDA(row, column, row+column))', 'A1, B1, LAMBDA(row, column, row+column)'];
      const result = MAKEARRAY.calculate(matches as RegExpMatchArray, context);
      
      expect(result).toEqual([
        [2, 3],
        [3, 4],
        [4, 5]
      ]);
    });

    it('無効な行数または列数でエラー', () => {
      const context: FormulaContext = {
        data: [],
        row: 0,
        col: 0
      };
      
      const matches = ['MAKEARRAY(0, 5, LAMBDA(row, column, row*column))', '0, 5, LAMBDA(row, column, row*column)'];
      const result = MAKEARRAY.calculate(matches as RegExpMatchArray, context);
      
      expect(result).toBe(FormulaError.VALUE);
    });

    it('非常に大きなサイズでエラー', () => {
      const context: FormulaContext = {
        data: [],
        row: 0,
        col: 0
      };
      
      const matches = ['MAKEARRAY(2000000, 20000, LAMBDA(row, column, row*column))', '2000000, 20000, LAMBDA(row, column, row*column)'];
      const result = MAKEARRAY.calculate(matches as RegExpMatchArray, context);
      
      expect(result).toBe(FormulaError.VALUE);
    });
  });
});