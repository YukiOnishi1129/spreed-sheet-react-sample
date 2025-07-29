/* eslint-disable @typescript-eslint/no-non-null-assertion */
import { describe, it, expect } from 'vitest';
import { 
  ABS, ROUND, TRUNC,
  SQRT, POWER,
  PI, RADIANS, DEGREES,
  SUM, MOD, SUMIF,
  INT, RAND, RANDBETWEEN,
  MAX, MIN, AVERAGE, COUNT
} from '../mathLogic';
import type { FormulaContext } from '../../shared/types';

// モックコンテキスト作成
const createContext = (data: (string | number | boolean | null)[][]): FormulaContext => ({
  data: data.map(row => row.map(cell => ({ value: cell }))),
  row: 0,
  col: 0
});

// テスト用のマッチ配列作成
const createMatches = (formula: string, ...groups: string[]): RegExpMatchArray => {
  const matches = [formula, ...groups] as RegExpMatchArray;
  matches.index = 0;
  matches.input = formula;
  return matches;
};

describe('数学関数のテスト', () => {
  describe('基本的な数学関数', () => {
    it('ABS関数が正しく動作する', () => {
      const context = createContext([[-5], [10], [-3.14]]);
      
      // ABS(A1) = 5
      const result1 = ABS.calculate!(createMatches('ABS(A1)', 'A1'), context);
      expect(result1).toBe(5);
      
      // ABS(A2) = 10
      const result2 = ABS.calculate!(createMatches('ABS(A2)', 'A2'), { ...context, row: 1 });
      expect(result2).toBe(10);
      
      // ABS(A3) = 3.14
      const result3 = ABS.calculate!(createMatches('ABS(A3)', 'A3'), { ...context, row: 2 });
      expect(result3).toBe(3.14);
    });

    it('INT関数が正しく動作する', () => {
      const context = createContext([[5.9], [0.1], [-5.9]]);
      
      expect(INT.calculate!(createMatches('INT(A1)', 'A1'), context)).toBe(5);
      expect(INT.calculate!(createMatches('INT(A2)', 'A2'), { ...context, row: 1 })).toBe(0);
      expect(INT.calculate!(createMatches('INT(A3)', 'A3'), { ...context, row: 2 })).toBe(-6);
    });

    it('ROUND関数が正しく動作する', () => {
      const context = createContext([[3.14159, 2]]);
      
      // ROUND(3.14159, 2) = 3.14
      const result = ROUND.calculate!(createMatches('ROUND(A1,B1)', 'A1,B1'), context);
      expect(result).toBe(3.14);
    });

    it('TRUNC関数が正しく動作する', () => {
      const context = createContext([[3.14159, 2]]);
      
      // TRUNC(3.14159, 2) = 3.14
      const result = TRUNC.calculate!(createMatches('TRUNC(A1,B1)', 'A1', 'B1'), context);
      expect(result).toBe(3.14);
    });
  });

  describe('指数・対数関数', () => {
    it('SQRT関数が正しく動作する', () => {
      const context = createContext([[16], [25], [2]]);
      
      expect(SQRT.calculate!(createMatches('SQRT(A1)', 'A1'), context)).toBe(4);
      expect(SQRT.calculate!(createMatches('SQRT(A2)', 'A2'), { ...context, row: 1 })).toBe(5);
      expect(SQRT.calculate!(createMatches('SQRT(A3)', 'A3'), { ...context, row: 2 })).toBeCloseTo(1.414213, 5);
    });

    it('POWER関数が正しく動作する', () => {
      const context = createContext([[2, 3]]);
      
      // POWER(2, 3) = 8
      const result = POWER.calculate!(createMatches('POWER(A1,B1)', 'A1', 'B1'), context);
      expect(result).toBe(8);
    });

    it('MAX/MIN関数が正しく動作する', () => {
      const context = createContext([[10], [40], [20], [30]]);
      
      // MAX(A1:A4) = 40
      expect(MAX.calculate!(createMatches('MAX(A1:A4)', 'A1:A4'), context)).toBe(40);
      
      // MIN(A1:A4) = 10
      expect(MIN.calculate!(createMatches('MIN(A1:A4)', 'A1:A4'), context)).toBe(10);
    });
  });

  describe('その他の関数', () => {
    it('PI関数が正しく動作する', () => {
      const context = createContext([[]]);
      const result = PI.calculate!(createMatches('PI()'), context);
      expect(result).toBeCloseTo(Math.PI, 5);
    });

    it('RADIANS/DEGREES関数が正しく動作する', () => {
      const context = createContext([[180], [Math.PI]]);
      
      // RADIANS(180) = π
      expect(RADIANS.calculate!(createMatches('RADIANS(A1)', 'A1'), context)).toBeCloseTo(Math.PI, 5);
      
      // DEGREES(π) = 180
      expect(DEGREES.calculate!(createMatches('DEGREES(A2)', 'A2'), { ...context, row: 1 })).toBeCloseTo(180, 5);
    });

    it('RAND/RANDBETWEEN関数が正しく動作する', () => {
      const context = createContext([[1, 10]]);
      
      // RAND() returns a value between 0 and 1
      const randResult = RAND.calculate!(createMatches('RAND()'), context);
      expect(typeof randResult).toBe('number');
      expect(randResult).toBeGreaterThanOrEqual(0);
      expect(randResult).toBeLessThan(1);
      
      // RANDBETWEEN(1, 10) returns a value between 1 and 10
      const randBetweenResult = RANDBETWEEN.calculate!(
        createMatches('RANDBETWEEN(A1,B1)', 'A1', 'B1'), 
        context
      );
      expect(typeof randBetweenResult).toBe('number');
      expect(randBetweenResult).toBeGreaterThanOrEqual(1);
      expect(randBetweenResult).toBeLessThanOrEqual(10);
    });
  });

  describe('集計関数', () => {
    it('SUM関数が正しく動作する', () => {
      const context = createContext([[1], [2], [3], [4], [5]]);
      
      // SUM(A1:A5) = 15
      const result = SUM.calculate!(createMatches('SUM(A1:A5)', 'A1:A5'), context);
      expect(result).toBe(15);
    });

    it('AVERAGE/COUNT関数が正しく動作する', () => {
      const context = createContext([[10], [20], [30], ['text'], [40]]);
      
      // AVERAGE(A1:A5) = 25 (テキストは無視)
      const avgResult = AVERAGE.calculate!(createMatches('AVERAGE(A1:A5)', 'A1:A5'), context);
      expect(avgResult).toBe(25);
      
      // COUNT(A1:A5) = 4 (数値のみカウント)
      const countResult = COUNT.calculate!(createMatches('COUNT(A1:A5)', 'A1:A5'), context);
      expect(countResult).toBe(4);
    });

    it('SUMIF関数が正しく動作する', () => {
      const context = createContext([
        [1, 10],
        [2, 20],
        [3, 30],
        [2, 40]
      ]);
      
      // SUMIF(A1:A4, 2, B1:B4) = 60 (20 + 40)
      const result = SUMIF.calculate!(createMatches('SUMIF(A1:A4,2,B1:B4)', 'A1:A4', '2', 'B1:B4'), context);
      expect(result).toBe(60);
    });

    it('MOD関数が正しく動作する', () => {
      const context = createContext([[10, 3]]);
      
      // MOD(10, 3) = 1
      const result = MOD.calculate!(createMatches('MOD(A1,B1)', 'A1', 'B1'), context);
      expect(result).toBe(1);
    });
  });

  describe('エラーハンドリング', () => {
    it('無効な入力でエラーを返す', () => {
      const context = createContext([['text'], [-1]]);
      
      // SQRT("text") = #NUM! (テキストは0として扱われる)
      expect(SQRT.calculate!(createMatches('SQRT(A1)', 'A1'), context)).toBe('#NUM!');
      
      // SQRT(-1) = #NUM!
      expect(SQRT.calculate!(createMatches('SQRT(A2)', 'A2'), { ...context, row: 1 })).toBe('#NUM!');
    });
  });
});