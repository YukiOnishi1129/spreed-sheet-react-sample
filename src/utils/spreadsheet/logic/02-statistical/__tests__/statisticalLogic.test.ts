/* eslint-disable @typescript-eslint/no-non-null-assertion */
import { describe, it, expect } from 'vitest';
import { 
  AVERAGE, COUNT, COUNTA, COUNTBLANK,
  MAX, MIN, MEDIAN,
  STDEV_S, VAR_S
} from '../basicStatisticsLogic';
import { 
  GAMMALN, PERCENTILE_EXC, PERCENTILE_INC,
  CORREL
} from '../advancedStatisticsLogic';
import { 
  CHISQ_TEST, F_TEST,
  NORM_DIST, NORM_INV
} from '../distributionLogic';
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

describe('基本統計関数のテスト', () => {
  describe('平均値関数', () => {
    it('AVERAGE関数が正しく動作する', () => {
      const context = createContext([[10], [20], [30], [40]]);
      
      // AVERAGE(A1:A4) = 25
      const result = AVERAGE.calculate!(createMatches('AVERAGE(A1:A4)', 'A1:A4'), context);
      expect(result).toBe(25);
    });

  });

  describe('カウント関数', () => {
    it('COUNT関数が正しく動作する', () => {
      const context = createContext([[10], [20], ['text'], [30], ['']]);
      
      // COUNT(A1:A5) = 3
      const result = COUNT.calculate!(createMatches('COUNT(A1:A5)', 'A1:A5'), context);
      expect(result).toBe(3);
    });

    it('COUNTA関数が正しく動作する', () => {
      const context = createContext([[10], [20], ['text'], [''], [null]]);
      
      // COUNTA(A1:A5) = 3
      const result = COUNTA.calculate!(createMatches('COUNTA(A1:A5)', 'A1:A5'), context);
      expect(result).toBe(3);
    });

    it('COUNTBLANK関数が正しく動作する', () => {
      const context = createContext([[10], [''], ['text'], [''], [null]]);
      
      // COUNTBLANK(A1:A5) = 3
      const result = COUNTBLANK.calculate!(createMatches('COUNTBLANK(A1:A5)', 'A1:A5'), context);
      expect(result).toBe(3);
    });

  });

  describe('最大値・最小値・中央値', () => {
    it('MAX関数が正しく動作する', () => {
      const context = createContext([[10], [40], [20], [30]]);
      
      // MAX(A1:A4) = 40
      const result = MAX.calculate!(createMatches('MAX(A1:A4)', 'A1:A4'), context);
      expect(result).toBe(40);
    });

    it('MIN関数が正しく動作する', () => {
      const context = createContext([[10], [40], [20], [30]]);
      
      // MIN(A1:A4) = 10
      const result = MIN.calculate!(createMatches('MIN(A1:A4)', 'A1:A4'), context);
      expect(result).toBe(10);
    });

    it('MEDIAN関数が正しく動作する', () => {
      const context = createContext([[1], [2], [3], [4], [5]]);
      
      // MEDIAN(A1:A5) = 3
      const result = MEDIAN.calculate!(createMatches('MEDIAN(A1:A5)', 'A1:A5'), context);
      expect(result).toBe(3);
    });
  });

  describe('標準偏差・分散', () => {
    it('STDEV_S関数が正しく動作する', () => {
      const context = createContext([[2], [4], [4], [4], [5], [5], [7], [9]]);
      
      // STDEV_S(A1:A8) ≈ 2.138
      const result = STDEV_S.calculate!(createMatches('STDEV.S(A1:A8)', 'A1:A8'), context);
      expect(result).toBeCloseTo(2.138, 2);
    });

    it('VAR_S関数が正しく動作する', () => {
      const context = createContext([[2], [4], [4], [4], [5], [5], [7], [9]]);
      
      // VAR_S(A1:A8) ≈ 4.571
      const result = VAR_S.calculate!(createMatches('VAR.S(A1:A8)', 'A1:A8'), context);
      expect(result).toBeCloseTo(4.571, 2);
    });
  });
});

describe('高度な統計関数のテスト', () => {
  describe('GAMMALN関数', () => {
    it('GAMMALN関数が正しく動作する', () => {
      const context = createContext([[2], [3], [4], [5]]);
      
      // GAMMALN(2) = 0
      expect(GAMMALN.calculate!(createMatches('GAMMALN(A1)', 'A1'), context)).toBe(0);
      
      // GAMMALN(3) = 0.693147
      expect(GAMMALN.calculate!(
        createMatches('GAMMALN(A2)', 'A2'), 
        { ...context, row: 1 }
      )).toBeCloseTo(0.693147, 5);
      
      // GAMMALN(4) = 1.791759
      expect(GAMMALN.calculate!(
        createMatches('GAMMALN(A3)', 'A3'), 
        { ...context, row: 2 }
      )).toBeCloseTo(1.791759, 5);
      
      // GAMMALN(5) = 3.178054
      expect(GAMMALN.calculate!(
        createMatches('GAMMALN(A4)', 'A4'), 
        { ...context, row: 3 }
      )).toBeCloseTo(3.178054, 5);
    });
  });

  describe('パーセンタイル関数', () => {
    it('PERCENTILE.INC関数が正しく動作する', () => {
      const context = createContext([[1], [2], [3], [4], [5]]);
      
      // PERCENTILE.INC(A1:A5, 0.25) = 2
      const result = PERCENTILE_INC.calculate!(
        createMatches('PERCENTILE.INC(A1:A5,0.25)', 'A1:A5', '0.25'),
        context
      );
      expect(result).toBe(2);
    });

    it('PERCENTILE.EXC関数が正しく動作する', () => {
      const context = createContext([[1], [2], [3], [4], [5]]);
      
      // PERCENTILE.EXC(A1:A5, 0.25) = 1.75
      const result = PERCENTILE_EXC.calculate!(
        createMatches('PERCENTILE.EXC(A1:A5,0.25)', 'A1:A5', '0.25'),
        context
      );
      expect(result).toBe(1.5);
    });
  });

  describe('相関係数', () => {
    it('CORREL関数が正しく動作する', () => {
      const context = createContext([
        [1, 2],
        [2, 4],
        [3, 6],
        [4, 8]
      ]);
      
      // CORREL(A1:A4, B1:B4) = 1 (完全な正の相関)
      const result = CORREL.calculate!(
        createMatches('CORREL(A1:A4,B1:B4)', 'A1:A4', 'B1:B4'),
        context
      );
      expect(result).toBeCloseTo(1, 5);
    });
  });
});

describe('分布関数のテスト', () => {
  describe('正規分布', () => {
    it('NORM.DIST関数が正しく動作する', () => {
      const context = createContext([[0, 0, 1, 'TRUE']]);
      
      // NORM.DIST(0, 0, 1, TRUE) = 0.5
      const result = NORM_DIST.calculate!(
        createMatches('NORM.DIST(A1,B1,C1,D1)', 'A1', 'B1', 'C1', 'D1'),
        context
      );
      expect(result).toBeCloseTo(0.5, 5);
    });

    it('NORM.INV関数が正しく動作する', () => {
      const context = createContext([[0.5, 0, 1]]);
      
      // NORM.INV(0.5, 0, 1) = 0
      const result = NORM_INV.calculate!(
        createMatches('NORM.INV(A1,B1,C1)', 'A1', 'B1', 'C1'),
        context
      );
      expect(result).toBeCloseTo(0, 5);
    });
  });

  describe('検定関数', () => {
    it('CHISQ.TEST関数が正しく動作する', () => {
      const context = createContext([[1, 2, 1.5, 2.5], [3, 4, 3.5, 4.5]]);
      
      // CHISQ.TEST(A1:B2, C1:D2)
      const result = CHISQ_TEST.calculate!(
        createMatches('CHISQ.TEST(A1:B2,C1:D2)', 'A1:B2', 'C1:D2'),
        context
      );
      expect(typeof result).toBe('number');
      expect(result).toBeGreaterThan(0);
      expect(result).toBeLessThanOrEqual(1);
    });

    it('F.TEST関数が正しく動作する', () => {
      const context = createContext([
        [10], [20], [30], [40],
        [15], [25], [35], [45]
      ]);
      
      // F.TEST(A1:A4, A5:A8) 
      const result = F_TEST.calculate!(
        createMatches('F.TEST(A1:A4,A5:A8)', 'A1:A4', 'A5:A8'),
        context
      );
      expect(typeof result).toBe('number');
      expect(result).toBeGreaterThanOrEqual(0);
      expect(result).toBeLessThanOrEqual(1);
    });
  });

  describe('エラーハンドリング', () => {
    it('無効な入力でエラーを返す', () => {
      const context = createContext([['text']]);
      
      // AVERAGE("text") = #DIV/0! (数値がない場合)
      expect(AVERAGE.calculate!(createMatches('AVERAGE(A1)', 'A1'), context)).toBe('#DIV/0!');
      
      // STDEV_S(A1) = #DIV/0! (1つの値では計算不可)
      const context2 = createContext([[10]]);
      expect(STDEV_S.calculate!(createMatches('STDEV.S(A1)', 'A1'), context2)).toBe('#DIV/0!');
      
      // PERCENTILE.EXC(A1:A5, 2) = #NUM! (kが範囲外)
      const context3 = createContext([[1], [2], [3], [4], [5]]);
      expect(PERCENTILE_EXC.calculate!(
        createMatches('PERCENTILE.EXC(A1:A5,2)', 'A1:A5', '2'),
        context3
      )).toBe('#NUM!');
    });
  });
});