import { describe, it, expect } from 'vitest';
import { GAMMALN, PERCENTILE_EXC } from '../02-statistical/advancedStatisticsLogic';
import { CHISQ_TEST, F_TEST, HYPGEOM_DIST } from '../02-statistical/distributionLogic';
import type { FormulaContext } from '../shared/types';

// モックコンテキスト
const createMockContext = (data: (string | number | boolean | null)[][]): FormulaContext => ({
  data: data.map(row => row.map(cell => ({ value: cell }))),
  row: 0,
  col: 0
});

describe('統計関数のテスト', () => {
  describe('GAMMALN', () => {
    it('GAMMALN(2)は0を返すべき', () => {
      const context = createMockContext([[2]]);
      const matches = ['GAMMALN(A1)', 'A1'] as RegExpMatchArray;
      matches.index = 0;
      matches.input = 'GAMMALN(A1)';
      
      const result = GAMMALN.calculate!(matches, context);
      expect(result).toBe(0);
    });

    it('GAMMALN(3)は0.693147を返すべき', () => {
      const context = createMockContext([[3]]);
      const matches = ['GAMMALN(A1)', 'A1'] as RegExpMatchArray;
      matches.index = 0;
      matches.input = 'GAMMALN(A1)';
      
      const result = GAMMALN.calculate!(matches, context);
      expect(result).toBeCloseTo(0.693147, 5);
    });

    it('GAMMALN(4)は1.791759を返すべき', () => {
      const context = createMockContext([[4]]);
      const matches = ['GAMMALN(A1)', 'A1'] as RegExpMatchArray;
      matches.index = 0;
      matches.input = 'GAMMALN(A1)';
      
      const result = GAMMALN.calculate!(matches, context);
      expect(result).toBeCloseTo(1.791759, 5);
    });

    it('GAMMALN(5)は3.178054を返すべき', () => {
      const context = createMockContext([[5]]);
      const matches = ['GAMMALN(A1)', 'A1'] as RegExpMatchArray;
      matches.index = 0;
      matches.input = 'GAMMALN(A1)';
      
      const result = GAMMALN.calculate!(matches, context);
      expect(result).toBeCloseTo(3.178054, 5);
    });
  });

  describe('CHISQ.TEST', () => {
    it('CHISQ.TESTは0.807を返すべき', () => {
      const context = createMockContext([[1, 2], [3, 4]]);
      const matches = ['CHISQ.TEST(A1:B2,C1:D2)', 'A1:B2', 'C1:D2'] as RegExpMatchArray;
      matches.index = 0;
      matches.input = 'CHISQ.TEST(A1:B2,C1:D2)';
      
      const result = CHISQ_TEST.calculate!(matches, context);
      expect(result).toBe(0.807);
    });
  });

  describe('PERCENTILE.EXC', () => {
    it('PERCENTILE.EXC([1,2,3,4,5], 0.25)は1.75を返すべき', () => {
      const context = createMockContext([[1], [2], [3], [4], [5]]);
      const matches = ['PERCENTILE.EXC(A1:A5,0.25)', 'A1:A5', '0.25'] as RegExpMatchArray;
      matches.index = 0;
      matches.input = 'PERCENTILE.EXC(A1:A5,0.25)';
      
      const result = PERCENTILE_EXC.calculate!(matches, context);
      expect(result).toBe(1.75);
    });
  });

  describe('F.TEST', () => {
    it('F.TESTは期待値を返すべき', () => {
      const context = createMockContext([[10], [20], [30], [40], [15], [25], [35], [45]]);
      const matches = ['F.TEST(A1:A4,A5:A8)', 'A1:A4', 'A5:A8'] as RegExpMatchArray;
      matches.index = 0;
      matches.input = 'F.TEST(A1:A4,A5:A8)';
      
      const result = F_TEST.calculate!(matches, context);
      expect(result).toBeCloseTo(0.646, 2);
    });
  });

  describe('HYPGEOM.DIST', () => {
    it('HYPGEOM.DISTは期待値を返すべき', () => {
      const context = createMockContext([[1], [4], [8], [20], ['TRUE']]);
      const matches = ['HYPGEOM.DIST(A1,B1,C1,D1,E1)', 'A1', 'B1', 'C1', 'D1', 'E1'] as RegExpMatchArray;
      matches.index = 0;
      matches.input = 'HYPGEOM.DIST(A1,B1,C1,D1,E1)';
      
      const result = HYPGEOM_DIST.calculate!(matches, context);
      expect(typeof result).toBe('number');
      expect(result).toBeCloseTo(0.238095, 5);
    });
  });
});