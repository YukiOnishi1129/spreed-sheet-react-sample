import { describe, it, expect } from 'vitest';
import type { FormulaContext } from '../../shared/types';
import {
  MODE_SNGL, MODE_MULT,
  RANK_AVG, RANK_EQ
} from '../modeRankLogic';
import { FormulaError } from '../../shared/types';

// Helper function to create FormulaContext
function createContext(data: (string | number | boolean | null)[][]): FormulaContext {
  return {
    data: data.map(row => row.map(cell => ({ value: cell }))),
    row: 0,
    col: 0
  };
}

describe('Mode and Rank Functions', () => {
  const mockContext = createContext([
    [1, 5, 10, 'text', 100],   // Row 1 (A1-E1)
    [2, 5, 20, '', 90],        // Row 2 (A2-E2)
    [2, 5, 30, null, 90],      // Row 3 (A3-E3)
    [3, 6, 40, null, 80],      // Row 4 (A4-E4)
    [3, 6, 50, null, 70],      // Row 5 (A5-E5)
    [3, null, null, null, null], // Row 6 (A6-E6)
    [4, null, null, null, null]  // Row 7 (A7-E7)
  ]);

  describe('Mode Functions', () => {
    describe('MODE.SNGL', () => {
      it('should find single most frequent value', () => {
        const matches = ['MODE.SNGL(A1:A7)', 'A1:A7'] as RegExpMatchArray;
        const result = MODE_SNGL.calculate(matches, mockContext);
        expect(result).toBe(3); // appears 3 times
      });

      it('should return first mode when multiple exist', () => {
        const matches = ['MODE.SNGL(B1:B5)', 'B1:B5'] as RegExpMatchArray;
        const result = MODE_SNGL.calculate(matches, mockContext);
        expect(result).toBe(5); // both 5 and 6 appear, but 5 comes first
      });

      it('should return NA error when no mode exists', () => {
        const matches = ['MODE.SNGL(C1:C5)', 'C1:C5'] as RegExpMatchArray;
        const result = MODE_SNGL.calculate(matches, mockContext);
        expect(result).toBe(FormulaError.NA);
      });

      it('should ignore non-numeric values', () => {
        const matches = ['MODE.SNGL(A1:D3)', 'A1:D3'] as RegExpMatchArray;
        const result = MODE_SNGL.calculate(matches, mockContext);
        expect(result).toBe(5); // 5 appears 3 times, which is most frequent
      });

      it('should return NUM error for no numeric values', () => {
        const matches = ['MODE.SNGL(D1:D3)', 'D1:D3'] as RegExpMatchArray;
        const result = MODE_SNGL.calculate(matches, mockContext);
        expect(result).toBe(FormulaError.NUM);
      });

      it('should handle single value', () => {
        const matches = ['MODE.SNGL(A1)', 'A1'] as RegExpMatchArray;
        const result = MODE_SNGL.calculate(matches, mockContext);
        expect(result).toBe(1);
      });

      it('should handle multiple ranges', () => {
        const matches = ['MODE.SNGL(A1:A3, A5:A6)', 'A1:A3, A5:A6'] as RegExpMatchArray;
        const result = MODE_SNGL.calculate(matches, mockContext);
        expect(result).toBe(3); // 3 appears twice in combined ranges
      });
    });

    describe('MODE.MULT', () => {
      it('should return array of all modes', () => {
        const matches = ['MODE.MULT(B1:B5)', 'B1:B5'] as RegExpMatchArray;
        const result = MODE_MULT.calculate(matches, mockContext);
        expect(result).toEqual([5]); // 5 appears 3 times, which is most frequent
      });

      it('should return single value array when one mode', () => {
        const matches = ['MODE.MULT(A1:A7)', 'A1:A7'] as RegExpMatchArray;
        const result = MODE_MULT.calculate(matches, mockContext);
        expect(result).toEqual([3]);
      });

      it('should return NA error when no mode exists', () => {
        const matches = ['MODE.MULT(C1:C5)', 'C1:C5'] as RegExpMatchArray;
        const result = MODE_MULT.calculate(matches, mockContext);
        expect(result).toBe(FormulaError.NA);
      });

      it('should handle complex multi-modal data', () => {
        const mockContextMultiModal = createContext([
          [1],
          [1],
          [2],
          [2],
          [3],
          [3],
          [4],
          [4]
        ]);
        const matches = ['MODE.MULT(A1:A8)', 'A1:A8'] as RegExpMatchArray;
        const result = MODE_MULT.calculate(matches, mockContextMultiModal);
        expect(result).toEqual([1, 2, 3, 4]); // all appear twice
      });

      it('should return NUM error for no numeric values', () => {
        const matches = ['MODE.MULT(D1:D3)', 'D1:D3'] as RegExpMatchArray;
        const result = MODE_MULT.calculate(matches, mockContext);
        expect(result).toBe(FormulaError.NUM);
      });
    });
  });

  describe('Rank Functions', () => {
    describe('RANK.AVG', () => {
      it('should rank with average for ties in descending order', () => {
        const matches = ['RANK.AVG(90, E1:E5)', '90', 'E1:E5'] as RegExpMatchArray;
        const result = RANK_AVG.calculate(matches, mockContext);
        expect(result).toBe(2.5); // tied for 2nd and 3rd, so average = 2.5
      });

      it('should rank in ascending order when specified', () => {
        const matches = ['RANK.AVG(90, E1:E5, 1)', '90', 'E1:E5', '1'] as RegExpMatchArray;
        const result = RANK_AVG.calculate(matches, mockContext);
        expect(result).toBe(3.5); // tied for 3rd and 4th in ascending order
      });

      it('should handle no ties', () => {
        const matches = ['RANK.AVG(30, C1:C5)', '30', 'C1:C5'] as RegExpMatchArray;
        const result = RANK_AVG.calculate(matches, mockContext);
        expect(result).toBe(3); // 50, 40, 30, 20, 10 in descending
      });

      it('should return NA error for value not in list', () => {
        const matches = ['RANK.AVG(60, C1:C5)', '60', 'C1:C5'] as RegExpMatchArray;
        const result = RANK_AVG.calculate(matches, mockContext);
        expect(result).toBe(FormulaError.NA);
      });

      it('should ignore non-numeric values', () => {
        const matches = ['RANK.AVG(2, A1:A4)', '2', 'A1:A4'] as RegExpMatchArray;
        const result = RANK_AVG.calculate(matches, mockContext);
        expect(result).toBe(2.5); // 2 appears twice, tied for 2nd and 3rd
      });

      it('should handle single value', () => {
        const matches = ['RANK.AVG(10, C1)', '10', 'C1'] as RegExpMatchArray;
        const result = RANK_AVG.calculate(matches, mockContext);
        expect(result).toBe(1);
      });

      it('should return NUM error for empty array', () => {
        const matches = ['RANK.AVG(5, D1:D3)', '5', 'D1:D3'] as RegExpMatchArray;
        const result = RANK_AVG.calculate(matches, mockContext);
        expect(result).toBe(FormulaError.NUM);
      });
    });

    describe('RANK.EQ', () => {
      it('should rank with same rank for ties', () => {
        const matches = ['RANK.EQ(90, E1:E5)', '90', 'E1:E5'] as RegExpMatchArray;
        const result = RANK_EQ.calculate(matches, mockContext);
        expect(result).toBe(2); // tied values get the same rank
      });

      it('should rank in ascending order when specified', () => {
        const matches = ['RANK.EQ(90, E1:E5, 1)', '90', 'E1:E5', '1'] as RegExpMatchArray;
        const result = RANK_EQ.calculate(matches, mockContext);
        expect(result).toBe(3); // in ascending: 70, 80, 90(tied), 90(tied), 100
      });

      it('should handle unique values correctly', () => {
        const matches = ['RANK.EQ(30, C1:C5)', '30', 'C1:C5'] as RegExpMatchArray;
        const result = RANK_EQ.calculate(matches, mockContext);
        expect(result).toBe(3);
      });

      it('should return NA error for value not in list', () => {
        const matches = ['RANK.EQ(60, C1:C5)', '60', 'C1:C5'] as RegExpMatchArray;
        const result = RANK_EQ.calculate(matches, mockContext);
        expect(result).toBe(FormulaError.NA);
      });

      it('should handle multiple occurrences of same value', () => {
        const matches = ['RANK.EQ(3, A1:A7)', '3', 'A1:A7'] as RegExpMatchArray;
        const result = RANK_EQ.calculate(matches, mockContext);
        expect(result).toBe(2); // 4 is 1st, 3 is 2nd (appears 3 times)
      });

      it('should treat order=0 as descending', () => {
        const matches = ['RANK.EQ(90, E1:E5, 0)', '90', 'E1:E5', '0'] as RegExpMatchArray;
        const result = RANK_EQ.calculate(matches, mockContext);
        expect(result).toBe(2);
      });

      it('should return VALUE error for non-numeric value', () => {
        const matches = ['RANK.EQ(text, E1:E5)', 'text', 'E1:E5'] as RegExpMatchArray;
        const result = RANK_EQ.calculate(matches, mockContext);
        expect(result).toBe(FormulaError.VALUE);
      });
    });
  });
});