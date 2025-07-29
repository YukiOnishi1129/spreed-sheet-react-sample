import { describe, it, expect } from 'vitest';
import type { FormulaContext } from '../../shared/types';
import { MDETERM, MINVERSE, MMULT, MUNIT } from '../matrixLogic';
import { FormulaError } from '../../shared/types';

// Helper function to create FormulaContext
function createContext(data: any[][]): FormulaContext {
  return {
    data,
    row: 0,
    col: 0
  };
}

describe('Matrix Functions', () => {
  const mockContext = createContext([
    [1, 2, 5, 6, 1, 4, 7, 'text', 3],    // Row 1 (A1-I1)
    [3, 4, 7, 8, 2, 5, 8, 2, 4],         // Row 2 (A2-I2)
    [null, null, null, null, 3, 6, 9, null, null]  // Row 3 (A3-I3)
  ]);

  describe('MDETERM', () => {
    it('should calculate determinant of 2x2 matrix', () => {
      const matches = ['MDETERM(A1:B2)', 'A1:B2'] as RegExpMatchArray;
      const result = MDETERM.calculate(matches, mockContext);
      expect(result).toBe(-2); // 1*4 - 2*3 = -2
    });

    it('should calculate determinant of 3x3 matrix', () => {
      const matches = ['MDETERM(E1:G3)', 'E1:G3'] as RegExpMatchArray;
      const result = MDETERM.calculate(matches, mockContext);
      expect(result).toBeCloseTo(0); // This is a singular matrix
    });

    it('should return VALUE error for non-numeric values', () => {
      const matches = ['MDETERM(H1:I2)', 'H1:I2'] as RegExpMatchArray;
      const result = MDETERM.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });

    it('should return VALUE error for non-square matrix', () => {
      const matches = ['MDETERM(A1:C2)', 'A1:C2'] as RegExpMatchArray;
      const result = MDETERM.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });
  });

  describe('MINVERSE', () => {
    it('should calculate inverse of 2x2 matrix', () => {
      const matches = ['MINVERSE(A1:B2)', 'A1:B2'] as RegExpMatchArray;
      const result = MINVERSE.calculate(matches, mockContext);
      expect(result).toEqual([
        [-2, 1],
        [1.5, -0.5]
      ]);
    });

    it('should return single value for 1x1 matrix', () => {
      const matches = ['MINVERSE(A1:A1)', 'A1:A1'] as RegExpMatchArray;
      const result = MINVERSE.calculate(matches, mockContext);
      expect(result).toBe(1); // 1/1 = 1
    });

    it('should return NUM error for singular matrix', () => {
      const matches = ['MINVERSE(E1:G3)', 'E1:G3'] as RegExpMatchArray;
      const result = MINVERSE.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });

    it('should return VALUE error for non-numeric values', () => {
      const matches = ['MINVERSE(H1:I2)', 'H1:I2'] as RegExpMatchArray;
      const result = MINVERSE.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });

    it('should return NUM error for non-square matrix', () => {
      const matches = ['MINVERSE(A1:C2)', 'A1:C2'] as RegExpMatchArray;
      const result = MINVERSE.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });
  });

  describe('MMULT', () => {
    it('should multiply two 2x2 matrices', () => {
      const matches = ['MMULT(A1:B2, C1:D2)', 'A1:B2', 'C1:D2'] as RegExpMatchArray;
      const result = MMULT.calculate(matches, mockContext);
      expect(result).toEqual([
        [19, 22],  // 1*5 + 2*7 = 19, 1*6 + 2*8 = 22
        [43, 50]   // 3*5 + 4*7 = 43, 3*6 + 4*8 = 50
      ]);
    });

    it('should return single value for 1x1 result', () => {
      const matches = ['MMULT(A1:A1, A1:A1)', 'A1:A1', 'A1:A1'] as RegExpMatchArray;
      const result = MMULT.calculate(matches, mockContext);
      expect(result).toBe(1); // 1*1 = 1
    });

    it('should return VALUE error for dimension mismatch', () => {
      const matches = ['MMULT(A1:B2, E1:F3)', 'A1:B2', 'E1:F3'] as RegExpMatchArray;
      const result = MMULT.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });

    it('should return VALUE error for non-numeric values', () => {
      const matches = ['MMULT(H1:I2, A1:B2)', 'H1:I2', 'A1:B2'] as RegExpMatchArray;
      const result = MMULT.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });
  });

  describe('MUNIT', () => {
    it('should create 3x3 identity matrix', () => {
      const matches = ['MUNIT(3)', '3'] as RegExpMatchArray;
      const result = MUNIT.calculate(matches, mockContext);
      expect(result).toEqual([
        [1, 0, 0],
        [0, 1, 0],
        [0, 0, 1]
      ]);
    });

    it('should return single value 1 for size 1', () => {
      const matches = ['MUNIT(1)', '1'] as RegExpMatchArray;
      const result = MUNIT.calculate(matches, mockContext);
      expect(result).toBe(1);
    });

    it('should return VALUE error for non-integer size', () => {
      const matches = ['MUNIT(2.5)', '2.5'] as RegExpMatchArray;
      const result = MUNIT.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });

    it('should return VALUE error for negative size', () => {
      const matches = ['MUNIT(-1)', '-1'] as RegExpMatchArray;
      const result = MUNIT.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });

    it('should return VALUE error for zero size', () => {
      const matches = ['MUNIT(0)', '0'] as RegExpMatchArray;
      const result = MUNIT.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });

    it('should return NUM error for size > 1000', () => {
      const matches = ['MUNIT(1001)', '1001'] as RegExpMatchArray;
      const result = MUNIT.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });

    it('should return VALUE error for non-numeric input', () => {
      const matches = ['MUNIT(text)', 'text'] as RegExpMatchArray;
      const result = MUNIT.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });
  });
});