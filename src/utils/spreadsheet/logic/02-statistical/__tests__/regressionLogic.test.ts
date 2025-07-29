import { describe, it, expect } from 'vitest';
import type { FormulaContext } from '../../shared/types';
import {
  SLOPE, INTERCEPT, RSQ, STEYX,
  FORECAST, LINEST
} from '../regressionLogic';
import { FormulaError } from '../../shared/types';

// Helper function to create FormulaContext
function createContext(data: any[][]): FormulaContext {
  return {
    data,
    row: 0,
    col: 0
  };
}

describe('Regression Analysis Functions', () => {
  const mockContext = createContext([
    // Perfect linear relationship: y = 2x + 1
    [1, 3, 2.8, 1, 'text', 10],     // Row 1 (A1-F1)
    [2, 5, 5.2, 4, '', 20],         // Row 2 (A2-F2)
    [3, 7, 6.9, 9, null, null],     // Row 3 (A3-F3)
    [4, 9, 9.1, 16, null, null],    // Row 4 (A4-F4)
    [5, 11, 10.8, 25, null, null]   // Row 5 (A5-F5)
  ]);

  describe('SLOPE', () => {
    it('should calculate slope of perfect linear relationship', () => {
      const matches = ['SLOPE(B1:B5, A1:A5)', 'B1:B5', 'A1:A5'] as RegExpMatchArray;
      const result = SLOPE.calculate(matches, mockContext);
      expect(result).toBe(2);
    });

    it('should calculate slope of non-perfect linear relationship', () => {
      const matches = ['SLOPE(C1:C5, A1:A5)', 'C1:C5', 'A1:A5'] as RegExpMatchArray;
      const result = SLOPE.calculate(matches, mockContext);
      expect(result).toBeCloseTo(2, 1);
    });

    it('should return NA error for different array lengths', () => {
      const matches = ['SLOPE(B1:B3, A1:A5)', 'B1:B3', 'A1:A5'] as RegExpMatchArray;
      expect(SLOPE.calculate(matches, mockContext)).toBe(FormulaError.NA);
    });

    it('should return DIV0 error for zero variance in x', () => {
      const mockContextSameX = {
        cells: {
          A1: { value: 5 },
          A2: { value: 5 },
          A3: { value: 5 },
          B1: { value: 1 },
          B2: { value: 2 },
          B3: { value: 3 },
        }
      };
      const matches = ['SLOPE(B1:B3, A1:A3)', 'B1:B3', 'A1:A3'] as RegExpMatchArray;
      expect(SLOPE.calculate(matches, mockContextSameX)).toBe(FormulaError.DIV0);
    });

    it('should ignore non-numeric values', () => {
      const matches = ['SLOPE(B1:E3, A1:D3)', 'B1:E3', 'A1:D3'] as RegExpMatchArray;
      const result = SLOPE.calculate(matches, mockContext);
      expect(result).toBeCloseTo(2, 1);
    });

    it('should return DIV0 error for less than 2 data points', () => {
      const matches = ['SLOPE(B1, A1)', 'B1', 'A1'] as RegExpMatchArray;
      expect(SLOPE.calculate(matches, mockContext)).toBe(FormulaError.DIV0);
    });
  });

  describe('INTERCEPT', () => {
    it('should calculate intercept of perfect linear relationship', () => {
      const matches = ['INTERCEPT(B1:B5, A1:A5)', 'B1:B5', 'A1:A5'] as RegExpMatchArray;
      const result = INTERCEPT.calculate(matches, mockContext);
      expect(result).toBe(1);
    });

    it('should calculate intercept of non-perfect linear relationship', () => {
      const matches = ['INTERCEPT(C1:C5, A1:A5)', 'C1:C5', 'A1:A5'] as RegExpMatchArray;
      const result = INTERCEPT.calculate(matches, mockContext);
      expect(result).toBeCloseTo(1, 1);
    });

    it('should return NA error for different array lengths', () => {
      const matches = ['INTERCEPT(B1:B3, A1:A5)', 'B1:B3', 'A1:A5'] as RegExpMatchArray;
      expect(INTERCEPT.calculate(matches, mockContext)).toBe(FormulaError.NA);
    });

    it('should work with slope to reconstruct line', () => {
      const slopeMatches = ['SLOPE(B1:B5, A1:A5)', 'B1:B5', 'A1:A5'] as RegExpMatchArray;
      const slope = SLOPE.calculate(slopeMatches, mockContext);
      
      const interceptMatches = ['INTERCEPT(B1:B5, A1:A5)', 'B1:B5', 'A1:A5'] as RegExpMatchArray;
      const intercept = INTERCEPT.calculate(interceptMatches, mockContext);
      
      // Verify: y = slope * x + intercept
      expect(slope * 3 + intercept).toBe(7); // Point (3, 7)
    });
  });

  describe('RSQ', () => {
    it('should calculate R-squared for perfect correlation', () => {
      const matches = ['RSQ(B1:B5, A1:A5)', 'B1:B5', 'A1:A5'] as RegExpMatchArray;
      const result = RSQ.calculate(matches, mockContext);
      expect(result).toBe(1);
    });

    it('should calculate R-squared for imperfect correlation', () => {
      const matches = ['RSQ(C1:C5, A1:A5)', 'C1:C5', 'A1:A5'] as RegExpMatchArray;
      const result = RSQ.calculate(matches, mockContext);
      expect(result).toBeGreaterThan(0.95);
      expect(result).toBeLessThan(1);
    });

    it('should calculate lower R-squared for non-linear data', () => {
      const matches = ['RSQ(D1:D5, A1:A5)', 'D1:D5', 'A1:A5'] as RegExpMatchArray;
      const result = RSQ.calculate(matches, mockContext);
      expect(result).toBeGreaterThan(0);
      expect(result).toBeLessThan(0.95);
    });

    it('should return NA error for different array lengths', () => {
      const matches = ['RSQ(B1:B3, A1:A5)', 'B1:B3', 'A1:A5'] as RegExpMatchArray;
      expect(RSQ.calculate(matches, mockContext)).toBe(FormulaError.NA);
    });

    it('should return DIV0 error for constant y values', () => {
      const mockContextConstY = {
        cells: {
          A1: { value: 1 },
          A2: { value: 2 },
          A3: { value: 3 },
          B1: { value: 5 },
          B2: { value: 5 },
          B3: { value: 5 },
        }
      };
      const matches = ['RSQ(B1:B3, A1:A3)', 'B1:B3', 'A1:A3'] as RegExpMatchArray;
      expect(RSQ.calculate(matches, mockContextConstY)).toBe(FormulaError.DIV0);
    });
  });

  describe('STEYX', () => {
    it('should return 0 for perfect linear relationship', () => {
      const matches = ['STEYX(B1:B5, A1:A5)', 'B1:B5', 'A1:A5'] as RegExpMatchArray;
      const result = STEYX.calculate(matches, mockContext);
      expect(result).toBe(0);
    });

    it('should calculate standard error for imperfect data', () => {
      const matches = ['STEYX(C1:C5, A1:A5)', 'C1:C5', 'A1:A5'] as RegExpMatchArray;
      const result = STEYX.calculate(matches, mockContext);
      expect(result).toBeGreaterThan(0);
      expect(result).toBeLessThan(1);
    });

    it('should return larger error for non-linear data', () => {
      const matches = ['STEYX(D1:D5, A1:A5)', 'D1:D5', 'A1:A5'] as RegExpMatchArray;
      const result = STEYX.calculate(matches, mockContext);
      expect(result).toBeGreaterThan(1);
    });

    it('should return NA error for different array lengths', () => {
      const matches = ['STEYX(B1:B3, A1:A5)', 'B1:B3', 'A1:A5'] as RegExpMatchArray;
      expect(STEYX.calculate(matches, mockContext)).toBe(FormulaError.NA);
    });

    it('should return DIV0 error for less than 3 data points', () => {
      const matches = ['STEYX(F1:F2, A1:A2)', 'F1:F2', 'A1:A2'] as RegExpMatchArray;
      expect(STEYX.calculate(matches, mockContext)).toBe(FormulaError.DIV0);
    });
  });

  describe('FORECAST', () => {
    it('should predict value on perfect linear trend', () => {
      const matches = ['FORECAST(6, B1:B5, A1:A5)', '6', 'B1:B5', 'A1:A5'] as RegExpMatchArray;
      const result = FORECAST.calculate(matches, mockContext);
      expect(result).toBe(13); // 2*6 + 1
    });

    it('should extrapolate backwards', () => {
      const matches = ['FORECAST(0, B1:B5, A1:A5)', '0', 'B1:B5', 'A1:A5'] as RegExpMatchArray;
      const result = FORECAST.calculate(matches, mockContext);
      expect(result).toBe(1); // 2*0 + 1
    });

    it('should interpolate between points', () => {
      const matches = ['FORECAST(2.5, B1:B5, A1:A5)', '2.5', 'B1:B5', 'A1:A5'] as RegExpMatchArray;
      const result = FORECAST.calculate(matches, mockContext);
      expect(result).toBe(6); // 2*2.5 + 1
    });

    it('should work with imperfect data', () => {
      const matches = ['FORECAST(6, C1:C5, A1:A5)', '6', 'C1:C5', 'A1:A5'] as RegExpMatchArray;
      const result = FORECAST.calculate(matches, mockContext);
      expect(result).toBeCloseTo(13, 1);
    });

    it('should return NA error for different array lengths', () => {
      const matches = ['FORECAST(6, B1:B3, A1:A5)', '6', 'B1:B3', 'A1:A5'] as RegExpMatchArray;
      expect(FORECAST.calculate(matches, mockContext)).toBe(FormulaError.NA);
    });

    it('should return VALUE error for non-numeric x', () => {
      const matches = ['FORECAST(text, B1:B5, A1:A5)', 'text', 'B1:B5', 'A1:A5'] as RegExpMatchArray;
      expect(FORECAST.calculate(matches, mockContext)).toBe(FormulaError.VALUE);
    });
  });

  describe('LINEST', () => {
    it('should return array of regression statistics', () => {
      const matches = ['LINEST(B1:B5, A1:A5, TRUE, TRUE)', 'B1:B5', 'A1:A5', 'TRUE', 'TRUE'] as RegExpMatchArray;
      const result = LINEST.calculate(matches, mockContext);
      
      expect(Array.isArray(result)).toBe(true);
      expect(result).toHaveLength(2);
      expect(result[0]).toHaveLength(2);
      expect(result[1]).toHaveLength(2);
      
      // First row: [slope, intercept]
      expect(result[0][0]).toBe(2);     // slope
      expect(result[0][1]).toBe(1);     // intercept
      
      // Second row should contain standard errors (0 for perfect fit)
      expect(result[1][0]).toBe(0);     // se of slope
      expect(result[1][1]).toBe(0);     // se of intercept
    });

    it('should return only coefficients when stats=FALSE', () => {
      const matches = ['LINEST(B1:B5, A1:A5, TRUE, FALSE)', 'B1:B5', 'A1:A5', 'TRUE', 'FALSE'] as RegExpMatchArray;
      const result = LINEST.calculate(matches, mockContext);
      
      expect(result).toEqual([2, 1]); // [slope, intercept]
    });

    it('should force intercept to 0 when const=FALSE', () => {
      const matches = ['LINEST(B1:B5, A1:A5, FALSE, FALSE)', 'B1:B5', 'A1:A5', 'FALSE', 'FALSE'] as RegExpMatchArray;
      const result = LINEST.calculate(matches, mockContext);
      
      expect(Array.isArray(result)).toBe(true);
      expect(result[0]).toBeCloseTo(2.2, 1); // slope will be different without intercept
      expect(result[1]).toBe(0); // no intercept
    });

    it('should handle imperfect data', () => {
      const matches = ['LINEST(C1:C5, A1:A5, TRUE, TRUE)', 'C1:C5', 'A1:A5', 'TRUE', 'TRUE'] as RegExpMatchArray;
      const result = LINEST.calculate(matches, mockContext);
      
      expect(result[0][0]).toBeCloseTo(2, 1);     // slope ≈ 2
      expect(result[0][1]).toBeCloseTo(1, 1);     // intercept ≈ 1
      expect(result[1][0]).toBeGreaterThan(0);    // non-zero standard error
      expect(result[1][1]).toBeGreaterThan(0);    // non-zero standard error
    });

    it('should return NA error for different array lengths', () => {
      const matches = ['LINEST(B1:B3, A1:A5, TRUE, TRUE)', 'B1:B3', 'A1:A5', 'TRUE', 'TRUE'] as RegExpMatchArray;
      expect(LINEST.calculate(matches, mockContext)).toBe(FormulaError.NA);
    });

    it('should default to const=TRUE and stats=FALSE', () => {
      const matches = ['LINEST(B1:B5, A1:A5)', 'B1:B5', 'A1:A5'] as RegExpMatchArray;
      const result = LINEST.calculate(matches, mockContext);
      
      expect(result).toEqual([2, 1]); // Same as const=TRUE, stats=FALSE
    });

    it('should handle boolean string values', () => {
      const matches = ['LINEST(B1:B5, A1:A5, "TRUE", "FALSE")', 'B1:B5', 'A1:A5', '"TRUE"', '"FALSE"'] as RegExpMatchArray;
      const result = LINEST.calculate(matches, mockContext);
      
      expect(result).toEqual([2, 1]);
    });
  });
});