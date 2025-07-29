import { describe, it, expect } from 'vitest';
import type { FormulaContext } from '../../shared/types';
import {
  FISHER, FISHERINV,
  PHI, GAUSS,
  PROB, BINOM_INV
} from '../otherStatisticalLogic';
import { FormulaError } from '../../shared/types';

// Helper function to create FormulaContext
function createContext(data: (string | number | boolean | null)[][]): FormulaContext {
  return {
    data: data.map(row => row.map(cell => ({ value: cell }))),
    row: 0,
    col: 0
  };
}

describe('Other Statistical Functions', () => {
  const mockContext = createContext([
    [0, 1, 0.1, 'text'],      // Row 1 (A1-D1)
    [0.5, 2, 0.2, -2],        // Row 2 (A2-D2)
    [0.75, 3, 0.3, 1.5],      // Row 3 (A3-D3)
    [0.9, 4, 0.2, null],      // Row 4 (A4-D4)
    [0.99, 5, 0.2, null]      // Row 5 (A5-D5)
  ]);

  describe('Fisher Transformation Functions', () => {
    describe('FISHER', () => {
      it('should calculate Fisher transformation', () => {
        const matches = ['FISHER(0.5)', '0.5'] as RegExpMatchArray;
        const result = FISHER.calculate(matches, mockContext);
        expect(result).toBeCloseTo(0.5493, 4);
      });

      it('should handle zero', () => {
        const matches = ['FISHER(0)', '0'] as RegExpMatchArray;
        const result = FISHER.calculate(matches, mockContext);
        expect(result).toBe(0);
      });

      it('should handle negative values', () => {
        const matches = ['FISHER(-0.5)', '-0.5'] as RegExpMatchArray;
        const result = FISHER.calculate(matches, mockContext);
        expect(result).toBeCloseTo(-0.5493, 4);
      });

      it('should return NUM error for values outside (-1, 1)', () => {
        const matches1 = ['FISHER(1)', '1'] as RegExpMatchArray;
        expect(FISHER.calculate(matches1, mockContext)).toBe(FormulaError.NUM);
        
        const matches2 = ['FISHER(-1)', '-1'] as RegExpMatchArray;
        expect(FISHER.calculate(matches2, mockContext)).toBe(FormulaError.NUM);
        
        const matches3 = ['FISHER(1.5)', '1.5'] as RegExpMatchArray;
        expect(FISHER.calculate(matches3, mockContext)).toBe(FormulaError.NUM);
      });

      it('should return VALUE error for non-numeric input', () => {
        const matches = ['FISHER(text)', 'text'] as RegExpMatchArray;
        expect(FISHER.calculate(matches, mockContext)).toBe(FormulaError.VALUE);
      });

      it('should work with cell references', () => {
        const matches = ['FISHER(A2)', 'A2'] as RegExpMatchArray;
        const result = FISHER.calculate(matches, mockContext);
        expect(result).toBeCloseTo(0.5493, 4);
      });
    });

    describe('FISHERINV', () => {
      it('should calculate inverse Fisher transformation', () => {
        const matches = ['FISHERINV(0.5493)', '0.5493'] as RegExpMatchArray;
        const result = FISHERINV.calculate(matches, mockContext);
        expect(result).toBeCloseTo(0.5, 4);
      });

      it('should handle zero', () => {
        const matches = ['FISHERINV(0)', '0'] as RegExpMatchArray;
        const result = FISHERINV.calculate(matches, mockContext);
        expect(result).toBe(0);
      });

      it('should handle negative values', () => {
        const matches = ['FISHERINV(-0.5493)', '-0.5493'] as RegExpMatchArray;
        const result = FISHERINV.calculate(matches, mockContext);
        expect(result).toBeCloseTo(-0.5, 4);
      });

      it('should handle large values', () => {
        const matches = ['FISHERINV(2)', '2'] as RegExpMatchArray;
        const result = FISHERINV.calculate(matches, mockContext);
        expect(result).toBeCloseTo(0.9640, 4);
      });

      it('should return VALUE error for non-numeric input', () => {
        const matches = ['FISHERINV(text)', 'text'] as RegExpMatchArray;
        expect(FISHERINV.calculate(matches, mockContext)).toBe(FormulaError.VALUE);
      });

      it('should be inverse of FISHER', () => {
        const value = 0.75;
        const matches1 = ['FISHER(0.75)', '0.75'] as RegExpMatchArray;
        const fisherResult = FISHER.calculate(matches1, mockContext);
        
        const matches2 = [`FISHERINV(${fisherResult})`, `${fisherResult}`] as RegExpMatchArray;
        const invResult = FISHERINV.calculate(matches2, mockContext);
        
        expect(invResult).toBeCloseTo(value, 10);
      });
    });
  });

  describe('Normal Distribution Functions', () => {
    describe('PHI', () => {
      it('should calculate standard normal density', () => {
        const matches = ['PHI(0)', '0'] as RegExpMatchArray;
        const result = PHI.calculate(matches, mockContext);
        expect(result).toBeCloseTo(0.3989, 4);
      });

      it('should handle positive values', () => {
        const matches = ['PHI(1)', '1'] as RegExpMatchArray;
        const result = PHI.calculate(matches, mockContext);
        expect(result).toBeCloseTo(0.2420, 4);
      });

      it('should handle negative values', () => {
        const matches = ['PHI(-1)', '-1'] as RegExpMatchArray;
        const result = PHI.calculate(matches, mockContext);
        expect(result).toBeCloseTo(0.2420, 4); // symmetric
      });

      it('should return VALUE error for non-numeric input', () => {
        const matches = ['PHI(text)', 'text'] as RegExpMatchArray;
        expect(PHI.calculate(matches, mockContext)).toBe(FormulaError.VALUE);
      });
    });

    describe('GAUSS', () => {
      it('should calculate Gauss error function', () => {
        const matches = ['GAUSS(0)', '0'] as RegExpMatchArray;
        const result = GAUSS.calculate(matches, mockContext);
        expect(result).toBe(0); // NORM.S.DIST(0) - 0.5 = 0
      });

      it('should handle positive values', () => {
        const matches = ['GAUSS(1)', '1'] as RegExpMatchArray;
        const result = GAUSS.calculate(matches, mockContext);
        expect(result).toBeCloseTo(0.3413, 4);
      });

      it('should handle negative values', () => {
        const matches = ['GAUSS(-1)', '-1'] as RegExpMatchArray;
        const result = GAUSS.calculate(matches, mockContext);
        expect(result).toBeCloseTo(-0.3413, 4);
      });

      it('should approach 0.5 for large positive values', () => {
        const matches = ['GAUSS(5)', '5'] as RegExpMatchArray;
        const result = GAUSS.calculate(matches, mockContext);
        expect(result).toBeCloseTo(0.5, 4);
      });

      it('should approach -0.5 for large negative values', () => {
        const matches = ['GAUSS(-5)', '-5'] as RegExpMatchArray;
        const result = GAUSS.calculate(matches, mockContext);
        expect(result).toBeCloseTo(-0.5, 4);
      });

      it('should return VALUE error for non-numeric input', () => {
        const matches = ['GAUSS(text)', 'text'] as RegExpMatchArray;
        expect(GAUSS.calculate(matches, mockContext)).toBe(FormulaError.VALUE);
      });
    });
  });

  describe('Probability Functions', () => {
    describe('PROB', () => {
      it('should calculate probability for exact range', () => {
        const matches = ['PROB(B1:B5, C1:C5, 2, 4)', 'B1:B5', 'C1:C5', '2', '4'] as RegExpMatchArray;
        const result = PROB.calculate(matches, mockContext);
        expect(result).toBe(0.7); // 0.2 + 0.3 + 0.2
      });

      it('should handle single value range', () => {
        const matches = ['PROB(B1:B5, C1:C5, 3, 3)', 'B1:B5', 'C1:C5', '3', '3'] as RegExpMatchArray;
        const result = PROB.calculate(matches, mockContext);
        expect(result).toBe(0.3);
      });

      it('should handle lower limit only', () => {
        const matches = ['PROB(B1:B5, C1:C5, 4)', 'B1:B5', 'C1:C5', '4'] as RegExpMatchArray;
        const result = PROB.calculate(matches, mockContext);
        expect(result).toBe(0.4); // 0.2 + 0.2
      });

      it('should return 0 for out of range values', () => {
        const matches = ['PROB(B1:B5, C1:C5, 6, 7)', 'B1:B5', 'C1:C5', '6', '7'] as RegExpMatchArray;
        const result = PROB.calculate(matches, mockContext);
        expect(result).toBe(0);
      });

      it('should return NA error for different array lengths', () => {
        const matches = ['PROB(B1:B3, C1:C5, 2, 4)', 'B1:B3', 'C1:C5', '2', '4'] as RegExpMatchArray;
        expect(PROB.calculate(matches, mockContext)).toBe(FormulaError.NA);
      });

      it('should return NUM error for probabilities not summing to 1', () => {
        const mockContextBadProb = createContext([
          [0, 1, 0.1, 'text'],      // Row 1 (A1-D1)
          [0.5, 2, 0.1, -2],        // Row 2 (A2-D2)
          [0.75, 3, 0.1, 1.5],      // Row 3 (A3-D3)
          [0.9, 4, 0.1, null],      // Row 4 (A4-D4)
          [0.99, 5, 0.1, null]      // Row 5 (A5-D5) - sum = 0.5
        ]);
        const matches = ['PROB(B1:B5, C1:C5, 2, 4)', 'B1:B5', 'C1:C5', '2', '4'] as RegExpMatchArray;
        expect(PROB.calculate(matches, mockContextBadProb)).toBe(FormulaError.NUM);
      });

      it('should return NUM error for invalid probabilities', () => {
        const mockContextNegProb = createContext([
          [0, 1, -0.1, 'text'],     // Row 1 (A1-D1) - negative probability
          [0.5, 2, 0.2, -2],        // Row 2 (A2-D2)
          [0.75, 3, 0.3, 1.5],      // Row 3 (A3-D3)
          [0.9, 4, 0.2, null],      // Row 4 (A4-D4)
          [0.99, 5, 0.2, null]      // Row 5 (A5-D5)
        ]);
        const matches = ['PROB(B1:B5, C1:C5, 2, 4)', 'B1:B5', 'C1:C5', '2', '4'] as RegExpMatchArray;
        expect(PROB.calculate(matches, mockContextNegProb)).toBe(FormulaError.NUM);
      });
    });

    describe('BINOM.INV', () => {
      it('should find smallest value where cumulative binomial >= alpha', () => {
        const matches = ['BINOM.INV(10, 0.5, 0.5)', '10', '0.5', '0.5'] as RegExpMatchArray;
        const result = BINOM_INV.calculate(matches, mockContext);
        expect(result).toBe(5);
      });

      it('should return 0 for very small alpha', () => {
        const matches = ['BINOM.INV(10, 0.5, 0.001)', '10', '0.5', '0.001'] as RegExpMatchArray;
        const result = BINOM_INV.calculate(matches, mockContext);
        expect(result).toBe(0);
      });

      it('should return n for alpha = 1', () => {
        const matches = ['BINOM.INV(10, 0.5, 1)', '10', '0.5', '1'] as RegExpMatchArray;
        const result = BINOM_INV.calculate(matches, mockContext);
        expect(result).toBe(10);
      });

      it('should handle edge case probabilities', () => {
        const matches1 = ['BINOM.INV(10, 0, 0.5)', '10', '0', '0.5'] as RegExpMatchArray;
        const result1 = BINOM_INV.calculate(matches1, mockContext);
        expect(result1).toBe(0); // p=0 means always 0 successes
        
        const matches2 = ['BINOM.INV(10, 1, 0.5)', '10', '1', '0.5'] as RegExpMatchArray;
        const result2 = BINOM_INV.calculate(matches2, mockContext);
        expect(result2).toBe(10); // p=1 means always n successes
      });

      it('should return NUM error for invalid trials', () => {
        const matches1 = ['BINOM.INV(-1, 0.5, 0.5)', '-1', '0.5', '0.5'] as RegExpMatchArray;
        expect(BINOM_INV.calculate(matches1, mockContext)).toBe(FormulaError.NUM);
        
        const matches2 = ['BINOM.INV(10.5, 0.5, 0.5)', '10.5', '0.5', '0.5'] as RegExpMatchArray;
        expect(BINOM_INV.calculate(matches2, mockContext)).toBe(FormulaError.NUM);
      });

      it('should return NUM error for probability outside [0,1]', () => {
        const matches1 = ['BINOM.INV(10, -0.1, 0.5)', '10', '-0.1', '0.5'] as RegExpMatchArray;
        expect(BINOM_INV.calculate(matches1, mockContext)).toBe(FormulaError.NUM);
        
        const matches2 = ['BINOM.INV(10, 1.1, 0.5)', '10', '1.1', '0.5'] as RegExpMatchArray;
        expect(BINOM_INV.calculate(matches2, mockContext)).toBe(FormulaError.NUM);
      });

      it('should return NUM error for alpha outside (0,1)', () => {
        const matches1 = ['BINOM.INV(10, 0.5, 0)', '10', '0.5', '0'] as RegExpMatchArray;
        expect(BINOM_INV.calculate(matches1, mockContext)).toBe(FormulaError.NUM);
        
        const matches2 = ['BINOM.INV(10, 0.5, 1.1)', '10', '0.5', '1.1'] as RegExpMatchArray;
        expect(BINOM_INV.calculate(matches2, mockContext)).toBe(FormulaError.NUM);
      });

      it('should return VALUE error for non-numeric inputs', () => {
        const matches = ['BINOM.INV(text, 0.5, 0.5)', 'text', '0.5', '0.5'] as RegExpMatchArray;
        expect(BINOM_INV.calculate(matches, mockContext)).toBe(FormulaError.VALUE);
      });
    });
  });
});