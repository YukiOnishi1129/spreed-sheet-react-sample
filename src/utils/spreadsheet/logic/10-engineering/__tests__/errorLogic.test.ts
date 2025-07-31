import { describe, it, expect } from 'vitest';
import {
  ERF, ERFC, ERF_PRECISE, ERFC_PRECISE
} from '../errorLogic';
import { FormulaError } from '../../shared/types';
import type { FormulaContext } from '../../shared/types';

// Helper function to create FormulaContext
const createContext = (data: (string | number | boolean | null)[][]): FormulaContext => ({
  data: data.map(row => row.map(cell => ({ value: cell }))),
  row: 0,
  col: 0
});

describe('Error Functions', () => {
  const mockContext = createContext([
    [0, 0.5, 1, 1.5, 2, 2.5, 3], // x values
    [-3, -2, -1, -0.5, 0], // negative x values
    [0.1, 0.01, 0.001], // small x values
    [10, 20, 30], // large x values
  ]);

  describe('ERF Function (Error Function)', () => {
    it('should calculate erf(0) = 0', () => {
      const matches = ['ERF(0)', '0'] as RegExpMatchArray;
      const result = ERF.calculate(matches, mockContext);
      expect(result).toBe(0);
    });

    it('should calculate erf for positive values', () => {
      const matches = ['ERF(1)', '1'] as RegExpMatchArray;
      const result = ERF.calculate(matches, mockContext);
      expect(result).toBeCloseTo(0.8427, 4); // erf(1) ≈ 0.8427
    });

    it('should calculate erf for negative values', () => {
      const matches = ['ERF(-1)', '-1'] as RegExpMatchArray;
      const result = ERF.calculate(matches, mockContext);
      expect(result).toBeCloseTo(-0.8427, 4); // erf is odd function
    });

    it('should approach 1 for large positive values', () => {
      const matches = ['ERF(3)', '3'] as RegExpMatchArray;
      const result = ERF.calculate(matches, mockContext);
      expect(result).toBeGreaterThan(0.999);
      expect(result).toBeLessThanOrEqual(1);
    });

    it('should approach -1 for large negative values', () => {
      const matches = ['ERF(-3)', '-3'] as RegExpMatchArray;
      const result = ERF.calculate(matches, mockContext);
      expect(result).toBeLessThan(-0.999);
      expect(result).toBeGreaterThanOrEqual(-1);
    });

    it('should calculate erf between bounds', () => {
      const matches = ['ERF(0.5, 1)', '0.5', '1'] as RegExpMatchArray;
      const result = ERF.calculate(matches, mockContext);
      // erf(1) - erf(0.5) ≈ 0.8427 - 0.5205 = 0.3222
      expect(result).toBeCloseTo(0.3222, 3);
    });

    it('should handle swapped bounds', () => {
      const matches = ['ERF(1, 0.5)', '1', '0.5'] as RegExpMatchArray;
      const result = ERF.calculate(matches, mockContext);
      expect(result).toBeCloseTo(-0.3222, 3); // Negative of above
    });

    it('should handle small values accurately', () => {
      const matches = ['ERF(0.1)', '0.1'] as RegExpMatchArray;
      const result = ERF.calculate(matches, mockContext);
      // For small x, erf(x) ≈ 2x/√π
      expect(result).toBeCloseTo(0.1125, 4);
    });

    it('should return VALUE error for non-numeric input', () => {
      const matches = ['ERF("text")', '"text"'] as RegExpMatchArray;
      const result = ERF.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });

    it('should handle cell references', () => {
      const matches = ['ERF(B1)', 'B1'] as RegExpMatchArray;
      const result = ERF.calculate(matches, mockContext);
      expect(result).toBeCloseTo(0.5205, 4); // erf(0.5)
    });
  });

  describe('ERFC Function (Complementary Error Function)', () => {
    it('should calculate erfc(0) = 1', () => {
      const matches = ['ERFC(0)', '0'] as RegExpMatchArray;
      const result = ERFC.calculate(matches, mockContext);
      expect(result).toBe(1);
    });

    it('should calculate erfc for positive values', () => {
      const matches = ['ERFC(1)', '1'] as RegExpMatchArray;
      const result = ERFC.calculate(matches, mockContext);
      expect(result).toBeCloseTo(0.1573, 4); // erfc(1) = 1 - erf(1)
    });

    it('should calculate erfc for negative values', () => {
      const matches = ['ERFC(-1)', '-1'] as RegExpMatchArray;
      const result = ERFC.calculate(matches, mockContext);
      expect(result).toBeCloseTo(1.8427, 4); // erfc(-1) = 1 - erf(-1)
    });

    it('should approach 0 for large positive values', () => {
      const matches = ['ERFC(5)', '5'] as RegExpMatchArray;
      const result = ERFC.calculate(matches, mockContext);
      expect(result).toBeGreaterThan(0);
      expect(result).toBeLessThan(0.0001);
    });

    it('should approach 2 for large negative values', () => {
      const matches = ['ERFC(-3)', '-3'] as RegExpMatchArray;
      const result = ERFC.calculate(matches, mockContext);
      expect(result).toBeGreaterThan(1.999);
      expect(result).toBeLessThanOrEqual(2);
    });

    it('should verify erfc(x) = 1 - erf(x)', () => {
      const erfMatches = ['ERF(0.7)', '0.7'] as RegExpMatchArray;
      const erfResult = ERF.calculate(erfMatches, mockContext) as number;
      
      const erfcMatches = ['ERFC(0.7)', '0.7'] as RegExpMatchArray;
      const erfcResult = ERFC.calculate(erfcMatches, mockContext) as number;
      
      expect(erfResult + erfcResult).toBeCloseTo(1, 10);
    });

    it('should handle very large values without overflow', () => {
      const matches = ['ERFC(10)', '10'] as RegExpMatchArray;
      const result = ERFC.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeGreaterThan(0);
      expect(result).toBeLessThan(1e-10);
    });

    it('should return VALUE error for non-numeric input', () => {
      const matches = ['ERFC("text")', '"text"'] as RegExpMatchArray;
      const result = ERFC.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });
  });

  describe('ERF.PRECISE Function (Precise Error Function)', () => {
    it('should calculate erf.precise(0) = 0', () => {
      const matches = ['ERF.PRECISE(0)', '0'] as RegExpMatchArray;
      const result = ERF_PRECISE.calculate(matches, mockContext);
      expect(result).toBe(0);
    });

    it('should match ERF for single argument', () => {
      const erfMatches = ['ERF(1.5)', '1.5'] as RegExpMatchArray;
      const erfResult = ERF.calculate(erfMatches, mockContext);
      
      const preciseMatches = ['ERF.PRECISE(1.5)', '1.5'] as RegExpMatchArray;
      const preciseResult = ERF_PRECISE.calculate(preciseMatches, mockContext);
      
      expect(preciseResult).toBe(erfResult);
    });

    it('should calculate for positive values', () => {
      const matches = ['ERF.PRECISE(2)', '2'] as RegExpMatchArray;
      const result = ERF_PRECISE.calculate(matches, mockContext);
      expect(result).toBeCloseTo(0.9953, 4);
    });

    it('should handle negative values', () => {
      const matches = ['ERF.PRECISE(-0.5)', '-0.5'] as RegExpMatchArray;
      const result = ERF_PRECISE.calculate(matches, mockContext);
      expect(result).toBeCloseTo(-0.5205, 4);
    });

    it('should handle very small values', () => {
      const matches = ['ERF.PRECISE(0.001)', '0.001'] as RegExpMatchArray;
      const result = ERF_PRECISE.calculate(matches, mockContext);
      // For small x, erf(x) ≈ 2x/√π
      const expected = 2 * 0.001 / Math.sqrt(Math.PI);
      expect(result).toBeCloseTo(expected, 6);
    });

    it('should return VALUE error for non-numeric input', () => {
      const matches = ['ERF.PRECISE("invalid")', '"invalid"'] as RegExpMatchArray;
      const result = ERF_PRECISE.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });
  });

  describe('ERFC.PRECISE Function (Precise Complementary Error Function)', () => {
    it('should calculate erfc.precise(0) = 1', () => {
      const matches = ['ERFC.PRECISE(0)', '0'] as RegExpMatchArray;
      const result = ERFC_PRECISE.calculate(matches, mockContext);
      expect(result).toBe(1);
    });

    it('should match ERFC for positive values', () => {
      const erfcMatches = ['ERFC(0.8)', '0.8'] as RegExpMatchArray;
      const erfcResult = ERFC.calculate(erfcMatches, mockContext);
      
      const preciseMatches = ['ERFC.PRECISE(0.8)', '0.8'] as RegExpMatchArray;
      const preciseResult = ERFC_PRECISE.calculate(preciseMatches, mockContext);
      
      expect(preciseResult).toBe(erfcResult);
    });

    it('should calculate for large positive values', () => {
      const matches = ['ERFC.PRECISE(4)', '4'] as RegExpMatchArray;
      const result = ERFC_PRECISE.calculate(matches, mockContext);
      expect(result).toBeGreaterThan(0);
      expect(result).toBeLessThan(0.00001);
    });

    it('should handle negative values', () => {
      const matches = ['ERFC.PRECISE(-2)', '-2'] as RegExpMatchArray;
      const result = ERFC_PRECISE.calculate(matches, mockContext);
      expect(result).toBeCloseTo(1.9953, 4); // 1 + erf(2)
    });

    it('should verify erfc.precise(x) = 1 - erf.precise(x)', () => {
      const erfMatches = ['ERF.PRECISE(1.2)', '1.2'] as RegExpMatchArray;
      const erfResult = ERF_PRECISE.calculate(erfMatches, mockContext) as number;
      
      const erfcMatches = ['ERFC.PRECISE(1.2)', '1.2'] as RegExpMatchArray;
      const erfcResult = ERFC_PRECISE.calculate(erfcMatches, mockContext) as number;
      
      expect(erfResult + erfcResult).toBeCloseTo(1, 10);
    });

    it('should return VALUE error for non-numeric input', () => {
      const matches = ['ERFC.PRECISE("invalid")', '"invalid"'] as RegExpMatchArray;
      const result = ERFC_PRECISE.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });
  });

  describe('Edge Cases and Integration', () => {
    it('should handle extreme precision for small values', () => {
      const x = 1e-10;
      const matches = ['ERF(' + x + ')', x.toString()] as RegExpMatchArray;
      const result = ERF.calculate(matches, mockContext) as number;
      
      // For very small x, erf(x) ≈ 2x/√π
      const expected = 2 * x / Math.sqrt(Math.PI);
      expect(result).toBeCloseTo(expected, 15);
    });

    it('should verify symmetry properties', () => {
      // erf(-x) = -erf(x)
      const posMatches = ['ERF(1.5)', '1.5'] as RegExpMatchArray;
      const posResult = ERF.calculate(posMatches, mockContext) as number;
      
      const negMatches = ['ERF(-1.5)', '-1.5'] as RegExpMatchArray;
      const negResult = ERF.calculate(negMatches, mockContext) as number;
      
      expect(negResult).toBeCloseTo(-posResult, 10);
    });

    it('should verify boundary integral property', () => {
      // erf(∞) - erf(-∞) = 2
      const largeMatches = ['ERF(10)', '10'] as RegExpMatchArray;
      const largeResult = ERF.calculate(largeMatches, mockContext) as number;
      
      const smallMatches = ['ERF(-10)', '-10'] as RegExpMatchArray;
      const smallResult = ERF.calculate(smallMatches, mockContext) as number;
      
      expect(largeResult - smallResult).toBeCloseTo(2, 6);
    });

    it('should handle the Gaussian integral relationship', () => {
      // ∫[0,x] e^(-t²)dt = √π/2 * erf(x)
      const matches = ['ERF(1)', '1'] as RegExpMatchArray;
      const erfResult = ERF.calculate(matches, mockContext) as number;
      
      // The integral value
      const integralValue = Math.sqrt(Math.PI) / 2 * erfResult;
      expect(integralValue).toBeCloseTo(0.7468, 3);
    });

    it('should handle cell references', () => {
      const matches = ['ERFC(A2)', 'A2'] as RegExpMatchArray;
      const result = ERFC.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeCloseTo(2, 1); // erfc(-3) ≈ 2
    });

    it('should compare regular and precise versions', () => {
      const testValues = [0.1, 0.5, 1, 2, 3];
      
      for (const x of testValues) {
        const erfMatches = ['ERF(' + x + ')', x.toString()] as RegExpMatchArray;
        const erfResult = ERF.calculate(erfMatches, mockContext);
        
        const preciseMatches = ['ERF.PRECISE(' + x + ')', x.toString()] as RegExpMatchArray;
        const preciseResult = ERF_PRECISE.calculate(preciseMatches, mockContext);
        
        expect(erfResult).toBe(preciseResult);
      }
    });
  });
});