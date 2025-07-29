import { describe, it, expect } from 'vitest';
import {
  BESSELI, BESSELJ, BESSELK, BESSELY
} from '../besselLogic';
import { FormulaError } from '../../shared/types';
import type { FormulaContext } from '../../shared/types';

// Helper function to create FormulaContext
const createContext = (data: (string | number | boolean | null)[][]): FormulaContext => ({
  data: data.map(row => row.map(cell => ({ value: cell }))),
  row: 0,
  col: 0
});

describe('Bessel Functions', () => {
  const mockContext = createContext([
    [0, 1, 2, 3, 4, 5], // n values (order)
    [0.5, 1.0, 1.5, 2.0, 2.5, 3.0], // x values
    [-1, -0.5, 0, 0.5, 1, 1.5], // more x values
    [10, 20, 30, 40, 50], // large x values
  ]);

  describe('BESSELI Function (Modified Bessel Function of First Kind)', () => {
    it('should calculate BESSELI for order 0', () => {
      const matches = ['BESSELI(1, 0)', '1', '0'] as RegExpMatchArray;
      const result = BESSELI.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeCloseTo(1.2661, 4); // Known value for I₀(1)
    });

    it('should calculate BESSELI for order 1', () => {
      const matches = ['BESSELI(1, 1)', '1', '1'] as RegExpMatchArray;
      const result = BESSELI.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeCloseTo(0.5652, 4); // Known value for I₁(1)
    });

    it('should handle x = 0', () => {
      const matches = ['BESSELI(0, 0)', '0', '0'] as RegExpMatchArray;
      const result = BESSELI.calculate(matches, mockContext);
      expect(result).toBe(1); // I₀(0) = 1
    });

    it('should handle x = 0 for n > 0', () => {
      const matches = ['BESSELI(0, 2)', '0', '2'] as RegExpMatchArray;
      const result = BESSELI.calculate(matches, mockContext);
      expect(result).toBe(0); // Iₙ(0) = 0 for n > 0
    });

    it('should increase with x for fixed n', () => {
      const smallMatches = ['BESSELI(1, 0)', '1', '0'] as RegExpMatchArray;
      const small = BESSELI.calculate(smallMatches, mockContext) as number;
      
      const largeMatches = ['BESSELI(2, 0)', '2', '0'] as RegExpMatchArray;
      const large = BESSELI.calculate(largeMatches, mockContext) as number;
      
      expect(large).toBeGreaterThan(small);
    });

    it('should return NUM error for negative n', () => {
      const matches = ['BESSELI(1, -1)', '1', '-1'] as RegExpMatchArray;
      const result = BESSELI.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });

    it('should return NUM error for non-integer n', () => {
      const matches = ['BESSELI(1, 1.5)', '1', '1.5'] as RegExpMatchArray;
      const result = BESSELI.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });

    it('should handle cell references', () => {
      const matches = ['BESSELI(B2, A1)', 'B2', 'A1'] as RegExpMatchArray;
      const result = BESSELI.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
    });
  });

  describe('BESSELJ Function (Bessel Function of First Kind)', () => {
    it('should calculate BESSELJ for order 0', () => {
      const matches = ['BESSELJ(1, 0)', '1', '0'] as RegExpMatchArray;
      const result = BESSELJ.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeCloseTo(0.7652, 4); // Known value for J₀(1)
    });

    it('should calculate BESSELJ for order 1', () => {
      const matches = ['BESSELJ(1, 1)', '1', '1'] as RegExpMatchArray;
      const result = BESSELJ.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeCloseTo(0.4401, 4); // Known value for J₁(1)
    });

    it('should handle x = 0', () => {
      const matches = ['BESSELJ(0, 0)', '0', '0'] as RegExpMatchArray;
      const result = BESSELJ.calculate(matches, mockContext);
      expect(result).toBe(1); // J₀(0) = 1
    });

    it('should handle x = 0 for n > 0', () => {
      const matches = ['BESSELJ(0, 3)', '0', '3'] as RegExpMatchArray;
      const result = BESSELJ.calculate(matches, mockContext);
      expect(result).toBe(0); // Jₙ(0) = 0 for n > 0
    });

    it('should oscillate for large x', () => {
      // BESSELJ oscillates between positive and negative values
      const matches = ['BESSELJ(10, 0)', '10', '0'] as RegExpMatchArray;
      const result = BESSELJ.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(Math.abs(result as number)).toBeLessThan(1); // Bounded oscillation
    });

    it('should return NUM error for negative n', () => {
      const matches = ['BESSELJ(1, -2)', '1', '-2'] as RegExpMatchArray;
      const result = BESSELJ.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });

    it('should return VALUE error for invalid input', () => {
      const matches = ['BESSELJ("text", 0)', '"text"', '0'] as RegExpMatchArray;
      const result = BESSELJ.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });
  });

  describe('BESSELK Function (Modified Bessel Function of Second Kind)', () => {
    it('should calculate BESSELK for order 0', () => {
      const matches = ['BESSELK(1, 0)', '1', '0'] as RegExpMatchArray;
      const result = BESSELK.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeCloseTo(0.4210, 3); // Known value for K₀(1)
    });

    it('should calculate BESSELK for order 1', () => {
      const matches = ['BESSELK(1, 1)', '1', '1'] as RegExpMatchArray;
      const result = BESSELK.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeCloseTo(0.6019, 3); // Known value for K₁(1)
    });

    it('should decrease exponentially for large x', () => {
      const smallMatches = ['BESSELK(2, 0)', '2', '0'] as RegExpMatchArray;
      const small = BESSELK.calculate(smallMatches, mockContext) as number;
      
      const largeMatches = ['BESSELK(5, 0)', '5', '0'] as RegExpMatchArray;
      const large = BESSELK.calculate(largeMatches, mockContext) as number;
      
      expect(large).toBeLessThan(small);
      expect(large).toBeGreaterThan(0);
    });

    it('should return NUM error for x <= 0', () => {
      const matches = ['BESSELK(0, 0)', '0', '0'] as RegExpMatchArray;
      const result = BESSELK.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });

    it('should return NUM error for negative x', () => {
      const matches = ['BESSELK(-1, 0)', '-1', '0'] as RegExpMatchArray;
      const result = BESSELK.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });

    it('should return NUM error for negative n', () => {
      const matches = ['BESSELK(1, -1)', '1', '-1'] as RegExpMatchArray;
      const result = BESSELK.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });

    it('should handle large orders', () => {
      const matches = ['BESSELK(2, 5)', '2', '5'] as RegExpMatchArray;
      const result = BESSELK.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeGreaterThan(0);
    });
  });

  describe('BESSELY Function (Bessel Function of Second Kind)', () => {
    it('should calculate BESSELY for order 0', () => {
      const matches = ['BESSELY(1, 0)', '1', '0'] as RegExpMatchArray;
      const result = BESSELY.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeCloseTo(0.0883, 3); // Known value for Y₀(1)
    });

    it('should calculate BESSELY for order 1', () => {
      const matches = ['BESSELY(1, 1)', '1', '1'] as RegExpMatchArray;
      const result = BESSELY.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeCloseTo(-0.7812, 3); // Known value for Y₁(1)
    });

    it('should oscillate for large x', () => {
      const matches = ['BESSELY(10, 0)', '10', '0'] as RegExpMatchArray;
      const result = BESSELY.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(Math.abs(result as number)).toBeLessThan(1); // Bounded oscillation
    });

    it('should return NUM error for x <= 0', () => {
      const matches = ['BESSELY(0, 0)', '0', '0'] as RegExpMatchArray;
      const result = BESSELY.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });

    it('should return NUM error for negative x', () => {
      const matches = ['BESSELY(-2, 0)', '-2', '0'] as RegExpMatchArray;
      const result = BESSELY.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });

    it('should return NUM error for negative n', () => {
      const matches = ['BESSELY(1, -3)', '1', '-3'] as RegExpMatchArray;
      const result = BESSELY.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });

    it('should return VALUE error for invalid input', () => {
      const matches = ['BESSELY(1, "text")', '1', '"text"'] as RegExpMatchArray;
      const result = BESSELY.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });
  });

  describe('Edge Cases and Integration', () => {
    it('should verify Bessel function relationships', () => {
      // For small arguments, I₀(x) ≈ 1 + x²/4
      const x = 0.1;
      const matches = ['BESSELI(0.1, 0)', '0.1', '0'] as RegExpMatchArray;
      const result = BESSELI.calculate(matches, mockContext) as number;
      const approx = 1 + (x * x) / 4;
      expect(result).toBeCloseTo(approx, 3);
    });

    it('should handle very large orders gracefully', () => {
      const matches = ['BESSELJ(1, 100)', '1', '100'] as RegExpMatchArray;
      const result = BESSELJ.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(Math.abs(result as number)).toBeLessThan(1e-10); // Very small for large n
    });

    it('should verify Wronskian relation for J and Y', () => {
      // J_n(x)*Y_{n+1}(x) - J_{n+1}(x)*Y_n(x) = 2/(πx)
      const x = 2;
      const n = 1;
      
      const j1Matches = ['BESSELJ(2, 1)', '2', '1'] as RegExpMatchArray;
      const j1 = BESSELJ.calculate(j1Matches, mockContext) as number;
      
      const j2Matches = ['BESSELJ(2, 2)', '2', '2'] as RegExpMatchArray;
      const j2 = BESSELJ.calculate(j2Matches, mockContext) as number;
      
      const y1Matches = ['BESSELY(2, 1)', '2', '1'] as RegExpMatchArray;
      const y1 = BESSELY.calculate(y1Matches, mockContext) as number;
      
      const y2Matches = ['BESSELY(2, 2)', '2', '2'] as RegExpMatchArray;
      const y2 = BESSELY.calculate(y2Matches, mockContext) as number;
      
      const wronskian = j1 * y2 - j2 * y1;
      const expected = 2 / (Math.PI * x);
      
      expect(wronskian).toBeCloseTo(expected, 2);
    });

    it('should handle cell references', () => {
      const matches = ['BESSELK(B1, A2)', 'B1', 'A2'] as RegExpMatchArray;
      const result = BESSELK.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
    });
  });
});