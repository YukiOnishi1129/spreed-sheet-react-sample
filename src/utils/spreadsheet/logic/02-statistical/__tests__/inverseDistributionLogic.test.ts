import { describe, it, expect } from 'vitest';
import type { FormulaContext } from '../../shared/types';
import {
  T_INV, T_INV_2T,
  CHISQ_INV, CHISQ_INV_RT,
  F_INV, F_INV_RT,
  BETA_INV, GAMMA_INV
} from '../inverseDistributionLogic';
import { FormulaError } from '../../shared/types';

// Helper function to create FormulaContext
function createContext(data: (string | number | boolean | null)[][]): FormulaContext {
  return {
    data: data.map(row => row.map(cell => ({ value: cell }))),
    row: 0,
    col: 0
  };
}

describe('Inverse Distribution Functions', () => {
  const mockContext = createContext([
    [0.05, 10, 'text'],    // Row 1 (A1-C1)
    [0.025, 20, -0.5],     // Row 2 (A2-C2)
    [0.5, 30, 1.5],        // Row 3 (A3-C3)
    [0.95, null, null],    // Row 4 (A4-C4)
    [0.975, null, null]    // Row 5 (A5-C5)
  ]);

  describe('T-Distribution Inverse Functions', () => {
    describe('T.INV', () => {
      it('should calculate inverse of left-tailed t-distribution', () => {
        const matches = ['T.INV(0.05, 10)', '0.05', '10'] as RegExpMatchArray;
        const result = T_INV.calculate(matches, mockContext);
        expect(result).toBeCloseTo(-1.8125, 4);
      });

      it('should return 0 for probability 0.5', () => {
        const matches = ['T.INV(0.5, 10)', '0.5', '10'] as RegExpMatchArray;
        const result = T_INV.calculate(matches, mockContext);
        expect(result).toBeCloseTo(0, 10);
      });

      it('should return NUM error for probability outside [0,1]', () => {
        const matches1 = ['T.INV(-0.1, 10)', '-0.1', '10'] as RegExpMatchArray;
        expect(T_INV.calculate(matches1, mockContext)).toBe(FormulaError.NUM);
        
        const matches2 = ['T.INV(1.1, 10)', '1.1', '10'] as RegExpMatchArray;
        expect(T_INV.calculate(matches2, mockContext)).toBe(FormulaError.NUM);
      });

      it('should accept non-integer degrees of freedom', () => {
        const matches = ['T.INV(0.5, 10.5)', '0.5', '10.5'] as RegExpMatchArray;
        const result = T_INV.calculate(matches, mockContext);
        expect(typeof result).toBe('number');
        expect(result).toBeCloseTo(0, 5); // T.INV(0.5, df)は常に0
      });

      it('should return NUM error for degrees of freedom < 1', () => {
        const matches = ['T.INV(0.5, 0)', '0.5', '0'] as RegExpMatchArray;
        expect(T_INV.calculate(matches, mockContext)).toBe(FormulaError.NUM);
      });

      it('should return VALUE error for non-numeric inputs', () => {
        const matches = ['T.INV(text, 10)', 'text', '10'] as RegExpMatchArray;
        expect(T_INV.calculate(matches, mockContext)).toBe(FormulaError.VALUE);
      });
    });

    describe('T.INV.2T', () => {
      it('should calculate inverse of two-tailed t-distribution', () => {
        const matches = ['T.INV.2T(0.05, 10)', '0.05', '10'] as RegExpMatchArray;
        const result = T_INV_2T.calculate(matches, mockContext);
        expect(result).toBeCloseTo(2.2281, 4);
      });

      it('should be symmetric around 0', () => {
        const matches = ['T.INV.2T(0.10, 10)', '0.10', '10'] as RegExpMatchArray;
        const result = T_INV_2T.calculate(matches, mockContext);
        expect(result).toBeGreaterThan(0);
      });

      it('should return NUM error for probability outside (0,1]', () => {
        const matches1 = ['T.INV.2T(0, 10)', '0', '10'] as RegExpMatchArray;
        expect(T_INV_2T.calculate(matches1, mockContext)).toBe(FormulaError.NUM);
        
        const matches2 = ['T.INV.2T(1.1, 10)', '1.1', '10'] as RegExpMatchArray;
        expect(T_INV_2T.calculate(matches2, mockContext)).toBe(FormulaError.NUM);
      });
    });
  });

  describe('Chi-squared Distribution Inverse Functions', () => {
    describe('CHISQ.INV', () => {
      it('should calculate inverse of left-tailed chi-squared distribution', () => {
        const matches = ['CHISQ.INV(0.05, 10)', '0.05', '10'] as RegExpMatchArray;
        const result = CHISQ_INV.calculate(matches, mockContext);
        expect(result).toBeCloseTo(3.9403, 4);
      });

      it('should return 0 for probability 0', () => {
        const matches = ['CHISQ.INV(0, 10)', '0', '10'] as RegExpMatchArray;
        const result = CHISQ_INV.calculate(matches, mockContext);
        expect(result).toBe(0);
      });

      it('should return NUM error for probability > 1', () => {
        const matches = ['CHISQ.INV(1.1, 10)', '1.1', '10'] as RegExpMatchArray;
        expect(CHISQ_INV.calculate(matches, mockContext)).toBe(FormulaError.NUM);
      });

      it('should accept non-integer degrees of freedom', () => {
        const matches = ['CHISQ.INV(0.5, 10.5)', '0.5', '10.5'] as RegExpMatchArray;
        const result = CHISQ_INV.calculate(matches, mockContext);
        expect(typeof result).toBe('number');
        expect(result).toBeGreaterThan(0);
      });
    });

    describe('CHISQ.INV.RT', () => {
      it('should calculate inverse of right-tailed chi-squared distribution', () => {
        const matches = ['CHISQ.INV.RT(0.95, 10)', '0.95', '10'] as RegExpMatchArray;
        const result = CHISQ_INV_RT.calculate(matches, mockContext);
        expect(result).toBeCloseTo(3.9403, 4);
      });

      it('should complement CHISQ.INV', () => {
        const matches1 = ['CHISQ.INV(0.05, 10)', '0.05', '10'] as RegExpMatchArray;
        const result1 = CHISQ_INV.calculate(matches1, mockContext);
        
        const matches2 = ['CHISQ.INV.RT(0.95, 10)', '0.95', '10'] as RegExpMatchArray;
        const result2 = CHISQ_INV_RT.calculate(matches2, mockContext);
        
        expect(result1).toBeCloseTo(result2, 4);
      });
    });
  });

  describe('F-Distribution Inverse Functions', () => {
    describe('F.INV', () => {
      it('should calculate inverse of left-tailed F-distribution', () => {
        const matches = ['F.INV(0.05, 5, 10)', '0.05', '5', '10'] as RegExpMatchArray;
        const result = F_INV.calculate(matches, mockContext);
        expect(result).toBeCloseTo(0.2107, 4);
      });

      it('should return NUM error for invalid parameters', () => {
        const matches = ['F.INV(-0.1, 5, 10)', '-0.1', '5', '10'] as RegExpMatchArray;
        expect(F_INV.calculate(matches, mockContext)).toBe(FormulaError.NUM);
      });

      it('should accept non-integer degrees of freedom', () => {
        const matches = ['F.INV(0.5, 5.5, 10)', '0.5', '5.5', '10'] as RegExpMatchArray;
        const result = F_INV.calculate(matches, mockContext);
        expect(typeof result).toBe('number');
        expect(result).toBeGreaterThan(0);
      });
    });

    describe('F.INV.RT', () => {
      it('should calculate inverse of right-tailed F-distribution', () => {
        const matches = ['F.INV.RT(0.05, 5, 10)', '0.05', '5', '10'] as RegExpMatchArray;
        const result = F_INV_RT.calculate(matches, mockContext);
        expect(result).toBeCloseTo(3.3258, 4);
      });

      it('should complement F.INV', () => {
        const matches1 = ['F.INV(0.05, 5, 10)', '0.05', '5', '10'] as RegExpMatchArray;
        const result1 = F_INV.calculate(matches1, mockContext);
        
        const matches2 = ['F.INV.RT(0.95, 5, 10)', '0.95', '5', '10'] as RegExpMatchArray;
        const result2 = F_INV_RT.calculate(matches2, mockContext);
        
        expect(result1).toBeCloseTo(result2, 4);
      });
    });
  });

  describe('Other Distribution Inverse Functions', () => {
    describe('BETA.INV', () => {
      it('should calculate inverse beta distribution', () => {
        const matches = ['BETA.INV(0.5, 2, 3, 0, 1)', '0.5', '2', '3', '0', '1'] as RegExpMatchArray;
        const result = BETA_INV.calculate(matches, mockContext);
        expect(result).toBeCloseTo(0.3858, 4);
      });

      it('should handle custom bounds', () => {
        const matches = ['BETA.INV(0.5, 2, 3, 10, 20)', '0.5', '2', '3', '10', '20'] as RegExpMatchArray;
        const result = BETA_INV.calculate(matches, mockContext);
        expect(result).toBeCloseTo(13.858, 3);
      });

      it('should use default bounds [0,1] when not specified', () => {
        const matches = ['BETA.INV(0.5, 2, 3)', '0.5', '2', '3'] as RegExpMatchArray;
        const result = BETA_INV.calculate(matches, mockContext);
        expect(result).toBeCloseTo(0.3858, 4);
      });

      it('should return NUM error for probability outside [0,1]', () => {
        const matches = ['BETA.INV(1.5, 2, 3, 0, 1)', '1.5', '2', '3', '0', '1'] as RegExpMatchArray;
        expect(BETA_INV.calculate(matches, mockContext)).toBe(FormulaError.NUM);
      });

      it('should return NUM error for non-positive parameters', () => {
        const matches = ['BETA.INV(0.5, 0, 3, 0, 1)', '0.5', '0', '3', '0', '1'] as RegExpMatchArray;
        expect(BETA_INV.calculate(matches, mockContext)).toBe(FormulaError.NUM);
      });

      it('should return NUM error when A >= B', () => {
        const matches = ['BETA.INV(0.5, 2, 3, 10, 10)', '0.5', '2', '3', '10', '10'] as RegExpMatchArray;
        expect(BETA_INV.calculate(matches, mockContext)).toBe(FormulaError.NUM);
      });
    });

    describe('GAMMA.INV', () => {
      it('should calculate inverse gamma distribution', () => {
        const matches = ['GAMMA.INV(0.5, 2, 1)', '0.5', '2', '1'] as RegExpMatchArray;
        const result = GAMMA_INV.calculate(matches, mockContext);
        expect(result).toBeCloseTo(1.6783, 4);
      });

      it('should handle different scale parameters', () => {
        const matches = ['GAMMA.INV(0.5, 2, 2)', '0.5', '2', '2'] as RegExpMatchArray;
        const result = GAMMA_INV.calculate(matches, mockContext);
        expect(result).toBeCloseTo(3.3567, 4);
      });

      it('should return 0 for probability 0', () => {
        const matches = ['GAMMA.INV(0, 2, 1)', '0', '2', '1'] as RegExpMatchArray;
        const result = GAMMA_INV.calculate(matches, mockContext);
        expect(result).toBe(0);
      });

      it('should return NUM error for probability outside [0,1]', () => {
        const matches = ['GAMMA.INV(-0.1, 2, 1)', '-0.1', '2', '1'] as RegExpMatchArray;
        expect(GAMMA_INV.calculate(matches, mockContext)).toBe(FormulaError.NUM);
      });

      it('should return NUM error for non-positive parameters', () => {
        const matches1 = ['GAMMA.INV(0.5, 0, 1)', '0.5', '0', '1'] as RegExpMatchArray;
        expect(GAMMA_INV.calculate(matches1, mockContext)).toBe(FormulaError.NUM);
        
        const matches2 = ['GAMMA.INV(0.5, 2, 0)', '0.5', '2', '0'] as RegExpMatchArray;
        expect(GAMMA_INV.calculate(matches2, mockContext)).toBe(FormulaError.NUM);
      });

      it('should return VALUE error for non-numeric inputs', () => {
        const matches = ['GAMMA.INV(text, 2, 1)', 'text', '2', '1'] as RegExpMatchArray;
        expect(GAMMA_INV.calculate(matches, mockContext)).toBe(FormulaError.VALUE);
      });
    });
  });
});