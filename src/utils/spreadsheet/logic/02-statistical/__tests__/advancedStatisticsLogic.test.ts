import { describe, it, expect } from 'vitest';
import {
  COVARIANCE_P, COVARIANCE_S, CORREL, PEARSON, DEVSQ,
  KURT, SKEW, SKEW_P, STANDARDIZE,
  PERCENTILE_INC, PERCENTILE_EXC, QUARTILE_INC, QUARTILE_EXC,
  PERCENTRANK, GAMMALN
} from '../advancedStatisticsLogic';
import { FormulaError } from '../../shared/types';
import type { FormulaContext } from '../../shared/types';

// Helper function to create FormulaContext
function createContext(data: any[][]): FormulaContext {
  return {
    data,
    row: 0,
    col: 0
  };
}

describe('Advanced Statistics Functions', () => {
  const mockContext = createContext([
    [1, 2, 'text', 10, 1, 1],        // Row 1 (A1-F1)
    [2, 4, '', 20, 2, 1],            // Row 2 (A2-F2)
    [3, 6, null, 30, 2, 2],          // Row 3 (A3-F3)
    [4, 8, null, 40, 3, 3],          // Row 4 (A4-F4)
    [5, 10, null, 50, 4, 5],         // Row 5 (A5-F5)
    [null, null, null, null, null, 8],  // Row 6 (A6-F6)
    [null, null, null, null, null, 13]  // Row 7 (A7-F7)
  ]);

  describe('Covariance Functions', () => {
    describe('COVARIANCE.P', () => {
      it('should calculate population covariance', () => {
        const matches = ['COVARIANCE.P(A1:A5, B1:B5)', 'A1:A5', 'B1:B5'] as RegExpMatchArray;
        const result = COVARIANCE_P.calculate(matches, mockContext);
        expect(result).toBe(4); // Perfect linear relationship, covariance = 4
      });

      it('should return NA error for different array lengths', () => {
        const matches = ['COVARIANCE.P(A1:A3, B1:B5)', 'A1:A3', 'B1:B5'] as RegExpMatchArray;
        const result = COVARIANCE_P.calculate(matches, mockContext);
        expect(result).toBe(FormulaError.NA);
      });

      it('should return NA error for empty arrays', () => {
        const matches = ['COVARIANCE.P(C1:C3, C1:C3)', 'C1:C3', 'C1:C3'] as RegExpMatchArray;
        const result = COVARIANCE_P.calculate(matches, mockContext);
        expect(result).toBe(FormulaError.NA);
      });

      it('should handle single values', () => {
        const matches = ['COVARIANCE.P(A1, B1)', 'A1', 'B1'] as RegExpMatchArray;
        const result = COVARIANCE_P.calculate(matches, mockContext);
        expect(result).toBe(0); // Single value pairs have no variance
      });
    });

    describe('COVARIANCE.S', () => {
      it('should calculate sample covariance', () => {
        const matches = ['COVARIANCE.S(A1:A5, B1:B5)', 'A1:A5', 'B1:B5'] as RegExpMatchArray;
        const result = COVARIANCE_S.calculate(matches, mockContext);
        expect(result).toBe(5); // Sample covariance = 5
      });

      it('should return NA error for arrays with less than 2 elements', () => {
        const matches = ['COVARIANCE.S(A1, B1)', 'A1', 'B1'] as RegExpMatchArray;
        const result = COVARIANCE_S.calculate(matches, mockContext);
        expect(result).toBe(FormulaError.NA);
      });
    });
  });

  describe('Correlation Functions', () => {
    describe('CORREL', () => {
      it('should calculate correlation coefficient', () => {
        const matches = ['CORREL(A1:A5, B1:B5)', 'A1:A5', 'B1:B5'] as RegExpMatchArray;
        const result = CORREL.calculate(matches, mockContext);
        expect(result).toBe(1); // Perfect positive correlation
      });

      it('should return NA error for different array lengths', () => {
        const matches = ['CORREL(A1:A3, B1:B5)', 'A1:A3', 'B1:B5'] as RegExpMatchArray;
        const result = CORREL.calculate(matches, mockContext);
        expect(result).toBe(FormulaError.NA);
      });

      it('should return DIV0 error when standard deviation is zero', () => {
        const matches = ['CORREL(A1:A1, B1:B1)', 'A1:A1', 'B1:B1'] as RegExpMatchArray;
        const result = CORREL.calculate(matches, mockContext);
        expect(result).toBe(FormulaError.DIV0);
      });
    });

    describe('PEARSON', () => {
      it('should calculate Pearson correlation coefficient (same as CORREL)', () => {
        const matches = ['PEARSON(A1:A5, B1:B5)', 'A1:A5', 'B1:B5'] as RegExpMatchArray;
        const result = PEARSON.calculate(matches, mockContext);
        expect(result).toBe(1); // Same as CORREL
      });
    });
  });

  describe('Deviation and Distribution Functions', () => {
    describe('DEVSQ', () => {
      it('should calculate sum of squared deviations', () => {
        const matches = ['DEVSQ(A1:A5)', 'A1:A5'] as RegExpMatchArray;
        const result = DEVSQ.calculate(matches, mockContext);
        expect(result).toBe(10); // Sum of squared deviations from mean(3)
      });

      it('should return 0 for single value', () => {
        const matches = ['DEVSQ(A1)', 'A1'] as RegExpMatchArray;
        const result = DEVSQ.calculate(matches, mockContext);
        expect(result).toBe(0);
      });

      it('should return NUM error for empty array', () => {
        const matches = ['DEVSQ(C1:C3)', 'C1:C3'] as RegExpMatchArray;
        const result = DEVSQ.calculate(matches, mockContext);
        expect(result).toBe(FormulaError.NUM);
      });
    });

    describe('KURT', () => {
      it('should calculate kurtosis', () => {
        const matches = ['KURT(D1:D5)', 'D1:D5'] as RegExpMatchArray;
        const result = KURT.calculate(matches, mockContext);
        expect(result).toBeCloseTo(-1.2); // Kurtosis of uniform distribution
      });

      it('should return DIV0 error for less than 4 values', () => {
        const matches = ['KURT(A1:A3)', 'A1:A3'] as RegExpMatchArray;
        const result = KURT.calculate(matches, mockContext);
        expect(result).toBe(FormulaError.DIV0);
      });
    });

    describe('SKEW', () => {
      it('should calculate sample skewness', () => {
        const matches = ['SKEW(A1:A5)', 'A1:A5'] as RegExpMatchArray;
        const result = SKEW.calculate(matches, mockContext);
        expect(result).toBe(0); // Symmetric distribution
      });

      it('should return DIV0 error for less than 3 values', () => {
        const matches = ['SKEW(A1:A2)', 'A1:A2'] as RegExpMatchArray;
        const result = SKEW.calculate(matches, mockContext);
        expect(result).toBe(FormulaError.DIV0);
      });
    });

    describe('SKEW.P', () => {
      it('should calculate population skewness', () => {
        const matches = ['SKEW.P(A1:A5)', 'A1:A5'] as RegExpMatchArray;
        const result = SKEW_P.calculate(matches, mockContext);
        expect(result).toBe(0); // Symmetric distribution
      });

      it('should return DIV0 error for less than 3 values', () => {
        const matches = ['SKEW.P(A1:A2)', 'A1:A2'] as RegExpMatchArray;
        const result = SKEW_P.calculate(matches, mockContext);
        expect(result).toBe(FormulaError.DIV0);
      });
    });
  });

  describe('Standardization', () => {
    describe('STANDARDIZE', () => {
      it('should standardize a value', () => {
        const matches = ['STANDARDIZE(3, 3, 2)', '3', '3', '2'] as RegExpMatchArray;
        const result = STANDARDIZE.calculate(matches, mockContext);
        expect(result).toBe(0); // (3-3)/2 = 0
      });

      it('should calculate z-score correctly', () => {
        const matches = ['STANDARDIZE(5, 3, 2)', '5', '3', '2'] as RegExpMatchArray;
        const result = STANDARDIZE.calculate(matches, mockContext);
        expect(result).toBe(1); // (5-3)/2 = 1
      });

      it('should return NUM error for zero standard deviation', () => {
        const matches = ['STANDARDIZE(5, 3, 0)', '5', '3', '0'] as RegExpMatchArray;
        const result = STANDARDIZE.calculate(matches, mockContext);
        expect(result).toBe(FormulaError.NUM);
      });

      it('should return VALUE error for non-numeric inputs', () => {
        const matches = ['STANDARDIZE(text, 3, 2)', 'text', '3', '2'] as RegExpMatchArray;
        const result = STANDARDIZE.calculate(matches, mockContext);
        expect(result).toBe(FormulaError.VALUE);
      });
    });
  });

  describe('Percentile Functions', () => {
    describe('PERCENTILE.INC', () => {
      it('should calculate percentile with inclusive method', () => {
        const matches = ['PERCENTILE.INC(A1:A5, 0.5)', 'A1:A5', '0.5'] as RegExpMatchArray;
        const result = PERCENTILE_INC.calculate(matches, mockContext);
        expect(result).toBe(3); // Median
      });

      it('should return minimum for k=0', () => {
        const matches = ['PERCENTILE.INC(A1:A5, 0)', 'A1:A5', '0'] as RegExpMatchArray;
        const result = PERCENTILE_INC.calculate(matches, mockContext);
        expect(result).toBe(1);
      });

      it('should return maximum for k=1', () => {
        const matches = ['PERCENTILE.INC(A1:A5, 1)', 'A1:A5', '1'] as RegExpMatchArray;
        const result = PERCENTILE_INC.calculate(matches, mockContext);
        expect(result).toBe(5);
      });

      it('should return NUM error for k outside [0,1]', () => {
        const matches = ['PERCENTILE.INC(A1:A5, 1.5)', 'A1:A5', '1.5'] as RegExpMatchArray;
        const result = PERCENTILE_INC.calculate(matches, mockContext);
        expect(result).toBe(FormulaError.NUM);
      });
    });

    describe('PERCENTILE.EXC', () => {
      it('should calculate percentile with exclusive method', () => {
        const matches = ['PERCENTILE.EXC(A1:A5, 0.5)', 'A1:A5', '0.5'] as RegExpMatchArray;
        const result = PERCENTILE_EXC.calculate(matches, mockContext);
        expect(result).toBe(3); // Median
      });

      it('should return NUM error for k=0 or k=1', () => {
        const matches1 = ['PERCENTILE.EXC(A1:A5, 0)', 'A1:A5', '0'] as RegExpMatchArray;
        expect(PERCENTILE_EXC.calculate(matches1, mockContext)).toBe(FormulaError.NUM);
        
        const matches2 = ['PERCENTILE.EXC(A1:A5, 1)', 'A1:A5', '1'] as RegExpMatchArray;
        expect(PERCENTILE_EXC.calculate(matches2, mockContext)).toBe(FormulaError.NUM);
      });
    });

    describe('QUARTILE.INC', () => {
      it('should calculate quartiles with inclusive method', () => {
        const matches1 = ['QUARTILE.INC(A1:A5, 0)', 'A1:A5', '0'] as RegExpMatchArray;
        expect(QUARTILE_INC.calculate(matches1, mockContext)).toBe(1); // Min
        
        const matches2 = ['QUARTILE.INC(A1:A5, 1)', 'A1:A5', '1'] as RegExpMatchArray;
        expect(QUARTILE_INC.calculate(matches2, mockContext)).toBe(2); // Q1
        
        const matches3 = ['QUARTILE.INC(A1:A5, 2)', 'A1:A5', '2'] as RegExpMatchArray;
        expect(QUARTILE_INC.calculate(matches3, mockContext)).toBe(3); // Median
        
        const matches4 = ['QUARTILE.INC(A1:A5, 3)', 'A1:A5', '3'] as RegExpMatchArray;
        expect(QUARTILE_INC.calculate(matches4, mockContext)).toBe(4); // Q3
        
        const matches5 = ['QUARTILE.INC(A1:A5, 4)', 'A1:A5', '4'] as RegExpMatchArray;
        expect(QUARTILE_INC.calculate(matches5, mockContext)).toBe(5); // Max
      });

      it('should return NUM error for invalid quartile', () => {
        const matches = ['QUARTILE.INC(A1:A5, 5)', 'A1:A5', '5'] as RegExpMatchArray;
        expect(QUARTILE_INC.calculate(matches, mockContext)).toBe(FormulaError.NUM);
      });
    });
  });

  describe('Rank Functions', () => {
    describe('PERCENTRANK', () => {
      it('should calculate percent rank', () => {
        const matches = ['PERCENTRANK(A1:A5, 3)', 'A1:A5', '3'] as RegExpMatchArray;
        const result = PERCENTRANK.calculate(matches, mockContext);
        expect(result).toBe(0.5); // 3 is at 50th percentile
      });

      it('should return 0 for minimum value', () => {
        const matches = ['PERCENTRANK(A1:A5, 1)', 'A1:A5', '1'] as RegExpMatchArray;
        const result = PERCENTRANK.calculate(matches, mockContext);
        expect(result).toBe(0);
      });

      it('should return 1 for maximum value', () => {
        const matches = ['PERCENTRANK(A1:A5, 5)', 'A1:A5', '5'] as RegExpMatchArray;
        const result = PERCENTRANK.calculate(matches, mockContext);
        expect(result).toBe(1);
      });

      it('should interpolate for values not in array', () => {
        const matches = ['PERCENTRANK(A1:A5, 2.5)', 'A1:A5', '2.5'] as RegExpMatchArray;
        const result = PERCENTRANK.calculate(matches, mockContext);
        expect(result).toBe(0.375); // Interpolated value
      });

      it('should return NA error for value outside range', () => {
        const matches = ['PERCENTRANK(A1:A5, 10)', 'A1:A5', '10'] as RegExpMatchArray;
        const result = PERCENTRANK.calculate(matches, mockContext);
        expect(result).toBe(FormulaError.NA);
      });
    });
  });

  describe('Special Functions', () => {
    describe('GAMMALN', () => {
      it('should calculate natural log of gamma function', () => {
        const matches = ['GAMMALN(1)', '1'] as RegExpMatchArray;
        const result = GAMMALN.calculate(matches, mockContext);
        expect(result).toBeCloseTo(0); // ln(Γ(1)) = ln(0!) = ln(1) = 0
      });

      it('should calculate for integer values', () => {
        const matches = ['GAMMALN(5)', '5'] as RegExpMatchArray;
        const result = GAMMALN.calculate(matches, mockContext);
        expect(result).toBeCloseTo(Math.log(24)); // ln(Γ(5)) = ln(4!) = ln(24)
      });

      it('should return NUM error for non-positive values', () => {
        const matches = ['GAMMALN(0)', '0'] as RegExpMatchArray;
        const result = GAMMALN.calculate(matches, mockContext);
        expect(result).toBe(FormulaError.NUM);
      });

      it('should return VALUE error for non-numeric input', () => {
        const matches = ['GAMMALN(text)', 'text'] as RegExpMatchArray;
        const result = GAMMALN.calculate(matches, mockContext);
        expect(result).toBe(FormulaError.VALUE);
      });
    });
  });
});