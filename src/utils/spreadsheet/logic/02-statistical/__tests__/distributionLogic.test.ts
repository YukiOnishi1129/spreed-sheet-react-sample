import { describe, it, expect } from 'vitest';
import type { FormulaContext } from '../../shared/types';
import {
  NORM_DIST, NORM_INV, NORM_S_DIST, NORM_S_INV,
  EXPON_DIST, BINOM_DIST, POISSON_DIST,
  GAMMA_DIST, LOGNORM_DIST, LOGNORM_INV,
  WEIBULL_DIST, T_DIST, T_DIST_2T, T_DIST_RT,
  CHISQ_DIST, CHISQ_DIST_RT, F_DIST, F_DIST_RT,
  BETA_DIST, NEGBINOM_DIST, HYPGEOM_DIST,
  CONFIDENCE_NORM, CONFIDENCE_T,
  Z_TEST, T_TEST, F_TEST, CHISQ_TEST
} from '../distributionLogic';
import { FormulaError } from '../../shared/types';

// Helper function to create FormulaContext
function createContext(data: any[][]): FormulaContext {
  return {
    data,
    row: 0,
    col: 0
  };
}

describe('Distribution Functions', () => {
  const mockContext = createContext([
    [0, 100, 98],   // Row 1 (A1-C1)
    [1, 105, 102],  // Row 2 (A2-C2)
    [2, 95, 99],    // Row 3 (A3-C3)
    [3, 110, 101],  // Row 4 (A4-C4)
    [4, 90, 100]    // Row 5 (A5-C5)
  ]);

  describe('Normal Distribution Functions', () => {
    describe('NORM.DIST', () => {
      it('should calculate normal distribution PDF', () => {
        const matches = ['NORM.DIST(0, 0, 1, FALSE)', '0', '0', '1', 'FALSE'] as RegExpMatchArray;
        const result = NORM_DIST.calculate(matches, mockContext);
        expect(result).toBeCloseTo(0.3989, 4); // Standard normal PDF at 0
      });

      it('should calculate normal distribution CDF', () => {
        const matches = ['NORM.DIST(0, 0, 1, TRUE)', '0', '0', '1', 'TRUE'] as RegExpMatchArray;
        const result = NORM_DIST.calculate(matches, mockContext);
        expect(result).toBe(0.5); // Standard normal CDF at 0
      });

      it('should return NUM error for non-positive standard deviation', () => {
        const matches = ['NORM.DIST(0, 0, 0, TRUE)', '0', '0', '0', 'TRUE'] as RegExpMatchArray;
        expect(NORM_DIST.calculate(matches, mockContext)).toBe(FormulaError.NUM);
      });
    });

    describe('NORM.INV', () => {
      it('should calculate inverse normal distribution', () => {
        const matches = ['NORM.INV(0.5, 0, 1)', '0.5', '0', '1'] as RegExpMatchArray;
        const result = NORM_INV.calculate(matches, mockContext);
        expect(result).toBeCloseTo(0, 10);
      });

      it('should return NUM error for probability outside [0,1]', () => {
        const matches = ['NORM.INV(1.5, 0, 1)', '1.5', '0', '1'] as RegExpMatchArray;
        expect(NORM_INV.calculate(matches, mockContext)).toBe(FormulaError.NUM);
      });
    });

    describe('NORM.S.DIST', () => {
      it('should calculate standard normal distribution', () => {
        const matches = ['NORM.S.DIST(0, TRUE)', '0', 'TRUE'] as RegExpMatchArray;
        const result = NORM_S_DIST.calculate(matches, mockContext);
        expect(result).toBe(0.5);
      });

      it('should calculate standard normal PDF', () => {
        const matches = ['NORM.S.DIST(0, FALSE)', '0', 'FALSE'] as RegExpMatchArray;
        const result = NORM_S_DIST.calculate(matches, mockContext);
        expect(result).toBeCloseTo(0.3989, 4);
      });
    });

    describe('NORM.S.INV', () => {
      it('should calculate inverse standard normal', () => {
        const matches = ['NORM.S.INV(0.975)', '0.975'] as RegExpMatchArray;
        const result = NORM_S_INV.calculate(matches, mockContext);
        expect(result).toBeCloseTo(1.96, 2);
      });
    });
  });

  describe('Exponential Distribution', () => {
    describe('EXPON.DIST', () => {
      it('should calculate exponential distribution PDF', () => {
        const matches = ['EXPON.DIST(1, 1, FALSE)', '1', '1', 'FALSE'] as RegExpMatchArray;
        const result = EXPON_DIST.calculate(matches, mockContext);
        expect(result).toBeCloseTo(0.3679, 4);
      });

      it('should calculate exponential distribution CDF', () => {
        const matches = ['EXPON.DIST(1, 1, TRUE)', '1', '1', 'TRUE'] as RegExpMatchArray;
        const result = EXPON_DIST.calculate(matches, mockContext);
        expect(result).toBeCloseTo(0.6321, 4);
      });

      it('should return NUM error for negative x or lambda', () => {
        const matches = ['EXPON.DIST(-1, 1, TRUE)', '-1', '1', 'TRUE'] as RegExpMatchArray;
        expect(EXPON_DIST.calculate(matches, mockContext)).toBe(FormulaError.NUM);
      });
    });
  });

  describe('Binomial Distribution', () => {
    describe('BINOM.DIST', () => {
      it('should calculate binomial distribution PMF', () => {
        const matches = ['BINOM.DIST(2, 5, 0.5, FALSE)', '2', '5', '0.5', 'FALSE'] as RegExpMatchArray;
        const result = BINOM_DIST.calculate(matches, mockContext);
        expect(result).toBeCloseTo(0.3125, 4);
      });

      it('should calculate binomial distribution CDF', () => {
        const matches = ['BINOM.DIST(2, 5, 0.5, TRUE)', '2', '5', '0.5', 'TRUE'] as RegExpMatchArray;
        const result = BINOM_DIST.calculate(matches, mockContext);
        expect(result).toBe(0.5);
      });

      it('should return NUM error for invalid parameters', () => {
        const matches = ['BINOM.DIST(6, 5, 0.5, FALSE)', '6', '5', '0.5', 'FALSE'] as RegExpMatchArray;
        expect(BINOM_DIST.calculate(matches, mockContext)).toBe(FormulaError.NUM);
      });
    });
  });

  describe('Poisson Distribution', () => {
    describe('POISSON.DIST', () => {
      it('should calculate Poisson distribution PMF', () => {
        const matches = ['POISSON.DIST(2, 3, FALSE)', '2', '3', 'FALSE'] as RegExpMatchArray;
        const result = POISSON_DIST.calculate(matches, mockContext);
        expect(result).toBeCloseTo(0.2240, 4);
      });

      it('should calculate Poisson distribution CDF', () => {
        const matches = ['POISSON.DIST(2, 3, TRUE)', '2', '3', 'TRUE'] as RegExpMatchArray;
        const result = POISSON_DIST.calculate(matches, mockContext);
        expect(result).toBeCloseTo(0.4232, 4);
      });
    });
  });

  describe('T-Distribution Functions', () => {
    describe('T.DIST', () => {
      it('should calculate left-tailed t-distribution', () => {
        const matches = ['T.DIST(0, 10, TRUE)', '0', '10', 'TRUE'] as RegExpMatchArray;
        const result = T_DIST.calculate(matches, mockContext);
        expect(result).toBe(0.5);
      });

      it('should return NUM error for non-integer degrees of freedom', () => {
        const matches = ['T.DIST(0, 10.5, TRUE)', '0', '10.5', 'TRUE'] as RegExpMatchArray;
        expect(T_DIST.calculate(matches, mockContext)).toBe(FormulaError.NUM);
      });
    });

    describe('T.DIST.2T', () => {
      it('should calculate two-tailed t-distribution', () => {
        const matches = ['T.DIST.2T(2, 10)', '2', '10'] as RegExpMatchArray;
        const result = T_DIST_2T.calculate(matches, mockContext);
        expect(result).toBeCloseTo(0.0734, 4);
      });

      it('should return NUM error for negative x', () => {
        const matches = ['T.DIST.2T(-2, 10)', '-2', '10'] as RegExpMatchArray;
        expect(T_DIST_2T.calculate(matches, mockContext)).toBe(FormulaError.NUM);
      });
    });

    describe('T.DIST.RT', () => {
      it('should calculate right-tailed t-distribution', () => {
        const matches = ['T.DIST.RT(0, 10)', '0', '10'] as RegExpMatchArray;
        const result = T_DIST_RT.calculate(matches, mockContext);
        expect(result).toBe(0.5);
      });
    });
  });

  describe('Chi-squared Distribution', () => {
    describe('CHISQ.DIST', () => {
      it('should calculate chi-squared distribution CDF', () => {
        const matches = ['CHISQ.DIST(5, 5, TRUE)', '5', '5', 'TRUE'] as RegExpMatchArray;
        const result = CHISQ_DIST.calculate(matches, mockContext);
        expect(result).toBeCloseTo(0.5841, 4);
      });

      it('should return NUM error for negative x', () => {
        const matches = ['CHISQ.DIST(-1, 5, TRUE)', '-1', '5', 'TRUE'] as RegExpMatchArray;
        expect(CHISQ_DIST.calculate(matches, mockContext)).toBe(FormulaError.NUM);
      });
    });

    describe('CHISQ.DIST.RT', () => {
      it('should calculate right-tailed chi-squared distribution', () => {
        const matches = ['CHISQ.DIST.RT(5, 5)', '5', '5'] as RegExpMatchArray;
        const result = CHISQ_DIST_RT.calculate(matches, mockContext);
        expect(result).toBeCloseTo(0.4159, 4);
      });
    });
  });

  describe('F-Distribution', () => {
    describe('F.DIST', () => {
      it('should calculate F-distribution CDF', () => {
        const matches = ['F.DIST(1, 5, 10, TRUE)', '1', '5', '10', 'TRUE'] as RegExpMatchArray;
        const result = F_DIST.calculate(matches, mockContext);
        expect(result).toBeCloseTo(0.5, 1);
      });
    });

    describe('F.DIST.RT', () => {
      it('should calculate right-tailed F-distribution', () => {
        const matches = ['F.DIST.RT(1, 5, 10)', '1', '5', '10'] as RegExpMatchArray;
        const result = F_DIST_RT.calculate(matches, mockContext);
        expect(result).toBeCloseTo(0.5, 1);
      });
    });
  });

  describe('Other Distributions', () => {
    describe('BETA.DIST', () => {
      it('should calculate beta distribution', () => {
        const matches = ['BETA.DIST(0.5, 2, 3, TRUE, 0, 1)', '0.5', '2', '3', 'TRUE', '0', '1'] as RegExpMatchArray;
        const result = BETA_DIST.calculate(matches, mockContext);
        expect(result).toBeCloseTo(0.6875, 4);
      });

      it('should return NUM error for x outside bounds', () => {
        const matches = ['BETA.DIST(1.5, 2, 3, TRUE, 0, 1)', '1.5', '2', '3', 'TRUE', '0', '1'] as RegExpMatchArray;
        expect(BETA_DIST.calculate(matches, mockContext)).toBe(FormulaError.NUM);
      });
    });

    describe('GAMMA.DIST', () => {
      it('should calculate gamma distribution', () => {
        const matches = ['GAMMA.DIST(2, 2, 1, TRUE)', '2', '2', '1', 'TRUE'] as RegExpMatchArray;
        const result = GAMMA_DIST.calculate(matches, mockContext);
        expect(result).toBeCloseTo(0.5940, 4);
      });
    });

    describe('LOGNORM.DIST', () => {
      it('should calculate lognormal distribution', () => {
        const matches = ['LOGNORM.DIST(1, 0, 1, TRUE)', '1', '0', '1', 'TRUE'] as RegExpMatchArray;
        const result = LOGNORM_DIST.calculate(matches, mockContext);
        expect(result).toBe(0.5);
      });
    });

    describe('WEIBULL.DIST', () => {
      it('should calculate Weibull distribution', () => {
        const matches = ['WEIBULL.DIST(1, 1, 1, TRUE)', '1', '1', '1', 'TRUE'] as RegExpMatchArray;
        const result = WEIBULL_DIST.calculate(matches, mockContext);
        expect(result).toBeCloseTo(0.6321, 4);
      });
    });

    describe('NEGBINOM.DIST', () => {
      it('should calculate negative binomial distribution', () => {
        const matches = ['NEGBINOM.DIST(5, 3, 0.4, FALSE)', '5', '3', '0.4', 'FALSE'] as RegExpMatchArray;
        const result = NEGBINOM_DIST.calculate(matches, mockContext);
        expect(result).toBeCloseTo(0.1382, 4);
      });
    });

    describe('HYPGEOM.DIST', () => {
      it('should calculate hypergeometric distribution', () => {
        const matches = ['HYPGEOM.DIST(1, 4, 8, 20, FALSE)', '1', '4', '8', '20', 'FALSE'] as RegExpMatchArray;
        const result = HYPGEOM_DIST.calculate(matches, mockContext);
        expect(result).toBeCloseTo(0.3633, 4);
      });
    });
  });

  describe('Confidence Interval Functions', () => {
    describe('CONFIDENCE.NORM', () => {
      it('should calculate confidence interval using normal distribution', () => {
        const matches = ['CONFIDENCE.NORM(0.05, 10, 100)', '0.05', '10', '100'] as RegExpMatchArray;
        const result = CONFIDENCE_NORM.calculate(matches, mockContext);
        expect(result).toBeCloseTo(1.96, 2);
      });

      it('should return NUM error for invalid alpha', () => {
        const matches = ['CONFIDENCE.NORM(0, 10, 100)', '0', '10', '100'] as RegExpMatchArray;
        expect(CONFIDENCE_NORM.calculate(matches, mockContext)).toBe(FormulaError.NUM);
      });
    });

    describe('CONFIDENCE.T', () => {
      it('should calculate confidence interval using t-distribution', () => {
        const matches = ['CONFIDENCE.T(0.05, 10, 30)', '0.05', '10', '30'] as RegExpMatchArray;
        const result = CONFIDENCE_T.calculate(matches, mockContext);
        expect(result).toBeCloseTo(3.659, 3);
      });
    });
  });

  describe('Hypothesis Testing Functions', () => {
    describe('Z.TEST', () => {
      it('should perform z-test', () => {
        const matches = ['Z.TEST(B1:B5, 100, 10)', 'B1:B5', '100', '10'] as RegExpMatchArray;
        const result = Z_TEST.calculate(matches, mockContext);
        expect(result).toBeCloseTo(0.5, 1);
      });
    });

    describe('T.TEST', () => {
      it('should perform paired t-test', () => {
        const matches = ['T.TEST(B1:B5, C1:C5, 2, 1)', 'B1:B5', 'C1:C5', '2', '1'] as RegExpMatchArray;
        const result = T_TEST.calculate(matches, mockContext);
        expect(result).toBeCloseTo(0.5, 1);
      });

      it('should return NA error for different array lengths in paired test', () => {
        const matches = ['T.TEST(B1:B3, C1:C5, 2, 1)', 'B1:B3', 'C1:C5', '2', '1'] as RegExpMatchArray;
        expect(T_TEST.calculate(matches, mockContext)).toBe(FormulaError.NA);
      });
    });

    describe('F.TEST', () => {
      it('should perform F-test', () => {
        const matches = ['F.TEST(B1:B5, C1:C5)', 'B1:B5', 'C1:C5'] as RegExpMatchArray;
        const result = F_TEST.calculate(matches, mockContext);
        expect(result).toBeGreaterThan(0);
        expect(result).toBeLessThan(1);
      });
    });

    describe('CHISQ.TEST', () => {
      it('should perform chi-squared test', () => {
        const matches = ['CHISQ.TEST(B1:B5, C1:C5)', 'B1:B5', 'C1:C5'] as RegExpMatchArray;
        const result = CHISQ_TEST.calculate(matches, mockContext);
        expect(result).toBeGreaterThan(0);
        expect(result).toBeLessThan(1);
      });
    });
  });
});