import { describe, it, expect } from 'vitest';
import {
  PMT, PV, FV, NPV, IRR, PPMT, IPMT, RATE
} from '../financialLogic';
import { FormulaError } from '../../shared/types';
import type { FormulaContext } from '../../shared/types';

// Helper function to create FormulaContext
const createContext = (data: (string | number | boolean | null)[][]): FormulaContext => ({
  data: data.map(row => row.map(cell => ({ value: cell }))),
  row: 0,
  col: 0
});

describe('Financial Functions', () => {
  const mockContext = createContext([
    [0.05, 60, 300000, 0, 1], // rate, nper, pv, fv, type
    [0.08, 10, -1000, 500, 0], // Another loan scenario
    [-100, 200, 300, -200, 100], // Cash flows for NPV/IRR
    [12, 5] // periods and rate data
  ]);

  describe('PMT Function (Payment)', () => {
    it('should calculate monthly payment for loan', () => {
      // 5% annual rate (0.05/12 monthly), 60 months, $300,000 loan
      const matches = ['PMT(0.004167, 60, 300000)', '0.004167', '60', '300000'] as RegExpMatchArray;
      const result = PMT.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeCloseTo(-5661.43, 1); // Typical mortgage payment
    });

    it('should handle zero interest rate', () => {
      const matches = ['PMT(0, 60, 300000)', '0', '60', '300000'] as RegExpMatchArray;
      const result = PMT.calculate(matches, mockContext);
      expect(result).toBeCloseTo(-5000, 2); // 300000 / 60
    });

    it('should handle beginning-of-period payments', () => {
      const matches = ['PMT(0.004167, 60, 300000, 0, 1)', '0.004167', '60', '300000', '0', '1'] as RegExpMatchArray;
      const result = PMT.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeCloseTo(-5637.93, 1); // Slightly less due to beginning payment
    });

    it('should include future value in calculation', () => {
      const matches = ['PMT(0.004167, 60, 300000, 10000)', '0.004167', '60', '300000', '10000'] as RegExpMatchArray;
      const result = PMT.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeLessThan(-5659); // Should be more negative due to future value
    });

    it('should return VALUE error for invalid inputs', () => {
      const matches = ['PMT("invalid", 60, 300000)', '"invalid"', '60', '300000'] as RegExpMatchArray;
      const result = PMT.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });

    it('should return DIV0 error for zero periods', () => {
      const matches = ['PMT(0.05, 0, 300000)', '0.05', '0', '300000'] as RegExpMatchArray;
      const result = PMT.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.DIV0);
    });

    it('should handle cell references', () => {
      const matches = ['PMT(A1, B1, C1)', 'A1', 'B1', 'C1'] as RegExpMatchArray;
      const result = PMT.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
    });
  });

  describe('PV Function (Present Value)', () => {
    it('should calculate present value of annuity', () => {
      // 5% rate, 10 periods, $1000 payment
      const matches = ['PV(0.05, 10, 1000)', '0.05', '10', '1000'] as RegExpMatchArray;
      const result = PV.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeCloseTo(-7721.73, 1);
    });

    it('should handle zero interest rate', () => {
      const matches = ['PV(0, 10, 1000)', '0', '10', '1000'] as RegExpMatchArray;
      const result = PV.calculate(matches, mockContext);
      expect(result).toBe(-10000); // Simple: 1000 * 10
    });

    it('should handle future value parameter', () => {
      const matches = ['PV(0.05, 10, 1000, 5000)', '0.05', '10', '1000', '5000'] as RegExpMatchArray;
      const result = PV.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeLessThan(-7721); // More negative due to future value
    });

    it('should handle beginning-of-period payments', () => {
      const matches = ['PV(0.05, 10, 1000, 0, 1)', '0.05', '10', '1000', '0', '1'] as RegExpMatchArray;
      const result = PV.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeCloseTo(-8107.82, 1); // Higher due to earlier payments
    });

    it('should return VALUE error for invalid inputs', () => {
      const matches = ['PV("invalid", 10, 1000)', '"invalid"', '10', '1000'] as RegExpMatchArray;
      const result = PV.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });

    it('should handle cell references', () => {
      const matches = ['PV(A2, B2, C2)', 'A2', 'B2', 'C2'] as RegExpMatchArray;
      const result = PV.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
    });
  });

  describe('FV Function (Future Value)', () => {
    it('should calculate future value of annuity', () => {
      // 5% rate, 10 periods, $1000 payment
      const matches = ['FV(0.05, 10, 1000)', '0.05', '10', '1000'] as RegExpMatchArray;
      const result = FV.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeCloseTo(-12577.89, 1);
    });

    it('should handle zero interest rate', () => {
      const matches = ['FV(0, 10, 1000)', '0', '10', '1000'] as RegExpMatchArray;
      const result = FV.calculate(matches, mockContext);
      expect(result).toBe(-10000); // Simple: 1000 * 10
    });

    it('should handle present value parameter', () => {
      const matches = ['FV(0.05, 10, 1000, 5000)', '0.05', '10', '1000', '5000'] as RegExpMatchArray;
      const result = FV.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeLessThan(-12577); // More negative due to present value
    });

    it('should handle beginning-of-period payments', () => {
      const matches = ['FV(0.05, 10, 1000, 0, 1)', '0.05', '10', '1000', '0', '1'] as RegExpMatchArray;
      const result = FV.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeCloseTo(-13206.79, 1); // Higher due to earlier payments
    });

    it('should return VALUE error for invalid inputs', () => {
      const matches = ['FV(0.05, "invalid", 1000)', '0.05', '"invalid"', '1000'] as RegExpMatchArray;
      const result = FV.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });
  });

  describe('NPV Function (Net Present Value)', () => {
    it('should calculate NPV from cash flow series', () => {
      const npvContext = createContext([
        [-1000, 200, 300, 400, 500] // Initial investment + cash flows
      ]);
      
      const matches = ['NPV(0.1, A1:E1)', '0.1', 'A1:E1'] as RegExpMatchArray;
      const result = NPV.calculate(matches, npvContext);
      expect(typeof result).toBe('number');
      expect(result).toBeCloseTo(65.26, 1); // Positive NPV
    });

    it('should handle individual cell arguments', () => {
      const matches = ['NPV(0.1, -1000, 200, 300, 400, 500)', '0.1', '-1000, 200, 300, 400, 500'] as RegExpMatchArray;
      const result = NPV.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeCloseTo(65.26, 1);
    });

    it('should handle negative NPV', () => {
      const negNpvContext = createContext([
        [-2000, 200, 300, 400, 500] // Large initial investment
      ]);
      
      const matches = ['NPV(0.15, A1:E1)', '0.15', 'A1:E1'] as RegExpMatchArray;
      const result = NPV.calculate(matches, negNpvContext);
      expect(typeof result).toBe('number');
      expect(result).toBeLessThan(0); // Negative NPV
    });

    it('should return VALUE error for invalid discount rate', () => {
      const matches = ['NPV("invalid", A1:E1)', '"invalid"', 'A1:E1'] as RegExpMatchArray;
      const result = NPV.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });

    it('should return VALUE error for empty cash flows', () => {
      const emptyContext = createContext([[]]);
      const matches = ['NPV(0.1, A1:A1)', '0.1', 'A1:A1'] as RegExpMatchArray;
      const result = NPV.calculate(matches, emptyContext);
      expect(result).toBe(FormulaError.VALUE);
    });

    it('should filter out non-numeric values', () => {
      const mixedContext = createContext([
        [-1000, 'text', 300, null, 500]
      ]);
      
      const matches = ['NPV(0.1, A1:E1)', '0.1', 'A1:E1'] as RegExpMatchArray;
      const result = NPV.calculate(matches, mixedContext);
      expect(typeof result).toBe('number');
      // Should only consider -1000, 300, 500
    });
  });

  describe('IRR Function (Internal Rate of Return)', () => {
    it('should calculate IRR for profitable investment', () => {
      const irrContext = createContext([
        [-1000, 300, 400, 500, 600] // Initial investment + returns
      ]);
      
      const matches = ['IRR(A1:E1)', 'A1:E1'] as RegExpMatchArray;
      const result = IRR.calculate(matches, irrContext);
      expect(typeof result).toBe('number');
      expect(result).toBeCloseTo(0.2489, 3); // ~24.9% return
    });

    it('should handle custom guess value', () => {
      const irrContext = createContext([
        [-1000, 300, 400, 500, 600]
      ]);
      
      const matches = ['IRR(A1:E1, 0.2)', 'A1:E1', '0.2'] as RegExpMatchArray;
      const result = IRR.calculate(matches, irrContext);
      expect(typeof result).toBe('number');
      expect(result).toBeCloseTo(0.2489, 3);
    });

    it('should return NUM error for no positive and negative values', () => {
      const noMixContext = createContext([
        [100, 200, 300, 400] // All positive
      ]);
      
      const matches = ['IRR(A1:D1)', 'A1:D1'] as RegExpMatchArray;
      const result = IRR.calculate(matches, noMixContext);
      expect(result).toBe(FormulaError.NUM);
    });

    it('should return NUM error for non-convergent series', () => {
      const badContext = createContext([
        [-1, 0.1, 0.1, 0.1, 0.1] // Very low returns
      ]);
      
      const matches = ['IRR(A1:E1)', 'A1:E1'] as RegExpMatchArray;
      const result = IRR.calculate(matches, badContext);
      // May return NUM error if it doesn't converge
      expect(typeof result).toBe('number' || result === FormulaError.NUM);
    });

    it('should handle cell reference for guess', () => {
      const irrContext = createContext([
        [-1000, 300, 400, 500, 600],
        [0.15] // Guess value in B1
      ]);
      
      const matches = ['IRR(A1:E1, B1)', 'A1:E1', 'B1'] as RegExpMatchArray;
      const result = IRR.calculate(matches, irrContext);
      expect(typeof result).toBe('number');
    });
  });

  describe('PPMT Function (Principal Payment)', () => {
    it('should calculate principal payment for specific period', () => {
      // 5% rate, period 1, 60 total periods, $300k loan
      const matches = ['PPMT(0.004167, 1, 60, 300000)', '0.004167', '1', '60', '300000'] as RegExpMatchArray;
      const result = PPMT.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeLessThan(0); // Payment should be negative
      expect(result).toBeCloseTo(-4411.32, 1);
    });

    it('should handle zero interest rate', () => {
      const matches = ['PPMT(0, 1, 60, 300000)', '0', '1', '60', '300000'] as RegExpMatchArray;
      const result = PPMT.calculate(matches, mockContext);
      expect(result).toBeCloseTo(-5000, 2); // All payment is principal when rate = 0
    });

    it('should return NUM error for invalid period', () => {
      const matches = ['PPMT(0.05, 0, 60, 300000)', '0.05', '0', '60', '300000'] as RegExpMatchArray;
      const result = PPMT.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });

    it('should return NUM error for period greater than nper', () => {
      const matches = ['PPMT(0.05, 70, 60, 300000)', '0.05', '70', '60', '300000'] as RegExpMatchArray;
      const result = PPMT.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });

    it('should handle beginning-of-period payments', () => {
      const matches = ['PPMT(0.004167, 1, 60, 300000, 0, 1)', '0.004167', '1', '60', '300000', '0', '1'] as RegExpMatchArray;
      const result = PPMT.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeLessThan(0);
    });

    it('should return VALUE error for invalid inputs', () => {
      const matches = ['PPMT("invalid", 1, 60, 300000)', '"invalid"', '1', '60', '300000'] as RegExpMatchArray;
      const result = PPMT.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });
  });

  describe('IPMT Function (Interest Payment)', () => {
    it('should calculate interest payment for specific period', () => {
      // 5% rate, period 1, 60 total periods, $300k loan
      const matches = ['IPMT(0.004167, 1, 60, 300000)', '0.004167', '1', '60', '300000'] as RegExpMatchArray;
      const result = IPMT.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeLessThan(0); // Interest payment should be negative
      expect(result).toBeCloseTo(-1250.1, 1); // Interest on $300k at 0.4167% monthly
    });

    it('should return zero for zero interest rate', () => {
      const matches = ['IPMT(0, 1, 60, 300000)', '0', '1', '60', '300000'] as RegExpMatchArray;
      const result = IPMT.calculate(matches, mockContext);
      expect(result).toBe(0); // No interest when rate = 0
    });

    it('should handle beginning-of-period payments', () => {
      const matches = ['IPMT(0.004167, 1, 60, 300000, 0, 1)', '0.004167', '1', '60', '300000', '0', '1'] as RegExpMatchArray;
      const result = IPMT.calculate(matches, mockContext);
      expect(result).toBe(0); // No interest in first period for beginning payments
    });

    it('should calculate interest for later periods', () => {
      const matches = ['IPMT(0.004167, 30, 60, 300000)', '0.004167', '30', '60', '300000'] as RegExpMatchArray;
      const result = IPMT.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeLessThan(0); // Should still be negative but smaller than period 1
    });

    it('should return NUM error for invalid period', () => {
      const matches = ['IPMT(0.05, 0, 60, 300000)', '0.05', '0', '60', '300000'] as RegExpMatchArray;
      const result = IPMT.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });

    it('should return VALUE error for invalid inputs', () => {
      const matches = ['IPMT(0.05, "invalid", 60, 300000)', '0.05', '"invalid"', '60', '300000'] as RegExpMatchArray;
      const result = IPMT.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });
  });

  describe('RATE Function (Interest Rate)', () => {
    it('should calculate interest rate from loan parameters', () => {
      // 60 periods, $5659 payment, $300k present value
      const matches = ['RATE(60, -5659, 300000)', '60', '-5659', '300000'] as RegExpMatchArray;
      const result = RATE.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeCloseTo(0.004167, 4); // Should be close to monthly rate
    });

    it('should handle future value parameter', () => {
      const matches = ['RATE(60, -5659, 300000, 10000)', '60', '-5659', '300000', '10000'] as RegExpMatchArray;
      const result = RATE.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      // Rate should be different due to future value
    });

    it('should handle beginning-of-period payments', () => {
      const matches = ['RATE(60, -5636, 300000, 0, 1)', '60', '-5636', '300000', '0', '1'] as RegExpMatchArray;
      const result = RATE.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeCloseTo(0.004167, 3);
    });

    it('should handle custom guess', () => {
      const matches = ['RATE(60, -5659, 300000, 0, 0, 0.05)', '60', '-5659', '300000', '0', '0', '0.05'] as RegExpMatchArray;
      const result = RATE.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
    });

    it('should return VALUE error for invalid inputs', () => {
      const matches = ['RATE("invalid", -5659, 300000)', '"invalid"', '-5659', '300000'] as RegExpMatchArray;
      const result = RATE.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });

    it('should return NUM error for non-convergent scenarios', () => {
      // Impossible scenario: positive payment and positive present value
      const matches = ['RATE(60, 5659, 300000)', '60', '5659', '300000'] as RegExpMatchArray;
      const result = RATE.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });

    it('should return VALUE error for zero or negative periods', () => {
      const matches = ['RATE(0, -5659, 300000)', '0', '-5659', '300000'] as RegExpMatchArray;
      const result = RATE.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });
  });

  describe('Edge Cases and Integration', () => {
    it('should handle PMT-PV relationship', () => {
      // Calculate PMT, then use it to verify PV
      const pmtMatches = ['PMT(0.05, 10, 100000)', '0.05', '10', '100000'] as RegExpMatchArray;
      const pmt = PMT.calculate(pmtMatches, mockContext) as number;
      
      const pvMatches = ['PV(0.05, 10, ' + pmt + ')', '0.05', '10', pmt.toString()] as RegExpMatchArray;
      const pv = PV.calculate(pvMatches, mockContext) as number;
      
      expect(pv).toBeCloseTo(100000, 0); // Should recover original PV (absolute value)
    });

    it('should handle PPMT + IPMT = PMT relationship', () => {
      const pmtMatches = ['PMT(0.05, 10, 100000)', '0.05', '10', '100000'] as RegExpMatchArray;
      const pmt = PMT.calculate(pmtMatches, mockContext) as number;
      
      const ppmtMatches = ['PPMT(0.05, 1, 10, 100000)', '0.05', '1', '10', '100000'] as RegExpMatchArray;
      const ppmt = PPMT.calculate(ppmtMatches, mockContext) as number;
      
      const ipmtMatches = ['IPMT(0.05, 1, 10, 100000)', '0.05', '1', '10', '100000'] as RegExpMatchArray;
      const ipmt = IPMT.calculate(ipmtMatches, mockContext) as number;
      
      expect(ppmt + ipmt).toBeCloseTo(pmt, 2); // PPMT + IPMT should equal PMT
    });

    it('should handle large numbers without overflow', () => {
      const matches = ['PMT(0.001, 360, 1000000)', '0.001', '360', '1000000'] as RegExpMatchArray;
      const result = PMT.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(isFinite(result as number)).toBe(true);
    });

    it('should handle very small interest rates', () => {
      const matches = ['PMT(0.0001, 60, 100000)', '0.0001', '60', '100000'] as RegExpMatchArray;
      const result = PMT.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeCloseTo(-1671.75, 1);
    });

    it('should handle negative present values correctly', () => {
      const matches = ['PMT(0.05, 10, -100000)', '0.05', '10', '-100000'] as RegExpMatchArray;
      const result = PMT.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeGreaterThan(0); // Should be positive for negative PV
    });
  });
});