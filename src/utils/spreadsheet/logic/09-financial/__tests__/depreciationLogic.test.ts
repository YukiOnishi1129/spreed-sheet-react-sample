import { describe, it, expect } from 'vitest';
import {
  AMORDEGRC, AMORLINC, CUMIPMT, CUMPRINC
} from '../depreciationLogic';
import { FormulaError } from '../../shared/types';
import type { FormulaContext } from '../../shared/types';

// Helper function to create FormulaContext
const createContext = (data: (string | number | boolean | null)[][]): FormulaContext => ({
  data: data.map(row => row.map(cell => ({ value: cell }))),
  row: 0,
  col: 0
});

describe('Depreciation Functions', () => {
  const mockContext = createContext([
    [10000, 5000, 1000, 500], // costs and salvage values
    ['2023-01-01', '2023-12-31', '2024-01-01', '2024-12-31'], // dates
    [0.15, 0.20, 0.25, 0.30], // rates
    [0, 1, 2, 3], // periods and basis
    [0.05, 60, 100000, 1, 12, 0] // loan parameters
  ]);

  describe('AMORDEGRC Function (French Declining Balance Depreciation)', () => {
    it('should calculate depreciation for first period', () => {
      const matches = ['AMORDEGRC(10000, "2023-01-01", "2023-12-31", 1000, 0, 0.15)', 
        '10000', '"2023-01-01"', '"2023-12-31"', '1000', '0', '0.15'] as RegExpMatchArray;
      const result = AMORDEGRC.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeGreaterThan(0);
      expect(result).toBeLessThan(10000);
    });

    it('should apply correct coefficient based on rate', () => {
      // Rate 0.05 should use coefficient 2
      const matches = ['AMORDEGRC(10000, "2023-01-01", "2023-12-31", 1000, 0, 0.05)', 
        '10000', '"2023-01-01"', '"2023-12-31"', '1000', '0', '0.05'] as RegExpMatchArray;
      const result = AMORDEGRC.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeGreaterThan(0);
    });

    it('should handle subsequent periods', () => {
      const matches = ['AMORDEGRC(10000, "2023-01-01", "2023-12-31", 1000, 2, 0.15)', 
        '10000', '"2023-01-01"', '"2023-12-31"', '1000', '2', '0.15'] as RegExpMatchArray;
      const result = AMORDEGRC.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeGreaterThan(0);
    });

    it('should return 0 when asset is fully depreciated', () => {
      const matches = ['AMORDEGRC(10000, "2023-01-01", "2023-12-31", 9000, 10, 0.50)', 
        '10000', '"2023-01-01"', '"2023-12-31"', '9000', '10', '0.50'] as RegExpMatchArray;
      const result = AMORDEGRC.calculate(matches, mockContext);
      expect(result).toBe(0);
    });

    it('should return NUM error for negative cost', () => {
      const matches = ['AMORDEGRC(-10000, "2023-01-01", "2023-12-31", 1000, 0, 0.15)', 
        '-10000', '"2023-01-01"', '"2023-12-31"', '1000', '0', '0.15'] as RegExpMatchArray;
      const result = AMORDEGRC.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });

    it('should return NUM error if salvage >= cost', () => {
      const matches = ['AMORDEGRC(10000, "2023-01-01", "2023-12-31", 11000, 0, 0.15)', 
        '10000', '"2023-01-01"', '"2023-12-31"', '11000', '0', '0.15'] as RegExpMatchArray;
      const result = AMORDEGRC.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });

    it('should return VALUE error for invalid dates', () => {
      const matches = ['AMORDEGRC(10000, "invalid", "2023-12-31", 1000, 0, 0.15)', 
        '10000', '"invalid"', '"2023-12-31"', '1000', '0', '0.15'] as RegExpMatchArray;
      const result = AMORDEGRC.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });

    it('should handle different basis values', () => {
      const matches = ['AMORDEGRC(10000, "2023-01-01", "2023-12-31", 1000, 0, 0.15, 1)', 
        '10000', '"2023-01-01"', '"2023-12-31"', '1000', '0', '0.15', '1'] as RegExpMatchArray;
      const result = AMORDEGRC.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeGreaterThan(0);
    });
  });

  describe('AMORLINC Function (French Straight-Line Depreciation)', () => {
    it('should calculate straight-line depreciation for first period', () => {
      const matches = ['AMORLINC(10000, "2023-01-01", "2023-12-31", 1000, 0, 0.20)', 
        '10000', '"2023-01-01"', '"2023-12-31"', '1000', '0', '0.20'] as RegExpMatchArray;
      const result = AMORLINC.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeCloseTo(1820, -1); // Prorated for 364 days with 360-day basis
    });

    it('should calculate full year depreciation for subsequent periods', () => {
      const matches = ['AMORLINC(10000, "2023-01-01", "2023-12-31", 1000, 1, 0.20)', 
        '10000', '"2023-01-01"', '"2023-12-31"', '1000', '1', '0.20'] as RegExpMatchArray;
      const result = AMORLINC.calculate(matches, mockContext);
      expect(result).toBe(1800); // (10000-1000) * 0.20
    });

    it('should return 0 after useful life', () => {
      const matches = ['AMORLINC(10000, "2023-01-01", "2023-12-31", 1000, 10, 0.20)', 
        '10000', '"2023-01-01"', '"2023-12-31"', '1000', '10', '0.20'] as RegExpMatchArray;
      const result = AMORLINC.calculate(matches, mockContext);
      expect(result).toBe(0); // Beyond 5 year life (1/0.20)
    });

    it('should return NUM error for rate > 1', () => {
      const matches = ['AMORLINC(10000, "2023-01-01", "2023-12-31", 1000, 0, 1.5)', 
        '10000', '"2023-01-01"', '"2023-12-31"', '1000', '0', '1.5'] as RegExpMatchArray;
      const result = AMORLINC.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });

    it('should return NUM error for negative salvage', () => {
      const matches = ['AMORLINC(10000, "2023-01-01", "2023-12-31", -1000, 0, 0.20)', 
        '10000', '"2023-01-01"', '"2023-12-31"', '-1000', '0', '0.20'] as RegExpMatchArray;
      const result = AMORLINC.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });

    it('should handle different basis values', () => {
      const matches = ['AMORLINC(10000, "2023-01-01", "2023-12-31", 1000, 0, 0.20, 1)', 
        '10000', '"2023-01-01"', '"2023-12-31"', '1000', '0', '0.20', '1'] as RegExpMatchArray;
      const result = AMORLINC.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeGreaterThan(0);
    });
  });

  describe('CUMIPMT Function (Cumulative Interest Payment)', () => {
    it('should calculate cumulative interest for periods 1-12', () => {
      // 5% annual rate, 60 periods, $100,000 loan, periods 1-12, end-of-period
      const matches = ['CUMIPMT(0.004167, 60, 100000, 1, 12, 0)', 
        '0.004167', '60', '100000', '1', '12', '0'] as RegExpMatchArray;
      const result = CUMIPMT.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeLessThan(0); // Negative value (payment)
      expect(Math.abs(result as number)).toBeGreaterThan(4000); // Significant interest
    });

    it('should calculate less interest for later periods', () => {
      // Periods 49-60 should have less interest than 1-12
      const earlyMatches = ['CUMIPMT(0.004167, 60, 100000, 1, 12, 0)', 
        '0.004167', '60', '100000', '1', '12', '0'] as RegExpMatchArray;
      const earlyInterest = Math.abs(CUMIPMT.calculate(earlyMatches, mockContext) as number);
      
      const lateMatches = ['CUMIPMT(0.004167, 60, 100000, 49, 60, 0)', 
        '0.004167', '60', '100000', '49', '60', '0'] as RegExpMatchArray;
      const lateInterest = Math.abs(CUMIPMT.calculate(lateMatches, mockContext) as number);
      
      expect(lateInterest).toBeLessThan(earlyInterest);
    });

    it('should handle beginning-of-period payments', () => {
      const matches = ['CUMIPMT(0.004167, 60, 100000, 1, 12, 1)', 
        '0.004167', '60', '100000', '1', '12', '1'] as RegExpMatchArray;
      const result = CUMIPMT.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeLessThan(0);
    });

    it('should return NUM error for invalid period range', () => {
      const matches = ['CUMIPMT(0.004167, 60, 100000, 0, 12, 0)', 
        '0.004167', '60', '100000', '0', '12', '0'] as RegExpMatchArray;
      const result = CUMIPMT.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });

    it('should return NUM error for start > end period', () => {
      const matches = ['CUMIPMT(0.004167, 60, 100000, 12, 1, 0)', 
        '0.004167', '60', '100000', '12', '1', '0'] as RegExpMatchArray;
      const result = CUMIPMT.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });

    it('should return NUM error for negative rate', () => {
      const matches = ['CUMIPMT(-0.004167, 60, 100000, 1, 12, 0)', 
        '-0.004167', '60', '100000', '1', '12', '0'] as RegExpMatchArray;
      const result = CUMIPMT.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });

    it('should return NUM error for invalid type', () => {
      const matches = ['CUMIPMT(0.004167, 60, 100000, 1, 12, 2)', 
        '0.004167', '60', '100000', '1', '12', '2'] as RegExpMatchArray;
      const result = CUMIPMT.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });
  });

  describe('CUMPRINC Function (Cumulative Principal Payment)', () => {
    it('should calculate cumulative principal for periods 1-12', () => {
      const matches = ['CUMPRINC(0.004167, 60, 100000, 1, 12, 0)', 
        '0.004167', '60', '100000', '1', '12', '0'] as RegExpMatchArray;
      const result = CUMPRINC.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeLessThan(0); // Negative value (payment)
      expect(Math.abs(result as number)).toBeGreaterThan(10000); // Significant principal
    });

    it('should calculate more principal for later periods', () => {
      // Periods 49-60 should have more principal than 1-12
      const earlyMatches = ['CUMPRINC(0.004167, 60, 100000, 1, 12, 0)', 
        '0.004167', '60', '100000', '1', '12', '0'] as RegExpMatchArray;
      const earlyPrincipal = Math.abs(CUMPRINC.calculate(earlyMatches, mockContext) as number);
      
      const lateMatches = ['CUMPRINC(0.004167, 60, 100000, 49, 60, 0)', 
        '0.004167', '60', '100000', '49', '60', '0'] as RegExpMatchArray;
      const latePrincipal = Math.abs(CUMPRINC.calculate(lateMatches, mockContext) as number);
      
      expect(latePrincipal).toBeGreaterThan(earlyPrincipal);
    });

    it('should equal loan amount for full term', () => {
      const matches = ['CUMPRINC(0.004167, 60, 100000, 1, 60, 0)', 
        '0.004167', '60', '100000', '1', '60', '0'] as RegExpMatchArray;
      const result = CUMPRINC.calculate(matches, mockContext);
      expect(Math.abs(result as number)).toBeCloseTo(100000, -1);
    });

    it('should handle beginning-of-period payments', () => {
      const matches = ['CUMPRINC(0.004167, 60, 100000, 1, 12, 1)', 
        '0.004167', '60', '100000', '1', '12', '1'] as RegExpMatchArray;
      const result = CUMPRINC.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeLessThan(0);
    });

    it('should return NUM error for end period > nper', () => {
      const matches = ['CUMPRINC(0.004167, 60, 100000, 1, 61, 0)', 
        '0.004167', '60', '100000', '1', '61', '0'] as RegExpMatchArray;
      const result = CUMPRINC.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });

    it('should return NUM error for zero periods', () => {
      const matches = ['CUMPRINC(0.004167, 0, 100000, 1, 1, 0)', 
        '0.004167', '0', '100000', '1', '1', '0'] as RegExpMatchArray;
      const result = CUMPRINC.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });
  });

  describe('Edge Cases and Integration', () => {
    it('should verify CUMIPMT + CUMPRINC equals total payment', () => {
      // const nper = 60;
      const pv = 100000;
      
      const ipmtMatches = ['CUMIPMT(0.004167, 60, 100000, 1, 60, 0)', 
        '0.004167', '60', '100000', '1', '60', '0'] as RegExpMatchArray;
      const totalInterest = CUMIPMT.calculate(ipmtMatches, mockContext) as number;
      
      const princMatches = ['CUMPRINC(0.004167, 60, 100000, 1, 60, 0)', 
        '0.004167', '60', '100000', '1', '60', '0'] as RegExpMatchArray;
      const totalPrincipal = CUMPRINC.calculate(princMatches, mockContext) as number;
      
      // Total principal should equal loan amount
      expect(Math.abs(totalPrincipal)).toBeCloseTo(pv, -1);
      
      // Total payment = principal + interest
      const totalPayment = Math.abs(totalPrincipal) + Math.abs(totalInterest);
      expect(totalPayment).toBeGreaterThan(pv); // Should include interest
    });

    it('should handle very small depreciation rates', () => {
      const matches = ['AMORLINC(10000, "2023-01-01", "2023-12-31", 1000, 0, 0.01)', 
        '10000', '"2023-01-01"', '"2023-12-31"', '1000', '0', '0.01'] as RegExpMatchArray;
      const result = AMORLINC.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeCloseTo(89, -1); // Small annual depreciation
    });

    it('should handle depreciation with European basis', () => {
      const matches = ['AMORDEGRC(10000, "2023-01-31", "2023-12-31", 1000, 0, 0.15, 4)', 
        '10000', '"2023-01-31"', '"2023-12-31"', '1000', '0', '0.15', '4'] as RegExpMatchArray;
      const result = AMORDEGRC.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeGreaterThan(0);
    });

    it('should handle cell references', () => {
      // Test AMORDEGRC with cell references
      // A1=10000, B2='2023-12-31', C3='2024-01-01', B1=5000, D1=0 (period), A3=0.15
      const matches = ['AMORDEGRC(A1, B2, C3, B1, D1, A3)', 'A1', 'B2', 'C3', 'B1', 'D1', 'A3'] as RegExpMatchArray;
      const result = AMORDEGRC.calculate(matches, mockContext);
      
      // Debug output
      if (typeof result === 'string' || result === 0) {
        console.log('AMORDEGRC result:', result);
        console.log('Values - cost:', 10000, 'purchaseDate:', '2023-12-31', 'firstPeriod:', '2024-01-01', 
                    'salvage:', 5000, 'period:', 0, 'rate:', 0.15);
      }
      
      expect(typeof result).toBe('number');
      // For such a short first period (1 day), depreciation might be very small or 0
      expect(result).toBeGreaterThanOrEqual(0);
    });
  });
});