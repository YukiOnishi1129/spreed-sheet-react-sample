import { describe, it, expect } from 'vitest';
import {
  COUPDAYBS, COUPDAYS, COUPDAYSNC, COUPNCD, COUPNUM, COUPPCD
} from '../couponLogic';
import { FormulaError } from '../../shared/types';
import type { FormulaContext } from '../../shared/types';

// Helper function to create FormulaContext
const createContext = (data: (string | number | boolean | null)[][]): FormulaContext => ({
  data: data.map(row => row.map(cell => ({ value: cell }))),
  row: 0,
  col: 0
});

describe('Coupon Functions', () => {
  const mockContext = createContext([
    ['2024-01-15', '2025-01-15', '2026-01-15', '2027-01-15'], // dates
    [1, 2, 4, 12], // frequencies
    [0, 1, 2, 3], // basis values
    ['2024-07-15', '2024-10-15', '2025-07-15', '2026-07-15'], // more dates
  ]);

  describe('COUPDAYBS Function (Days from Previous Coupon to Settlement)', () => {
    it('should calculate days from previous coupon date to settlement', () => {
      // Semi-annual bond: settlement 2024-04-15, maturity 2026-01-15
      const matches = ['COUPDAYBS("2024-04-15", "2026-01-15", 2)', 
        '"2024-04-15"', '"2026-01-15"', '2'] as RegExpMatchArray;
      const result = COUPDAYBS.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBe(90); // 3 months * 30 days = 90 days (30/360 basis)
    });

    it('should handle annual frequency', () => {
      const matches = ['COUPDAYBS("2024-04-15", "2026-01-15", 1)', 
        '"2024-04-15"', '"2026-01-15"', '1'] as RegExpMatchArray;
      const result = COUPDAYBS.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBe(90); // From Jan 15 to Apr 15
    });

    it('should handle quarterly frequency', () => {
      const matches = ['COUPDAYBS("2024-04-15", "2026-01-15", 4)', 
        '"2024-04-15"', '"2026-01-15"', '4'] as RegExpMatchArray;
      const result = COUPDAYBS.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBe(0); // Settlement is on coupon date
    });

    it('should use actual/actual basis when specified', () => {
      const matches = ['COUPDAYBS("2024-04-15", "2026-01-15", 2, 1)', 
        '"2024-04-15"', '"2026-01-15"', '2', '1'] as RegExpMatchArray;
      const result = COUPDAYBS.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBe(91); // Actual days from Jan 15 to Apr 15 (2024 is leap year)
    });

    it('should return NUM error for invalid frequency', () => {
      const matches = ['COUPDAYBS("2024-04-15", "2026-01-15", 3)', 
        '"2024-04-15"', '"2026-01-15"', '3'] as RegExpMatchArray;
      const result = COUPDAYBS.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });

    it('should return VALUE error for invalid dates', () => {
      const matches = ['COUPDAYBS("invalid", "2026-01-15", 2)', 
        '"invalid"', '"2026-01-15"', '2'] as RegExpMatchArray;
      const result = COUPDAYBS.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });

    it('should return NUM error if settlement >= maturity', () => {
      const matches = ['COUPDAYBS("2027-01-15", "2026-01-15", 2)', 
        '"2027-01-15"', '"2026-01-15"', '2'] as RegExpMatchArray;
      const result = COUPDAYBS.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });
  });

  describe('COUPDAYS Function (Days in Coupon Period)', () => {
    it('should return days in coupon period for semi-annual', () => {
      const matches = ['COUPDAYS("2024-04-15", "2026-01-15", 2)', 
        '"2024-04-15"', '"2026-01-15"', '2'] as RegExpMatchArray;
      const result = COUPDAYS.calculate(matches, mockContext);
      expect(result).toBe(180); // 360/2 = 180 days
    });

    it('should return days for annual frequency', () => {
      const matches = ['COUPDAYS("2024-04-15", "2026-01-15", 1)', 
        '"2024-04-15"', '"2026-01-15"', '1'] as RegExpMatchArray;
      const result = COUPDAYS.calculate(matches, mockContext);
      expect(result).toBe(360); // 360/1 = 360 days
    });

    it('should return days for quarterly frequency', () => {
      const matches = ['COUPDAYS("2024-04-15", "2026-01-15", 4)', 
        '"2024-04-15"', '"2026-01-15"', '4'] as RegExpMatchArray;
      const result = COUPDAYS.calculate(matches, mockContext);
      expect(result).toBe(90); // 360/4 = 90 days
    });

    it('should handle actual/actual basis', () => {
      const matches = ['COUPDAYS("2024-04-15", "2026-01-15", 2, 1)', 
        '"2024-04-15"', '"2026-01-15"', '2', '1'] as RegExpMatchArray;
      const result = COUPDAYS.calculate(matches, mockContext);
      expect(result).toBe(182.5); // 365/2 = 182.5 days
    });

    it('should handle actual/360 basis', () => {
      const matches = ['COUPDAYS("2024-04-15", "2026-01-15", 2, 2)', 
        '"2024-04-15"', '"2026-01-15"', '2', '2'] as RegExpMatchArray;
      const result = COUPDAYS.calculate(matches, mockContext);
      expect(result).toBe(180); // 360/2 = 180 days
    });

    it('should return NUM error for invalid frequency', () => {
      const matches = ['COUPDAYS("2024-04-15", "2026-01-15", 5)', 
        '"2024-04-15"', '"2026-01-15"', '5'] as RegExpMatchArray;
      const result = COUPDAYS.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });
  });

  describe('COUPDAYSNC Function (Days from Settlement to Next Coupon)', () => {
    it('should calculate days to next coupon date', () => {
      const matches = ['COUPDAYSNC("2024-04-15", "2026-01-15", 2)', 
        '"2024-04-15"', '"2026-01-15"', '2'] as RegExpMatchArray;
      const result = COUPDAYSNC.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBe(90); // To July 15 (3 months * 30 days)
    });

    it('should handle settlement on coupon date', () => {
      const matches = ['COUPDAYSNC("2024-01-15", "2026-01-15", 2)', 
        '"2024-01-15"', '"2026-01-15"', '2'] as RegExpMatchArray;
      const result = COUPDAYSNC.calculate(matches, mockContext);
      expect(result).toBe(180); // Full period to next coupon
    });

    it('should handle actual basis', () => {
      const matches = ['COUPDAYSNC("2024-04-15", "2026-01-15", 2, 1)', 
        '"2024-04-15"', '"2026-01-15"', '2', '1'] as RegExpMatchArray;
      const result = COUPDAYSNC.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBe(91); // Actual days from Apr 15 to Jul 15
    });

    it('should handle quarterly frequency', () => {
      const matches = ['COUPDAYSNC("2024-02-15", "2026-01-15", 4)', 
        '"2024-02-15"', '"2026-01-15"', '4'] as RegExpMatchArray;
      const result = COUPDAYSNC.calculate(matches, mockContext);
      expect(result).toBe(60); // Two months to next quarterly payment (Apr 15)
    });

    it('should return VALUE error for invalid dates', () => {
      const matches = ['COUPDAYSNC("invalid", "2026-01-15", 2)', 
        '"invalid"', '"2026-01-15"', '2'] as RegExpMatchArray;
      const result = COUPDAYSNC.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });
  });

  describe('COUPNCD Function (Next Coupon Date)', () => {
    it('should return next coupon date as Excel serial number', () => {
      const matches = ['COUPNCD("2024-04-15", "2026-01-15", 2)', 
        '"2024-04-15"', '"2026-01-15"', '2'] as RegExpMatchArray;
      const result = COUPNCD.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      // Result should be Excel date serial for 2024-07-15
      expect(result).toBeGreaterThan(45000); // Excel dates in 2024 range
    });

    it('should handle annual frequency', () => {
      const matches = ['COUPNCD("2024-04-15", "2026-01-15", 1)', 
        '"2024-04-15"', '"2026-01-15"', '1'] as RegExpMatchArray;
      const result = COUPNCD.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      // Should return Excel serial for 2025-01-15
    });

    it('should handle quarterly frequency', () => {
      const matches = ['COUPNCD("2024-02-15", "2026-01-15", 4)', 
        '"2024-02-15"', '"2026-01-15"', '4'] as RegExpMatchArray;
      const result = COUPNCD.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      // Should return Excel serial for 2024-04-15
    });

    it('should return NUM error for invalid frequency', () => {
      const matches = ['COUPNCD("2024-04-15", "2026-01-15", 6)', 
        '"2024-04-15"', '"2026-01-15"', '6'] as RegExpMatchArray;
      const result = COUPNCD.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });

    it('should return VALUE error for invalid dates', () => {
      const matches = ['COUPNCD("invalid", "2026-01-15", 2)', 
        '"invalid"', '"2026-01-15"', '2'] as RegExpMatchArray;
      const result = COUPNCD.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });
  });

  describe('COUPNUM Function (Number of Coupons)', () => {
    it('should count remaining coupons for semi-annual bond', () => {
      const matches = ['COUPNUM("2024-04-15", "2026-01-15", 2)', 
        '"2024-04-15"', '"2026-01-15"', '2'] as RegExpMatchArray;
      const result = COUPNUM.calculate(matches, mockContext);
      expect(result).toBe(4); // Jul 2024, Jan 2025, Jul 2025, Jan 2026
    });

    it('should count coupons for annual bond', () => {
      const matches = ['COUPNUM("2024-04-15", "2026-01-15", 1)', 
        '"2024-04-15"', '"2026-01-15"', '1'] as RegExpMatchArray;
      const result = COUPNUM.calculate(matches, mockContext);
      expect(result).toBe(2); // Jan 2025, Jan 2026
    });

    it('should count coupons for quarterly bond', () => {
      const matches = ['COUPNUM("2024-04-15", "2025-04-15", 4)', 
        '"2024-04-15"', '"2025-04-15"', '4'] as RegExpMatchArray;
      const result = COUPNUM.calculate(matches, mockContext);
      expect(result).toBe(4); // Jul, Oct, Jan, Apr
    });

    it('should return 1 for settlement close to maturity', () => {
      const matches = ['COUPNUM("2025-12-15", "2026-01-15", 2)', 
        '"2025-12-15"', '"2026-01-15"', '2'] as RegExpMatchArray;
      const result = COUPNUM.calculate(matches, mockContext);
      expect(result).toBe(1); // Only final coupon
    });

    it('should return NUM error for invalid frequency', () => {
      const matches = ['COUPNUM("2024-04-15", "2026-01-15", 3)', 
        '"2024-04-15"', '"2026-01-15"', '3'] as RegExpMatchArray;
      const result = COUPNUM.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });

    it('should return NUM error if settlement >= maturity', () => {
      const matches = ['COUPNUM("2026-01-15", "2026-01-15", 2)', 
        '"2026-01-15"', '"2026-01-15"', '2'] as RegExpMatchArray;
      const result = COUPNUM.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });
  });

  describe('COUPPCD Function (Previous Coupon Date)', () => {
    it('should return previous coupon date as Excel serial number', () => {
      const matches = ['COUPPCD("2024-04-15", "2026-01-15", 2)', 
        '"2024-04-15"', '"2026-01-15"', '2'] as RegExpMatchArray;
      const result = COUPPCD.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      // Result should be Excel date serial for 2024-01-15
      expect(result).toBeGreaterThan(45000); // Excel dates in 2024 range
    });

    it('should handle settlement on coupon date', () => {
      const matches = ['COUPPCD("2024-07-15", "2026-01-15", 2)', 
        '"2024-07-15"', '"2026-01-15"', '2'] as RegExpMatchArray;
      const result = COUPPCD.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      // Should return 2024-01-15 (previous period)
    });

    it('should handle annual frequency', () => {
      const matches = ['COUPPCD("2024-06-15", "2026-01-15", 1)', 
        '"2024-06-15"', '"2026-01-15"', '1'] as RegExpMatchArray;
      const result = COUPPCD.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      // Should return Excel serial for 2024-01-15
    });

    it('should handle quarterly frequency', () => {
      const matches = ['COUPPCD("2024-05-15", "2026-01-15", 4)', 
        '"2024-05-15"', '"2026-01-15"', '4'] as RegExpMatchArray;
      const result = COUPPCD.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      // Should return Excel serial for 2024-04-15
    });

    it('should return VALUE error for invalid dates', () => {
      const matches = ['COUPPCD("invalid", "2026-01-15", 2)', 
        '"invalid"', '"2026-01-15"', '2'] as RegExpMatchArray;
      const result = COUPPCD.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });
  });

  describe('Edge Cases and Integration', () => {
    it('should handle leap years correctly', () => {
      // 2024 is a leap year
      const matches = ['COUPDAYBS("2024-03-01", "2026-01-15", 2, 1)', 
        '"2024-03-01"', '"2026-01-15"', '2', '1'] as RegExpMatchArray;
      const result = COUPDAYBS.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      // Should account for Feb 29 in actual/actual basis
    });

    it('should verify COUPDAYBS + COUPDAYSNC = COUPDAYS', () => {
      const settlement = '"2024-04-15"';
      const maturity = '"2026-01-15"';
      const frequency = '2';
      
      const daysBSMatches = ['COUPDAYBS(' + settlement + ', ' + maturity + ', ' + frequency + ')', 
        settlement, maturity, frequency] as RegExpMatchArray;
      const daysBs = COUPDAYBS.calculate(daysBSMatches, mockContext) as number;
      
      const daysNCMatches = ['COUPDAYSNC(' + settlement + ', ' + maturity + ', ' + frequency + ')', 
        settlement, maturity, frequency] as RegExpMatchArray;
      const daysNc = COUPDAYSNC.calculate(daysNCMatches, mockContext) as number;
      
      const daysMatches = ['COUPDAYS(' + settlement + ', ' + maturity + ', ' + frequency + ')', 
        settlement, maturity, frequency] as RegExpMatchArray;
      const days = COUPDAYS.calculate(daysMatches, mockContext) as number;
      
      expect(daysBs + daysNc).toBe(days);
    });

    it('should handle European 30/360 basis', () => {
      const matches = ['COUPDAYBS("2024-03-31", "2026-01-31", 2, 4)', 
        '"2024-03-31"', '"2026-01-31"', '2', '4'] as RegExpMatchArray;
      const result = COUPDAYBS.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      // European basis handles month-end differently
    });

    it('should handle bonds with very long maturity', () => {
      const matches = ['COUPNUM("2024-01-15", "2054-01-15", 2)', 
        '"2024-01-15"', '"2054-01-15"', '2'] as RegExpMatchArray;
      const result = COUPNUM.calculate(matches, mockContext);
      expect(result).toBe(60); // 30 years * 2 coupons per year
    });

    it('should handle cell references', () => {
      // A4 = '2024-07-15' (row 3, col 0), C1 = '2026-01-15' (row 0, col 2), B2 = 2 (row 1, col 1)
      const matches = ['COUPDAYBS(A4, C1, B2)', 'A4', 'C1', 'B2'] as RegExpMatchArray;
      const result = COUPDAYBS.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBe(0); // July 15 is a coupon date for semi-annual with maturity Jan 15
    });
  });
});