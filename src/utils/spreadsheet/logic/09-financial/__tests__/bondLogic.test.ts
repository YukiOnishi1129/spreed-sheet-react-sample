import { describe, it, expect } from 'vitest';
import {
  ACCRINT, ACCRINTM, DISC, DURATION, MDURATION, PRICE, YIELD
} from '../bondLogic';
import { FormulaError } from '../../shared/types';
import type { FormulaContext } from '../../shared/types';

// Helper function to create FormulaContext
const createContext = (data: (string | number | boolean | null)[][]): FormulaContext => ({
  data: data.map(row => row.map(cell => ({ value: cell }))),
  row: 0,
  col: 0
});

describe('Bond Functions', () => {
  const mockContext = createContext([
    ['2024-01-01', '2024-07-01', '2025-01-01', '2026-01-01'], // dates
    [0.05, 0.06, 0.07, 0.08], // rates
    [1000, 100, 95, 105], // prices/par values
    [1, 2, 4, 0], // frequencies/basis
    ['2023-01-01', '2025-06-30', '2024-12-31', '2023-07-01'] // more dates
  ]);

  describe('ACCRINT Function (Accrued Interest)', () => {
    it('should calculate accrued interest for periodic payments', () => {
      const matches = ['ACCRINT("2023-01-01", "2023-07-01", "2024-01-01", 0.05, 1000, 2)', 
        '"2023-01-01"', '"2023-07-01"', '"2024-01-01"', '0.05', '1000', '2'] as RegExpMatchArray;
      const result = ACCRINT.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeCloseTo(50, 0); // 1000 * 0.05 * (360/360)
    });

    it('should handle quarterly frequency', () => {
      const matches = ['ACCRINT("2023-01-01", "2023-04-01", "2023-07-01", 0.08, 1000, 4)', 
        '"2023-01-01"', '"2023-04-01"', '"2023-07-01"', '0.08', '1000', '4'] as RegExpMatchArray;
      const result = ACCRINT.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeGreaterThan(0);
    });

    it('should use different day count basis', () => {
      const matches = ['ACCRINT("2023-01-01", "2023-07-01", "2024-01-01", 0.05, 1000, 2, 1)', 
        '"2023-01-01"', '"2023-07-01"', '"2024-01-01"', '0.05', '1000', '2', '1'] as RegExpMatchArray;
      const result = ACCRINT.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeGreaterThan(0);
    });

    it('should return VALUE error for invalid dates', () => {
      const matches = ['ACCRINT("invalid", "2023-07-01", "2024-01-01", 0.05, 1000, 2)', 
        '"invalid"', '"2023-07-01"', '"2024-01-01"', '0.05', '1000', '2'] as RegExpMatchArray;
      const result = ACCRINT.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });

    it('should return NUM error for invalid frequency', () => {
      const matches = ['ACCRINT("2023-01-01", "2023-07-01", "2024-01-01", 0.05, 1000, 3)', 
        '"2023-01-01"', '"2023-07-01"', '"2024-01-01"', '0.05', '1000', '3'] as RegExpMatchArray;
      const result = ACCRINT.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });

    it('should return NUM error for negative rate', () => {
      const matches = ['ACCRINT("2023-01-01", "2023-07-01", "2024-01-01", -0.05, 1000, 2)', 
        '"2023-01-01"', '"2023-07-01"', '"2024-01-01"', '-0.05', '1000', '2'] as RegExpMatchArray;
      const result = ACCRINT.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });
  });

  describe('ACCRINTM Function (Accrued Interest at Maturity)', () => {
    it('should calculate accrued interest for maturity securities', () => {
      const matches = ['ACCRINTM("2023-01-01", "2024-01-01", 0.05, 1000)', 
        '"2023-01-01"', '"2024-01-01"', '0.05', '1000'] as RegExpMatchArray;
      const result = ACCRINTM.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeCloseTo(50, 0); // 1000 * 0.05 * (360/360) for one year
    });

    it('should handle different basis', () => {
      const matches = ['ACCRINTM("2023-01-01", "2024-01-01", 0.05, 1000, 1)', 
        '"2023-01-01"', '"2024-01-01"', '0.05', '1000', '1'] as RegExpMatchArray;
      const result = ACCRINTM.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeCloseTo(50, 0); // Using actual/actual basis
    });

    it('should return NUM error if issue date >= maturity date', () => {
      const matches = ['ACCRINTM("2024-01-01", "2023-01-01", 0.05, 1000)', 
        '"2024-01-01"', '"2023-01-01"', '0.05', '1000'] as RegExpMatchArray;
      const result = ACCRINTM.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });

    it('should return VALUE error for invalid dates', () => {
      const matches = ['ACCRINTM("invalid", "2024-01-01", 0.05, 1000)', 
        '"invalid"', '"2024-01-01"', '0.05', '1000'] as RegExpMatchArray;
      const result = ACCRINTM.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });

    it('should return NUM error for negative par value', () => {
      const matches = ['ACCRINTM("2023-01-01", "2024-01-01", 0.05, -1000)', 
        '"2023-01-01"', '"2024-01-01"', '0.05', '-1000'] as RegExpMatchArray;
      const result = ACCRINTM.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });
  });

  describe('DISC Function (Discount Rate)', () => {
    it('should calculate discount rate', () => {
      const matches = ['DISC("2023-01-01", "2024-01-01", 95, 100)', 
        '"2023-01-01"', '"2024-01-01"', '95', '100'] as RegExpMatchArray;
      const result = DISC.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeCloseTo(0.05, 2); // (100-95)/100 * (360/360)
    });

    it('should return NUM error for invalid price', () => {
      const matches = ['DISC("2023-01-01", "2024-01-01", 0, 100)', 
        '"2023-01-01"', '"2024-01-01"', '0', '100'] as RegExpMatchArray;
      const result = DISC.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });

    it('should return NUM error if settlement >= maturity', () => {
      const matches = ['DISC("2024-01-01", "2023-01-01", 95, 100)', 
        '"2024-01-01"', '"2023-01-01"', '95', '100'] as RegExpMatchArray;
      const result = DISC.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });

    it('should handle cell references', () => {
      // Use C1 for maturity date instead of A3 (which is a number)
      const matches = ['DISC(A1, C1, C3, B3)', 'A1', 'C1', 'C3', 'B3'] as RegExpMatchArray;
      const result = DISC.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
    });
  });

  describe('DURATION Function (Macaulay Duration)', () => {
    it('should calculate duration for bond', () => {
      const matches = ['DURATION("2023-01-01", "2028-01-01", 0.05, 0.06, 2)', 
        '"2023-01-01"', '"2028-01-01"', '0.05', '0.06', '2'] as RegExpMatchArray;
      const result = DURATION.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeGreaterThan(0);
      expect(result).toBeLessThan(5); // Should be less than maturity
    });

    it('should handle annual payments', () => {
      const matches = ['DURATION("2023-01-01", "2028-01-01", 0.05, 0.06, 1)', 
        '"2023-01-01"', '"2028-01-01"', '0.05', '0.06', '1'] as RegExpMatchArray;
      const result = DURATION.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeGreaterThan(0);
    });

    it('should return NUM error for invalid frequency', () => {
      const matches = ['DURATION("2023-01-01", "2028-01-01", 0.05, 0.06, 3)', 
        '"2023-01-01"', '"2028-01-01"', '0.05', '0.06', '3'] as RegExpMatchArray;
      const result = DURATION.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });

    it('should return NUM error for negative coupon', () => {
      const matches = ['DURATION("2023-01-01", "2028-01-01", -0.05, 0.06, 2)', 
        '"2023-01-01"', '"2028-01-01"', '-0.05', '0.06', '2'] as RegExpMatchArray;
      const result = DURATION.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });

    it('should return VALUE error for invalid dates', () => {
      const matches = ['DURATION("invalid", "2028-01-01", 0.05, 0.06, 2)', 
        '"invalid"', '"2028-01-01"', '0.05', '0.06', '2'] as RegExpMatchArray;
      const result = DURATION.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });
  });

  describe('MDURATION Function (Modified Duration)', () => {
    it('should calculate modified duration', () => {
      const matches = ['MDURATION("2023-01-01", "2028-01-01", 0.05, 0.06, 2)', 
        '"2023-01-01"', '"2028-01-01"', '0.05', '0.06', '2'] as RegExpMatchArray;
      const result = MDURATION.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeGreaterThan(0);
      // Modified duration should be less than Macaulay duration
    });

    it('should be less than Macaulay duration', () => {
      const durationMatches = ['DURATION("2023-01-01", "2028-01-01", 0.05, 0.06, 2)', 
        '"2023-01-01"', '"2028-01-01"', '0.05', '0.06', '2'] as RegExpMatchArray;
      const duration = DURATION.calculate(durationMatches, mockContext) as number;
      
      const mdurationMatches = ['MDURATION("2023-01-01", "2028-01-01", 0.05, 0.06, 2)', 
        '"2023-01-01"', '"2028-01-01"', '0.05', '0.06', '2'] as RegExpMatchArray;
      const mduration = MDURATION.calculate(mdurationMatches, mockContext) as number;
      
      expect(mduration).toBeLessThan(duration);
      expect(mduration).toBeCloseTo(duration / 1.06, 2);
    });

    it('should return NUM error for invalid parameters', () => {
      const matches = ['MDURATION("2023-01-01", "2028-01-01", -0.05, 0.06, 2)', 
        '"2023-01-01"', '"2028-01-01"', '-0.05', '0.06', '2'] as RegExpMatchArray;
      const result = MDURATION.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });
  });

  describe('PRICE Function (Bond Price)', () => {
    it('should calculate bond price', () => {
      const matches = ['PRICE("2023-01-01", "2028-01-01", 0.05, 0.06, 100, 2)', 
        '"2023-01-01"', '"2028-01-01"', '0.05', '0.06', '100', '2'] as RegExpMatchArray;
      const result = PRICE.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeGreaterThan(0);
      expect(result).toBeLessThan(110); // Reasonable price range
    });

    it('should handle zero coupon bonds', () => {
      const matches = ['PRICE("2023-01-01", "2028-01-01", 0, 0.06, 100, 2)', 
        '"2023-01-01"', '"2028-01-01"', '0', '0.06', '100', '2'] as RegExpMatchArray;
      const result = PRICE.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeLessThan(100); // Zero coupon should trade at discount
    });

    it('should return NUM error for invalid frequency', () => {
      const matches = ['PRICE("2023-01-01", "2028-01-01", 0.05, 0.06, 100, 5)', 
        '"2023-01-01"', '"2028-01-01"', '0.05', '0.06', '100', '5'] as RegExpMatchArray;
      const result = PRICE.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });

    it('should return NUM error for negative yield', () => {
      const matches = ['PRICE("2023-01-01", "2028-01-01", 0.05, -0.06, 100, 2)', 
        '"2023-01-01"', '"2028-01-01"', '0.05', '-0.06', '100', '2'] as RegExpMatchArray;
      const result = PRICE.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });

    it('should return VALUE error for invalid dates', () => {
      const matches = ['PRICE("invalid", "2028-01-01", 0.05, 0.06, 100, 2)', 
        '"invalid"', '"2028-01-01"', '0.05', '0.06', '100', '2'] as RegExpMatchArray;
      const result = PRICE.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });
  });

  describe('YIELD Function (Bond Yield)', () => {
    it('should calculate bond yield', () => {
      const matches = ['YIELD("2023-01-01", "2028-01-01", 0.05, 95, 100, 2)', 
        '"2023-01-01"', '"2028-01-01"', '0.05', '95', '100', '2'] as RegExpMatchArray;
      const result = YIELD.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeGreaterThan(0.05); // Should be higher than coupon if price < redemption
    });

    it('should converge for various price levels', () => {
      // Premium bond
      const premiumMatches = ['YIELD("2023-01-01", "2028-01-01", 0.06, 105, 100, 2)', 
        '"2023-01-01"', '"2028-01-01"', '0.06', '105', '100', '2'] as RegExpMatchArray;
      const premiumYield = YIELD.calculate(premiumMatches, mockContext);
      expect(typeof premiumYield).toBe('number');
      expect(premiumYield).toBeLessThan(0.06); // Yield < coupon for premium bond
      
      // Discount bond
      const discountMatches = ['YIELD("2023-01-01", "2028-01-01", 0.04, 95, 100, 2)', 
        '"2023-01-01"', '"2028-01-01"', '0.04', '95', '100', '2'] as RegExpMatchArray;
      const discountYield = YIELD.calculate(discountMatches, mockContext);
      expect(typeof discountYield).toBe('number');
      expect(discountYield).toBeGreaterThan(0.04); // Yield > coupon for discount bond
    });

    it('should return NUM error for zero price', () => {
      const matches = ['YIELD("2023-01-01", "2028-01-01", 0.05, 0, 100, 2)', 
        '"2023-01-01"', '"2028-01-01"', '0.05', '0', '100', '2'] as RegExpMatchArray;
      const result = YIELD.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });

    it('should return NUM error for invalid frequency', () => {
      const matches = ['YIELD("2023-01-01", "2028-01-01", 0.05, 95, 100, 3)', 
        '"2023-01-01"', '"2028-01-01"', '0.05', '95', '100', '3'] as RegExpMatchArray;
      const result = YIELD.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });

    it('should return VALUE error for invalid dates', () => {
      const matches = ['YIELD("invalid", "2028-01-01", 0.05, 95, 100, 2)', 
        '"invalid"', '"2028-01-01"', '0.05', '95', '100', '2'] as RegExpMatchArray;
      const result = YIELD.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });
  });

  describe('Edge Cases', () => {
    it('should handle bonds close to maturity', () => {
      const matches = ['PRICE("2023-12-01", "2024-01-01", 0.05, 0.06, 100, 2)', 
        '"2023-12-01"', '"2024-01-01"', '0.05', '0.06', '100', '2'] as RegExpMatchArray;
      const result = PRICE.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeCloseTo(99.92, 1); // Should be close to redemption
    });

    it('should handle European basis (30/360E)', () => {
      const matches = ['ACCRINT("2023-01-31", "2023-07-31", "2024-01-31", 0.05, 1000, 2, 4)', 
        '"2023-01-31"', '"2023-07-31"', '"2024-01-31"', '0.05', '1000', '2', '4'] as RegExpMatchArray;
      const result = ACCRINT.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeGreaterThan(0);
    });

    it('should handle actual/365 basis', () => {
      const matches = ['ACCRINTM("2023-01-01", "2024-01-01", 0.05, 1000, 3)', 
        '"2023-01-01"', '"2024-01-01"', '0.05', '1000', '3'] as RegExpMatchArray;
      const result = ACCRINTM.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeCloseTo(50, 0); // Using actual/365 basis
    });

    it('should handle very long maturity bonds', () => {
      const matches = ['DURATION("2023-01-01", "2053-01-01", 0.03, 0.04, 2)', 
        '"2023-01-01"', '"2053-01-01"', '0.03', '0.04', '2'] as RegExpMatchArray;
      const result = DURATION.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeGreaterThan(10); // Long duration
      expect(result).toBeLessThan(30); // But still reasonable
    });
  });
});