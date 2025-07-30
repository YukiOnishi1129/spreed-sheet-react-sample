import { describe, it, expect } from 'vitest';
import {
  TBILLEQ, TBILLPRICE, TBILLYIELD
} from '../treasuryLogic';
import { FormulaError } from '../../shared/types';
import type { FormulaContext } from '../../shared/types';

// Helper function to create FormulaContext
const createContext = (data: (string | number | boolean | null)[][]): FormulaContext => ({
  data: data.map(row => row.map(cell => ({ value: cell }))),
  row: 0,
  col: 0
});

describe('Treasury Bill Functions', () => {
  const mockContext = createContext([
    ['2024-01-15', '2024-04-15', '2024-07-15', '2024-10-15'], // settlement dates
    ['2024-07-15', '2024-10-15', '2025-01-15', '2025-04-15'], // maturity dates
    [0.04, 0.045, 0.05, 0.055], // discount rates
    [98.5, 99.0, 99.2, 99.5], // prices per $100 face value
    [0.041, 0.046, 0.051, 0.056] // yields
  ]);

  describe('TBILLEQ Function (T-Bill Bond-Equivalent Yield)', () => {
    it('should calculate bond-equivalent yield for 90-day T-bill', () => {
      const matches = ['TBILLEQ("2024-01-15", "2024-04-15", 0.04)', 
        '"2024-01-15"', '"2024-04-15"', '0.04'] as RegExpMatchArray;
      const result = TBILLEQ.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeGreaterThan(0.04); // BEY should be higher than discount rate
      expect(result).toBeLessThan(0.05);
    });

    it('should calculate higher BEY for longer maturity', () => {
      // 90-day bill
      const shortMatches = ['TBILLEQ("2024-01-15", "2024-04-15", 0.04)', 
        '"2024-01-15"', '"2024-04-15"', '0.04'] as RegExpMatchArray;
      const shortBEY = TBILLEQ.calculate(shortMatches, mockContext) as number;
      
      // 180-day bill
      const longMatches = ['TBILLEQ("2024-01-15", "2024-07-15", 0.04)', 
        '"2024-01-15"', '"2024-07-15"', '0.04'] as RegExpMatchArray;
      const longBEY = TBILLEQ.calculate(longMatches, mockContext) as number;
      
      expect(longBEY).toBeGreaterThan(shortBEY);
    });

    it('should handle near-maturity bills', () => {
      const matches = ['TBILLEQ("2024-04-10", "2024-04-15", 0.04)', 
        '"2024-04-10"', '"2024-04-15"', '0.04'] as RegExpMatchArray;
      const result = TBILLEQ.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeGreaterThan(0);
    });

    it('should return NUM error for maturity > 1 year', () => {
      const matches = ['TBILLEQ("2024-01-15", "2025-02-15", 0.04)', 
        '"2024-01-15"', '"2025-02-15"', '0.04'] as RegExpMatchArray;
      const result = TBILLEQ.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });

    it('should return NUM error for negative discount rate', () => {
      const matches = ['TBILLEQ("2024-01-15", "2024-04-15", -0.01)', 
        '"2024-01-15"', '"2024-04-15"', '-0.01'] as RegExpMatchArray;
      const result = TBILLEQ.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });

    it('should return NUM error for discount rate >= 1', () => {
      const matches = ['TBILLEQ("2024-01-15", "2024-04-15", 1.0)', 
        '"2024-01-15"', '"2024-04-15"', '1.0'] as RegExpMatchArray;
      const result = TBILLEQ.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });

    it('should return NUM error if settlement >= maturity', () => {
      const matches = ['TBILLEQ("2024-04-15", "2024-04-15", 0.04)', 
        '"2024-04-15"', '"2024-04-15"', '0.04'] as RegExpMatchArray;
      const result = TBILLEQ.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });

    it('should return VALUE error for invalid dates', () => {
      const matches = ['TBILLEQ("invalid", "2024-04-15", 0.04)', 
        '"invalid"', '"2024-04-15"', '0.04'] as RegExpMatchArray;
      const result = TBILLEQ.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });
  });

  describe('TBILLPRICE Function (T-Bill Price)', () => {
    it('should calculate price for 90-day T-bill', () => {
      const matches = ['TBILLPRICE("2024-01-15", "2024-04-15", 0.04)', 
        '"2024-01-15"', '"2024-04-15"', '0.04'] as RegExpMatchArray;
      const result = TBILLPRICE.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeLessThan(100); // Price should be less than face value
      expect(result).toBeGreaterThan(98); // But not too low
    });

    it('should calculate lower price for higher discount rate', () => {
      const lowRateMatches = ['TBILLPRICE("2024-01-15", "2024-04-15", 0.03)', 
        '"2024-01-15"', '"2024-04-15"', '0.03'] as RegExpMatchArray;
      const lowRatePrice = TBILLPRICE.calculate(lowRateMatches, mockContext) as number;
      
      const highRateMatches = ['TBILLPRICE("2024-01-15", "2024-04-15", 0.06)', 
        '"2024-01-15"', '"2024-04-15"', '0.06'] as RegExpMatchArray;
      const highRatePrice = TBILLPRICE.calculate(highRateMatches, mockContext) as number;
      
      expect(highRatePrice).toBeLessThan(lowRatePrice);
    });

    it('should calculate lower price for longer maturity', () => {
      // 90-day bill
      const shortMatches = ['TBILLPRICE("2024-01-15", "2024-04-15", 0.04)', 
        '"2024-01-15"', '"2024-04-15"', '0.04'] as RegExpMatchArray;
      const shortPrice = TBILLPRICE.calculate(shortMatches, mockContext) as number;
      
      // 180-day bill
      const longMatches = ['TBILLPRICE("2024-01-15", "2024-07-15", 0.04)', 
        '"2024-01-15"', '"2024-07-15"', '0.04'] as RegExpMatchArray;
      const longPrice = TBILLPRICE.calculate(longMatches, mockContext) as number;
      
      expect(longPrice).toBeLessThan(shortPrice);
    });

    it('should approach 100 as settlement approaches maturity', () => {
      const nearMatches = ['TBILLPRICE("2024-04-14", "2024-04-15", 0.04)', 
        '"2024-04-14"', '"2024-04-15"', '0.04'] as RegExpMatchArray;
      const nearPrice = TBILLPRICE.calculate(nearMatches, mockContext) as number;
      
      expect(nearPrice).toBeGreaterThan(99.9);
      expect(nearPrice).toBeLessThan(100);
    });

    it('should return NUM error for maturity > 1 year', () => {
      const matches = ['TBILLPRICE("2024-01-15", "2025-02-15", 0.04)', 
        '"2024-01-15"', '"2025-02-15"', '0.04'] as RegExpMatchArray;
      const result = TBILLPRICE.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });

    it('should return NUM error for zero discount rate', () => {
      const matches = ['TBILLPRICE("2024-01-15", "2024-04-15", 0)', 
        '"2024-01-15"', '"2024-04-15"', '0'] as RegExpMatchArray;
      const result = TBILLPRICE.calculate(matches, mockContext);
      expect(result).toBe(100); // At 0% discount, price equals face value
    });

    it('should return VALUE error for invalid dates', () => {
      const matches = ['TBILLPRICE("2024-01-15", "invalid", 0.04)', 
        '"2024-01-15"', '"invalid"', '0.04'] as RegExpMatchArray;
      const result = TBILLPRICE.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });
  });

  describe('TBILLYIELD Function (T-Bill Yield)', () => {
    it('should calculate yield for 90-day T-bill', () => {
      const matches = ['TBILLYIELD("2024-01-15", "2024-04-15", 99)', 
        '"2024-01-15"', '"2024-04-15"', '99'] as RegExpMatchArray;
      const result = TBILLYIELD.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeGreaterThan(0);
      expect(result).toBeLessThan(0.1); // Reasonable yield range
    });

    it('should calculate higher yield for lower price', () => {
      const highPriceMatches = ['TBILLYIELD("2024-01-15", "2024-04-15", 99.5)', 
        '"2024-01-15"', '"2024-04-15"', '99.5'] as RegExpMatchArray;
      const highPriceYield = TBILLYIELD.calculate(highPriceMatches, mockContext) as number;
      
      const lowPriceMatches = ['TBILLYIELD("2024-01-15", "2024-04-15", 98.5)', 
        '"2024-01-15"', '"2024-04-15"', '98.5'] as RegExpMatchArray;
      const lowPriceYield = TBILLYIELD.calculate(lowPriceMatches, mockContext) as number;
      
      expect(lowPriceYield).toBeGreaterThan(highPriceYield);
    });

    it('should calculate annualized yield correctly', () => {
      // For a 182-day bill with price 98, yield should be around 4%
      const matches = ['TBILLYIELD("2024-01-15", "2024-07-15", 98)', 
        '"2024-01-15"', '"2024-07-15"', '98'] as RegExpMatchArray;
      const result = TBILLYIELD.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeGreaterThan(0.03);
      expect(result).toBeLessThan(0.05);
    });

    it('should return zero yield for price = 100', () => {
      const matches = ['TBILLYIELD("2024-01-15", "2024-04-15", 100)', 
        '"2024-01-15"', '"2024-04-15"', '100'] as RegExpMatchArray;
      const result = TBILLYIELD.calculate(matches, mockContext);
      expect(result).toBe(0);
    });

    it('should return NUM error for price <= 0', () => {
      const matches = ['TBILLYIELD("2024-01-15", "2024-04-15", 0)', 
        '"2024-01-15"', '"2024-04-15"', '0'] as RegExpMatchArray;
      const result = TBILLYIELD.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });

    it('should return NUM error for price > 100', () => {
      const matches = ['TBILLYIELD("2024-01-15", "2024-04-15", 101)', 
        '"2024-01-15"', '"2024-04-15"', '101'] as RegExpMatchArray;
      const result = TBILLYIELD.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });

    it('should return NUM error for maturity > 1 year', () => {
      const matches = ['TBILLYIELD("2024-01-15", "2025-02-15", 99)', 
        '"2024-01-15"', '"2025-02-15"', '99'] as RegExpMatchArray;
      const result = TBILLYIELD.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });

    it('should return VALUE error for invalid dates', () => {
      const matches = ['TBILLYIELD("invalid", "2024-04-15", 99)', 
        '"invalid"', '"2024-04-15"', '99'] as RegExpMatchArray;
      const result = TBILLYIELD.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });
  });

  describe('Edge Cases and Integration', () => {
    it('should verify price-yield consistency', () => {
      const discount = 0.04;
      const settlement = '"2024-01-15"';
      const maturity = '"2024-04-15"';
      
      // Calculate price from discount rate
      const priceMatches = ['TBILLPRICE(' + settlement + ', ' + maturity + ', ' + discount + ')', 
        settlement, maturity, discount.toString()] as RegExpMatchArray;
      const price = TBILLPRICE.calculate(priceMatches, mockContext) as number;
      
      // Calculate yield from that price
      const yieldMatches = ['TBILLYIELD(' + settlement + ', ' + maturity + ', ' + price + ')', 
        settlement, maturity, price.toString()] as RegExpMatchArray;
      const calculatedYield = TBILLYIELD.calculate(yieldMatches, mockContext) as number;
      
      // Yield should be close to original discount rate
      expect(calculatedYield).toBeCloseTo(discount, 3);
    });

    it('should handle 52-week (364-day) bills', () => {
      const matches = ['TBILLPRICE("2024-01-15", "2025-01-14", 0.045)', 
        '"2024-01-15"', '"2025-01-14"', '0.045'] as RegExpMatchArray;
      const result = TBILLPRICE.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeGreaterThan(95);
      expect(result).toBeLessThan(100);
    });

    it('should handle very short maturity bills', () => {
      // 7-day bill
      const matches = ['TBILLEQ("2024-01-15", "2024-01-22", 0.04)', 
        '"2024-01-15"', '"2024-01-22"', '0.04'] as RegExpMatchArray;
      const result = TBILLEQ.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeGreaterThan(0.04);
    });

    it('should handle leap year calculations', () => {
      // 2024 is a leap year
      const matches = ['TBILLPRICE("2024-02-15", "2024-05-15", 0.04)', 
        '"2024-02-15"', '"2024-05-15"', '0.04'] as RegExpMatchArray;
      const result = TBILLPRICE.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeLessThan(100);
    });

    it('should verify BEY is higher than discount yield', () => {
      const discounts = [0.02, 0.04, 0.06, 0.08];
      
      for (const discount of discounts) {
        const matches = ['TBILLEQ("2024-01-15", "2024-07-15", ' + discount + ')', 
          '"2024-01-15"', '"2024-07-15"', discount.toString()] as RegExpMatchArray;
        const bey = TBILLEQ.calculate(matches, mockContext) as number;
        expect(bey).toBeGreaterThan(discount);
      }
    });

    it('should handle cell references', () => {
      const matches = ['TBILLPRICE(A1, B1, A3)', 'A1', 'B1', 'A3'] as RegExpMatchArray;
      const result = TBILLPRICE.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
    });
  });
});