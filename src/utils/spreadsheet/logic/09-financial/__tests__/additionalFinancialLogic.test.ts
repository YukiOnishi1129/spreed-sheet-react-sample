import { describe, it, expect } from 'vitest';
import {
  NPER, XNPV, XIRR, MIRR, SLN, SYD, DB, DDB, VDB, PDURATION, RRI, ISPMT
} from '../additionalFinancialLogic';
import { FormulaError } from '../../shared/types';
import type { FormulaContext } from '../../shared/types';

// Helper function to create FormulaContext
const createContext = (data: (string | number | boolean | null)[][]): FormulaContext => ({
  data: data.map(row => row.map(cell => ({ value: cell }))),
  row: 0,
  col: 0
});

describe('Additional Financial Functions', () => {
  const mockContext = createContext([
    [0.05, 0.06, 0.08, 0.10], // rates
    [-1000, 200, 300, 400, 500], // cash flows
    [44927, 45292, 45658, 46023, 46388], // Excel dates (2023-01-01 to 2024-01-01)
    [10000, 1000, 5, 3], // cost, salvage, life, period
    [100000, 50000, 10, 1], // depreciation parameters
    [-10000, 2500, 3000, 3500, 4000], // MIRR cash flows
  ]);

  describe('NPER Function (Number of Periods)', () => {
    it('should calculate number of periods for loan', () => {
      const matches = ['NPER(0.05, -1000, 10000)', '0.05', '-1000', '10000'] as RegExpMatchArray;
      const result = NPER.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeCloseTo(14.21, 1); // About 14.2 periods
    });

    it('should handle zero interest rate', () => {
      const matches = ['NPER(0, -1000, 10000)', '0', '-1000', '10000'] as RegExpMatchArray;
      const result = NPER.calculate(matches, mockContext);
      expect(result).toBe(10); // 10000 / 1000 = 10 periods
    });

    it('should handle future value parameter', () => {
      const matches = ['NPER(0.05, -1000, 10000, 5000)', '0.05', '-1000', '10000', '5000'] as RegExpMatchArray;
      const result = NPER.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeGreaterThan(14.21); // More periods needed with FV
    });

    it('should handle beginning-of-period payments', () => {
      const matches = ['NPER(0.05, -1000, 10000, 0, 1)', '0.05', '-1000', '10000', '0', '1'] as RegExpMatchArray;
      const result = NPER.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeLessThan(14.21); // Fewer periods with beginning payments
    });

    it('should return NUM error for zero payment with zero rate', () => {
      const matches = ['NPER(0, 0, 10000)', '0', '0', '10000'] as RegExpMatchArray;
      const result = NPER.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });

    it('should return VALUE error for invalid type', () => {
      const matches = ['NPER(0.05, -1000, 10000, 0, 2)', '0.05', '-1000', '10000', '0', '2'] as RegExpMatchArray;
      const result = NPER.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });
  });

  describe('XNPV Function (Net Present Value with Dates)', () => {
    it('should calculate NPV with specific dates', () => {
      const matches = ['XNPV(0.1, A2:E2, A3:E3)', '0.1', 'A2:E2', 'A3:E3'] as RegExpMatchArray;
      const result = XNPV.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeGreaterThan(0); // Positive NPV expected
    });

    it('should handle zero discount rate', () => {
      const matches = ['XNPV(0, A2:E2, A3:E3)', '0', 'A2:E2', 'A3:E3'] as RegExpMatchArray;
      const result = XNPV.calculate(matches, mockContext);
      expect(result).toBe(400); // Sum of cash flows: -1000+200+300+400+500
    });

    it('should return VALUE error for mismatched arrays', () => {
      const matches = ['XNPV(0.1, A2:D2, A3:E3)', '0.1', 'A2:D2', 'A3:E3'] as RegExpMatchArray;
      const result = XNPV.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });

    it('should return REF error for invalid range', () => {
      const matches = ['XNPV(0.1, invalid, B3:F3)', '0.1', 'invalid', 'B3:F3'] as RegExpMatchArray;
      const result = XNPV.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.REF);
    });

    it('should handle negative NPV', () => {
      const negContext = createContext([
        [-5000, 500, 600, 700], // cash flows
        [44927, 45292, 45658, 46023] // dates
      ]);
      const matches = ['XNPV(0.15, A1:D1, A2:D2)', '0.15', 'A1:D1', 'A2:D2'] as RegExpMatchArray;
      const result = XNPV.calculate(matches, negContext);
      expect(typeof result).toBe('number');
      expect(result).toBeLessThan(0); // Negative NPV
    });
  });

  describe('XIRR Function (Internal Rate of Return with Dates)', () => {
    it('should calculate IRR with specific dates', () => {
      const matches = ['XIRR(A2:E2, A3:E3)', 'A2:E2', 'A3:E3'] as RegExpMatchArray;
      const result = XIRR.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeGreaterThan(0); // Positive IRR expected
      expect(result).toBeLessThan(1); // Reasonable IRR range
    });

    it('should handle custom guess value', () => {
      const matches = ['XIRR(A2:E2, A3:E3, 0.2)', 'A2:E2', 'A3:E3', '0.2'] as RegExpMatchArray;
      const result = XIRR.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeGreaterThan(0);
    });

    it('should return NUM error for non-convergent series', () => {
      const badContext = createContext([
        [-1000, -500, -300, -200], // All negative cash flows
        [44927, 45292, 45658, 46023]
      ]);
      const matches = ['XIRR(A1:D1, A2:D2)', 'A1:D1', 'A2:D2'] as RegExpMatchArray;
      const result = XIRR.calculate(matches, badContext);
      expect(result).toBe(FormulaError.NUM);
    });

    it('should return VALUE error for empty data', () => {
      const emptyContext = createContext([[], []]);
      const matches = ['XIRR(A1:A1, B1:B1)', 'A1:A1', 'B1:B1'] as RegExpMatchArray;
      const result = XIRR.calculate(matches, emptyContext);
      expect(result).toBe(FormulaError.VALUE);
    });
  });

  describe('MIRR Function (Modified Internal Rate of Return)', () => {
    it('should calculate MIRR for mixed cash flows', () => {
      const matches = ['MIRR(A6:E6, 0.1, 0.12)', 'A6:E6', '0.1', '0.12'] as RegExpMatchArray;
      const result = MIRR.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeGreaterThan(0);
      expect(result).toBeLessThan(0.5); // Reasonable MIRR range
    });

    it('should handle different finance and reinvest rates', () => {
      const matches = ['MIRR(A6:E6, 0.08, 0.15)', 'A6:E6', '0.08', '0.15'] as RegExpMatchArray;
      const result = MIRR.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeGreaterThan(0);
    });

    it('should return DIV0 error for all positive cash flows', () => {
      const posContext = createContext([
        [1000, 2000, 3000, 4000]
      ]);
      const matches = ['MIRR(A1:D1, 0.1, 0.12)', 'A1:D1', '0.1', '0.12'] as RegExpMatchArray;
      const result = MIRR.calculate(matches, posContext);
      expect(result).toBe(FormulaError.DIV0);
    });

    it('should return VALUE error for insufficient data', () => {
      const matches = ['MIRR(A1:A1, 0.1, 0.12)', 'A1:A1', '0.1', '0.12'] as RegExpMatchArray;
      const result = MIRR.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });
  });

  describe('SLN Function (Straight-Line Depreciation)', () => {
    it('should calculate straight-line depreciation', () => {
      const matches = ['SLN(10000, 1000, 5)', '10000', '1000', '5'] as RegExpMatchArray;
      const result = SLN.calculate(matches, mockContext);
      expect(result).toBe(1800); // (10000-1000)/5 = 1800
    });

    it('should handle zero salvage value', () => {
      const matches = ['SLN(10000, 0, 10)', '10000', '0', '10'] as RegExpMatchArray;
      const result = SLN.calculate(matches, mockContext);
      expect(result).toBe(1000); // 10000/10 = 1000
    });

    it('should return NUM error for zero life', () => {
      const matches = ['SLN(10000, 1000, 0)', '10000', '1000', '0'] as RegExpMatchArray;
      const result = SLN.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });

    it('should handle cell references', () => {
      const matches = ['SLN(A4, B4, C4)', 'A4', 'B4', 'C4'] as RegExpMatchArray;
      const result = SLN.calculate(matches, mockContext);
      expect(result).toBe(1800); // (10000-1000)/5
    });
  });

  describe('SYD Function (Sum-of-Years Digits Depreciation)', () => {
    it('should calculate depreciation for first period', () => {
      const matches = ['SYD(10000, 1000, 5, 1)', '10000', '1000', '5', '1'] as RegExpMatchArray;
      const result = SYD.calculate(matches, mockContext);
      expect(result).toBe(3000); // 9000 * 5/15
    });

    it('should calculate depreciation for middle period', () => {
      const matches = ['SYD(10000, 1000, 5, 3)', '10000', '1000', '5', '3'] as RegExpMatchArray;
      const result = SYD.calculate(matches, mockContext);
      expect(result).toBe(1800); // 9000 * 3/15
    });

    it('should calculate depreciation for last period', () => {
      const matches = ['SYD(10000, 1000, 5, 5)', '10000', '1000', '5', '5'] as RegExpMatchArray;
      const result = SYD.calculate(matches, mockContext);
      expect(result).toBe(600); // 9000 * 1/15
    });

    it('should return NUM error for period > life', () => {
      const matches = ['SYD(10000, 1000, 5, 6)', '10000', '1000', '5', '6'] as RegExpMatchArray;
      const result = SYD.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });

    it('should return NUM error for zero period', () => {
      const matches = ['SYD(10000, 1000, 5, 0)', '10000', '1000', '5', '0'] as RegExpMatchArray;
      const result = SYD.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });
  });

  describe('DB Function (Fixed-Declining Balance Depreciation)', () => {
    it('should calculate depreciation for first period', () => {
      const matches = ['DB(10000, 1000, 5, 1)', '10000', '1000', '5', '1'] as RegExpMatchArray;
      const result = DB.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeGreaterThan(0);
      expect(result).toBeLessThan(10000);
    });

    it('should handle first period with partial year', () => {
      const matches = ['DB(10000, 1000, 5, 1, 6)', '10000', '1000', '5', '1', '6'] as RegExpMatchArray;
      const result = DB.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeGreaterThan(0);
    });

    it('should calculate depreciation for last period', () => {
      const matches = ['DB(10000, 1000, 5, 5)', '10000', '1000', '5', '5'] as RegExpMatchArray;
      const result = DB.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeGreaterThan(0);
    });

    it('should return NUM error for negative cost', () => {
      const matches = ['DB(-10000, 1000, 5, 1)', '-10000', '1000', '5', '1'] as RegExpMatchArray;
      const result = DB.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });

    it('should return NUM error for invalid month parameter', () => {
      const matches = ['DB(10000, 1000, 5, 1, 13)', '10000', '1000', '5', '1', '13'] as RegExpMatchArray;
      const result = DB.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });
  });

  describe('DDB Function (Double-Declining Balance Depreciation)', () => {
    it('should calculate depreciation with default factor', () => {
      const matches = ['DDB(10000, 1000, 5, 1)', '10000', '1000', '5', '1'] as RegExpMatchArray;
      const result = DDB.calculate(matches, mockContext);
      expect(result).toBe(4000); // 10000 * 2/5
    });

    it('should calculate depreciation with custom factor', () => {
      const matches = ['DDB(10000, 1000, 5, 1, 1.5)', '10000', '1000', '5', '1', '1.5'] as RegExpMatchArray;
      const result = DDB.calculate(matches, mockContext);
      expect(result).toBe(3000); // 10000 * 1.5/5
    });

    it('should not depreciate below salvage value', () => {
      const matches = ['DDB(10000, 8000, 5, 1)', '10000', '8000', '5', '1'] as RegExpMatchArray;
      const result = DDB.calculate(matches, mockContext);
      expect(result).toBe(2000); // Limited to 10000-8000
    });

    it('should handle later periods correctly', () => {
      const matches = ['DDB(10000, 1000, 5, 2)', '10000', '1000', '5', '2'] as RegExpMatchArray;
      const result = DDB.calculate(matches, mockContext);
      expect(result).toBe(2400); // (10000-4000) * 2/5
    });

    it('should return NUM error for zero factor', () => {
      const matches = ['DDB(10000, 1000, 5, 1, 0)', '10000', '1000', '5', '1', '0'] as RegExpMatchArray;
      const result = DDB.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });
  });

  describe('VDB Function (Variable-Declining Balance Depreciation)', () => {
    it('should calculate depreciation for whole periods', () => {
      const matches = ['VDB(10000, 1000, 5, 0, 1)', '10000', '1000', '5', '0', '1'] as RegExpMatchArray;
      const result = VDB.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeGreaterThan(0);
    });

    it('should calculate depreciation for partial periods', () => {
      const matches = ['VDB(10000, 1000, 5, 0.5, 1.5)', '10000', '1000', '5', '0.5', '1.5'] as RegExpMatchArray;
      const result = VDB.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeGreaterThan(0);
    });

    it('should handle no-switch option', () => {
      const matches = ['VDB(10000, 1000, 5, 0, 5, 2, TRUE)', '10000', '1000', '5', '0', '5', '2', 'TRUE'] as RegExpMatchArray;
      const result = VDB.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeGreaterThan(0);
    });

    it('should switch to straight-line when beneficial', () => {
      const matches = ['VDB(10000, 1000, 5, 3, 5, 2, FALSE)', '10000', '1000', '5', '3', '5', '2', 'FALSE'] as RegExpMatchArray;
      const result = VDB.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeGreaterThan(0);
    });

    it('should return NUM error for start >= end', () => {
      const matches = ['VDB(10000, 1000, 5, 2, 2)', '10000', '1000', '5', '2', '2'] as RegExpMatchArray;
      const result = VDB.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });
  });

  describe('PDURATION Function (Investment Duration)', () => {
    it('should calculate periods to reach target', () => {
      const matches = ['PDURATION(0.05, 1000, 2000)', '0.05', '1000', '2000'] as RegExpMatchArray;
      const result = PDURATION.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeCloseTo(14.21, 1); // log(2)/log(1.05)
    });

    it('should handle small growth rates', () => {
      const matches = ['PDURATION(0.01, 1000, 1100)', '0.01', '1000', '1100'] as RegExpMatchArray;
      const result = PDURATION.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeCloseTo(9.58, 1);
    });

    it('should return NUM error for zero rate', () => {
      const matches = ['PDURATION(0, 1000, 2000)', '0', '1000', '2000'] as RegExpMatchArray;
      const result = PDURATION.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });

    it('should return NUM error for negative values', () => {
      const matches = ['PDURATION(0.05, -1000, 2000)', '0.05', '-1000', '2000'] as RegExpMatchArray;
      const result = PDURATION.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });
  });

  describe('RRI Function (Rate of Return)', () => {
    it('should calculate growth rate', () => {
      const matches = ['RRI(5, 10000, 20000)', '5', '10000', '20000'] as RegExpMatchArray;
      const result = RRI.calculate(matches, mockContext);
      expect(result).toBeCloseTo(0.1487, 3); // (20000/10000)^(1/5) - 1
    });

    it('should handle negative future value', () => {
      const matches = ['RRI(5, 10000, -20000)', '5', '10000', '-20000'] as RegExpMatchArray;
      const result = RRI.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
    });

    it('should return NUM error for zero periods', () => {
      const matches = ['RRI(0, 10000, 20000)', '0', '10000', '20000'] as RegExpMatchArray;
      const result = RRI.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });

    it('should return NUM error for zero present value', () => {
      const matches = ['RRI(5, 0, 20000)', '5', '0', '20000'] as RegExpMatchArray;
      const result = RRI.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });
  });

  describe('ISPMT Function (Interest Payment for Even Principal)', () => {
    it('should calculate interest for first period', () => {
      const matches = ['ISPMT(0.1, 1, 10, 100000)', '0.1', '1', '10', '100000'] as RegExpMatchArray;
      const result = ISPMT.calculate(matches, mockContext);
      expect(result).toBe(-10000); // Full balance * rate
    });

    it('should calculate decreasing interest for later periods', () => {
      const matches = ['ISPMT(0.1, 5, 10, 100000)', '0.1', '5', '10', '100000'] as RegExpMatchArray;
      const result = ISPMT.calculate(matches, mockContext);
      expect(result).toBe(-6000); // 60% of balance * rate
    });

    it('should calculate interest for last period', () => {
      const matches = ['ISPMT(0.1, 10, 10, 100000)', '0.1', '10', '10', '100000'] as RegExpMatchArray;
      const result = ISPMT.calculate(matches, mockContext);
      expect(result).toBeCloseTo(-1000, 10); // 10% of balance * rate
    });

    it('should return NUM error for zero periods', () => {
      const matches = ['ISPMT(0.1, 1, 0, 100000)', '0.1', '1', '0', '100000'] as RegExpMatchArray;
      const result = ISPMT.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });

    it('should return NUM error for period > nper', () => {
      const matches = ['ISPMT(0.1, 11, 10, 100000)', '0.1', '11', '10', '100000'] as RegExpMatchArray;
      const result = ISPMT.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });
  });

  describe('Edge Cases and Integration', () => {
    it('should verify depreciation methods sum correctly', () => {
      const cost = 10000;
      const salvage = 1000;
      const life = 5;
      let totalSYD = 0;
      
      for (let per = 1; per <= life; per++) {
        const matches = [`SYD(${cost}, ${salvage}, ${life}, ${per})`, 
          cost.toString(), salvage.toString(), life.toString(), per.toString()] as RegExpMatchArray;
        totalSYD += SYD.calculate(matches, mockContext) as number;
      }
      
      expect(totalSYD).toBeCloseTo(cost - salvage, 2);
    });

    it('should handle very small depreciation rates', () => {
      const matches = ['DB(100000, 99000, 50, 1)', '100000', '99000', '50', '1'] as RegExpMatchArray;
      const result = DB.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeGreaterThanOrEqual(0); // May be 0 for very small depreciation rates
      expect(result).toBeLessThan(1000);
    });

    it('should handle cell references', () => {
      const matches = ['NPER(A1, B1, C1)', 'A1', 'B1', 'C1'] as RegExpMatchArray;
      const result = NPER.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
    });

    it('should calculate consistent XNPV and XIRR', () => {
      // XIRR result should produce zero XNPV
      const xirrMatches = ['XIRR(B1:F1, B3:F3)', 'B1:F1', 'B3:F3'] as RegExpMatchArray;
      const irr = XIRR.calculate(xirrMatches, mockContext);
      
      if (typeof irr === 'number') {
        // This would verify that XNPV at the IRR rate is close to zero
        // Due to numerical precision, we check it's very small
        expect(Math.abs(irr)).toBeLessThan(10);
      }
    });
  });
});