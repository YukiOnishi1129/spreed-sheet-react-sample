import { describe, it, expect } from 'vitest';
import {
  DATEDIF, NOW, DATE, YEAR, MONTH, DAY,
  WEEKDAY, DAYS, EDATE, TODAY, EOMONTH,
  TIME, HOUR, MINUTE, SECOND,
  WEEKNUM, DAYS360, YEARFRAC,
  DATEVALUE, TIMEVALUE, ISOWEEKNUM
} from '../dateLogic';
import { FormulaError } from '../../shared/types';
import type { FormulaContext } from '../../shared/types';

// モックコンテキスト作成
const createContext = (data: (string | number | boolean | null)[][]): FormulaContext => ({
  data: data.map(row => row.map(cell => ({ value: cell }))),
  row: 0,
  col: 0
});

describe('Date Functions', () => {
  const mockContext = createContext([
    ['2024-01-15', '2024-12-31', '2023-01-01', '2024-02-29', 'invalid'],
    ['2025-01-15', '2024-01-01', '2024-03-01', '2024-04-30', 45292],
    [new Date('2024-01-15'), new Date('2024-12-31'), '15:30:45', '23:59:59', '00:00:01'],
    ['1/15/2024', '12/31/2024', '2024', '1', '15'],
    ['text', '', null, true, false],
  ]);


  describe('DATEDIF', () => {
    it('should calculate difference in days', () => {
      const matches = ['DATEDIF(A1, B1, "D")', 'A1', 'B1', '"D"'] as RegExpMatchArray;
      const result = DATEDIF.calculate(matches, mockContext);
      expect(result).toBe(351); // Days from Jan 15 to Dec 31, 2024
    });

    it('should calculate difference in months', () => {
      const matches = ['DATEDIF(A1, B1, "M")', 'A1', 'B1', '"M"'] as RegExpMatchArray;
      const result = DATEDIF.calculate(matches, mockContext);
      expect(result).toBe(11); // Complete months
    });

    it('should calculate difference in years', () => {
      const matches = ['DATEDIF(A1, A2, "Y")', 'A1', 'A2', '"Y"'] as RegExpMatchArray;
      const result = DATEDIF.calculate(matches, mockContext);
      expect(result).toBe(1); // 1 complete year
    });

    it('should calculate YM (months ignoring years)', () => {
      const matches = ['DATEDIF(A1, B1, "YM")', 'A1', 'B1', '"YM"'] as RegExpMatchArray;
      const result = DATEDIF.calculate(matches, mockContext);
      expect(result).toBe(11); // Jan to Dec = 11 months
    });

    it('should calculate YD (days ignoring years)', () => {
      const matches = ['DATEDIF(A1, B1, "YD")', 'A1', 'B1', '"YD"'] as RegExpMatchArray;
      const result = DATEDIF.calculate(matches, mockContext);
      expect(result).toBe(351); // Days from Jan 15 to Dec 31
    });

    it('should calculate MD (days ignoring months and years)', () => {
      const matches = ['DATEDIF(A1, B1, "MD")', 'A1', 'B1', '"MD"'] as RegExpMatchArray;
      const result = DATEDIF.calculate(matches, mockContext);
      expect(result).toBe(16); // 31 - 15 = 16
    });

    it('should handle string dates', () => {
      const matches = ['DATEDIF("2024-01-01", "2024-12-31", "D")', '"2024-01-01"', '"2024-12-31"', '"D"'] as RegExpMatchArray;
      const result = DATEDIF.calculate(matches, mockContext);
      expect(result).toBe(365); // 2024 is a leap year
    });

    it('should return NUM error for invalid dates', () => {
      const matches = ['DATEDIF(A1, E1, "D")', 'A1', 'E1', '"D"'] as RegExpMatchArray;
      const result = DATEDIF.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });

    it('should return NUM error when end date is before start date', () => {
      const matches = ['DATEDIF(B1, A1, "D")', 'B1', 'A1', '"D"'] as RegExpMatchArray;
      const result = DATEDIF.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });

    it('should return VALUE error for invalid unit', () => {
      const matches = ['DATEDIF(A1, B1, "X")', 'A1', 'B1', '"X"'] as RegExpMatchArray;
      const result = DATEDIF.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });
  });

  describe('NOW', () => {
    it('should return current date and time as Excel serial', () => {
      const matches = ['NOW()'] as RegExpMatchArray;
      const result = NOW.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeGreaterThan(44000); // After year 2020
      expect(result).toBeLessThan(50000); // Before year 2037
    });
  });

  describe('TODAY', () => {
    it('should return current date as Excel serial', () => {
      const matches = ['TODAY()'] as RegExpMatchArray;
      const result = TODAY.calculate(matches, mockContext);
      expect(typeof result).toBe('number');
      expect(result).toBeGreaterThan(44000); // After year 2020
      expect(result).toBeLessThan(50000); // Before year 2037
      expect(result % 1).toBe(0); // Should be a whole number (no time component)
    });
  });

  describe('DATE', () => {
    it('should create date from year, month, day', () => {
      const matches = ['DATE(2024, 1, 15)', '2024', '1', '15'] as RegExpMatchArray;
      const result = DATE.calculate(matches, mockContext);
      expect(result).toBe(45306); // Excel serial for 2024-01-15
    });

    it('should handle cell references', () => {
      const matches = ['DATE(C4, D4, E4)', 'C4', 'D4', 'E4'] as RegExpMatchArray;
      const result = DATE.calculate(matches, mockContext);
      expect(result).toBe(45306); // Excel serial for 2024-01-15
    });

    it('should handle month overflow', () => {
      const matches = ['DATE(2024, 13, 1)', '2024', '13', '1'] as RegExpMatchArray;
      const result = DATE.calculate(matches, mockContext);
      expect(result).toBe(45658); // 2025-01-01
    });

    it('should handle day overflow', () => {
      const matches = ['DATE(2024, 1, 32)', '2024', '1', '32'] as RegExpMatchArray;
      const result = DATE.calculate(matches, mockContext);
      expect(result).toBe(45323); // 2024-02-01
    });

    it('should return VALUE error for non-numeric inputs', () => {
      const matches = ['DATE(text, 1, 1)', 'text', '1', '1'] as RegExpMatchArray;
      const result = DATE.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });
  });

  describe('Date Component Functions', () => {
    describe('YEAR', () => {
      it('should extract year from date', () => {
        const matches = ['YEAR(A1)', 'A1'] as RegExpMatchArray;
        const result = YEAR.calculate(matches, mockContext);
        expect(result).toBe(2024);
      });

      it('should handle Excel serial numbers', () => {
        const matches = ['YEAR(45292)', '45292'] as RegExpMatchArray;
        const result = YEAR.calculate(matches, mockContext);
        expect(result).toBe(2024);
      });

      it('should return VALUE error for invalid date', () => {
        const matches = ['YEAR(E1)', 'E1'] as RegExpMatchArray;
        const result = YEAR.calculate(matches, mockContext);
        expect(result).toBe(FormulaError.VALUE);
      });
    });

    describe('MONTH', () => {
      it('should extract month from date', () => {
        const matches = ['MONTH(A1)', 'A1'] as RegExpMatchArray;
        const result = MONTH.calculate(matches, mockContext);
        expect(result).toBe(1);
      });

      it('should extract month from December date', () => {
        const matches = ['MONTH(B1)', 'B1'] as RegExpMatchArray;
        const result = MONTH.calculate(matches, mockContext);
        expect(result).toBe(12);
      });
    });

    describe('DAY', () => {
      it('should extract day from date', () => {
        const matches = ['DAY(A1)', 'A1'] as RegExpMatchArray;
        const result = DAY.calculate(matches, mockContext);
        expect(result).toBe(15);
      });

      it('should extract day from end of month', () => {
        const matches = ['DAY(B1)', 'B1'] as RegExpMatchArray;
        const result = DAY.calculate(matches, mockContext);
        expect(result).toBe(31);
      });
    });
  });

  describe('WEEKDAY', () => {
    it('should return weekday with type 1 (Sunday=1)', () => {
      const matches = ['WEEKDAY(A1)', 'A1'] as RegExpMatchArray;
      const result = WEEKDAY.calculate(matches, mockContext);
      expect(result).toBe(2); // Monday = 2
    });

    it('should return weekday with type 2 (Monday=1)', () => {
      const matches = ['WEEKDAY(A1, 2)', 'A1', '2'] as RegExpMatchArray;
      const result = WEEKDAY.calculate(matches, mockContext);
      expect(result).toBe(1); // Monday = 1
    });

    it('should return weekday with type 3 (Monday=0)', () => {
      const matches = ['WEEKDAY(A1, 3)', 'A1', '3'] as RegExpMatchArray;
      const result = WEEKDAY.calculate(matches, mockContext);
      expect(result).toBe(0); // Monday = 0
    });

    it('should return VALUE error for invalid date', () => {
      const matches = ['WEEKDAY(E1)', 'E1'] as RegExpMatchArray;
      const result = WEEKDAY.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });
  });

  describe('DAYS', () => {
    it('should calculate days between dates', () => {
      const matches = ['DAYS(B1, A1)', 'B1', 'A1'] as RegExpMatchArray;
      const result = DAYS.calculate(matches, mockContext);
      expect(result).toBe(351);
    });

    it('should handle negative days', () => {
      const matches = ['DAYS(A1, B1)', 'A1', 'B1'] as RegExpMatchArray;
      const result = DAYS.calculate(matches, mockContext);
      expect(result).toBe(-351);
    });

    it('should return VALUE error for invalid dates', () => {
      const matches = ['DAYS(E1, A1)', 'E1', 'A1'] as RegExpMatchArray;
      const result = DAYS.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });
  });

  describe('EDATE', () => {
    it('should add months to date', () => {
      const matches = ['EDATE(A1, 3)', 'A1', '3'] as RegExpMatchArray;
      const result = EDATE.calculate(matches, mockContext);
      expect(result).toBe(45397); // 2024-04-15
    });

    it('should subtract months from date', () => {
      const matches = ['EDATE(A1, -3)', 'A1', '-3'] as RegExpMatchArray;
      const result = EDATE.calculate(matches, mockContext);
      expect(result).toBe(45214); // 2023-10-15
    });

    it('should handle month-end dates', () => {
      const matches = ['EDATE(D1, 1)', 'D1', '1'] as RegExpMatchArray;
      const result = EDATE.calculate(matches, mockContext);
      expect(result).toBe(45380); // 2024-03-29 (Feb 29 + 1 month)
    });

    it('should return VALUE error for invalid date', () => {
      const matches = ['EDATE(E1, 1)', 'E1', '1'] as RegExpMatchArray;
      const result = EDATE.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });
  });

  describe('EOMONTH', () => {
    it('should return end of current month', () => {
      const matches = ['EOMONTH(A1, 0)', 'A1', '0'] as RegExpMatchArray;
      const result = EOMONTH.calculate(matches, mockContext);
      expect(result).toBe(45322); // 2024-01-31
    });

    it('should return end of next month', () => {
      const matches = ['EOMONTH(A1, 1)', 'A1', '1'] as RegExpMatchArray;
      const result = EOMONTH.calculate(matches, mockContext);
      expect(result).toBe(45351); // 2024-02-29 (leap year)
    });

    it('should return end of previous month', () => {
      const matches = ['EOMONTH(A1, -1)', 'A1', '-1'] as RegExpMatchArray;
      const result = EOMONTH.calculate(matches, mockContext);
      expect(result).toBe(45291); // 2023-12-31
    });

    it('should handle leap year February', () => {
      const matches = ['EOMONTH(D1, 0)', 'D1', '0'] as RegExpMatchArray;
      const result = EOMONTH.calculate(matches, mockContext);
      expect(result).toBe(45351); // 2024-02-29
    });
  });

  describe('Time Functions', () => {
    describe('TIME', () => {
      it('should create time from hours, minutes, seconds', () => {
        const matches = ['TIME(15, 30, 45)', '15', '30', '45'] as RegExpMatchArray;
        const result = TIME.calculate(matches, mockContext);
        expect(result).toBeCloseTo(0.6463542, 7); // 15:30:45 as fraction of day
      });

      it('should handle hour overflow', () => {
        const matches = ['TIME(25, 0, 0)', '25', '0', '0'] as RegExpMatchArray;
        const result = TIME.calculate(matches, mockContext);
        expect(result).toBeCloseTo(0.0416667, 7); // 01:00:00
      });

      it('should handle minute overflow', () => {
        const matches = ['TIME(1, 65, 0)', '1', '65', '0'] as RegExpMatchArray;
        const result = TIME.calculate(matches, mockContext);
        expect(result).toBeCloseTo(0.0868056, 7); // 02:05:00
      });

      it('should return VALUE error for non-numeric inputs', () => {
        const matches = ['TIME(text, 0, 0)', 'text', '0', '0'] as RegExpMatchArray;
        const result = TIME.calculate(matches, mockContext);
        expect(result).toBe(FormulaError.VALUE);
      });
    });

    describe('HOUR', () => {
      it('should extract hour from time string', () => {
        const matches = ['HOUR(C3)', 'C3'] as RegExpMatchArray;
        const result = HOUR.calculate(matches, mockContext);
        expect(result).toBe(15);
      });

      it('should extract hour from decimal time', () => {
        const matches = ['HOUR(0.75)', '0.75'] as RegExpMatchArray;
        const result = HOUR.calculate(matches, mockContext);
        expect(result).toBe(18); // 0.75 * 24 = 18
      });
    });

    describe('MINUTE', () => {
      it('should extract minute from time string', () => {
        const matches = ['MINUTE(C3)', 'C3'] as RegExpMatchArray;
        const result = MINUTE.calculate(matches, mockContext);
        expect(result).toBe(30);
      });
    });

    describe('SECOND', () => {
      it('should extract second from time string', () => {
        const matches = ['SECOND(C3)', 'C3'] as RegExpMatchArray;
        const result = SECOND.calculate(matches, mockContext);
        expect(result).toBe(45);
      });
    });
  });

  describe('Week Functions', () => {
    describe('WEEKNUM', () => {
      it('should calculate week number with type 1', () => {
        const matches = ['WEEKNUM(A1)', 'A1'] as RegExpMatchArray;
        const result = WEEKNUM.calculate(matches, mockContext);
        expect(result).toBe(3); // Week 3 of 2024
      });

      it('should calculate week number with type 2', () => {
        const matches = ['WEEKNUM(A1, 2)', 'A1', '2'] as RegExpMatchArray;
        const result = WEEKNUM.calculate(matches, mockContext);
        expect(result).toBe(3);
      });

      it('should return VALUE error for invalid date', () => {
        const matches = ['WEEKNUM(E1)', 'E1'] as RegExpMatchArray;
        const result = WEEKNUM.calculate(matches, mockContext);
        expect(result).toBe(FormulaError.VALUE);
      });
    });

    describe('ISOWEEKNUM', () => {
      it('should calculate ISO week number', () => {
        const matches = ['ISOWEEKNUM(A1)', 'A1'] as RegExpMatchArray;
        const result = ISOWEEKNUM.calculate(matches, mockContext);
        expect(result).toBe(3); // ISO week 3 of 2024
      });

      it('should handle year boundary correctly', () => {
        const matches = ['ISOWEEKNUM("2024-01-01")', '"2024-01-01"'] as RegExpMatchArray;
        const result = ISOWEEKNUM.calculate(matches, mockContext);
        expect(result).toBe(1);
      });
    });
  });

  describe('DAYS360', () => {
    it('should calculate days using 360-day year', () => {
      const matches = ['DAYS360(A1, B1)', 'A1', 'B1'] as RegExpMatchArray;
      const result = DAYS360.calculate(matches, mockContext);
      expect(result).toBe(346); // Using 30-day months
    });

    it('should handle European method', () => {
      const matches = ['DAYS360(A1, B1, TRUE)', 'A1', 'B1', 'TRUE'] as RegExpMatchArray;
      const result = DAYS360.calculate(matches, mockContext);
      expect(result).toBe(346);
    });

    it('should return VALUE error for invalid dates', () => {
      const matches = ['DAYS360(E1, B1)', 'E1', 'B1'] as RegExpMatchArray;
      const result = DAYS360.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });
  });

  describe('YEARFRAC', () => {
    it('should calculate year fraction with basis 0', () => {
      const matches = ['YEARFRAC(A1, B1)', 'A1', 'B1'] as RegExpMatchArray;
      const result = YEARFRAC.calculate(matches, mockContext);
      expect(result).toBeCloseTo(0.9611, 4);
    });

    it('should calculate year fraction with basis 1 (actual/actual)', () => {
      const matches = ['YEARFRAC(A1, B1, 1)', 'A1', 'B1', '1'] as RegExpMatchArray;
      const result = YEARFRAC.calculate(matches, mockContext);
      expect(result).toBeCloseTo(0.9617, 4);
    });

    it('should calculate year fraction with basis 2 (actual/360)', () => {
      const matches = ['YEARFRAC(A1, B1, 2)', 'A1', 'B1', '2'] as RegExpMatchArray;
      const result = YEARFRAC.calculate(matches, mockContext);
      expect(result).toBeCloseTo(0.975, 3);
    });

    it('should calculate year fraction with basis 3 (actual/365)', () => {
      const matches = ['YEARFRAC(A1, B1, 3)', 'A1', 'B1', '3'] as RegExpMatchArray;
      const result = YEARFRAC.calculate(matches, mockContext);
      expect(result).toBeCloseTo(0.9616, 4);
    });

    it('should calculate year fraction with basis 4 (30/360 European)', () => {
      const matches = ['YEARFRAC(A1, B1, 4)', 'A1', 'B1', '4'] as RegExpMatchArray;
      const result = YEARFRAC.calculate(matches, mockContext);
      expect(result).toBeCloseTo(0.9611, 4);
    });

    it('should return NUM error for invalid basis', () => {
      const matches = ['YEARFRAC(A1, B1, 5)', 'A1', 'B1', '5'] as RegExpMatchArray;
      const result = YEARFRAC.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });
  });

  describe('Date/Time Parsing Functions', () => {
    describe('DATEVALUE', () => {
      it('should convert date string to Excel serial', () => {
        const matches = ['DATEVALUE("2024-01-15")', '"2024-01-15"'] as RegExpMatchArray;
        const result = DATEVALUE.calculate(matches, mockContext);
        expect(result).toBe(45306);
      });

      it('should handle various date formats', () => {
        const matches = ['DATEVALUE("1/15/2024")', '"1/15/2024"'] as RegExpMatchArray;
        const result = DATEVALUE.calculate(matches, mockContext);
        expect(result).toBe(45306);
      });

      it('should return VALUE error for invalid date string', () => {
        const matches = ['DATEVALUE("invalid")', '"invalid"'] as RegExpMatchArray;
        const result = DATEVALUE.calculate(matches, mockContext);
        expect(result).toBe(FormulaError.VALUE);
      });
    });

    describe('TIMEVALUE', () => {
      it('should convert time string to decimal', () => {
        const matches = ['TIMEVALUE("15:30:45")', '"15:30:45"'] as RegExpMatchArray;
        const result = TIMEVALUE.calculate(matches, mockContext);
        expect(result).toBeCloseTo(0.6463542, 7);
      });

      it('should handle 24-hour time', () => {
        const matches = ['TIMEVALUE("23:59:59")', '"23:59:59"'] as RegExpMatchArray;
        const result = TIMEVALUE.calculate(matches, mockContext);
        expect(result).toBeCloseTo(0.9999884, 7);
      });

      it('should return VALUE error for invalid time string', () => {
        const matches = ['TIMEVALUE("25:00:00")', '"25:00:00"'] as RegExpMatchArray;
        const result = TIMEVALUE.calculate(matches, mockContext);
        expect(result).toBe(FormulaError.VALUE);
      });
    });
  });
});