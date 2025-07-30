import { describe, it, expect } from 'vitest';
import {
  NETWORKDAYS, NETWORKDAYS_INTL, WORKDAY, WORKDAY_INTL
} from '../workdayLogic';
import { FormulaError } from '../../shared/types';
import type { FormulaContext } from '../../shared/types';

// Helper function to create FormulaContext
const createContext = (data: (string | number | boolean | null)[][]): FormulaContext => ({
  data: data.map(row => row.map(cell => ({ value: cell }))),
  row: 0,
  col: 0
});

describe('Workday Functions', () => {
  const mockContext = createContext([
    ['2024-01-15', '2024-01-31', '2024-01-01', '2024-12-25', 'invalid'],
    ['2024-02-01', '2024-02-29', '2024-07-04', '2024-11-28', 45292],
    ['2024-03-01', '2024-03-15', '2024-01-15', '2024-01-16', 10],
    ['2024-06-01', '2024-06-30', '2024-01-17', '1', -10],
    ['text', '', null, true, false],
  ]);

  describe('NETWORKDAYS', () => {
    it('should calculate working days between dates', () => {
      const matches = ['NETWORKDAYS(A1, B1)', 'A1', 'B1'] as RegExpMatchArray;
      const result = NETWORKDAYS.calculate(matches, mockContext);
      expect(result).toBe(13); // Jan 15-31, excluding weekends
    });

    it('should handle dates in reverse order', () => {
      const matches = ['NETWORKDAYS(B1, A1)', 'B1', 'A1'] as RegExpMatchArray;
      const result = NETWORKDAYS.calculate(matches, mockContext);
      expect(result).toBe(-13); // Negative when end < start
    });

    it('should exclude holidays', () => {
      const matches = ['NETWORKDAYS(A1, B1, C3:D3)', 'A1', 'B1', 'C3:D3'] as RegExpMatchArray;
      const result = NETWORKDAYS.calculate(matches, mockContext);
      expect(result).toBe(11); // Excludes Jan 15 and 16 as holidays (13 - 2 = 11)
    });

    it('should handle single holiday', () => {
      const matches = ['NETWORKDAYS(A1, B1, C3)', 'A1', 'B1', 'C3'] as RegExpMatchArray;
      const result = NETWORKDAYS.calculate(matches, mockContext);
      expect(result).toBe(12); // Excludes Jan 15 as holiday
    });

    it('should handle string dates', () => {
      const matches = ['NETWORKDAYS("2024-01-01", "2024-01-31")', '"2024-01-01"', '"2024-01-31"'] as RegExpMatchArray;
      const result = NETWORKDAYS.calculate(matches, mockContext);
      expect(result).toBe(23); // January 2024 working days
    });

    it('should handle Excel serial numbers', () => {
      const matches = ['NETWORKDAYS(E2, 45322)', 'E2', '45322'] as RegExpMatchArray;
      const result = NETWORKDAYS.calculate(matches, mockContext);
      expect(result).toBe(23); // From Jan 1 to Jan 31, 2024
    });

    it('should return VALUE error for invalid start date', () => {
      const matches = ['NETWORKDAYS(E1, B1)', 'E1', 'B1'] as RegExpMatchArray;
      const result = NETWORKDAYS.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });

    it('should return VALUE error for invalid end date', () => {
      const matches = ['NETWORKDAYS(A1, E1)', 'A1', 'E1'] as RegExpMatchArray;
      const result = NETWORKDAYS.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });

    it('should handle same start and end date', () => {
      const matches = ['NETWORKDAYS(A1, A1)', 'A1', 'A1'] as RegExpMatchArray;
      const result = NETWORKDAYS.calculate(matches, mockContext);
      expect(result).toBe(1); // Monday counts as 1 working day
    });

    it('should handle weekend dates', () => {
      const matches = ['NETWORKDAYS("2024-01-13", "2024-01-14")', '"2024-01-13"', '"2024-01-14"'] as RegExpMatchArray;
      const result = NETWORKDAYS.calculate(matches, mockContext);
      expect(result).toBe(0); // Saturday to Sunday = 0 working days
    });
  });

  describe('NETWORKDAYS.INTL', () => {
    it('should calculate working days with default weekend (Sat-Sun)', () => {
      const matches = ['NETWORKDAYS.INTL(A1, B1)', 'A1', 'B1'] as RegExpMatchArray;
      const result = NETWORKDAYS_INTL.calculate(matches, mockContext);
      expect(result).toBe(13); // Same as NETWORKDAYS
    });

    it('should handle different weekend definitions', () => {
      // Weekend = Sunday only (code 11)
      const matches = ['NETWORKDAYS.INTL(A1, B1, 11)', 'A1', 'B1', '11'] as RegExpMatchArray;
      const result = NETWORKDAYS_INTL.calculate(matches, mockContext);
      expect(result).toBe(15); // More working days with only Sunday as weekend
    });

    it('should handle custom weekend string', () => {
      // Custom weekend: "0000011" = Friday and Saturday
      const matches = ['NETWORKDAYS.INTL(A1, B1, "0000011")', 'A1', 'B1', '"0000011"'] as RegExpMatchArray;
      const result = NETWORKDAYS_INTL.calculate(matches, mockContext);
      expect(result).toBe(13); // Different weekend days
    });

    it('should handle weekend codes 1-7', () => {
      // Code 2 = Sunday-Monday weekend
      const matches = ['NETWORKDAYS.INTL(A1, B1, 2)', 'A1', 'B1', '2'] as RegExpMatchArray;
      const result = NETWORKDAYS_INTL.calculate(matches, mockContext);
      expect(result).toBe(12); // With Sun-Mon weekends, there are 12 working days
    });

    it('should exclude holidays with custom weekend', () => {
      const matches = ['NETWORKDAYS.INTL(A1, B1, 1, C3)', 'A1', 'B1', '1', 'C3'] as RegExpMatchArray;
      const result = NETWORKDAYS_INTL.calculate(matches, mockContext);
      expect(result).toBe(12); // Excludes Jan 15 as holiday
    });

    it('should return VALUE error for invalid dates', () => {
      const matches = ['NETWORKDAYS.INTL(E1, B1)', 'E1', 'B1'] as RegExpMatchArray;
      const result = NETWORKDAYS_INTL.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });
  });

  describe('WORKDAY', () => {
    it('should calculate future workday', () => {
      const matches = ['WORKDAY(A1, 10)', 'A1', '10'] as RegExpMatchArray;
      const result = WORKDAY.calculate(matches, mockContext);
      expect(result).toBe(45320); // Jan 15 + 10 working days
    });

    it('should calculate past workday with negative days', () => {
      const matches = ['WORKDAY(B1, -10)', 'B1', '-10'] as RegExpMatchArray;
      const result = WORKDAY.calculate(matches, mockContext);
      expect(result).toBe(45308); // Jan 31 - 10 working days = Jan 17
    });

    it('should skip weekends', () => {
      const matches = ['WORKDAY("2024-01-12", 1)', '"2024-01-12"', '1'] as RegExpMatchArray;
      const result = WORKDAY.calculate(matches, mockContext);
      expect(result).toBe(45306); // Friday + 1 = Monday (skips weekend)
    });

    it('should handle holidays', () => {
      const matches = ['WORKDAY(A1, 1, C3)', 'A1', '1', 'C3'] as RegExpMatchArray;
      const result = WORKDAY.calculate(matches, mockContext);
      expect(result).toBe(45307); // Skip Jan 15 (holiday), result is Jan 16
    });

    it('should handle holiday range', () => {
      const matches = ['WORKDAY(A1, 5, C3:D3)', 'A1', '5', 'C3:D3'] as RegExpMatchArray;
      const result = WORKDAY.calculate(matches, mockContext);
      expect(result).toBe(45314); // Skip holidays in range
    });

    it('should handle zero days', () => {
      const matches = ['WORKDAY(A1, 0)', 'A1', '0'] as RegExpMatchArray;
      const result = WORKDAY.calculate(matches, mockContext);
      expect(result).toBe(45306); // Same date
    });

    it('should handle Excel serial input', () => {
      const matches = ['WORKDAY(E2, 10)', 'E2', '10'] as RegExpMatchArray;
      const result = WORKDAY.calculate(matches, mockContext);
      expect(result).toBe(45306); // Excel serial 45292 (Jan 1) + 10 working days = Jan 15
    });

    it('should return VALUE error for invalid date', () => {
      const matches = ['WORKDAY(E1, 10)', 'E1', '10'] as RegExpMatchArray;
      const result = WORKDAY.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });

    it('should return VALUE error for invalid days', () => {
      const matches = ['WORKDAY(A1, "text")', 'A1', '"text"'] as RegExpMatchArray;
      const result = WORKDAY.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });
  });

  describe('WORKDAY.INTL', () => {
    it('should calculate workday with default weekend', () => {
      const matches = ['WORKDAY.INTL(A1, 10)', 'A1', '10'] as RegExpMatchArray;
      const result = WORKDAY_INTL.calculate(matches, mockContext);
      expect(result).toBe(45320); // Same as WORKDAY
    });

    it('should handle different weekend definitions', () => {
      // Weekend = Sunday only (code 11)
      const matches = ['WORKDAY.INTL(A1, 10, 11)', 'A1', '10', '11'] as RegExpMatchArray;
      const result = WORKDAY_INTL.calculate(matches, mockContext);
      expect(result).toBe(45317); // Jan 15 + 10 working days with Sunday-only weekend = Jan 26
    });

    it('should handle custom weekend string', () => {
      // Custom weekend: "0000011" = Friday and Saturday
      const matches = ['WORKDAY.INTL("2024-01-11", 1, "0000011")', '"2024-01-11"', '1', '"0000011"'] as RegExpMatchArray;
      const result = WORKDAY_INTL.calculate(matches, mockContext);
      expect(result).toBe(45305); // Thursday + 1 = Sunday (skips Fri-Sat)
    });

    it('should handle negative days with custom weekend', () => {
      const matches = ['WORKDAY.INTL(B1, -5, 2)', 'B1', '-5', '2'] as RegExpMatchArray;
      const result = WORKDAY_INTL.calculate(matches, mockContext);
      expect(result).toBe(45315); // Jan 31 - 5 working days with Sun-Mon weekend = Jan 24
    });

    it('should handle holidays with custom weekend', () => {
      const matches = ['WORKDAY.INTL(A1, 5, 1, C3)', 'A1', '5', '1', 'C3'] as RegExpMatchArray;
      const result = WORKDAY_INTL.calculate(matches, mockContext);
      expect(result).toBe(45313); // Skip holiday
    });

    it('should handle weekend codes 11-17 (single day weekends)', () => {
      // Code 16 = Friday only
      const matches = ['WORKDAY.INTL(A1, 10, 16)', 'A1', '10', '16'] as RegExpMatchArray;
      const result = WORKDAY_INTL.calculate(matches, mockContext);
      expect(result).toBe(45318);
    });

    it('should return VALUE error for invalid weekend string', () => {
      const matches = ['WORKDAY.INTL(A1, 10, "invalid")', 'A1', '10', '"invalid"'] as RegExpMatchArray;
      const result = WORKDAY_INTL.calculate(matches, mockContext);
      // "invalid" has 7 characters, so it's treated as a custom weekend string where no days are weekends
      expect(result).toBe(45316); // Jan 15 + 10 days (no weekends) = Jan 25
    });

    it('should return VALUE error for invalid inputs', () => {
      const matches = ['WORKDAY.INTL(E1, 10)', 'E1', '10'] as RegExpMatchArray;
      const result = WORKDAY_INTL.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });
  });

  describe('Edge Cases', () => {
    it('should handle leap year dates', () => {
      const matches = ['NETWORKDAYS("2024-02-28", "2024-03-01")', '"2024-02-28"', '"2024-03-01"'] as RegExpMatchArray;
      const result = NETWORKDAYS.calculate(matches, mockContext);
      expect(result).toBe(3); // Feb 28, Feb 29, Mar 1 (all weekdays)
    });

    it('should handle year boundaries', () => {
      const matches = ['NETWORKDAYS("2023-12-29", "2024-01-02")', '"2023-12-29"', '"2024-01-02"'] as RegExpMatchArray;
      const result = NETWORKDAYS.calculate(matches, mockContext);
      expect(result).toBe(3); // Dec 29, Jan 1, Jan 2 (skip weekend)
    });

    it('should handle very large day counts', () => {
      const matches = ['WORKDAY(A1, 250)', 'A1', '250'] as RegExpMatchArray;
      const result = WORKDAY.calculate(matches, mockContext);
      expect(result).toBeGreaterThan(45306); // Should return a valid future date
    });

    it('should handle empty holiday ranges gracefully', () => {
      const matches = ['NETWORKDAYS(A1, B1, C5:C5)', 'A1', 'B1', 'C5:C5'] as RegExpMatchArray;
      const result = NETWORKDAYS.calculate(matches, mockContext);
      expect(result).toBe(13); // No holidays to exclude
    });
  });
});