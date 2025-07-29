// dateUtils.test.ts
import { describe, it, expect, beforeEach } from 'vitest';
import {
  setGlobalDateConfig,
  getGlobalDateConfig,
  createDateInTimezone,
  formatDate,
  getWeekday,
  getWeekNumber,
  getISOWeekNumber,
  parseTimeString,
  getDateInTimezone,
  parseDate,
  excelSerialToDate,
  dateToExcelSerial,
  isValidDate,
  diffDays,
  diffMonths,
  diffYears,
  diffMonthsYM,
  diffDaysYD,
  diffDaysMD,
  now,
  today,
  TIMEZONES
} from '../dateUtils';

describe('Date Utils', () => {
  // Reset global config before each test
  beforeEach(() => {
    setGlobalDateConfig({
      timezone: Intl.DateTimeFormat().resolvedOptions().timeZone,
      locale: 'ja-JP'
    });
  });

  describe('Global Date Configuration', () => {
    it('should set and get global date config', () => {
      setGlobalDateConfig({
        timezone: TIMEZONES.US_EASTERN,
        locale: 'en-US'
      });
      
      const config = getGlobalDateConfig();
      expect(config.timezone).toBe(TIMEZONES.US_EASTERN);
      expect(config.locale).toBe('en-US');
    });

    it('should merge partial config updates', () => {
      setGlobalDateConfig({ timezone: TIMEZONES.UK });
      const config1 = getGlobalDateConfig();
      expect(config1.timezone).toBe(TIMEZONES.UK);
      expect(config1.locale).toBe('ja-JP'); // Should keep existing locale
      
      setGlobalDateConfig({ locale: 'fr-FR' });
      const config2 = getGlobalDateConfig();
      expect(config2.timezone).toBe(TIMEZONES.UK); // Should keep existing timezone
      expect(config2.locale).toBe('fr-FR');
    });
  });

  describe('createDateInTimezone', () => {
    it('should create date in default timezone', () => {
      const date = createDateInTimezone(2023, 3, 15);
      expect(date.getFullYear()).toBe(2023);
      expect(date.getMonth()).toBe(2); // 0-indexed
      expect(date.getDate()).toBe(15);
    });

    it('should create date in specific timezone', () => {
      const date = createDateInTimezone(2023, 3, 15, { timezone: TIMEZONES.US_PACIFIC });
      expect(date.getFullYear()).toBe(2023);
      expect(date.getMonth()).toBe(2);
      expect(date.getDate()).toBe(15);
    });

    it('should handle timezone differences correctly', () => {
      // Create same date in different timezones
      const japanDate = createDateInTimezone(2023, 1, 1, { timezone: TIMEZONES.JAPAN });
      const nyDate = createDateInTimezone(2023, 1, 1, { timezone: TIMEZONES.US_EASTERN });
      
      // Both should represent Jan 1, 2023 in their respective timezones
      expect(japanDate.getFullYear()).toBe(2023);
      expect(japanDate.getMonth()).toBe(0);
      expect(japanDate.getDate()).toBe(1);
      
      expect(nyDate.getFullYear()).toBe(2023);
      expect(nyDate.getMonth()).toBe(0);
      expect(nyDate.getDate()).toBe(1);
    });
  });

  describe('formatDate', () => {
    it('should format date with default locale', () => {
      const date = new Date(2023, 2, 15); // March 15, 2023
      const formatted = formatDate(date);
      expect(formatted).toMatch(/2023.*03.*15/); // Should contain year, month, day
    });

    it('should format date with custom format', () => {
      const date = new Date(2023, 2, 15);
      const formatted = formatDate(date, 'YYYY-MM-DD');
      expect(formatted).toBe('2023-03-15');
    });

    it('should format date with different locale', () => {
      const date = new Date(2023, 2, 15);
      const formatted = formatDate(date, undefined, { locale: 'en-US' });
      expect(formatted).toMatch(/03.*15.*2023/); // US format
    });

    it('should handle custom format patterns', () => {
      const date = new Date(2023, 11, 25); // December 25, 2023
      const formatted = formatDate(date, 'DD/MM/YYYY');
      expect(formatted).toBe('25/12/2023');
    });
  });

  describe('getWeekday', () => {
    it('should get weekday with Monday as start (default)', () => {
      // March 15, 2023 is Wednesday
      const date = new Date(2023, 2, 15);
      const weekday = getWeekday(date);
      expect(weekday).toBe(3); // Wednesday = 3 when Monday = 1
    });

    it('should get weekday with Sunday as start', () => {
      const date = new Date(2023, 2, 15); // Wednesday
      const weekday = getWeekday(date, 0);
      expect(weekday).toBe(4); // Wednesday = 4 when Sunday = 1
    });

    it('should handle Sunday correctly', () => {
      const date = new Date(2023, 2, 12); // Sunday
      expect(getWeekday(date, 1)).toBe(7); // Sunday = 7 when Monday = 1
      expect(getWeekday(date, 0)).toBe(1); // Sunday = 1 when Sunday = 1
    });

    it('should respect timezone config', () => {
      const date = new Date(2023, 2, 15);
      const weekday = getWeekday(date, 1, { timezone: TIMEZONES.JAPAN });
      expect(typeof weekday).toBe('number');
      expect(weekday).toBeGreaterThanOrEqual(1);
      expect(weekday).toBeLessThanOrEqual(7);
    });
  });

  describe('getWeekNumber', () => {
    it('should calculate week number with Sunday start (type 1)', () => {
      const date = new Date(2023, 0, 1); // Jan 1, 2023 (Sunday)
      const weekNum = getWeekNumber(date, 1);
      expect(weekNum).toBe(1);
    });

    it('should calculate week number with Monday start (type 2)', () => {
      const date = new Date(2023, 0, 2); // Jan 2, 2023 (Monday)
      const weekNum = getWeekNumber(date, 2);
      expect(weekNum).toBe(1);
    });

    it('should handle mid-year dates', () => {
      const date = new Date(2023, 6, 15); // July 15, 2023
      const weekNum = getWeekNumber(date);
      expect(weekNum).toBeGreaterThan(25);
      expect(weekNum).toBeLessThan(35);
    });

    it('should handle year boundaries', () => {
      const date = new Date(2023, 11, 31); // Dec 31, 2023
      const weekNum = getWeekNumber(date);
      expect(weekNum).toBeGreaterThanOrEqual(52);
      expect(weekNum).toBeLessThanOrEqual(53);
    });
  });

  describe('getISOWeekNumber', () => {
    it('should calculate ISO week number correctly', () => {
      // January 4 is always in week 1
      const date = new Date(2023, 0, 4); // Jan 4, 2023
      const weekNum = getISOWeekNumber(date);
      expect(weekNum).toBe(1);
    });

    it('should handle year boundaries', () => {
      // Dec 31, 2022 is in ISO week 52 of 2022
      const date1 = new Date(2022, 11, 31);
      const weekNum1 = getISOWeekNumber(date1);
      expect(weekNum1).toBe(52);
      
      // Jan 1, 2023 is in ISO week 52 of 2022
      const date2 = new Date(2023, 0, 1);
      const weekNum2 = getISOWeekNumber(date2);
      expect(weekNum2).toBe(52);
    });

    it('should handle week 53', () => {
      // Some years have 53 weeks
      const date = new Date(2020, 11, 31); // Dec 31, 2020
      const weekNum = getISOWeekNumber(date);
      expect(weekNum).toBeGreaterThanOrEqual(52);
      expect(weekNum).toBeLessThanOrEqual(53);
    });

    it('should handle first week of year correctly', () => {
      // First Thursday of 2023 is Jan 5
      const date = new Date(2023, 0, 5);
      const weekNum = getISOWeekNumber(date);
      expect(weekNum).toBe(1);
    });
  });

  describe('parseTimeString', () => {
    it('should parse time with hours and minutes', () => {
      const time = parseTimeString('14:30');
      expect(time).toBeCloseTo(0.604166667, 8);
    });

    it('should parse time with hours, minutes, and seconds', () => {
      const time = parseTimeString('14:30:45');
      expect(time).toBeCloseTo(0.604687500, 8);
    });

    it('should handle single digit values', () => {
      const time = parseTimeString('2:5:9');
      expect(time).toBeCloseTo(0.086921296, 8);
    });

    it('should return null for invalid format', () => {
      expect(parseTimeString('invalid')).toBeNull();
      expect(parseTimeString('25:00:00')).toBeNull();
      expect(parseTimeString('12:60:00')).toBeNull();
      expect(parseTimeString('12:30:60')).toBeNull();
    });

    it('should handle midnight and noon', () => {
      expect(parseTimeString('00:00:00')).toBe(0);
      expect(parseTimeString('12:00:00')).toBe(0.5);
      expect(parseTimeString('23:59:59')).toBeCloseTo(0.999988426, 8);
    });
  });

  describe('getDateInTimezone', () => {
    it('should get date in specific timezone', () => {
      const date = new Date(2023, 2, 15, 12, 30, 45);
      const tokyoDate = getDateInTimezone(date, TIMEZONES.JAPAN);
      
      expect(tokyoDate instanceof Date).toBe(true);
      expect(tokyoDate.getFullYear()).toBe(2023);
    });

    it('should handle timezone conversions', () => {
      const date = new Date('2023-03-15T00:00:00Z'); // UTC midnight
      const nyDate = getDateInTimezone(date, TIMEZONES.US_EASTERN);
      const tokyoDate = getDateInTimezone(date, TIMEZONES.JAPAN);
      
      // Both should be valid dates
      expect(isValidDate(nyDate)).toBe(true);
      expect(isValidDate(tokyoDate)).toBe(true);
    });
  });

  describe('parseDate', () => {
    it('should parse Date objects', () => {
      const input = new Date(2023, 2, 15);
      const result = parseDate(input);
      expect(result).toEqual(input);
    });

    it('should parse Excel serial numbers', () => {
      const result = parseDate(44970); // Excel serial for 2023-02-13
      expect(result).toBeInstanceOf(Date);
      expect(result?.getFullYear()).toBe(2023);
      expect(result?.getMonth()).toBe(1); // February
    });

    it('should parse ISO date strings', () => {
      const result = parseDate('2023-03-15');
      expect(result?.getFullYear()).toBe(2023);
      expect(result?.getMonth()).toBe(2); // March
      expect(result?.getDate()).toBe(15);
    });

    it('should parse MM/DD/YYYY format', () => {
      const result = parseDate('03/15/2023');
      expect(result?.getFullYear()).toBe(2023);
      expect(result?.getMonth()).toBe(2); // March
      expect(result?.getDate()).toBe(15);
    });

    it('should return null for invalid inputs', () => {
      expect(parseDate(null)).toBeNull();
      expect(parseDate(undefined)).toBeNull();
      expect(parseDate('')).toBeNull();
      expect(parseDate('invalid')).toBeNull();
      expect(parseDate({})).toBeNull();
    });

    it('should handle edge cases', () => {
      // Invalid date object
      const invalidDate = new Date('invalid');
      expect(parseDate(invalidDate)).toBeNull();
      
      // Zero Excel serial
      const zeroSerial = parseDate(0);
      expect(zeroSerial).toBeInstanceOf(Date);
    });
  });

  describe('Excel Serial Conversion', () => {
    it('should convert Excel serial to date', () => {
      const date = excelSerialToDate(44927); // Jan 1, 2023
      expect(date.getFullYear()).toBe(2023);
      expect(date.getMonth()).toBe(0); // January
      expect(date.getDate()).toBe(1);
    });

    it('should convert date to Excel serial', () => {
      const date = new Date(2023, 0, 1); // Jan 1, 2023
      const serial = dateToExcelSerial(date);
      expect(serial).toBe(44927);
    });

    it('should handle Excel epoch', () => {
      const date = excelSerialToDate(1); // Jan 1, 1900
      expect(date.getFullYear()).toBe(1900);
      expect(date.getMonth()).toBe(0);
      expect(date.getDate()).toBe(1);
    });

    it('should round trip conversion', () => {
      const originalDate = new Date(2023, 2, 15);
      const serial = dateToExcelSerial(originalDate);
      const convertedDate = excelSerialToDate(serial);
      
      expect(convertedDate.getFullYear()).toBe(originalDate.getFullYear());
      expect(convertedDate.getMonth()).toBe(originalDate.getMonth());
      expect(convertedDate.getDate()).toBe(originalDate.getDate());
    });
  });

  describe('isValidDate', () => {
    it('should validate valid dates', () => {
      expect(isValidDate(new Date(2023, 2, 15))).toBe(true);
      expect(isValidDate(new Date('2023-03-15'))).toBe(true);
      expect(isValidDate(new Date(Date.now()))).toBe(true);
    });

    it('should reject invalid dates', () => {
      expect(isValidDate(new Date('invalid'))).toBe(false);
      expect(isValidDate(new Date(NaN))).toBe(false);
      expect(isValidDate('not a date' as any)).toBe(false);
      expect(isValidDate(null as any)).toBe(false);
    });
  });

  describe('Date Difference Functions', () => {
    describe('diffDays', () => {
      it('should calculate positive day difference', () => {
        const start = new Date(2023, 2, 1);
        const end = new Date(2023, 2, 15);
        expect(diffDays(start, end)).toBe(14);
      });

      it('should calculate negative day difference', () => {
        const start = new Date(2023, 2, 15);
        const end = new Date(2023, 2, 1);
        expect(diffDays(start, end)).toBe(-14);
      });

      it('should handle same date', () => {
        const date = new Date(2023, 2, 15);
        expect(diffDays(date, date)).toBe(0);
      });

      it('should ignore time component', () => {
        const start = new Date(2023, 2, 1, 10, 30, 45);
        const end = new Date(2023, 2, 1, 20, 45, 30);
        expect(diffDays(start, end)).toBe(0);
      });
    });

    describe('diffMonths', () => {
      it('should calculate month difference within same year', () => {
        const start = new Date(2023, 0, 15); // Jan
        const end = new Date(2023, 3, 20); // Apr
        expect(diffMonths(start, end)).toBe(3);
      });

      it('should calculate month difference across years', () => {
        const start = new Date(2022, 10, 15); // Nov 2022
        const end = new Date(2023, 1, 20); // Feb 2023
        expect(diffMonths(start, end)).toBe(3);
      });

      it('should handle negative differences', () => {
        const start = new Date(2023, 3, 15);
        const end = new Date(2023, 0, 20);
        expect(diffMonths(start, end)).toBe(-3);
      });
    });

    describe('diffYears', () => {
      it('should calculate year difference', () => {
        const start = new Date(2020, 5, 15);
        const end = new Date(2023, 5, 15);
        expect(diffYears(start, end)).toBe(3);
      });

      it('should consider month and day', () => {
        const start = new Date(2020, 5, 15);
        const end = new Date(2023, 4, 14); // One day before 3 full years
        expect(diffYears(start, end)).toBe(2);
      });

      it('should handle same year', () => {
        const start = new Date(2023, 0, 1);
        const end = new Date(2023, 11, 31);
        expect(diffYears(start, end)).toBe(0);
      });
    });

    describe('diffMonthsYM', () => {
      it('should calculate months ignoring years', () => {
        const start = new Date(2022, 1, 15); // Feb
        const end = new Date(2023, 4, 20); // May
        expect(diffMonthsYM(start, end)).toBe(3);
      });

      it('should handle day consideration', () => {
        const start = new Date(2023, 1, 20); // Feb 20
        const end = new Date(2023, 4, 15); // May 15
        expect(diffMonthsYM(start, end)).toBe(2); // Only 2 full months
      });

      it('should wrap around year boundary', () => {
        const start = new Date(2022, 10, 15); // Nov
        const end = new Date(2023, 1, 20); // Feb
        expect(diffMonthsYM(start, end)).toBe(3);
      });
    });

    describe('diffDaysYD', () => {
      it('should calculate days ignoring years', () => {
        const start = new Date(2022, 0, 15); // Jan 15
        const end = new Date(2023, 2, 10); // Mar 10
        expect(diffDaysYD(start, end)).toBe(54);
      });

      it('should handle dates in same year', () => {
        const start = new Date(2023, 0, 15);
        const end = new Date(2023, 2, 10);
        expect(diffDaysYD(start, end)).toBe(54);
      });

      it('should handle negative case with wrap', () => {
        const start = new Date(2022, 11, 15); // Dec 15
        const end = new Date(2023, 0, 10); // Jan 10
        expect(diffDaysYD(start, end)).toBe(26);
      });
    });

    describe('diffDaysMD', () => {
      it('should calculate days ignoring months', () => {
        const start = new Date(2023, 0, 10);
        const end = new Date(2023, 1, 25);
        expect(diffDaysMD(start, end)).toBe(15);
      });

      it('should handle wrap around month', () => {
        const start = new Date(2023, 0, 25);
        const end = new Date(2023, 1, 10);
        expect(diffDaysMD(start, end)).toBe(16); // 31-25+10
      });

      it('should handle same day of month', () => {
        const start = new Date(2023, 0, 15);
        const end = new Date(2023, 1, 15);
        expect(diffDaysMD(start, end)).toBe(0);
      });

      it('should handle February edge cases', () => {
        const start = new Date(2023, 0, 30); // Jan 30
        const end = new Date(2023, 1, 28); // Feb 28 (non-leap)
        expect(diffDaysMD(start, end)).toBe(29); // 31-30+28
      });
    });
  });

  describe('now and today', () => {
    it('should return current date/time', () => {
      const before = new Date();
      const nowResult = now();
      const after = new Date();
      
      expect(nowResult.getTime()).toBeGreaterThanOrEqual(before.getTime());
      expect(nowResult.getTime()).toBeLessThanOrEqual(after.getTime());
    });

    it('should return today with no time component', () => {
      const todayResult = today();
      expect(todayResult.getHours()).toBe(0);
      expect(todayResult.getMinutes()).toBe(0);
      expect(todayResult.getSeconds()).toBe(0);
      expect(todayResult.getMilliseconds()).toBe(0);
    });

    it('should respect timezone config', () => {
      const nowJapan = now({ timezone: TIMEZONES.JAPAN });
      const nowNY = now({ timezone: TIMEZONES.US_EASTERN });
      
      expect(nowJapan instanceof Date).toBe(true);
      expect(nowNY instanceof Date).toBe(true);
    });

    it('should return different times for different timezones', () => {
      const todayJapan = today({ timezone: TIMEZONES.JAPAN });
      const todayHawaii = today({ timezone: TIMEZONES.US_HAWAII });
      
      // Both should be valid dates with no time component
      expect(todayJapan.getHours()).toBe(0);
      expect(todayHawaii.getHours()).toBe(0);
    });
  });

  describe('TIMEZONES constant', () => {
    it('should contain expected timezone values', () => {
      expect(TIMEZONES.JAPAN).toBe('Asia/Tokyo');
      expect(TIMEZONES.US_EASTERN).toBe('America/New_York');
      expect(TIMEZONES.US_PACIFIC).toBe('America/Los_Angeles');
      expect(TIMEZONES.UK).toBe('Europe/London');
      expect(TIMEZONES.UTC).toBe('UTC');
    });

    it('should have all expected timezone properties', () => {
      const expectedTimezones = [
        'JAPAN', 'US_EASTERN', 'US_CENTRAL', 'US_MOUNTAIN', 
        'US_PACIFIC', 'US_ALASKA', 'US_HAWAII', 'UK', 
        'FRANCE', 'GERMANY', 'AUSTRALIA_SYDNEY', 'CHINA', 
        'KOREA', 'UTC'
      ];
      
      expectedTimezones.forEach(tz => {
        expect(TIMEZONES).toHaveProperty(tz);
      });
    });
  });
});