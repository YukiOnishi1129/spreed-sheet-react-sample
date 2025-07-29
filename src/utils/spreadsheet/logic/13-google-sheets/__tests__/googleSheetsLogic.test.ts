import { describe, it, expect } from 'vitest';
import {
  JOIN, ARRAYFORMULA, QUERY, REGEXMATCH, REGEXEXTRACT, REGEXREPLACE, FLATTEN
} from '../googleSheetsLogic';
import type { FormulaContext } from '../../shared/types';

// Helper function to create FormulaContext
const createContext = (data: (string | number | boolean | null)[][]): FormulaContext => ({
  data: data.map(row => row.map(cell => ({ value: cell }))),
  row: 0,
  col: 0
});

describe('Google Sheets Functions', () => {
  const mockContext = createContext([
    ['Apple', 'Banana', 'Cherry', 'Date'], // fruits
    ['Red', 'Yellow', 'Red', 'Brown'], // colors
    [100, 150, 120, 80], // quantities
    ['A1', 'B2', 'C3', 'D4'], // codes
    ['apple@example.com', 'banana@test.com', 'cherry@domain.org'], // emails
    ['Phone: 123-456-7890', 'Tel: 987-654-3210', 'Mobile: 555-0123'], // phone numbers
    ['2024-01-15', '2024-02-20', '2024-03-25'], // dates
  ]);

  describe('JOIN Function', () => {
    it('should join values with delimiter', () => {
      const matches = ['JOIN(",", A1:D1)', '","', 'A1:D1'] as RegExpMatchArray;
      const result = JOIN.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should join with custom delimiter', () => {
      const matches = ['JOIN(" - ", B1:B4)', '" - "', 'B1:B4'] as RegExpMatchArray;
      const result = JOIN.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should handle multiple arguments', () => {
      const matches = ['JOIN(" ", "Hello", "World")', '" "', '"Hello", "World"'] as RegExpMatchArray;
      const result = JOIN.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });
  });

  describe('ARRAYFORMULA Function', () => {
    it('should apply formula to range', () => {
      const matches = ['ARRAYFORMULA(A1:D1)', 'A1:D1'] as RegExpMatchArray;
      const result = ARRAYFORMULA.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should handle mathematical operations', () => {
      const matches = ['ARRAYFORMULA(C1:C4 * 2)', 'C1:C4 * 2'] as RegExpMatchArray;
      const result = ARRAYFORMULA.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should handle complex expressions', () => {
      const matches = ['ARRAYFORMULA(IF(C1:C4>100, "High", "Low"))', 
        'IF(C1:C4>100, "High", "Low")'] as RegExpMatchArray;
      const result = ARRAYFORMULA.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should handle array operations', () => {
      const matches = ['ARRAYFORMULA(A1:D1 & " - " & B1:B4)', 'A1:D1 & " - " & B1:B4'] as RegExpMatchArray;
      const result = ARRAYFORMULA.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });
  });

  describe('QUERY Function', () => {
    it('should query with select statement', () => {
      const matches = ['QUERY(A1:C4, "SELECT A, C WHERE C > 100")', 
        'A1:C4', '"SELECT A, C WHERE C > 100"'] as RegExpMatchArray;
      const result = QUERY.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should query with order by', () => {
      const matches = ['QUERY(A1:C4, "SELECT * ORDER BY C DESC")', 
        'A1:C4', '"SELECT * ORDER BY C DESC"'] as RegExpMatchArray;
      const result = QUERY.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should query with group by', () => {
      const matches = ['QUERY(B1:C4, "SELECT B, SUM(C) GROUP BY B")', 
        'B1:C4', '"SELECT B, SUM(C) GROUP BY B"'] as RegExpMatchArray;
      const result = QUERY.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should query with headers parameter', () => {
      const matches = ['QUERY(A1:C4, "SELECT A, B", 1)', 
        'A1:C4', '"SELECT A, B"', '1'] as RegExpMatchArray;
      const result = QUERY.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });
  });

  describe('REGEXMATCH Function', () => {
    it('should match email pattern', () => {
      const matches = ['REGEXMATCH(A5, "[a-z]+@[a-z]+\\.[a-z]+")', 
        'A5', '"[a-z]+@[a-z]+\\\\.[a-z]+"'] as RegExpMatchArray;
      const result = REGEXMATCH.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should match phone pattern', () => {
      const matches = ['REGEXMATCH(A6, "\\d{3}-\\d{3}-\\d{4}")', 
        'A6', '"\\\\d{3}-\\\\d{3}-\\\\d{4}"'] as RegExpMatchArray;
      const result = REGEXMATCH.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should handle case insensitive match', () => {
      const matches = ['REGEXMATCH("Hello World", "hello", "i")', 
        '"Hello World"', '"hello"', '"i"'] as RegExpMatchArray;
      const result = REGEXMATCH.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should return false for no match', () => {
      const matches = ['REGEXMATCH(A1, "\\d+")', 'A1', '"\\\\d+"'] as RegExpMatchArray;
      const result = REGEXMATCH.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });
  });

  describe('REGEXEXTRACT Function', () => {
    it('should extract email domain', () => {
      const matches = ['REGEXEXTRACT(A5, "@([a-z]+\\.[a-z]+)")', 
        'A5', '"@([a-z]+\\\\.[a-z]+)"'] as RegExpMatchArray;
      const result = REGEXEXTRACT.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should extract phone area code', () => {
      const matches = ['REGEXEXTRACT(A6, "(\\d{3})-")', 
        'A6', '"(\\\\d{3})-"'] as RegExpMatchArray;
      const result = REGEXEXTRACT.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should extract year from date', () => {
      const matches = ['REGEXEXTRACT(A7, "(\\d{4})")', 
        'A7', '"(\\\\d{4})"'] as RegExpMatchArray;
      const result = REGEXEXTRACT.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should handle multiple capture groups', () => {
      const matches = ['REGEXEXTRACT("John Doe", "(\\w+) (\\w+)")', 
        '"John Doe"', '"(\\\\w+) (\\\\w+)"'] as RegExpMatchArray;
      const result = REGEXEXTRACT.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });
  });

  describe('REGEXREPLACE Function', () => {
    it('should replace phone format', () => {
      const matches = ['REGEXREPLACE(A6, "(\\d{3})-(\\d{3})-(\\d{4})", "($1) $2-$3")', 
        'A6', '"(\\\\d{3})-(\\\\d{3})-(\\\\d{4})"', '"($1) $2-$3"'] as RegExpMatchArray;
      const result = REGEXREPLACE.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should replace email domain', () => {
      const matches = ['REGEXREPLACE(A5, "@[a-z]+\\.[a-z]+", "@newdomain.com")', 
        'A5', '"@[a-z]+\\\\.[a-z]+"', '"@newdomain.com"'] as RegExpMatchArray;
      const result = REGEXREPLACE.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should replace all occurrences', () => {
      const matches = ['REGEXREPLACE("abc123def456", "\\d+", "X")', 
        '"abc123def456"', '"\\\\d+"', '"X"'] as RegExpMatchArray;
      const result = REGEXREPLACE.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should handle backreferences', () => {
      const matches = ['REGEXREPLACE("FirstName LastName", "(\\w+) (\\w+)", "$2, $1")', 
        '"FirstName LastName"', '"(\\\\w+) (\\\\w+)"', '"$2, $1"'] as RegExpMatchArray;
      const result = REGEXREPLACE.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });
  });

  describe('FLATTEN Function', () => {
    it('should flatten 2D range', () => {
      const matches = ['FLATTEN(A1:C4)', 'A1:C4'] as RegExpMatchArray;
      const result = FLATTEN.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should flatten single column', () => {
      const matches = ['FLATTEN(A1:A7)', 'A1:A7'] as RegExpMatchArray;
      const result = FLATTEN.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should flatten single row', () => {
      const matches = ['FLATTEN(A1:D1)', 'A1:D1'] as RegExpMatchArray;
      const result = FLATTEN.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should handle multiple ranges', () => {
      const matches = ['FLATTEN(A1:B2, C3:D4)', 'A1:B2, C3:D4'] as RegExpMatchArray;
      const result = FLATTEN.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });
  });

  describe('Edge Cases', () => {
    it('should handle empty ranges', () => {
      const matches = ['JOIN(",", A10:D10)', '","', 'A10:D10'] as RegExpMatchArray;
      const result = JOIN.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should handle invalid regex patterns', () => {
      const matches = ['REGEXMATCH(A1, "[invalid")', 'A1', '"[invalid"'] as RegExpMatchArray;
      const result = REGEXMATCH.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should handle null values in JOIN', () => {
      const matches = ['JOIN(",", A1:D1)', '","', 'A1:D1'] as RegExpMatchArray;
      const result = JOIN.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should handle complex QUERY', () => {
      const matches = ['QUERY(A1:C7, "SELECT A, B, AVG(C) WHERE B CONTAINS \'Red\' GROUP BY A, B PIVOT D")', 
        'A1:C7', '"SELECT A, B, AVG(C) WHERE B CONTAINS \'Red\' GROUP BY A, B PIVOT D"'] as RegExpMatchArray;
      const result = QUERY.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should handle all functions returning consistent error', () => {
      const functions = [
        JOIN, ARRAYFORMULA, QUERY, REGEXMATCH, REGEXEXTRACT, REGEXREPLACE, FLATTEN
      ];
      
      for (const func of functions) {
        const matches = ['FUNC("param")', '"param"'] as RegExpMatchArray;
        const result = func.calculate(matches, mockContext);
        expect(result).toBe('#N/A - Google Sheets specific functions not supported');
      }
    });
  });
});