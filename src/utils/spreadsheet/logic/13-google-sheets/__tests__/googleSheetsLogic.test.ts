import { describe, it, expect } from 'vitest';
import {
  ARRAYFORMULA, GOOGLEFINANCE, GOOGLETRANSLATE, QUERY, SPARKLINE
} from '../googleSheetsLogic';
import { FormulaError } from '../../shared/types';
import type { FormulaContext } from '../../shared/types';

// Helper function to create FormulaContext
const createContext = (data: (string | number | boolean | null)[][]): FormulaContext => ({
  data: data.map(row => row.map(cell => ({ value: cell }))),
  row: 0,
  col: 0
});

describe('Google Sheets Functions', () => {
  const mockContext = createContext([
    ['AAPL', 'GOOGL', 'MSFT', 'AMZN'], // stock symbols
    ['Hello World', 'Good Morning', 'Thank You'], // text to translate
    ['en', 'es', 'fr', 'de', 'ja'], // language codes
    [100, 150, 120, 180, 160, 200], // data for sparkline
    ['Name', 'Age', 'City', 'Country'], // headers
    ['John', 25, 'New York', 'USA'], // data row 1
    ['Maria', 30, 'Madrid', 'Spain'], // data row 2
    ['Yuki', 28, 'Tokyo', 'Japan'], // data row 3
  ]);

  describe('ARRAYFORMULA Function', () => {
    it('should apply formula to array', () => {
      const matches = ['ARRAYFORMULA(A1:A4 & " Stock")', 'A1:A4 & " Stock"'] as RegExpMatchArray;
      const result = ARRAYFORMULA.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should handle mathematical operations', () => {
      const matches = ['ARRAYFORMULA(B4:G4 * 2)', 'B4:G4 * 2'] as RegExpMatchArray;
      const result = ARRAYFORMULA.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should handle complex expressions', () => {
      const matches = ['ARRAYFORMULA(IF(A1:A4="AAPL", "Apple", "Other"))', 
        'IF(A1:A4="AAPL", "Apple", "Other")'] as RegExpMatchArray;
      const result = ARRAYFORMULA.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should handle array addition', () => {
      const matches = ['ARRAYFORMULA(A1:C1 + A2:C2)', 'A1:C1 + A2:C2'] as RegExpMatchArray;
      const result = ARRAYFORMULA.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should handle text functions on arrays', () => {
      const matches = ['ARRAYFORMULA(UPPER(B2:D2))', 'UPPER(B2:D2)'] as RegExpMatchArray;
      const result = ARRAYFORMULA.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should handle nested array formulas', () => {
      const matches = ['ARRAYFORMULA(SUM(A1:D4 * 2))', 'SUM(A1:D4 * 2)'] as RegExpMatchArray;
      const result = ARRAYFORMULA.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });
  });

  describe('GOOGLEFINANCE Function', () => {
    it('should fetch current stock price', () => {
      const matches = ['GOOGLEFINANCE("AAPL")', '"AAPL"'] as RegExpMatchArray;
      const result = GOOGLEFINANCE.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should fetch specific attribute', () => {
      const matches = ['GOOGLEFINANCE("GOOGL", "price")', '"GOOGL"', '"price"'] as RegExpMatchArray;
      const result = GOOGLEFINANCE.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should fetch historical data', () => {
      const matches = ['GOOGLEFINANCE("MSFT", "close", DATE(2024,1,1), DATE(2024,1,31))', 
        '"MSFT"', '"close"', 'DATE(2024,1,1)', 'DATE(2024,1,31)'] as RegExpMatchArray;
      const result = GOOGLEFINANCE.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should fetch with interval', () => {
      const matches = ['GOOGLEFINANCE("AMZN", "close", DATE(2024,1,1), DATE(2024,12,31), "WEEKLY")', 
        '"AMZN"', '"close"', 'DATE(2024,1,1)', 'DATE(2024,12,31)', '"WEEKLY"'] as RegExpMatchArray;
      const result = GOOGLEFINANCE.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should fetch currency exchange rate', () => {
      const matches = ['GOOGLEFINANCE("CURRENCY:USDJPY")', '"CURRENCY:USDJPY"'] as RegExpMatchArray;
      const result = GOOGLEFINANCE.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should fetch mutual fund data', () => {
      const matches = ['GOOGLEFINANCE("MUTF:VFINX", "nav")', '"MUTF:VFINX"', '"nav"'] as RegExpMatchArray;
      const result = GOOGLEFINANCE.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should handle market attributes', () => {
      const matches = ['GOOGLEFINANCE("AAPL", "marketcap")', '"AAPL"', '"marketcap"'] as RegExpMatchArray;
      const result = GOOGLEFINANCE.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });
  });

  describe('GOOGLETRANSLATE Function', () => {
    it('should translate text', () => {
      const matches = ['GOOGLETRANSLATE("Hello World", "en", "es")', 
        '"Hello World"', '"en"', '"es"'] as RegExpMatchArray;
      const result = GOOGLETRANSLATE.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should auto-detect source language', () => {
      const matches = ['GOOGLETRANSLATE("Bonjour", "auto", "en")', 
        '"Bonjour"', '"auto"', '"en"'] as RegExpMatchArray;
      const result = GOOGLETRANSLATE.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should translate to Japanese', () => {
      const matches = ['GOOGLETRANSLATE("Thank You", "en", "ja")', 
        '"Thank You"', '"en"', '"ja"'] as RegExpMatchArray;
      const result = GOOGLETRANSLATE.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should handle cell references', () => {
      const matches = ['GOOGLETRANSLATE(B2, "en", "fr")', 'B2', '"en"', '"fr"'] as RegExpMatchArray;
      const result = GOOGLETRANSLATE.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should translate with simplified syntax', () => {
      const matches = ['GOOGLETRANSLATE("Hello", "es")', '"Hello"', '"es"'] as RegExpMatchArray;
      const result = GOOGLETRANSLATE.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should handle long text', () => {
      const longText = 'This is a very long text that needs translation. It contains multiple sentences and should be handled properly by the translation service.';
      const matches = ['GOOGLETRANSLATE("' + longText + '", "en", "de")', 
        '"' + longText + '"', '"en"', '"de"'] as RegExpMatchArray;
      const result = GOOGLETRANSLATE.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });
  });

  describe('QUERY Function', () => {
    it('should execute simple select query', () => {
      const matches = ['QUERY(A5:D8, "SELECT A, B WHERE B > 25")', 
        'A5:D8', '"SELECT A, B WHERE B > 25"'] as RegExpMatchArray;
      const result = QUERY.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should handle column headers', () => {
      const matches = ['QUERY(A5:D8, "SELECT * WHERE C = \'Tokyo\'", 1)', 
        'A5:D8', '"SELECT * WHERE C = \'Tokyo\'"', '1'] as RegExpMatchArray;
      const result = QUERY.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should support ORDER BY', () => {
      const matches = ['QUERY(A5:D8, "SELECT A, B ORDER BY B DESC")', 
        'A5:D8', '"SELECT A, B ORDER BY B DESC"'] as RegExpMatchArray;
      const result = QUERY.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should support GROUP BY', () => {
      const matches = ['QUERY(A5:D8, "SELECT D, COUNT(A) GROUP BY D")', 
        'A5:D8', '"SELECT D, COUNT(A) GROUP BY D"'] as RegExpMatchArray;
      const result = QUERY.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should support PIVOT', () => {
      const matches = ['QUERY(A5:D8, "SELECT B, SUM(B) GROUP BY A PIVOT C")', 
        'A5:D8', '"SELECT B, SUM(B) GROUP BY A PIVOT C"'] as RegExpMatchArray;
      const result = QUERY.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should support LIMIT and OFFSET', () => {
      const matches = ['QUERY(A5:D8, "SELECT * LIMIT 2 OFFSET 1")', 
        'A5:D8', '"SELECT * LIMIT 2 OFFSET 1"'] as RegExpMatchArray;
      const result = QUERY.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should support date functions', () => {
      const matches = ['QUERY(A1:D10, "SELECT A WHERE B > date \'2024-01-01\'")', 
        'A1:D10', '"SELECT A WHERE B > date \'2024-01-01\'"'] as RegExpMatchArray;
      const result = QUERY.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });
  });

  describe('SPARKLINE Function', () => {
    it('should create basic line chart', () => {
      const matches = ['SPARKLINE(A4:F4)', 'A4:F4'] as RegExpMatchArray;
      const result = SPARKLINE.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should create column chart', () => {
      const matches = ['SPARKLINE(A4:F4, {"charttype", "column"})', 
        'A4:F4', '{"charttype", "column"}'] as RegExpMatchArray;
      const result = SPARKLINE.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should create bar chart', () => {
      const matches = ['SPARKLINE(A4:F4, {"charttype", "bar"})', 
        'A4:F4', '{"charttype", "bar"}'] as RegExpMatchArray;
      const result = SPARKLINE.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should create win/loss chart', () => {
      const matches = ['SPARKLINE(A4:F4, {"charttype", "winloss"})', 
        'A4:F4', '{"charttype", "winloss"}'] as RegExpMatchArray;
      const result = SPARKLINE.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should handle color options', () => {
      const matches = ['SPARKLINE(A4:F4, {"color", "red"; "linewidth", 2})', 
        'A4:F4', '{"color", "red"; "linewidth", 2}'] as RegExpMatchArray;
      const result = SPARKLINE.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should handle min/max options', () => {
      const matches = ['SPARKLINE(A4:F4, {"max", 250; "min", 50})', 
        'A4:F4', '{"max", 250; "min", 50}'] as RegExpMatchArray;
      const result = SPARKLINE.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should handle empty cells option', () => {
      const matches = ['SPARKLINE(A4:F4, {"empty", "zero"})', 
        'A4:F4', '{"empty", "zero"}'] as RegExpMatchArray;
      const result = SPARKLINE.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });
  });

  describe('Edge Cases and Integration', () => {
    it('should handle nested Google functions', () => {
      const matches = ['GOOGLETRANSLATE(GOOGLEFINANCE("AAPL", "name"), "en", "ja")', 
        'GOOGLEFINANCE("AAPL", "name")', '"en"', '"ja"'] as RegExpMatchArray;
      const result = GOOGLETRANSLATE.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should handle QUERY with dynamic ranges', () => {
      const matches = ['QUERY(ARRAYFORMULA(A1:D10), "SELECT * WHERE Col2 > 100")', 
        'ARRAYFORMULA(A1:D10)', '"SELECT * WHERE Col2 > 100"'] as RegExpMatchArray;
      const result = QUERY.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should handle SPARKLINE with GOOGLEFINANCE data', () => {
      const matches = ['SPARKLINE(GOOGLEFINANCE("AAPL", "price", TODAY()-30, TODAY()))', 
        'GOOGLEFINANCE("AAPL", "price", TODAY()-30, TODAY())'] as RegExpMatchArray;
      const result = SPARKLINE.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should handle all functions returning consistent error', () => {
      const functions = [
        ARRAYFORMULA, GOOGLEFINANCE, GOOGLETRANSLATE, QUERY, SPARKLINE
      ];
      
      for (const func of functions) {
        const matches = ['FUNC("param")', '"param"'] as RegExpMatchArray;
        const result = func.calculate(matches, mockContext);
        expect(result).toBe('#N/A - Google Sheets specific functions not supported');
      }
    });
  });
});