import { describe, it, expect } from 'vitest';
import {
  DETECTLANGUAGE, IMAGE, MAP, SPLIT, UNIQUE, JOIN, FLATTEN, TRANSPOSE
} from '../googleSheetsExtraLogic';
import { FormulaError } from '../../shared/types';
import type { FormulaContext } from '../../shared/types';

// Helper function to create FormulaContext
const createContext = (data: (string | number | boolean | null)[][]): FormulaContext => ({
  data: data.map(row => row.map(cell => ({ value: cell }))),
  row: 0,
  col: 0
});

describe('Google Sheets Extra Functions', () => {
  const mockContext = createContext([
    ['Hello World', 'Bonjour', 'こんにちは', 'Hola'], // texts in different languages
    ['https://example.com/image.png', 'https://example.com/logo.jpg'], // image URLs
    ['New York, NY', 'Tokyo, Japan', 'London, UK'], // locations
    ['apple,banana,orange', 'red;green;blue', 'one|two|three'], // delimited strings
    ['Apple', 'Banana', 'Apple', 'Orange', 'Banana', 'Apple'], // data with duplicates
    [1, 2, 3, 4, 5], // numbers
    ['A', 'B', 'C'], // letters
    [[1, 2], [3, 4], [5, 6]], // nested arrays representation
  ]);

  describe('DETECTLANGUAGE Function', () => {
    it('should detect English text', () => {
      const matches = ['DETECTLANGUAGE("Hello World")', '"Hello World"'] as RegExpMatchArray;
      const result = DETECTLANGUAGE.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should detect French text', () => {
      const matches = ['DETECTLANGUAGE("Bonjour le monde")', '"Bonjour le monde"'] as RegExpMatchArray;
      const result = DETECTLANGUAGE.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should detect Japanese text', () => {
      const matches = ['DETECTLANGUAGE("こんにちは世界")', '"こんにちは世界"'] as RegExpMatchArray;
      const result = DETECTLANGUAGE.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should detect Spanish text', () => {
      const matches = ['DETECTLANGUAGE("Hola mundo")', '"Hola mundo"'] as RegExpMatchArray;
      const result = DETECTLANGUAGE.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should handle mixed language text', () => {
      const matches = ['DETECTLANGUAGE("Hello こんにちは Bonjour")', '"Hello こんにちは Bonjour"'] as RegExpMatchArray;
      const result = DETECTLANGUAGE.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should handle cell references', () => {
      const matches = ['DETECTLANGUAGE(C1)', 'C1'] as RegExpMatchArray;
      const result = DETECTLANGUAGE.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });
  });

  describe('IMAGE Function', () => {
    it('should insert image with default mode', () => {
      const matches = ['IMAGE("https://example.com/image.png")', 
        '"https://example.com/image.png"'] as RegExpMatchArray;
      const result = IMAGE.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should insert image with fit mode', () => {
      const matches = ['IMAGE("https://example.com/logo.jpg", 1)', 
        '"https://example.com/logo.jpg"', '1'] as RegExpMatchArray;
      const result = IMAGE.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should insert image with stretch mode', () => {
      const matches = ['IMAGE("https://example.com/banner.png", 2)', 
        '"https://example.com/banner.png"', '2'] as RegExpMatchArray;
      const result = IMAGE.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should insert image with original size', () => {
      const matches = ['IMAGE("https://example.com/photo.jpg", 3)', 
        '"https://example.com/photo.jpg"', '3'] as RegExpMatchArray;
      const result = IMAGE.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should insert image with custom size', () => {
      const matches = ['IMAGE("https://example.com/icon.png", 4, 100, 100)', 
        '"https://example.com/icon.png"', '4', '100', '100'] as RegExpMatchArray;
      const result = IMAGE.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should handle data URIs', () => {
      const matches = ['IMAGE("data:image/png;base64,iVBORw0KGgo...")', 
        '"data:image/png;base64,iVBORw0KGgo..."'] as RegExpMatchArray;
      const result = IMAGE.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });
  });

  describe('MAP Function', () => {
    it('should create map with location', () => {
      const matches = ['MAP("New York, NY")', '"New York, NY"'] as RegExpMatchArray;
      const result = MAP.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should create map with zoom level', () => {
      const matches = ['MAP("Tokyo, Japan", 10)', '"Tokyo, Japan"', '10'] as RegExpMatchArray;
      const result = MAP.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should create map with coordinates', () => {
      const matches = ['MAP("40.7128,-74.0060", 12)', '"40.7128,-74.0060"', '12'] as RegExpMatchArray;
      const result = MAP.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should handle address with special characters', () => {
      const matches = ['MAP("123 Main St, Apt #4, New York")', 
        '"123 Main St, Apt #4, New York"'] as RegExpMatchArray;
      const result = MAP.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should handle international addresses', () => {
      const matches = ['MAP("東京都渋谷区")', '"東京都渋谷区"'] as RegExpMatchArray;
      const result = MAP.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should handle cell references', () => {
      const matches = ['MAP(A3)', 'A3'] as RegExpMatchArray;
      const result = MAP.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });
  });

  describe('SPLIT Function', () => {
    it('should split comma-delimited string', () => {
      const matches = ['SPLIT("apple,banana,orange", ",")', 
        '"apple,banana,orange"', '","'] as RegExpMatchArray;
      const result = SPLIT.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should split semicolon-delimited string', () => {
      const matches = ['SPLIT("red;green;blue", ";")', 
        '"red;green;blue"', '";"'] as RegExpMatchArray;
      const result = SPLIT.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should split by space', () => {
      const matches = ['SPLIT("Hello World Test", " ")', 
        '"Hello World Test"', '" "'] as RegExpMatchArray;
      const result = SPLIT.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should split by multiple delimiters', () => {
      const matches = ['SPLIT("apple,banana;orange|grape", ",;|", TRUE)', 
        '"apple,banana;orange|grape"', '",;|"', 'TRUE'] as RegExpMatchArray;
      const result = SPLIT.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should handle empty segments', () => {
      const matches = ['SPLIT("a,,b,,c", ",", FALSE)', 
        '"a,,b,,c"', '","', 'FALSE'] as RegExpMatchArray;
      const result = SPLIT.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should split cell reference', () => {
      const matches = ['SPLIT(A4, ",")', 'A4', '","'] as RegExpMatchArray;
      const result = SPLIT.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });
  });

  describe('UNIQUE Function', () => {
    it('should return unique values from range', () => {
      const matches = ['UNIQUE(A6:F6)', 'A6:F6'] as RegExpMatchArray;
      const result = UNIQUE.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should handle unique by columns', () => {
      const matches = ['UNIQUE(A1:D4, FALSE)', 'A1:D4', 'FALSE'] as RegExpMatchArray;
      const result = UNIQUE.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should handle unique by rows', () => {
      const matches = ['UNIQUE(A1:D4, TRUE)', 'A1:D4', 'TRUE'] as RegExpMatchArray;
      const result = UNIQUE.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should return exactly once option', () => {
      const matches = ['UNIQUE(A6:F6, FALSE, TRUE)', 'A6:F6', 'FALSE', 'TRUE'] as RegExpMatchArray;
      const result = UNIQUE.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should handle vertical data', () => {
      const matches = ['UNIQUE(A1:A8)', 'A1:A8'] as RegExpMatchArray;
      const result = UNIQUE.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should handle mixed data types', () => {
      const matches = ['UNIQUE(A1:C7)', 'A1:C7'] as RegExpMatchArray;
      const result = UNIQUE.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });
  });

  describe('JOIN Function', () => {
    it('should join with comma delimiter', () => {
      const matches = ['JOIN(",", A1:D1)', '","', 'A1:D1'] as RegExpMatchArray;
      const result = JOIN.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should join with space delimiter', () => {
      const matches = ['JOIN(" ", B2:D2)', '" "', 'B2:D2'] as RegExpMatchArray;
      const result = JOIN.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should join with custom delimiter', () => {
      const matches = ['JOIN(" - ", A7:C7)', '" - "', 'A7:C7'] as RegExpMatchArray;
      const result = JOIN.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should join vertical range', () => {
      const matches = ['JOIN("; ", A1:A4)', '"; "', 'A1:A4'] as RegExpMatchArray;
      const result = JOIN.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should handle empty cells', () => {
      const matches = ['JOIN(",", A1:F1)', '","', 'A1:F1'] as RegExpMatchArray;
      const result = JOIN.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should join multiple arguments', () => {
      const matches = ['JOIN(" ", "Hello", "World", "!")', '" "', '"Hello"', '"World"', '"!"'] as RegExpMatchArray;
      const result = JOIN.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });
  });

  describe('FLATTEN Function', () => {
    it('should flatten 2D range', () => {
      const matches = ['FLATTEN(A1:D4)', 'A1:D4'] as RegExpMatchArray;
      const result = FLATTEN.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should flatten multiple ranges', () => {
      const matches = ['FLATTEN(A1:B2, C3:D4)', 'A1:B2', 'C3:D4'] as RegExpMatchArray;
      const result = FLATTEN.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should flatten vertical range', () => {
      const matches = ['FLATTEN(A1:A8)', 'A1:A8'] as RegExpMatchArray;
      const result = FLATTEN.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should flatten horizontal range', () => {
      const matches = ['FLATTEN(A1:H1)', 'A1:H1'] as RegExpMatchArray;
      const result = FLATTEN.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should skip empty cells', () => {
      const matches = ['FLATTEN(A1:D10)', 'A1:D10'] as RegExpMatchArray;
      const result = FLATTEN.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });
  });

  describe('TRANSPOSE Function', () => {
    it('should transpose rectangular range', () => {
      const matches = ['TRANSPOSE(A1:D3)', 'A1:D3'] as RegExpMatchArray;
      const result = TRANSPOSE.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should transpose single row', () => {
      const matches = ['TRANSPOSE(A1:F1)', 'A1:F1'] as RegExpMatchArray;
      const result = TRANSPOSE.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should transpose single column', () => {
      const matches = ['TRANSPOSE(A1:A6)', 'A1:A6'] as RegExpMatchArray;
      const result = TRANSPOSE.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should transpose square range', () => {
      const matches = ['TRANSPOSE(A1:C3)', 'A1:C3'] as RegExpMatchArray;
      const result = TRANSPOSE.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should handle mixed data types', () => {
      const matches = ['TRANSPOSE(A1:C4)', 'A1:C4'] as RegExpMatchArray;
      const result = TRANSPOSE.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });
  });

  describe('Edge Cases and Integration', () => {
    it('should combine SPLIT and UNIQUE', () => {
      const matches = ['UNIQUE(SPLIT("a,b,a,c,b,d", ","))', 
        'SPLIT("a,b,a,c,b,d", ",")'] as RegExpMatchArray;
      const result = UNIQUE.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should combine JOIN and UNIQUE', () => {
      const matches = ['JOIN(", ", UNIQUE(A6:F6))', '", "', 'UNIQUE(A6:F6)'] as RegExpMatchArray;
      const result = JOIN.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should combine FLATTEN and TRANSPOSE', () => {
      const matches = ['FLATTEN(TRANSPOSE(A1:D3))', 'TRANSPOSE(A1:D3)'] as RegExpMatchArray;
      const result = FLATTEN.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should detect language of translated text', () => {
      const matches = ['DETECTLANGUAGE(GOOGLETRANSLATE("Hello", "ja"))', 
        'GOOGLETRANSLATE("Hello", "ja")'] as RegExpMatchArray;
      const result = DETECTLANGUAGE.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Google Sheets specific functions not supported');
    });

    it('should handle all functions returning consistent error', () => {
      const functions = [
        DETECTLANGUAGE, IMAGE, MAP, SPLIT, UNIQUE, JOIN, FLATTEN, TRANSPOSE
      ];
      
      for (const func of functions) {
        const matches = ['FUNC("param")', '"param"'] as RegExpMatchArray;
        const result = func.calculate(matches, mockContext);
        expect(result).toBe('#N/A - Google Sheets specific functions not supported');
      }
    });
  });
});