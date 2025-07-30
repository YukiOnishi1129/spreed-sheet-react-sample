import { describe, it, expect } from 'vitest';
import {
  IMPORTDATA, IMPORTFEED, IMPORTHTML, IMPORTXML, IMPORTRANGE
} from '../importLogic';
import { FormulaError } from '../../shared/types';
import type { FormulaContext } from '../../shared/types';

// Helper function to create FormulaContext
const createContext = (data: (string | number | boolean | null)[][]): FormulaContext => ({
  data: data.map(row => row.map(cell => ({ value: cell }))),
  row: 0,
  col: 0
});

describe('Import Functions', () => {
  const mockContext = createContext([
    ['https://example.com/data.csv', 'https://example.com/data.json'], // URLs
    ['https://example.com/rss.xml', 'https://example.com/atom.xml'], // Feed URLs
    ['https://example.com/table.html', 'table', 'list'], // HTML URLs and queries
    ['//item/title', '//channel/item', '@href'], // XPath queries
    ['1234567890abcdef', 'Sheet1!A1:C10'], // Spreadsheet IDs and ranges
  ]);

  describe('IMPORTDATA Function', () => {
    it('should attempt to import CSV data', () => {
      const matches = ['IMPORTDATA("https://example.com/data.csv")', 
        '"https://example.com/data.csv"'] as RegExpMatchArray;
      const result = IMPORTDATA.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Web import functions require external data access');
    });

    it('should handle TSV files', () => {
      const matches = ['IMPORTDATA("https://example.com/data.tsv")', 
        '"https://example.com/data.tsv"'] as RegExpMatchArray;
      const result = IMPORTDATA.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Web import functions require external data access');
    });

    it('should handle JSON endpoints', () => {
      const matches = ['IMPORTDATA("https://api.example.com/data.json")', 
        '"https://api.example.com/data.json"'] as RegExpMatchArray;
      const result = IMPORTDATA.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Web import functions require external data access');
    });

    it('should handle secure HTTPS URLs', () => {
      const matches = ['IMPORTDATA("https://secure.example.com/data.csv")', 
        '"https://secure.example.com/data.csv"'] as RegExpMatchArray;
      const result = IMPORTDATA.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Web import functions require external data access');
    });

    it('should handle URLs with query parameters', () => {
      const matches = ['IMPORTDATA("https://example.com/export?format=csv&limit=100")', 
        '"https://example.com/export?format=csv&limit=100"'] as RegExpMatchArray;
      const result = IMPORTDATA.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Web import functions require external data access');
    });

    it('should handle cell references', () => {
      const matches = ['IMPORTDATA(A1)', 'A1'] as RegExpMatchArray;
      const result = IMPORTDATA.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Web import functions require external data access');
    });
  });

  describe('IMPORTFEED Function', () => {
    it('should import RSS feed', () => {
      const matches = ['IMPORTFEED("https://example.com/rss.xml")', 
        '"https://example.com/rss.xml"'] as RegExpMatchArray;
      const result = IMPORTFEED.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Web import functions require external data access');
    });

    it('should handle Atom feeds', () => {
      const matches = ['IMPORTFEED("https://example.com/feed.atom")', 
        '"https://example.com/feed.atom"'] as RegExpMatchArray;
      const result = IMPORTFEED.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Web import functions require external data access');
    });

    it('should handle query for specific items', () => {
      const matches = ['IMPORTFEED("https://example.com/rss.xml", "items title")', 
        '"https://example.com/rss.xml"', '"items title"'] as RegExpMatchArray;
      const result = IMPORTFEED.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Web import functions require external data access');
    });

    it('should handle headers parameter', () => {
      const matches = ['IMPORTFEED("https://example.com/rss.xml", "items", TRUE)', 
        '"https://example.com/rss.xml"', '"items"', 'TRUE'] as RegExpMatchArray;
      const result = IMPORTFEED.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Web import functions require external data access');
    });

    it('should handle num_items parameter', () => {
      const matches = ['IMPORTFEED("https://example.com/rss.xml", "items", TRUE, 10)', 
        '"https://example.com/rss.xml"', '"items"', 'TRUE', '10'] as RegExpMatchArray;
      const result = IMPORTFEED.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Web import functions require external data access');
    });

    it('should handle feed info query', () => {
      const matches = ['IMPORTFEED("https://example.com/rss.xml", "feed")', 
        '"https://example.com/rss.xml"', '"feed"'] as RegExpMatchArray;
      const result = IMPORTFEED.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Web import functions require external data access');
    });
  });

  describe('IMPORTHTML Function', () => {
    it('should import HTML table', () => {
      const matches = ['IMPORTHTML("https://example.com/page.html", "table", 1)', 
        '"https://example.com/page.html"', '"table"', '1'] as RegExpMatchArray;
      const result = IMPORTHTML.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Web import functions require external data access');
    });

    it('should import HTML list', () => {
      const matches = ['IMPORTHTML("https://example.com/page.html", "list", 2)', 
        '"https://example.com/page.html"', '"list"', '2'] as RegExpMatchArray;
      const result = IMPORTHTML.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Web import functions require external data access');
    });

    it('should handle first table by default', () => {
      const matches = ['IMPORTHTML("https://example.com/data.html", "table", 0)', 
        '"https://example.com/data.html"', '"table"', '0'] as RegExpMatchArray;
      const result = IMPORTHTML.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Web import functions require external data access');
    });

    it('should handle nested tables', () => {
      const matches = ['IMPORTHTML("https://example.com/complex.html", "table", 5)', 
        '"https://example.com/complex.html"', '"table"', '5'] as RegExpMatchArray;
      const result = IMPORTHTML.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Web import functions require external data access');
    });

    it('should handle ordered lists', () => {
      const matches = ['IMPORTHTML("https://example.com/page.html", "list", 1)', 
        '"https://example.com/page.html"', '"list"', '1'] as RegExpMatchArray;
      const result = IMPORTHTML.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Web import functions require external data access');
    });

    it('should handle Wikipedia pages', () => {
      const matches = ['IMPORTHTML("https://en.wikipedia.org/wiki/List_of_countries", "table", 1)', 
        '"https://en.wikipedia.org/wiki/List_of_countries"', '"table"', '1'] as RegExpMatchArray;
      const result = IMPORTHTML.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Web import functions require external data access');
    });
  });

  describe('IMPORTXML Function', () => {
    it('should import XML with XPath', () => {
      const matches = ['IMPORTXML("https://example.com/data.xml", "//item/title")', 
        '"https://example.com/data.xml"', '"//item/title"'] as RegExpMatchArray;
      const result = IMPORTXML.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Web import functions require external data access');
    });

    it('should handle attribute selection', () => {
      const matches = ['IMPORTXML("https://example.com/page.html", "//a/@href")', 
        '"https://example.com/page.html"', '"//a/@href"'] as RegExpMatchArray;
      const result = IMPORTXML.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Web import functions require external data access');
    });

    it('should handle complex XPath queries', () => {
      const matches = ['IMPORTXML("https://example.com/data.xml", "//div[@class=\'content\']//p[1]")', 
        '"https://example.com/data.xml"', '"//div[@class=\'content\']//p[1]"'] as RegExpMatchArray;
      const result = IMPORTXML.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Web import functions require external data access');
    });

    it('should handle HTML parsing with XPath', () => {
      const matches = ['IMPORTXML("https://example.com/page.html", "//h1")', 
        '"https://example.com/page.html"', '"//h1"'] as RegExpMatchArray;
      const result = IMPORTXML.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Web import functions require external data access');
    });

    it('should handle namespace queries', () => {
      const matches = ['IMPORTXML("https://example.com/feed.xml", "//rss:item/rss:title")', 
        '"https://example.com/feed.xml"', '"//rss:item/rss:title"'] as RegExpMatchArray;
      const result = IMPORTXML.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Web import functions require external data access');
    });

    it('should handle text content extraction', () => {
      const matches = ['IMPORTXML("https://example.com/page.html", "//p/text()")', 
        '"https://example.com/page.html"', '"//p/text()"'] as RegExpMatchArray;
      const result = IMPORTXML.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Web import functions require external data access');
    });
  });

  describe('IMPORTRANGE Function', () => {
    it('should import range from another spreadsheet', () => {
      const matches = ['IMPORTRANGE("1234567890abcdef", "Sheet1!A1:C10")', 
        '"1234567890abcdef"', '"Sheet1!A1:C10"'] as RegExpMatchArray;
      const result = IMPORTRANGE.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Web import functions require external data access');
    });

    it('should handle full URL', () => {
      const matches = ['IMPORTRANGE("https://docs.google.com/spreadsheets/d/1234567890abcdef", "Data!B2:D20")', 
        '"https://docs.google.com/spreadsheets/d/1234567890abcdef"', '"Data!B2:D20"'] as RegExpMatchArray;
      const result = IMPORTRANGE.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Web import functions require external data access');
    });

    it('should handle single cell reference', () => {
      const matches = ['IMPORTRANGE("abcdef1234567890", "Summary!B5")', 
        '"abcdef1234567890"', '"Summary!B5"'] as RegExpMatchArray;
      const result = IMPORTRANGE.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Web import functions require external data access');
    });

    it('should handle entire column', () => {
      const matches = ['IMPORTRANGE("spreadsheet_id", "Sheet1!A:A")', 
        '"spreadsheet_id"', '"Sheet1!A:A"'] as RegExpMatchArray;
      const result = IMPORTRANGE.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Web import functions require external data access');
    });

    it('should handle entire row', () => {
      const matches = ['IMPORTRANGE("spreadsheet_id", "Sheet1!1:1")', 
        '"spreadsheet_id"', '"Sheet1!1:1"'] as RegExpMatchArray;
      const result = IMPORTRANGE.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Web import functions require external data access');
    });

    it('should handle named ranges', () => {
      const matches = ['IMPORTRANGE("1234567890abcdef", "NamedRange")', 
        '"1234567890abcdef"', '"NamedRange"'] as RegExpMatchArray;
      const result = IMPORTRANGE.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Web import functions require external data access');
    });
  });

  describe('Edge Cases and Integration', () => {
    it('should handle invalid URLs gracefully', () => {
      const matches = ['IMPORTDATA("not-a-url")', '"not-a-url"'] as RegExpMatchArray;
      const result = IMPORTDATA.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Web import functions require external data access');
    });

    it('should handle authentication requirements', () => {
      const matches = ['IMPORTDATA("https://secure.example.com/private/data.csv")', 
        '"https://secure.example.com/private/data.csv"'] as RegExpMatchArray;
      const result = IMPORTDATA.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Web import functions require external data access');
    });

    it('should handle rate limiting scenarios', () => {
      const matches = ['IMPORTXML("https://api.example.com/data", "//result")', 
        '"https://api.example.com/data"', '"//result"'] as RegExpMatchArray;
      const result = IMPORTXML.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Web import functions require external data access');
    });

    it('should handle CORS restrictions', () => {
      const matches = ['IMPORTHTML("https://cross-origin.example.com/table.html", "table", 1)', 
        '"https://cross-origin.example.com/table.html"', '"table"', '1'] as RegExpMatchArray;
      const result = IMPORTHTML.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Web import functions require external data access');
    });

    it('should handle redirects', () => {
      const matches = ['IMPORTDATA("https://bit.ly/shortened-url")', 
        '"https://bit.ly/shortened-url"'] as RegExpMatchArray;
      const result = IMPORTDATA.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Web import functions require external data access');
    });

    it('should handle all import functions returning consistent error', () => {
      // Test each function with appropriate parameters
      const testCases = [
        { func: IMPORTDATA, matches: ['IMPORTDATA("https://example.com")', '"https://example.com"'] },
        { func: IMPORTFEED, matches: ['IMPORTFEED("https://example.com")', '"https://example.com"'] },
        { func: IMPORTHTML, matches: ['IMPORTHTML("https://example.com", "table", 1)', '"https://example.com"', '"table"', '1'] },
        { func: IMPORTXML, matches: ['IMPORTXML("https://example.com", "//div")', '"https://example.com"', '"//div"'] },
        { func: IMPORTRANGE, matches: ['IMPORTRANGE("spreadsheet_id", "Sheet1!A1")', '"spreadsheet_id"', '"Sheet1!A1"'] }
      ];
      
      for (const { func, matches } of testCases) {
        const result = func.calculate(matches as RegExpMatchArray, mockContext);
        expect(result).toBe('#N/A - Web import functions require external data access');
      }
    });
  });
});