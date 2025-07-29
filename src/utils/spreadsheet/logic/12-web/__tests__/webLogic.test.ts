import { describe, it, expect } from 'vitest';
import {
  ENCODEURL, HYPERLINK, ISURL, WEBSERVICE
} from '../webLogic';
import { FormulaError } from '../../shared/types';
import type { FormulaContext } from '../../shared/types';

// Helper function to create FormulaContext
const createContext = (data: (string | number | boolean | null)[][]): FormulaContext => ({
  data: data.map(row => row.map(cell => ({ value: cell }))),
  row: 0,
  col: 0
});

describe('Web Functions', () => {
  const mockContext = createContext([
    ['https://example.com', 'http://test.org', 'ftp://files.com'], // URLs
    ['Hello World', 'Test & Example', 'Path/To/File'], // text to encode
    ['Click Here', 'Visit Site', 'Download'], // link text
    ['not-a-url', 'example.com', 'https://'], // invalid/partial URLs
    ['https://api.example.com/data', 'GET', 'POST'], // API endpoints
  ]);

  describe('ENCODEURL Function', () => {
    it('should encode URL with spaces', () => {
      const matches = ['ENCODEURL("Hello World")', '"Hello World"'] as RegExpMatchArray;
      const result = ENCODEURL.calculate(matches, mockContext);
      expect(result).toBe('Hello%20World');
    });

    it('should encode special characters', () => {
      const matches = ['ENCODEURL("Test & Example")', '"Test & Example"'] as RegExpMatchArray;
      const result = ENCODEURL.calculate(matches, mockContext);
      expect(result).toBe('Test%20%26%20Example');
    });

    it('should encode path separators', () => {
      const matches = ['ENCODEURL("Path/To/File")', '"Path/To/File"'] as RegExpMatchArray;
      const result = ENCODEURL.calculate(matches, mockContext);
      expect(result).toBe('Path%2FTo%2FFile');
    });

    it('should handle empty string', () => {
      const matches = ['ENCODEURL("")', '""'] as RegExpMatchArray;
      const result = ENCODEURL.calculate(matches, mockContext);
      expect(result).toBe('');
    });

    it('should encode query parameters', () => {
      const matches = ['ENCODEURL("name=John Doe&age=30")', '"name=John Doe&age=30"'] as RegExpMatchArray;
      const result = ENCODEURL.calculate(matches, mockContext);
      expect(result).toBe('name%3DJohn%20Doe%26age%3D30');
    });

    it('should encode international characters', () => {
      const matches = ['ENCODEURL("こんにちは")', '"こんにちは"'] as RegExpMatchArray;
      const result = ENCODEURL.calculate(matches, mockContext);
      expect(result).toBe('%E3%81%93%E3%82%93%E3%81%AB%E3%81%A1%E3%81%AF');
    });

    it('should handle already encoded URLs', () => {
      const matches = ['ENCODEURL("Hello%20World")', '"Hello%20World"'] as RegExpMatchArray;
      const result = ENCODEURL.calculate(matches, mockContext);
      expect(result).toBe('Hello%2520World'); // Double encoding
    });

    it('should return VALUE error for non-string input', () => {
      const matches = ['ENCODEURL(123)', '123'] as RegExpMatchArray;
      const result = ENCODEURL.calculate(matches, mockContext);
      expect(result).toBe('123'); // Numbers are converted to string
    });

    it('should handle cell references', () => {
      const matches = ['ENCODEURL(B2)', 'B2'] as RegExpMatchArray;
      const result = ENCODEURL.calculate(matches, mockContext);
      expect(result).toBe('Test%20%26%20Example');
    });
  });

  describe('HYPERLINK Function', () => {
    it('should create hyperlink with custom text', () => {
      const matches = ['HYPERLINK("https://example.com", "Click Here")', 
        '"https://example.com"', '"Click Here"'] as RegExpMatchArray;
      const result = HYPERLINK.calculate(matches, mockContext);
      expect(result).toBe('Click Here');
    });

    it('should use URL as text when link_label omitted', () => {
      const matches = ['HYPERLINK("https://example.com")', '"https://example.com"'] as RegExpMatchArray;
      const result = HYPERLINK.calculate(matches, mockContext);
      expect(result).toBe('https://example.com');
    });

    it('should handle email links', () => {
      const matches = ['HYPERLINK("mailto:test@example.com", "Email Us")', 
        '"mailto:test@example.com"', '"Email Us"'] as RegExpMatchArray;
      const result = HYPERLINK.calculate(matches, mockContext);
      expect(result).toBe('Email Us');
    });

    it('should handle file links', () => {
      const matches = ['HYPERLINK("file:///C:/Documents/file.pdf", "Open File")', 
        '"file:///C:/Documents/file.pdf"', '"Open File"'] as RegExpMatchArray;
      const result = HYPERLINK.calculate(matches, mockContext);
      expect(result).toBe('Open File');
    });

    it('should handle anchor links', () => {
      const matches = ['HYPERLINK("#Sheet2!A1", "Go to Sheet2")', 
        '"#Sheet2!A1"', '"Go to Sheet2"'] as RegExpMatchArray;
      const result = HYPERLINK.calculate(matches, mockContext);
      expect(result).toBe('Go to Sheet2');
    });

    it('should handle FTP links', () => {
      const matches = ['HYPERLINK("ftp://files.example.com/data.zip", "Download")', 
        '"ftp://files.example.com/data.zip"', '"Download"'] as RegExpMatchArray;
      const result = HYPERLINK.calculate(matches, mockContext);
      expect(result).toBe('Download');
    });

    it('should handle URLs with query parameters', () => {
      const matches = ['HYPERLINK("https://example.com?search=test&page=1", "Search Results")', 
        '"https://example.com?search=test&page=1"', '"Search Results"'] as RegExpMatchArray;
      const result = HYPERLINK.calculate(matches, mockContext);
      expect(result).toBe('Search Results');
    });

    it('should handle empty link label', () => {
      const matches = ['HYPERLINK("https://example.com", "")', 
        '"https://example.com"', '""'] as RegExpMatchArray;
      const result = HYPERLINK.calculate(matches, mockContext);
      expect(result).toBe('');
    });

    it('should handle cell references', () => {
      const matches = ['HYPERLINK(A1, C1)', 'A1', 'C1'] as RegExpMatchArray;
      const result = HYPERLINK.calculate(matches, mockContext);
      expect(result).toBe('Click Here');
    });
  });

  describe('ISURL Function', () => {
    it('should return TRUE for valid HTTP URL', () => {
      const matches = ['ISURL("https://example.com")', '"https://example.com"'] as RegExpMatchArray;
      const result = ISURL.calculate(matches, mockContext);
      expect(result).toBe(true);
    });

    it('should return TRUE for HTTP URL', () => {
      const matches = ['ISURL("http://test.org")', '"http://test.org"'] as RegExpMatchArray;
      const result = ISURL.calculate(matches, mockContext);
      expect(result).toBe(true);
    });

    it('should return TRUE for FTP URL', () => {
      const matches = ['ISURL("ftp://files.com")', '"ftp://files.com"'] as RegExpMatchArray;
      const result = ISURL.calculate(matches, mockContext);
      expect(result).toBe(true);
    });

    it('should return FALSE for missing protocol', () => {
      const matches = ['ISURL("example.com")', '"example.com"'] as RegExpMatchArray;
      const result = ISURL.calculate(matches, mockContext);
      expect(result).toBe(false);
    });

    it('should return FALSE for invalid URL', () => {
      const matches = ['ISURL("not-a-url")', '"not-a-url"'] as RegExpMatchArray;
      const result = ISURL.calculate(matches, mockContext);
      expect(result).toBe(false);
    });

    it('should return FALSE for incomplete URL', () => {
      const matches = ['ISURL("https://")', '"https://"'] as RegExpMatchArray;
      const result = ISURL.calculate(matches, mockContext);
      expect(result).toBe(false);
    });

    it('should handle URLs with paths', () => {
      const matches = ['ISURL("https://example.com/path/to/page")', 
        '"https://example.com/path/to/page"'] as RegExpMatchArray;
      const result = ISURL.calculate(matches, mockContext);
      expect(result).toBe(true);
    });

    it('should handle URLs with query parameters', () => {
      const matches = ['ISURL("https://example.com?param=value")', 
        '"https://example.com?param=value"'] as RegExpMatchArray;
      const result = ISURL.calculate(matches, mockContext);
      expect(result).toBe(true);
    });

    it('should handle URLs with fragments', () => {
      const matches = ['ISURL("https://example.com#section")', 
        '"https://example.com#section"'] as RegExpMatchArray;
      const result = ISURL.calculate(matches, mockContext);
      expect(result).toBe(true);
    });

    it('should handle URLs with port numbers', () => {
      const matches = ['ISURL("https://example.com:8080")', 
        '"https://example.com:8080"'] as RegExpMatchArray;
      const result = ISURL.calculate(matches, mockContext);
      expect(result).toBe(true);
    });

    it('should return FALSE for empty string', () => {
      const matches = ['ISURL("")', '""'] as RegExpMatchArray;
      const result = ISURL.calculate(matches, mockContext);
      expect(result).toBe(false);
    });

    it('should return FALSE for non-string values', () => {
      const matches = ['ISURL(123)', '123'] as RegExpMatchArray;
      const result = ISURL.calculate(matches, mockContext);
      expect(result).toBe(false);
    });

    it('should handle cell references', () => {
      const matches = ['ISURL(A1)', 'A1'] as RegExpMatchArray;
      const result = ISURL.calculate(matches, mockContext);
      expect(result).toBe(true); // https://example.com
    });
  });

  describe('WEBSERVICE Function', () => {
    it('should attempt to call web service', () => {
      const matches = ['WEBSERVICE("https://api.example.com/data")', 
        '"https://api.example.com/data"'] as RegExpMatchArray;
      const result = WEBSERVICE.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Web service calls require external access');
    });

    it('should handle API endpoints', () => {
      const matches = ['WEBSERVICE("https://api.example.com/v1/users")', 
        '"https://api.example.com/v1/users"'] as RegExpMatchArray;
      const result = WEBSERVICE.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Web service calls require external access');
    });

    it('should handle REST APIs', () => {
      const matches = ['WEBSERVICE("https://jsonplaceholder.typicode.com/posts/1")', 
        '"https://jsonplaceholder.typicode.com/posts/1"'] as RegExpMatchArray;
      const result = WEBSERVICE.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Web service calls require external access');
    });

    it('should handle GraphQL endpoints', () => {
      const matches = ['WEBSERVICE("https://api.example.com/graphql")', 
        '"https://api.example.com/graphql"'] as RegExpMatchArray;
      const result = WEBSERVICE.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Web service calls require external access');
    });

    it('should handle authenticated endpoints', () => {
      const matches = ['WEBSERVICE("https://api.example.com/private/data")', 
        '"https://api.example.com/private/data"'] as RegExpMatchArray;
      const result = WEBSERVICE.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Web service calls require external access');
    });

    it('should handle query parameters', () => {
      const matches = ['WEBSERVICE("https://api.example.com/search?q=test&limit=10")', 
        '"https://api.example.com/search?q=test&limit=10"'] as RegExpMatchArray;
      const result = WEBSERVICE.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Web service calls require external access');
    });

    it('should handle XML responses', () => {
      const matches = ['WEBSERVICE("https://api.example.com/data.xml")', 
        '"https://api.example.com/data.xml"'] as RegExpMatchArray;
      const result = WEBSERVICE.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Web service calls require external access');
    });

    it('should handle JSON responses', () => {
      const matches = ['WEBSERVICE("https://api.example.com/data.json")', 
        '"https://api.example.com/data.json"'] as RegExpMatchArray;
      const result = WEBSERVICE.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Web service calls require external access');
    });

    it('should handle cell references', () => {
      const matches = ['WEBSERVICE(A5)', 'A5'] as RegExpMatchArray;
      const result = WEBSERVICE.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Web service calls require external access');
    });
  });

  describe('Edge Cases and Integration', () => {
    it('should encode URLs for use in HYPERLINK', () => {
      const text = 'Search Term';
      const encodedMatches = ['ENCODEURL("' + text + '")', '"' + text + '"'] as RegExpMatchArray;
      const encoded = ENCODEURL.calculate(encodedMatches, mockContext);
      
      const hyperlinkMatches = ['HYPERLINK("https://example.com/search?q=' + encoded + '", "Search")', 
        '"https://example.com/search?q=' + encoded + '"', '"Search"'] as RegExpMatchArray;
      const result = HYPERLINK.calculate(hyperlinkMatches, mockContext);
      
      expect(result).toBe('Search');
    });

    it('should validate URLs before using in WEBSERVICE', () => {
      const url = 'https://api.example.com/data';
      
      const isurlMatches = ['ISURL("' + url + '")', '"' + url + '"'] as RegExpMatchArray;
      const isValid = ISURL.calculate(isurlMatches, mockContext);
      
      expect(isValid).toBe(true);
      
      if (isValid) {
        const webserviceMatches = ['WEBSERVICE("' + url + '")', '"' + url + '"'] as RegExpMatchArray;
        const result = WEBSERVICE.calculate(webserviceMatches, mockContext);
        expect(result).toBe('#N/A - Web service calls require external access');
      }
    });

    it('should handle internationalized domain names', () => {
      const matches = ['ISURL("https://例え.jp")', '"https://例え.jp"'] as RegExpMatchArray;
      const result = ISURL.calculate(matches, mockContext);
      expect(result).toBe(true);
    });

    it('should handle data URIs', () => {
      const matches = ['ISURL("data:text/plain;base64,SGVsbG8gV29ybGQ=")', 
        '"data:text/plain;base64,SGVsbG8gV29ybGQ="'] as RegExpMatchArray;
      const result = ISURL.calculate(matches, mockContext);
      expect(result).toBe(true);
    });

    it('should handle special protocols', () => {
      const protocols = ['tel:+1234567890', 'sms:+1234567890', 'geo:37.786,-122.406'];
      
      for (const protocol of protocols) {
        const matches = ['HYPERLINK("' + protocol + '", "Link")', 
          '"' + protocol + '"', '"Link"'] as RegExpMatchArray;
        const result = HYPERLINK.calculate(matches, mockContext);
        expect(result).toBe('Link');
      }
    });

    it('should handle extremely long URLs', () => {
      const longPath = 'a'.repeat(2000);
      const matches = ['ISURL("https://example.com/' + longPath + '")', 
        '"https://example.com/' + longPath + '"'] as RegExpMatchArray;
      const result = ISURL.calculate(matches, mockContext);
      expect(result).toBe(true);
    });
  });
});