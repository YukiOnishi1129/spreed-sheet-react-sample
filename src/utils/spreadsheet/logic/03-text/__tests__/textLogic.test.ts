import { describe, it, expect } from 'vitest';
import type { FormulaContext } from '../../shared/types';
import {
  CONCATENATE, CONCAT, LEFT, RIGHT, MID,
  LEN, UPPER, LOWER, PROPER,
  TRIM, SUBSTITUTE, REPLACE, FIND, SEARCH,
  TEXT, VALUE, CHAR, CODE,
  REPT, EXACT, TEXTJOIN, CLEAN, UNICHAR, UNICODE
} from '../textLogic';
import { FormulaError } from '../../shared/types';

// Helper function to create FormulaContext
function createContext(data: any[][]): FormulaContext {
  return {
    data,
    row: 0,
    col: 0
  };
}

describe('Text Functions', () => {
  const mockContext = createContext([
    ['Hello', 123, 'Hello World', 'text'],                  // Row 1 (A1-D1)
    ['World', 45.67, 'apple,banana,orange', ' '],           // Row 2 (A2-D2)
    ['  Test  ', true, 'Test123Test', ','],                 // Row 3 (A3-D3)
    ['UPPER CASE', false, null, null],                      // Row 4 (A4-D4)
    ['lower case', new Date('2024-01-15'), '', null]        // Row 5 (A5-D5)
  ]);

  describe('Concatenation Functions', () => {
    describe('CONCATENATE', () => {
      it('should concatenate multiple strings', () => {
        const matches = ['CONCATENATE(A1, " ", A2)', 'A1, " ", A2'] as RegExpMatchArray;
        const result = CONCATENATE.calculate(matches, mockContext);
        expect(result).toBe('Hello World');
      });

      it('should concatenate cell values with literals', () => {
        const matches = ['CONCATENATE("The answer is ", B1)', '"The answer is ", B1'] as RegExpMatchArray;
        const result = CONCATENATE.calculate(matches, mockContext);
        expect(result).toBe('The answer is 123');
      });

      it('should handle nested functions', () => {
        const matches = ['CONCATENATE(LEFT(A1, 3), RIGHT(A2, 3))', 'LEFT(A1, 3), RIGHT(A2, 3)'] as RegExpMatchArray;
        const result = CONCATENATE.calculate(matches, mockContext);
        expect(result).toBe('Helrld');
      });

      it('should handle empty values', () => {
        const matches = ['CONCATENATE(A1, C4, A2)', 'A1, C4, A2'] as RegExpMatchArray;
        const result = CONCATENATE.calculate(matches, mockContext);
        expect(result).toBe('HelloWorld');
      });

      it('should convert numbers to strings', () => {
        const matches = ['CONCATENATE(B1, B2)', 'B1, B2'] as RegExpMatchArray;
        const result = CONCATENATE.calculate(matches, mockContext);
        expect(result).toBe('12345.67');
      });

      it('should convert booleans to strings', () => {
        const matches = ['CONCATENATE(B3, " ", B4)', 'B3, " ", B4'] as RegExpMatchArray;
        const result = CONCATENATE.calculate(matches, mockContext);
        expect(result).toBe('true false');
      });
    });

    describe('CONCAT', () => {
      it('should concatenate range of cells', () => {
        const matches = ['CONCAT(A1:A2)', 'A1:A2'] as RegExpMatchArray;
        const result = CONCAT.calculate(matches, mockContext);
        expect(result).toBe('HelloWorld');
      });

      it('should concatenate multiple ranges and values', () => {
        const matches = ['CONCAT(A1:A2, " ", B1)', 'A1:A2, " ", B1'] as RegExpMatchArray;
        const result = CONCAT.calculate(matches, mockContext);
        expect(result).toBe('HelloWorld 123');
      });

      it('should handle empty cells in range', () => {
        const matches = ['CONCAT(C3:C5)', 'C3:C5'] as RegExpMatchArray;
        const result = CONCAT.calculate(matches, mockContext);
        expect(result).toBe('Test123Test');
      });
    });
  });

  describe('Substring Functions', () => {
    describe('LEFT', () => {
      it('should extract leftmost characters', () => {
        const matches = ['LEFT(A1, 3)', 'A1', '3'] as RegExpMatchArray;
        const result = LEFT.calculate(matches, mockContext);
        expect(result).toBe('Hel');
      });

      it('should default to 1 character', () => {
        const matches = ['LEFT(A1)', 'A1'] as RegExpMatchArray;
        const result = LEFT.calculate(matches, mockContext);
        expect(result).toBe('H');
      });

      it('should handle num_chars greater than text length', () => {
        const matches = ['LEFT(A1, 10)', 'A1', '10'] as RegExpMatchArray;
        const result = LEFT.calculate(matches, mockContext);
        expect(result).toBe('Hello');
      });

      it('should return VALUE error for negative num_chars', () => {
        const matches = ['LEFT(A1, -1)', 'A1', '-1'] as RegExpMatchArray;
        const result = LEFT.calculate(matches, mockContext);
        expect(result).toBe(FormulaError.VALUE);
      });
    });

    describe('RIGHT', () => {
      it('should extract rightmost characters', () => {
        const matches = ['RIGHT(A2, 3)', 'A2', '3'] as RegExpMatchArray;
        const result = RIGHT.calculate(matches, mockContext);
        expect(result).toBe('rld');
      });

      it('should default to 1 character', () => {
        const matches = ['RIGHT(A2)', 'A2'] as RegExpMatchArray;
        const result = RIGHT.calculate(matches, mockContext);
        expect(result).toBe('d');
      });

      it('should handle num_chars greater than text length', () => {
        const matches = ['RIGHT(A2, 10)', 'A2', '10'] as RegExpMatchArray;
        const result = RIGHT.calculate(matches, mockContext);
        expect(result).toBe('World');
      });
    });

    describe('MID', () => {
      it('should extract middle substring', () => {
        const matches = ['MID(C1, 7, 5)', 'C1', '7', '5'] as RegExpMatchArray;
        const result = MID.calculate(matches, mockContext);
        expect(result).toBe('World');
      });

      it('should handle start position beyond text length', () => {
        const matches = ['MID(A1, 10, 5)', 'A1', '10', '5'] as RegExpMatchArray;
        const result = MID.calculate(matches, mockContext);
        expect(result).toBe('');
      });

      it('should return VALUE error for invalid start', () => {
        const matches = ['MID(A1, 0, 5)', 'A1', '0', '5'] as RegExpMatchArray;
        const result = MID.calculate(matches, mockContext);
        expect(result).toBe(FormulaError.VALUE);
      });

      it('should return VALUE error for negative num_chars', () => {
        const matches = ['MID(A1, 1, -5)', 'A1', '1', '-5'] as RegExpMatchArray;
        const result = MID.calculate(matches, mockContext);
        expect(result).toBe(FormulaError.VALUE);
      });
    });
  });

  describe('Text Information Functions', () => {
    describe('LEN', () => {
      it('should return length of text', () => {
        const matches = ['LEN(A1)', 'A1'] as RegExpMatchArray;
        const result = LEN.calculate(matches, mockContext);
        expect(result).toBe(5);
      });

      it('should return length including spaces', () => {
        const matches = ['LEN(A3)', 'A3'] as RegExpMatchArray;
        const result = LEN.calculate(matches, mockContext);
        expect(result).toBe(9); // "  Test  "
      });

      it('should return 0 for empty string', () => {
        const matches = ['LEN(C5)', 'C5'] as RegExpMatchArray;
        const result = LEN.calculate(matches, mockContext);
        expect(result).toBe(0);
      });

      it('should convert numbers to string and count', () => {
        const matches = ['LEN(B1)', 'B1'] as RegExpMatchArray;
        const result = LEN.calculate(matches, mockContext);
        expect(result).toBe(3); // "123"
      });
    });
  });

  describe('Case Conversion Functions', () => {
    describe('UPPER', () => {
      it('should convert to uppercase', () => {
        const matches = ['UPPER(A1)', 'A1'] as RegExpMatchArray;
        const result = UPPER.calculate(matches, mockContext);
        expect(result).toBe('HELLO');
      });

      it('should handle already uppercase text', () => {
        const matches = ['UPPER(A4)', 'A4'] as RegExpMatchArray;
        const result = UPPER.calculate(matches, mockContext);
        expect(result).toBe('UPPER CASE');
      });
    });

    describe('LOWER', () => {
      it('should convert to lowercase', () => {
        const matches = ['LOWER(A4)', 'A4'] as RegExpMatchArray;
        const result = LOWER.calculate(matches, mockContext);
        expect(result).toBe('upper case');
      });

      it('should handle already lowercase text', () => {
        const matches = ['LOWER(A5)', 'A5'] as RegExpMatchArray;
        const result = LOWER.calculate(matches, mockContext);
        expect(result).toBe('lower case');
      });
    });

    describe('PROPER', () => {
      it('should capitalize first letter of each word', () => {
        const matches = ['PROPER(A5)', 'A5'] as RegExpMatchArray;
        const result = PROPER.calculate(matches, mockContext);
        expect(result).toBe('Lower Case');
      });

      it('should handle mixed case text', () => {
        const matches = ['PROPER("hello WORLD")', '"hello WORLD"'] as RegExpMatchArray;
        const result = PROPER.calculate(matches, mockContext);
        expect(result).toBe('Hello World');
      });
    });
  });

  describe('Text Modification Functions', () => {
    describe('TRIM', () => {
      it('should remove leading and trailing spaces', () => {
        const matches = ['TRIM(A3)', 'A3'] as RegExpMatchArray;
        const result = TRIM.calculate(matches, mockContext);
        expect(result).toBe('Test');
      });

      it('should reduce multiple spaces to single space', () => {
        const matches = ['TRIM("  Hello    World  ")', '"  Hello    World  "'] as RegExpMatchArray;
        const result = TRIM.calculate(matches, mockContext);
        expect(result).toBe('Hello World');
      });
    });

    describe('SUBSTITUTE', () => {
      it('should substitute all occurrences', () => {
        const matches = ['SUBSTITUTE(C3, "Test", "Demo")', 'C3', '"Test"', '"Demo"'] as RegExpMatchArray;
        const result = SUBSTITUTE.calculate(matches, mockContext);
        expect(result).toBe('Demo123Demo');
      });

      it('should substitute specific occurrence', () => {
        const matches = ['SUBSTITUTE(C3, "Test", "Demo", 2)', 'C3', '"Test"', '"Demo"', '2'] as RegExpMatchArray;
        const result = SUBSTITUTE.calculate(matches, mockContext);
        expect(result).toBe('Test123Demo');
      });

      it('should handle case sensitivity', () => {
        const matches = ['SUBSTITUTE(C3, "test", "Demo")', 'C3', '"test"', '"Demo"'] as RegExpMatchArray;
        const result = SUBSTITUTE.calculate(matches, mockContext);
        expect(result).toBe('Test123Test');
      });
    });

    describe('REPLACE', () => {
      it('should replace text at specific position', () => {
        const matches = ['REPLACE(C1, 7, 5, "Earth")', 'C1', '7', '5', '"Earth"'] as RegExpMatchArray;
        const result = REPLACE.calculate(matches, mockContext);
        expect(result).toBe('Hello Earth');
      });

      it('should insert text when num_chars is 0', () => {
        const matches = ['REPLACE(A1, 3, 0, "XXX")', 'A1', '3', '0', '"XXX"'] as RegExpMatchArray;
        const result = REPLACE.calculate(matches, mockContext);
        expect(result).toBe('HeXXXllo');
      });

      it('should return VALUE error for invalid start', () => {
        const matches = ['REPLACE(A1, 0, 2, "XX")', 'A1', '0', '2', '"XX"'] as RegExpMatchArray;
        const result = REPLACE.calculate(matches, mockContext);
        expect(result).toBe(FormulaError.VALUE);
      });
    });
  });

  describe('Search Functions', () => {
    describe('FIND', () => {
      it('should find case-sensitive position', () => {
        const matches = ['FIND("World", C1)', '"World"', 'C1'] as RegExpMatchArray;
        const result = FIND.calculate(matches, mockContext);
        expect(result).toBe(7);
      });

      it('should find from start position', () => {
        const matches = ['FIND("o", C1, 5)', '"o"', 'C1', '5'] as RegExpMatchArray;
        const result = FIND.calculate(matches, mockContext);
        expect(result).toBe(8);
      });

      it('should return VALUE error when not found', () => {
        const matches = ['FIND("xyz", C1)', '"xyz"', 'C1'] as RegExpMatchArray;
        const result = FIND.calculate(matches, mockContext);
        expect(result).toBe(FormulaError.VALUE);
      });

      it('should be case sensitive', () => {
        const matches = ['FIND("world", C1)', '"world"', 'C1'] as RegExpMatchArray;
        const result = FIND.calculate(matches, mockContext);
        expect(result).toBe(FormulaError.VALUE);
      });
    });

    describe('SEARCH', () => {
      it('should find case-insensitive position', () => {
        const matches = ['SEARCH("world", C1)', '"world"', 'C1'] as RegExpMatchArray;
        const result = SEARCH.calculate(matches, mockContext);
        expect(result).toBe(7);
      });

      it('should support wildcards', () => {
        const matches = ['SEARCH("W?rld", C1)', '"W?rld"', 'C1'] as RegExpMatchArray;
        const result = SEARCH.calculate(matches, mockContext);
        expect(result).toBe(7);
      });

      it('should support asterisk wildcard', () => {
        const matches = ['SEARCH("H*d", C1)', '"H*d"', 'C1'] as RegExpMatchArray;
        const result = SEARCH.calculate(matches, mockContext);
        expect(result).toBe(1);
      });
    });
  });

  describe('Conversion Functions', () => {
    describe('TEXT', () => {
      it('should format number with decimal places', () => {
        const matches = ['TEXT(B2, "0.0")', 'B2', '"0.0"'] as RegExpMatchArray;
        const result = TEXT.calculate(matches, mockContext);
        expect(result).toBe('45.7');
      });

      it('should format number with padding', () => {
        const matches = ['TEXT(B1, "000000")', 'B1', '"000000"'] as RegExpMatchArray;
        const result = TEXT.calculate(matches, mockContext);
        expect(result).toBe('000123');
      });

      it('should format with percentage', () => {
        const matches = ['TEXT(0.5, "0%")', '0.5', '"0%"'] as RegExpMatchArray;
        const result = TEXT.calculate(matches, mockContext);
        expect(result).toBe('50%');
      });
    });

    describe('VALUE', () => {
      it('should convert text to number', () => {
        const matches = ['VALUE("123.45")', '"123.45"'] as RegExpMatchArray;
        const result = VALUE.calculate(matches, mockContext);
        expect(result).toBe(123.45);
      });

      it('should handle percentage text', () => {
        const matches = ['VALUE("50%")', '"50%"'] as RegExpMatchArray;
        const result = VALUE.calculate(matches, mockContext);
        expect(result).toBe(0.5);
      });

      it('should return VALUE error for non-numeric text', () => {
        const matches = ['VALUE(A1)', 'A1'] as RegExpMatchArray;
        const result = VALUE.calculate(matches, mockContext);
        expect(result).toBe(FormulaError.VALUE);
      });
    });
  });

  describe('Character Functions', () => {
    describe('CHAR', () => {
      it('should return character from ASCII code', () => {
        const matches = ['CHAR(65)', '65'] as RegExpMatchArray;
        const result = CHAR.calculate(matches, mockContext);
        expect(result).toBe('A');
      });

      it('should return VALUE error for invalid code', () => {
        const matches = ['CHAR(0)', '0'] as RegExpMatchArray;
        const result = CHAR.calculate(matches, mockContext);
        expect(result).toBe(FormulaError.VALUE);
      });
    });

    describe('CODE', () => {
      it('should return ASCII code of first character', () => {
        const matches = ['CODE("A")', '"A"'] as RegExpMatchArray;
        const result = CODE.calculate(matches, mockContext);
        expect(result).toBe(65);
      });

      it('should handle multi-character text', () => {
        const matches = ['CODE(A1)', 'A1'] as RegExpMatchArray;
        const result = CODE.calculate(matches, mockContext);
        expect(result).toBe(72); // 'H' from "Hello"
      });

      it('should return VALUE error for empty text', () => {
        const matches = ['CODE(C5)', 'C5'] as RegExpMatchArray;
        const result = CODE.calculate(matches, mockContext);
        expect(result).toBe(FormulaError.VALUE);
      });
    });
  });

  describe('Other Text Functions', () => {
    describe('REPT', () => {
      it('should repeat text specified times', () => {
        const matches = ['REPT("Hi", 3)', '"Hi"', '3'] as RegExpMatchArray;
        const result = REPT.calculate(matches, mockContext);
        expect(result).toBe('HiHiHi');
      });

      it('should return empty string for 0 times', () => {
        const matches = ['REPT("Hi", 0)', '"Hi"', '0'] as RegExpMatchArray;
        const result = REPT.calculate(matches, mockContext);
        expect(result).toBe('');
      });

      it('should return VALUE error for negative times', () => {
        const matches = ['REPT("Hi", -1)', '"Hi"', '-1'] as RegExpMatchArray;
        const result = REPT.calculate(matches, mockContext);
        expect(result).toBe(FormulaError.VALUE);
      });
    });

    describe('EXACT', () => {
      it('should return TRUE for exact match', () => {
        const matches = ['EXACT(A1, "Hello")', 'A1', '"Hello"'] as RegExpMatchArray;
        const result = EXACT.calculate(matches, mockContext);
        expect(result).toBe(true);
      });

      it('should return FALSE for case mismatch', () => {
        const matches = ['EXACT(A1, "hello")', 'A1', '"hello"'] as RegExpMatchArray;
        const result = EXACT.calculate(matches, mockContext);
        expect(result).toBe(false);
      });

      it('should compare numbers as text', () => {
        const matches = ['EXACT(B1, "123")', 'B1', '"123"'] as RegExpMatchArray;
        const result = EXACT.calculate(matches, mockContext);
        expect(result).toBe(true);
      });
    });

    describe('TEXTJOIN', () => {
      it('should join text with delimiter', () => {
        const matches = ['TEXTJOIN(", ", TRUE, A1:A2)', '", "', 'TRUE', 'A1:A2'] as RegExpMatchArray;
        const result = TEXTJOIN.calculate(matches, mockContext);
        expect(result).toBe('Hello, World');
      });

      it('should skip empty when ignore_empty is TRUE', () => {
        const matches = ['TEXTJOIN("-", TRUE, C3:C5)', '"-"', 'TRUE', 'C3:C5'] as RegExpMatchArray;
        const result = TEXTJOIN.calculate(matches, mockContext);
        expect(result).toBe('Test123Test');
      });

      it('should include empty when ignore_empty is FALSE', () => {
        const matches = ['TEXTJOIN("-", FALSE, C3:C5)', '"-"', 'FALSE', 'C3:C5'] as RegExpMatchArray;
        const result = TEXTJOIN.calculate(matches, mockContext);
        expect(result).toBe('Test123Test--');
      });

      it('should handle multiple ranges', () => {
        const matches = ['TEXTJOIN(" ", TRUE, A1, A2, B1)', '" "', 'TRUE', 'A1, A2, B1'] as RegExpMatchArray;
        const result = TEXTJOIN.calculate(matches, mockContext);
        expect(result).toBe('Hello World 123');
      });
    });

    describe('CLEAN', () => {
      it('should remove non-printable characters', () => {
        const matches = ['CLEAN("Hello\x00World")', '"Hello\x00World"'] as RegExpMatchArray;
        const result = CLEAN.calculate(matches, mockContext);
        expect(result).toBe('HelloWorld');
      });

      it('should keep printable characters', () => {
        const matches = ['CLEAN(A1)', 'A1'] as RegExpMatchArray;
        const result = CLEAN.calculate(matches, mockContext);
        expect(result).toBe('Hello');
      });
    });

    describe('Unicode Functions', () => {
      describe('UNICHAR', () => {
        it('should return character from Unicode value', () => {
          const matches = ['UNICHAR(65)', '65'] as RegExpMatchArray;
          const result = UNICHAR.calculate(matches, mockContext);
          expect(result).toBe('A');
        });

        it('should handle extended Unicode', () => {
          const matches = ['UNICHAR(8364)', '8364'] as RegExpMatchArray;
          const result = UNICHAR.calculate(matches, mockContext);
          expect(result).toBe('€');
        });

        it('should return VALUE error for invalid code', () => {
          const matches = ['UNICHAR(0)', '0'] as RegExpMatchArray;
          const result = UNICHAR.calculate(matches, mockContext);
          expect(result).toBe(FormulaError.VALUE);
        });
      });

      describe('UNICODE', () => {
        it('should return Unicode value of character', () => {
          const matches = ['UNICODE("A")', '"A"'] as RegExpMatchArray;
          const result = UNICODE.calculate(matches, mockContext);
          expect(result).toBe(65);
        });

        it('should handle extended Unicode', () => {
          const matches = ['UNICODE("€")', '"€"'] as RegExpMatchArray;
          const result = UNICODE.calculate(matches, mockContext);
          expect(result).toBe(8364);
        });

        it('should return VALUE error for empty text', () => {
          const matches = ['UNICODE(C5)', 'C5'] as RegExpMatchArray;
          const result = UNICODE.calculate(matches, mockContext);
          expect(result).toBe(FormulaError.VALUE);
        });
      });
    });
  });
});