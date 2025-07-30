import { describe, it, expect } from 'vitest';
import {
  ROW, ROWS, COLUMN, COLUMNS, ADDRESS, AREAS, FORMULATEXT,
  XMATCH, LAMBDA, HYPERLINK, CHOOSEROWS, CHOOSECOLS, GETPIVOTDATA
} from '../additionalLookupLogic';
import { FormulaError } from '../../shared/types';
import type { FormulaContext } from '../../shared/types';

// Helper function to create FormulaContext
const createContext = (data: any[][], row = 0, col = 0): FormulaContext => ({
  data: data.map(row => row.map(cell => {
    // If cell is already an object with value property, use it as is
    if (typeof cell === 'object' && cell !== null && 'value' in cell) {
      return cell;
    }
    // Otherwise wrap it
    return { value: cell };
  })),
  row,
  col
});

describe('Additional Lookup and Reference Functions', () => {
  const mockContext = createContext([
    ['A1', 'B1', 'C1', 'D1', 'E1'],
    ['A2', 'B2', 'C2', 'D2', 'E2'],
    ['A3', 'B3', 'C3', 'D3', 'E3'],
    ['A4', 'B4', 'C4', 'D4', 'E4'],
    ['A5', 'B5', 'C5', 'D5', 'E5']
  ], 2, 1); // Current position at C3 (row 2, col 1)

  describe('ROW Function', () => {
    it('should return row number for cell reference', () => {
      const matches = ['ROW(B3)', 'B3'] as RegExpMatchArray;
      const result = ROW.calculate(matches, mockContext);
      expect(result).toBe(3);
    });

    it('should return current row when no argument', () => {
      const matches = ['ROW()', undefined] as RegExpMatchArray;
      const result = ROW.calculate(matches, mockContext);
      expect(result).toBe(3); // row 2 (0-based) + 1
    });

    it('should return first row for range reference', () => {
      const matches = ['ROW(A2:C5)', 'A2:C5'] as RegExpMatchArray;
      const result = ROW.calculate(matches, mockContext);
      expect(result).toBe(2);
    });

    it('should return VALUE error for invalid reference', () => {
      const matches = ['ROW(invalid)', 'invalid'] as RegExpMatchArray;
      const result = ROW.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });
  });

  describe('ROWS Function', () => {
    it('should return 1 for single cell', () => {
      const matches = ['ROWS(A1)', 'A1'] as RegExpMatchArray;
      const result = ROWS.calculate(matches, mockContext);
      expect(result).toBe(1);
    });

    it('should return number of rows in range', () => {
      const matches = ['ROWS(A1:C5)', 'A1:C5'] as RegExpMatchArray;
      const result = ROWS.calculate(matches, mockContext);
      expect(result).toBe(5);
    });

    it('should handle reversed range', () => {
      const matches = ['ROWS(A5:C1)', 'A5:C1'] as RegExpMatchArray;
      const result = ROWS.calculate(matches, mockContext);
      expect(result).toBe(5);
    });

    it('should return VALUE error for invalid reference', () => {
      const matches = ['ROWS(invalid)', 'invalid'] as RegExpMatchArray;
      const result = ROWS.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });
  });

  describe('COLUMN Function', () => {
    it('should return column number for cell reference', () => {
      const matches = ['COLUMN(B3)', 'B3'] as RegExpMatchArray;
      const result = COLUMN.calculate(matches, mockContext);
      expect(result).toBe(2);
    });

    it('should return current column when no argument', () => {
      const matches = ['COLUMN()', undefined] as RegExpMatchArray;
      const result = COLUMN.calculate(matches, mockContext);
      expect(result).toBe(2); // col 1 (0-based) + 1
    });

    it('should return first column for range reference', () => {
      const matches = ['COLUMN(B2:E5)', 'B2:E5'] as RegExpMatchArray;
      const result = COLUMN.calculate(matches, mockContext);
      expect(result).toBe(2);
    });

    it('should handle multi-letter columns', () => {
      const matches = ['COLUMN(AA1)', 'AA1'] as RegExpMatchArray;
      const result = COLUMN.calculate(matches, mockContext);
      expect(result).toBe(27);
    });
  });

  describe('COLUMNS Function', () => {
    it('should return 1 for single cell', () => {
      const matches = ['COLUMNS(A1)', 'A1'] as RegExpMatchArray;
      const result = COLUMNS.calculate(matches, mockContext);
      expect(result).toBe(1);
    });

    it('should return number of columns in range', () => {
      const matches = ['COLUMNS(A1:E3)', 'A1:E3'] as RegExpMatchArray;
      const result = COLUMNS.calculate(matches, mockContext);
      expect(result).toBe(5);
    });

    it('should handle reversed range', () => {
      const matches = ['COLUMNS(E1:A3)', 'E1:A3'] as RegExpMatchArray;
      const result = COLUMNS.calculate(matches, mockContext);
      expect(result).toBe(5);
    });
  });

  describe('ADDRESS Function', () => {
    it('should create absolute address (default)', () => {
      const matches = ['ADDRESS(3, 2)', '3', '2'] as RegExpMatchArray;
      const result = ADDRESS.calculate(matches, mockContext);
      expect(result).toBe('$B$3');
    });

    it('should create relative address', () => {
      const matches = ['ADDRESS(3, 2, 4)', '3', '2', '4'] as RegExpMatchArray;
      const result = ADDRESS.calculate(matches, mockContext);
      expect(result).toBe('B3');
    });

    it('should create mixed address (row absolute, column relative)', () => {
      const matches = ['ADDRESS(3, 2, 3)', '3', '2', '3'] as RegExpMatchArray;
      const result = ADDRESS.calculate(matches, mockContext);
      expect(result).toBe('B$3');
    });

    it('should create mixed address (row relative, column absolute)', () => {
      const matches = ['ADDRESS(3, 2, 2)', '3', '2', '2'] as RegExpMatchArray;
      const result = ADDRESS.calculate(matches, mockContext);
      expect(result).toBe('$B3');
    });

    it('should handle R1C1 format', () => {
      const matches = ['ADDRESS(3, 2, 1, FALSE)', '3', '2', '1', 'FALSE'] as RegExpMatchArray;
      const result = ADDRESS.calculate(matches, mockContext);
      expect(result).toBe('R3C2');
    });

    it('should include sheet name', () => {
      const matches = ['ADDRESS(3, 2, 1, TRUE, "Sheet1")', '3', '2', '1', 'TRUE', '"Sheet1"'] as RegExpMatchArray;
      const result = ADDRESS.calculate(matches, mockContext);
      expect(result).toBe('Sheet1!$B$3');
    });

    it('should return VALUE error for invalid row/column', () => {
      const matches = ['ADDRESS(0, 2)', '0', '2'] as RegExpMatchArray;
      const result = ADDRESS.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });
  });

  describe('AREAS Function', () => {
    it('should count single area', () => {
      const matches = ['AREAS(A1:B2)', 'A1:B2'] as RegExpMatchArray;
      const result = AREAS.calculate(matches, mockContext);
      expect(result).toBe(1);
    });

    it('should count multiple areas', () => {
      const matches = ['AREAS(A1:B2,C3:D4,E5)', 'A1:B2,C3:D4,E5'] as RegExpMatchArray;
      const result = AREAS.calculate(matches, mockContext);
      expect(result).toBe(3);
    });

    it('should return VALUE error for invalid reference', () => {
      const matches = ['AREAS(invalid)', 'invalid'] as RegExpMatchArray;
      const result = AREAS.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });
  });

  describe('FORMULATEXT Function', () => {
    it('should return NA error when cell has no formula', () => {
      const matches = ['FORMULATEXT(A1)', 'A1'] as RegExpMatchArray;
      const result = FORMULATEXT.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NA);
    });

    it('should return formula text when cell has formula', () => {
      const formulaContext = createContext([
        [{ value: 'A1', formula: '=SUM(B1:B3)' }]
      ]);
      
      const matches = ['FORMULATEXT(A1)', 'A1'] as RegExpMatchArray;
      const result = FORMULATEXT.calculate(matches, formulaContext);
      expect(result).toBe('=SUM(B1:B3)');
    });

    it('should return REF error for out of range cell', () => {
      const matches = ['FORMULATEXT(Z100)', 'Z100'] as RegExpMatchArray;
      const result = FORMULATEXT.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.REF);
    });
  });

  describe('XMATCH Function', () => {
    it('should find exact match (default)', () => {
      const searchContext = createContext([
        ['apple', 'banana', 'cherry', 'date']
      ]);
      
      const matches = ['XMATCH("banana", A1:D1)', '"banana"', 'A1:D1'] as RegExpMatchArray;
      const result = XMATCH.calculate(matches, searchContext);
      expect(result).toBe(2);
    });

    it('should find next smallest value', () => {
      const searchContext = createContext([
        [1, 5, 10, 15, 20]
      ]);
      
      const matches = ['XMATCH(12, A1:E1, -1)', '12', 'A1:E1', '-1'] as RegExpMatchArray;
      const result = XMATCH.calculate(matches, searchContext);
      expect(result).toBe(3); // 10 is the largest value <= 12
    });

    it('should find next largest value', () => {
      const searchContext = createContext([
        [1, 5, 10, 15, 20]
      ]);
      
      const matches = ['XMATCH(12, A1:E1, 1)', '12', 'A1:E1', '1'] as RegExpMatchArray;
      const result = XMATCH.calculate(matches, searchContext);
      expect(result).toBe(4); // 15 is the smallest value >= 12
    });

    it('should handle wildcard matching', () => {
      const searchContext = createContext([
        ['apple', 'banana', 'cherry', 'date']
      ]);
      
      const matches = ['XMATCH("ban*", A1:D1, 2)', '"ban*"', 'A1:D1', '2'] as RegExpMatchArray;
      const result = XMATCH.calculate(matches, searchContext);
      expect(result).toBe(2);
    });

    it('should return NA error when not found', () => {
      const searchContext = createContext([
        ['apple', 'banana', 'cherry', 'date']
      ]);
      
      const matches = ['XMATCH("grape", A1:D1)', '"grape"', 'A1:D1'] as RegExpMatchArray;
      const result = XMATCH.calculate(matches, searchContext);
      expect(result).toBe(FormulaError.NA);
    });
  });

  describe('LAMBDA Function', () => {
    it('should return VALUE error (not supported for direct cell use)', () => {
      const matches = ['LAMBDA(x, x*2)', 'x, x*2'] as RegExpMatchArray;
      const result = LAMBDA.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });
  });

  describe('HYPERLINK Function', () => {
    it('should return friendly name', () => {
      const matches = ['HYPERLINK("http://example.com", "Example")', '"http://example.com"', '"Example"'] as RegExpMatchArray;
      const result = HYPERLINK.calculate(matches, mockContext);
      expect(result).toBe('Example');
    });

    it('should use URL as friendly name when not provided', () => {
      const matches = ['HYPERLINK("http://example.com")', '"http://example.com"'] as RegExpMatchArray;
      const result = HYPERLINK.calculate(matches, mockContext);
      expect(result).toBe('http://example.com');
    });

    it('should handle cell references', () => {
      const linkContext = createContext([
        ['http://example.com', 'My Link']
      ]);
      
      const matches = ['HYPERLINK(A1, B1)', 'A1', 'B1'] as RegExpMatchArray;
      const result = HYPERLINK.calculate(matches, linkContext);
      expect(result).toBe('My Link');
    });
  });

  describe('CHOOSEROWS Function', () => {
    it('should select specific rows', () => {
      const matches = ['CHOOSEROWS(A1:C3, 1, 3)', 'A1:C3', '1, 3'] as RegExpMatchArray;
      const result = CHOOSEROWS.calculate(matches, mockContext);
      expect(Array.isArray(result)).toBe(true);
      if (Array.isArray(result)) {
        expect(result).toEqual([
          ['A1', 'B1', 'C1'],
          ['A3', 'B3', 'C3']
        ]);
      }
    });

    it('should return VALUE error for invalid row number', () => {
      const matches = ['CHOOSEROWS(A1:C3, 5)', 'A1:C3', '5'] as RegExpMatchArray;
      const result = CHOOSEROWS.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });

    it('should return REF error for invalid range', () => {
      const matches = ['CHOOSEROWS(invalid, 1)', 'invalid', '1'] as RegExpMatchArray;
      const result = CHOOSEROWS.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.REF);
    });
  });

  describe('CHOOSECOLS Function', () => {
    it('should select specific columns', () => {
      const matches = ['CHOOSECOLS(A1:C3, 1, 3)', 'A1:C3', '1, 3'] as RegExpMatchArray;
      const result = CHOOSECOLS.calculate(matches, mockContext);
      expect(Array.isArray(result)).toBe(true);
      if (Array.isArray(result)) {
        expect(result).toEqual([
          ['A1', 'C1'],
          ['A2', 'C2'],
          ['A3', 'C3']
        ]);
      }
    });

    it('should return VALUE error for invalid column number', () => {
      const matches = ['CHOOSECOLS(A1:C3, 5)', 'A1:C3', '5'] as RegExpMatchArray;
      const result = CHOOSECOLS.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });

    it('should return REF error for invalid range', () => {
      const matches = ['CHOOSECOLS(invalid, 1)', 'invalid', '1'] as RegExpMatchArray;
      const result = CHOOSECOLS.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.REF);
    });
  });

  describe('GETPIVOTDATA Function', () => {
    it('should return REF error (pivot table not implemented)', () => {
      const matches = ['GETPIVOTDATA("Sales", A1)', '"Sales"', 'A1'] as RegExpMatchArray;
      const result = GETPIVOTDATA.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.REF);
    });
  });

  describe('Edge Cases', () => {
    it('should handle ADDRESS with large column numbers', () => {
      const matches = ['ADDRESS(1, 702)', '1', '702'] as RegExpMatchArray; // Column ZZ
      const result = ADDRESS.calculate(matches, mockContext);
      expect(result).toBe('$ZZ$1');
    });

    it('should handle ROW with empty context', () => {
      const emptyContext = createContext([]);
      const matches = ['ROW()', undefined] as RegExpMatchArray;
      const result = ROW.calculate(matches, emptyContext);
      expect(result).toBe(1); // row 0 + 1
    });

    it('should handle XMATCH with reverse search mode', () => {
      const searchContext = createContext([
        ['a', 'b', 'c', 'b', 'a']
      ]);
      
      const matches = ['XMATCH("b", A1:E1, 0, -1)', '"b"', 'A1:E1', '0', '-1'] as RegExpMatchArray;
      const result = XMATCH.calculate(matches, searchContext);
      expect(result).toBe(4); // Last occurrence of 'b'
    });

    it('should handle ADDRESS with sheet name containing spaces', () => {
      const matches = ['ADDRESS(1, 1, 1, TRUE, "My Sheet")', '1', '1', '1', 'TRUE', '"My Sheet"'] as RegExpMatchArray;
      const result = ADDRESS.calculate(matches, mockContext);
      expect(result).toBe("'My Sheet'!$A$1");
    });

    it('should handle COLUMN calculation for multi-character columns correctly', () => {
      const matches = ['COLUMN(AB1)', 'AB1'] as RegExpMatchArray;
      const result = COLUMN.calculate(matches, mockContext);
      expect(result).toBe(28); // A=1, B=2, so AB = 1*26 + 2 = 28
    });
  });
});