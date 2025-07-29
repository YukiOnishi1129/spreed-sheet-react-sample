import { describe, it, expect } from 'vitest';
import {
  VLOOKUP, HLOOKUP, INDEX, MATCH, LOOKUP, XLOOKUP,
  OFFSET, INDIRECT, CHOOSE, TRANSPOSE, FILTER, SORT, UNIQUE
} from '../lookupLogic';
import { FormulaError } from '../../shared/types';
import type { FormulaContext } from '../../shared/types';

// Helper function to create FormulaContext
const createContext = (data: (string | number | boolean | null)[][]): FormulaContext => ({
  data: data.map(row => row.map(cell => ({ value: cell }))),
  row: 0,
  col: 0
});

describe('Lookup and Reference Functions', () => {
  const mockContext = createContext([
    ['Name', 'Age', 'City', 'Score', 100],
    ['Alice', 25, 'Tokyo', 85, 200],
    ['Bob', 30, 'Osaka', 92, 300],
    ['Charlie', 28, 'Kyoto', 78, 400],
    ['David', 35, 'Tokyo', 95, 500],
    [1, 2, 3, 4, 5],
    [10, 20, 30, 40, 50],
    ['Apple', 'Banana', 'Cherry', 'Date', 'Elderberry'],
    ['Red', 'Yellow', 'Red', 'Brown', 'Purple'],
    ['A', 'B', 'C', 'D', 'E']
  ]);

  describe('VLOOKUP Function', () => {
    it('should find exact match', () => {
      const matches = ['VLOOKUP("Bob", A1:D5, 2, FALSE)', '"Bob"', 'A1:D5', '2', 'FALSE'] as RegExpMatchArray;
      const result = VLOOKUP.calculate(matches, mockContext);
      expect(result).toBe(30);
    });

    it('should find exact match with different column', () => {
      const matches = ['VLOOKUP("Alice", A1:D5, 3, FALSE)', '"Alice"', 'A1:D5', '3', 'FALSE'] as RegExpMatchArray;
      const result = VLOOKUP.calculate(matches, mockContext);
      expect(result).toBe('Tokyo');
    });

    it('should return NA error when value not found in exact match', () => {
      const matches = ['VLOOKUP("Eve", A1:D5, 2, FALSE)', '"Eve"', 'A1:D5', '2', 'FALSE'] as RegExpMatchArray;
      const result = VLOOKUP.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NA);
    });

    it('should return REF error for invalid column index', () => {
      const matches = ['VLOOKUP("Alice", A1:D5, 10, FALSE)', '"Alice"', 'A1:D5', '10', 'FALSE'] as RegExpMatchArray;
      const result = VLOOKUP.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.REF);
    });

    it('should return VALUE error for column index less than 1', () => {
      const matches = ['VLOOKUP("Alice", A1:D5, 0, FALSE)', '"Alice"', 'A1:D5', '0', 'FALSE'] as RegExpMatchArray;
      const result = VLOOKUP.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });

    it('should handle approximate match (default)', () => {
      const matches = ['VLOOKUP("Bob", A1:D5, 2)', '"Bob"', 'A1:D5', '2'] as RegExpMatchArray;
      const result = VLOOKUP.calculate(matches, mockContext);
      expect(result).toBe(30);
    });
  });

  describe('HLOOKUP Function', () => {
    it('should find value in horizontal lookup', () => {
      const matches = ['HLOOKUP(2, A6:E7, 2, FALSE)', '2', 'A6:E7', '2', 'FALSE'] as RegExpMatchArray;
      const result = HLOOKUP.calculate(matches, mockContext);
      expect(result).toBe(20);
    });

    it('should return NA error when value not found', () => {
      const matches = ['HLOOKUP(6, A6:E7, 2, FALSE)', '6', 'A6:E7', '2', 'FALSE'] as RegExpMatchArray;
      const result = HLOOKUP.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NA);
    });

    it('should return REF error for invalid row index', () => {
      const matches = ['HLOOKUP(2, A6:E7, 5, FALSE)', '2', 'A6:E7', '5', 'FALSE'] as RegExpMatchArray;
      const result = HLOOKUP.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.REF);
    });
  });

  describe('INDEX Function', () => {
    it('should return value at specified row and column', () => {
      const matches = ['INDEX(A1:D5, 2, 3)', 'A1:D5', '2', '3'] as RegExpMatchArray;
      const result = INDEX.calculate(matches, mockContext);
      expect(result).toBe('Tokyo');
    });

    it('should return value at specified row (single column)', () => {
      const matches = ['INDEX(A1:A5, 3)', 'A1:A5', '3'] as RegExpMatchArray;
      const result = INDEX.calculate(matches, mockContext);
      expect(result).toBe('Bob');
    });

    it('should return REF error for out of range', () => {
      const matches = ['INDEX(A1:D5, 10, 1)', 'A1:D5', '10', '1'] as RegExpMatchArray;
      const result = INDEX.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.REF);
    });

    it('should return VALUE error for invalid row number', () => {
      const matches = ['INDEX(A1:D5, 0, 1)', 'A1:D5', '0', '1'] as RegExpMatchArray;
      const result = INDEX.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });
  });

  describe('MATCH Function', () => {
    it('should find exact match position', () => {
      const matches = ['MATCH("Bob", A1:A5, 0)', '"Bob"', 'A1:A5', '0'] as RegExpMatchArray;
      const result = MATCH.calculate(matches, mockContext);
      expect(result).toBe(3);
    });

    it('should find approximate match (less than or equal)', () => {
      const matches = ['MATCH(25, B1:B5, 1)', '25', 'B1:B5', '1'] as RegExpMatchArray;
      const result = MATCH.calculate(matches, mockContext);
      expect(result).toBe(2); // Age column has 25 at position 2
    });

    it('should return NA error when not found', () => {
      const matches = ['MATCH("Eve", A1:A5, 0)', '"Eve"', 'A1:A5', '0'] as RegExpMatchArray;
      const result = MATCH.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NA);
    });

    it('should handle numeric exact match', () => {
      const matches = ['MATCH(3, A6:E6, 0)', '3', 'A6:E6', '0'] as RegExpMatchArray;
      const result = MATCH.calculate(matches, mockContext);
      expect(result).toBe(3);
    });
  });

  describe('LOOKUP Function', () => {
    it('should perform vector lookup', () => {
      const matches = ['LOOKUP(3, A6:E6, A7:E7)', '3', 'A6:E6', 'A7:E7'] as RegExpMatchArray;
      const result = LOOKUP.calculate(matches, mockContext);
      expect(result).toBe(30);
    });

    it('should handle lookup without result vector', () => {
      const matches = ['LOOKUP(3, A6:E6)', '3', 'A6:E6'] as RegExpMatchArray;
      const result = LOOKUP.calculate(matches, mockContext);
      expect(result).toBe(3);
    });

    it('should return NA error when arrays have different lengths', () => {
      const matches = ['LOOKUP(3, A6:E6, A7:D7)', '3', 'A6:E6', 'A7:D7'] as RegExpMatchArray;
      const result = LOOKUP.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NA);
    });
  });

  describe('XLOOKUP Function', () => {
    it('should find exact match', () => {
      const matches = ['XLOOKUP("Bob", A1:A5, B1:B5)', '"Bob"', 'A1:A5', 'B1:B5'] as RegExpMatchArray;
      const result = XLOOKUP.calculate(matches, mockContext);
      expect(result).toBe(30);
    });

    it('should return if_not_found value when not found', () => {
      const matches = ['XLOOKUP("Eve", A1:A5, B1:B5, "Not Found")', '"Eve"', 'A1:A5', 'B1:B5', '"Not Found"'] as RegExpMatchArray;
      const result = XLOOKUP.calculate(matches, mockContext);
      expect(result).toBe('Not Found');
    });

    it('should return NA error when not found and no if_not_found', () => {
      const matches = ['XLOOKUP("Eve", A1:A5, B1:B5)', '"Eve"', 'A1:A5', 'B1:B5'] as RegExpMatchArray;
      const result = XLOOKUP.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NA);
    });

    it('should return REF error for mismatched array sizes', () => {
      const matches = ['XLOOKUP("Bob", A1:A5, B1:C5)', '"Bob"', 'A1:A5', 'B1:C5'] as RegExpMatchArray;
      const result = XLOOKUP.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.REF);
    });
  });

  describe('OFFSET Function', () => {
    it('should return offset cell value', () => {
      const matches = ['OFFSET(A1, 1, 1)', 'A1', '1', '1'] as RegExpMatchArray;
      const result = OFFSET.calculate(matches, mockContext);
      expect(result).toBe(25);
    });

    it('should handle zero offset', () => {
      const matches = ['OFFSET(A2, 0, 0)', 'A2', '0', '0'] as RegExpMatchArray;
      const result = OFFSET.calculate(matches, mockContext);
      expect(result).toBe('Alice');
    });

    it('should return REF error for negative position', () => {
      const matches = ['OFFSET(A1, -1, 0)', 'A1', '-1', '0'] as RegExpMatchArray;
      const result = OFFSET.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.REF);
    });

    it('should return VALUE error for invalid parameters', () => {
      const matches = ['OFFSET(A1, "text", 1)', 'A1', '"text"', '1'] as RegExpMatchArray;
      const result = OFFSET.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });
  });

  describe('INDIRECT Function', () => {
    it('should return value from indirect reference', () => {
      const matches = ['INDIRECT("A2")', '"A2"'] as RegExpMatchArray;
      const result = INDIRECT.calculate(matches, mockContext);
      expect(result).toBe('Alice');
    });

    it('should handle cell reference in another cell', () => {
      // This would require the cell to contain "B3" as text
      const matches = ['INDIRECT("B3")', '"B3"'] as RegExpMatchArray;
      const result = INDIRECT.calculate(matches, mockContext);
      expect(result).toBe(30);
    });

    it('should return REF error for invalid reference', () => {
      const matches = ['INDIRECT("invalid")', '"invalid"'] as RegExpMatchArray;
      const result = INDIRECT.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.REF);
    });
  });

  describe('CHOOSE Function', () => {
    it('should return value at specified index', () => {
      const matches = ['CHOOSE(2, "First", "Second", "Third")', '2', '"First", "Second", "Third"'] as RegExpMatchArray;
      const result = CHOOSE.calculate(matches, mockContext);
      expect(result).toBe('Second');
    });

    it('should return VALUE error for index out of range', () => {
      const matches = ['CHOOSE(5, "First", "Second", "Third")', '5', '"First", "Second", "Third"'] as RegExpMatchArray;
      const result = CHOOSE.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });

    it('should return VALUE error for invalid index', () => {
      const matches = ['CHOOSE(0, "First", "Second", "Third")', '0', '"First", "Second", "Third"'] as RegExpMatchArray;
      const result = CHOOSE.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });
  });

  describe('TRANSPOSE Function', () => {
    it('should transpose a range', () => {
      const matches = ['TRANSPOSE(A6:C6)', 'A6:C6'] as RegExpMatchArray;
      const result = TRANSPOSE.calculate(matches, mockContext);
      expect(Array.isArray(result)).toBe(true);
      if (Array.isArray(result)) {
        expect(result).toEqual([1, 2, 3]);
      }
    });

    it('should handle single cell', () => {
      const matches = ['TRANSPOSE(A1)', 'A1'] as RegExpMatchArray;
      const result = TRANSPOSE.calculate(matches, mockContext);
      expect(result).toBe('Name');
    });

    it('should return REF error for invalid range', () => {
      const matches = ['TRANSPOSE("invalid")', '"invalid"'] as RegExpMatchArray;
      const result = TRANSPOSE.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.REF);
    });
  });

  describe('FILTER Function', () => {
    it('should filter values based on criteria', () => {
      // Create a simple test case where we filter based on TRUE/FALSE values
      const filterContext = createContext([
        [1, 2, 3, 4, 5],
        [true, false, true, false, true]
      ]);
      
      const matches = ['FILTER(A1:E1, A2:E2)', 'A1:E1', 'A2:E2'] as RegExpMatchArray;
      const result = FILTER.calculate(matches, filterContext);
      expect(Array.isArray(result)).toBe(true);
      if (Array.isArray(result)) {
        expect(result).toEqual([1, 3, 5]);
      }
    });

    it('should return if_empty value when no matches', () => {
      const filterContext = createContext([
        [1, 2, 3, 4, 5],
        [false, false, false, false, false]
      ]);
      
      const matches = ['FILTER(A1:E1, A2:E2, "No Data")', 'A1:E1', 'A2:E2', '"No Data"'] as RegExpMatchArray;
      const result = FILTER.calculate(matches, filterContext);
      expect(result).toBe('No Data');
    });

    it('should return VALUE error for mismatched array sizes', () => {
      const matches = ['FILTER(A1:E1, A2:D2)', 'A1:E1', 'A2:D2'] as RegExpMatchArray;
      const result = FILTER.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });
  });

  describe('SORT Function', () => {
    it('should sort single column ascending', () => {
      const matches = ['SORT(A6:A9)', 'A6:A9'] as RegExpMatchArray;
      const result = SORT.calculate(matches, mockContext);
      expect(Array.isArray(result)).toBe(true);
    });

    it('should sort single column descending', () => {
      const matches = ['SORT(A6:A9, 1, -1)', 'A6:A9', '1', '-1'] as RegExpMatchArray;
      const result = SORT.calculate(matches, mockContext);
      expect(Array.isArray(result)).toBe(true);
    });

    it('should return VALUE error for by_col=TRUE (unsupported)', () => {
      const matches = ['SORT(A6:C8, 1, 1, TRUE)', 'A6:C8', '1', '1', 'TRUE'] as RegExpMatchArray;
      const result = SORT.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });

    it('should return VALUE error for invalid sort index', () => {
      const matches = ['SORT(A6:C8, 5, 1)', 'A6:C8', '5', '1'] as RegExpMatchArray;
      const result = SORT.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });
  });

  describe('UNIQUE Function', () => {
    it('should return unique values from single column', () => {
      const uniqueContext = createContext([
        [1, 2, 2, 3, 3, 3],
      ]);
      
      const matches = ['UNIQUE(A1:F1)', 'A1:F1'] as RegExpMatchArray;
      const result = UNIQUE.calculate(matches, uniqueContext);
      expect(Array.isArray(result)).toBe(true);
      if (Array.isArray(result)) {
        expect(result).toEqual([1, 2, 3]);
      }
    });

    it('should return exactly once occurrences', () => {
      const uniqueContext = createContext([
        [1, 2, 2, 3, 3, 3],
      ]);
      
      const matches = ['UNIQUE(A1:F1, FALSE, TRUE)', 'A1:F1', 'FALSE', 'TRUE'] as RegExpMatchArray;
      const result = UNIQUE.calculate(matches, uniqueContext);
      expect(Array.isArray(result)).toBe(true);
      if (Array.isArray(result)) {
        expect(result).toEqual([1]); // Only 1 appears exactly once
      }
    });

    it('should return VALUE error for by_col=TRUE (unsupported)', () => {
      const matches = ['UNIQUE(A1:C3, TRUE)', 'A1:C3', 'TRUE'] as RegExpMatchArray;
      const result = UNIQUE.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });
  });

  describe('Edge Cases', () => {
    it('should handle empty cells in VLOOKUP', () => {
      const emptyContext = createContext([
        ['', 'Value1'],
        ['Key', 'Value2']
      ]);
      
      const matches = ['VLOOKUP("Key", A1:B2, 2, FALSE)', '"Key"', 'A1:B2', '2', 'FALSE'] as RegExpMatchArray;
      const result = VLOOKUP.calculate(matches, emptyContext);
      expect(result).toBe('Value2');
    });

    it('should handle null values in arrays', () => {
      const nullContext = createContext([
        [null, 1, 2, null, 3]
      ]);
      
      const matches = ['MATCH(2, A1:E1, 0)', '2', 'A1:E1', '0'] as RegExpMatchArray;
      const result = MATCH.calculate(matches, nullContext);
      expect(result).toBe(3);
    });

    it('should handle string to number conversion in MATCH', () => {
      const matches = ['MATCH("30", B1:B5, 0)', '"30"', 'B1:B5', '0'] as RegExpMatchArray;
      const result = MATCH.calculate(matches, mockContext);
      expect(result).toBe(3);
    });
  });
});