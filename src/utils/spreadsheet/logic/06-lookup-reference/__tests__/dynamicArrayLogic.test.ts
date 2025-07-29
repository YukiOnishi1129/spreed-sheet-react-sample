import { describe, it, expect } from 'vitest';
import {
  SORTBY, TAKE, DROP, EXPAND, HSTACK, VSTACK,
  TOCOL, TOROW, WRAPROWS, WRAPCOLS
} from '../dynamicArrayLogic';
import { FormulaError } from '../../shared/types';
import type { FormulaContext } from '../../shared/types';

// Helper function to create FormulaContext
const createContext = (data: (string | number | boolean | null)[][]): FormulaContext => ({
  data: data.map(row => row.map(cell => ({ value: cell }))),
  row: 0,
  col: 0
});

describe('Dynamic Array Functions', () => {
  const mockContext = createContext([
    ['Name', 'Age', 'Score', 10],
    ['Alice', 25, 85, 20],
    ['Bob', 30, 92, 30],
    ['Charlie', 28, 78, 40],
    ['David', 35, 95, 50],
    [1, 2, 3, 4],
    [5, 6, 7, 8],
    [9, 10, 11, 12]
  ]);

  describe('SORTBY Function', () => {
    it('should sort array by another array in ascending order', () => {
      const matches = ['SORTBY(A2:A5, B2:B5)', 'A2:A5', 'B2:B5'] as RegExpMatchArray;
      const result = SORTBY.calculate(matches, mockContext);
      expect(Array.isArray(result)).toBe(true);
      if (Array.isArray(result)) {
        expect(result[0]).toEqual(['Alice']); // Age 25 (smallest)
        expect(result[1]).toEqual(['Charlie']); // Age 28
        expect(result[2]).toEqual(['Bob']); // Age 30
        expect(result[3]).toEqual(['David']); // Age 35 (largest)
      }
    });

    it('should sort array by another array in descending order', () => {
      const matches = ['SORTBY(A2:A5, B2:B5, -1)', 'A2:A5', 'B2:B5', '-1'] as RegExpMatchArray;
      const result = SORTBY.calculate(matches, mockContext);
      expect(Array.isArray(result)).toBe(true);
      if (Array.isArray(result)) {
        expect(result[0]).toEqual(['David']); // Age 35 (largest)
        expect(result[1]).toEqual(['Bob']); // Age 30
        expect(result[2]).toEqual(['Charlie']); // Age 28
        expect(result[3]).toEqual(['Alice']); // Age 25 (smallest)
      }
    });

    it('should return VALUE error for mismatched array sizes', () => {
      const matches = ['SORTBY(A2:A5, B2:B3)', 'A2:A5', 'B2:B3'] as RegExpMatchArray;
      const result = SORTBY.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });

    it('should return REF error for invalid range', () => {
      const matches = ['SORTBY(invalid, B2:B5)', 'invalid', 'B2:B5'] as RegExpMatchArray;
      const result = SORTBY.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.REF);
    });
  });

  describe('TAKE Function', () => {
    it('should take first n rows', () => {
      const matches = ['TAKE(A1:C3, 2)', 'A1:C3', '2'] as RegExpMatchArray;
      const result = TAKE.calculate(matches, mockContext);
      expect(Array.isArray(result)).toBe(true);
      if (Array.isArray(result)) {
        expect(result).toHaveLength(2);
        expect(result[0]).toEqual(['Name', 'Age', 'Score']);
        expect(result[1]).toEqual(['Alice', 25, 85]);
      }
    });

    it('should take last n rows with negative number', () => {
      const matches = ['TAKE(A1:C3, -2)', 'A1:C3', '-2'] as RegExpMatchArray;
      const result = TAKE.calculate(matches, mockContext);
      expect(Array.isArray(result)).toBe(true);
      if (Array.isArray(result)) {
        expect(result).toHaveLength(2);
        expect(result[0]).toEqual(['Alice', 25, 85]);
        expect(result[1]).toEqual(['Bob', 30, 92]);
      }
    });

    it('should take specific number of rows and columns', () => {
      const matches = ['TAKE(A1:C3, 2, 2)', 'A1:C3', '2', '2'] as RegExpMatchArray;
      const result = TAKE.calculate(matches, mockContext);
      expect(Array.isArray(result)).toBe(true);
      if (Array.isArray(result)) {
        expect(result).toHaveLength(2);
        expect(result[0]).toEqual(['Name', 'Age']);
        expect(result[1]).toEqual(['Alice', 25]);
      }
    });

    it('should return empty array for zero rows', () => {
      const matches = ['TAKE(A1:C3, 0)', 'A1:C3', '0'] as RegExpMatchArray;
      const result = TAKE.calculate(matches, mockContext);
      expect(result).toEqual([]);
    });
  });

  describe('DROP Function', () => {
    it('should drop first n rows', () => {
      const matches = ['DROP(A1:C3, 1)', 'A1:C3', '1'] as RegExpMatchArray;
      const result = DROP.calculate(matches, mockContext);
      expect(Array.isArray(result)).toBe(true);
      if (Array.isArray(result)) {
        expect(result).toHaveLength(2);
        expect(result[0]).toEqual(['Alice', 25, 85]);
        expect(result[1]).toEqual(['Bob', 30, 92]);
      }
    });

    it('should drop last n rows with negative number', () => {
      const matches = ['DROP(A1:C3, -1)', 'A1:C3', '-1'] as RegExpMatchArray;
      const result = DROP.calculate(matches, mockContext);
      expect(Array.isArray(result)).toBe(true);
      if (Array.isArray(result)) {
        expect(result).toHaveLength(2);
        expect(result[0]).toEqual(['Name', 'Age', 'Score']);
        expect(result[1]).toEqual(['Alice', 25, 85]);
      }
    });

    it('should drop specific number of rows and columns', () => {
      const matches = ['DROP(A1:C3, 1, 1)', 'A1:C3', '1', '1'] as RegExpMatchArray;
      const result = DROP.calculate(matches, mockContext);
      expect(Array.isArray(result)).toBe(true);
      if (Array.isArray(result)) {
        expect(result).toHaveLength(2);
        expect(result[0]).toEqual([25, 85]);
        expect(result[1]).toEqual([30, 92]);
      }
    });
  });

  describe('EXPAND Function', () => {
    it('should expand array with specified dimensions', () => {
      const matches = ['EXPAND(A1:B2, 3, 4)', 'A1:B2', '3', '4'] as RegExpMatchArray;
      const result = EXPAND.calculate(matches, mockContext);
      expect(Array.isArray(result)).toBe(true);
      if (Array.isArray(result)) {
        expect(result).toHaveLength(3);
        expect(result[0]).toHaveLength(4);
        expect(result[0]).toEqual(['Name', 'Age', FormulaError.NA, FormulaError.NA]);
        expect(result[1]).toEqual(['Alice', 25, FormulaError.NA, FormulaError.NA]);
        expect(result[2]).toEqual([FormulaError.NA, FormulaError.NA, FormulaError.NA, FormulaError.NA]);
      }
    });

    it('should expand array with custom pad value', () => {
      const matches = ['EXPAND(A1:B2, 3, 3, 0)', 'A1:B2', '3', '3', '0'] as RegExpMatchArray;
      const result = EXPAND.calculate(matches, mockContext);
      expect(Array.isArray(result)).toBe(true);
      if (Array.isArray(result)) {
        expect(result[0]).toEqual(['Name', 'Age', '0']);
        expect(result[2]).toEqual(['0', '0', '0']);
      }
    });

    it('should return VALUE error for smaller dimensions', () => {
      const matches = ['EXPAND(A1:C3, 2, 2)', 'A1:C3', '2', '2'] as RegExpMatchArray;
      const result = EXPAND.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });
  });

  describe('HSTACK Function', () => {
    it('should stack arrays horizontally', () => {
      const matches = ['HSTACK(A1:A3, B1:B3)', 'A1:A3, B1:B3'] as RegExpMatchArray;
      const result = HSTACK.calculate(matches, mockContext);
      expect(Array.isArray(result)).toBe(true);
      if (Array.isArray(result)) {
        expect(result).toHaveLength(3);
        expect(result[0]).toEqual(['Name', 'Age']);
        expect(result[1]).toEqual(['Alice', 25]);
        expect(result[2]).toEqual(['Bob', 30]);
      }
    });

    it('should stack arrays with different heights', () => {
      const matches = ['HSTACK(A1:A2, B1:B3)', 'A1:A2, B1:B3'] as RegExpMatchArray;
      const result = HSTACK.calculate(matches, mockContext);
      expect(Array.isArray(result)).toBe(true);
      if (Array.isArray(result)) {
        expect(result).toHaveLength(3);
        expect(result[2]).toEqual([FormulaError.NA, 30]); // First array padded with #N/A
      }
    });

    it('should stack single values', () => {
      const matches = ['HSTACK(A1, B1)', 'A1, B1'] as RegExpMatchArray;
      const result = HSTACK.calculate(matches, mockContext);
      expect(Array.isArray(result)).toBe(true);
      if (Array.isArray(result)) {
        expect(result).toEqual([['Name', 'Age']]);
      }
    });
  });

  describe('VSTACK Function', () => {
    it('should stack arrays vertically', () => {
      const matches = ['VSTACK(A1:C1, A2:C2)', 'A1:C1, A2:C2'] as RegExpMatchArray;
      const result = VSTACK.calculate(matches, mockContext);
      expect(Array.isArray(result)).toBe(true);
      if (Array.isArray(result)) {
        expect(result).toHaveLength(2);
        expect(result[0]).toEqual(['Name', 'Age', 'Score']);
        expect(result[1]).toEqual(['Alice', 25, 85]);
      }
    });

    it('should stack arrays with different widths', () => {
      const matches = ['VSTACK(A1:B1, A2:C2)', 'A1:B1, A2:C2'] as RegExpMatchArray;
      const result = VSTACK.calculate(matches, mockContext);
      expect(Array.isArray(result)).toBe(true);
      if (Array.isArray(result)) {
        expect(result).toHaveLength(2);
        expect(result[0]).toEqual(['Name', 'Age', FormulaError.NA]); // First array padded with #N/A
        expect(result[1]).toEqual(['Alice', 25, 85]);
      }
    });
  });

  describe('TOCOL Function', () => {
    it('should convert array to single column', () => {
      const matches = ['TOCOL(A1:B2)', 'A1:B2'] as RegExpMatchArray;
      const result = TOCOL.calculate(matches, mockContext);
      expect(Array.isArray(result)).toBe(true);
      if (Array.isArray(result)) {
        expect(result).toEqual([['Name'], ['Age'], ['Alice'], [25]]);
      }
    });

    it('should scan by column when specified', () => {
      const matches = ['TOCOL(A1:B2, 0, TRUE)', 'A1:B2', '0', 'TRUE'] as RegExpMatchArray;
      const result = TOCOL.calculate(matches, mockContext);
      expect(Array.isArray(result)).toBe(true);
      if (Array.isArray(result)) {
        expect(result).toEqual([['Name'], ['Alice'], ['Age'], [25]]);
      }
    });

    it('should ignore blanks when specified', () => {
      const testContext = createContext([
        ['A', ''],
        ['', 'B']
      ]);
      
      const matches = ['TOCOL(A1:B2, 1)', 'A1:B2', '1'] as RegExpMatchArray;
      const result = TOCOL.calculate(matches, testContext);
      expect(Array.isArray(result)).toBe(true);
      if (Array.isArray(result)) {
        expect(result).toEqual([['A'], ['B']]);
      }
    });
  });

  describe('TOROW Function', () => {
    it('should convert array to single row', () => {
      const matches = ['TOROW(A1:B2)', 'A1:B2'] as RegExpMatchArray;
      const result = TOROW.calculate(matches, mockContext);
      expect(Array.isArray(result)).toBe(true);
      if (Array.isArray(result)) {
        expect(result).toEqual([['Name', 'Age', 'Alice', 25]]);
      }
    });

    it('should scan by column when specified', () => {
      const matches = ['TOROW(A1:B2, 0, TRUE)', 'A1:B2', '0', 'TRUE'] as RegExpMatchArray;
      const result = TOROW.calculate(matches, mockContext);
      expect(Array.isArray(result)).toBe(true);
      if (Array.isArray(result)) {
        expect(result).toEqual([['Name', 'Alice', 'Age', 25]]);
      }
    });
  });

  describe('WRAPROWS Function', () => {
    it('should wrap vector into rows', () => {
      const matches = ['WRAPROWS(A6:D6, 2)', 'A6:D6', '2'] as RegExpMatchArray;
      const result = WRAPROWS.calculate(matches, mockContext);
      expect(Array.isArray(result)).toBe(true);
      if (Array.isArray(result)) {
        expect(result).toHaveLength(2);
        expect(result[0]).toEqual([1, 2]);
        expect(result[1]).toEqual([3, 4]);
      }
    });

    it('should pad incomplete rows', () => {
      const matches = ['WRAPROWS(A6:C6, 2)', 'A6:C6', '2'] as RegExpMatchArray;
      const result = WRAPROWS.calculate(matches, mockContext);
      expect(Array.isArray(result)).toBe(true);
      if (Array.isArray(result)) {
        expect(result).toHaveLength(2);
        expect(result[0]).toEqual([1, 2]);
        expect(result[1]).toEqual([3, FormulaError.NA]);
      }
    });

    it('should use custom pad value', () => {
      const matches = ['WRAPROWS(A6:C6, 2, 0)', 'A6:C6', '2', '0'] as RegExpMatchArray;
      const result = WRAPROWS.calculate(matches, mockContext);
      expect(Array.isArray(result)).toBe(true);
      if (Array.isArray(result)) {
        expect(result[1]).toEqual([3, '0']);
      }
    });

    it('should return VALUE error for invalid wrap count', () => {
      const matches = ['WRAPROWS(A6:C6, 0)', 'A6:C6', '0'] as RegExpMatchArray;
      const result = WRAPROWS.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });
  });

  describe('WRAPCOLS Function', () => {
    it('should wrap vector into columns', () => {
      const matches = ['WRAPCOLS(A6:D6, 2)', 'A6:D6', '2'] as RegExpMatchArray;
      const result = WRAPCOLS.calculate(matches, mockContext);
      expect(Array.isArray(result)).toBe(true);
      if (Array.isArray(result)) {
        expect(result).toHaveLength(2);
        expect(result[0]).toEqual([1, 3]);
        expect(result[1]).toEqual([2, 4]);
      }
    });

    it('should pad incomplete columns', () => {
      const matches = ['WRAPCOLS(A6:C6, 2)', 'A6:C6', '2'] as RegExpMatchArray;
      const result = WRAPCOLS.calculate(matches, mockContext);
      expect(Array.isArray(result)).toBe(true);
      if (Array.isArray(result)) {
        expect(result).toHaveLength(2);
        expect(result[0]).toEqual([1, 3]);
        expect(result[1]).toEqual([2, FormulaError.NA]);
      }
    });

    it('should use custom pad value', () => {
      const matches = ['WRAPCOLS(A6:C6, 2, 0)', 'A6:C6', '2', '0'] as RegExpMatchArray;
      const result = WRAPCOLS.calculate(matches, mockContext);
      expect(Array.isArray(result)).toBe(true);
      if (Array.isArray(result)) {
        expect(result[1]).toEqual([2, '0']);
      }
    });
  });

  describe('Edge Cases', () => {
    it('should handle TAKE with more rows than available', () => {
      const matches = ['TAKE(A1:A2, 5)', 'A1:A2', '5'] as RegExpMatchArray;
      const result = TAKE.calculate(matches, mockContext);
      expect(Array.isArray(result)).toBe(true);
      if (Array.isArray(result)) {
        expect(result).toHaveLength(2); // Only available rows
      }
    });

    it('should handle DROP with more rows than available', () => {
      const matches = ['DROP(A1:A2, 5)', 'A1:A2', '5'] as RegExpMatchArray;
      const result = DROP.calculate(matches, mockContext);
      expect(Array.isArray(result)).toBe(true);
      if (Array.isArray(result)) {
        expect(result).toEqual([]); // Empty array
      }
    });

    it('should handle SORTBY with null values', () => {
      const nullContext = createContext([
        ['A', 'B', 'C'],
        [1, null, 3]
      ]);
      
      const matches = ['SORTBY(A1:C1, A2:C2)', 'A1:C1', 'A2:C2'] as RegExpMatchArray;
      const result = SORTBY.calculate(matches, nullContext);
      expect(Array.isArray(result)).toBe(true);
      if (Array.isArray(result)) {
        expect(result[0]).toEqual(['A']); // 1 comes first
        expect(result[1]).toEqual(['C']); // 3 comes second
        expect(result[2]).toEqual(['B']); // null comes last
      }
    });

    it('should handle single cell range in HSTACK', () => {
      const matches = ['HSTACK(A1)', 'A1'] as RegExpMatchArray;
      const result = HSTACK.calculate(matches, mockContext);
      expect(Array.isArray(result)).toBe(true);
      if (Array.isArray(result)) {
        expect(result).toEqual([['Name']]);
      }
    });

    it('should handle REF error for invalid range in dynamic arrays', () => {
      const matches = ['TOCOL(invalid)', 'invalid'] as RegExpMatchArray;
      const result = TOCOL.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.REF);
    });
  });
});