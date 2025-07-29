import { describe, it, expect } from 'vitest';
import { SWITCH, LET } from '../newLogicalLogic';
import { FormulaError } from '../../shared/types';
import type { FormulaContext } from '../../shared/types';

// Helper function to create FormulaContext
const createContext = (data: (string | number | boolean | null)[][]): FormulaContext => ({
  data: data.map(row => row.map(cell => ({ value: cell }))),
  row: 0,
  col: 0
});

describe('New Logic Functions', () => {
  const mockContext = createContext([
    [1, 2, 3, 'A', 'B', 'C', 10, 20, 30, 100],
    ['Apple', 'Banana', 'Cherry', 'Dog', 'Cat', 'Bird', 5, 15, 25, 200],
    ['Red', 'Green', 'Blue', 'Sunday', 'Monday', 'Tuesday', 0, -5, 50, 300],
    [true, false, null, '', 'text', 'data', 7, 8, 9, 400],
    [0.5, 1.5, 2.5, 'X', 'Y', 'Z', 12, 18, 24, 500]
  ]);


  describe('SWITCH Function', () => {
    it('should return matching case value', () => {
      const matches = ['SWITCH(A1, 1, "One", 2, "Two", 3, "Three")', 'A1, 1, "One", 2, "Two", 3, "Three"'] as RegExpMatchArray;
      const result = SWITCH.calculate(matches, mockContext);
      expect(result).toBe('One');
    });

    it('should return second matching case', () => {
      const matches = ['SWITCH(B1, 1, "One", 2, "Two", 3, "Three")', 'B1, 1, "One", 2, "Two", 3, "Three"'] as RegExpMatchArray;
      const result = SWITCH.calculate(matches, mockContext);
      expect(result).toBe('Two');
    });

    it('should return default value when no match', () => {
      const matches = ['SWITCH(J1, 1, "One", 2, "Two", 3, "Three", "Default")', 'J1, 1, "One", 2, "Two", 3, "Three", "Default"'] as RegExpMatchArray;
      const result = SWITCH.calculate(matches, mockContext);
      expect(result).toBe('Default');
    });

    it('should return NA error when no match and no default', () => {
      const matches = ['SWITCH(J1, 1, "One", 2, "Two", 3, "Three")', 'J1, 1, "One", 2, "Two", 3, "Three"'] as RegExpMatchArray;
      const result = SWITCH.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NA);
    });

    it('should handle string matching', () => {
      const matches = ['SWITCH(D1, "A", "Letter A", "B", "Letter B", "C", "Letter C")', 'D1, "A", "Letter A", "B", "Letter B", "C", "Letter C"'] as RegExpMatchArray;
      const result = SWITCH.calculate(matches, mockContext);
      expect(result).toBe('Letter A');
    });

    it('should handle cell references in cases and values', () => {
      const matches = ['SWITCH(A1, A1, B1, B1, C1)', 'A1, A1, B1, B1, C1'] as RegExpMatchArray;
      const result = SWITCH.calculate(matches, mockContext);
      expect(result).toBe(2); // A1=1 matches A1, so return B1=2
    });

    it('should handle numeric return values', () => {
      const matches = ['SWITCH(A1, 1, 100, 2, 200, 3, 300)', 'A1, 1, 100, 2, 200, 3, 300'] as RegExpMatchArray;
      const result = SWITCH.calculate(matches, mockContext);
      expect(result).toBe(100);
    });

    it('should handle mixed data types', () => {
      const matches = ['SWITCH(A1, 1, A2, 2, B2, 3, C2)', 'A1, 1, A2, 2, B2, 3, C2'] as RegExpMatchArray;
      const result = SWITCH.calculate(matches, mockContext);
      expect(result).toBe('Apple');
    });

    it('should return VALUE error for insufficient arguments', () => {
      const matches = ['SWITCH(A1, 1)', 'A1, 1'] as RegExpMatchArray;
      const result = SWITCH.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });

    it('should handle zero values', () => {
      const matches = ['SWITCH(G3, 0, "Zero", 1, "One", "Other")', 'G3, 0, "Zero", 1, "One", "Other"'] as RegExpMatchArray;
      const result = SWITCH.calculate(matches, mockContext);
      expect(result).toBe('Zero');
    });

    it('should handle negative values', () => {
      const matches = ['SWITCH(H3, -5, "Negative Five", 0, "Zero", "Positive")', 'H3, -5, "Negative Five", 0, "Zero", "Positive"'] as RegExpMatchArray;
      const result = SWITCH.calculate(matches, mockContext);
      expect(result).toBe('Negative Five');
    });
  });

  describe('LET Function', () => {
    it('should define and use a single variable', () => {
      const matches = ['LET(x, 5, x * 2)', 'x, 5, x * 2'] as RegExpMatchArray;
      const result = LET.calculate(matches, mockContext);
      expect(result).toBe(10);
    });

    it('should handle multiple variables', () => {
      const matches = ['LET(x, 3, y, 4, x + y)', 'x, 3, y, 4, x + y'] as RegExpMatchArray;
      const result = LET.calculate(matches, mockContext);
      expect(result).toBe(7);
    });

    it('should handle cell references in variable values', () => {
      const matches = ['LET(x, A1, y, B1, x + y)', 'x, A1, y, B1, x + y'] as RegExpMatchArray;
      const result = LET.calculate(matches, mockContext);
      expect(result).toBe(3); // A1(1) + B1(2)
    });

    it('should handle complex mathematical expressions', () => {
      const matches = ['LET(a, 2, b, 3, c, 4, a * b + c)', 'a, 2, b, 3, c, 4, a * b + c'] as RegExpMatchArray;
      const result = LET.calculate(matches, mockContext);
      expect(result).toBe(10); // 2 * 3 + 4
    });

    it('should handle parentheses in expressions', () => {
      const matches = ['LET(x, 2, y, 3, (x + y) * 2)', 'x, 2, y, 3, (x + y) * 2'] as RegExpMatchArray;
      const result = LET.calculate(matches, mockContext);
      expect(result).toBe(10); // (2 + 3) * 2
    });

    it('should handle division', () => {
      const matches = ['LET(x, 10, y, 2, x / y)', 'x, 10, y, 2, x / y'] as RegExpMatchArray;
      const result = LET.calculate(matches, mockContext);
      expect(result).toBe(5);
    });

    it('should handle subtraction', () => {
      const matches = ['LET(x, 10, y, 3, x - y)', 'x, 10, y, 3, x - y'] as RegExpMatchArray;
      const result = LET.calculate(matches, mockContext);
      expect(result).toBe(7);
    });

    it('should handle decimal values', () => {
      const matches = ['LET(x, 2.5, y, 1.5, x + y)', 'x, 2.5, y, 1.5, x + y'] as RegExpMatchArray;
      const result = LET.calculate(matches, mockContext);
      expect(result).toBe(4);
    });

    it('should handle string variables in cell references', () => {
      const matches = ['LET(cellRef, A1, cellRef)', 'cellRef, A1, cellRef'] as RegExpMatchArray;
      const result = LET.calculate(matches, mockContext);
      expect(result).toBe(1);
    });

    it('should handle quoted string variables', () => {
      const matches = ['LET(greeting, "Hello", greeting)', 'greeting, "Hello", greeting'] as RegExpMatchArray;
      const result = LET.calculate(matches, mockContext);
      expect(result).toBe('Hello');
    });

    it('should return VALUE error for even number of arguments', () => {
      const matches = ['LET(x, 5, y, 10)', 'x, 5, y, 10'] as RegExpMatchArray; // No formula
      const result = LET.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });

    it('should return VALUE error for too few arguments', () => {
      const matches = ['LET(x)', 'x'] as RegExpMatchArray;
      const result = LET.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });

    it('should return VALUE error for division by zero', () => {
      const matches = ['LET(x, 10, y, 0, x / y)', 'x, 10, y, 0, x / y'] as RegExpMatchArray;
      const result = LET.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });

    it('should handle negative numbers', () => {
      const matches = ['LET(x, -5, y, 3, x + y)', 'x, -5, y, 3, x + y'] as RegExpMatchArray;
      const result = LET.calculate(matches, mockContext);
      expect(result).toBe(-2);
    });

    it('should handle order of operations', () => {
      const matches = ['LET(a, 2, b, 3, c, 4, a + b * c)', 'a, 2, b, 3, c, 4, a + b * c'] as RegExpMatchArray;
      const result = LET.calculate(matches, mockContext);
      expect(result).toBe(14); // 2 + (3 * 4)
    });

    it('should handle variables with cell reference values', () => {
      const matches = ['LET(sum, G1, product, H1, sum + product)', 'sum, G1, product, H1, sum + product'] as RegExpMatchArray;
      const result = LET.calculate(matches, mockContext);
      expect(result).toBe(30); // G1(10) + H1(20)
    });
  });

  describe('Edge Cases', () => {
    it('should handle SWITCH with boolean values', () => {
      const matches = ['SWITCH(A4, TRUE, "IsTrue", FALSE, "IsFalse", "Unknown")', 'A4, TRUE, "IsTrue", FALSE, "IsFalse", "Unknown"'] as RegExpMatchArray;
      const result = SWITCH.calculate(matches, mockContext);
      expect(result).toBe('IsTrue');
    });

    it('should handle SWITCH with null values', () => {
      const matches = ['SWITCH(C4, NULL, "IsNull", "", "IsEmpty", "HasValue")', 'C4, NULL, "IsNull", "", "IsEmpty", "HasValue"'] as RegExpMatchArray;
      const result = SWITCH.calculate(matches, mockContext);
      expect(result).toBe('HasValue'); // null doesn't match the string "NULL"
    });

    it('should handle LET with zero values', () => {
      const matches = ['LET(x, 0, y, 5, x + y)', 'x, 0, y, 5, x + y'] as RegExpMatchArray;
      const result = LET.calculate(matches, mockContext);
      expect(result).toBe(5);
    });

    it('should handle complex nested expressions in LET', () => {
      const matches = ['LET(x, 2, y, 3, z, 4, x * (y + z))', 'x, 2, y, 3, z, 4, x * (y + z)'] as RegExpMatchArray;
      const result = LET.calculate(matches, mockContext);
      expect(result).toBe(14); // 2 * (3 + 4)
    });

    it('should handle SWITCH with decimal matching', () => {
      const matches = ['SWITCH(A5, 0.5, "Half", 1.5, "One and Half", "Other")', 'A5, 0.5, "Half", 1.5, "One and Half", "Other"'] as RegExpMatchArray;
      const result = SWITCH.calculate(matches, mockContext);
      expect(result).toBe('Half');
    });

    it('should handle LET with single variable', () => {
      const matches = ['LET(value, 42, value)', 'value, 42, value'] as RegExpMatchArray;
      const result = LET.calculate(matches, mockContext);
      expect(result).toBe(42);
    });

    it('should handle SWITCH case sensitivity', () => {
      const matches = ['SWITCH(D1, "a", "Lower A", "A", "Upper A", "No Match")', 'D1, "a", "Lower A", "A", "Upper A", "No Match"'] as RegExpMatchArray;
      const result = SWITCH.calculate(matches, mockContext);
      expect(result).toBe('Upper A'); // Case sensitive matching
    });
  });
});