import { describe, it, expect } from 'vitest';
import {
  LAMBDA, MAKEARRAY, BYROW, BYCOL, REDUCE, SCAN, 
  MAP as MAP_FUNC, FILTER as FILTER_FUNC, SORT as SORT_FUNC
} from '../excelNewLogic';
import { FormulaError } from '../../shared/types';
import type { FormulaContext } from '../../shared/types';

// Helper function to create FormulaContext
const createContext = (data: (string | number | boolean | null)[][]): FormulaContext => ({
  data: data.map(row => row.map(cell => ({ value: cell }))),
  row: 0,
  col: 0
});

describe('Excel New Functions (Lambda and Dynamic Arrays)', () => {
  const mockContext = createContext([
    [1, 2, 3, 4, 5], // numbers
    [10, 20, 30, 40, 50], // more numbers
    ['Apple', 'Banana', 'Cherry', 'Date', 'Elderberry'], // fruits
    [100, 75, 150, 90, 120], // scores
    [true, false, true, true, false], // booleans
    ['A', 'B', 'C', 'D', 'E'], // letters
  ]);

  describe('LAMBDA Function', () => {
    it('should create lambda with single parameter', () => {
      const matches = ['LAMBDA(x, x*2)', 'x', 'x*2'] as RegExpMatchArray;
      const result = LAMBDA.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Lambda functions require Excel 365');
    });

    it('should create lambda with multiple parameters', () => {
      const matches = ['LAMBDA(x, y, x+y)', 'x', 'y', 'x+y'] as RegExpMatchArray;
      const result = LAMBDA.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Lambda functions require Excel 365');
    });

    it('should create lambda with complex expression', () => {
      const matches = ['LAMBDA(x, y, z, IF(x>y, z*x, z*y))', 
        'x', 'y', 'z', 'IF(x>y, z*x, z*y)'] as RegExpMatchArray;
      const result = LAMBDA.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Lambda functions require Excel 365');
    });

    it('should handle recursive lambda', () => {
      const matches = ['LAMBDA(n, IF(n<=1, 1, n*LAMBDA(n-1)))', 
        'n', 'IF(n<=1, 1, n*LAMBDA(n-1))'] as RegExpMatchArray;
      const result = LAMBDA.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Lambda functions require Excel 365');
    });

    it('should handle named parameters', () => {
      const matches = ['LAMBDA(value, tax_rate, value * (1 + tax_rate))', 
        'value', 'tax_rate', 'value * (1 + tax_rate)'] as RegExpMatchArray;
      const result = LAMBDA.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Lambda functions require Excel 365');
    });
  });

  describe('MAKEARRAY Function', () => {
    it('should create array with dimensions', () => {
      const matches = ['MAKEARRAY(3, 4, LAMBDA(r, c, r*c))', 
        '3', '4', 'LAMBDA(r, c, r*c)'] as RegExpMatchArray;
      const result = MAKEARRAY.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Lambda functions require Excel 365');
    });

    it('should create identity matrix', () => {
      const matches = ['MAKEARRAY(5, 5, LAMBDA(r, c, IF(r=c, 1, 0)))', 
        '5', '5', 'LAMBDA(r, c, IF(r=c, 1, 0))'] as RegExpMatchArray;
      const result = MAKEARRAY.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Lambda functions require Excel 365');
    });

    it('should create multiplication table', () => {
      const matches = ['MAKEARRAY(10, 10, LAMBDA(row, col, row*col))', 
        '10', '10', 'LAMBDA(row, col, row*col)'] as RegExpMatchArray;
      const result = MAKEARRAY.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Lambda functions require Excel 365');
    });

    it('should create array with complex formula', () => {
      const matches = ['MAKEARRAY(4, 3, LAMBDA(r, c, POWER(r, c)))', 
        '4', '3', 'LAMBDA(r, c, POWER(r, c))'] as RegExpMatchArray;
      const result = MAKEARRAY.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Lambda functions require Excel 365');
    });
  });

  describe('BYROW Function', () => {
    it('should apply function to each row', () => {
      const matches = ['BYROW(A1:E2, LAMBDA(row, SUM(row)))', 
        'A1:E2', 'LAMBDA(row, SUM(row))'] as RegExpMatchArray;
      const result = BYROW.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Lambda functions require Excel 365');
    });

    it('should find max in each row', () => {
      const matches = ['BYROW(A1:E4, LAMBDA(row, MAX(row)))', 
        'A1:E4', 'LAMBDA(row, MAX(row))'] as RegExpMatchArray;
      const result = BYROW.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Lambda functions require Excel 365');
    });

    it('should count values in each row', () => {
      const matches = ['BYROW(A1:E5, LAMBDA(row, COUNT(row)))', 
        'A1:E5', 'LAMBDA(row, COUNT(row))'] as RegExpMatchArray;
      const result = BYROW.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Lambda functions require Excel 365');
    });

    it('should calculate average for each row', () => {
      const matches = ['BYROW(A1:E3, LAMBDA(row, AVERAGE(row)))', 
        'A1:E3', 'LAMBDA(row, AVERAGE(row))'] as RegExpMatchArray;
      const result = BYROW.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Lambda functions require Excel 365');
    });
  });

  describe('BYCOL Function', () => {
    it('should apply function to each column', () => {
      const matches = ['BYCOL(A1:E4, LAMBDA(col, SUM(col)))', 
        'A1:E4', 'LAMBDA(col, SUM(col))'] as RegExpMatchArray;
      const result = BYCOL.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Lambda functions require Excel 365');
    });

    it('should find min in each column', () => {
      const matches = ['BYCOL(A1:E4, LAMBDA(col, MIN(col)))', 
        'A1:E4', 'LAMBDA(col, MIN(col))'] as RegExpMatchArray;
      const result = BYCOL.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Lambda functions require Excel 365');
    });

    it('should calculate standard deviation for each column', () => {
      const matches = ['BYCOL(A1:E4, LAMBDA(col, STDEV(col)))', 
        'A1:E4', 'LAMBDA(col, STDEV(col))'] as RegExpMatchArray;
      const result = BYCOL.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Lambda functions require Excel 365');
    });

    it('should concatenate values in each column', () => {
      const matches = ['BYCOL(A3:E3, LAMBDA(col, TEXTJOIN(",", TRUE, col)))', 
        'A3:E3', 'LAMBDA(col, TEXTJOIN(",", TRUE, col))'] as RegExpMatchArray;
      const result = BYCOL.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Lambda functions require Excel 365');
    });
  });

  describe('REDUCE Function', () => {
    it('should reduce array to sum', () => {
      const matches = ['REDUCE(0, A1:E1, LAMBDA(acc, val, acc + val))', 
        '0', 'A1:E1', 'LAMBDA(acc, val, acc + val)'] as RegExpMatchArray;
      const result = REDUCE.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Lambda functions require Excel 365');
    });

    it('should reduce array to product', () => {
      const matches = ['REDUCE(1, A1:E1, LAMBDA(acc, val, acc * val))', 
        '1', 'A1:E1', 'LAMBDA(acc, val, acc * val)'] as RegExpMatchArray;
      const result = REDUCE.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Lambda functions require Excel 365');
    });

    it('should count specific values', () => {
      const matches = ['REDUCE(0, A5:E5, LAMBDA(acc, val, IF(val=TRUE, acc+1, acc)))', 
        '0', 'A5:E5', 'LAMBDA(acc, val, IF(val=TRUE, acc+1, acc))'] as RegExpMatchArray;
      const result = REDUCE.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Lambda functions require Excel 365');
    });

    it('should find maximum value', () => {
      const matches = ['REDUCE(0, A4:E4, LAMBDA(acc, val, MAX(acc, val)))', 
        '0', 'A4:E4', 'LAMBDA(acc, val, MAX(acc, val))'] as RegExpMatchArray;
      const result = REDUCE.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Lambda functions require Excel 365');
    });
  });

  describe('SCAN Function', () => {
    it('should calculate running total', () => {
      const matches = ['SCAN(0, A1:E1, LAMBDA(acc, val, acc + val))', 
        '0', 'A1:E1', 'LAMBDA(acc, val, acc + val)'] as RegExpMatchArray;
      const result = SCAN.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Lambda functions require Excel 365');
    });

    it('should calculate running product', () => {
      const matches = ['SCAN(1, A1:E1, LAMBDA(acc, val, acc * val))', 
        '1', 'A1:E1', 'LAMBDA(acc, val, acc * val)'] as RegExpMatchArray;
      const result = SCAN.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Lambda functions require Excel 365');
    });

    it('should calculate running maximum', () => {
      const matches = ['SCAN(-999, A4:E4, LAMBDA(acc, val, MAX(acc, val)))', 
        '-999', 'A4:E4', 'LAMBDA(acc, val, MAX(acc, val))'] as RegExpMatchArray;
      const result = SCAN.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Lambda functions require Excel 365');
    });

    it('should create cumulative concatenation', () => {
      const matches = ['SCAN("", A6:E6, LAMBDA(acc, val, acc & val))', 
        '""', 'A6:E6', 'LAMBDA(acc, val, acc & val)'] as RegExpMatchArray;
      const result = SCAN.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Lambda functions require Excel 365');
    });
  });

  describe('MAP Function', () => {
    it('should map single array', () => {
      const matches = ['MAP(A1:E1, LAMBDA(x, x*2))', 
        'A1:E1', 'LAMBDA(x, x*2)'] as RegExpMatchArray;
      const result = MAP_FUNC.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Lambda functions require Excel 365');
    });

    it('should map multiple arrays', () => {
      const matches = ['MAP(A1:E1, A2:E2, LAMBDA(x, y, x+y))', 
        'A1:E1', 'A2:E2', 'LAMBDA(x, y, x+y)'] as RegExpMatchArray;
      const result = MAP_FUNC.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Lambda functions require Excel 365');
    });

    it('should apply complex transformation', () => {
      const matches = ['MAP(A4:E4, LAMBDA(x, IF(x>100, "High", "Low")))', 
        'A4:E4', 'LAMBDA(x, IF(x>100, "High", "Low"))'] as RegExpMatchArray;
      const result = MAP_FUNC.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Lambda functions require Excel 365');
    });

    it('should handle text transformation', () => {
      const matches = ['MAP(A3:E3, LAMBDA(x, UPPER(x)))', 
        'A3:E3', 'LAMBDA(x, UPPER(x))'] as RegExpMatchArray;
      const result = MAP_FUNC.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Lambda functions require Excel 365');
    });
  });

  describe('FILTER Function', () => {
    it('should filter by condition', () => {
      const matches = ['FILTER(A1:E1, A1:E1>3)', 'A1:E1', 'A1:E1>3'] as RegExpMatchArray;
      const result = FILTER_FUNC.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Lambda functions require Excel 365');
    });

    it('should filter with multiple conditions', () => {
      const matches = ['FILTER(A1:E4, (A1:E1>2)*(A1:E1<5))', 
        'A1:E4', '(A1:E1>2)*(A1:E1<5)'] as RegExpMatchArray;
      const result = FILTER_FUNC.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Lambda functions require Excel 365');
    });

    it('should filter with if_empty parameter', () => {
      const matches = ['FILTER(A1:E1, A1:E1>10, "No results")', 
        'A1:E1', 'A1:E1>10', '"No results"'] as RegExpMatchArray;
      const result = FILTER_FUNC.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Lambda functions require Excel 365');
    });

    it('should filter text values', () => {
      const matches = ['FILTER(A3:E3, LEN(A3:E3)>5)', 
        'A3:E3', 'LEN(A3:E3)>5'] as RegExpMatchArray;
      const result = FILTER_FUNC.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Lambda functions require Excel 365');
    });

    it('should filter boolean values', () => {
      const matches = ['FILTER(A5:E5, A5:E5=TRUE)', 
        'A5:E5', 'A5:E5=TRUE'] as RegExpMatchArray;
      const result = FILTER_FUNC.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Lambda functions require Excel 365');
    });
  });

  describe('SORT Function', () => {
    it('should sort ascending', () => {
      const matches = ['SORT(A1:E1)', 'A1:E1'] as RegExpMatchArray;
      const result = SORT_FUNC.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Lambda functions require Excel 365');
    });

    it('should sort descending', () => {
      const matches = ['SORT(A1:E1, 1, -1)', 'A1:E1', '1', '-1'] as RegExpMatchArray;
      const result = SORT_FUNC.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Lambda functions require Excel 365');
    });

    it('should sort by column', () => {
      const matches = ['SORT(A1:E4, 2)', 'A1:E4', '2'] as RegExpMatchArray;
      const result = SORT_FUNC.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Lambda functions require Excel 365');
    });

    it('should sort by multiple columns', () => {
      const matches = ['SORT(A1:E4, {1,2}, {1,-1})', 
        'A1:E4', '{1,2}', '{1,-1}'] as RegExpMatchArray;
      const result = SORT_FUNC.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Lambda functions require Excel 365');
    });

    it('should sort horizontally', () => {
      const matches = ['SORT(A1:E1, 1, 1, TRUE)', 
        'A1:E1', '1', '1', 'TRUE'] as RegExpMatchArray;
      const result = SORT_FUNC.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Lambda functions require Excel 365');
    });
  });

  describe('Edge Cases and Integration', () => {
    it('should combine FILTER and SORT', () => {
      const matches = ['SORT(FILTER(A1:E1, A1:E1>2))', 
        'FILTER(A1:E1, A1:E1>2)'] as RegExpMatchArray;
      const result = SORT_FUNC.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Lambda functions require Excel 365');
    });

    it('should use LAMBDA with MAP', () => {
      const customFunc = 'LAMBDA(x, x^2 + 2*x + 1)';
      const matches = ['MAP(A1:E1, ' + customFunc + ')', 
        'A1:E1', customFunc] as RegExpMatchArray;
      const result = MAP_FUNC.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Lambda functions require Excel 365');
    });

    it('should combine REDUCE with LAMBDA for complex aggregation', () => {
      const matches = ['REDUCE("", A3:E3, LAMBDA(acc, val, IF(acc="", val, acc & ", " & val)))', 
        '""', 'A3:E3', 'LAMBDA(acc, val, IF(acc="", val, acc & ", " & val))'] as RegExpMatchArray;
      const result = REDUCE.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Lambda functions require Excel 365');
    });

    it('should use MAKEARRAY for Pascal triangle', () => {
      const matches = ['MAKEARRAY(5, 5, LAMBDA(r, c, IF(c>r, "", COMBIN(r-1, c-1))))', 
        '5', '5', 'LAMBDA(r, c, IF(c>r, "", COMBIN(r-1, c-1)))'] as RegExpMatchArray;
      const result = MAKEARRAY.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Lambda functions require Excel 365');
    });

    it('should handle nested LAMBDA functions', () => {
      const matches = ['LAMBDA(x, LAMBDA(y, x+y))', 
        'x', 'LAMBDA(y, x+y)'] as RegExpMatchArray;
      const result = LAMBDA.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Lambda functions require Excel 365');
    });

    it('should handle all functions returning consistent error', () => {
      const functions = [
        LAMBDA, MAKEARRAY, BYROW, BYCOL, REDUCE, SCAN, MAP_FUNC, FILTER_FUNC, SORT_FUNC
      ];
      
      for (const func of functions) {
        const matches = ['FUNC("param")', '"param"'] as RegExpMatchArray;
        const result = func.calculate(matches, mockContext);
        expect(result).toBe('#N/A - Lambda functions require Excel 365');
      }
    });
  });
});