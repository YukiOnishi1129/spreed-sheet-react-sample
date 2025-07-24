import { describe, it, expect } from 'vitest';
import { matchFormula } from '../logic';
import type { FormulaContext } from '../logic/shared/types';

describe('Formula Integration Tests', () => {
  const createContext = (data: (string | number | null)[][]): FormulaContext => ({
    data: data.map(row => row.map(cell => ({ value: cell }))),
    row: 0,
    col: 0
  });

  describe('Complex Formula Calculations', () => {
    it('should calculate nested formulas correctly', () => {
      const context = createContext([
        [100, 200, 300],  // A1, B1, C1
        [50, 75, 100],    // A2, B2, C2
        [null, null, null] // A3, B3, C3 (for results)
      ]);

      // Test SUM of range
      const sumResult = matchFormula('SUM(A1:C1)');
      expect(sumResult).toBeTruthy();
      if (sumResult) {
        const value = sumResult.function.calculate(sumResult.matches, context);
        expect(value).toBe(600); // 100 + 200 + 300
      }

      // Test AVERAGE with mixed values
      const avgResult = matchFormula('AVERAGE(A2:C2)');
      expect(avgResult).toBeTruthy();
      if (avgResult) {
        const value = avgResult.function.calculate(avgResult.matches, context);
        expect(value).toBe(75); // (50 + 75 + 100) / 3
      }
    });

    it('should handle conditional calculations', () => {
      const context = createContext([
        [120000, 100000], // A1: 売上, B1: 目標
        [85000, 100000],  // A2: 売上, B2: 目標
      ]);

      // Test IF with comparison
      const if1Result = matchFormula('IF(A1>=B1,"達成","未達成")');
      expect(if1Result).toBeTruthy();
      if (if1Result) {
        const value = if1Result.function.calculate(if1Result.matches, context);
        expect(value).toBe('達成');
      }

      // Change context for second row
      const context2 = { ...context, row: 1 };
      const if2Result = matchFormula('IF(A2>=B2,"達成","未達成")');
      expect(if2Result).toBeTruthy();
      if (if2Result) {
        const value = if2Result.function.calculate(if2Result.matches, context2);
        expect(value).toBe('未達成');
      }
    });

    it('should handle text functions', () => {
      const context = createContext([
        ['山田', '太郎'],
        ['test@example.com', null]
      ]);

      // Test CONCATENATE
      const concatResult = matchFormula('CONCATENATE(A1," ",B1)');
      expect(concatResult).toBeTruthy();
      if (concatResult) {
        const value = concatResult.function.calculate(concatResult.matches, context);
        expect(value).toBe('山田 太郎');
      }

      // Test LEFT
      const leftResult = matchFormula('LEFT(A2,4)');
      expect(leftResult).toBeTruthy();
      if (leftResult) {
        const value = leftResult.function.calculate(leftResult.matches, { ...context, row: 1 });
        expect(value).toBe('test');
      }
    });

    it('should handle date calculations', () => {
      const context = createContext([
        ['2024-01-01', '2024-12-31'],
        ['2024-03-15', '2024-06-20']
      ]);

      // Test DAYS
      const daysResult = matchFormula('DAYS(B1,A1)');
      expect(daysResult).toBeTruthy();
      if (daysResult) {
        const value = daysResult.function.calculate(daysResult.matches, context);
        expect(value).toBe(365);
      }

      // Test YEAR
      const yearResult = matchFormula('YEAR(A1)');
      expect(yearResult).toBeTruthy();
      if (yearResult) {
        const value = yearResult.function.calculate(yearResult.matches, context);
        expect(value).toBe(2024);
      }
    });

    it('should handle statistical functions', () => {
      const context = createContext([
        [85, 92, 78, 88, 95, 82, null, 90]
      ]);

      // Test MAX
      const maxResult = matchFormula('MAX(A1:H1)');
      expect(maxResult).toBeTruthy();
      if (maxResult) {
        const value = maxResult.function.calculate(maxResult.matches, context);
        expect(value).toBe(95);
      }

      // Test MIN
      const minResult = matchFormula('MIN(A1:H1)');
      expect(minResult).toBeTruthy();
      if (minResult) {
        const value = minResult.function.calculate(minResult.matches, context);
        expect(value).toBe(78);
      }

      // Test COUNT
      const countResult = matchFormula('COUNT(A1:H1)');
      expect(countResult).toBeTruthy();
      if (countResult) {
        const value = countResult.function.calculate(countResult.matches, context);
        expect(value).toBe(7); // null is not counted
      }
    });
  });

  describe('Error Handling', () => {
    it('should handle invalid formulas', () => {
      const result = matchFormula('INVALID_FUNCTION()');
      expect(result).toBeNull();
    });

    it('should handle missing arguments', () => {
      const context = createContext([[1, 2, 3]]);
      
      const result = matchFormula('SUM()');
      expect(result).toBeTruthy();
      if (result) {
        const value = result.function.calculate(result.matches, context);
        expect(value).toBe(0); // Empty sum should return 0
      }
    });
  });
});