import { describe, it, expect } from 'vitest';
import {
  DSTDEV, DSTDEVP, DVAR, DVARP
} from '../statisticalDatabaseLogic';
import { FormulaError } from '../../shared/types';
import type { FormulaContext } from '../../shared/types';

// Helper function to create FormulaContext
const createContext = (data: (string | number | boolean | null)[][]): FormulaContext => ({
  data: data.map(row => row.map(cell => ({ value: cell }))),
  row: 0,
  col: 0
});

describe('Statistical Database Functions', () => {
  // Sample database: Test scores by subject and student
  const mockContext = createContext([
    // Database area (A1:D6)
    ['Student', 'Subject', 'Score', 'Grade'],
    ['Alice', 'Math', 85, 'B'],
    ['Bob', 'Math', 92, 'A'],
    ['Charlie', 'Math', 78, 'C'],
    ['Alice', 'English', 88, 'B'],
    ['Bob', 'English', 95, 'A'],
    // Criteria area (F1:G2)
    ['Subject'],
    ['Math']
  ]);

  describe('DSTDEV Function (Sample Standard Deviation)', () => {
    it('should calculate standard deviation for Math scores', () => {
      const matches = ['DSTDEV(A1:D6, "Score", F1:G2)', 'A1:D6', 'Score', 'F1:G2'] as RegExpMatchArray;
      const result = DSTDEV.calculate(matches, mockContext);
      
      // Manual calculation: scores are 85, 92, 78
      // Mean = (85 + 92 + 78) / 3 = 85
      // Variance = [(85-85)² + (92-85)² + (78-85)²] / (3-1) = [0 + 49 + 49] / 2 = 49
      // Standard deviation = √49 = 7
      expect(result).toBeCloseTo(7, 5);
    });

    it('should return DIV0 error for insufficient data points', () => {
      const singlePointContext = createContext([
        ['Student', 'Subject', 'Score'],
        ['Alice', 'Math', 85],
        ['Subject'],
        ['Math']
      ]);
      
      const matches = ['DSTDEV(A1:C2, "Score", A3:A4)', 'A1:C2', 'Score', 'A3:A4'] as RegExpMatchArray;
      const result = DSTDEV.calculate(matches, singlePointContext);
      expect(result).toBe(FormulaError.DIV0);
    });

    it('should handle no matching records', () => {
      const noMatchContext = createContext([
        ['Student', 'Subject', 'Score'],
        ['Alice', 'Math', 85],
        ['Bob', 'English', 92],
        ['Subject'],
        ['Science'] // No Science records
      ]);
      
      const matches = ['DSTDEV(A1:C3, "Score", A4:A5)', 'A1:C3', 'Score', 'A4:A5'] as RegExpMatchArray;
      const result = DSTDEV.calculate(matches, noMatchContext);
      expect(result).toBe(FormulaError.DIV0);
    });

    it('should return VALUE error for invalid database reference', () => {
      const matches = ['DSTDEV(invalid, "Score", F1:G2)', 'invalid', 'Score', 'F1:G2'] as RegExpMatchArray;
      const result = DSTDEV.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.REF);
    });

    it('should handle complex criteria with comparison operators', () => {
      const complexContext = createContext([
        ['Student', 'Subject', 'Score'],
        ['Alice', 'Math', 85],
        ['Bob', 'Math', 92],
        ['Charlie', 'Math', 78],
        ['David', 'Math', 95],
        ['Score'],
        ['>80'] // Scores greater than 80
      ]);
      
      const matches = ['DSTDEV(A1:C5, "Score", A6:A7)', 'A1:C5', 'Score', 'A6:A7'] as RegExpMatchArray;
      const result = DSTDEV.calculate(matches, complexContext);
      
      // Matching scores: 85, 92, 95
      // Mean = (85 + 92 + 95) / 3 = 90.67
      // Sample std dev calculation
      const mean = (85 + 92 + 95) / 3;
      const variance = (Math.pow(85 - mean, 2) + Math.pow(92 - mean, 2) + Math.pow(95 - mean, 2)) / 2;
      const expectedStdDev = Math.sqrt(variance);
      
      expect(result).toBeCloseTo(expectedStdDev, 5);
    });
  });

  describe('DSTDEVP Function (Population Standard Deviation)', () => {
    it('should calculate population standard deviation for Math scores', () => {
      const matches = ['DSTDEVP(A1:D6, "Score", F1:G2)', 'A1:D6', 'Score', 'F1:G2'] as RegExpMatchArray;
      const result = DSTDEVP.calculate(matches, mockContext);
      
      // Manual calculation: scores are 85, 92, 78
      // Mean = (85 + 92 + 78) / 3 = 85
      // Population variance = [(85-85)² + (92-85)² + (78-85)²] / 3 = [0 + 49 + 49] / 3 = 32.67
      // Population standard deviation = √32.67 ≈ 5.71
      expect(result).toBeCloseTo(5.715, 3);
    });

    it('should return DIV0 error for no matching records', () => {
      const noMatchContext = createContext([
        ['Student', 'Subject', 'Score'],
        ['Alice', 'Math', 85],
        ['Subject'],
        ['Science']
      ]);
      
      const matches = ['DSTDEVP(A1:C2, "Score", A3:A4)', 'A1:C2', 'Score', 'A3:A4'] as RegExpMatchArray;
      const result = DSTDEVP.calculate(matches, noMatchContext);
      expect(result).toBe(FormulaError.DIV0);
    });

    it('should handle single data point', () => {
      const singlePointContext = createContext([
        ['Student', 'Subject', 'Score'],
        ['Alice', 'Math', 85],
        ['Subject'],
        ['Math']
      ]);
      
      const matches = ['DSTDEVP(A1:C2, "Score", A3:A4)', 'A1:C2', 'Score', 'A3:A4'] as RegExpMatchArray;
      const result = DSTDEVP.calculate(matches, singlePointContext);
      expect(result).toBe(0); // Population std dev of single value is 0
    });

    it('should return REF error for invalid criteria reference', () => {
      const matches = ['DSTDEVP(A1:D6, "Score", invalid)', 'A1:D6', 'Score', 'invalid'] as RegExpMatchArray;
      const result = DSTDEVP.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.REF);
    });
  });

  describe('DVAR Function (Sample Variance)', () => {
    it('should calculate sample variance for Math scores', () => {
      const matches = ['DVAR(A1:D6, "Score", F1:G2)', 'A1:D6', 'Score', 'F1:G2'] as RegExpMatchArray;
      const result = DVAR.calculate(matches, mockContext);
      
      // Manual calculation: scores are 85, 92, 78
      // Mean = (85 + 92 + 78) / 3 = 85
      // Sample variance = [(85-85)² + (92-85)² + (78-85)²] / (3-1) = [0 + 49 + 49] / 2 = 49
      expect(result).toBeCloseTo(49, 5);
    });

    it('should return DIV0 error for insufficient data points', () => {
      const singlePointContext = createContext([
        ['Student', 'Subject', 'Score'],
        ['Alice', 'Math', 85],
        ['Subject'],
        ['Math']
      ]);
      
      const matches = ['DVAR(A1:C2, "Score", A3:A4)', 'A1:C2', 'Score', 'A3:A4'] as RegExpMatchArray;
      const result = DVAR.calculate(matches, singlePointContext);
      expect(result).toBe(FormulaError.DIV0);
    });

    it('should handle identical values', () => {
      const identicalContext = createContext([
        ['Student', 'Subject', 'Score'],
        ['Alice', 'Math', 90],
        ['Bob', 'Math', 90],
        ['Charlie', 'Math', 90],
        ['Subject'],
        ['Math']
      ]);
      
      const matches = ['DVAR(A1:C4, "Score", A5:A6)', 'A1:C4', 'Score', 'A5:A6'] as RegExpMatchArray;
      const result = DVAR.calculate(matches, identicalContext);
      expect(result).toBe(0); // Variance of identical values is 0
    });

    it('should handle mixed numeric and non-numeric values', () => {
      const mixedContext = createContext([
        ['Student', 'Subject', 'Score'],
        ['Alice', 'Math', 85],
        ['Bob', 'Math', 'N/A'],
        ['Charlie', 'Math', 92],
        ['David', 'Math', null],
        ['Subject'],
        ['Math']
      ]);
      
      const matches = ['DVAR(A1:C5, "Score", A6:A7)', 'A1:C5', 'Score', 'A6:A7'] as RegExpMatchArray;
      const result = DVAR.calculate(matches, mixedContext);
      
      // Only 85 and 92 are numeric and match criteria
      // Mean = (85 + 92) / 2 = 88.5
      // Sample variance = [(85-88.5)² + (92-88.5)²] / (2-1) = [12.25 + 12.25] / 1 = 24.5
      expect(result).toBeCloseTo(24.5, 5);
    });
  });

  describe('DVARP Function (Population Variance)', () => {
    it('should calculate population variance for Math scores', () => {
      const matches = ['DVARP(A1:D6, "Score", F1:G2)', 'A1:D6', 'Score', 'F1:G2'] as RegExpMatchArray;
      const result = DVARP.calculate(matches, mockContext);
      
      // Manual calculation: scores are 85, 92, 78
      // Mean = (85 + 92 + 78) / 3 = 85
      // Population variance = [(85-85)² + (92-85)² + (78-85)²] / 3 = [0 + 49 + 49] / 3 ≈ 32.67
      expect(result).toBeCloseTo(32.666667, 5);
    });

    it('should return DIV0 error for no matching records', () => {
      const noMatchContext = createContext([
        ['Student', 'Subject', 'Score'],
        ['Alice', 'Math', 85],
        ['Subject'],
        ['Science']
      ]);
      
      const matches = ['DVARP(A1:C2, "Score", A3:A4)', 'A1:C2', 'Score', 'A3:A4'] as RegExpMatchArray;
      const result = DVARP.calculate(matches, noMatchContext);
      expect(result).toBe(FormulaError.DIV0);
    });

    it('should handle single data point', () => {
      const singlePointContext = createContext([
        ['Student', 'Subject', 'Score'],
        ['Alice', 'Math', 85],
        ['Subject'],
        ['Math']
      ]);
      
      const matches = ['DVARP(A1:C2, "Score", A3:A4)', 'A1:C2', 'Score', 'A3:A4'] as RegExpMatchArray;
      const result = DVARP.calculate(matches, singlePointContext);
      expect(result).toBe(0); // Population variance of single value is 0
    });

    it('should handle large numbers', () => {
      const largeNumberContext = createContext([
        ['Item', 'Category', 'Value'],
        ['A', 'Type1', 10000],
        ['B', 'Type1', 20000],
        ['C', 'Type1', 30000],
        ['Category'],
        ['Type1']
      ]);
      
      const matches = ['DVARP(A1:C4, "Value", A5:A6)', 'A1:C4', 'Value', 'A5:A6'] as RegExpMatchArray;
      const result = DVARP.calculate(matches, largeNumberContext);
      
      // Values: 10000, 20000, 30000
      // Mean = 20000
      // Population variance = [(10000-20000)² + (20000-20000)² + (30000-20000)²] / 3
      // = [100000000 + 0 + 100000000] / 3 = 66666666.67
      expect(result).toBeCloseTo(66666666.67, 2);
    });
  });

  describe('Edge Cases and Error Handling', () => {
    it('should handle field specified by number', () => {
      // Using field index 3 instead of field name "Score"
      const matches = ['DSTDEV(A1:D6, 3, F1:G2)', 'A1:D6', '3', 'F1:G2'] as RegExpMatchArray;
      const result = DSTDEV.calculate(matches, mockContext);
      expect(result).toBeCloseTo(7, 5);
    });

    it('should handle case-insensitive field matching', () => {
      const caseContext = createContext([
        ['student', 'subject', 'score'],
        ['Alice', 'Math', 85],
        ['Bob', 'Math', 92],
        ['SUBJECT'],
        ['Math']
      ]);
      
      const matches = ['DSTDEV(A1:C3, "SCORE", A4:A5)', 'A1:C3', 'SCORE', 'A4:A5'] as RegExpMatchArray;
      const result = DSTDEV.calculate(matches, caseContext);
      
      // Should find the "score" field despite case difference
      const expectedStdDev = Math.sqrt(Math.pow(85 - 88.5, 2) + Math.pow(92 - 88.5, 2));
      expect(result).toBeCloseTo(expectedStdDev, 5);
    });

    it('should handle empty criteria cells correctly', () => {
      const emptyCriteriaContext = createContext([
        ['Student', 'Subject', 'Score'],
        ['Alice', 'Math', 85],
        ['Bob', 'English', 92],
        ['Charlie', 'Math', 78],
        ['Subject'],
        [''] // Empty criteria - should match all
      ]);
      
      const matches = ['DVAR(A1:C4, "Score", A5:A6)', 'A1:C4', 'Score', 'A5:A6'] as RegExpMatchArray;
      const result = DVAR.calculate(matches, emptyCriteriaContext);
      
      // All scores: 85, 92, 78
      const mean = (85 + 92 + 78) / 3;
      const expectedVar = (Math.pow(85 - mean, 2) + Math.pow(92 - mean, 2) + Math.pow(78 - mean, 2)) / 2;
      expect(result).toBeCloseTo(expectedVar, 5);
    });

    it('should return VALUE error for malformed input', () => {
      const matches = ['DSTDEV(A1:D6, "Score", F1:G2)', 'A1:D6', 'Score', 'F1:G2'] as RegExpMatchArray;
      const malformedContext = createContext([]);
      const result = DSTDEV.calculate(matches, malformedContext);
      expect(result).toBe(FormulaError.VALUE);
    });

    it('should handle criteria with inequality operators', () => {
      const inequalityContext = createContext([
        ['Student', 'Subject', 'Score', 'Grade'],
        ['Alice', 'Math', 85, 'B'],
        ['Bob', 'Math', 92, 'A'],
        ['Charlie', 'Math', 78, 'C'],
        ['David', 'Math', 95, 'A'],
        ['Score'],
        ['>=85'] // Scores 85 and above
      ]);
      
      const matches = ['DVARP(A1:D5, "Score", A6:A7)', 'A1:D5', 'Score', 'A6:A7'] as RegExpMatchArray;
      const result = DVARP.calculate(matches, inequalityContext);
      
      // Matching scores: 85, 92, 95
      const values = [85, 92, 95];
      const mean = values.reduce((sum, val) => sum + val, 0) / values.length;
      const expectedVar = values.reduce((sum, val) => sum + Math.pow(val - mean, 2), 0) / values.length;
      
      expect(result).toBeCloseTo(expectedVar, 5);
    });
  });
});