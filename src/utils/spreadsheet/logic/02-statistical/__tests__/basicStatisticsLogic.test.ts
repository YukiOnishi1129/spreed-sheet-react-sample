import { describe, it, expect } from 'vitest';
import type { FormulaContext } from '../../shared/types';
import {
  AVERAGE, COUNT, COUNTA, COUNTBLANK,
  MAX, MIN, MEDIAN, MODE,
  STDEV_S, STDEV_P, VAR_S, VAR_P,
  AVEDEV, GEOMEAN, HARMEAN,
  LARGE, SMALL, RANK, TRIMMEAN,
  STDEVA, STDEVPA, VARA, VARPA
} from '../basicStatisticsLogic';
import { FormulaError } from '../../shared/types';

// Helper function to create FormulaContext
function createContext(data: (string | number | boolean | null)[][]): FormulaContext {
  return {
    data: data.map(row => row.map(cell => ({ value: cell }))),
    row: 0,
    col: 0
  };
}

describe('Basic Statistics Functions', () => {
  const mockContext = createContext([
    [1, 'text', 10, 2, -5],  // Row 1 (A1-E1)
    [2, '', 20, 3, -2],      // Row 2 (A2-E2)
    [3, null, 30, 2, 0],     // Row 3 (A3-E3)
    [4, 0, 20, 2, 3],        // Row 4 (A4-E4)
    [5, false, 10, 5, 7]     // Row 5 (A5-E5)
  ]);

  describe('Basic Aggregation Functions', () => {
    describe('AVERAGE', () => {
      it('should calculate average of numbers', () => {
        const matches = ['AVERAGE(A1:A5)', 'A1:A5'] as RegExpMatchArray;
        expect(AVERAGE.calculate(matches, mockContext)).toBe(3);
      });

      it('should handle multiple ranges', () => {
        const matches = ['AVERAGE(A1:A3, C1:C2)', 'A1:A3, C1:C2'] as RegExpMatchArray;
        expect(AVERAGE.calculate(matches, mockContext)).toBe(9.2); // (1+2+3+10+20)/5
      });

      it('should ignore non-numeric values', () => {
        const matches = ['AVERAGE(A1:A3, B1:B3)', 'A1:A3, B1:B3'] as RegExpMatchArray;
        expect(AVERAGE.calculate(matches, mockContext)).toBe(2); // (1+2+3)/3
      });

      it('should return DIV0 error for no numeric values', () => {
        const matches = ['AVERAGE(B1:B3)', 'B1:B3'] as RegExpMatchArray;
        expect(AVERAGE.calculate(matches, mockContext)).toBe(FormulaError.DIV0);
      });
    });

    describe('COUNT', () => {
      it('should count numeric cells', () => {
        const matches = ['COUNT(A1:A5)', 'A1:A5'] as RegExpMatchArray;
        expect(COUNT.calculate(matches, mockContext)).toBe(5);
      });

      it('should count only numeric values', () => {
        const matches = ['COUNT(A1:B5)', 'A1:B5'] as RegExpMatchArray;
        expect(COUNT.calculate(matches, mockContext)).toBe(6); // 5 from A + 1 from B4
      });

      it('should handle multiple ranges', () => {
        const matches = ['COUNT(A1:A3, C1:C3)', 'A1:A3, C1:C3'] as RegExpMatchArray;
        expect(COUNT.calculate(matches, mockContext)).toBe(6);
      });

      it('should return 0 for no numeric values', () => {
        const matches = ['COUNT(B1:B3)', 'B1:B3'] as RegExpMatchArray;
        expect(COUNT.calculate(matches, mockContext)).toBe(0);
      });
    });

    describe('COUNTA', () => {
      it('should count non-empty cells', () => {
        const matches = ['COUNTA(A1:A5)', 'A1:A5'] as RegExpMatchArray;
        expect(COUNTA.calculate(matches, mockContext)).toBe(5);
      });

      it('should count cells with any value', () => {
        const matches = ['COUNTA(B1:B5)', 'B1:B5'] as RegExpMatchArray;
        expect(COUNTA.calculate(matches, mockContext)).toBe(3); // text, 0, false
      });

      it('should handle multiple ranges', () => {
        const matches = ['COUNTA(A1:A2, B1:B2)', 'A1:A2, B1:B2'] as RegExpMatchArray;
        expect(COUNTA.calculate(matches, mockContext)).toBe(3); // 1, 2, text
      });
    });

    describe('COUNTBLANK', () => {
      it('should count empty cells', () => {
        const matches = ['COUNTBLANK(B1:B5)', 'B1:B5'] as RegExpMatchArray;
        expect(COUNTBLANK.calculate(matches, mockContext)).toBe(2); // empty string and null
      });

      it('should return 0 for range with no empty cells', () => {
        const matches = ['COUNTBLANK(A1:A5)', 'A1:A5'] as RegExpMatchArray;
        expect(COUNTBLANK.calculate(matches, mockContext)).toBe(0);
      });
    });
  });

  describe('Min/Max Functions', () => {
    describe('MAX', () => {
      it('should find maximum value', () => {
        const matches = ['MAX(A1:A5)', 'A1:A5'] as RegExpMatchArray;
        expect(MAX.calculate(matches, mockContext)).toBe(5);
      });

      it('should handle multiple ranges', () => {
        const matches = ['MAX(A1:A5, C1:C5)', 'A1:A5, C1:C5'] as RegExpMatchArray;
        expect(MAX.calculate(matches, mockContext)).toBe(30);
      });

      it('should return 0 for no numeric values', () => {
        const matches = ['MAX(B1:B3)', 'B1:B3'] as RegExpMatchArray;
        expect(MAX.calculate(matches, mockContext)).toBe(0);
      });
    });

    describe('MIN', () => {
      it('should find minimum value', () => {
        const matches = ['MIN(A1:A5)', 'A1:A5'] as RegExpMatchArray;
        expect(MIN.calculate(matches, mockContext)).toBe(1);
      });

      it('should handle negative values', () => {
        const matches = ['MIN(E1:E5)', 'E1:E5'] as RegExpMatchArray;
        expect(MIN.calculate(matches, mockContext)).toBe(-5);
      });

      it('should return 0 for no numeric values', () => {
        const matches = ['MIN(B1:B3)', 'B1:B3'] as RegExpMatchArray;
        expect(MIN.calculate(matches, mockContext)).toBe(0);
      });
    });
  });

  describe('Central Tendency Functions', () => {
    describe('MEDIAN', () => {
      it('should calculate median for odd count', () => {
        const matches = ['MEDIAN(A1:A5)', 'A1:A5'] as RegExpMatchArray;
        expect(MEDIAN.calculate(matches, mockContext)).toBe(3);
      });

      it('should calculate median for even count', () => {
        const matches = ['MEDIAN(A1:A4)', 'A1:A4'] as RegExpMatchArray;
        expect(MEDIAN.calculate(matches, mockContext)).toBe(2.5);
      });

      it('should return NUM error for no numeric values', () => {
        const matches = ['MEDIAN(B1:B3)', 'B1:B3'] as RegExpMatchArray;
        expect(MEDIAN.calculate(matches, mockContext)).toBe(FormulaError.NUM);
      });
    });

    describe('MODE', () => {
      it('should find most frequent value', () => {
        const matches = ['MODE(D1:D5)', 'D1:D5'] as RegExpMatchArray;
        expect(MODE.calculate(matches, mockContext)).toBe(2); // appears 3 times
      });

      it('should return first mode when multiple exist', () => {
        const matches = ['MODE(C1:C5)', 'C1:C5'] as RegExpMatchArray;
        expect(MODE.calculate(matches, mockContext)).toBe(10); // both 10 and 20 appear twice
      });

      it('should return NA error when no mode exists', () => {
        const matches = ['MODE(A1:A5)', 'A1:A5'] as RegExpMatchArray;
        expect(MODE.calculate(matches, mockContext)).toBe(FormulaError.NA);
      });
    });
  });

  describe('Variance and Standard Deviation Functions', () => {
    describe('STDEV.S', () => {
      it('should calculate sample standard deviation', () => {
        const matches = ['STDEV.S(A1:A5)', 'A1:A5'] as RegExpMatchArray;
        expect(STDEV_S.calculate(matches, mockContext)).toBeCloseTo(1.581, 3);
      });

      it('should return DIV0 error for single value', () => {
        const matches = ['STDEV.S(A1)', 'A1'] as RegExpMatchArray;
        expect(STDEV_S.calculate(matches, mockContext)).toBe(FormulaError.DIV0);
      });
    });

    describe('STDEV.P', () => {
      it('should calculate population standard deviation', () => {
        const matches = ['STDEV.P(A1:A5)', 'A1:A5'] as RegExpMatchArray;
        expect(STDEV_P.calculate(matches, mockContext)).toBeCloseTo(1.414, 3);
      });

      it('should return 0 for single value', () => {
        const matches = ['STDEV.P(A1)', 'A1'] as RegExpMatchArray;
        expect(STDEV_P.calculate(matches, mockContext)).toBe(0);
      });
    });

    describe('VAR.S', () => {
      it('should calculate sample variance', () => {
        const matches = ['VAR.S(A1:A5)', 'A1:A5'] as RegExpMatchArray;
        expect(VAR_S.calculate(matches, mockContext)).toBe(2.5);
      });
    });

    describe('VAR.P', () => {
      it('should calculate population variance', () => {
        const matches = ['VAR.P(A1:A5)', 'A1:A5'] as RegExpMatchArray;
        expect(VAR_P.calculate(matches, mockContext)).toBe(2);
      });
    });
  });

  describe('Mean Deviation Functions', () => {
    describe('AVEDEV', () => {
      it('should calculate average absolute deviation', () => {
        const matches = ['AVEDEV(A1:A5)', 'A1:A5'] as RegExpMatchArray;
        expect(AVEDEV.calculate(matches, mockContext)).toBe(1.2);
      });

      it('should return 0 for single value', () => {
        const matches = ['AVEDEV(A1)', 'A1'] as RegExpMatchArray;
        expect(AVEDEV.calculate(matches, mockContext)).toBe(0);
      });
    });
  });

  describe('Special Mean Functions', () => {
    describe('GEOMEAN', () => {
      it('should calculate geometric mean', () => {
        const matches = ['GEOMEAN(A1:A5)', 'A1:A5'] as RegExpMatchArray;
        expect(GEOMEAN.calculate(matches, mockContext)).toBeCloseTo(2.605, 3);
      });

      it('should return NUM error for non-positive values', () => {
        const matches = ['GEOMEAN(E1:E5)', 'E1:E5'] as RegExpMatchArray;
        expect(GEOMEAN.calculate(matches, mockContext)).toBe(FormulaError.NUM);
      });
    });

    describe('HARMEAN', () => {
      it('should calculate harmonic mean', () => {
        const matches = ['HARMEAN(A1:A5)', 'A1:A5'] as RegExpMatchArray;
        expect(HARMEAN.calculate(matches, mockContext)).toBeCloseTo(2.19, 2);
      });

      it('should return NUM error for non-positive values', () => {
        const matches = ['HARMEAN(E1:E5)', 'E1:E5'] as RegExpMatchArray;
        expect(HARMEAN.calculate(matches, mockContext)).toBe(FormulaError.NUM);
      });
    });

    describe('TRIMMEAN', () => {
      it('should calculate trimmed mean', () => {
        const matches = ['TRIMMEAN(A1:A5, 0.2)', 'A1:A5', '0.2'] as RegExpMatchArray;
        expect(TRIMMEAN.calculate(matches, mockContext)).toBe(3); // Trim 20% from each end
      });

      it('should return normal mean for 0 percent', () => {
        const matches = ['TRIMMEAN(A1:A5, 0)', 'A1:A5', '0'] as RegExpMatchArray;
        expect(TRIMMEAN.calculate(matches, mockContext)).toBe(3);
      });

      it('should return NUM error for percent >= 1', () => {
        const matches = ['TRIMMEAN(A1:A5, 1)', 'A1:A5', '1'] as RegExpMatchArray;
        expect(TRIMMEAN.calculate(matches, mockContext)).toBe(FormulaError.NUM);
      });
    });
  });

  describe('Position Functions', () => {
    describe('LARGE', () => {
      it('should find k-th largest value', () => {
        const matches = ['LARGE(A1:A5, 2)', 'A1:A5', '2'] as RegExpMatchArray;
        expect(LARGE.calculate(matches, mockContext)).toBe(4);
      });

      it('should return largest for k=1', () => {
        const matches = ['LARGE(A1:A5, 1)', 'A1:A5', '1'] as RegExpMatchArray;
        expect(LARGE.calculate(matches, mockContext)).toBe(5);
      });

      it('should return NUM error for k > count', () => {
        const matches = ['LARGE(A1:A5, 6)', 'A1:A5', '6'] as RegExpMatchArray;
        expect(LARGE.calculate(matches, mockContext)).toBe(FormulaError.NUM);
      });
    });

    describe('SMALL', () => {
      it('should find k-th smallest value', () => {
        const matches = ['SMALL(A1:A5, 2)', 'A1:A5', '2'] as RegExpMatchArray;
        expect(SMALL.calculate(matches, mockContext)).toBe(2);
      });

      it('should return smallest for k=1', () => {
        const matches = ['SMALL(A1:A5, 1)', 'A1:A5', '1'] as RegExpMatchArray;
        expect(SMALL.calculate(matches, mockContext)).toBe(1);
      });
    });

    describe('RANK', () => {
      it('should rank value in descending order by default', () => {
        const matches = ['RANK(3, A1:A5)', '3', 'A1:A5'] as RegExpMatchArray;
        expect(RANK.calculate(matches, mockContext)).toBe(3); // 5,4,3,2,1
      });

      it('should rank in ascending order when specified', () => {
        const matches = ['RANK(3, A1:A5, 1)', '3', 'A1:A5', '1'] as RegExpMatchArray;
        expect(RANK.calculate(matches, mockContext)).toBe(3); // 1,2,3,4,5
      });

      it('should return NA error for value not in list', () => {
        const matches = ['RANK(10, A1:A5)', '10', 'A1:A5'] as RegExpMatchArray;
        expect(RANK.calculate(matches, mockContext)).toBe(FormulaError.NA);
      });
    });
  });

  describe('Extended Variance Functions', () => {
    describe('STDEVA', () => {
      it('should include text as 0 in calculation', () => {
        const matches = ['STDEVA(A1:A3, B1)', 'A1:A3, B1'] as RegExpMatchArray;
        const result = STDEVA.calculate(matches, mockContext);
        expect(result).toBeGreaterThan(0); // Should treat text as 0
      });
    });

    describe('STDEVPA', () => {
      it('should include text as 0 in population calculation', () => {
        const matches = ['STDEVPA(A1:A3, B1)', 'A1:A3, B1'] as RegExpMatchArray;
        const result = STDEVPA.calculate(matches, mockContext);
        expect(result).toBeGreaterThan(0);
      });
    });

    describe('VARA', () => {
      it('should include text as 0 in variance calculation', () => {
        const matches = ['VARA(A1:A3, B1)', 'A1:A3, B1'] as RegExpMatchArray;
        const result = VARA.calculate(matches, mockContext);
        expect(result).toBeGreaterThan(0);
      });
    });

    describe('VARPA', () => {
      it('should include text as 0 in population variance', () => {
        const matches = ['VARPA(A1:A3, B1)', 'A1:A3, B1'] as RegExpMatchArray;
        const result = VARPA.calculate(matches, mockContext);
        expect(result).toBeGreaterThan(0);
      });
    });
  });
});