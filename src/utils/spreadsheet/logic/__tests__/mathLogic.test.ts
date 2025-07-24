import { describe, it, expect } from 'vitest';
import { SUM, AVERAGE, COUNT, MAX, MIN, ROUND, ROUNDUP, ROUNDDOWN } from '../01-math-trigonometry/mathLogic';
import type { FormulaContext } from '../shared/types';

describe('Math Functions', () => {
  const createContext = (data: (string | number | null)[][]): FormulaContext => ({
    data: data.map(row => row.map(cell => ({ value: cell }))),
    row: 0,
    col: 0
  });

  describe('SUM function', () => {
    it('should sum single range', () => {
      const context = createContext([
        [10, 20, 30],
        [40, 50, 60]
      ]);
      
      const matches = ['SUM(A1:C1)', 'A1:C1'];
      const result = SUM.calculate(matches as RegExpMatchArray, context);
      expect(result).toBe(60); // 10 + 20 + 30
    });

    it('should sum multiple ranges', () => {
      const context = createContext([
        [10, 20, 30],
        [40, 50, 60]
      ]);
      
      const matches = ['SUM(A1:B1,C2)', 'A1:B1,C2'];
      const result = SUM.calculate(matches as RegExpMatchArray, context);
      expect(result).toBe(90); // 10 + 20 + 60
    });

    it('should handle empty cells', () => {
      const context = createContext([
        [10, null, 30],
        [40, '', 60]
      ]);
      
      const matches = ['SUM(A1:C1)', 'A1:C1'];
      const result = SUM.calculate(matches as RegExpMatchArray, context);
      expect(result).toBe(40); // 10 + 0 + 30
    });
  });

  describe('AVERAGE function', () => {
    it('should calculate average correctly', () => {
      const context = createContext([
        [10, 20, 30],
        [40, 50, 60]
      ]);
      
      const matches = ['AVERAGE(A1:C1)', 'A1:C1'];
      const result = AVERAGE.calculate(matches as RegExpMatchArray, context);
      expect(result).toBe(20); // (10 + 20 + 30) / 3
    });

    it('should skip non-numeric values', () => {
      const context = createContext([
        [10, 'text', 30],
        [40, null, 60]
      ]);
      
      const matches = ['AVERAGE(A1:C1)', 'A1:C1'];
      const result = AVERAGE.calculate(matches as RegExpMatchArray, context);
      expect(result).toBe(20); // (10 + 30) / 2
    });
  });

  describe('COUNT function', () => {
    it('should count numeric values', () => {
      const context = createContext([
        [10, 'text', 30],
        [40, null, 60]
      ]);
      
      const matches = ['COUNT(A1:C2)', 'A1:C2'];
      const result = COUNT.calculate(matches as RegExpMatchArray, context);
      expect(result).toBe(4); // 10, 30, 40, 60
    });
  });

  describe('MAX/MIN functions', () => {
    it('MAX should return maximum value', () => {
      const context = createContext([
        [10, 50, 30],
        [40, 20, 60]
      ]);
      
      const matches = ['MAX(A1:C2)', 'A1:C2'];
      const result = MAX.calculate(matches as RegExpMatchArray, context);
      expect(result).toBe(60);
    });

    it('MIN should return minimum value', () => {
      const context = createContext([
        [10, 50, 30],
        [40, 20, 60]
      ]);
      
      const matches = ['MIN(A1:C2)', 'A1:C2'];
      const result = MIN.calculate(matches as RegExpMatchArray, context);
      expect(result).toBe(10);
    });
  });

  describe('ROUND functions', () => {
    it('ROUND should round to specified digits', () => {
      const context = createContext([[3.14159]]);
      
      const matches = ['ROUND(3.14159,2)', '3.14159', '2'];
      const result = ROUND.calculate(matches as RegExpMatchArray, context);
      expect(result).toBe(3.14);
    });

    it('ROUNDUP should round up', () => {
      const context = createContext([[3.141]]);
      
      const matches = ['ROUNDUP(3.141,2)', '3.141', '2'];
      const result = ROUNDUP.calculate(matches as RegExpMatchArray, context);
      expect(result).toBe(3.15);
    });

    it('ROUNDDOWN should round down', () => {
      const context = createContext([[3.149]]);
      
      const matches = ['ROUNDDOWN(3.149,2)', '3.149', '2'];
      const result = ROUNDDOWN.calculate(matches as RegExpMatchArray, context);
      expect(result).toBe(3.14);
    });
  });
});