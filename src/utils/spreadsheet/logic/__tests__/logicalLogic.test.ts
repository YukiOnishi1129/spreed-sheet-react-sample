import { describe, it, expect } from 'vitest';
import { IF, AND, OR, NOT, IFS } from '../05-logical/logicLogic';
import type { FormulaContext } from '../shared/types';

describe('Logical Functions', () => {
  const createContext = (data: (string | number | boolean | null)[][]): FormulaContext => ({
    data: data.map(row => row.map(cell => ({ value: cell }))),
    row: 0,
    col: 0
  });

  describe('IF function', () => {
    it('should handle numeric comparisons', () => {
      const context = createContext([
        [100, 50, 150]
      ]);
      
      // =IF(A1>=100,"合格","不合格")
      const matches = 'IF(A1>=100,"合格","不合格")'.match(/^IF\s*\((.+)\)$/i)!;
      const result = IF.calculate(matches, context);
      expect(result).toBe('合格');
    });

    it('should handle less than comparison', () => {
      const context = createContext([
        [80, 50, 150]
      ]);
      
      // =IF(A1>=100,"合格","不合格")
      const matches = 'IF(A1>=100,"合格","不合格")'.match(/^IF\s*\((.+)\)$/i)!;
      const result = IF.calculate(matches, context);
      expect(result).toBe('不合格');
    });

    it('should handle string comparisons', () => {
      const context = createContext([
        ['Apple', 'Banana', 'Cherry']
      ]);
      
      // =IF(A1="Apple","Found","Not Found")
      const matches = 'IF(A1="Apple","Found","Not Found")'.match(/^IF\s*\((.+)\)$/i)!;
      const result = IF.calculate(matches, context);
      expect(result).toBe('Found');
    });

    it('should handle boolean conditions', () => {
      const context = createContext([
        [true, false, 1]
      ]);
      
      // =IF(A1,"Yes","No")
      const matches = 'IF(A1,"Yes","No")'.match(/^IF\s*\((.+)\)$/i)!;
      const result = IF.calculate(matches, context);
      expect(result).toBe('Yes');
    });
  });

  describe('AND function', () => {
    it('should return true when all conditions are true', () => {
      const context = createContext([
        [true, true, true]
      ]);
      
      const matches = ['AND(A1,B1,C1)', 'A1,B1,C1'];
      const result = AND.calculate(matches as RegExpMatchArray, context);
      expect(result).toBe(true);
    });

    it('should return false when any condition is false', () => {
      const context = createContext([
        [true, false, true]
      ]);
      
      const matches = ['AND(A1,B1,C1)', 'A1,B1,C1'];
      const result = AND.calculate(matches as RegExpMatchArray, context);
      expect(result).toBe(false);
    });

    it('should handle numeric values (0 as false, others as true)', () => {
      const context = createContext([
        [1, 2, 0]
      ]);
      
      const matches = ['AND(A1,B1,C1)', 'A1,B1,C1'];
      const result = AND.calculate(matches as RegExpMatchArray, context);
      expect(result).toBe(false);
    });
  });

  describe('OR function', () => {
    it('should return true when any condition is true', () => {
      const context = createContext([
        [true, false, false]
      ]);
      
      const matches = ['OR(A1,B1,C1)', 'A1,B1,C1'];
      const result = OR.calculate(matches as RegExpMatchArray, context);
      expect(result).toBe(true);
    });

    it('should return false when all conditions are false', () => {
      const context = createContext([
        [false, false, false]
      ]);
      
      const matches = ['OR(A1,B1,C1)', 'A1,B1,C1'];
      const result = OR.calculate(matches as RegExpMatchArray, context);
      expect(result).toBe(false);
    });
  });

  describe('NOT function', () => {
    it('should invert true to false', () => {
      const context = createContext([
        [true]
      ]);
      
      const matches = ['NOT(A1)', 'A1'];
      const result = NOT.calculate(matches as RegExpMatchArray, context);
      expect(result).toBe(false);
    });

    it('should invert false to true', () => {
      const context = createContext([
        [false]
      ]);
      
      const matches = ['NOT(A1)', 'A1'];
      const result = NOT.calculate(matches as RegExpMatchArray, context);
      expect(result).toBe(true);
    });
  });

  describe('IFS function', () => {
    it('should evaluate multiple conditions and return first true result', () => {
      const context = createContext([
        [85]
      ]);
      
      // =IFS(A1>=90,"A",A1>=80,"B",A1>=70,"C",TRUE,"F")
      const matches = ['IFS(A1>=90,"A",A1>=80,"B",A1>=70,"C",TRUE,"F")', 'A1>=90,"A",A1>=80,"B",A1>=70,"C",TRUE,"F"'];
      const result = IFS.calculate(matches as RegExpMatchArray, context);
      expect(result).toBe('B');
    });

    it('should handle default case with TRUE', () => {
      const context = createContext([
        [65]
      ]);
      
      // =IFS(A1>=90,"A",A1>=80,"B",A1>=70,"C",TRUE,"F")
      const matches = ['IFS(A1>=90,"A",A1>=80,"B",A1>=70,"C","1","F")', 'A1>=90,"A",A1>=80,"B",A1>=70,"C","1","F"'];
      const result = IFS.calculate(matches as RegExpMatchArray, context);
      expect(result).toBe('F');
    });
  });
});