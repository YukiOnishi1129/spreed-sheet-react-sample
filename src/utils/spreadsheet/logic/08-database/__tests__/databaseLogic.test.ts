import { describe, it, expect } from 'vitest';
import {
  DSUM, DAVERAGE, DCOUNT, DCOUNTA, DMAX, DMIN, DPRODUCT, DGET
} from '../databaseLogic';
import { FormulaError } from '../../shared/types';
import type { FormulaContext } from '../../shared/types';

// Helper function to create FormulaContext
const createContext = (data: (string | number | boolean | null)[][]): FormulaContext => ({
  data: data.map(row => row.map(cell => ({ value: cell }))),
  row: 0,
  col: 0
});

describe('Database Functions', () => {
  // Sample database: Employee data
  const mockContext = createContext([
    // Headers row (A1:D1)
    ['Name', 'Department', 'Salary', 'Age'],
    // Data rows (A2:D6)
    ['Alice', 'Sales', 50000, 25],
    ['Bob', 'Engineering', 75000, 30],
    ['Charlie', 'Sales', 45000, 28],
    ['David', 'Engineering', 80000, 35],
    ['Eve', 'Marketing', 55000, 26],
    // Criteria area (F1:H3)
    ['', '', 'Department', 'Salary'],
    ['', '', 'Sales', '>40000'],
    ['', '', 'Engineering', '']
  ]);

  describe('DSUM Function', () => {
    it('should sum salaries for Sales department', () => {
      const matches = ['DSUM(A1:D6, "Salary", F1:G2)', 'A1:D6', '"Salary"', 'F1:G2'] as RegExpMatchArray;
      const result = DSUM.calculate(matches, mockContext);
      expect(result).toBe(95000); // Alice(50000) + Charlie(45000)
    });

    it('should sum using field index instead of name', () => {
      const matches = ['DSUM(A1:D6, 3, F1:G2)', 'A1:D6', '3', 'F1:G2'] as RegExpMatchArray;
      const result = DSUM.calculate(matches, mockContext);
      expect(result).toBe(95000); // Same as above, using column index 3 for Salary
    });

    it('should return 0 when no matching records', () => {
      const noMatchContext = createContext([
        ['Name', 'Department', 'Salary'],
        ['Alice', 'Sales', 50000],
        ['Bob', 'Engineering', 75000],
        ['Department'],
        ['HR'] // No HR department in data
      ]);
      
      const matches = ['DSUM(A1:C3, "Salary", A4:A5)', 'A1:C3', '"Salary"', 'A4:A5'] as RegExpMatchArray;
      const result = DSUM.calculate(matches, noMatchContext);
      expect(result).toBe(0);
    });

    it('should return VALUE error for invalid field', () => {
      const matches = ['DSUM(A1:D6, "InvalidField", F1:G2)', 'A1:D6', '"InvalidField"', 'F1:G2'] as RegExpMatchArray;
      const result = DSUM.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });

    it('should handle complex criteria with comparison operators', () => {
      const matches = ['DSUM(A1:D6, "Salary", G1:H2)', 'A1:D6', '"Salary"', 'G1:H2'] as RegExpMatchArray;
      const result = DSUM.calculate(matches, mockContext);
      expect(result).toBe(50000); // Only Alice meets >40000 in Sales
    });
  });

  describe('DAVERAGE Function', () => {
    it('should calculate average salary for Sales department', () => {
      const matches = ['DAVERAGE(A1:D6, "Salary", F1:G2)', 'A1:D6', '"Salary"', 'F1:G2'] as RegExpMatchArray;
      const result = DAVERAGE.calculate(matches, mockContext);
      expect(result).toBe(47500); // (50000 + 45000) / 2
    });

    it('should return DIV0 error when no matching records', () => {
      const noMatchContext = createContext([
        ['Name', 'Department', 'Salary'],
        ['Alice', 'Sales', 50000],
        ['Department'],
        ['HR'] // No HR department in data
      ]);
      
      const matches = ['DAVERAGE(A1:C2, "Salary", A3:A4)', 'A1:C2', '"Salary"', 'A3:A4'] as RegExpMatchArray;
      const result = DAVERAGE.calculate(matches, noMatchContext);
      expect(result).toBe(FormulaError.DIV0);
    });

    it('should handle non-numeric values gracefully', () => {
      const mixedContext = createContext([
        ['Name', 'Value'],
        ['Alice', 100],
        ['Bob', 'N/A'],
        ['Charlie', 200],
        ['Name'],
        ['Alice']
      ]);
      
      const matches = ['DAVERAGE(A1:B4, "Value", A5:A6)', 'A1:B4', '"Value"', 'A5:A6'] as RegExpMatchArray;
      const result = DAVERAGE.calculate(matches, mixedContext);
      expect(result).toBe(100); // Only Alice's 100 is considered
    });
  });

  describe('DCOUNT Function', () => {
    it('should count numeric values in Sales department', () => {
      const matches = ['DCOUNT(A1:D6, "Salary", F1:G2)', 'A1:D6', '"Salary"', 'F1:G2'] as RegExpMatchArray;
      const result = DCOUNT.calculate(matches, mockContext);
      expect(result).toBe(2); // Alice and Charlie
    });

    it('should return 0 when no matching numeric records', () => {
      const textContext = createContext([
        ['Name', 'Department', 'Description'],
        ['Alice', 'Sales', 'Manager'],
        ['Bob', 'Sales', 'Associate'],
        ['Department'],
        ['Sales']
      ]);
      
      const matches = ['DCOUNT(A1:C3, "Description", A4:A5)', 'A1:C3', '"Description"', 'A4:A5'] as RegExpMatchArray;
      const result = DCOUNT.calculate(matches, textContext);
      expect(result).toBe(0); // Description field has no numeric values
    });

    it('should handle age field counting', () => {
      const matches = ['DCOUNT(A1:D6, "Age", F1:G2)', 'A1:D6', '"Age"', 'F1:G2'] as RegExpMatchArray;
      const result = DCOUNT.calculate(matches, mockContext);
      expect(result).toBe(2); // Alice(25) and Charlie(28)
    });
  });

  describe('DCOUNTA Function', () => {
    it('should count non-empty values in Sales department', () => {
      const matches = ['DCOUNTA(A1:D6, "Name", F1:G2)', 'A1:D6', '"Name"', 'F1:G2'] as RegExpMatchArray;
      const result = DCOUNTA.calculate(matches, mockContext);
      expect(result).toBe(2); // Alice and Charlie
    });

    it('should not count empty or null values', () => {
      const emptyContext = createContext([
        ['Name', 'Department', 'Bonus'],
        ['Alice', 'Sales', 1000],
        ['Bob', 'Sales', null],
        ['Charlie', 'Sales', ''],
        ['Department'],
        ['Sales']
      ]);
      
      const matches = ['DCOUNTA(A1:C4, "Bonus", A5:A6)', 'A1:C4', '"Bonus"', 'A5:A6'] as RegExpMatchArray;
      const result = DCOUNTA.calculate(matches, emptyContext);
      expect(result).toBe(1); // Only Alice has non-empty bonus
    });

    it('should count text values', () => {
      const matches = ['DCOUNTA(A1:D6, "Department", F1:G2)', 'A1:D6', '"Department"', 'F1:G2'] as RegExpMatchArray;
      const result = DCOUNTA.calculate(matches, mockContext);
      expect(result).toBe(2); // Alice and Charlie both have "Sales"
    });
  });

  describe('DMAX Function', () => {
    it('should find maximum salary in Sales department', () => {
      const matches = ['DMAX(A1:D6, "Salary", F1:G2)', 'A1:D6', '"Salary"', 'F1:G2'] as RegExpMatchArray;
      const result = DMAX.calculate(matches, mockContext);
      expect(result).toBe(50000); // Alice's salary
    });

    it('should return 0 when no matching records', () => {
      const noMatchContext = createContext([
        ['Name', 'Department', 'Salary'],
        ['Alice', 'Sales', 50000],
        ['Department'],
        ['HR']
      ]);
      
      const matches = ['DMAX(A1:C2, "Salary", A3:A4)', 'A1:C2', '"Salary"', 'A3:A4'] as RegExpMatchArray;
      const result = DMAX.calculate(matches, noMatchContext);
      expect(result).toBe(0);
    });

    it('should handle negative numbers', () => {
      const negativeContext = createContext([
        ['Name', 'Loss'],
        ['Alice', -100],
        ['Bob', -50],
        ['Name'],
        ['Alice']
      ]);
      
      const matches = ['DMAX(A1:B3, "Loss", A4:A5)', 'A1:B3', '"Loss"', 'A4:A5'] as RegExpMatchArray;
      const result = DMAX.calculate(matches, negativeContext);
      expect(result).toBe(-100); // Alice's loss
    });
  });

  describe('DMIN Function', () => {
    it('should find minimum salary in Sales department', () => {
      const matches = ['DMIN(A1:D6, "Salary", F1:G2)', 'A1:D6', '"Salary"', 'F1:G2'] as RegExpMatchArray;
      const result = DMIN.calculate(matches, mockContext);
      expect(result).toBe(45000); // Charlie's salary
    });

    it('should return 0 when no matching records', () => {
      const noMatchContext = createContext([
        ['Name', 'Department', 'Salary'],
        ['Alice', 'Sales', 50000],
        ['Department'],
        ['HR']
      ]);
      
      const matches = ['DMIN(A1:C2, "Salary", A3:A4)', 'A1:C2', '"Salary"', 'A3:A4'] as RegExpMatchArray;
      const result = DMIN.calculate(matches, noMatchContext);
      expect(result).toBe(0);
    });

    it('should find minimum age in Engineering department', () => {
      const matches = ['DMIN(A1:D6, "Age", G1:H3)', 'A1:D6', '"Age"', 'G1:H3'] as RegExpMatchArray;
      const result = DMIN.calculate(matches, mockContext);
      expect(result).toBe(30); // Bob's age
    });
  });

  describe('DPRODUCT Function', () => {
    it('should calculate product of salaries in Sales department', () => {
      const matches = ['DPRODUCT(A1:D6, "Salary", F1:G2)', 'A1:D6', '"Salary"', 'F1:G2'] as RegExpMatchArray;
      const result = DPRODUCT.calculate(matches, mockContext);
      expect(result).toBe(2250000000); // 50000 * 45000
    });

    it('should return 0 when no matching records', () => {
      const noMatchContext = createContext([
        ['Name', 'Department', 'Salary'],
        ['Alice', 'Sales', 50000],
        ['Department'],
        ['HR']
      ]);
      
      const matches = ['DPRODUCT(A1:C2, "Salary", A3:A4)', 'A1:C2', '"Salary"', 'A3:A4'] as RegExpMatchArray;
      const result = DPRODUCT.calculate(matches, noMatchContext);
      expect(result).toBe(0);
    });

    it('should handle single matching record', () => {
      const singleContext = createContext([
        ['Name', 'Department', 'Value'],
        ['Alice', 'Sales', 5],
        ['Bob', 'Engineering', 3],
        ['Department'],
        ['Sales']
      ]);
      
      const matches = ['DPRODUCT(A1:C3, "Value", A4:A5)', 'A1:C3', '"Value"', 'A4:A5'] as RegExpMatchArray;
      const result = DPRODUCT.calculate(matches, singleContext);
      expect(result).toBe(5); // Only Alice matches
    });
  });

  describe('DGET Function', () => {
    it('should get single matching value', () => {
      const uniqueContext = createContext([
        ['Name', 'Department', 'Salary'],
        ['Alice', 'Sales', 50000],
        ['Bob', 'Engineering', 75000],
        ['Name'],
        ['Alice']
      ]);
      
      const matches = ['DGET(A1:C3, "Salary", A4:A5)', 'A1:C3', '"Salary"', 'A4:A5'] as RegExpMatchArray;
      const result = DGET.calculate(matches, uniqueContext);
      expect(result).toBe(50000);
    });

    it('should return VALUE error when no records match', () => {
      const matches = ['DGET(A1:D6, "Salary", F1:G2)', 'A1:D6', '"Salary"', 'F1:G2'] as RegExpMatchArray;
      const result = DGET.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM); // Multiple records match Sales department
    });

    it('should return NUM error when multiple records match', () => {
      const noMatchContext = createContext([
        ['Name', 'Department', 'Salary'],
        ['Alice', 'Sales', 50000],
        ['Department'],
        ['HR']
      ]);
      
      const matches = ['DGET(A1:C2, "Salary", A3:A4)', 'A1:C2', '"Salary"', 'A3:A4'] as RegExpMatchArray;
      const result = DGET.calculate(matches, noMatchContext);
      expect(result).toBe(FormulaError.VALUE);
    });

    it('should get text value', () => {
      const textContext = createContext([
        ['ID', 'Name', 'Status'],
        [1, 'Alice', 'Active'],
        [2, 'Bob', 'Inactive'],
        ['ID'],
        [1]
      ]);
      
      const matches = ['DGET(A1:C3, "Status", A4:A5)', 'A1:C3', '"Status"', 'A4:A5'] as RegExpMatchArray;
      const result = DGET.calculate(matches, textContext);
      expect(result).toBe('Active');
    });
  });

  describe('Edge Cases and Error Handling', () => {
    it('should handle empty database', () => {
      const emptyContext = createContext([]);
      const matches = ['DSUM(A1:C1, "Field", A1:A1)', 'A1:C1', '"Field"', 'A1:A1'] as RegExpMatchArray;
      const result = DSUM.calculate(matches, emptyContext);
      expect(result).toBe(FormulaError.VALUE);
    });

    it('should handle database with only headers', () => {
      const headersOnlyContext = createContext([
        ['Name', 'Salary']
      ]);
      const matches = ['DSUM(A1:B1, "Salary", A1:A1)', 'A1:B1', '"Salary"', 'A1:A1'] as RegExpMatchArray;
      const result = DSUM.calculate(matches, headersOnlyContext);
      expect(result).toBe(FormulaError.VALUE);
    });

    it('should handle wildcard criteria', () => {
      const wildcardContext = createContext([
        ['Name', 'Department'],
        ['Alice', 'Sales'],
        ['Alex', 'Sales'],
        ['Bob', 'Support'],
        ['Name'],
        ['Al*'] // Wildcard for names starting with "Al"
      ]);
      
      const matches = ['DCOUNTA(A1:B4, "Name", A5:A6)', 'A1:B4', '"Name"', 'A5:A6'] as RegExpMatchArray;
      const result = DCOUNTA.calculate(matches, wildcardContext);
      expect(result).toBe(2); // Alice and Alex
    });

    it('should handle criteria with <= operator', () => {
      const comparisonContext = createContext([
        ['Name', 'Age'],
        ['Alice', 25],
        ['Bob', 30],
        ['Charlie', 35],
        ['Age'],
        ['<=30']
      ]);
      
      const matches = ['DCOUNT(A1:B4, "Age", A5:A6)', 'A1:B4', '"Age"', 'A5:A6'] as RegExpMatchArray;
      const result = DCOUNT.calculate(matches, comparisonContext);
      expect(result).toBe(2); // Alice(25) and Bob(30)
    });

    it('should handle criteria with >= operator', () => {
      const comparisonContext = createContext([
        ['Name', 'Age'],
        ['Alice', 25],
        ['Bob', 30],
        ['Charlie', 35],
        ['Age'],
        ['>=30']
      ]);
      
      const matches = ['DCOUNT(A1:B4, "Age", A5:A6)', 'A1:B4', '"Age"', 'A5:A6'] as RegExpMatchArray;
      const result = DCOUNT.calculate(matches, comparisonContext);
      expect(result).toBe(2); // Bob(30) and Charlie(35)
    });

    it('should handle multiple criteria columns', () => {
      const multiCriteriaContext = createContext([
        ['Name', 'Department', 'Age', 'Salary'],
        ['Alice', 'Sales', 25, 50000],
        ['Bob', 'Sales', 30, 45000],
        ['Charlie', 'Engineering', 28, 75000],
        ['Department', 'Age'],
        ['Sales', '>25']
      ]);
      
      const matches = ['DSUM(A1:D4, "Salary", A5:B6)', 'A1:D4', '"Salary"', 'A5:B6'] as RegExpMatchArray;
      const result = DSUM.calculate(matches, multiCriteriaContext);
      expect(result).toBe(45000); // Only Bob meets both criteria (Sales AND Age > 25)
    });
  });
});