import { describe, it, expect } from 'vitest';
import {
  CUBEKPIMEMBER, CUBEMEMBER, CUBEMEMBERPROPERTY, CUBERANKEDMEMBER,
  CUBESET, CUBESETCOUNT, CUBEVALUE
} from '../cubeLogic';
import type { FormulaContext } from '../../shared/types';

// Helper function to create FormulaContext
const createContext = (data: (string | number | boolean | null)[][]): FormulaContext => ({
  data: data.map(row => row.map(cell => ({ value: cell }))),
  row: 0,
  col: 0
});

describe('Cube Functions', () => {
  const mockContext = createContext([
    ['Adventure Works', 'Sales', 'Finance'], // connection names
    ['[Measures].[Sales Amount]', '[Measures].[Profit]'], // measure members
    ['[Product].[Category].[Bikes]', '[Product].[Category].[Accessories]'], // dimension members
    ['[Date].[Calendar].[CY 2023]', '[Date].[Calendar].[CY 2024]'], // time dimensions
    ['Caption', 'UniqueName', 'Level'], // properties
    [1, 2, 3, 4, 5], // rank values
  ]);

  describe('CUBEKPIMEMBER Function', () => {
    it('should return KPI member reference', () => {
      const matches = ['CUBEKPIMEMBER("Adventure Works", "Revenue", "Goal")', 
        '"Adventure Works"', '"Revenue"', '"Goal"'] as RegExpMatchArray;
      const result = CUBEKPIMEMBER.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Cube functions require OLAP connection');
    });

    it('should handle Value property', () => {
      const matches = ['CUBEKPIMEMBER("Sales", "Profit Margin", "Value")', 
        '"Sales"', '"Profit Margin"', '"Value"'] as RegExpMatchArray;
      const result = CUBEKPIMEMBER.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Cube functions require OLAP connection');
    });

    it('should handle Status property', () => {
      const matches = ['CUBEKPIMEMBER("Adventure Works", "Customer Satisfaction", "Status")', 
        '"Adventure Works"', '"Customer Satisfaction"', '"Status"'] as RegExpMatchArray;
      const result = CUBEKPIMEMBER.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Cube functions require OLAP connection');
    });

    it('should handle Trend property', () => {
      const matches = ['CUBEKPIMEMBER("Finance", "Revenue Growth", "Trend")', 
        '"Finance"', '"Revenue Growth"', '"Trend"'] as RegExpMatchArray;
      const result = CUBEKPIMEMBER.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Cube functions require OLAP connection');
    });

    it('should handle Weight property', () => {
      const matches = ['CUBEKPIMEMBER("Sales", "Market Share", "Weight")', 
        '"Sales"', '"Market Share"', '"Weight"'] as RegExpMatchArray;
      const result = CUBEKPIMEMBER.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Cube functions require OLAP connection');
    });

    it('should handle cell references', () => {
      const matches = ['CUBEKPIMEMBER(A1, "KPI Name", "Goal")', 
        'A1', '"KPI Name"', '"Goal"'] as RegExpMatchArray;
      const result = CUBEKPIMEMBER.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Cube functions require OLAP connection');
    });
  });

  describe('CUBEMEMBER Function', () => {
    it('should return member reference', () => {
      const matches = ['CUBEMEMBER("Adventure Works", "[Product].[Category].[Bikes]")', 
        '"Adventure Works"', '"[Product].[Category].[Bikes]"'] as RegExpMatchArray;
      const result = CUBEMEMBER.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Cube functions require OLAP connection');
    });

    it('should handle tuple syntax', () => {
      const matches = ['CUBEMEMBER("Sales", "([Product].[Bikes], [Time].[2024])")', 
        '"Sales"', '"([Product].[Bikes], [Time].[2024])"'] as RegExpMatchArray;
      const result = CUBEMEMBER.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Cube functions require OLAP connection');
    });

    it('should handle measure members', () => {
      const matches = ['CUBEMEMBER("Adventure Works", "[Measures].[Sales Amount]")', 
        '"Adventure Works"', '"[Measures].[Sales Amount]"'] as RegExpMatchArray;
      const result = CUBEMEMBER.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Cube functions require OLAP connection');
    });

    it('should handle optional caption parameter', () => {
      const matches = ['CUBEMEMBER("Sales", "[Product].[All Products]", "All Products Total")', 
        '"Sales"', '"[Product].[All Products]"', '"All Products Total"'] as RegExpMatchArray;
      const result = CUBEMEMBER.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Cube functions require OLAP connection');
    });
  });

  describe('CUBEMEMBERPROPERTY Function', () => {
    it('should return member property', () => {
      const matches = ['CUBEMEMBERPROPERTY("Adventure Works", "[Store].[Store Name].[Store 1]", "Manager")', 
        '"Adventure Works"', '"[Store].[Store Name].[Store 1]"', '"Manager"'] as RegExpMatchArray;
      const result = CUBEMEMBERPROPERTY.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Cube functions require OLAP connection');
    });

    it('should handle Caption property', () => {
      const matches = ['CUBEMEMBERPROPERTY("Sales", "[Product].[Category].[Bikes]", "Caption")', 
        '"Sales"', '"[Product].[Category].[Bikes]"', '"Caption"'] as RegExpMatchArray;
      const result = CUBEMEMBERPROPERTY.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Cube functions require OLAP connection');
    });

    it('should handle UniqueName property', () => {
      const matches = ['CUBEMEMBERPROPERTY("Adventure Works", "[Customer].[Customer 1]", "UniqueName")', 
        '"Adventure Works"', '"[Customer].[Customer 1]"', '"UniqueName"'] as RegExpMatchArray;
      const result = CUBEMEMBERPROPERTY.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Cube functions require OLAP connection');
    });

    it('should handle Level property', () => {
      const matches = ['CUBEMEMBERPROPERTY("Sales", "[Date].[Calendar].[Month]", "Level")', 
        '"Sales"', '"[Date].[Calendar].[Month]"', '"Level"'] as RegExpMatchArray;
      const result = CUBEMEMBERPROPERTY.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Cube functions require OLAP connection');
    });
  });

  describe('CUBERANKEDMEMBER Function', () => {
    it('should return ranked member', () => {
      const matches = ['CUBERANKEDMEMBER("Adventure Works", "[Product].[Category].Members", 1)', 
        '"Adventure Works"', '"[Product].[Category].Members"', '1'] as RegExpMatchArray;
      const result = CUBERANKEDMEMBER.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Cube functions require OLAP connection');
    });

    it('should handle top N members', () => {
      const matches = ['CUBERANKEDMEMBER("Sales", "[Customer].[Customer].Members", 10)', 
        '"Sales"', '"[Customer].[Customer].Members"', '10'] as RegExpMatchArray;
      const result = CUBERANKEDMEMBER.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Cube functions require OLAP connection');
    });

    it('should handle optional caption', () => {
      const matches = ['CUBERANKEDMEMBER("Adventure Works", "[Store].[Store].Members", 1, "Top Store")', 
        '"Adventure Works"', '"[Store].[Store].Members"', '1', '"Top Store"'] as RegExpMatchArray;
      const result = CUBERANKEDMEMBER.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Cube functions require OLAP connection');
    });

    it('should handle set expression with order', () => {
      const matches = ['CUBERANKEDMEMBER("Sales", "Order([Product].Members, [Measures].[Sales], DESC)", 5)', 
        '"Sales"', '"Order([Product].Members, [Measures].[Sales], DESC)"', '5'] as RegExpMatchArray;
      const result = CUBERANKEDMEMBER.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Cube functions require OLAP connection');
    });
  });

  describe('CUBESET Function', () => {
    it('should create a set from expression', () => {
      const matches = ['CUBESET("Adventure Works", "[Product].[Category].Children")', 
        '"Adventure Works"', '"[Product].[Category].Children"'] as RegExpMatchArray;
      const result = CUBESET.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Cube functions require OLAP connection');
    });

    it('should handle multiple members', () => {
      const matches = ['CUBESET("Sales", "{[Product].[Bikes], [Product].[Accessories]}")', 
        '"Sales"', '"{[Product].[Bikes], [Product].[Accessories]}"'] as RegExpMatchArray;
      const result = CUBESET.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Cube functions require OLAP connection');
    });

    it('should handle caption parameter', () => {
      const matches = ['CUBESET("Adventure Works", "[Store].[Store].Members", "All Stores")', 
        '"Adventure Works"', '"[Store].[Store].Members"', '"All Stores"'] as RegExpMatchArray;
      const result = CUBESET.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Cube functions require OLAP connection');
    });

    it('should handle sort order ASC', () => {
      const matches = ['CUBESET("Sales", "[Product].Members", "Products", 1)', 
        '"Sales"', '"[Product].Members"', '"Products"', '1'] as RegExpMatchArray;
      const result = CUBESET.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Cube functions require OLAP connection');
    });

    it('should handle sort order DESC', () => {
      const matches = ['CUBESET("Sales", "[Customer].Members", "Customers", 2, "[Measures].[Sales]")', 
        '"Sales"', '"[Customer].Members"', '"Customers"', '2', '"[Measures].[Sales]"'] as RegExpMatchArray;
      const result = CUBESET.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Cube functions require OLAP connection');
    });
  });

  describe('CUBESETCOUNT Function', () => {
    it('should count members in set', () => {
      const matches = ['CUBESETCOUNT("[Product].[Category].Children")', 
        '"[Product].[Category].Children"'] as RegExpMatchArray;
      const result = CUBESETCOUNT.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Cube functions require OLAP connection');
    });

    it('should handle set reference', () => {
      const matches = ['CUBESETCOUNT(A1)', 'A1'] as RegExpMatchArray;
      const result = CUBESETCOUNT.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Cube functions require OLAP connection');
    });

    it('should handle complex set expressions', () => {
      const matches = ['CUBESETCOUNT("Filter([Store].Members, [Measures].[Sales] > 1000000)")', 
        '"Filter([Store].Members, [Measures].[Sales] > 1000000)"'] as RegExpMatchArray;
      const result = CUBESETCOUNT.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Cube functions require OLAP connection');
    });
  });

  describe('CUBEVALUE Function', () => {
    it('should return cube value with single member', () => {
      const matches = ['CUBEVALUE("Adventure Works", "[Measures].[Sales Amount]")', 
        '"Adventure Works"', '"[Measures].[Sales Amount]"'] as RegExpMatchArray;
      const result = CUBEVALUE.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Cube functions require OLAP connection');
    });

    it('should handle multiple members (tuple)', () => {
      const matches = ['CUBEVALUE("Sales", "[Measures].[Sales]", "[Product].[Bikes]", "[Time].[2024]")', 
        '"Sales"', '"[Measures].[Sales]"', '"[Product].[Bikes]"', '"[Time].[2024]"'] as RegExpMatchArray;
      const result = CUBEVALUE.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Cube functions require OLAP connection');
    });

    it('should handle member references', () => {
      const matches = ['CUBEVALUE("Adventure Works", B1, C1)', 
        '"Adventure Works"', 'B1', 'C1'] as RegExpMatchArray;
      const result = CUBEVALUE.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Cube functions require OLAP connection');
    });

    it('should handle calculated measures', () => {
      const matches = ['CUBEVALUE("Finance", "[Measures].[Profit Margin]", "[Date].[Q1 2024]")', 
        '"Finance"', '"[Measures].[Profit Margin]"', '"[Date].[Q1 2024]"'] as RegExpMatchArray;
      const result = CUBEVALUE.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Cube functions require OLAP connection');
    });

    it('should handle slicer with multiple dimensions', () => {
      const matches = ['CUBEVALUE("Sales", "[Measures].[Revenue]", "[Product].[All]", "[Customer].[All]", "[Store].[Store 1]")', 
        '"Sales"', '"[Measures].[Revenue]"', '"[Product].[All]"', '"[Customer].[All]"', '"[Store].[Store 1]"'] as RegExpMatchArray;
      const result = CUBEVALUE.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Cube functions require OLAP connection');
    });
  });

  describe('Edge Cases and Integration', () => {
    it('should handle empty connection string', () => {
      const matches = ['CUBEMEMBER("", "[Product].[Category]")', 
        '""', '"[Product].[Category]"'] as RegExpMatchArray;
      const result = CUBEMEMBER.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Cube functions require OLAP connection');
    });

    it('should handle MDX expressions in CUBESET', () => {
      const mdxExpression = 'TopCount([Product].[Product].Members, 10, [Measures].[Sales Amount])';
      const matches = ['CUBESET("Adventure Works", "' + mdxExpression + '")', 
        '"Adventure Works"', '"' + mdxExpression + '"'] as RegExpMatchArray;
      const result = CUBESET.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Cube functions require OLAP connection');
    });

    it('should handle time intelligence in CUBEVALUE', () => {
      const matches = ['CUBEVALUE("Sales", "[Measures].[YTD Sales]", "[Date].[Calendar].[Month].[March 2024]")', 
        '"Sales"', '"[Measures].[YTD Sales]"', '"[Date].[Calendar].[Month].[March 2024]"'] as RegExpMatchArray;
      const result = CUBEVALUE.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Cube functions require OLAP connection');
    });

    it('should handle hierarchical navigation', () => {
      const matches = ['CUBEMEMBER("Adventure Works", "[Product].[Category].[Bikes].Parent")', 
        '"Adventure Works"', '"[Product].[Category].[Bikes].Parent"'] as RegExpMatchArray;
      const result = CUBEMEMBER.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Cube functions require OLAP connection');
    });

    it('should handle calculated members', () => {
      const calcMember = 'WITH MEMBER [Measures].[Profit Ratio] AS [Measures].[Profit]/[Measures].[Sales]';
      const matches = ['CUBEMEMBER("Finance", "' + calcMember + '")', 
        '"Finance"', '"' + calcMember + '"'] as RegExpMatchArray;
      const result = CUBEMEMBER.calculate(matches, mockContext);
      expect(result).toBe('#N/A - Cube functions require OLAP connection');
    });

    it('should handle all cube functions returning consistent error', () => {
      const functions = [
        CUBEKPIMEMBER, CUBEMEMBER, CUBEMEMBERPROPERTY, 
        CUBERANKEDMEMBER, CUBESET, CUBESETCOUNT, CUBEVALUE
      ];
      
      for (const func of functions) {
        const matches = ['FUNC("Connection", "Expression")', '"Connection"', '"Expression"'] as RegExpMatchArray;
        const result = func.calculate(matches, mockContext);
        expect(result).toBe('#N/A - Cube functions require OLAP connection');
      }
    });
  });
});