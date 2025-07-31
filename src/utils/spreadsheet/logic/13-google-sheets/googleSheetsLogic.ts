// Google Sheets specific functions

import type { CustomFormula, FormulaResult, FormulaContext } from '../shared/types';
import { FormulaError } from '../shared/types';
import { getCellValue, getCellRangeValues, parseCellRange } from '../shared/utils';
import { matchFormula } from '../index';

// JOIN関数の実装（配列を文字列に結合）
export const JOIN: CustomFormula = {
  name: 'JOIN',
  pattern: /\bJOIN\(([^,]+),\s*(.+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, separatorRef, arrayRef] = matches;
    
    try {
      // Get separator
      let separator = getCellValue(separatorRef.trim(), context);
      if (separator === separatorRef.trim() || separator === null || separator === undefined) {
        // Treat as literal string, remove quotes
        separator = separatorRef.trim().replace(/^["']|["']$/g, '');
      }
      
      // Get array values
      const arrayValues = getCellRangeValues(arrayRef.trim(), context);
      
      if (!Array.isArray(arrayValues) || arrayValues.length === 0) {
        return '';
      }
      
      // Convert all values to strings and filter out null/undefined/empty
      const stringValues = arrayValues
        .map(val => {
          if (val === null || val === undefined || val === '') {
            return '';
          }
          return String(val);
        })
        .filter(val => val !== '');
      
      return stringValues.join(String(separator));
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// ARRAYFORMULA関数の実装（配列数式を作成）
export const ARRAYFORMULA: CustomFormula = {
  name: 'ARRAYFORMULA',
  pattern: /ARRAYFORMULA\((.+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, formula] = matches;
    
    try {
      // For array formulas, we need to detect ranges and apply the formula to each element
      // This is a simplified implementation
      
      // Check if the formula contains range references
      const rangeMatches = formula.match(/([A-Z]+\d+:[A-Z]+\d+)/g);
      
      if (!rangeMatches || rangeMatches.length === 0) {
        // No ranges found, evaluate as normal formula
        const matchResult = matchFormula(formula.toUpperCase());
        if (matchResult) {
          const { function: func, matches: formulaMatches } = matchResult;
          return func.calculate(formulaMatches, context);
        } else {
          // Try to evaluate as expression
          return formula;
        }
      }
      
      // Get the first range to determine array size
      const firstRange = rangeMatches[0];
      const rangeCoords = parseCellRange(firstRange);
      
      if (!rangeCoords) {
        return FormulaError.REF;
      }
      
      const { start, end } = rangeCoords;
      const rows = end.row - start.row + 1;
      const cols = end.col - start.col + 1;
      
      // Create result array
      const results: (string | number | boolean | null)[][] = [];
      
      // Apply formula to each cell in the range
      for (let r = 0; r < rows; r++) {
        const row: (string | number | boolean | null)[] = [];
        for (let c = 0; c < cols; c++) {
          try {
            // Replace range references with individual cell references
            let cellFormula = formula;
            for (const range of rangeMatches) {
              const currentRangeCoords = parseCellRange(range);
              if (currentRangeCoords) {
                const cellRow = currentRangeCoords.start.row + r;
                const cellCol = currentRangeCoords.start.col + c;
                const cellRef = String.fromCharCode(65 + cellCol) + (cellRow + 1);
                cellFormula = cellFormula.replace(range, cellRef);
              }
            }
            
            // Evaluate the cell formula
            const matchResult = matchFormula(cellFormula.toUpperCase());
            if (matchResult) {
              const { function: func, matches: formulaMatches } = matchResult;
              const result = func.calculate(formulaMatches, context);
              row.push(result as (string | number | boolean | null));
            } else {
              // Get cell value directly
              const cellValue = context.data[start.row + r]?.[start.col + c];
              const value = cellValue?.value ?? cellValue?.v ?? cellValue ?? null;
              row.push(value as (string | number | boolean | null));
            }
          } catch {
            row.push(FormulaError.VALUE as string);
          }
        }
        results.push(row);
      }
      
      // Return single value if 1x1, otherwise return array
      if (results.length === 1 && results[0].length === 1) {
        return results[0][0] as FormulaResult;
      }
      
      return results as FormulaResult;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// QUERY関数の実装（簡易版 - SQLライクなクエリ）
export const QUERY: CustomFormula = {
  name: 'QUERY',
  pattern: /QUERY\(([^,]+)(?:,\s*"([^"]*)")?(?:,\s*(-?\d+))?\)/i,
  calculate: (): FormulaResult => {
    return '#N/A - Google Sheets specific functions not supported';
  }
};


// REGEXMATCH関数の実装（正規表現マッチング）
export const REGEXMATCH: CustomFormula = {
  name: 'REGEXMATCH',
  pattern: /REGEXMATCH\(([^,]+),\s*"([^"]+)"\)/i,
  calculate: (): FormulaResult => {
    return '#N/A - Google Sheets specific functions not supported';
  }
};

// REGEXEXTRACT関数の実装（正規表現による抽出）
export const REGEXEXTRACT: CustomFormula = {
  name: 'REGEXEXTRACT',
  pattern: /REGEXEXTRACT\(([^,]+),\s*"([^"]+)"\)/i,
  calculate: (): FormulaResult => {
    return '#N/A - Google Sheets specific functions not supported';
  }
};

// REGEXREPLACE関数の実装（正規表現による置換）
export const REGEXREPLACE: CustomFormula = {
  name: 'REGEXREPLACE',
  pattern: /REGEXREPLACE\(([^,]+),\s*"([^"]+)",\s*"([^"]*)"\)/i,
  calculate: (): FormulaResult => {
    return '#N/A - Google Sheets specific functions not supported';
  }
};

// FLATTEN関数の実装（配列を1次元化）
export const FLATTEN: CustomFormula = {
  name: 'FLATTEN',
  pattern: /FLATTEN\((.+)\)/i,
  calculate: (): FormulaResult => {
    return '#N/A - Google Sheets specific functions not supported';
  }
};