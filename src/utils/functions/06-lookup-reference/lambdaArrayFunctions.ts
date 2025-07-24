// Lambda配列関数の実装

import type { CustomFormula, FormulaContext, FormulaResult } from '../shared/types';
import { FormulaError } from '../shared/types';
import { getCellValue } from '../shared/utils';

// BYROW関数の実装（行ごとに関数を適用）
export const BYROW: CustomFormula = {
  name: 'BYROW',
  pattern: /BYROW\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, arrayRef, lambdaRef] = matches;
    
    try {
      // データ配列を取得
      const data: any[][] = [];
      const rangeMatch = arrayRef.trim().match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
      
      if (!rangeMatch) {
        return FormulaError.REF;
      }
      
      const [, startCol, startRowStr, endCol, endRowStr] = rangeMatch;
      const startRow = parseInt(startRowStr) - 1;
      const endRow = parseInt(endRowStr) - 1;
      const startColIndex = startCol.charCodeAt(0) - 65;
      const endColIndex = endCol.charCodeAt(0) - 65;
      
      for (let row = startRow; row <= endRow; row++) {
        const rowData: any[] = [];
        for (let col = startColIndex; col <= endColIndex; col++) {
          const cell = context.data[row]?.[col];
          rowData.push(cell ? cell.value : null);
        }
        data.push(rowData);
      }
      
      // Lambda関数の解析（簡易実装）
      // 実際のLAMBDA関数のサポートにはより複雑な実装が必要
      const lambdaStr = lambdaRef.trim().toUpperCase();
      
      const results: any[] = [];
      
      for (const row of data) {
        let result: any;
        
        // 簡易的な関数適用
        if (lambdaStr.includes('SUM')) {
          result = row.reduce((sum, val) => {
            const num = Number(val);
            return isNaN(num) ? sum : sum + num;
          }, 0);
        } else if (lambdaStr.includes('AVERAGE')) {
          const nums = row.filter(val => !isNaN(Number(val))).map(Number);
          result = nums.length > 0 ? nums.reduce((sum, val) => sum + val, 0) / nums.length : 0;
        } else if (lambdaStr.includes('COUNT')) {
          result = row.filter(val => !isNaN(Number(val))).length;
        } else if (lambdaStr.includes('MAX')) {
          const nums = row.filter(val => !isNaN(Number(val))).map(Number);
          result = nums.length > 0 ? Math.max(...nums) : 0;
        } else if (lambdaStr.includes('MIN')) {
          const nums = row.filter(val => !isNaN(Number(val))).map(Number);
          result = nums.length > 0 ? Math.min(...nums) : 0;
        } else {
          // デフォルトは合計
          result = row.reduce((sum, val) => {
            const num = Number(val);
            return isNaN(num) ? sum : sum + num;
          }, 0);
        }
        
        results.push([result]);
      }
      
      return results;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// BYCOL関数の実装（列ごとに関数を適用）
export const BYCOL: CustomFormula = {
  name: 'BYCOL',
  pattern: /BYCOL\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, arrayRef, lambdaRef] = matches;
    
    try {
      // データ配列を取得
      const data: any[][] = [];
      const rangeMatch = arrayRef.trim().match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
      
      if (!rangeMatch) {
        return FormulaError.REF;
      }
      
      const [, startCol, startRowStr, endCol, endRowStr] = rangeMatch;
      const startRow = parseInt(startRowStr) - 1;
      const endRow = parseInt(endRowStr) - 1;
      const startColIndex = startCol.charCodeAt(0) - 65;
      const endColIndex = endCol.charCodeAt(0) - 65;
      
      for (let row = startRow; row <= endRow; row++) {
        const rowData: any[] = [];
        for (let col = startColIndex; col <= endColIndex; col++) {
          const cell = context.data[row]?.[col];
          rowData.push(cell ? cell.value : null);
        }
        data.push(rowData);
      }
      
      // Lambda関数の解析（簡易実装）
      const lambdaStr = lambdaRef.trim().toUpperCase();
      
      const results: any[] = [];
      const numCols = data[0]?.length ?? 0;
      
      for (let col = 0; col < numCols; col++) {
        const column = data.map(row => row[col]);
        let result: any;
        
        // 簡易的な関数適用
        if (lambdaStr.includes('SUM')) {
          result = column.reduce((sum, val) => {
            const num = Number(val);
            return isNaN(num) ? sum : sum + num;
          }, 0);
        } else if (lambdaStr.includes('AVERAGE')) {
          const nums = column.filter(val => !isNaN(Number(val))).map(Number);
          result = nums.length > 0 ? nums.reduce((sum, val) => sum + val, 0) / nums.length : 0;
        } else if (lambdaStr.includes('COUNT')) {
          result = column.filter(val => !isNaN(Number(val))).length;
        } else if (lambdaStr.includes('MAX')) {
          const nums = column.filter(val => !isNaN(Number(val))).map(Number);
          result = nums.length > 0 ? Math.max(...nums) : 0;
        } else if (lambdaStr.includes('MIN')) {
          const nums = column.filter(val => !isNaN(Number(val))).map(Number);
          result = nums.length > 0 ? Math.min(...nums) : 0;
        } else {
          // デフォルトは合計
          result = column.reduce((sum, val) => {
            const num = Number(val);
            return isNaN(num) ? sum : sum + num;
          }, 0);
        }
        
        results.push(result);
      }
      
      return [results];
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// MAP関数の実装（配列の各要素に関数を適用）
export const MAP: CustomFormula = {
  name: 'MAP',
  pattern: /MAP\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, arrayRef, lambdaRef] = matches;
    
    try {
      // データ配列を取得
      const data: any[][] = [];
      const rangeMatch = arrayRef.trim().match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
      
      if (!rangeMatch) {
        return FormulaError.REF;
      }
      
      const [, startCol, startRowStr, endCol, endRowStr] = rangeMatch;
      const startRow = parseInt(startRowStr) - 1;
      const endRow = parseInt(endRowStr) - 1;
      const startColIndex = startCol.charCodeAt(0) - 65;
      const endColIndex = endCol.charCodeAt(0) - 65;
      
      for (let row = startRow; row <= endRow; row++) {
        const rowData: any[] = [];
        for (let col = startColIndex; col <= endColIndex; col++) {
          const cell = context.data[row]?.[col];
          rowData.push(cell ? cell.value : null);
        }
        data.push(rowData);
      }
      
      // Lambda関数の解析（簡易実装）
      const lambdaStr = lambdaRef.trim().toUpperCase();
      
      const results: any[][] = [];
      
      for (const row of data) {
        const resultRow: any[] = [];
        for (const value of row) {
          let result: any;
          
          // 簡易的な関数適用
          if (lambdaStr.includes('*2')) {
            const num = Number(value);
            result = isNaN(num) ? value : num * 2;
          } else if (lambdaStr.includes('+1')) {
            const num = Number(value);
            result = isNaN(num) ? value : num + 1;
          } else if (lambdaStr.includes('SQRT')) {
            const num = Number(value);
            result = isNaN(num) || num < 0 ? FormulaError.NUM : Math.sqrt(num);
          } else if (lambdaStr.includes('UPPER')) {
            result = String(value).toUpperCase();
          } else if (lambdaStr.includes('LOWER')) {
            result = String(value).toLowerCase();
          } else {
            // デフォルトはそのまま返す
            result = value;
          }
          
          resultRow.push(result);
        }
        results.push(resultRow);
      }
      
      return results;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// REDUCE関数の実装（配列を単一の値に縮約）
export const REDUCE: CustomFormula = {
  name: 'REDUCE',
  pattern: /REDUCE\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, initialRef, arrayRef, lambdaRef] = matches;
    
    try {
      // 初期値を取得
      const initial = getCellValue(initialRef.trim(), context) ?? 0;
      
      // データ配列を取得
      const values: any[] = [];
      
      if (arrayRef.includes(':')) {
        const rangeMatch = arrayRef.trim().match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
        if (!rangeMatch) {
          return FormulaError.REF;
        }
        
        const [, startCol, startRowStr, endCol, endRowStr] = rangeMatch;
        const startRow = parseInt(startRowStr) - 1;
        const endRow = parseInt(endRowStr) - 1;
        const startColIndex = startCol.charCodeAt(0) - 65;
        const endColIndex = endCol.charCodeAt(0) - 65;
        
        for (let row = startRow; row <= endRow; row++) {
          for (let col = startColIndex; col <= endColIndex; col++) {
            const cell = context.data[row]?.[col];
            if (cell) {
              values.push(cell.value);
            }
          }
        }
      } else {
        values.push(getCellValue(arrayRef.trim(), context));
      }
      
      // Lambda関数の解析（簡易実装）
      const lambdaStr = lambdaRef.trim().toUpperCase();
      
      let result: any = initial;
      
      for (const value of values) {
        // 簡易的な関数適用
        if (lambdaStr.includes('+')) {
          const num = Number(value);
          if (!isNaN(num)) {
            result = Number(result) + num;
          }
        } else if (lambdaStr.includes('*')) {
          const num = Number(value);
          if (!isNaN(num)) {
            result = Number(result) * num;
          }
        } else if (lambdaStr.includes('MAX')) {
          const num = Number(value);
          if (!isNaN(num)) {
            result = Math.max(Number(result), num);
          }
        } else if (lambdaStr.includes('MIN')) {
          const num = Number(value);
          if (!isNaN(num)) {
            result = Math.min(Number(result), num);
          }
        } else {
          // デフォルトは加算
          const num = Number(value);
          if (!isNaN(num)) {
            result = Number(result) + num;
          }
        }
      }
      
      return result;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// SCAN関数の実装（配列の累積計算）
export const SCAN: CustomFormula = {
  name: 'SCAN',
  pattern: /SCAN\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, initialRef, arrayRef, lambdaRef] = matches;
    
    try {
      // 初期値を取得
      const initial = getCellValue(initialRef.trim(), context) ?? 0;
      
      // データ配列を取得
      const data: any[][] = [];
      const rangeMatch = arrayRef.trim().match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
      
      if (!rangeMatch) {
        return FormulaError.REF;
      }
      
      const [, startCol, startRowStr, endCol, endRowStr] = rangeMatch;
      const startRow = parseInt(startRowStr) - 1;
      const endRow = parseInt(endRowStr) - 1;
      const startColIndex = startCol.charCodeAt(0) - 65;
      const endColIndex = endCol.charCodeAt(0) - 65;
      
      for (let row = startRow; row <= endRow; row++) {
        const rowData: any[] = [];
        for (let col = startColIndex; col <= endColIndex; col++) {
          const cell = context.data[row]?.[col];
          rowData.push(cell ? cell.value : null);
        }
        data.push(rowData);
      }
      
      // Lambda関数の解析（簡易実装）
      const lambdaStr = lambdaRef.trim().toUpperCase();
      
      const results: any[][] = [];
      let accumulator = initial;
      
      for (const row of data) {
        const resultRow: any[] = [];
        for (const value of row) {
          // 簡易的な関数適用
          if (lambdaStr.includes('+')) {
            const num = Number(value);
            if (!isNaN(num)) {
              accumulator = Number(accumulator) + num;
            }
          } else if (lambdaStr.includes('*')) {
            const num = Number(value);
            if (!isNaN(num)) {
              accumulator = Number(accumulator) * num;
            }
          } else if (lambdaStr.includes('MAX')) {
            const num = Number(value);
            if (!isNaN(num)) {
              accumulator = Math.max(Number(accumulator), num);
            }
          } else if (lambdaStr.includes('MIN')) {
            const num = Number(value);
            if (!isNaN(num)) {
              accumulator = Math.min(Number(accumulator), num);
            }
          } else {
            // デフォルトは加算
            const num = Number(value);
            if (!isNaN(num)) {
              accumulator = Number(accumulator) + num;
            }
          }
          
          resultRow.push(accumulator);
        }
        results.push(resultRow);
      }
      
      return results;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// MAKEARRAY関数の実装（指定サイズの配列を作成）
export const MAKEARRAY: CustomFormula = {
  name: 'MAKEARRAY',
  pattern: /MAKEARRAY\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, rowsRef, colsRef, lambdaRef] = matches;
    
    try {
      const rows = parseInt(getCellValue(rowsRef.trim(), context)?.toString() ?? rowsRef.trim());
      const cols = parseInt(getCellValue(colsRef.trim(), context)?.toString() ?? colsRef.trim());
      
      if (isNaN(rows) || isNaN(cols) || rows <= 0 || cols <= 0) {
        return FormulaError.VALUE;
      }
      
      // Lambda関数の解析（簡易実装）
      const lambdaStr = lambdaRef.trim().toUpperCase();
      
      const results: any[][] = [];
      
      for (let row = 0; row < rows; row++) {
        const resultRow: any[] = [];
        for (let col = 0; col < cols; col++) {
          let result: any;
          
          // 簡易的な関数適用
          if (lambdaStr.includes('ROW') && lambdaStr.includes('COLUMN')) {
            // 行番号 + 列番号
            result = (row + 1) + (col + 1);
          } else if (lambdaStr.includes('ROW') && lambdaStr.includes('*')) {
            // 行番号 * 列番号
            result = (row + 1) * (col + 1);
          } else if (lambdaStr.includes('ROW')) {
            // 行番号のみ
            result = row + 1;
          } else if (lambdaStr.includes('COLUMN')) {
            // 列番号のみ
            result = col + 1;
          } else if (lambdaStr.includes('RAND')) {
            // ランダム値
            result = Math.random();
          } else {
            // デフォルトは行番号 * 列番号
            result = (row + 1) * (col + 1);
          }
          
          resultRow.push(result);
        }
        results.push(resultRow);
      }
      
      return results;
    } catch {
      return FormulaError.VALUE;
    }
  }
};