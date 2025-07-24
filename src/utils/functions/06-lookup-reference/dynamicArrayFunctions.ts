// 動的配列関数の実装

import type { CustomFormula, FormulaContext, FormulaResult } from '../shared/types';
import { FormulaError } from '../shared/types';
import { getCellValue } from '../shared/utils';

// SORTBY関数の実装（別の配列で並べ替え）
export const SORTBY: CustomFormula = {
  name: 'SORTBY',
  pattern: /SORTBY\(([^,]+),\s*([^,)]+)(?:,\s*([^,)]+))?(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, arrayRef, sortArrayRef, order1Ref] = matches;
    
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
      
      // ソート配列を取得
      const sortValues1: any[] = [];
      const sortRange1Match = sortArrayRef.trim().match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
      
      if (!sortRange1Match) {
        return FormulaError.REF;
      }
      
      const [, s1StartCol, s1StartRowStr, , s1EndRowStr] = sortRange1Match;
      const s1StartRow = parseInt(s1StartRowStr) - 1;
      const s1EndRow = parseInt(s1EndRowStr) - 1;
      const s1StartColIndex = s1StartCol.charCodeAt(0) - 65;
      
      for (let row = s1StartRow; row <= s1EndRow; row++) {
        const cell = context.data[row]?.[s1StartColIndex];
        sortValues1.push(cell ? cell.value : null);
      }
      
      if (data.length !== sortValues1.length) {
        return FormulaError.VALUE;
      }
      
      const order1 = order1Ref ? parseInt(getCellValue(order1Ref.trim(), context)?.toString() ?? order1Ref.trim()) : 1;
      
      // インデックス配列を作成
      const indices = Array.from({ length: data.length }, (_, i) => i);
      
      // ソート
      indices.sort((a, b) => {
        const valA = sortValues1[a];
        const valB = sortValues1[b];
        
        if (valA === null || valA === undefined) return 1;
        if (valB === null || valB === undefined) return -1;
        
        let comparison = 0;
        if (typeof valA === 'number' && typeof valB === 'number') {
          comparison = valA - valB;
        } else {
          comparison = String(valA).localeCompare(String(valB));
        }
        
        return order1 === -1 ? -comparison : comparison;
      });
      
      // ソートされた結果を返す
      return indices.map(i => data[i]);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// TAKE関数の実装（行/列を取得）
export const TAKE: CustomFormula = {
  name: 'TAKE',
  pattern: /TAKE\(([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, arrayRef, rowsRef, colsRef] = matches;
    
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
      
      const rows = parseInt(getCellValue(rowsRef.trim(), context)?.toString() ?? rowsRef.trim());
      const cols = colsRef ? parseInt(getCellValue(colsRef.trim(), context)?.toString() ?? colsRef.trim()) : 
                             (data[0]?.length ?? 0);
      
      if (isNaN(rows) || (colsRef && isNaN(cols))) {
        return FormulaError.VALUE;
      }
      
      if (rows === 0 || cols === 0) {
        return [];
      }
      
      // 行数を取得
      const takeRows = rows > 0 ? 
        data.slice(0, Math.min(rows, data.length)) :
        data.slice(Math.max(0, data.length + rows));
      
      // 列数を取得
      if (cols !== data[0]?.length) {
        return takeRows.map(row => {
          if (cols > 0) {
            return row.slice(0, Math.min(cols, row.length));
          } else {
            return row.slice(Math.max(0, row.length + cols));
          }
        });
      }
      
      return takeRows;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// DROP関数の実装（行/列を除外）
export const DROP: CustomFormula = {
  name: 'DROP',
  pattern: /DROP\(([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, arrayRef, rowsRef, colsRef] = matches;
    
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
      
      const rows = parseInt(getCellValue(rowsRef.trim(), context)?.toString() ?? rowsRef.trim());
      const cols = colsRef ? parseInt(getCellValue(colsRef.trim(), context)?.toString() ?? colsRef.trim()) : 0;
      
      if (isNaN(rows) || (colsRef && isNaN(cols))) {
        return FormulaError.VALUE;
      }
      
      // 行を除外
      let result = [...data];
      if (rows > 0) {
        result = result.slice(rows);
      } else if (rows < 0) {
        result = result.slice(0, Math.max(0, result.length + rows));
      }
      
      // 列を除外
      if (cols !== 0) {
        result = result.map(row => {
          if (cols > 0) {
            return row.slice(cols);
          } else {
            return row.slice(0, Math.max(0, row.length + cols));
          }
        });
      }
      
      return result;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// EXPAND関数の実装（配列を拡張）
export const EXPAND: CustomFormula = {
  name: 'EXPAND',
  pattern: /EXPAND\(([^,]+),\s*([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, arrayRef, rowsRef, colsRef, padRef] = matches;
    
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
      
      const newRows = parseInt(getCellValue(rowsRef.trim(), context)?.toString() ?? rowsRef.trim());
      const newCols = parseInt(getCellValue(colsRef.trim(), context)?.toString() ?? colsRef.trim());
      const padValue = padRef ? getCellValue(padRef.trim(), context) ?? padRef.trim() : FormulaError.NA;
      
      if (isNaN(newRows) || isNaN(newCols)) {
        return FormulaError.VALUE;
      }
      
      if (newRows < data.length || newCols < (data[0]?.length ?? 0)) {
        return FormulaError.VALUE;
      }
      
      // 結果配列を作成
      const result: any[][] = [];
      
      for (let row = 0; row < newRows; row++) {
        const resultRow: any[] = [];
        for (let col = 0; col < newCols; col++) {
          if (row < data.length && col < data[row].length) {
            resultRow.push(data[row][col]);
          } else {
            resultRow.push(padValue);
          }
        }
        result.push(resultRow);
      }
      
      return result;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// HSTACK関数の実装（水平方向に結合）
export const HSTACK: CustomFormula = {
  name: 'HSTACK',
  pattern: /HSTACK\((.+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, argsStr] = matches;
    
    try {
      const args = argsStr.split(',').map(arg => arg.trim());
      const arrays: any[][][] = [];
      let maxRows = 0;
      
      // 各配列を取得
      for (const arg of args) {
        if (arg.includes(':')) {
          const rangeMatch = arg.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
          if (!rangeMatch) {
            return FormulaError.REF;
          }
          
          const [, startCol, startRowStr, endCol, endRowStr] = rangeMatch;
          const startRow = parseInt(startRowStr) - 1;
          const endRow = parseInt(endRowStr) - 1;
          const startColIndex = startCol.charCodeAt(0) - 65;
          const endColIndex = endCol.charCodeAt(0) - 65;
          
          const array: any[][] = [];
          for (let row = startRow; row <= endRow; row++) {
            const rowData: any[] = [];
            for (let col = startColIndex; col <= endColIndex; col++) {
              const cell = context.data[row]?.[col];
              rowData.push(cell ? cell.value : null);
            }
            array.push(rowData);
          }
          
          arrays.push(array);
          maxRows = Math.max(maxRows, array.length);
        } else {
          // 単一セルの場合
          const value = getCellValue(arg, context);
          arrays.push([[value]]);
          maxRows = Math.max(maxRows, 1);
        }
      }
      
      // 水平結合
      const result: any[][] = [];
      for (let row = 0; row < maxRows; row++) {
        const resultRow: any[] = [];
        for (const array of arrays) {
          if (row < array.length) {
            resultRow.push(...array[row]);
          } else {
            // 短い配列の場合、#N/Aで埋める
            resultRow.push(...Array(array[0].length).fill(FormulaError.NA));
          }
        }
        result.push(resultRow);
      }
      
      return result;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// VSTACK関数の実装（垂直方向に結合）
export const VSTACK: CustomFormula = {
  name: 'VSTACK',
  pattern: /VSTACK\((.+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, argsStr] = matches;
    
    try {
      const args = argsStr.split(',').map(arg => arg.trim());
      const arrays: any[][][] = [];
      let maxCols = 0;
      
      // 各配列を取得
      for (const arg of args) {
        if (arg.includes(':')) {
          const rangeMatch = arg.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
          if (!rangeMatch) {
            return FormulaError.REF;
          }
          
          const [, startCol, startRowStr, endCol, endRowStr] = rangeMatch;
          const startRow = parseInt(startRowStr) - 1;
          const endRow = parseInt(endRowStr) - 1;
          const startColIndex = startCol.charCodeAt(0) - 65;
          const endColIndex = endCol.charCodeAt(0) - 65;
          
          const array: any[][] = [];
          for (let row = startRow; row <= endRow; row++) {
            const rowData: any[] = [];
            for (let col = startColIndex; col <= endColIndex; col++) {
              const cell = context.data[row]?.[col];
              rowData.push(cell ? cell.value : null);
            }
            array.push(rowData);
          }
          
          arrays.push(array);
          maxCols = Math.max(maxCols, array[0]?.length ?? 0);
        } else {
          // 単一セルの場合
          const value = getCellValue(arg, context);
          arrays.push([[value]]);
          maxCols = Math.max(maxCols, 1);
        }
      }
      
      // 垂直結合
      const result: any[][] = [];
      for (const array of arrays) {
        for (const row of array) {
          const resultRow = [...row];
          // 短い行の場合、#N/Aで埋める
          while (resultRow.length < maxCols) {
            resultRow.push(FormulaError.NA);
          }
          result.push(resultRow);
        }
      }
      
      return result;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// TOCOL関数の実装（列に変換）
export const TOCOL: CustomFormula = {
  name: 'TOCOL',
  pattern: /TOCOL\(([^,)]+)(?:,\s*([^,)]+))?(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, arrayRef, ignoreRef, scanByColRef] = matches;
    
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
      
      const ignore = ignoreRef ? parseInt(getCellValue(ignoreRef.trim(), context)?.toString() ?? ignoreRef.trim()) : 0;
      const scanByCol = scanByColRef ? getCellValue(scanByColRef.trim(), context)?.toString().toLowerCase() === 'true' : false;
      
      const result: any[] = [];
      
      // スキャン方向に応じて要素を収集
      if (scanByCol) {
        // 列ごとにスキャン
        for (let col = 0; col < (data[0]?.length ?? 0); col++) {
          for (let row = 0; row < data.length; row++) {
            const value = data[row][col];
            if (shouldIncludeValue(value, ignore)) {
              result.push(value);
            }
          }
        }
      } else {
        // 行ごとにスキャン（デフォルト）
        for (const row of data) {
          for (const value of row) {
            if (shouldIncludeValue(value, ignore)) {
              result.push(value);
            }
          }
        }
      }
      
      // 単一列の2次元配列として返す
      return result.map(val => [val]);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// TOROW関数の実装（行に変換）
export const TOROW: CustomFormula = {
  name: 'TOROW',
  pattern: /TOROW\(([^,)]+)(?:,\s*([^,)]+))?(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, arrayRef, ignoreRef, scanByColRef] = matches;
    
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
      
      const ignore = ignoreRef ? parseInt(getCellValue(ignoreRef.trim(), context)?.toString() ?? ignoreRef.trim()) : 0;
      const scanByCol = scanByColRef ? getCellValue(scanByColRef.trim(), context)?.toString().toLowerCase() === 'true' : false;
      
      const result: any[] = [];
      
      // スキャン方向に応じて要素を収集
      if (scanByCol) {
        // 列ごとにスキャン
        for (let col = 0; col < (data[0]?.length ?? 0); col++) {
          for (let row = 0; row < data.length; row++) {
            const value = data[row][col];
            if (shouldIncludeValue(value, ignore)) {
              result.push(value);
            }
          }
        }
      } else {
        // 行ごとにスキャン（デフォルト）
        for (const row of data) {
          for (const value of row) {
            if (shouldIncludeValue(value, ignore)) {
              result.push(value);
            }
          }
        }
      }
      
      // 単一行の2次元配列として返す
      return [result];
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// WRAPROWS関数の実装（行で折り返し）
export const WRAPROWS: CustomFormula = {
  name: 'WRAPROWS',
  pattern: /WRAPROWS\(([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, vectorRef, wrapCountRef, padRef] = matches;
    
    try {
      // ベクトルデータを取得
      const values: any[] = [];
      
      if (vectorRef.includes(':')) {
        const rangeMatch = vectorRef.trim().match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
        if (!rangeMatch) {
          return FormulaError.REF;
        }
        
        const [, startCol, startRowStr, endCol, endRowStr] = rangeMatch;
        const startRow = parseInt(startRowStr) - 1;
        const endRow = parseInt(endRowStr) - 1;
        const startColIndex = startCol.charCodeAt(0) - 65;
        const endColIndex = endCol.charCodeAt(0) - 65;
        
        // 範囲内のすべての値を収集
        for (let row = startRow; row <= endRow; row++) {
          for (let col = startColIndex; col <= endColIndex; col++) {
            const cell = context.data[row]?.[col];
            values.push(cell ? cell.value : null);
          }
        }
      } else {
        values.push(getCellValue(vectorRef.trim(), context));
      }
      
      const wrapCount = parseInt(getCellValue(wrapCountRef.trim(), context)?.toString() ?? wrapCountRef.trim());
      const padValue = padRef ? getCellValue(padRef.trim(), context) ?? padRef.trim() : FormulaError.NA;
      
      if (isNaN(wrapCount) || wrapCount <= 0) {
        return FormulaError.VALUE;
      }
      
      // 行で折り返し
      const result: any[][] = [];
      for (let i = 0; i < values.length; i += wrapCount) {
        const row: any[] = [];
        for (let j = 0; j < wrapCount; j++) {
          if (i + j < values.length) {
            row.push(values[i + j]);
          } else {
            row.push(padValue);
          }
        }
        result.push(row);
      }
      
      return result;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// WRAPCOLS関数の実装（列で折り返し）
export const WRAPCOLS: CustomFormula = {
  name: 'WRAPCOLS',
  pattern: /WRAPCOLS\(([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, vectorRef, wrapCountRef, padRef] = matches;
    
    try {
      // ベクトルデータを取得
      const values: any[] = [];
      
      if (vectorRef.includes(':')) {
        const rangeMatch = vectorRef.trim().match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
        if (!rangeMatch) {
          return FormulaError.REF;
        }
        
        const [, startCol, startRowStr, endCol, endRowStr] = rangeMatch;
        const startRow = parseInt(startRowStr) - 1;
        const endRow = parseInt(endRowStr) - 1;
        const startColIndex = startCol.charCodeAt(0) - 65;
        const endColIndex = endCol.charCodeAt(0) - 65;
        
        // 範囲内のすべての値を収集
        for (let row = startRow; row <= endRow; row++) {
          for (let col = startColIndex; col <= endColIndex; col++) {
            const cell = context.data[row]?.[col];
            values.push(cell ? cell.value : null);
          }
        }
      } else {
        values.push(getCellValue(vectorRef.trim(), context));
      }
      
      const wrapCount = parseInt(getCellValue(wrapCountRef.trim(), context)?.toString() ?? wrapCountRef.trim());
      const padValue = padRef ? getCellValue(padRef.trim(), context) ?? padRef.trim() : FormulaError.NA;
      
      if (isNaN(wrapCount) || wrapCount <= 0) {
        return FormulaError.VALUE;
      }
      
      // 列で折り返し（転置が必要）
      const numCols = Math.ceil(values.length / wrapCount);
      const result: any[][] = Array(wrapCount).fill(null).map(() => []);
      
      for (let i = 0; i < values.length; i++) {
        const row = i % wrapCount;
        result[row].push(values[i]);
      }
      
      // パディング
      for (let row = 0; row < wrapCount; row++) {
        while (result[row].length < numCols) {
          result[row].push(padValue);
        }
      }
      
      return result;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// ヘルパー関数：値を含めるかどうかを判定
function shouldIncludeValue(value: any, ignore: number): boolean {
  switch (ignore) {
    case 0: // すべて含む
      return true;
    case 1: // 空白を無視
      return value !== null && value !== undefined && value !== '';
    case 2: // エラーを無視
      return typeof value !== 'string' || !value.startsWith('#');
    case 3: // 空白とエラーを無視
      return value !== null && value !== undefined && value !== '' && 
             (typeof value !== 'string' || !value.startsWith('#'));
    default:
      return true;
  }
}