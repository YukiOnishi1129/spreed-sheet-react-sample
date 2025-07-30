// 関数処理の共通ユーティリティ

import type { FormulaContext } from './types';
import { FormulaError } from './types';
import { parseDate as parseDateNew } from './dateUtils';

// セル参照（A1形式）から座標を取得
export const parseCellReference = (cellRef: string): { col: number; row: number } | null => {
  // Remove $ signs for absolute references
  const cleanRef = cellRef.replace(/\$/g, '');
  const match = cleanRef.match(/^([A-Z]+)(\d+)$/);
  if (!match) return null;
  
  // 列番号を計算（A=0, B=1, ..., Z=25, AA=26, ...）
  let col = 0;
  for (let i = 0; i < match[1].length; i++) {
    col = col * 26 + (match[1].charCodeAt(i) - 64);
  }
  col--; // 0ベースに変換
  
  const row = parseInt(match[2]) - 1; // 0ベースに変換
  
  return { col, row };
};

// セル範囲（A1:B2形式）から座標範囲を取得
export const parseCellRange = (range: string): { start: { col: number; row: number }; end: { col: number; row: number } } | null => {
  // Remove $ signs for absolute references
  const cleanRange = range.replace(/\$/g, '');
  
  // Check for full column references (e.g., D:E)
  const columnMatch = cleanRange.match(/^([A-Z]+):([A-Z]+)$/);
  if (columnMatch) {
    const startCol = columnMatch[1].charCodeAt(0) - 'A'.charCodeAt(0);
    const endCol = columnMatch[2].charCodeAt(0) - 'A'.charCodeAt(0);
    // Use a reasonable default range for full column references
    return { 
      start: { col: startCol, row: 0 }, 
      end: { col: endCol, row: 999 } // Reasonable max rows
    };
  }
  
  // Standard range (e.g., A1:B2)
  const match = cleanRange.match(/^([A-Z]+\d+):([A-Z]+\d+)$/);
  if (!match) return null;
  
  const start = parseCellReference(match[1]);
  const end = parseCellReference(match[2]);
  
  if (!start || !end) return null;
  
  
  return { start, end };
};

// セル参照から値を取得
export const getCellValue = (cellRef: string, context: FormulaContext): unknown => {
  // If cellRef is not a string or is empty, try to return it as a direct value
  if (!cellRef || typeof cellRef !== 'string') {
    return cellRef;
  }
  
  const coords = parseCellReference(cellRef);
  if (!coords) {
    // If it's not a valid cell reference, try to parse it as a direct value
    const num = Number(cellRef);
    if (!isNaN(num)) {
      return num;
    }
    // Return the string value itself if it's not a cell reference
    return cellRef;
  }
  
  const { col, row } = coords;
  if (row < 0 || row >= context.data.length || col < 0 || col >= context.data[0]?.length) {
    return FormulaError.REF;
  }
  
  const cellData = context.data[row][col];
  
  
  // 配列データの場合、セルは直接値として格納されている
  // 文字列、数値、null、undefinedの場合はそのまま返す
  if (typeof cellData === 'string' || typeof cellData === 'number' || cellData === null || cellData === undefined) {
    return cellData;
  }
  
  // セルデータがオブジェクトの場合のみプロパティを探す
  if (cellData && typeof cellData === 'object') {
    // If the object itself has a numeric or string representation, return it
    // Check for value property (handling null as a valid value)
    if ('value' in cellData) {
      return cellData.value;
    }
    if ('v' in cellData) {
      return cellData.v;
    }
    if ('_' in cellData) {
      return cellData._;
    }
    // If we have a formula cell that hasn't been calculated yet, return empty
    if ('formula' in cellData || 'f' in cellData) {
      return '';
    }
    return cellData;
  }
  return cellData;
};

// セル範囲から値の配列を取得
export const getCellRangeValues = (range: string, context: FormulaContext): unknown[] => {
  const rangeCoords = parseCellRange(range);
  if (!rangeCoords) {
    // Check if it's a single cell reference
    const singleCell = parseCellReference(range);
    if (singleCell) {
      const value = getCellValue(range, context);
      return [value];
    }
    return [];
  }
  
  const { start, end } = rangeCoords;
  const values: unknown[] = [];
  
  
  
  for (let row = start.row; row <= end.row; row++) {
    for (let col = start.col; col <= end.col; col++) {
      if (row >= 0 && row < context.data.length && col >= 0 && col < context.data[0]?.length) {
        const cellData = context.data[row][col];
        let value: unknown;
        
        
        // Handle both direct values and object values
        if (typeof cellData === 'string' || typeof cellData === 'number' || cellData === null || cellData === undefined) {
          value = cellData;
        } else if (cellData && typeof cellData === 'object') {
          // Check for value property (handling null as a valid value)
          if ('value' in cellData) {
            value = cellData.value;
          } else if ('v' in cellData) {
            value = cellData.v;
          } else if ('_' in cellData) {
            value = cellData._;
          } else if ('formula' in cellData || 'f' in cellData) {
            // Skip formula cells that haven't been calculated yet
            continue;
          } else {
            value = cellData;
          }
        } else {
          value = cellData;
        }
        
        // Always push the value, even if it's empty (needed for COUNTBLANK)
        values.push(value);
      }
    }
  }
  
  
  return values;
};

// 日付文字列をDateオブジェクトに変換（互換性のため）
export const parseDate = (dateValue: unknown): Date | null => {
  return parseDateNew(String(dateValue));
};

// Dateオブジェクトをそのまま返す（互換性のため）
export const toDate = (dateObj: Date | null): Date | null => {
  return dateObj;
};

// 値を数値に変換
export const toNumber = (value: unknown): number | null => {
  if (typeof value === 'number') return value;
  if (typeof value === 'string') {
    const num = parseFloat(value);
    if (!isNaN(num)) return num;
  }
  if (value instanceof Date) {
    return excelDateSerial(value);
  }
  return null;
};

// DateオブジェクトをExcelシリアル値に変換
export const excelDateSerial = (date: Date): number => {
  const excelEpoch = new Date(1899, 11, 30);
  return Math.floor((date.getTime() - excelEpoch.getTime()) / (1000 * 60 * 60 * 24));
};

// 文字列の引用符を除去
export const unquoteString = (str: string): string => {
  if (str.startsWith('"') && str.endsWith('"')) {
    return str.slice(1, -1).replace(/\\"/g, '"');
  }
  return str;
};

// 簡単な算術式を評価する関数
export const evaluateExpression = (expression: string, context: FormulaContext): number | string => {
  try {
    // セル参照を値に置換
    const cellRefPattern = /[A-Z]+\d+/g;
    let evaluatedExpr = expression;
    
    const cellRefs = expression.match(cellRefPattern);
    if (cellRefs) {
      for (const cellRef of cellRefs) {
        const value = getCellValue(cellRef, context);
        if (typeof value === 'number') {
          evaluatedExpr = evaluatedExpr.replace(cellRef, value.toString());
        } else if (value !== null && value !== undefined && value !== '') {
          // 数値でない場合は文字列として扱う
          return expression; // 元の式を返す
        }
      }
    }
    
    // 基本的な算術演算子のみを許可（セキュリティのため）
    if (!/^[\d\s+\-*/().,]+$/.test(evaluatedExpr)) {
      return expression; // 安全でない文字が含まれている場合は元の式を返す
    }
    
    // 算術式を評価
    const result = Function('"use strict"; return (' + evaluatedExpr + ')')();
    return typeof result === 'number' ? result : expression;
  } catch (error) {
    // エラーの場合は元の式を返す
    return expression;
  }
};

// 式を評価してエラーをチェックする関数
export const evaluateFormulaWithErrorCheck = (expression: string, context: FormulaContext): string | number => {
  try {
    // セル参照を値に置換
    const cellRefPattern = /[A-Z]+\d+/g;
    let evaluatedExpr = expression;
    const replacements: Array<{ref: string, value: number}> = [];
    
    const cellRefs = expression.match(cellRefPattern);
    if (cellRefs) {
      for (const cellRef of cellRefs) {
        const value = getCellValue(cellRef, context);
        if (typeof value === 'number') {
          replacements.push({ref: cellRef, value});
          evaluatedExpr = evaluatedExpr.replace(cellRef, value.toString());
        } else if (value === null || value === undefined || value === '') {
          evaluatedExpr = evaluatedExpr.replace(cellRef, '0');
          replacements.push({ref: cellRef, value: 0});
        } else {
          // 数値でない場合はVALUEエラー
          return '#VALUE!';
        }
      }
    }
    
    // 除算をチェック
    if (evaluatedExpr.includes('/')) {
      // 除算部分を抽出してゼロ除算をチェック
      const divisionParts = evaluatedExpr.split('/');
      for (let i = 1; i < divisionParts.length; i++) {
        // 次の演算子までの部分を取得
        const divisorMatch = divisionParts[i].match(/^[\s]*(\d+(?:\.\d+)?)/);
        if (divisorMatch && parseFloat(divisorMatch[1]) === 0) {
          return '#DIV/0!';
        }
      }
    }
    
    // 基本的な算術演算子のみを許可（セキュリティのため）
    if (!/^[\d\s+\-*/().,]+$/.test(evaluatedExpr)) {
      return '#NAME?'; // 安全でない文字が含まれている場合はNAMEエラー
    }
    
    // 算術式を評価
    const result = Function('"use strict"; return (' + evaluatedExpr + ')')();
    
    // 無限大やNaNのチェック
    if (!isFinite(result)) {
      return '#DIV/0!';
    }
    if (isNaN(result)) {
      return '#VALUE!';
    }
    
    return typeof result === 'number' ? result : '#VALUE!';
  } catch (error) {
    // エラーの場合は#VALUE!エラーを返す
    return '#VALUE!';
  }
};