// データベース関数の実装

import type { CustomFormula, FormulaContext, FormulaResult } from './types';
import { FormulaError } from './types';
import { getCellValue } from './utils';

// データベース範囲を2次元配列として取得するヘルパー関数
function getDatabaseArray(databaseRef: string, context: FormulaContext): unknown[][] {
  if (!databaseRef.includes(':')) {
    return [[getCellValue(databaseRef, context)]];
  }

  const [startCell, endCell] = databaseRef.split(':');
  const startMatch = startCell.match(/([A-Z]+)(\d+)/);
  const endMatch = endCell.match(/([A-Z]+)(\d+)/);
  
  if (!startMatch || !endMatch) {
    return [];
  }

  const [, startCol, startRowStr] = startMatch;
  const [, endCol, endRowStr] = endMatch;
  
  const startRow = parseInt(startRowStr);
  const endRow = parseInt(endRowStr);
  
  const startColIndex = startCol.split('').reduce((acc, char) => acc * 26 + char.charCodeAt(0) - 64, 0) - 1;
  const endColIndex = endCol.split('').reduce((acc, char) => acc * 26 + char.charCodeAt(0) - 64, 0) - 1;

  const result: unknown[][] = [];
  
  for (let row = startRow; row <= endRow; row++) {
    const rowData: unknown[] = [];
    for (let col = startColIndex; col <= endColIndex; col++) {
      const colName = String.fromCharCode(65 + col);
      const cellRef = `${colName}${row}`;
      rowData.push(getCellValue(cellRef, context));
    }
    result.push(rowData);
  }
  
  return result;
}

// フィールドインデックスを取得するヘルパー関数
function getFieldIndex(field: unknown, headers: unknown[]): number {
  if (typeof field === 'number') {
    return field - 1; // 1ベースから0ベースに変換
  }
  
  if (typeof field === 'string') {
    const index = headers.findIndex(header => 
      String(header).toLowerCase() === field.toLowerCase()
    );
    return index;
  }
  
  return -1;
}

// 条件に一致するかチェックするヘルパー関数
function matchesCriteria(row: unknown[], headers: unknown[], criteria: unknown[][]): boolean {
  if (criteria.length < 2) return true; // 条件がない場合は全て一致
  
  const criteriaHeaders = criteria[0];
  
  for (let criteriaCol = 0; criteriaCol < criteriaHeaders.length; criteriaCol++) {
    const criteriaHeader = criteriaHeaders?.[criteriaCol];
    const fieldIndex = getFieldIndex(criteriaHeader, headers);
    
    if (fieldIndex === -1) continue;
    
    // 各条件行をチェック
    let matchesAnyRow = false;
    for (let criteriaRow = 1; criteriaRow < criteria.length; criteriaRow++) {
      const criteriaValue = criteria[criteriaRow]?.[criteriaCol];
      if (criteriaValue === null || criteriaValue === undefined || criteriaValue === '') {
        continue;
      }
      
      const cellValue = row[fieldIndex] as unknown;
      
      // 条件文字列の解析
      const criteriaStr = String(criteriaValue);
      
      if (criteriaStr.startsWith('>=')) {
        const value = parseFloat(criteriaStr.substring(2));
        if (!isNaN(value) && Number(cellValue) >= value) {
          matchesAnyRow = true;
          break;
        }
      } else if (criteriaStr.startsWith('<=')) {
        const value = parseFloat(criteriaStr.substring(2));
        if (!isNaN(value) && Number(cellValue) <= value) {
          matchesAnyRow = true;
          break;
        }
      } else if (criteriaStr.startsWith('>')) {
        const value = parseFloat(criteriaStr.substring(1));
        if (!isNaN(value) && Number(cellValue) > value) {
          matchesAnyRow = true;
          break;
        }
      } else if (criteriaStr.startsWith('<')) {
        const value = parseFloat(criteriaStr.substring(1));
        if (!isNaN(value) && Number(cellValue) < value) {
          matchesAnyRow = true;
          break;
        }
      } else if (criteriaStr.startsWith('=')) {
        const value = criteriaStr.substring(1);
        if (String(cellValue) === value) {
          matchesAnyRow = true;
          break;
        }
      } else if (criteriaStr.includes('*') || criteriaStr.includes('?')) {
        // ワイルドカード一致
        const regex = new RegExp(
          criteriaStr.replace(/\*/g, '.*').replace(/\?/g, '.'),
          'i'
        );
        if (regex.test(String(cellValue))) {
          matchesAnyRow = true;
          break;
        }
      } else {
        // 完全一致
        if (String(cellValue) === criteriaStr) {
          matchesAnyRow = true;
          break;
        }
      }
    }
    
    if (!matchesAnyRow) {
      return false;
    }
  }
  
  return true;
}

// DSUM関数（条件付き合計）
export const DSUM: CustomFormula = {
  name: 'DSUM',
  pattern: /DSUM\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, databaseRef, fieldRef, criteriaRef] = matches;
    
    try {
      const database = getDatabaseArray(databaseRef.trim(), context);
      const criteria = getDatabaseArray(criteriaRef.trim(), context);
      
      if (database.length < 2 || criteria.length < 1) {
        return FormulaError.VALUE;
      }
      
      const headers = database[0];
      const field = getCellValue(fieldRef.trim(), context) ?? fieldRef.trim();
      const fieldIndex = getFieldIndex(field, headers);
      
      if (fieldIndex === -1) {
        return FormulaError.VALUE;
      }
      
      let sum = 0;
      for (let i = 1; i < database.length; i++) {
        const row = database[i];
        if (matchesCriteria(row, headers, criteria)) {
          const value = Number(row[fieldIndex]);
          if (!isNaN(value)) {
            sum += value;
          }
        }
      }
      
      return sum;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// DAVERAGE関数（条件付き平均）
export const DAVERAGE: CustomFormula = {
  name: 'DAVERAGE',
  pattern: /DAVERAGE\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, databaseRef, fieldRef, criteriaRef] = matches;
    
    try {
      const database = getDatabaseArray(databaseRef.trim(), context);
      const criteria = getDatabaseArray(criteriaRef.trim(), context);
      
      if (database.length < 2 || criteria.length < 1) {
        return FormulaError.VALUE;
      }
      
      const headers = database[0];
      const field = getCellValue(fieldRef.trim(), context) ?? fieldRef.trim();
      const fieldIndex = getFieldIndex(field, headers);
      
      if (fieldIndex === -1) {
        return FormulaError.VALUE;
      }
      
      let sum = 0;
      let count = 0;
      for (let i = 1; i < database.length; i++) {
        const row = database[i];
        if (matchesCriteria(row, headers, criteria)) {
          const value = Number(row[fieldIndex]);
          if (!isNaN(value)) {
            sum += value;
            count++;
          }
        }
      }
      
      return count === 0 ? FormulaError.DIV0 : sum / count;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// DCOUNT関数（条件付きカウント）
export const DCOUNT: CustomFormula = {
  name: 'DCOUNT',
  pattern: /DCOUNT\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, databaseRef, fieldRef, criteriaRef] = matches;
    
    try {
      const database = getDatabaseArray(databaseRef.trim(), context);
      const criteria = getDatabaseArray(criteriaRef.trim(), context);
      
      if (database.length < 2 || criteria.length < 1) {
        return FormulaError.VALUE;
      }
      
      const headers = database[0];
      const field = getCellValue(fieldRef.trim(), context) ?? fieldRef.trim();
      const fieldIndex = getFieldIndex(field, headers);
      
      if (fieldIndex === -1) {
        return FormulaError.VALUE;
      }
      
      let count = 0;
      for (let i = 1; i < database.length; i++) {
        const row = database[i];
        if (matchesCriteria(row, headers, criteria)) {
          const value = Number(row[fieldIndex]);
          if (!isNaN(value)) {
            count++;
          }
        }
      }
      
      return count;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// DCOUNTA関数（条件付きカウント、空白以外）
export const DCOUNTA: CustomFormula = {
  name: 'DCOUNTA',
  pattern: /DCOUNTA\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, databaseRef, fieldRef, criteriaRef] = matches;
    
    try {
      const database = getDatabaseArray(databaseRef.trim(), context);
      const criteria = getDatabaseArray(criteriaRef.trim(), context);
      
      if (database.length < 2 || criteria.length < 1) {
        return FormulaError.VALUE;
      }
      
      const headers = database[0];
      const field = getCellValue(fieldRef.trim(), context) ?? fieldRef.trim();
      const fieldIndex = getFieldIndex(field, headers);
      
      if (fieldIndex === -1) {
        return FormulaError.VALUE;
      }
      
      let count = 0;
      for (let i = 1; i < database.length; i++) {
        const row = database[i];
        if (matchesCriteria(row, headers, criteria)) {
          const value = row[fieldIndex] as unknown;
          if (value !== null && value !== undefined && value !== '') {
            count++;
          }
        }
      }
      
      return count;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// DMAX関数（条件付き最大値）
export const DMAX: CustomFormula = {
  name: 'DMAX',
  pattern: /DMAX\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, databaseRef, fieldRef, criteriaRef] = matches;
    
    try {
      const database = getDatabaseArray(databaseRef.trim(), context);
      const criteria = getDatabaseArray(criteriaRef.trim(), context);
      
      if (database.length < 2 || criteria.length < 1) {
        return FormulaError.VALUE;
      }
      
      const headers = database[0];
      const field = getCellValue(fieldRef.trim(), context) ?? fieldRef.trim();
      const fieldIndex = getFieldIndex(field, headers);
      
      if (fieldIndex === -1) {
        return FormulaError.VALUE;
      }
      
      let max = -Infinity;
      let hasValue = false;
      for (let i = 1; i < database.length; i++) {
        const row = database[i];
        if (matchesCriteria(row, headers, criteria)) {
          const value = Number(row[fieldIndex]);
          if (!isNaN(value)) {
            max = Math.max(max, value);
            hasValue = true;
          }
        }
      }
      
      return hasValue ? max : 0;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// DMIN関数（条件付き最小値）
export const DMIN: CustomFormula = {
  name: 'DMIN',
  pattern: /DMIN\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, databaseRef, fieldRef, criteriaRef] = matches;
    
    try {
      const database = getDatabaseArray(databaseRef.trim(), context);
      const criteria = getDatabaseArray(criteriaRef.trim(), context);
      
      if (database.length < 2 || criteria.length < 1) {
        return FormulaError.VALUE;
      }
      
      const headers = database[0];
      const field = getCellValue(fieldRef.trim(), context) ?? fieldRef.trim();
      const fieldIndex = getFieldIndex(field, headers);
      
      if (fieldIndex === -1) {
        return FormulaError.VALUE;
      }
      
      let min = Infinity;
      let hasValue = false;
      for (let i = 1; i < database.length; i++) {
        const row = database[i];
        if (matchesCriteria(row, headers, criteria)) {
          const value = Number(row[fieldIndex]);
          if (!isNaN(value)) {
            min = Math.min(min, value);
            hasValue = true;
          }
        }
      }
      
      return hasValue ? min : 0;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// DPRODUCT関数（条件付き積）
export const DPRODUCT: CustomFormula = {
  name: 'DPRODUCT',
  pattern: /DPRODUCT\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, databaseRef, fieldRef, criteriaRef] = matches;
    
    try {
      const database = getDatabaseArray(databaseRef.trim(), context);
      const criteria = getDatabaseArray(criteriaRef.trim(), context);
      
      if (database.length < 2 || criteria.length < 1) {
        return FormulaError.VALUE;
      }
      
      const headers = database[0];
      const field = getCellValue(fieldRef.trim(), context) ?? fieldRef.trim();
      const fieldIndex = getFieldIndex(field, headers);
      
      if (fieldIndex === -1) {
        return FormulaError.VALUE;
      }
      
      let product = 1;
      let hasValue = false;
      for (let i = 1; i < database.length; i++) {
        const row = database[i];
        if (matchesCriteria(row, headers, criteria)) {
          const value = Number(row[fieldIndex]);
          if (!isNaN(value)) {
            product *= value;
            hasValue = true;
          }
        }
      }
      
      return hasValue ? product : 0;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// DGET関数（条件に一致する値を取得）
export const DGET: CustomFormula = {
  name: 'DGET',
  pattern: /DGET\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, databaseRef, fieldRef, criteriaRef] = matches;
    
    try {
      const database = getDatabaseArray(databaseRef.trim(), context);
      const criteria = getDatabaseArray(criteriaRef.trim(), context);
      
      if (database.length < 2 || criteria.length < 1) {
        return FormulaError.VALUE;
      }
      
      const headers = database[0];
      const field = getCellValue(fieldRef.trim(), context) ?? fieldRef.trim();
      const fieldIndex = getFieldIndex(field, headers);
      
      if (fieldIndex === -1) {
        return FormulaError.VALUE;
      }
      
      const matchedValues: unknown[] = [];
      for (let i = 1; i < database.length; i++) {
        const row = database[i];
        if (matchesCriteria(row, headers, criteria)) {
          matchedValues.push(row[fieldIndex]);
        }
      }
      
      if (matchedValues.length === 0) {
        return FormulaError.VALUE;
      } else if (matchedValues.length === 1) {
        return matchedValues[0] as FormulaResult;
      } else {
        return FormulaError.NUM; // 複数の値が一致
      }
    } catch {
      return FormulaError.VALUE;
    }
  }
};