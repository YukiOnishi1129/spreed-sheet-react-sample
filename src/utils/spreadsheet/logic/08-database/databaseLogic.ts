// データベース関数の実装

import type { CustomFormula, FormulaContext, FormulaResult } from '../shared/types';
import { FormulaError } from '../shared/types';
import { getCellValue } from '../shared/utils';

// Convert column letter to index (A=0, B=1, ... Z=25, AA=26, etc.)
function columnToIndex(column: string): number {
  let index = 0;
  for (let i = 0; i < column.length; i++) {
    index = index * 26 + (column.charCodeAt(i) - 64);
  }
  return index - 1;
}

// Convert index to column letter (0=A, 1=B, ... 25=Z, 26=AA, etc.)
function indexToColumn(index: number): string {
  let column = '';
  index++; // Convert to 1-based
  while (index > 0) {
    const remainder = (index - 1) % 26;
    column = String.fromCharCode(65 + remainder) + column;
    index = Math.floor((index - 1) / 26);
  }
  return column;
}

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
  
  const startColIndex = columnToIndex(startCol);
  const endColIndex = columnToIndex(endCol);

  const result: unknown[][] = [];
  
  for (let row = startRow; row <= endRow; row++) {
    const rowData: unknown[] = [];
    for (let col = startColIndex; col <= endColIndex; col++) {
      const colName = indexToColumn(col);
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
    // Remove quotes if present
    const cleanField = field.replace(/^["']|["']$/g, '');
    const index = headers.findIndex(header => 
      String(header).toLowerCase() === cleanField.toLowerCase()
    );
    return index;
  }
  
  return -1;
}

// 条件に一致するかチェックするヘルパー関数
function matchesCriteria(row: unknown[], headers: unknown[], criteria: unknown[][]): boolean {
  if (criteria.length < 2) return true; // 条件がない場合は全て一致
  
  const criteriaHeaders = criteria[0];
  
  // For each criteria row (OR condition)
  for (let criteriaRow = 1; criteriaRow < criteria.length; criteriaRow++) {
    let matchesAllColumns = true;
    let hasAnyCriteriaInRow = false;
    
    // For each criteria column (AND condition within a row)
    for (let criteriaCol = 0; criteriaCol < criteriaHeaders.length; criteriaCol++) {
      const criteriaHeader = criteriaHeaders?.[criteriaCol];
      const criteriaValue = criteria[criteriaRow]?.[criteriaCol];
      
      // Skip empty criteria
      if (!criteriaHeader || criteriaValue === null || criteriaValue === undefined || criteriaValue === '') {
        continue;
      }
      
      hasAnyCriteriaInRow = true;
      const fieldIndex = getFieldIndex(criteriaHeader, headers);
      if (fieldIndex === -1) {
        matchesAllColumns = false;
        break;
      }
      
      const cellValue = row[fieldIndex] as unknown;
      const criteriaStr = String(criteriaValue);
      
      let matches = false;
      
      // 条件文字列の解析
      if (criteriaStr.startsWith('>=')) {
        const value = parseFloat(criteriaStr.substring(2));
        if (!isNaN(value) && Number(cellValue) >= value) {
          matches = true;
        }
      } else if (criteriaStr.startsWith('<=')) {
        const value = parseFloat(criteriaStr.substring(2));
        if (!isNaN(value) && Number(cellValue) <= value) {
          matches = true;
        }
      } else if (criteriaStr.startsWith('>')) {
        const value = parseFloat(criteriaStr.substring(1));
        if (!isNaN(value) && Number(cellValue) > value) {
          matches = true;
        }
      } else if (criteriaStr.startsWith('<')) {
        const value = parseFloat(criteriaStr.substring(1));
        if (!isNaN(value) && Number(cellValue) < value) {
          matches = true;
        }
      } else if (criteriaStr.startsWith('=')) {
        const value = criteriaStr.substring(1);
        if (String(cellValue) === value) {
          matches = true;
        }
      } else if (criteriaStr.includes('*') || criteriaStr.includes('?')) {
        // ワイルドカード一致
        const regex = new RegExp(
          '^' + criteriaStr.replace(/\*/g, '.*').replace(/\?/g, '.') + '$',
          'i'
        );
        if (regex.test(String(cellValue))) {
          matches = true;
        }
      } else {
        // 完全一致
        if (String(cellValue).toLowerCase() === criteriaStr.toLowerCase()) {
          matches = true;
        }
      }
      
      if (!matches) {
        matchesAllColumns = false;
        break;
      }
    }
    
    // If all columns in this row match and the row has criteria, the record matches
    if (matchesAllColumns && hasAnyCriteriaInRow) {
      return true;
    }
  }
  
  return false;
}

// DSUM関数（条件付き合計）
export const DSUM: CustomFormula = {
  name: 'DSUM',
  pattern: /\bDSUM\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, databaseRef, fieldRef, criteriaRef] = matches;
    
    try {
      const database = getDatabaseArray(databaseRef.trim(), context);
      const criteria = getDatabaseArray(criteriaRef.trim(), context);
      
      if (database.length < 2 || criteria.length < 1) {
        return FormulaError.VALUE;
      }
      
      const headers = database[0];
      // Handle field reference - could be a cell reference or a direct field name/number
      let field = fieldRef.trim();
      // Check if field is a number
      const fieldNum = Number(field);
      if (!isNaN(fieldNum)) {
        field = fieldNum;
      } else if (/^[A-Z]+\d+$/.test(field)) {
        // If field looks like a cell reference (e.g., B1), get its value
        field = getCellValue(field, context) ?? field;
      }
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
  pattern: /\bDAVERAGE\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, databaseRef, fieldRef, criteriaRef] = matches;
    
    try {
      const database = getDatabaseArray(databaseRef.trim(), context);
      const criteria = getDatabaseArray(criteriaRef.trim(), context);
      
      if (database.length < 2 || criteria.length < 1) {
        return FormulaError.VALUE;
      }
      
      const headers = database[0];
      // Handle field reference - could be a cell reference or a direct field name/number
      let field = fieldRef.trim();
      // Check if field is a number
      const fieldNum = Number(field);
      if (!isNaN(fieldNum)) {
        field = fieldNum;
      } else if (/^[A-Z]+\d+$/.test(field)) {
        // If field looks like a cell reference (e.g., B1), get its value
        field = getCellValue(field, context) ?? field;
      }
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
  pattern: /\bDCOUNT\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, databaseRef, fieldRef, criteriaRef] = matches;
    
    try {
      const database = getDatabaseArray(databaseRef.trim(), context);
      const criteria = getDatabaseArray(criteriaRef.trim(), context);
      
      if (database.length < 2 || criteria.length < 1) {
        return FormulaError.VALUE;
      }
      
      const headers = database[0];
      // Handle field reference - could be a cell reference or a direct field name/number
      let field = fieldRef.trim();
      // Check if field is a number
      const fieldNum = Number(field);
      if (!isNaN(fieldNum)) {
        field = fieldNum;
      } else if (/^[A-Z]+\d+$/.test(field)) {
        // If field looks like a cell reference (e.g., B1), get its value
        field = getCellValue(field, context) ?? field;
      }
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
  pattern: /\bDCOUNTA\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, databaseRef, fieldRef, criteriaRef] = matches;
    
    try {
      const database = getDatabaseArray(databaseRef.trim(), context);
      const criteria = getDatabaseArray(criteriaRef.trim(), context);
      
      if (database.length < 2 || criteria.length < 1) {
        return FormulaError.VALUE;
      }
      
      const headers = database[0];
      // Handle field reference - could be a cell reference or a direct field name/number
      let field = fieldRef.trim();
      // Check if field is a number
      const fieldNum = Number(field);
      if (!isNaN(fieldNum)) {
        field = fieldNum;
      } else if (/^[A-Z]+\d+$/.test(field)) {
        // If field looks like a cell reference (e.g., B1), get its value
        field = getCellValue(field, context) ?? field;
      }
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
  pattern: /\bDMAX\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, databaseRef, fieldRef, criteriaRef] = matches;
    
    try {
      const database = getDatabaseArray(databaseRef.trim(), context);
      const criteria = getDatabaseArray(criteriaRef.trim(), context);
      
      if (database.length < 2 || criteria.length < 1) {
        return FormulaError.VALUE;
      }
      
      const headers = database[0];
      // Handle field reference - could be a cell reference or a direct field name/number
      let field = fieldRef.trim();
      // Check if field is a number
      const fieldNum = Number(field);
      if (!isNaN(fieldNum)) {
        field = fieldNum;
      } else if (/^[A-Z]+\d+$/.test(field)) {
        // If field looks like a cell reference (e.g., B1), get its value
        field = getCellValue(field, context) ?? field;
      }
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
  pattern: /\bDMIN\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, databaseRef, fieldRef, criteriaRef] = matches;
    
    try {
      const database = getDatabaseArray(databaseRef.trim(), context);
      const criteria = getDatabaseArray(criteriaRef.trim(), context);
      
      if (database.length < 2 || criteria.length < 1) {
        return FormulaError.VALUE;
      }
      
      const headers = database[0];
      // Handle field reference - could be a cell reference or a direct field name/number
      let field = fieldRef.trim();
      // Check if field is a number
      const fieldNum = Number(field);
      if (!isNaN(fieldNum)) {
        field = fieldNum;
      } else if (/^[A-Z]+\d+$/.test(field)) {
        // If field looks like a cell reference (e.g., B1), get its value
        field = getCellValue(field, context) ?? field;
      }
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
  pattern: /\bDPRODUCT\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, databaseRef, fieldRef, criteriaRef] = matches;
    
    try {
      const database = getDatabaseArray(databaseRef.trim(), context);
      const criteria = getDatabaseArray(criteriaRef.trim(), context);
      
      if (database.length < 2 || criteria.length < 1) {
        return FormulaError.VALUE;
      }
      
      const headers = database[0];
      // Handle field reference - could be a cell reference or a direct field name/number
      let field = fieldRef.trim();
      // Check if field is a number
      const fieldNum = Number(field);
      if (!isNaN(fieldNum)) {
        field = fieldNum;
      } else if (/^[A-Z]+\d+$/.test(field)) {
        // If field looks like a cell reference (e.g., B1), get its value
        field = getCellValue(field, context) ?? field;
      }
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
  pattern: /\bDGET\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, databaseRef, fieldRef, criteriaRef] = matches;
    
    try {
      const database = getDatabaseArray(databaseRef.trim(), context);
      const criteria = getDatabaseArray(criteriaRef.trim(), context);
      
      if (database.length < 2 || criteria.length < 1) {
        return FormulaError.VALUE;
      }
      
      const headers = database[0];
      // Handle field reference - could be a cell reference or a direct field name/number
      let field = fieldRef.trim();
      // Check if field is a number
      const fieldNum = Number(field);
      if (!isNaN(fieldNum)) {
        field = fieldNum;
      } else if (/^[A-Z]+\d+$/.test(field)) {
        // If field looks like a cell reference (e.g., B1), get its value
        field = getCellValue(field, context) ?? field;
      }
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
        return FormulaError.NUM; // No matching records
      } else if (matchedValues.length === 1) {
        return matchedValues[0] as FormulaResult;
      } else {
        return FormulaError.NUM; // Multiple matching records
      }
    } catch {
      return FormulaError.VALUE;
    }
  }
};