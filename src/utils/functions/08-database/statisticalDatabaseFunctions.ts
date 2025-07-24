// データベース統計関数

import type { CustomFormula, FormulaContext, FormulaResult } from '../shared/types';
import { FormulaError } from '../shared/types';

// データベース関数用のヘルパー：条件に一致する行を取得
function getMatchingRows(
  database: any[][],
  field: string | number,
  criteria: any[][]
): { values: number[], fieldIndex: number } {
  if (database.length < 2 || criteria.length < 2) {
    return { values: [], fieldIndex: -1 };
  }

  // ヘッダー行を取得
  const headers = database[0].map(h => String(h).toUpperCase());
  const criteriaHeaders = criteria[0].map(h => String(h).toUpperCase());

  // フィールドのインデックスを特定
  let fieldIndex = -1;
  if (typeof field === 'number') {
    fieldIndex = field - 1;
  } else {
    const fieldName = String(field).toUpperCase();
    fieldIndex = headers.indexOf(fieldName);
  }

  if (fieldIndex < 0 || fieldIndex >= headers.length) {
    return { values: [], fieldIndex: -1 };
  }

  // 条件に一致する行の値を収集
  const values: number[] = [];

  for (let row = 1; row < database.length; row++) {
    let match = true;

    // すべての条件をチェック
    for (let cCol = 0; cCol < criteriaHeaders.length; cCol++) {
      const criteriaHeader = criteriaHeaders[cCol];
      const dbCol = headers.indexOf(criteriaHeader);

      if (dbCol >= 0) {
        for (let cRow = 1; cRow < criteria.length; cRow++) {
          const criteriaValue = criteria[cRow][cCol];
          if (criteriaValue !== null && criteriaValue !== undefined && criteriaValue !== '') {
            const dbValue = database[row][dbCol];
            
            // 条件の評価
            const criteriaStr = String(criteriaValue);
            if (criteriaStr.startsWith('>') || criteriaStr.startsWith('<') || criteriaStr.startsWith('=') || criteriaStr.startsWith('!')) {
              // 比較演算子を含む条件
              const operator = criteriaStr.match(/^([><=!]+)/)?.[1] || '';
              const compareValue = criteriaStr.substring(operator.length);
              const dbNum = parseFloat(String(dbValue));
              const compareNum = parseFloat(compareValue);

              if (!isNaN(dbNum) && !isNaN(compareNum)) {
                switch (operator) {
                  case '>':
                    match = dbNum > compareNum;
                    break;
                  case '>=':
                    match = dbNum >= compareNum;
                    break;
                  case '<':
                    match = dbNum < compareNum;
                    break;
                  case '<=':
                    match = dbNum <= compareNum;
                    break;
                  case '<>':
                  case '!=':
                    match = dbNum !== compareNum;
                    break;
                  case '=':
                    match = dbNum === compareNum;
                    break;
                  default:
                    match = String(dbValue).toUpperCase() === criteriaStr.toUpperCase();
                }
              } else {
                match = String(dbValue).toUpperCase() === criteriaStr.toUpperCase();
              }
            } else {
              // 完全一致
              match = String(dbValue).toUpperCase() === criteriaStr.toUpperCase();
            }

            if (!match) break;
          }
        }
      }
      if (!match) break;
    }

    if (match) {
      const value = parseFloat(String(database[row][fieldIndex]));
      if (!isNaN(value)) {
        values.push(value);
      }
    }
  }

  return { values, fieldIndex };
}

// DSTDEV関数の実装（条件付き標準偏差：標本）
export const DSTDEV: CustomFormula = {
  name: 'DSTDEV',
  pattern: /DSTDEV\(([^,]+),\s*"?([^",]+)"?,\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, databaseRef, field, criteriaRef] = matches;
    
    try {
      // データベースと条件範囲を解析
      const dbRange = databaseRef.trim().match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
      const critRange = criteriaRef.trim().match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
      
      if (!dbRange || !critRange) {
        return FormulaError.REF;
      }
      
      // データベース配列を構築
      const database: any[][] = [];
      const [, dbStartCol, dbStartRowStr, dbEndCol, dbEndRowStr] = dbRange;
      const dbStartRow = parseInt(dbStartRowStr) - 1;
      const dbEndRow = parseInt(dbEndRowStr) - 1;
      const dbStartColIndex = dbStartCol.charCodeAt(0) - 65;
      const dbEndColIndex = dbEndCol.charCodeAt(0) - 65;
      
      for (let row = dbStartRow; row <= dbEndRow; row++) {
        const rowData: any[] = [];
        for (let col = dbStartColIndex; col <= dbEndColIndex; col++) {
          const cell = context.data[row]?.[col];
          rowData.push(cell ? cell.value : null);
        }
        database.push(rowData);
      }
      
      // 条件配列を構築
      const criteria: any[][] = [];
      const [, critStartCol, critStartRowStr, critEndCol, critEndRowStr] = critRange;
      const critStartRow = parseInt(critStartRowStr) - 1;
      const critEndRow = parseInt(critEndRowStr) - 1;
      const critStartColIndex = critStartCol.charCodeAt(0) - 65;
      const critEndColIndex = critEndCol.charCodeAt(0) - 65;
      
      for (let row = critStartRow; row <= critEndRow; row++) {
        const rowData: any[] = [];
        for (let col = critStartColIndex; col <= critEndColIndex; col++) {
          const cell = context.data[row]?.[col];
          rowData.push(cell ? cell.value : null);
        }
        criteria.push(rowData);
      }
      
      // 条件に一致する値を取得
      const { values } = getMatchingRows(database, field, criteria);
      
      if (values.length < 2) {
        return FormulaError.DIV0;
      }
      
      // 標準偏差（標本）を計算
      const mean = values.reduce((sum, val) => sum + val, 0) / values.length;
      const squaredDiffs = values.map(val => Math.pow(val - mean, 2));
      const variance = squaredDiffs.reduce((sum, val) => sum + val, 0) / (values.length - 1);
      
      return Math.sqrt(variance);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// DSTDEVP関数の実装（条件付き標準偏差：母集団）
export const DSTDEVP: CustomFormula = {
  name: 'DSTDEVP',
  pattern: /DSTDEVP\(([^,]+),\s*"?([^",]+)"?,\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, databaseRef, field, criteriaRef] = matches;
    
    try {
      // データベースと条件範囲を解析
      const dbRange = databaseRef.trim().match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
      const critRange = criteriaRef.trim().match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
      
      if (!dbRange || !critRange) {
        return FormulaError.REF;
      }
      
      // データベース配列を構築
      const database: any[][] = [];
      const [, dbStartCol, dbStartRowStr, dbEndCol, dbEndRowStr] = dbRange;
      const dbStartRow = parseInt(dbStartRowStr) - 1;
      const dbEndRow = parseInt(dbEndRowStr) - 1;
      const dbStartColIndex = dbStartCol.charCodeAt(0) - 65;
      const dbEndColIndex = dbEndCol.charCodeAt(0) - 65;
      
      for (let row = dbStartRow; row <= dbEndRow; row++) {
        const rowData: any[] = [];
        for (let col = dbStartColIndex; col <= dbEndColIndex; col++) {
          const cell = context.data[row]?.[col];
          rowData.push(cell ? cell.value : null);
        }
        database.push(rowData);
      }
      
      // 条件配列を構築
      const criteria: any[][] = [];
      const [, critStartCol, critStartRowStr, critEndCol, critEndRowStr] = critRange;
      const critStartRow = parseInt(critStartRowStr) - 1;
      const critEndRow = parseInt(critEndRowStr) - 1;
      const critStartColIndex = critStartCol.charCodeAt(0) - 65;
      const critEndColIndex = critEndCol.charCodeAt(0) - 65;
      
      for (let row = critStartRow; row <= critEndRow; row++) {
        const rowData: any[] = [];
        for (let col = critStartColIndex; col <= critEndColIndex; col++) {
          const cell = context.data[row]?.[col];
          rowData.push(cell ? cell.value : null);
        }
        criteria.push(rowData);
      }
      
      // 条件に一致する値を取得
      const { values } = getMatchingRows(database, field, criteria);
      
      if (values.length === 0) {
        return FormulaError.DIV0;
      }
      
      // 標準偏差（母集団）を計算
      const mean = values.reduce((sum, val) => sum + val, 0) / values.length;
      const squaredDiffs = values.map(val => Math.pow(val - mean, 2));
      const variance = squaredDiffs.reduce((sum, val) => sum + val, 0) / values.length;
      
      return Math.sqrt(variance);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// DVAR関数の実装（条件付き分散：標本）
export const DVAR: CustomFormula = {
  name: 'DVAR',
  pattern: /DVAR\(([^,]+),\s*"?([^",]+)"?,\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, databaseRef, field, criteriaRef] = matches;
    
    try {
      // データベースと条件範囲を解析
      const dbRange = databaseRef.trim().match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
      const critRange = criteriaRef.trim().match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
      
      if (!dbRange || !critRange) {
        return FormulaError.REF;
      }
      
      // データベース配列を構築
      const database: any[][] = [];
      const [, dbStartCol, dbStartRowStr, dbEndCol, dbEndRowStr] = dbRange;
      const dbStartRow = parseInt(dbStartRowStr) - 1;
      const dbEndRow = parseInt(dbEndRowStr) - 1;
      const dbStartColIndex = dbStartCol.charCodeAt(0) - 65;
      const dbEndColIndex = dbEndCol.charCodeAt(0) - 65;
      
      for (let row = dbStartRow; row <= dbEndRow; row++) {
        const rowData: any[] = [];
        for (let col = dbStartColIndex; col <= dbEndColIndex; col++) {
          const cell = context.data[row]?.[col];
          rowData.push(cell ? cell.value : null);
        }
        database.push(rowData);
      }
      
      // 条件配列を構築
      const criteria: any[][] = [];
      const [, critStartCol, critStartRowStr, critEndCol, critEndRowStr] = critRange;
      const critStartRow = parseInt(critStartRowStr) - 1;
      const critEndRow = parseInt(critEndRowStr) - 1;
      const critStartColIndex = critStartCol.charCodeAt(0) - 65;
      const critEndColIndex = critEndCol.charCodeAt(0) - 65;
      
      for (let row = critStartRow; row <= critEndRow; row++) {
        const rowData: any[] = [];
        for (let col = critStartColIndex; col <= critEndColIndex; col++) {
          const cell = context.data[row]?.[col];
          rowData.push(cell ? cell.value : null);
        }
        criteria.push(rowData);
      }
      
      // 条件に一致する値を取得
      const { values } = getMatchingRows(database, field, criteria);
      
      if (values.length < 2) {
        return FormulaError.DIV0;
      }
      
      // 分散（標本）を計算
      const mean = values.reduce((sum, val) => sum + val, 0) / values.length;
      const squaredDiffs = values.map(val => Math.pow(val - mean, 2));
      const variance = squaredDiffs.reduce((sum, val) => sum + val, 0) / (values.length - 1);
      
      return variance;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// DVARP関数の実装（条件付き分散：母集団）
export const DVARP: CustomFormula = {
  name: 'DVARP',
  pattern: /DVARP\(([^,]+),\s*"?([^",]+)"?,\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, databaseRef, field, criteriaRef] = matches;
    
    try {
      // データベースと条件範囲を解析
      const dbRange = databaseRef.trim().match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
      const critRange = criteriaRef.trim().match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
      
      if (!dbRange || !critRange) {
        return FormulaError.REF;
      }
      
      // データベース配列を構築
      const database: any[][] = [];
      const [, dbStartCol, dbStartRowStr, dbEndCol, dbEndRowStr] = dbRange;
      const dbStartRow = parseInt(dbStartRowStr) - 1;
      const dbEndRow = parseInt(dbEndRowStr) - 1;
      const dbStartColIndex = dbStartCol.charCodeAt(0) - 65;
      const dbEndColIndex = dbEndCol.charCodeAt(0) - 65;
      
      for (let row = dbStartRow; row <= dbEndRow; row++) {
        const rowData: any[] = [];
        for (let col = dbStartColIndex; col <= dbEndColIndex; col++) {
          const cell = context.data[row]?.[col];
          rowData.push(cell ? cell.value : null);
        }
        database.push(rowData);
      }
      
      // 条件配列を構築
      const criteria: any[][] = [];
      const [, critStartCol, critStartRowStr, critEndCol, critEndRowStr] = critRange;
      const critStartRow = parseInt(critStartRowStr) - 1;
      const critEndRow = parseInt(critEndRowStr) - 1;
      const critStartColIndex = critStartCol.charCodeAt(0) - 65;
      const critEndColIndex = critEndCol.charCodeAt(0) - 65;
      
      for (let row = critStartRow; row <= critEndRow; row++) {
        const rowData: any[] = [];
        for (let col = critStartColIndex; col <= critEndColIndex; col++) {
          const cell = context.data[row]?.[col];
          rowData.push(cell ? cell.value : null);
        }
        criteria.push(rowData);
      }
      
      // 条件に一致する値を取得
      const { values } = getMatchingRows(database, field, criteria);
      
      if (values.length === 0) {
        return FormulaError.DIV0;
      }
      
      // 分散（母集団）を計算
      const mean = values.reduce((sum, val) => sum + val, 0) / values.length;
      const squaredDiffs = values.map(val => Math.pow(val - mean, 2));
      const variance = squaredDiffs.reduce((sum, val) => sum + val, 0) / values.length;
      
      return variance;
    } catch {
      return FormulaError.VALUE;
    }
  }
};