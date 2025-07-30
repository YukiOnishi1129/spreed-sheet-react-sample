// データベース統計関数

import type { CustomFormula, FormulaContext, FormulaResult } from '../shared/types';
import { FormulaError } from '../shared/types';

// データベース関数用のヘルパー：条件に一致する行を取得
function getMatchingRows(
  database: (number | string | boolean | null)[][],
  field: string | number,
  criteria: (number | string | boolean | null)[][]
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
    // フィールドが数値の場合は1ベースのインデックスとして扱う
    const numField = parseInt(String(field));
    if (!isNaN(numField) && numField > 0) {
      fieldIndex = numField - 1;
    }
  } else {
    // フィールド名で検索（大文字小文字を区別しない）
    const fieldName = String(field).toUpperCase();
    fieldIndex = headers.indexOf(fieldName);
  }

  if (fieldIndex < 0 || fieldIndex >= headers.length) {
    return { values: [], fieldIndex: -1 };
  }

  // 条件に一致する行の値を収集
  const values: number[] = [];

  // データベースの各行をチェック
  for (let row = 1; row < database.length; row++) {
    let rowMatches = false;

    // 条件の各行をチェック（いずれかの条件行に一致すればOK）
    for (let cRow = 1; cRow < criteria.length; cRow++) {
      let allCriteriaMatch = true;

      // この条件行のすべての列をチェック
      for (let cCol = 0; cCol < criteriaHeaders.length; cCol++) {
        const criteriaHeader = criteriaHeaders[cCol];
        const dbCol = headers.indexOf(criteriaHeader);

        if (dbCol >= 0) {
          const criteriaValue = criteria[cRow][cCol];
          
          // 空の条件はスキップ（すべてに一致）
          if (criteriaValue === null || criteriaValue === undefined || criteriaValue === '') {
            continue;
          }

          const dbValue = database[row][dbCol];
          const criteriaStr = String(criteriaValue);
          let matches = false;

          // 比較演算子を含む条件の処理
          const operatorMatch = criteriaStr.match(/^([><=!]+)(.*)$/);
          if (operatorMatch) {
            const operator = operatorMatch[1];
            const compareValue = operatorMatch[2];
            const dbNum = parseFloat(String(dbValue));
            const compareNum = parseFloat(compareValue);

            if (!isNaN(dbNum) && !isNaN(compareNum)) {
              switch (operator) {
                case '>':
                  matches = dbNum > compareNum;
                  break;
                case '>=':
                  matches = dbNum >= compareNum;
                  break;
                case '<':
                  matches = dbNum < compareNum;
                  break;
                case '<=':
                  matches = dbNum <= compareNum;
                  break;
                case '<>':
                case '!=':
                  matches = dbNum !== compareNum;
                  break;
                case '=':
                  matches = dbNum === compareNum;
                  break;
              }
            } else {
              // 数値でない場合は文字列として比較
              matches = String(dbValue).toUpperCase() === criteriaStr.toUpperCase();
            }
          } else {
            // 完全一致（大文字小文字を区別しない）
            matches = String(dbValue).toUpperCase() === criteriaStr.toUpperCase();
          }

          if (!matches) {
            allCriteriaMatch = false;
            break;
          }
        }
      }

      if (allCriteriaMatch) {
        rowMatches = true;
        break;
      }
    }

    if (rowMatches) {
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
  pattern: /\bDSTDEV\(([^,]+),\s*"?([^",]+)"?,\s*([^)]+)\)/i,
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
      const database: (number | string | boolean | null)[][] = [];
      const [, dbStartCol, dbStartRowStr, dbEndCol, dbEndRowStr] = dbRange;
      const dbStartRow = parseInt(dbStartRowStr) - 1;
      const dbEndRow = parseInt(dbEndRowStr) - 1;
      const dbStartColIndex = dbStartCol.charCodeAt(0) - 65;
      const dbEndColIndex = dbEndCol.charCodeAt(0) - 65;
      
      // コンテキストデータが空の場合のチェック
      if (!context.data || context.data.length === 0) {
        return FormulaError.VALUE;
      }
      
      for (let row = dbStartRow; row <= dbEndRow; row++) {
        const rowData: (number | string | boolean | null)[] = [];
        for (let col = dbStartColIndex; col <= dbEndColIndex; col++) {
          const cell = context.data[row]?.[col];
          rowData.push(cell ? cell.value as (number | string | boolean | null) : null);
        }
        database.push(rowData);
      }
      
      // 条件配列を構築
      const criteria: (number | string | boolean | null)[][] = [];
      const [, critStartCol, critStartRowStr, critEndCol, critEndRowStr] = critRange;
      const critStartRow = parseInt(critStartRowStr) - 1;
      const critEndRow = parseInt(critEndRowStr) - 1;
      const critStartColIndex = critStartCol.charCodeAt(0) - 65;
      const critEndColIndex = critEndCol.charCodeAt(0) - 65;
      
      for (let row = critStartRow; row <= critEndRow; row++) {
        const rowData: (number | string | boolean | null)[] = [];
        for (let col = critStartColIndex; col <= critEndColIndex; col++) {
          const cell = context.data[row]?.[col];
          rowData.push(cell ? cell.value as (number | string | boolean | null) : null);
        }
        criteria.push(rowData);
      }
      
      // フィールドパラメータの処理
      let processedField: string | number = field;
      if (field.match(/^\d+$/)) {
        processedField = parseInt(field);
      } else if (field.startsWith('"') && field.endsWith('"')) {
        processedField = field.slice(1, -1);
      }
      
      // 条件に一致する値を取得
      const { values } = getMatchingRows(database, processedField, criteria);
      
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
  pattern: /\bDSTDEVP\(([^,]+),\s*"?([^",]+)"?,\s*([^)]+)\)/i,
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
      const database: (number | string | boolean | null)[][] = [];
      const [, dbStartCol, dbStartRowStr, dbEndCol, dbEndRowStr] = dbRange;
      const dbStartRow = parseInt(dbStartRowStr) - 1;
      const dbEndRow = parseInt(dbEndRowStr) - 1;
      const dbStartColIndex = dbStartCol.charCodeAt(0) - 65;
      const dbEndColIndex = dbEndCol.charCodeAt(0) - 65;
      
      // コンテキストデータが空の場合のチェック
      if (!context.data || context.data.length === 0) {
        return FormulaError.VALUE;
      }
      
      for (let row = dbStartRow; row <= dbEndRow; row++) {
        const rowData: (number | string | boolean | null)[] = [];
        for (let col = dbStartColIndex; col <= dbEndColIndex; col++) {
          const cell = context.data[row]?.[col];
          rowData.push(cell ? cell.value as (number | string | boolean | null) : null);
        }
        database.push(rowData);
      }
      
      // 条件配列を構築
      const criteria: (number | string | boolean | null)[][] = [];
      const [, critStartCol, critStartRowStr, critEndCol, critEndRowStr] = critRange;
      const critStartRow = parseInt(critStartRowStr) - 1;
      const critEndRow = parseInt(critEndRowStr) - 1;
      const critStartColIndex = critStartCol.charCodeAt(0) - 65;
      const critEndColIndex = critEndCol.charCodeAt(0) - 65;
      
      for (let row = critStartRow; row <= critEndRow; row++) {
        const rowData: (number | string | boolean | null)[] = [];
        for (let col = critStartColIndex; col <= critEndColIndex; col++) {
          const cell = context.data[row]?.[col];
          rowData.push(cell ? cell.value as (number | string | boolean | null) : null);
        }
        criteria.push(rowData);
      }
      
      // フィールドパラメータの処理
      let processedField: string | number = field;
      if (field.match(/^\d+$/)) {
        processedField = parseInt(field);
      } else if (field.startsWith('"') && field.endsWith('"')) {
        processedField = field.slice(1, -1);
      }
      
      // 条件に一致する値を取得
      const { values } = getMatchingRows(database, processedField, criteria);
      
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
  pattern: /\bDVAR\(([^,]+),\s*"?([^",]+)"?,\s*([^)]+)\)/i,
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
      const database: (number | string | boolean | null)[][] = [];
      const [, dbStartCol, dbStartRowStr, dbEndCol, dbEndRowStr] = dbRange;
      const dbStartRow = parseInt(dbStartRowStr) - 1;
      const dbEndRow = parseInt(dbEndRowStr) - 1;
      const dbStartColIndex = dbStartCol.charCodeAt(0) - 65;
      const dbEndColIndex = dbEndCol.charCodeAt(0) - 65;
      
      // コンテキストデータが空の場合のチェック
      if (!context.data || context.data.length === 0) {
        return FormulaError.VALUE;
      }
      
      for (let row = dbStartRow; row <= dbEndRow; row++) {
        const rowData: (number | string | boolean | null)[] = [];
        for (let col = dbStartColIndex; col <= dbEndColIndex; col++) {
          const cell = context.data[row]?.[col];
          rowData.push(cell ? cell.value as (number | string | boolean | null) : null);
        }
        database.push(rowData);
      }
      
      // 条件配列を構築
      const criteria: (number | string | boolean | null)[][] = [];
      const [, critStartCol, critStartRowStr, critEndCol, critEndRowStr] = critRange;
      const critStartRow = parseInt(critStartRowStr) - 1;
      const critEndRow = parseInt(critEndRowStr) - 1;
      const critStartColIndex = critStartCol.charCodeAt(0) - 65;
      const critEndColIndex = critEndCol.charCodeAt(0) - 65;
      
      for (let row = critStartRow; row <= critEndRow; row++) {
        const rowData: (number | string | boolean | null)[] = [];
        for (let col = critStartColIndex; col <= critEndColIndex; col++) {
          const cell = context.data[row]?.[col];
          rowData.push(cell ? cell.value as (number | string | boolean | null) : null);
        }
        criteria.push(rowData);
      }
      
      // フィールドパラメータの処理
      let processedField: string | number = field;
      if (field.match(/^\d+$/)) {
        processedField = parseInt(field);
      } else if (field.startsWith('"') && field.endsWith('"')) {
        processedField = field.slice(1, -1);
      }
      
      // 条件に一致する値を取得
      const { values } = getMatchingRows(database, processedField, criteria);
      
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
  pattern: /\bDVARP\(([^,]+),\s*"?([^",]+)"?,\s*([^)]+)\)/i,
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
      const database: (number | string | boolean | null)[][] = [];
      const [, dbStartCol, dbStartRowStr, dbEndCol, dbEndRowStr] = dbRange;
      const dbStartRow = parseInt(dbStartRowStr) - 1;
      const dbEndRow = parseInt(dbEndRowStr) - 1;
      const dbStartColIndex = dbStartCol.charCodeAt(0) - 65;
      const dbEndColIndex = dbEndCol.charCodeAt(0) - 65;
      
      // コンテキストデータが空の場合のチェック
      if (!context.data || context.data.length === 0) {
        return FormulaError.VALUE;
      }
      
      for (let row = dbStartRow; row <= dbEndRow; row++) {
        const rowData: (number | string | boolean | null)[] = [];
        for (let col = dbStartColIndex; col <= dbEndColIndex; col++) {
          const cell = context.data[row]?.[col];
          rowData.push(cell ? cell.value as (number | string | boolean | null) : null);
        }
        database.push(rowData);
      }
      
      // 条件配列を構築
      const criteria: (number | string | boolean | null)[][] = [];
      const [, critStartCol, critStartRowStr, critEndCol, critEndRowStr] = critRange;
      const critStartRow = parseInt(critStartRowStr) - 1;
      const critEndRow = parseInt(critEndRowStr) - 1;
      const critStartColIndex = critStartCol.charCodeAt(0) - 65;
      const critEndColIndex = critEndCol.charCodeAt(0) - 65;
      
      for (let row = critStartRow; row <= critEndRow; row++) {
        const rowData: (number | string | boolean | null)[] = [];
        for (let col = critStartColIndex; col <= critEndColIndex; col++) {
          const cell = context.data[row]?.[col];
          rowData.push(cell ? cell.value as (number | string | boolean | null) : null);
        }
        criteria.push(rowData);
      }
      
      // フィールドパラメータの処理
      let processedField: string | number = field;
      if (field.match(/^\d+$/)) {
        processedField = parseInt(field);
      } else if (field.startsWith('"') && field.endsWith('"')) {
        processedField = field.slice(1, -1);
      }
      
      // 条件に一致する値を取得
      const { values } = getMatchingRows(database, processedField, criteria);
      
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