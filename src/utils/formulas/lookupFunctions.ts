// 検索関数の実装

import type { CustomFormula, FormulaContext, FormulaResult } from './types';
import { FormulaError } from './types';
import { getCellValue, getCellRangeValues } from './utils';

// VLOOKUP関数の実装
export const VLOOKUP: CustomFormula = {
  name: 'VLOOKUP',
  pattern: /VLOOKUP\(([^,]+),\s*([^,]+),\s*(\d+)(?:,\s*(TRUE|FALSE|0|1))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, lookupValue, tableArray, colIndex, rangeLookup] = matches;
    
    try {
      // 検索値を取得
      const searchValue = getCellValue(lookupValue.trim(), context) || lookupValue.trim().replace(/^["']|["']$/g, '');
      
      // 列インデックスを数値に変換
      const columnIndex = parseInt(colIndex);
      if (columnIndex < 1) {
        return FormulaError.VALUE;
      }
      
      // テーブル配列から値を取得
      const tableValues = getCellRangeValues(tableArray.trim(), context);
      
      // 範囲検索の設定（デフォルトはTRUE）
      const exactMatch = rangeLookup && (rangeLookup.toUpperCase() === 'FALSE' || rangeLookup === '0');
      
      // テーブルが2次元配列になっているかチェック
      if (!Array.isArray(tableValues) || tableValues.length === 0) {
        return FormulaError.REF;
      }
      
      // 行数と列数を計算（範囲から推定）
      const rangeParts = tableArray.trim().split(':');
      if (rangeParts.length !== 2) {
        return FormulaError.REF;
      }
      
      // セル範囲から行数と列数を計算
      const [startCell, endCell] = rangeParts;
      const startMatch = startCell.match(/([A-Z]+)(\d+)/);
      const endMatch = endCell.match(/([A-Z]+)(\d+)/);
      
      if (!startMatch || !endMatch) {
        return FormulaError.REF;
      }
      
      const startCol = startMatch[1].charCodeAt(0) - 'A'.charCodeAt(0);
      const endCol = endMatch[1].charCodeAt(0) - 'A'.charCodeAt(0);
      const cols = endCol - startCol + 1;
      
      if (columnIndex > cols) {
        return FormulaError.REF;
      }
      
      const rows = Math.ceil(tableValues.length / cols);
      
      // 2次元配列に変換
      const table: unknown[][] = [];
      for (let i = 0; i < rows; i++) {
        const row: unknown[] = [];
        for (let j = 0; j < cols; j++) {
          const index = i * cols + j;
          row.push(index < tableValues.length ? tableValues[index] : null);
        }
        table.push(row);
      }
      
      // 検索実行
      for (let i = 0; i < table.length; i++) {
        const firstColumnValue = table[i][0];
        
        if (exactMatch) {
          // 完全一致検索
          if (String(firstColumnValue) === String(searchValue)) {
            return table[i][columnIndex - 1] || null;
          }
        } else {
          // 範囲検索（近似一致）
          const numericSearch = Number(searchValue);
          const numericFirst = Number(firstColumnValue);
          
          if (!isNaN(numericSearch) && !isNaN(numericFirst)) {
            if (numericFirst <= numericSearch) {
              // 次の行をチェック
              if (i === table.length - 1 || Number(table[i + 1][0]) > numericSearch) {
                return table[i][columnIndex - 1] || null;
              }
            }
          } else if (String(firstColumnValue) === String(searchValue)) {
            return table[i][columnIndex - 1] || null;
          }
        }
      }
      
      return FormulaError.NA;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// HLOOKUP関数の実装
export const HLOOKUP: CustomFormula = {
  name: 'HLOOKUP',
  pattern: /HLOOKUP\(([^,]+),\s*([^,]+),\s*(\d+)(?:,\s*(TRUE|FALSE|0|1))?\)/i,
  calculate: () => null // HyperFormulaが処理
};

// INDEX関数の実装
export const INDEX: CustomFormula = {
  name: 'INDEX',
  pattern: /INDEX\(([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, arrayRange, rowNum, colNum] = matches;
    
    try {
      // 配列の値を取得
      const arrayValues = getCellRangeValues(arrayRange.trim(), context);
      
      // 行番号を取得
      const rowIndex = parseInt(getCellValue(rowNum.trim(), context)?.toString() || rowNum.trim()) - 1;
      
      // 列番号を取得（省略可能）
      const colIndex = colNum 
        ? parseInt(getCellValue(colNum.trim(), context)?.toString() || colNum.trim()) - 1
        : 0;
      
      if (rowIndex < 0 || colIndex < 0) {
        return FormulaError.VALUE;
      }
      
      // 範囲から行数と列数を計算
      const rangeParts = arrayRange.trim().split(':');
      if (rangeParts.length !== 2) {
        return FormulaError.REF;
      }
      
      const [startCell, endCell] = rangeParts;
      const startMatch = startCell.match(/([A-Z]+)(\d+)/);
      const endMatch = endCell.match(/([A-Z]+)(\d+)/);
      
      if (!startMatch || !endMatch) {
        return FormulaError.REF;
      }
      
      const startCol = startMatch[1].charCodeAt(0) - 'A'.charCodeAt(0);
      const endCol = endMatch[1].charCodeAt(0) - 'A'.charCodeAt(0);
      const cols = endCol - startCol + 1;
      const rows = Math.ceil(arrayValues.length / cols);
      
      if (rowIndex >= rows || colIndex >= cols) {
        return FormulaError.REF;
      }
      
      // 2次元配列のインデックスを1次元配列のインデックスに変換
      const flatIndex = rowIndex * cols + colIndex;
      
      if (flatIndex >= arrayValues.length) {
        return FormulaError.REF;
      }
      
      return arrayValues[flatIndex] || null;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// MATCH関数の実装
export const MATCH: CustomFormula = {
  name: 'MATCH',
  pattern: /MATCH\(([^,]+),\s*([^,]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, lookupValue, lookupArray, matchType] = matches;
    
    try {
      // 検索値を取得
      const searchValue = getCellValue(lookupValue.trim(), context) || lookupValue.trim().replace(/^["']|["']$/g, '');
      
      // 配列の値を取得
      const arrayValues = getCellRangeValues(lookupArray.trim(), context);
      
      // 一致タイプを取得（デフォルトは1）
      const matchTypeNum = matchType ? parseInt(matchType) : 1;
      
      switch (matchTypeNum) {
        case 1: // 近似一致（昇順）
          {
            let bestMatch = -1;
            const numericSearch = Number(searchValue);
            
            for (let i = 0; i < arrayValues.length; i++) {
              const currentValue = arrayValues[i];
              const numericCurrent = Number(currentValue);
              
              if (!isNaN(numericSearch) && !isNaN(numericCurrent)) {
                if (numericCurrent <= numericSearch) {
                  bestMatch = i;
                } else {
                  break; // 配列は昇順ソートされている前提
                }
              } else if (String(currentValue) === String(searchValue)) {
                return i + 1; // 1ベースのインデックス
              }
            }
            
            return bestMatch >= 0 ? bestMatch + 1 : FormulaError.NA;
          }
        
        case 0: // 完全一致
          for (let i = 0; i < arrayValues.length; i++) {
            if (String(arrayValues[i]) === String(searchValue)) {
              return i + 1; // 1ベースのインデックス
            }
          }
          return FormulaError.NA;
        
        case -1: // 近似一致（降順）
          {
            let bestMatch = -1;
            const numericSearch = Number(searchValue);
            
            for (let i = 0; i < arrayValues.length; i++) {
              const currentValue = arrayValues[i];
              const numericCurrent = Number(currentValue);
              
              if (!isNaN(numericSearch) && !isNaN(numericCurrent)) {
                if (numericCurrent >= numericSearch) {
                  bestMatch = i;
                } else {
                  break; // 配列は降順ソートされている前提
                }
              } else if (String(currentValue) === String(searchValue)) {
                return i + 1; // 1ベースのインデックス
              }
            }
            
            return bestMatch >= 0 ? bestMatch + 1 : FormulaError.NA;
          }
        
        default:
          return FormulaError.VALUE;
      }
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// LOOKUP関数の実装
export const LOOKUP: CustomFormula = {
  name: 'LOOKUP',
  pattern: /LOOKUP\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// XLOOKUP関数の実装（手動実装が必要）
export const XLOOKUP: CustomFormula = {
  name: 'XLOOKUP',
  pattern: /XLOOKUP\(([^,]+),\s*([^,]+),\s*([^,]+)(?:,\s*([^,)]+))?(?:,\s*([^,)]+))?(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, lookupValue, lookupArray, returnArray, ifNotFound] = matches;
    
    console.log('XLOOKUP計算:', { lookupValue, lookupArray, returnArray });
    
    // セル参照から値を取得
    let searchValue: unknown;
    if (lookupValue.match(/^[A-Z]+\d+$/)) {
      searchValue = getCellValue(lookupValue, context);
    } else {
      searchValue = lookupValue.replace(/^"|"$/g, '');
    }
    
    // 検索配列の値を取得
    const lookupValues = getCellRangeValues(lookupArray, context);
    // 返り値配列の値を取得
    const returnValues = getCellRangeValues(returnArray, context);
    
    if (lookupValues.length !== returnValues.length) {
      return FormulaError.REF;
    }
    
    // 完全一致で検索（デフォルト）
    for (let i = 0; i < lookupValues.length; i++) {
      if (String(lookupValues[i]).toLowerCase() === String(searchValue).toLowerCase()) {
        return returnValues[i] as FormulaResult;
      }
    }
    
    // 見つからない場合
    if (ifNotFound) {
      return ifNotFound.replace(/^"|"$/g, '');
    }
    
    return FormulaError.NA;
  }
};