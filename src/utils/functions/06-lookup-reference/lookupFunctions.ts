// 検索関数の実装

import type { CustomFormula, FormulaContext, FormulaResult } from '../shared/types';
import { FormulaError } from '../shared/types';
import { getCellValue, getCellRangeValues } from '../shared/utils';

// VLOOKUP関数の実装
export const VLOOKUP: CustomFormula = {
  name: 'VLOOKUP',
  pattern: /VLOOKUP\(([^,]+),\s*([^,]+),\s*(\d+)(?:,\s*(TRUE|FALSE|0|1))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, lookupValue, tableArray, colIndex, rangeLookup] = matches;
    
    try {
      // 検索値を取得
      const searchValue = getCellValue(lookupValue.trim(), context) ?? lookupValue.trim().replace(/^["']|["']$/g, '');
      
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
            const result = table[i][columnIndex - 1];
            return result === undefined ? null : result as FormulaResult;
          }
        } else {
          // 範囲検索（近似一致）
          const numericSearch = Number(searchValue);
          const numericFirst = Number(firstColumnValue);
          
          if (!isNaN(numericSearch) && !isNaN(numericFirst)) {
            if (numericFirst <= numericSearch) {
              // 次の行をチェック
              if (i === table.length - 1 || Number(table[i + 1][0]) > numericSearch) {
                const result = table[i][columnIndex - 1];
                return result === undefined ? null : result as FormulaResult;
              }
            }
          } else if (String(firstColumnValue) === String(searchValue)) {
            const result = table[i][columnIndex - 1];
            return result === undefined ? null : result as FormulaResult;
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
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, lookupValue, tableArray, rowIndex, rangeLookup] = matches;
    
    try {
      // 検索値を取得
      const searchValue = getCellValue(lookupValue.trim(), context) ?? lookupValue.trim().replace(/^["']|["']$/g, '');
      
      // 行インデックスを数値に変換
      const rowIndexNum = parseInt(rowIndex);
      if (rowIndexNum < 1) {
        return FormulaError.VALUE;
      }
      
      // テーブル配列から値を取得
      const tableValues = getCellRangeValues(tableArray.trim(), context);
      
      // 範囲検索の設定（デフォルトはTRUE）
      const exactMatch = rangeLookup && (rangeLookup.toUpperCase() === 'FALSE' || rangeLookup === '0');
      
      // テーブルが配列になっているかチェック
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
      const startRow = parseInt(startMatch[2]);
      const endRow = parseInt(endMatch[2]);
      
      const cols = endCol - startCol + 1;
      const rows = endRow - startRow + 1;
      
      if (rowIndexNum > rows) {
        return FormulaError.REF;
      }
      
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
      
      // 最初の行で検索
      const firstRow = table[0];
      
      for (let j = 0; j < firstRow.length; j++) {
        const cellValue = firstRow[j];
        
        if (exactMatch) {
          // 完全一致検索
          if (String(cellValue) === String(searchValue)) {
            const result = table[rowIndexNum - 1][j];
            return result === undefined ? null : result as FormulaResult;
          }
        } else {
          // 範囲検索（近似一致）
          const numericSearch = Number(searchValue);
          const numericCell = Number(cellValue);
          
          if (!isNaN(numericSearch) && !isNaN(numericCell)) {
            if (numericCell <= numericSearch) {
              // 次の列をチェック
              if (j === firstRow.length - 1 || Number(firstRow[j + 1]) > numericSearch) {
                const result = table[rowIndexNum - 1][j];
                return result === undefined ? null : result as FormulaResult;
              }
            }
          } else if (String(cellValue) === String(searchValue)) {
            const result = table[rowIndexNum - 1][j];
            return result === undefined ? null : result as FormulaResult;
          }
        }
      }
      
      return FormulaError.NA;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// INDEX関数の実装
export const INDEX: CustomFormula = {
  name: 'INDEX',
  pattern: /INDEX\(([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, arrayRange, rowNum, colNum] = matches;
    
    try {
      // 配列の値を取得
      const arrayValues = getCellRangeValues(arrayRange.trim(), context);
      
      // 行番号を取得
      const rowIndex = parseInt(getCellValue(rowNum.trim(), context)?.toString() ?? rowNum.trim()) - 1;
      
      // 列番号を取得（省略可能）
      const colIndex = colNum 
        ? parseInt(getCellValue(colNum.trim(), context)?.toString() ?? colNum.trim()) - 1
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
      
      const result = arrayValues[flatIndex];
      return result === undefined ? null : result as FormulaResult;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// MATCH関数の実装
export const MATCH: CustomFormula = {
  name: 'MATCH',
  pattern: /MATCH\(([^,]+),\s*([^,]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, lookupValue, lookupArray, matchType] = matches;
    
    try {
      // 検索値を取得
      const searchValue = getCellValue(lookupValue.trim(), context) ?? lookupValue.trim().replace(/^["']|["']$/g, '');
      
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
  pattern: /LOOKUP\(([^,]+),\s*([^,]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, lookupValue, lookupVector, resultVector] = matches;
    
    try {
      // 検索値を取得
      const searchValue = getCellValue(lookupValue.trim(), context) ?? lookupValue.trim().replace(/^["']|["']$/g, '');
      
      // 検索ベクトルの値を取得
      const lookupValues = getCellRangeValues(lookupVector.trim(), context);
      
      // 結果ベクトルの値を取得（省略時は検索ベクトルと同じ）
      let returnValues = lookupValues;
      if (resultVector) {
        returnValues = getCellRangeValues(resultVector.trim(), context);
      }
      
      if (lookupValues.length !== returnValues.length) {
        return FormulaError.NA;
      }
      
      // 近似一致検索（配列が昇順に並んでいることを前提）
      let bestMatch = -1;
      const numericSearch = Number(searchValue);
      
      for (let i = 0; i < lookupValues.length; i++) {
        const currentValue = lookupValues[i];
        const numericCurrent = Number(currentValue);
        
        if (!isNaN(numericSearch) && !isNaN(numericCurrent)) {
          if (numericCurrent <= numericSearch) {
            bestMatch = i;
          } else {
            break;
          }
        } else if (String(currentValue) <= String(searchValue)) {
          bestMatch = i;
        } else {
          break;
        }
      }
      
      if (bestMatch >= 0) {
        const result = returnValues[bestMatch];
        return result as FormulaResult;
      }
      
      return FormulaError.NA;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// XLOOKUP関数の実装（手動実装が必要）
export const XLOOKUP: CustomFormula = {
  name: 'XLOOKUP',
  pattern: /XLOOKUP\(([^,]+),\s*([^,]+),\s*([^,]+)(?:,\s*([^,)]+))?(?:,\s*([^,)]+))?(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, lookupValue, lookupArray, returnArray, ifNotFound] = matches;
    
    // XLOOKUP計算実行
    
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

// OFFSET関数の実装
export const OFFSET: CustomFormula = {
  name: 'OFFSET',
  pattern: /OFFSET\(([^,]+),\s*([^,]+),\s*([^,)]+)(?:,\s*([^,)]+))?(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, referenceRef, rowsRef, colsRef, heightRef, widthRef] = matches;
    
    try {
      // 基準セル参照を取得
      const reference = referenceRef.trim();
      const rows = parseInt(getCellValue(rowsRef.trim(), context)?.toString() ?? rowsRef.trim());
      const cols = parseInt(getCellValue(colsRef.trim(), context)?.toString() ?? colsRef.trim());
      const height = heightRef ? parseInt(getCellValue(heightRef.trim(), context)?.toString() ?? heightRef.trim()) : 1;
      const width = widthRef ? parseInt(getCellValue(widthRef.trim(), context)?.toString() ?? widthRef.trim()) : 1;
      
      if (isNaN(rows) || isNaN(cols) || isNaN(height) || isNaN(width) || height < 1 || width < 1) {
        return FormulaError.VALUE;
      }
      
      // 基準セルの位置を解析
      const refMatch = reference.match(/([A-Z]+)(\d+)/);
      if (!refMatch) {
        return FormulaError.REF;
      }
      
      const [, refCol, refRowStr] = refMatch;
      const refRow = parseInt(refRowStr);
      const refColIndex = refCol.split('').reduce((acc, char) => acc * 26 + char.charCodeAt(0) - 64, 0) - 1;
      
      // オフセット後の位置を計算
      const newRow = refRow + rows;
      const newColIndex = refColIndex + cols;
      
      if (newRow < 1 || newColIndex < 0) {
        return FormulaError.REF;
      }
      
      // 列名を生成
      const getColumnName = (index: number): string => {
        let result = '';
        while (index >= 0) {
          result = String.fromCharCode(65 + (index % 26)) + result;
          index = Math.floor(index / 26) - 1;
        }
        return result;
      };
      
      // 単一セルの場合
      if (height === 1 && width === 1) {
        const newColName = getColumnName(newColIndex);
        const newCellRef = `${newColName}${newRow}`;
        return getCellValue(newCellRef, context) as FormulaResult;
      }
      
      // 範囲を構築
      const startColName = getColumnName(newColIndex);
      const endColName = getColumnName(newColIndex + width - 1);
      const startRow = newRow;
      const endRow = newRow + height - 1;
      
      const rangeRef = `${startColName}${startRow}:${endColName}${endRow}`;
      
      try {
        return getCellRangeValues(rangeRef, context) as FormulaResult;
      } catch {
        return FormulaError.REF;
      }
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// INDIRECT関数の実装
export const INDIRECT: CustomFormula = {
  name: 'INDIRECT',
  pattern: /INDIRECT\(([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, refTextRef, a1Ref] = matches;
    
    try {
      // 参照テキストを取得
      const refText = getCellValue(refTextRef.trim(), context) ?? refTextRef.trim().replace(/^['"]|['"]$/g, '');
      
      // A1形式かどうか（デフォルトはTRUE）
      const a1Style = a1Ref ? (a1Ref.trim().toUpperCase() !== 'FALSE' && a1Ref.trim() !== '0') : true;
      
      if (!a1Style) {
        // R1C1形式は未対応
        return FormulaError.VALUE;
      }
      
      if (typeof refText !== 'string') {
        return FormulaError.REF;
      }
      
      // セル参照として解釈
      if (refText.includes(':')) {
        // 範囲参照
        return getCellRangeValues(refText, context) as FormulaResult;
      } else {
        // 単一セル参照
        return getCellValue(refText, context) as FormulaResult;
      }
    } catch {
      return FormulaError.REF;
    }
  }
};

// CHOOSE関数の実装
export const CHOOSE: CustomFormula = {
  name: 'CHOOSE',
  pattern: /CHOOSE\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, indexRef, valuesRef] = matches;
    
    try {
      // インデックスを取得
      const index = parseInt(getCellValue(indexRef.trim(), context)?.toString() ?? indexRef.trim());
      
      if (isNaN(index) || index < 1 || !Number.isInteger(index)) {
        return FormulaError.VALUE;
      }
      
      // 値のリストを分割
      const valueList = valuesRef.split(',').map(val => val.trim());
      
      if (index > valueList.length) {
        return FormulaError.VALUE;
      }
      
      // 選択された値を返す
      const selectedValue = valueList[index - 1];
      
      // セル参照の場合は値を取得、そうでなければそのまま返す
      if (selectedValue.match(/^[A-Z]+\d+$/)) {
        return getCellValue(selectedValue, context) as FormulaResult;
      } else {
        return selectedValue.replace(/^['"]|['"]$/g, '');
      }
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// TRANSPOSE関数の実装
export const TRANSPOSE: CustomFormula = {
  name: 'TRANSPOSE',
  pattern: /TRANSPOSE\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, arrayRef] = matches;
    
    try {
      // 配列を取得
      const values = getCellRangeValues(arrayRef.trim(), context);
      
      if (!Array.isArray(values) || values.length === 0) {
        return FormulaError.VALUE;
      }
      
      // 範囲から行数と列数を計算
      const rangeParts = arrayRef.trim().split(':');
      if (rangeParts.length !== 2) {
        return values as FormulaResult; // 単一セルの場合はそのまま返す
      }
      
      const [startCell, endCell] = rangeParts;
      const startMatch = startCell.match(/([A-Z]+)(\d+)/);
      const endMatch = endCell.match(/([A-Z]+)(\d+)/);
      
      if (!startMatch || !endMatch) {
        return FormulaError.REF;
      }
      
      const startCol = startMatch[1].charCodeAt(0) - 'A'.charCodeAt(0);
      const endCol = endMatch[1].charCodeAt(0) - 'A'.charCodeAt(0);
      const startRow = parseInt(startMatch[2]);
      const endRow = parseInt(endMatch[2]);
      
      const cols = endCol - startCol + 1;
      const rows = endRow - startRow + 1;
      
      // 2次元配列に変換
      const matrix: unknown[][] = [];
      for (let i = 0; i < rows; i++) {
        const row: unknown[] = [];
        for (let j = 0; j < cols; j++) {
          const index = i * cols + j;
          row.push(index < values.length ? values[index] : null);
        }
        matrix.push(row);
      }
      
      // 転置
      const transposed: unknown[][] = [];
      for (let j = 0; j < cols; j++) {
        const newRow: unknown[] = [];
        for (let i = 0; i < rows; i++) {
          newRow.push(matrix[i][j]);
        }
        transposed.push(newRow);
      }
      
      // 1次元配列として返す（フラット化）
      const result = transposed.flat();
      return result.length === 1 ? result[0] as FormulaResult : result as FormulaResult;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// FILTER関数の実装
export const FILTER: CustomFormula = {
  name: 'FILTER',
  pattern: /FILTER\(([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, arrayRef, includeRef, ifEmptyRef] = matches;
    
    try {
      // 配列データを取得
      const arrayValues = getCellRangeValues(arrayRef.trim(), context);
      const includeValues = getCellRangeValues(includeRef.trim(), context);
      
      if (arrayValues.length !== includeValues.length) {
        return FormulaError.VALUE;
      }
      
      // フィルタ条件に基づいて要素を抽出
      const filteredValues: unknown[] = [];
      for (let i = 0; i < arrayValues.length; i++) {
        const includeValue = includeValues[i];
        // TRUE、1、または0以外の数値をTRUEとして扱う
        if (includeValue === true || includeValue === 1 || 
            (typeof includeValue === 'string' && includeValue.toUpperCase() === 'TRUE') ||
            (typeof includeValue === 'number' && includeValue !== 0)) {
          filteredValues.push(arrayValues[i]);
        }
      }
      
      // 結果が空の場合
      if (filteredValues.length === 0) {
        if (ifEmptyRef) {
          const emptyValue = getCellValue(ifEmptyRef.trim(), context) ?? ifEmptyRef.trim().replace(/^['"]|['"]$/g, '');
          return emptyValue as FormulaResult;
        }
        return FormulaError.CALC; // #CALC!エラー
      }
      
      return filteredValues.length === 1 ? filteredValues[0] as FormulaResult : filteredValues as FormulaResult;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// SORT関数の実装
export const SORT: CustomFormula = {
  name: 'SORT',
  pattern: /SORT\(([^,)]+)(?:,\s*([^,)]+))?(?:,\s*([^,)]+))?(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, arrayRef, sortIndexRef, sortOrderRef, byColRef] = matches;
    
    try {
      // 配列データを取得
      const values = getCellRangeValues(arrayRef.trim(), context);
      
      if (!Array.isArray(values) || values.length === 0) {
        return FormulaError.VALUE;
      }
      
      // ソートインデックス（デフォルトは1列目）
      const sortIndex = sortIndexRef ? parseInt(getCellValue(sortIndexRef.trim(), context)?.toString() ?? sortIndexRef.trim()) : 1;
      
      // ソート順序（デフォルトは昇順=1、降順=-1）
      const sortOrder = sortOrderRef ? parseInt(getCellValue(sortOrderRef.trim(), context)?.toString() ?? sortOrderRef.trim()) : 1;
      
      // 列でソートするか（デフォルトはFALSE=行でソート）
      const byCol = byColRef ? (getCellValue(byColRef.trim(), context)?.toString().toUpperCase() === 'TRUE' || byColRef.trim() === '1') : false;
      
      if (byCol) {
        // 列でのソートは複雑なため、簡単な実装
        return FormulaError.VALUE; // 未対応
      }
      
      // 範囲から行数と列数を計算
      const rangeParts = arrayRef.trim().split(':');
      if (rangeParts.length !== 2) {
        // 単一列の場合、そのままソート
        const sortedValues = [...values].sort((a, b) => {
          const aVal = a === null || a === undefined ? '' : String(a);
          const bVal = b === null || b === undefined ? '' : String(b);
          
          // 数値として比較可能な場合
          const aNum = Number(aVal);
          const bNum = Number(bVal);
          if (!isNaN(aNum) && !isNaN(bNum)) {
            return sortOrder === 1 ? aNum - bNum : bNum - aNum;
          }
          
          // 文字列として比較
          return sortOrder === 1 ? aVal.localeCompare(bVal) : bVal.localeCompare(aVal);
        });
        
        return sortedValues as FormulaResult;
      }
      
      const [startCell, endCell] = rangeParts;
      const startMatch = startCell.match(/([A-Z]+)(\d+)/);
      const endMatch = endCell.match(/([A-Z]+)(\d+)/);
      
      if (!startMatch || !endMatch) {
        return FormulaError.REF;
      }
      
      const startCol = startMatch[1].charCodeAt(0) - 'A'.charCodeAt(0);
      const endCol = endMatch[1].charCodeAt(0) - 'A'.charCodeAt(0);
      const startRow = parseInt(startMatch[2]);
      const endRow = parseInt(endMatch[2]);
      
      const cols = endCol - startCol + 1;
      const rows = endRow - startRow + 1;
      
      if (sortIndex < 1 || sortIndex > cols) {
        return FormulaError.VALUE;
      }
      
      // 2次元配列に変換
      const matrix: unknown[][] = [];
      for (let i = 0; i < rows; i++) {
        const row: unknown[] = [];
        for (let j = 0; j < cols; j++) {
          const index = i * cols + j;
          row.push(index < values.length ? values[index] : null);
        }
        matrix.push(row);
      }
      
      // 指定された列でソート
      const sortColIndex = sortIndex - 1;
      matrix.sort((rowA, rowB) => {
        const aVal = rowA[sortColIndex];
        const bVal = rowB[sortColIndex];
        
        const aStr = aVal === null || aVal === undefined ? '' : String(aVal);
        const bStr = bVal === null || bVal === undefined ? '' : String(bVal);
        
        // 数値として比較可能な場合
        const aNum = Number(aStr);
        const bNum = Number(bStr);
        if (!isNaN(aNum) && !isNaN(bNum)) {
          return sortOrder === 1 ? aNum - bNum : bNum - aNum;
        }
        
        // 文字列として比較
        return sortOrder === 1 ? aStr.localeCompare(bStr) : bStr.localeCompare(aStr);
      });
      
      // 1次元配列に戻す
      const result = matrix.flat();
      return result as FormulaResult;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// UNIQUE関数の実装
export const UNIQUE: CustomFormula = {
  name: 'UNIQUE',
  pattern: /UNIQUE\(([^,)]+)(?:,\s*([^,)]+))?(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, arrayRef, byColRef, exactlyOnceRef] = matches;
    
    try {
      // 配列データを取得
      const values = getCellRangeValues(arrayRef.trim(), context);
      
      if (!Array.isArray(values) || values.length === 0) {
        return FormulaError.VALUE;
      }
      
      // 列で一意性を判定するか（デフォルトはFALSE=行で判定）
      const byCol = byColRef ? (getCellValue(byColRef.trim(), context)?.toString().toUpperCase() === 'TRUE' || byColRef.trim() === '1') : false;
      
      // 一度だけ出現する要素のみを返すか（デフォルトはFALSE=最初の出現を残す）
      const exactlyOnce = exactlyOnceRef ? (getCellValue(exactlyOnceRef.trim(), context)?.toString().toUpperCase() === 'TRUE' || exactlyOnceRef.trim() === '1') : false;
      
      if (byCol) {
        // 列での一意性判定は複雑なため、簡単な実装
        return FormulaError.VALUE; // 未対応
      }
      
      // 単純な配列の場合
      const rangeParts = arrayRef.trim().split(':');
      if (rangeParts.length !== 2) {
        // 単一列の場合
        const uniqueValues: unknown[] = [];
        const seen = new Map<string, number>();
        
        // 出現回数をカウント
        for (const value of values) {
          const key = String(value ?? '');
          seen.set(key, (seen.get(key) ?? 0) + 1);
        }
        
        // exactlyOnceフラグに基づいて結果を決定
        for (const value of values) {
          const key = String(value ?? '');
          const count = seen.get(key) ?? 0;
          
          if (exactlyOnce) {
            // 一度だけ出現する要素
            if (count === 1 && !uniqueValues.some(v => String(v ?? '') === key)) {
              uniqueValues.push(value);
            }
          } else {
            // 最初の出現を残す
            if (!uniqueValues.some(v => String(v ?? '') === key)) {
              uniqueValues.push(value);
            }
          }
        }
        
        return uniqueValues.length === 1 ? uniqueValues[0] as FormulaResult : uniqueValues as FormulaResult;
      }
      
      // 2次元配列の場合（行での一意性判定）
      const [startCell, endCell] = rangeParts;
      const startMatch = startCell.match(/([A-Z]+)(\d+)/);
      const endMatch = endCell.match(/([A-Z]+)(\d+)/);
      
      if (!startMatch || !endMatch) {
        return FormulaError.REF;
      }
      
      const startCol = startMatch[1].charCodeAt(0) - 'A'.charCodeAt(0);
      const endCol = endMatch[1].charCodeAt(0) - 'A'.charCodeAt(0);
      const startRow = parseInt(startMatch[2]);
      const endRow = parseInt(endMatch[2]);
      
      const cols = endCol - startCol + 1;
      const rows = endRow - startRow + 1;
      
      // 2次元配列に変換
      const matrix: unknown[][] = [];
      for (let i = 0; i < rows; i++) {
        const row: unknown[] = [];
        for (let j = 0; j < cols; j++) {
          const index = i * cols + j;
          row.push(index < values.length ? values[index] : null);
        }
        matrix.push(row);
      }
      
      // 行の一意性を判定
      const uniqueRows: unknown[][] = [];
      const rowKeys = new Map<string, number>();
      
      // 各行をキーとして出現回数をカウント
      for (const row of matrix) {
        const key = row.map(cell => String(cell ?? '')).join('|');
        rowKeys.set(key, (rowKeys.get(key) ?? 0) + 1);
      }
      
      // exactlyOnceフラグに基づいて結果を決定
      for (const row of matrix) {
        const key = row.map(cell => String(cell ?? '')).join('|');
        const count = rowKeys.get(key) ?? 0;
        
        if (exactlyOnce) {
          // 一度だけ出現する行
          if (count === 1 && !uniqueRows.some(r => r.map(cell => String(cell ?? '')).join('|') === key)) {
            uniqueRows.push(row);
          }
        } else {
          // 最初の出現を残す
          if (!uniqueRows.some(r => r.map(cell => String(cell ?? '')).join('|') === key)) {
            uniqueRows.push(row);
          }
        }
      }
      
      // 1次元配列に戻す
      const result = uniqueRows.flat();
      return result as FormulaResult;
    } catch {
      return FormulaError.VALUE;
    }
  }
};