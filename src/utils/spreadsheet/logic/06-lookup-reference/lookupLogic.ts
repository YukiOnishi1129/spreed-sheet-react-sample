// 検索関数の実装

import type { CustomFormula, FormulaContext, FormulaResult } from '../shared/types';
import { FormulaError } from '../shared/types';
import { getCellValue, getCellRangeValues, parseCellRange } from '../shared/utils';
import { matchFormula } from '../index';

// VLOOKUP関数の実装
export const VLOOKUP: CustomFormula = {
  name: 'VLOOKUP',
  pattern: /VLOOKUP\(([^,]+),\s*([^,]+),\s*(\d+)(?:,\s*(TRUE|FALSE|0|1))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, lookupValue, tableArray, colIndex, rangeLookup] = matches;
    
    try {
      // 検索値を取得
      let searchValue = getCellValue(lookupValue.trim(), context);
      // If it's a string literal with quotes, remove them
      if (searchValue === lookupValue.trim()) {
        searchValue = lookupValue.trim().replace(/^["']|["']$/g, '');
      }
      
      // 列インデックスを数値に変換
      const columnIndex = parseInt(colIndex);
      if (columnIndex < 1) {
        return FormulaError.VALUE;
      }
      
      // 範囲を解析
      const rangeCoords = parseCellRange(tableArray.trim());
      if (!rangeCoords) {
        return FormulaError.REF;
      }
      
      const { start, end } = rangeCoords;
      const numCols = end.col - start.col + 1;
      
      if (columnIndex > numCols) {
        return FormulaError.REF;
      }
      
      // 範囲検索の設定（デフォルトはTRUE）
      const exactMatch = rangeLookup && (rangeLookup.toUpperCase() === 'FALSE' || rangeLookup === '0');
      
      // 各行を検索
      for (let row = start.row; row <= end.row; row++) {
        // 最初の列の値を取得
        const firstColValue = context.data[row]?.[start.col];
        let compareValue: unknown;
        
        // Handle both direct values and object values
        if (typeof firstColValue === 'string' || typeof firstColValue === 'number' || firstColValue === null || firstColValue === undefined) {
          compareValue = firstColValue;
        } else if (firstColValue && typeof firstColValue === 'object') {
          compareValue = firstColValue.value ?? firstColValue.v ?? firstColValue._ ?? firstColValue;
        } else {
          compareValue = firstColValue;
        }
        
        // Skip empty cells in the lookup column
        if (compareValue === '' || compareValue === null || compareValue === undefined) {
          continue;
        }
        
        if (exactMatch) {
          // 完全一致検索
          if (String(compareValue) === String(searchValue)) {
            // 指定された列の値を返す
            const targetCol = start.col + columnIndex - 1;
            const resultCell = context.data[row]?.[targetCol];
            
            // Handle both direct values and object values
            if (typeof resultCell === 'string' || typeof resultCell === 'number' || resultCell === null || resultCell === undefined) {
              return resultCell as FormulaResult;
            } else if (resultCell && typeof resultCell === 'object') {
              return (resultCell.value ?? resultCell.v ?? resultCell._ ?? resultCell) as FormulaResult;
            } else {
              return resultCell as FormulaResult;
            }
          }
        }
      }
      
      // 近似一致検索（Excelの正確な動作）
      if (!exactMatch) {
        let bestMatch: { row: number; value: unknown } | null = null;
        
        // データを昇順でソートされていると仮定し、最大の≤値を見つける
        for (let row = start.row; row <= end.row; row++) {
          const firstColValue = context.data[row]?.[start.col];
          let compareValue: unknown;
          
          // Handle both direct values and object values
          if (typeof firstColValue === 'string' || typeof firstColValue === 'number' || firstColValue === null || firstColValue === undefined) {
            compareValue = firstColValue;
          } else if (firstColValue && typeof firstColValue === 'object') {
            compareValue = firstColValue.value ?? firstColValue.v ?? firstColValue._ ?? firstColValue;
          } else {
            compareValue = firstColValue;
          }
          
          // 数値比較を試行
          const numericCompare = Number(compareValue);
          const numericSearch = Number(searchValue);
          
          if (!isNaN(numericCompare) && !isNaN(numericSearch)) {
            // 数値比較
            if (numericCompare <= numericSearch) {
              if (!bestMatch || Number(bestMatch.value) < numericCompare) {
                bestMatch = { row, value: compareValue };
              }
            }
          } else {
            // 文字列比較
            const stringCompare = String(compareValue);
            const stringSearch = String(searchValue);
            
            if (stringCompare <= stringSearch) {
              if (!bestMatch || String(bestMatch.value) < stringCompare) {
                bestMatch = { row, value: compareValue };
              }
            }
          }
        }
        
        if (bestMatch) {
          const targetCol = start.col + columnIndex - 1;
          const resultCell = context.data[bestMatch.row]?.[targetCol];
          
          // Handle both direct values and object values  
          if (typeof resultCell === 'string' || typeof resultCell === 'number' || resultCell === null || resultCell === undefined) {
            return resultCell as FormulaResult;
          } else if (resultCell && typeof resultCell === 'object') {
            return (resultCell.value ?? resultCell.v ?? resultCell._ ?? resultCell) as FormulaResult;
          } else {
            return resultCell as FormulaResult;
          }
        }
      }
      
      // 値が見つからない場合
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
      // Parse search value
      let searchValue = getCellValue(lookupValue.trim(), context);
      if (searchValue === lookupValue.trim()) {
        searchValue = lookupValue.trim().replace(/^["']|["']$/g, '');
      }
      
      // Parse row index
      const rowIndexNum = parseInt(rowIndex);
      if (rowIndexNum < 1) {
        return FormulaError.VALUE;
      }
      
      // Parse range coordinates
      const rangeCoords = parseCellRange(tableArray.trim());
      if (!rangeCoords) {
        return FormulaError.REF;
      }
      
      const { start, end } = rangeCoords;
      const rows = end.row - start.row + 1;
      const cols = end.col - start.col + 1;
      
      if (rowIndexNum > rows) {
        return FormulaError.REF;
      }
      
      // Determine search mode
      const exactMatch = rangeLookup && (rangeLookup.toUpperCase() === 'FALSE' || rangeLookup === '0');
      
      // Search in the first row
      for (let col = 0; col < cols; col++) {
        const cellRow = start.row;
        const cellCol = start.col + col;
        const cellValue = context.data[cellRow]?.[cellCol];
        let compareValue: unknown;
        
        // Handle both direct values and object values
        if (typeof cellValue === 'string' || typeof cellValue === 'number' || cellValue === null || cellValue === undefined) {
          compareValue = cellValue;
        } else if (cellValue && typeof cellValue === 'object') {
          compareValue = cellValue.value ?? cellValue.v ?? cellValue._ ?? cellValue;
        } else {
          compareValue = cellValue;
        }
        
        // Skip empty cells
        if (compareValue === '' || compareValue === null || compareValue === undefined) {
          continue;
        }
        
        if (exactMatch) {
          // Exact match search
          if (String(compareValue) === String(searchValue)) {
            // Get result from the specified row
            const resultRow = start.row + rowIndexNum - 1;
            const resultCol = start.col + col;
            const resultCell = context.data[resultRow]?.[resultCol];
            
            // Handle both direct values and object values
            if (typeof resultCell === 'string' || typeof resultCell === 'number' || resultCell === null || resultCell === undefined) {
              return resultCell as FormulaResult;
            } else if (resultCell && typeof resultCell === 'object') {
              return (resultCell.value ?? resultCell.v ?? resultCell._ ?? resultCell) as FormulaResult;
            } else {
              return resultCell as FormulaResult;
            }
          }
        }
      }
      
      // Handle approximate match (range lookup = TRUE, default)
      if (!exactMatch) {
        let bestMatch: { col: number; value: unknown } | null = null;
        
        // Find the largest value <= search value (assumes sorted data)
        for (let col = 0; col < cols; col++) {
          const cellRow = start.row;
          const cellCol = start.col + col;
          const cellValue = context.data[cellRow]?.[cellCol];
          let compareValue: unknown;
          
          if (typeof cellValue === 'string' || typeof cellValue === 'number' || cellValue === null || cellValue === undefined) {
            compareValue = cellValue;
          } else if (cellValue && typeof cellValue === 'object') {
            compareValue = cellValue.value ?? cellValue.v ?? cellValue._ ?? cellValue;
          } else {
            compareValue = cellValue;
          }
          
          // Skip empty cells
          if (compareValue === '' || compareValue === null || compareValue === undefined) {
            continue;
          }
          
          // Attempt numeric comparison first
          const numSearchValue = typeof searchValue === 'number' ? searchValue : parseFloat(String(searchValue));
          const numCompareValue = typeof compareValue === 'number' ? compareValue : parseFloat(String(compareValue));
          
          if (!isNaN(numSearchValue) && !isNaN(numCompareValue)) {
            if (numCompareValue <= numSearchValue) {
              if (!bestMatch || numCompareValue > (bestMatch.value as number)) {
                bestMatch = { col, value: numCompareValue };
              }
            }
          } else if (String(compareValue) <= String(searchValue)) {
            // String comparison
            if (!bestMatch || String(compareValue) > String(bestMatch.value)) {
              bestMatch = { col, value: compareValue };
            }
          }
        }
        
        if (bestMatch) {
          const resultRow = start.row + rowIndexNum - 1;
          const resultCol = start.col + bestMatch.col;
          const resultCell = context.data[resultRow]?.[resultCol];
          
          if (typeof resultCell === 'string' || typeof resultCell === 'number' || resultCell === null || resultCell === undefined) {
            return resultCell as FormulaResult;
          } else if (resultCell && typeof resultCell === 'object') {
            return (resultCell.value ?? resultCell.v ?? resultCell._ ?? resultCell) as FormulaResult;
          } else {
            return resultCell as FormulaResult;
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
      // 範囲を解析
      const rangeCoords = parseCellRange(arrayRange.trim());
      if (!rangeCoords) {
        return FormulaError.REF;
      }
      
      const { start, end } = rangeCoords;
      
      // 行番号を取得（1ベース）
      // MATCH関数の結果など、ネストされた関数の評価結果を処理
      let rowValue;
      const trimmedRowNum = rowNum.trim();
      
      // 数値の場合
      if (!isNaN(Number(trimmedRowNum))) {
        rowValue = Number(trimmedRowNum);
      } 
      // セル参照の場合
      else if (trimmedRowNum.match(/^[A-Z]+\d+$/)) {
        rowValue = getCellValue(trimmedRowNum, context);
      }
      // MATCH関数などの結果の場合
      else {
        // ネストされた関数の評価結果として数値が来ることを想定
        const nestedMatch = matchFormula(trimmedRowNum);
        if (nestedMatch) {
          rowValue = nestedMatch.function.calculate(nestedMatch.matches, context);
        } else {
          rowValue = trimmedRowNum;
        }
      }
      
      const rowIndex = parseInt(String(rowValue)) - 1;
      
      // 列番号を取得（省略可能、1ベース）
      let colIndex = 0;
      if (colNum) {
        const colValue = getCellValue(colNum.trim(), context) ?? colNum.trim();
        colIndex = parseInt(String(colValue)) - 1;
      }
      
      if (isNaN(rowIndex) || rowIndex < 0) {
        return FormulaError.VALUE;
      }
      
      if (isNaN(colIndex) || colIndex < 0) {
        return FormulaError.VALUE;
      }
      
      // 実際のセル位置を計算
      const targetRow = start.row + rowIndex;
      const targetCol = start.col + colIndex;
      
      // 範囲チェック
      if (targetRow > end.row || targetCol > end.col) {
        return FormulaError.REF;
      }
      
      // データを直接取得
      if (targetRow >= 0 && targetRow < context.data.length && 
          targetCol >= 0 && targetCol < context.data[0]?.length) {
        const cellData = context.data[targetRow][targetCol];
        
        // Handle both direct values and object values
        let result: FormulaResult;
        if (typeof cellData === 'string' || typeof cellData === 'number' || cellData === null || cellData === undefined) {
          result = cellData as FormulaResult;
        } else if (cellData && typeof cellData === 'object') {
          result = (cellData.value ?? cellData.v ?? cellData._ ?? cellData) as FormulaResult;
        } else {
          result = cellData as FormulaResult;
        }
        return result;
      }
      
      return FormulaError.REF;
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
      let searchValue;
      const trimmedLookup = lookupValue.trim();
      
      // 文字列リテラル（引用符で囲まれている）の場合
      if ((trimmedLookup.startsWith('"') && trimmedLookup.endsWith('"')) ||
          (trimmedLookup.startsWith("'") && trimmedLookup.endsWith("'"))) {
        searchValue = trimmedLookup.slice(1, -1);
      } else {
        // 数値の直接入力の場合
        const numValue = parseFloat(trimmedLookup);
        if (!isNaN(numValue)) {
          searchValue = numValue;
        } else {
          // セル参照の場合
          searchValue = getCellValue(trimmedLookup, context);
          if (searchValue === FormulaError.REF || searchValue === null || searchValue === undefined) {
            searchValue = trimmedLookup;
          }
        }
      }
      
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
            const arrayVal = arrayValues[i];
            // Try exact match first
            if (arrayVal === searchValue) {
              return i + 1; // 1ベースのインデックス
            }
            // Try string comparison
            if (String(arrayVal) === String(searchValue)) {
              return i + 1; // 1ベースのインデックス
            }
            // Try numeric comparison if both can be converted to numbers
            const numArray = Number(arrayVal);
            const numSearch = Number(searchValue);
            if (!isNaN(numArray) && !isNaN(numSearch) && numArray === numSearch) {
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
  pattern: /\bLOOKUP\(([^,]+),\s*([^,]+)(?:,\s*([^)]+))?\)/i,
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
  pattern: /\bXLOOKUP\(([^,]+),\s*([^,]+),\s*([^,]+)(?:,\s*([^,)]+))?(?:,\s*([^,)]+))?(?:,\s*([^)]+))?\)/i,
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
  pattern: /\bOFFSET\(([^,]+),\s*([^,]+),\s*([^,)]+)(?:,\s*([^,)]+))?(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, referenceRef, rowsRef, colsRef, heightRef, widthRef] = matches;
    
    try {
      // 基準セル参照を取得
      const reference = referenceRef.trim();
      let rowsValue: string | number = rowsRef.trim();
      let colsValue: string | number = colsRef.trim();
      
      // セル参照の場合のみ値を取得
      if (rowsRef.match(/^[A-Z]+\d+$/)) {
        const cellValue = getCellValue(rowsRef.trim(), context);
        if (cellValue !== FormulaError.REF) {
          rowsValue = cellValue as string | number;
        }
      }
      if (colsRef.match(/^[A-Z]+\d+$/)) {
        const cellValue = getCellValue(colsRef.trim(), context);
        if (cellValue !== FormulaError.REF) {
          colsValue = cellValue as string | number;
        }
      }
      const rows = typeof rowsValue === 'number' ? rowsValue : parseInt(rowsValue.toString());
      const cols = typeof colsValue === 'number' ? colsValue : parseInt(colsValue.toString());
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
  pattern: /\bINDIRECT\(([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, refTextRef, a1Ref] = matches;
    
    try {
      // 参照テキストを取得
      let refText = getCellValue(refTextRef.trim(), context) ?? refTextRef.trim();
      
      // 文字列の引用符を除去
      if (typeof refText === 'string') {
        refText = refText.replace(/^['"]|['"]$/g, '');
      }
      
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
        const result = getCellRangeValues(refText, context);
        // Check if it's a valid range
        if (!result || result.length === 0) {
          return FormulaError.REF;
        }
        return result as FormulaResult;
      } else {
        // 単一セル参照 - check if it's a valid cell reference
        const cellMatch = refText.match(/^[A-Z]+\d+$/);
        if (!cellMatch) {
          return FormulaError.REF;
        }
        const result = getCellValue(refText, context);
        // If getCellValue returns the same string, it means it's not a valid reference
        if (result === refText) {
          return FormulaError.REF;
        }
        return result as FormulaResult;
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
  pattern: /\bTRANSPOSE\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, arrayRef] = matches;
    
    try {
      // Check for invalid reference first
      const trimmedRef = arrayRef.trim();
      
      // Check if it's a string literal (should return REF error)
      if ((trimmedRef.startsWith('"') && trimmedRef.endsWith('"')) ||
          (trimmedRef.startsWith("'") && trimmedRef.endsWith("'"))) {
        return FormulaError.REF;
      }
      
      // 配列を取得
      const values = getCellRangeValues(trimmedRef, context);
      
      if (!Array.isArray(values) || values.length === 0) {
        return FormulaError.VALUE;
      }
      
      // 範囲から行数と列数を計算
      const rangeParts = trimmedRef.split(':');
      if (rangeParts.length !== 2) {
        // 単一セルの場合
        const singleValue = getCellValue(trimmedRef, context);
        return singleValue as FormulaResult;
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
      
      // 2次元配列として返す
      return transposed as FormulaResult;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// FILTER関数の実装
export const FILTER: CustomFormula = {
  name: 'FILTER',
  pattern: /\bFILTER\(([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, arrayRef, includeRef, ifEmptyRef] = matches;
    
    try {
      // Parse array range
      const arrayRange = parseCellRange(arrayRef.trim());
      if (!arrayRange) {
        return FormulaError.REF;
      }
      
      // Get array data
      const { start: arrayStart, end: arrayEnd } = arrayRange;
      const rows = arrayEnd.row - arrayStart.row + 1;
      const cols = arrayEnd.col - arrayStart.col + 1;
      
      // Extract array values
      const arrayData: unknown[][] = [];
      for (let r = 0; r < rows; r++) {
        const row: unknown[] = [];
        for (let c = 0; c < cols; c++) {
          const cellValue = context.data[arrayStart.row + r]?.[arrayStart.col + c];
          row.push(cellValue?.value ?? cellValue?.v ?? cellValue ?? null);
        }
        arrayData.push(row);
      }
      
      // Handle condition evaluation
      let conditionResults: boolean[] = [];
      
      // Check if includeRef contains a comparison operator
      const comparisonMatch = includeRef.trim().match(/^([^<>=!]+)([<>=!]+)(.+)$/);
      if (comparisonMatch) {
        const [, leftSide, operator, rightSide] = comparisonMatch;
        
        // Parse left side range
        const conditionRange = parseCellRange(leftSide.trim());
        if (!conditionRange) {
          return FormulaError.REF;
        }
        
        // Get right side value
        let rightValue: unknown;
        if (rightSide.trim().match(/^-?\d+(\.\d+)?$/)) {
          rightValue = parseFloat(rightSide.trim());
        } else {
          rightValue = getCellValue(rightSide.trim(), context);
        }
        
        // Extract condition values and evaluate
        const { start: condStart, end: condEnd } = conditionRange;
        const condRows = condEnd.row - condStart.row + 1;
        
        if (condRows !== rows) {
          return FormulaError.VALUE;
        }
        
        for (let r = 0; r < condRows; r++) {
          const condValue = context.data[condStart.row + r]?.[condStart.col];
          const leftValue = condValue?.value ?? condValue?.v ?? condValue ?? null;
          
          let result = false;
          if (leftValue !== null && rightValue !== null) {
            const left = typeof leftValue === 'number' ? leftValue : parseFloat(String(leftValue));
            const right = typeof rightValue === 'number' ? rightValue : parseFloat(String(rightValue));
            
            if (!isNaN(left) && !isNaN(right)) {
              switch (operator) {
                case '>': result = left > right; break;
                case '<': result = left < right; break;
                case '>=': result = left >= right; break;
                case '<=': result = left <= right; break;
                case '=': case '==': result = left === right; break;
                case '<>': case '!=': result = left !== right; break;
              }
            }
          }
          conditionResults.push(result);
        }
      } else {
        // Direct boolean array
        const includeValues = getCellRangeValues(includeRef.trim(), context);
        if (includeValues.length !== rows) {
          return FormulaError.VALUE;
        }
        
        conditionResults = includeValues.map(val => {
          return val === true || val === 1 || 
                 (typeof val === 'string' && val.toUpperCase() === 'TRUE') ||
                 (typeof val === 'number' && val !== 0);
        });
      }
      
      // Filter rows based on conditions
      const filteredRows: unknown[][] = [];
      for (let i = 0; i < rows; i++) {
        if (conditionResults[i]) {
          filteredRows.push(arrayData[i]);
        }
      }
      
      // Handle empty result
      if (filteredRows.length === 0) {
        if (ifEmptyRef) {
          let emptyValue = getCellValue(ifEmptyRef.trim(), context);
          if (emptyValue === ifEmptyRef.trim()) {
            emptyValue = ifEmptyRef.trim();
          }
          if (typeof emptyValue === 'string' && 
              ((emptyValue.startsWith('"') && emptyValue.endsWith('"')) ||
               (emptyValue.startsWith("'") && emptyValue.endsWith("'")))) {
            emptyValue = emptyValue.slice(1, -1);
          }
          return emptyValue as FormulaResult;
        }
        return FormulaError.CALC;
      }
      
      // Return results
      if (cols > 1) {
        return filteredRows as FormulaResult;
      } else {
        // Single column - return as spill array
        return filteredRows.map(row => row[0]) as FormulaResult;
      }
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// SORT関数の実装
export const SORT: CustomFormula = {
  name: 'SORT',
  pattern: /\bSORT\(([^,)]+)(?:,\s*([^,)]+))?(?:,\s*([^,)]+))?(?:,\s*([^)]+))?\)/i,
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
  pattern: /\bUNIQUE\(([^,)]+)(?:,\s*([^,)]+))?(?:,\s*([^)]+))?\)/i,
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
      // const rangeParts = arrayRef.trim().split(':');
      // Always treat as single dimensional array for now
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
      
      return uniqueValues as FormulaResult;
    } catch {
      return FormulaError.VALUE;
    }
  }
};