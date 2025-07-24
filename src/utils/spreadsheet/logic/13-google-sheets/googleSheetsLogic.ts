// Google Sheets specific functions

import type { CustomFormula, FormulaContext, FormulaResult } from '../shared/types';
import { FormulaError } from '../shared/types';
import { getCellValue, getCellRangeValues } from '../shared/utils';

// JOIN関数の実装（配列を文字列に結合）
export const JOIN: CustomFormula = {
  name: 'JOIN',
  pattern: /JOIN\(([^,]+),\s*(.+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, delimiterRef, valuesRef] = matches;
    
    try {
      // デリミターを取得
      let delimiter = getCellValue(delimiterRef.trim(), context)?.toString() ?? delimiterRef.trim();
      if (delimiter.startsWith('"') && delimiter.endsWith('"')) {
        delimiter = delimiter.slice(1, -1);
      }
      
      const values: string[] = [];
      
      // 値を収集
      if (valuesRef.includes(':')) {
        // セル範囲の場合
        const rangeValues = getCellRangeValues(valuesRef.trim(), context);
        rangeValues.forEach(value => {
          if (value !== null && value !== undefined && value !== '') {
            values.push(String(value));
          }
        });
      } else {
        // 複数の引数の場合
        const args = valuesRef.split(',').map(arg => arg.trim());
        for (const arg of args) {
          let value: unknown;
          if (arg.match(/^[A-Z]+\d+$/)) {
            value = getCellValue(arg, context);
          } else if (arg.startsWith('"') && arg.endsWith('"')) {
            value = arg.slice(1, -1);
          } else {
            value = arg;
          }
          
          if (value !== null && value !== undefined && value !== '') {
            values.push(String(value));
          }
        }
      }
      
      return values.join(delimiter);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// ARRAYFORMULA関数の実装（配列数式を作成）
export const ARRAYFORMULA: CustomFormula = {
  name: 'ARRAYFORMULA',
  pattern: /ARRAYFORMULA\((.+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const formulaContent = matches[1].trim();
    
    try {
      // 簡単な実装：セル範囲の演算をサポート
      // 例: ARRAYFORMULA(A1:A5 * B1:B5)
      const operationMatch = formulaContent.match(/([A-Z]+\d+:[A-Z]+\d+)\s*([+\-*/])\s*([A-Z]+\d+:[A-Z]+\d+)/);
      
      if (operationMatch) {
        const [, range1Ref, operator, range2Ref] = operationMatch;
        
        const range1Values = getCellRangeValues(range1Ref, context);
        const range2Values = getCellRangeValues(range2Ref, context);
        
        if (range1Values.length !== range2Values.length) {
          return FormulaError.VALUE;
        }
        
        const results: number[] = [];
        
        for (let i = 0; i < range1Values.length; i++) {
          const val1 = Number(range1Values[i]);
          const val2 = Number(range2Values[i]);
          
          if (isNaN(val1) || isNaN(val2)) {
            return FormulaError.VALUE;
          }
          
          let result: number;
          switch (operator) {
            case '+':
              result = val1 + val2;
              break;
            case '-':
              result = val1 - val2;
              break;
            case '*':
              result = val1 * val2;
              break;
            case '/':
              if (val2 === 0) return FormulaError.DIV0;
              result = val1 / val2;
              break;
            default:
              return FormulaError.VALUE;
          }
          
          results.push(result);
        }
        
        // 結果を2次元配列として返す（列ベクトル）
        return results.map(val => [val]) as number[][];
      }
      
      // 単一セル範囲の場合はそのまま返す
      if (formulaContent.match(/^[A-Z]+\d+:[A-Z]+\d+$/)) {
        const values = getCellRangeValues(formulaContent, context);
        return values.map(val => [val]) as (string | number)[][];
      }
      
      // その他の場合は未サポート
      return FormulaError.VALUE;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// QUERY関数の実装（簡易版 - SQLライクなクエリ）
export const QUERY: CustomFormula = {
  name: 'QUERY',
  pattern: /QUERY\(([^,]+)(?:,\s*"([^"]*)")?(?:,\s*(-?\d+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, dataRef, queryString = '', headersParam = '1'] = matches;
    
    try {
      const headers = parseInt(headersParam);
      
      // データ範囲を取得
      if (!dataRef.includes(':')) {
        return FormulaError.REF;
      }
      
      // 範囲を解析
      const rangeMatch = dataRef.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
      if (!rangeMatch) {
        return FormulaError.REF;
      }
      
      const [, startCol, startRowStr, endCol, endRowStr] = rangeMatch;
      const startRow = parseInt(startRowStr) - 1;
      const endRow = parseInt(endRowStr) - 1;
      const startColIndex = startCol.charCodeAt(0) - 65;
      const endColIndex = endCol.charCodeAt(0) - 65;
      
      // データを2次元配列として取得
      const data: unknown[][] = [];
      for (let row = startRow; row <= endRow; row++) {
        const rowData: unknown[] = [];
        for (let col = startColIndex; col <= endColIndex; col++) {
          if (context.data[row]?.[col]) {
            rowData.push(context.data[row][col].value);
          } else {
            rowData.push(null);
          }
        }
        data.push(rowData);
      }
      
      if (data.length === 0) {
        return [];
      }
      
      // クエリが空の場合は全データを返す
      if (!queryString) {
        return data as (string | number)[][];
      }
      
      // 簡単なSELECTクエリをサポート
      const selectMatch = queryString.match(/SELECT\s+(.+?)(?:\s+WHERE\s+(.+?))?(?:\s+ORDER\s+BY\s+(.+?))?(?:\s+LIMIT\s+(\d+))?$/i);
      
      if (!selectMatch) {
        return FormulaError.VALUE;
      }
      
      const [, selectClause, whereClause, orderByClause, limitClause] = selectMatch;
      
      let result = data;
      const headerRow = headers > 0 ? data.slice(0, headers) : [];
      let dataRows = headers > 0 ? data.slice(headers) : data;
      
      // WHERE句の処理（簡易版）
      if (whereClause) {
        const whereMatch = whereClause.match(/([A-Z])\s*([=<>]+)\s*(.+)/i);
        if (whereMatch) {
          const [, colLetter, operator, valueStr] = whereMatch;
          const colIndex = colLetter.charCodeAt(0) - 65;
          
          let compareValue: unknown = valueStr.trim();
          if (typeof compareValue === 'string' && compareValue.startsWith("'") && compareValue.endsWith("'")) {
            compareValue = compareValue.slice(1, -1);
          } else {
            const num = parseFloat(compareValue as string);
            if (!isNaN(num)) {
              compareValue = num;
            }
          }
          
          dataRows = dataRows.filter(row => {
            const cellValue = row[colIndex];
            
            switch (operator) {
              case '=':
                return cellValue === compareValue;
              case '<>':
              case '!=':
                return cellValue !== compareValue;
              case '>':
                return Number(cellValue) > Number(compareValue);
              case '<':
                return Number(cellValue) < Number(compareValue);
              case '>=':
                return Number(cellValue) >= Number(compareValue);
              case '<=':
                return Number(cellValue) <= Number(compareValue);
              default:
                return false;
            }
          });
        }
      }
      
      // ORDER BY句の処理（簡易版）
      if (orderByClause) {
        const orderMatch = orderByClause.match(/([A-Z])(?:\s+(ASC|DESC))?/i);
        if (orderMatch) {
          const [, colLetter, direction = 'ASC'] = orderMatch;
          const colIndex = colLetter.charCodeAt(0) - 65;
          
          dataRows.sort((a, b) => {
            const aVal = a[colIndex];
            const bVal = b[colIndex];
            
            let comparison = 0;
            if (typeof aVal === 'number' && typeof bVal === 'number') {
              comparison = aVal - bVal;
            } else {
              comparison = String(aVal).localeCompare(String(bVal));
            }
            
            return direction.toUpperCase() === 'DESC' ? -comparison : comparison;
          });
        }
      }
      
      // LIMIT句の処理
      if (limitClause) {
        const limit = parseInt(limitClause);
        dataRows = dataRows.slice(0, limit);
      }
      
      // SELECT句の処理（簡易版 - カラム選択）
      if (selectClause !== '*') {
        const columns = selectClause.split(',').map(col => col.trim());
        const colIndices: number[] = [];
        
        for (const col of columns) {
          if (col.match(/^[A-Z]$/i)) {
            colIndices.push(col.charCodeAt(0) - 65);
          }
        }
        
        if (colIndices.length > 0) {
          dataRows = dataRows.map(row => 
            colIndices.map(idx => row[idx])
          );
        }
      }
      
      // ヘッダーとデータを結合して返す
      result = [...headerRow, ...dataRows];
      
      return result as (string | number)[][];
    } catch {
      return FormulaError.VALUE;
    }
  }
};


// REGEXMATCH関数の実装（正規表現マッチング）
export const REGEXMATCH: CustomFormula = {
  name: 'REGEXMATCH',
  pattern: /REGEXMATCH\(([^,]+),\s*"([^"]+)"\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, textRef, regexPattern] = matches;
    
    try {
      // テキストを取得
      let text = getCellValue(textRef.trim(), context)?.toString() ?? textRef.trim();
      if (text.startsWith('"') && text.endsWith('"')) {
        text = text.slice(1, -1);
      }
      
      // 正規表現を作成してマッチング
      const regex = new RegExp(regexPattern);
      return regex.test(text);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// REGEXEXTRACT関数の実装（正規表現による抽出）
export const REGEXEXTRACT: CustomFormula = {
  name: 'REGEXEXTRACT',
  pattern: /REGEXEXTRACT\(([^,]+),\s*"([^"]+)"\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, textRef, regexPattern] = matches;
    
    try {
      // テキストを取得
      let text = getCellValue(textRef.trim(), context)?.toString() ?? textRef.trim();
      if (text.startsWith('"') && text.endsWith('"')) {
        text = text.slice(1, -1);
      }
      
      // 正規表現を作成してマッチング
      const regex = new RegExp(regexPattern);
      const matchResult = text.match(regex);
      
      if (matchResult) {
        // キャプチャグループがある場合は最初のグループを返す
        return matchResult[1] || matchResult[0];
      }
      
      return FormulaError.NA;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// REGEXREPLACE関数の実装（正規表現による置換）
export const REGEXREPLACE: CustomFormula = {
  name: 'REGEXREPLACE',
  pattern: /REGEXREPLACE\(([^,]+),\s*"([^"]+)",\s*"([^"]*)"\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, textRef, regexPattern, replacement] = matches;
    
    try {
      // テキストを取得
      let text = getCellValue(textRef.trim(), context)?.toString() ?? textRef.trim();
      if (text.startsWith('"') && text.endsWith('"')) {
        text = text.slice(1, -1);
      }
      
      // 正規表現を作成して置換
      const regex = new RegExp(regexPattern, 'g');
      return text.replace(regex, replacement);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// FLATTEN関数の実装（配列を1次元化）
export const FLATTEN: CustomFormula = {
  name: 'FLATTEN',
  pattern: /FLATTEN\((.+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const rangeRef = matches[1].trim();
    
    try {
      const values: unknown[] = [];
      
      if (rangeRef.includes(':')) {
        // 範囲を解析
        const rangeMatch = rangeRef.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
        if (!rangeMatch) {
          return FormulaError.REF;
        }
        
        const [, startCol, startRowStr, endCol, endRowStr] = rangeMatch;
        const startRow = parseInt(startRowStr) - 1;
        const endRow = parseInt(endRowStr) - 1;
        const startColIndex = startCol.charCodeAt(0) - 65;
        const endColIndex = endCol.charCodeAt(0) - 65;
        
        // 行ごと、列ごとに値を収集
        for (let row = startRow; row <= endRow; row++) {
          for (let col = startColIndex; col <= endColIndex; col++) {
            if (context.data[row]?.[col]) {
              const value = context.data[row][col].value;
              if (value !== null && value !== undefined && value !== '') {
                values.push(value);
              }
            }
          }
        }
      } else {
        // 単一セルの場合
        const value = getCellValue(rangeRef, context);
        if (value !== null && value !== undefined && value !== '') {
          values.push(value);
        }
      }
      
      // 縦方向の配列として返す
      return values.map(val => [val]) as (string | number)[][];
    } catch {
      return FormulaError.VALUE;
    }
  }
};