// Lambda配列関数の実装

import type { CustomFormula, FormulaContext, FormulaResult } from '../shared/types';
import { FormulaError } from '../shared/types';
import { getCellValue, parseCellRange, evaluateExpression, toNumber } from '../shared/utils';
import { matchFormula } from '../index';

// Lambda式の評価を行うヘルパー関数
const evaluateLambda = (
  lambdaExpression: string,
  paramNames: string[],
  paramValues: unknown[],
  context: FormulaContext
): FormulaResult => {
  try {
    // LAMBDA関数の形式を解析
    const lambdaMatch = lambdaExpression.match(/LAMBDA\(([^,)]+(?:,[^,)]+)*),\s*(.+)\)$/i);
    if (lambdaMatch) {
      const [, params, body] = lambdaMatch;
      const lambdaParams = params.split(',').map(p => p.trim());
      
      // パラメータを実際の値に置換
      let evaluatedBody = body.trim();
      
      // 単一のパラメータが配列の場合、特別な処理
      if (lambdaParams.length === 1 && paramValues.length === 1 && Array.isArray(paramValues[0])) {
        const arrayParam = paramValues[0] as unknown[];
        
        // SUM(row)のような関数呼び出しを検出
        if (evaluatedBody.toUpperCase() === `SUM(${lambdaParams[0].toUpperCase()})`) {
          const sum = arrayParam.reduce((acc: number, val) => {
            const num = toNumber(val);
            return num !== null ? acc + num : acc;
          }, 0);
          return sum;
        }
        
        if (evaluatedBody.toUpperCase() === `AVERAGE(${lambdaParams[0].toUpperCase()})`) {
          const nums = arrayParam.map(val => toNumber(val)).filter(n => n !== null) as number[];
          return nums.length > 0 ? nums.reduce((a, b) => a + b, 0) / nums.length : 0;
        }
        
        if (evaluatedBody.toUpperCase() === `COUNT(${lambdaParams[0].toUpperCase()})`) {
          return arrayParam.filter(val => toNumber(val) !== null).length;
        }
        
        if (evaluatedBody.toUpperCase() === `MAX(${lambdaParams[0].toUpperCase()})`) {
          const nums = arrayParam.map(val => toNumber(val)).filter(n => n !== null) as number[];
          return nums.length > 0 ? Math.max(...nums) : 0;
        }
        
        if (evaluatedBody.toUpperCase() === `MIN(${lambdaParams[0].toUpperCase()})`) {
          const nums = arrayParam.map(val => toNumber(val)).filter(n => n !== null) as number[];
          return nums.length > 0 ? Math.min(...nums) : 0;
        }
      }
      
      // 通常のパラメータ置換
      lambdaParams.forEach((param, index) => {
        if (index < paramValues.length) {
          const value = paramValues[index];
          const regex = new RegExp(`\\b${param}\\b`, 'gi');
          
          // 配列の場合はJSON形式で置換
          if (Array.isArray(value)) {
            evaluatedBody = evaluatedBody.replace(regex, JSON.stringify(value));
          } else {
            evaluatedBody = evaluatedBody.replace(regex, String(value));
          }
        }
      });
      
      // 特別な関数を含む場合はフォーミュラ評価へ
      if (evaluatedBody.toUpperCase().includes('MAX(') || 
          evaluatedBody.toUpperCase().includes('MIN(') ||
          evaluatedBody.toUpperCase().includes('SUM(') ||
          evaluatedBody.toUpperCase().includes('AVERAGE(')) {
        return evaluateFormulaExpression(evaluatedBody, context);
      }
      
      // 簡単な演算子を含む式の評価
      if (/^[a-z0-9\s+\-*/().,]+$/i.test(evaluatedBody)) {
        try {
          // 安全な評価のためにeval の代わりに計算を実行
          // eslint-disable-next-line @typescript-eslint/no-unsafe-call, @typescript-eslint/no-implied-eval
          const result = new Function('return (' + evaluatedBody + ')')() as unknown;
          if (typeof result === 'number' && !isNaN(result)) {
            return result;
          }
        } catch {
          // 評価に失敗した場合は下記の処理に進む
        }
      }
      
      // 関数式を評価
      return evaluateFormulaExpression(evaluatedBody, context);
    }
    
    // LAMBDA関数でない場合、単純な関数名の可能性
    const functionName = lambdaExpression.trim().toUpperCase();
    
    // 組み込み関数を処理
    if (functionName === 'SUM' || lambdaExpression.toUpperCase().includes('SUM(')) {
      const sum = paramValues.reduce((acc: number, val) => {
        const num = toNumber(val);
        return num !== null ? acc + num : acc;
      }, 0);
      return sum;
    }
    
    if (functionName === 'AVERAGE' || lambdaExpression.toUpperCase().includes('AVERAGE(')) {
      const nums = paramValues.map(val => toNumber(val)).filter(n => n !== null) as number[];
      return nums.length > 0 ? nums.reduce((a, b) => a + b, 0) / nums.length : 0;
    }
    
    if (functionName === 'COUNT' || lambdaExpression.toUpperCase().includes('COUNT(')) {
      return paramValues.filter(val => toNumber(val) !== null).length;
    }
    
    if (functionName === 'MAX' || lambdaExpression.toUpperCase().includes('MAX(')) {
      const nums = paramValues.map(val => toNumber(val)).filter(n => n !== null) as number[];
      return nums.length > 0 ? Math.max(...nums) : 0;
    }
    
    if (functionName === 'MIN' || lambdaExpression.toUpperCase().includes('MIN(')) {
      const nums = paramValues.map(val => toNumber(val)).filter(n => n !== null) as number[];
      return nums.length > 0 ? Math.min(...nums) : 0;
    }
    
    // その他の関数式を評価
    let expression = lambdaExpression;
    paramNames.forEach((param, index) => {
      if (index < paramValues.length) {
        const value = paramValues[index];
        const regex = new RegExp(`\\b${param}\\b`, 'gi');
        expression = expression.replace(regex, String(value));
      }
    });
    
    return evaluateFormulaExpression(expression, context);
  } catch {
    return FormulaError.VALUE;
  }
};

// 数式を評価するヘルパー関数
const evaluateFormulaExpression = (expression: string, context: FormulaContext): FormulaResult => {
  try {
    // MAX関数の特別処理
    const maxMatch = expression.match(/MAX\(([^,]+),\s*([^)]+)\)/i);
    if (maxMatch) {
      const [, arg1, arg2] = maxMatch;
      const val1 = isNaN(Number(arg1)) ? getCellValue(arg1.trim(), context) : Number(arg1);
      const val2 = isNaN(Number(arg2)) ? getCellValue(arg2.trim(), context) : Number(arg2);
      const num1 = toNumber(val1);
      const num2 = toNumber(val2);
      if (num1 !== null && num2 !== null) {
        return Math.max(num1, num2);
      }
    }
    
    // MIN関数の特別処理
    const minMatch = expression.match(/MIN\(([^,]+),\s*([^)]+)\)/i);
    if (minMatch) {
      const [, arg1, arg2] = minMatch;
      const val1 = isNaN(Number(arg1)) ? getCellValue(arg1.trim(), context) : Number(arg1);
      const val2 = isNaN(Number(arg2)) ? getCellValue(arg2.trim(), context) : Number(arg2);
      const num1 = toNumber(val1);
      const num2 = toNumber(val2);
      if (num1 !== null && num2 !== null) {
        return Math.min(num1, num2);
      }
    }
    
    // 組み込み関数をチェック
    const matchResult = matchFormula(expression.toUpperCase());
    if (matchResult) {
      const { function: func, matches } = matchResult;
      return func.calculate(matches, context);
    }
    
    // 単純な算術式を評価
    const result = evaluateExpression(expression, context);
    return result;
  } catch {
    return FormulaError.VALUE;
  }
};

// 配列データを取得するヘルパー関数
const getArrayData = (arrayRef: string, context: FormulaContext): (number | string | boolean | null)[][] => {
  const rangeCoords = parseCellRange(arrayRef.trim());
  if (!rangeCoords) {
    throw new Error('Invalid range');
  }
  
  const { start, end } = rangeCoords;
  const data: (number | string | boolean | null)[][] = [];
  
  for (let row = start.row; row <= end.row && row < context.data.length; row++) {
    const rowData: (number | string | boolean | null)[] = [];
    for (let col = start.col; col <= end.col && col < context.data[row]?.length; col++) {
      const cellValue = getCellValue(`${String.fromCharCode(65 + col)}${row + 1}`, context);
      rowData.push(cellValue as (number | string | boolean | null));
    }
    data.push(rowData);
  }
  
  return data;
};

// BYROW関数の実装（行ごとに関数を適用）
export const BYROW: CustomFormula = {
  name: 'BYROW',
  pattern: /BYROW\(([^,]+),\s*(.+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, arrayRef, lambdaRef] = matches;
    
    try {
      // データ配列を取得
      const data = getArrayData(arrayRef, context);
      const results: (number | string | boolean | null)[][] = [];
      
      // 各行に対してLambda関数を適用
      for (const row of data) {
        const result = evaluateLambda(lambdaRef.trim(), ['row'], [row], context);
        
        // 結果を単一セルの配列として追加
        if (Array.isArray(result) && Array.isArray(result[0])) {
          // 二次元配列の場合は最初の要素を取得
          results.push([(result[0] as (number | string | boolean | null)[])[0]]);
        } else if (Array.isArray(result)) {
          results.push([result[0] as (number | string | boolean | null)]);
        } else {
          results.push([result as (number | string | boolean | null)]);
        }
      }
      
      return results;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// BYCOL関数の実装（列ごとに関数を適用）
export const BYCOL: CustomFormula = {
  name: 'BYCOL',
  pattern: /BYCOL\(([^,]+),\s*(.+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, arrayRef, lambdaRef] = matches;
    
    try {
      // データ配列を取得
      const data = getArrayData(arrayRef, context);
      
      if (data.length === 0 || data[0].length === 0) {
        return FormulaError.VALUE;
      }
      
      const results: (number | string | boolean | null)[][] = [[]]; // 1行の結果配列
      const numCols = data[0].length;
      
      // 各列に対してLambda関数を適用
      for (let col = 0; col < numCols; col++) {
        const column = data.map(row => row[col]);
        const result = evaluateLambda(lambdaRef.trim(), ['column'], [column], context);
        
        // 結果を横一列に追加
        if (Array.isArray(result) && Array.isArray(result[0])) {
          // 二次元配列の場合は最初の要素を取得
          results[0].push((result[0] as (number | string | boolean | null)[])[0]);
        } else if (Array.isArray(result)) {
          results[0].push(result[0] as (number | string | boolean | null));
        } else {
          results[0].push(result as (number | string | boolean | null));
        }
      }
      
      return results;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// MAP関数の実装（配列の各要素に関数を適用）
export const MAP: CustomFormula = {
  name: 'MAP',
  pattern: /MAP\((.+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, args] = matches;
    
    try {
      // 引数を解析（複数の配列とLambda関数）
      const argParts = parseArguments(args);
      if (argParts.length < 2) {
        return FormulaError.VALUE;
      }
      
      const lambdaRef = argParts[argParts.length - 1];
      const arrayRefs = argParts.slice(0, -1);
      
      // 各配列のデータを取得
      const arrays = arrayRefs.map(ref => getArrayData(ref, context));
      
      if (arrays.length === 0) {
        return FormulaError.VALUE;
      }
      
      // すべての配列のサイズを確認
      const rows = arrays[0].length;
      const cols = arrays[0][0]?.length ?? 0;
      
      for (const arr of arrays) {
        if (arr.length !== rows || (arr[0]?.length ?? 0) !== cols) {
          return FormulaError.VALUE; // サイズが一致しない
        }
      }
      
      const results: (number | string | boolean | null)[][] = [];
      
      // 各要素に対してLambda関数を適用
      for (let row = 0; row < rows; row++) {
        const resultRow: (number | string | boolean | null)[] = [];
        for (let col = 0; col < cols; col++) {
          // 各配列から対応する要素を取得
          const values = arrays.map(arr => arr[row][col]);
          const result = evaluateLambda(lambdaRef.trim(), ['a', 'b', 'c', 'd', 'e'].slice(0, values.length), values, context);
          
          if (Array.isArray(result) && Array.isArray(result[0])) {
            // 二次元配列の場合は最初の要素を取得
            resultRow.push((result[0] as (number | string | boolean | null)[])[0]);
          } else if (Array.isArray(result)) {
            resultRow.push(result[0] as (number | string | boolean | null));
          } else {
            resultRow.push(result as (number | string | boolean | null));
          }
        }
        results.push(resultRow);
      }
      
      return results;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// 引数を解析するヘルパー関数
const parseArguments = (args: string): string[] => {
  const result: string[] = [];
  let current = '';
  let depth = 0;
  let inQuote = false;
  
  for (let i = 0; i < args.length; i++) {
    const char = args[i];
    
    if (char === '"' && (i === 0 || args[i - 1] !== '\\')) {
      inQuote = !inQuote;
    }
    
    if (!inQuote) {
      if (char === '(') depth++;
      else if (char === ')') depth--;
      else if (char === ',' && depth === 0) {
        result.push(current.trim());
        current = '';
        continue;
      }
    }
    
    current += char;
  }
  
  if (current.trim()) {
    result.push(current.trim());
  }
  
  return result;
};

// REDUCE関数の実装（配列を単一の値に縮約）
export const REDUCE: CustomFormula = {
  name: 'REDUCE',
  pattern: /REDUCE\((.+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, args] = matches;
    
    try {
      const argParts = parseArguments(args);
      if (argParts.length < 3) {
        return FormulaError.VALUE;
      }
      
      const [initialRef, arrayRef, lambdaRef] = argParts;
      
      // 初期値を取得
      let initial: unknown;
      if (initialRef.match(/^-?\d+(\.\d+)?$/)) {
        initial = parseFloat(initialRef);
      } else if (initialRef.startsWith('"') && initialRef.endsWith('"')) {
        initial = initialRef.slice(1, -1);
      } else {
        initial = getCellValue(initialRef, context);
      }
      
      // データ配列を取得（フラット化）
      const data = getArrayData(arrayRef, context);
      const values: (number | string | boolean | null)[] = [];
      
      for (const row of data) {
        for (const value of row) {
          if (value !== null && value !== '') {
            values.push(value);
          }
        }
      }
      
      // Lambda関数を順次適用
      let accumulator = initial;
      
      // 特別なケース: &演算子の処理
      if (lambdaRef.includes('&')) {
        for (const value of values) {
          // 文字列連結
          accumulator = String(accumulator) + String(value);
        }
        return accumulator as string;
      }
      
      for (const value of values) {
        accumulator = evaluateLambda(lambdaRef.trim(), ['accumulator', 'value'], [accumulator, value], context);
        
        // エラーが発生した場合は即座に返す
        if (typeof accumulator === 'string' && accumulator.startsWith('#')) {
          return accumulator;
        }
      }
      
      // accumulator の型をチェックして適切な FormulaResult を返す
      if (typeof accumulator === 'string' || typeof accumulator === 'number' || typeof accumulator === 'boolean' || accumulator === null) {
        return accumulator;
      } else if (Array.isArray(accumulator)) {
        // 配列の場合はそのまま返せない可能性があるため、最初の要素を返す
        if (Array.isArray(accumulator[0])) {
          return accumulator as (number | string | boolean | null)[][];
        } else {
          return accumulator as (number | string | boolean | null)[];
        }
      } else if (accumulator instanceof Date) {
        return accumulator;
      } else {
        return FormulaError.VALUE;
      }
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// SCAN関数の実装（配列の累積計算）
export const SCAN: CustomFormula = {
  name: 'SCAN',
  pattern: /SCAN\((.+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, args] = matches;
    
    try {
      const argParts = parseArguments(args);
      if (argParts.length < 3) {
        return FormulaError.VALUE;
      }
      
      const [initialRef, arrayRef, lambdaRef] = argParts;
      
      // 初期値を取得
      let initial: unknown;
      if (initialRef.match(/^-?\d+(\.\d+)?$/)) {
        initial = parseFloat(initialRef);
      } else if (initialRef.startsWith('"') && initialRef.endsWith('"')) {
        initial = initialRef.slice(1, -1);
      } else {
        initial = getCellValue(initialRef, context);
      }
      
      // データ配列を取得
      const data = getArrayData(arrayRef, context);
      const results: (number | string | boolean | null)[][] = [];
      let accumulator = initial;
      
      // 各要素に対して累積計算を実行
      for (let row = 0; row < data.length; row++) {
        const resultRow: (number | string | boolean | null)[] = [];
        for (let col = 0; col < data[row].length; col++) {
          const value = data[row][col];
          
          if (value !== null && value !== '') {
            accumulator = evaluateLambda(lambdaRef.trim(), ['accumulator', 'value'], [accumulator, value], context);
            
            // エラーが発生した場合
            if (typeof accumulator === 'string' && accumulator.startsWith('#')) {
              resultRow.push(accumulator as string);
            } else if (Array.isArray(accumulator)) {
              resultRow.push(accumulator[0] as (number | string | boolean | null));
            } else {
              resultRow.push(accumulator as (number | string | boolean | null));
            }
          } else {
            // 空のセルの場合は前の累積値を保持
            if (Array.isArray(accumulator)) {
              resultRow.push(accumulator[0] as (number | string | boolean | null));
            } else {
              resultRow.push(accumulator as (number | string | boolean | null));
            }
          }
        }
        results.push(resultRow);
      }
      
      return results;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// MAKEARRAY関数の実装（指定サイズの配列を作成）
export const MAKEARRAY: CustomFormula = {
  name: 'MAKEARRAY',
  pattern: /MAKEARRAY\((.+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, args] = matches;
    
    try {
      const argParts = parseArguments(args);
      if (argParts.length < 3) {
        return FormulaError.VALUE;
      }
      
      const [rowsRef, colsRef, lambdaRef] = argParts;
      
      // 行数と列数を取得
      let rows: number;
      let cols: number;
      
      if (rowsRef.match(/^\d+$/)) {
        rows = parseInt(rowsRef);
      } else {
        const rowsValue = getCellValue(rowsRef, context);
        rows = parseInt(String(rowsValue));
      }
      
      if (colsRef.match(/^\d+$/)) {
        cols = parseInt(colsRef);
      } else {
        const colsValue = getCellValue(colsRef, context);
        cols = parseInt(String(colsValue));
      }
      
      if (isNaN(rows) || isNaN(cols) || rows <= 0 || cols <= 0 || rows > 1048576 || cols > 16384) {
        return FormulaError.VALUE;
      }
      
      const results: (number | string | boolean | null)[][] = [];
      
      // 各セルに対してLambda関数を適用
      for (let row = 0; row < rows; row++) {
        const resultRow: (number | string | boolean | null)[] = [];
        for (let col = 0; col < cols; col++) {
          // Lambda関数内でROW()とCOLUMN()を使用できるようにコンテキストを準備
          const lambdaExpression = lambdaRef.trim()
            .replace(/\bROW\(\)/gi, String(row + 1))
            .replace(/\bCOLUMN\(\)/gi, String(col + 1));
          
          const result = evaluateLambda(lambdaExpression, ['row', 'column'], [row + 1, col + 1], context);
          
          if (Array.isArray(result) && Array.isArray(result[0])) {
            // 二次元配列の場合は最初の要素を取得
            resultRow.push((result[0] as (number | string | boolean | null)[])[0]);
          } else if (Array.isArray(result)) {
            resultRow.push(result[0] as (number | string | boolean | null));
          } else {
            resultRow.push(result as (number | string | boolean | null));
          }
        }
        results.push(resultRow);
      }
      
      return results;
    } catch {
      return FormulaError.VALUE;
    }
  }
};