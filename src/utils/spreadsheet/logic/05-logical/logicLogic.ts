// 論理関数の実装

import type { CustomFormula, FormulaContext, FormulaResult } from '../shared/types';
import { FormulaError } from '../shared/types';
import { getCellValue } from '../shared/utils';

// 引数を論理値に変換するユーティリティ関数
function parseArgumentsToLogical(args: string, context: FormulaContext): boolean[] {
  const values: boolean[] = [];
  
  // 引数をカンマで分割
  const parts = args.split(',').map(part => part.trim());
  
  for (const part of parts) {
    const cellValue = getCellValue(part, context) ?? part;
    
    // 数値の場合：0はfalse、それ以外はtrue
    if (typeof cellValue === 'number') {
      values.push(cellValue !== 0);
    }
    // 文字列の場合："TRUE"/"FALSE"または"1"/"0"をチェック
    else if (typeof cellValue === 'string') {
      const str = cellValue.toLowerCase();
      if (str === 'true' || str === '1') {
        values.push(true);
      } else if (str === 'false' || str === '0') {
        values.push(false);
      } else {
        // 文字列は一般的にtrueとして扱う（空文字列以外）
        values.push(cellValue !== '');
      }
    }
    // 論理値の場合
    else if (typeof cellValue === 'boolean') {
      values.push(cellValue);
    }
    // nullまたはundefinedの場合
    else if (cellValue === null || cellValue === undefined) {
      values.push(false);
    }
    else {
      values.push(Boolean(cellValue));
    }
  }
  
  return values;
}

// 単一の値を論理値に変換
function toLogical(value: unknown): boolean {
  if (typeof value === 'number') {
    return value !== 0;
  }
  if (typeof value === 'string') {
    const str = value.toLowerCase();
    if (str === 'true' || str === '1') return true;
    if (str === 'false' || str === '0') return false;
    return value !== '';
  }
  if (typeof value === 'boolean') {
    return value;
  }
  if (value === null || value === undefined) {
    return false;
  }
  return Boolean(value);
}

// IF関数の実装
export const IF: CustomFormula = {
  name: 'IF',
  pattern: /IF\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, conditionRef, trueValueRef, falseValueRef] = matches;
    
    const conditionValue = getCellValue(conditionRef, context) ?? conditionRef;
    const condition = toLogical(conditionValue);
    
    if (condition) {
      const trueValue = getCellValue(trueValueRef, context) ?? trueValueRef;
      // 文字列の引用符を除去
      if (typeof trueValue === 'string' && trueValue.startsWith('"') && trueValue.endsWith('"')) {
        return trueValue.slice(1, -1);
      }
      return trueValue as FormulaResult;
    } else {
      const falseValue = getCellValue(falseValueRef, context) ?? falseValueRef;
      // 文字列の引用符を除去
      if (typeof falseValue === 'string' && falseValue.startsWith('"') && falseValue.endsWith('"')) {
        return falseValue.slice(1, -1);
      }
      return falseValue as FormulaResult;
    }
  }
};

// AND関数の実装
export const AND: CustomFormula = {
  name: 'AND',
  pattern: /AND\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, args] = matches;
    const logicalValues = parseArgumentsToLogical(args, context);
    
    if (logicalValues.length === 0) {
      return FormulaError.VALUE;
    }
    
    // すべてがtrueの場合のみtrue
    return logicalValues.every(value => value === true);
  }
};

// OR関数の実装
export const OR: CustomFormula = {
  name: 'OR',
  pattern: /OR\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, args] = matches;
    const logicalValues = parseArgumentsToLogical(args, context);
    
    if (logicalValues.length === 0) {
      return FormulaError.VALUE;
    }
    
    // いずれかがtrueの場合にtrue
    return logicalValues.some(value => value === true);
  }
};

// NOT関数の実装
export const NOT: CustomFormula = {
  name: 'NOT',
  pattern: /NOT\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, valueRef] = matches;
    
    const value = getCellValue(valueRef, context) ?? valueRef;
    const logicalValue = toLogical(value);
    
    return !logicalValue;
  }
};

// IFS関数の実装（手動実装が必要）
export const IFS: CustomFormula = {
  name: 'IFS',
  pattern: /IFS\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const argsString = matches[1];
    
    // 引数をパースする（カンマで分割、ただし関数内のカンマは無視）
    const args: string[] = [];
    let currentArg = '';
    let parenDepth = 0;
    let inQuotes = false;
    
    for (let i = 0; i < argsString.length; i++) {
      const char = argsString[i];
      
      if (char === '"' && argsString[i - 1] !== '\\') {
        inQuotes = !inQuotes;
      }
      
      if (!inQuotes) {
        if (char === '(') parenDepth++;
        else if (char === ')') parenDepth--;
        else if (char === ',' && parenDepth === 0) {
          args.push(currentArg.trim());
          currentArg = '';
          continue;
        }
      }
      
      currentArg += char;
    }
    
    if (currentArg.trim()) {
      args.push(currentArg.trim());
    }
    
    // 条件と値のペアを評価
    for (let i = 0; i < args.length; i += 2) {
      if (i + 1 >= args.length) break;
      
      const conditionArg = args[i];
      const valueArg = args[i + 1];
      
      // 条件を評価（簡単な実装）
      let conditionResult = false;
      
      // セル参照の場合
      if (conditionArg.match(/^[A-Z]+\d+$/)) {
        const cellValue = getCellValue(conditionArg, context);
        conditionResult = Boolean(cellValue);
      } else {
        // 簡単な比較演算子の処理
        const comparisonMatch = conditionArg.match(/([A-Z]+\d+|"[^"]*"|\d+)\s*([><=!]+)\s*([A-Z]+\d+|"[^"]*"|\d+)/);
        if (comparisonMatch) {
          const [, leftParam, operator, rightParam] = comparisonMatch;
          let left = leftParam;
          let right = rightParam;
          
          // セル参照の場合は値を取得
          if (left.match(/^[A-Z]+\d+$/)) {
            left = String(getCellValue(left, context));
          } else if (left.startsWith('"') && left.endsWith('"')) {
            left = left.slice(1, -1);
          }
          
          if (right.match(/^[A-Z]+\d+$/)) {
            right = String(getCellValue(right, context));
          } else if (right.startsWith('"') && right.endsWith('"')) {
            right = right.slice(1, -1);
          }
          
          // 数値比較の場合
          const leftNum = parseFloat(left);
          const rightNum = parseFloat(right);
          
          if (!isNaN(leftNum) && !isNaN(rightNum)) {
            switch (operator) {
              case '>': conditionResult = leftNum > rightNum; break;
              case '<': conditionResult = leftNum < rightNum; break;
              case '>=': conditionResult = leftNum >= rightNum; break;
              case '<=': conditionResult = leftNum <= rightNum; break;
              case '=': conditionResult = leftNum === rightNum; break;
              case '<>':
              case '!=': conditionResult = leftNum !== rightNum; break;
            }
          } else {
            // 文字列比較
            switch (operator) {
              case '=': conditionResult = String(left) === String(right); break;
              case '<>':
              case '!=': conditionResult = String(left) !== String(right); break;
            }
          }
        }
      }
      
      if (conditionResult) {
        // 戻り値を処理
        if (valueArg.match(/^[A-Z]+\d+$/)) {
          return getCellValue(valueArg, context) as FormulaResult;
        } else if (valueArg.startsWith('"') && valueArg.endsWith('"')) {
          return valueArg.slice(1, -1);
        } else {
          const num = parseFloat(valueArg);
          return !isNaN(num) ? num : valueArg;
        }
      }
    }
    
    return FormulaError.NA;
  }
};

// XOR関数の実装（排他的論理和）
export const XOR: CustomFormula = {
  name: 'XOR',
  pattern: /XOR\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const args = matches[1].split(',').map(arg => arg.trim());
    let trueCount = 0;
    
    for (const arg of args) {
      let value = false;
      
      // セル参照の場合
      if (arg.match(/^[A-Z]+\d+$/)) {
        const cellValue = getCellValue(arg, context);
        value = Boolean(cellValue);
      } else {
        // 値の評価
        const trimmed = arg.replace(/"/g, '');
        if (trimmed === 'TRUE' || trimmed === '1') {
          value = true;
        } else if (trimmed === 'FALSE' || trimmed === '0') {
          value = false;
        } else {
          const numValue = parseFloat(trimmed);
          value = !isNaN(numValue) ? numValue !== 0 : trimmed.length > 0;
        }
      }
      
      if (value) {
        trueCount++;
      }
    }
    
    // XORは奇数個の引数がtrueの場合にtrueを返す
    return trueCount % 2 === 1;
  }
};

// TRUE関数の実装
export const TRUE: CustomFormula = {
  name: 'TRUE',
  pattern: /TRUE\(\)/i,
  calculate: (): FormulaResult => {
    return true;
  }
};

// FALSE関数の実装
export const FALSE: CustomFormula = {
  name: 'FALSE',
  pattern: /FALSE\(\)/i,
  calculate: (): FormulaResult => {
    return false;
  }
};

// IFERROR関数の実装（エラー時の値指定）
export const IFERROR: CustomFormula = {
  name: 'IFERROR',
  pattern: /IFERROR\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const valueArg = matches[1].trim();
    const errorValue = matches[2].replace(/"/g, '').trim();
    
    // セル参照の場合
    if (valueArg.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(valueArg, context);
      // エラーチェック（文字列でエラーを表現）
      if (typeof cellValue === 'string' && cellValue.startsWith('#') && cellValue.endsWith('!')) {
        return errorValue;
      }
      return cellValue as FormulaResult;
    }
    
    // 値がエラーかどうかをチェック
    if (valueArg.startsWith('#') && valueArg.endsWith('!')) {
      return errorValue;
    }
    
    // 数値変換を試す
    const numValue = parseFloat(valueArg);
    if (!isNaN(numValue)) {
      return numValue;
    }
    
    return valueArg.replace(/"/g, '');
  }
};

// IFNA関数の実装（#N/Aエラー時の値）
export const IFNA: CustomFormula = {
  name: 'IFNA',
  pattern: /IFNA\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const valueArg = matches[1].trim();
    const naValue = matches[2].replace(/"/g, '').trim();
    
    // セル参照の場合
    if (valueArg.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(valueArg, context);
      // #N/Aエラーチェック
      if (cellValue === '#N/A' || cellValue === '#N/A!') {
        return naValue;
      }
      return cellValue as FormulaResult;
    }
    
    // #N/Aエラーの場合のみ置換
    if (valueArg === '#N/A' || valueArg === '#N/A!') {
      return naValue;
    }
    
    // 数値変換を試す
    const numValue = parseFloat(valueArg);
    if (!isNaN(numValue)) {
      return numValue;
    }
    
    return valueArg.replace(/"/g, '');
  }
};