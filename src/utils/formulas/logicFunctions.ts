// 論理関数の実装

import type { CustomFormula, FormulaContext, FormulaResult } from './types';
import { FormulaError } from './types';
import { getCellValue } from './utils';

// IF関数の実装
export const IF: CustomFormula = {
  name: 'IF',
  pattern: /IF\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// AND関数の実装
export const AND: CustomFormula = {
  name: 'AND',
  pattern: /AND\(([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// OR関数の実装
export const OR: CustomFormula = {
  name: 'OR',
  pattern: /OR\(([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// NOT関数の実装
export const NOT: CustomFormula = {
  name: 'NOT',
  pattern: /NOT\(([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
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