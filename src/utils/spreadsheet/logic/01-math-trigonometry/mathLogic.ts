// 数学関数の実装

import type { CustomFormula, FormulaContext } from '../shared/types';
import { FormulaError } from '../shared/types';
import { getCellValue, getCellRangeValues } from '../shared/utils';

// 引数を数値の配列に変換するユーティリティ関数
function parseArgumentsToNumbers(args: string, context: FormulaContext): number[] {
  const values: number[] = [];
  
  // 引数をカンマで分割
  const parts = args.split(',').map(part => part.trim());
  
  for (const part of parts) {
    // 範囲指定（A1:B2のような形式）かチェック
    if (part.includes(':')) {
      const rangeValues = getCellRangeValues(part, context);
      for (const value of rangeValues) {
        const num = Number(value);
        if (!isNaN(num)) {
          values.push(num);
        }
      }
    } else {
      // 単一のセルまたは値
      const cellValue = getCellValue(part, context) ?? part;
      const num = Number(cellValue);
      if (!isNaN(num)) {
        values.push(num);
      }
    }
  }
  
  return values;
}

// 条件を評価するユーティリティ関数
function evaluateCondition(value: unknown, condition: string): boolean {
  // 条件を解析（例: ">5", ">=10", "<>0", "text"など）
  const stringValue = String(value);
  const numericValue = Number(value);
  
  // 数値比較の正規表現
  const numericConditionMatch = condition.match(/^(>=|<=|<>|>|<|=)?(.+)$/);
  if (numericConditionMatch) {
    const [, operator, operandStr] = numericConditionMatch;
    const operand = Number(operandStr);
    
    if (!isNaN(numericValue) && !isNaN(operand)) {
      switch (operator) {
        case '>': return numericValue > operand;
        case '<': return numericValue < operand;
        case '>=': return numericValue >= operand;
        case '<=': return numericValue <= operand;
        case '<>': return numericValue !== operand;
        case '=':
        case undefined: return numericValue === operand;
        default: return false;
      }
    }
  }
  
  // 文字列の完全一致
  if (condition.startsWith('"') && condition.endsWith('"')) {
    const conditionText = condition.slice(1, -1);
    return stringValue === conditionText;
  }
  
  // ワイルドカード対応（*と?）
  if (condition.includes('*') || condition.includes('?')) {
    const regex = new RegExp(
      '^' + condition
        .replace(/\*/g, '.*')
        .replace(/\?/g, '.')
        .replace(/[.+^${}()|[\]\\]/g, '\\$&') + '$',
      'i'
    );
    return regex.test(stringValue);
  }
  
  // デフォルトは文字列の完全一致
  return stringValue === condition;
}

// SUMIF関数の実装
export const SUMIF: CustomFormula = {
  name: 'SUMIF',
  pattern: /SUMIF\(([^,]+),\s*([^,]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, range, criteria, sumRange] = matches;
    
    try {
      // 範囲の値を取得
      const rangeValues = getCellRangeValues(range.trim(), context);
      
      // 合計範囲が指定されていない場合は検索範囲と同じ
      const sumRangeValues = sumRange 
        ? getCellRangeValues(sumRange.trim(), context)
        : rangeValues;
      
      if (rangeValues.length !== sumRangeValues.length) {
        return FormulaError.VALUE;
      }
      
      // 条件文字列を処理（引用符を除去）
      const cleanCriteria = criteria.replace(/^["']|["']$/g, '');
      
      let sum = 0;
      for (let i = 0; i < rangeValues.length; i++) {
        if (evaluateCondition(rangeValues[i], cleanCriteria)) {
          const value = Number(sumRangeValues[i]);
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

// COUNTIF関数の実装
export const COUNTIF: CustomFormula = {
  name: 'COUNTIF',
  pattern: /COUNTIF\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, range, criteria] = matches;
    
    try {
      // 範囲の値を取得
      const rangeValues = getCellRangeValues(range.trim(), context);
      
      // 条件文字列を処理（引用符を除去）
      const cleanCriteria = criteria.replace(/^["']|["']$/g, '');
      
      let count = 0;
      for (const value of rangeValues) {
        if (evaluateCondition(value, cleanCriteria)) {
          count++;
        }
      }
      
      return count;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// AVERAGEIF関数の実装
export const AVERAGEIF: CustomFormula = {
  name: 'AVERAGEIF',
  pattern: /AVERAGEIF\(([^,]+),\s*([^,]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, range, criteria, averageRange] = matches;
    
    try {
      // 範囲の値を取得
      const rangeValues = getCellRangeValues(range.trim(), context);
      
      // 平均範囲が指定されていない場合は検索範囲と同じ
      const avgRangeValues = averageRange 
        ? getCellRangeValues(averageRange.trim(), context)
        : rangeValues;
      
      if (rangeValues.length !== avgRangeValues.length) {
        return FormulaError.VALUE;
      }
      
      // 条件文字列を処理（引用符を除去）
      const cleanCriteria = criteria.replace(/^["']|["']$/g, '');
      
      let sum = 0;
      let count = 0;
      
      for (let i = 0; i < rangeValues.length; i++) {
        // Skip empty cells in the range
        if (rangeValues[i] === null || rangeValues[i] === undefined || rangeValues[i] === '') {
          continue;
        }
        
        if (evaluateCondition(rangeValues[i], cleanCriteria)) {
          const value = Number(avgRangeValues[i]);
          if (!isNaN(value)) {
            sum += value;
            count++;
          }
        }
      }
      
      return count > 0 ? sum / count : 0; // Return 0 instead of #DIV/0! for no matches
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// SUM関数の実装
export const SUM: CustomFormula = {
  name: 'SUM',
  pattern: /SUM\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, args] = matches;
    const numbers = parseArgumentsToNumbers(args, context);
    
    if (numbers.length === 0) {
      return 0;
    }
    
    return numbers.reduce((sum, num) => sum + num, 0);
  }
};

// AVERAGE関数の実装
export const AVERAGE: CustomFormula = {
  name: 'AVERAGE',
  pattern: /AVERAGE\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, args] = matches;
    const numbers = parseArgumentsToNumbers(args, context);
    
    if (numbers.length === 0) {
      return FormulaError.DIV0;
    }
    
    const sum = numbers.reduce((total, num) => total + num, 0);
    return sum / numbers.length;
  }
};

// COUNT関数の実装
export const COUNT: CustomFormula = {
  name: 'COUNT',
  pattern: /COUNT\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, args] = matches;
    const numbers = parseArgumentsToNumbers(args, context);
    
    return numbers.length;
  }
};

// MAX関数の実装
export const MAX: CustomFormula = {
  name: 'MAX',
  pattern: /MAX\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, args] = matches;
    const numbers = parseArgumentsToNumbers(args, context);
    
    if (numbers.length === 0) {
      return 0; // Excel returns 0 for MAX of empty range
    }
    
    return Math.max(...numbers);
  }
};

// MIN関数の実装
export const MIN: CustomFormula = {
  name: 'MIN',
  pattern: /MIN\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, args] = matches;
    const numbers = parseArgumentsToNumbers(args, context);
    
    if (numbers.length === 0) {
      return 0; // Excel returns 0 for MIN of empty range
    }
    
    return Math.min(...numbers);
  }
};

// ROUND関数の実装
export const ROUND: CustomFormula = {
  name: 'ROUND',
  pattern: /ROUND\((.+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, argsString] = matches;
    
    // 引数を正しく分割（括弧内のカンマは無視）
    const args: string[] = [];
    let currentArg = '';
    let parenDepth = 0;
    
    for (let i = 0; i < argsString.length; i++) {
      const char = argsString[i];
      
      if (char === '(') parenDepth++;
      else if (char === ')') parenDepth--;
      else if (char === ',' && parenDepth === 0) {
        args.push(currentArg.trim());
        currentArg = '';
        continue;
      }
      
      currentArg += char;
    }
    
    if (currentArg.trim()) {
      args.push(currentArg.trim());
    }
    
    
    const numberRef = args[0];
    const digitsRef = args[1];
    
    let number: number;
    let digits: number;
    
    // 数値の取得（演算式の場合も考慮）
    if (numberRef.includes('/') || numberRef.includes('*') || numberRef.includes('+') || numberRef.includes('-')) {
      // 簡単な演算を処理
      const parts = numberRef.split(/([+\-*/])/);
      
      if (parts.length === 3) {
        const [leftPart, operator, rightPart] = parts;
        let leftValue: number;
        let rightValue: number;
        
        // 左辺の値を取得
        if (leftPart.trim().match(/^[A-Z]+\d+$/)) {
          const cellValue = getCellValue(leftPart.trim(), context);
          // オブジェクトの場合、valueプロパティを取得
          if (cellValue && typeof cellValue === 'object' && 'value' in cellValue) {
            leftValue = Number(cellValue.value);
          } else {
            leftValue = Number(cellValue);
          }
        } else {
          leftValue = Number(leftPart.trim());
        }
        
        // 右辺の値を取得
        if (rightPart.trim().match(/^[A-Z]+\d+$/)) {
          const cellValue = getCellValue(rightPart.trim(), context);
          // オブジェクトの場合、valueプロパティを取得
          if (cellValue && typeof cellValue === 'object' && 'value' in cellValue) {
            rightValue = Number(cellValue.value);
          } else {
            rightValue = Number(cellValue);
          }
        } else {
          rightValue = Number(rightPart.trim());
        }
        
        // 演算を実行
        switch (operator) {
          case '+': number = leftValue + rightValue; break;
          case '-': number = leftValue - rightValue; break;
          case '*': number = leftValue * rightValue; break;
          case '/': number = leftValue / rightValue; break;
          default: number = NaN;
        }
      } else {
        number = Number(getCellValue(numberRef, context) ?? numberRef);
      }
    } else {
      const value = getCellValue(numberRef, context);
      if (value && typeof value === 'object' && 'value' in value) {
        number = Number(value.value);
      } else if (value !== null && value !== undefined && value !== FormulaError.REF) {
        number = Number(value);
      } else {
        number = Number(numberRef);
      }
    }
    
    // 桁数の取得
    const digitsValue = getCellValue(digitsRef, context);
    digits = digitsValue !== null && digitsValue !== undefined && digitsValue !== FormulaError.REF
      ? Number(digitsValue)
      : Number(digitsRef);
    
    if (isNaN(number) || isNaN(digits)) {
      return FormulaError.VALUE;
    }
    
    const factor = Math.pow(10, digits);
    return Math.round(number * factor) / factor;
  }
};

// ABS関数の実装（絶対値）
export const ABS: CustomFormula = {
  name: 'ABS',
  pattern: /ABS\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, numberRef] = matches;
    
    const number = Number(getCellValue(numberRef, context) ?? numberRef);
    
    if (isNaN(number)) {
      return FormulaError.VALUE;
    }
    
    return Math.abs(number);
  }
};

// SQRT関数の実装（平方根）
export const SQRT: CustomFormula = {
  name: 'SQRT',
  pattern: /SQRT\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, numberRef] = matches;
    
    const number = Number(getCellValue(numberRef, context) ?? numberRef);
    
    if (isNaN(number) || number < 0) {
      return FormulaError.NUM;
    }
    
    return Math .sqrt(number);
  }
};

// POWER関数の実装（べき乗）
export const POWER: CustomFormula = {
  name: 'POWER',
  pattern: /POWER\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, baseRef, exponentRef] = matches;
    
    let base: number;
    let exponent: number;
    
    // ベース値の取得
    const baseValue = getCellValue(baseRef.trim(), context);
    if (baseValue !== null && baseValue !== undefined && baseValue !== FormulaError.REF) {
      base = Number(baseValue);
    } else {
      base = Number(baseRef.trim());
    }
    
    // 指数値の取得
    const expValue = getCellValue(exponentRef.trim(), context);
    if (expValue !== null && expValue !== undefined && expValue !== FormulaError.REF) {
      exponent = Number(expValue);
    } else {
      exponent = Number(exponentRef.trim());
    }
    
    if (isNaN(base) || isNaN(exponent)) {
      return FormulaError.VALUE;
    }
    
    const result = Math.pow(base, exponent);
    
    if (!isFinite(result)) {
      return FormulaError.NUM;
    }
    
    return result;
  }
};

// MOD関数の実装（剰余）
export const MOD: CustomFormula = {
  name: 'MOD',
  pattern: /MOD\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, numberRef, divisorRef] = matches;
    
    const number = Number(getCellValue(numberRef, context) ?? numberRef);
    const divisor = Number(getCellValue(divisorRef, context) ?? divisorRef);
    
    if (isNaN(number) || isNaN(divisor) || divisor === 0) {
      return FormulaError.DIV0;
    }
    
    return number % divisor;
  }
};

// INT関数の実装（整数部分）
export const INT: CustomFormula = {
  name: 'INT',
  pattern: /INT\(([^)]+)\)/i,
  calculate: (matches, context) => {
    const valueRef = matches[1].trim();
    let value: number;
    
    if (valueRef.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(valueRef, context);
      value = parseFloat(String(cellValue ?? '0'));
    } else {
      value = parseFloat(valueRef);
    }
    
    if (isNaN(value)) return FormulaError.VALUE;
    return Math.floor(value);
  }
};

// TRUNC関数の実装（小数部分切り捨て）
export const TRUNC: CustomFormula = {
  name: 'TRUNC',
  pattern: /TRUNC\(([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, numberRef, digitsRef] = matches;
    
    const number = Number(getCellValue(numberRef, context) ?? numberRef);
    const digits = digitsRef ? Number(getCellValue(digitsRef, context) ?? digitsRef) : 0;
    
    if (isNaN(number) || isNaN(digits)) {
      return FormulaError.VALUE;
    }
    
    const factor = Math.pow(10, digits);
    return Math.trunc(number * factor) / factor;
  }
};

// RAND関数の実装（0以上1未満の乱数）
export const RAND: CustomFormula = {
  name: 'RAND',
  pattern: /RAND\(\)/i,
  calculate: () => {
    return Math.random();
  }
};

// RANDBETWEEN関数の実装（指定範囲の整数乱数）
export const RANDBETWEEN: CustomFormula = {
  name: 'RANDBETWEEN',
  pattern: /RANDBETWEEN\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches, context) => {
    const minRef = matches[1].trim();
    const maxRef = matches[2].trim();
    
    let min: number, max: number;
    
    if (minRef.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(minRef, context);
      min = parseInt(String(cellValue ?? '0'));
    } else {
      min = parseInt(minRef);
    }
    
    if (maxRef.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(maxRef, context);
      max = parseInt(String(cellValue ?? '0'));
    } else {
      max = parseInt(maxRef);
    }
    
    if (isNaN(min) || isNaN(max)) return FormulaError.VALUE;
    if (min > max) return FormulaError.NUM;
    return Math.floor(Math.random() * (max - min + 1)) + min;
  }
};

// PI関数の実装（円周率）
export const PI: CustomFormula = {
  name: 'PI',
  pattern: /PI\(\)/i,
  calculate: () => {
    return Math.PI;
  }
};

// DEGREES関数の実装（ラジアンを度に変換）
export const DEGREES: CustomFormula = {
  name: 'DEGREES',
  pattern: /DEGREES\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, angleRef] = matches;
    
    const angleValue = getCellValue(angleRef, context) ?? angleRef;
    const angle = Number(angleValue);
    
    if (isNaN(angle)) {
      return FormulaError.VALUE;
    }
    
    // ラジアンを度に変換
    return angle * (180 / Math.PI);
  }
};

// RADIANS関数の実装（度をラジアンに変換）
export const RADIANS: CustomFormula = {
  name: 'RADIANS',
  pattern: /RADIANS\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, angleRef] = matches;
    
    const angleValue = getCellValue(angleRef, context) ?? angleRef;
    const angle = Number(angleValue);
    
    if (isNaN(angle)) {
      return FormulaError.VALUE;
    }
    
    // 度をラジアンに変換
    return angle * (Math.PI / 180);
  }
};

// SIN関数の実装（正弦）

// LOG関数の実装（対数）
export const LOG: CustomFormula = {
  name: 'LOG',
  pattern: /LOG\(([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, numberRef, baseRef] = matches;
    
    const number = Number(getCellValue(numberRef, context) ?? numberRef);
    const base = baseRef ? Number(getCellValue(baseRef, context) ?? baseRef) : 10;
    
    if (isNaN(number) || isNaN(base) || number <= 0 || base <= 0 || base === 1) {
      return FormulaError.NUM;
    }
    
    return Math.log(number) / Math.log(base);
  }
};

// LOG10関数の実装（常用対数）
export const LOG10: CustomFormula = {
  name: 'LOG10',
  pattern: /LOG10\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, numberRef] = matches;
    
    const number = Number(getCellValue(numberRef, context) ?? numberRef);
    
    if (isNaN(number) || number <= 0) {
      return FormulaError.NUM;
    }
    
    return Math.log10(number);
  }
};

// LN関数の実装（自然対数）
export const LN: CustomFormula = {
  name: 'LN',
  pattern: /LN\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, numberRef] = matches;
    
    const number = Number(getCellValue(numberRef, context) ?? numberRef);
    
    if (isNaN(number) || number <= 0) {
      return FormulaError.NUM;
    }
    
    return Math.log(number);
  }
};

// EXP関数の実装（指数関数）
export const EXP: CustomFormula = {
  name: 'EXP',
  pattern: /EXP\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, numberRef] = matches;
    
    const number = Number(getCellValue(numberRef, context) ?? numberRef);
    
    if (isNaN(number)) {
      return FormulaError.VALUE;
    }
    
    const result = Math.exp(number);
    
    if (!isFinite(result)) {
      return FormulaError.NUM;
    }
    
    return result;
  }
};

// ASIN関数の実装（逆正弦）

// ROUNDUP関数の実装（切り上げ）
export const ROUNDUP: CustomFormula = {
  name: 'ROUNDUP',
  pattern: /ROUNDUP\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, numberRef, digitsRef] = matches;
    
    const number = Number(getCellValue(numberRef, context) ?? numberRef);
    const digits = Number(getCellValue(digitsRef, context) ?? digitsRef);
    
    if (isNaN(number) || isNaN(digits)) {
      return FormulaError.VALUE;
    }
    
    const factor = Math.pow(10, digits);
    if (number >= 0) {
      return Math.ceil(number * factor) / factor;
    } else {
      return Math.floor(number * factor) / factor;
    }
  }
};

// ROUNDDOWN関数の実装（切り下げ）
export const ROUNDDOWN: CustomFormula = {
  name: 'ROUNDDOWN',
  pattern: /ROUNDDOWN\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, numberRef, digitsRef] = matches;
    
    const number = Number(getCellValue(numberRef, context) ?? numberRef);
    const digits = Number(getCellValue(digitsRef, context) ?? digitsRef);
    
    if (isNaN(number) || isNaN(digits)) {
      return FormulaError.VALUE;
    }
    
    const factor = Math.pow(10, digits);
    if (number >= 0) {
      return Math.floor(number * factor) / factor;
    } else {
      return Math.ceil(number * factor) / factor;
    }
  }
};

// CEILING関数の実装（基準値の倍数に切り上げ）
export const CEILING: CustomFormula = {
  name: 'CEILING',
  pattern: /CEILING\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, numberRef, significanceRef] = matches;
    
    const number = Number(getCellValue(numberRef, context) ?? numberRef);
    const significance = Number(getCellValue(significanceRef, context) ?? significanceRef);
    
    if (isNaN(number) || isNaN(significance)) {
      return FormulaError.VALUE;
    }
    
    if (significance === 0) {
      return 0;
    }
    
    if (number === 0) {
      return 0;
    }
    
    if (number > 0 && significance > 0) {
      return Math.ceil(number / significance) * significance;
    } else if (number < 0 && significance < 0) {
      return Math.floor(number / significance) * significance;
    } else {
      return FormulaError.NUM;
    }
  }
};

// FLOOR関数の実装（基準値の倍数に切り下げ）
export const FLOOR: CustomFormula = {
  name: 'FLOOR',
  pattern: /FLOOR\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, numberRef, significanceRef] = matches;
    
    const number = Number(getCellValue(numberRef, context) ?? numberRef);
    const significance = Number(getCellValue(significanceRef, context) ?? significanceRef);
    
    if (isNaN(number) || isNaN(significance)) {
      return FormulaError.VALUE;
    }
    
    if (significance === 0) {
      return 0;
    }
    
    if (number === 0) {
      return 0;
    }
    
    if (number > 0 && significance > 0) {
      return Math.floor(number / significance) * significance;
    } else if (number < 0 && significance < 0) {
      return Math.ceil(number / significance) * significance;
    } else {
      return FormulaError.NUM;
    }
  }
};

// SIGN関数の実装（符号）
export const SIGN: CustomFormula = {
  name: 'SIGN',
  pattern: /SIGN\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, numberRef] = matches;
    
    const number = Number(getCellValue(numberRef, context) ?? numberRef);
    
    if (isNaN(number)) {
      return FormulaError.VALUE;
    }
    
    return Math.sign(number);
  }
};

// FACT関数の実装（階乗）
export const FACT: CustomFormula = {
  name: 'FACT',
  pattern: /FACT\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, numberRef] = matches;
    
    const number = Number(getCellValue(numberRef, context) ?? numberRef);
    
    if (isNaN(number) || number < 0 || !Number.isInteger(number)) {
      return FormulaError.NUM;
    }
    
    if (number > 170) {
      return FormulaError.NUM; // オーバーフロー防止
    }
    
    let result = 1;
    for (let i = 2; i <= number; i++) {
      result *= i;
    }
    
    return result;
  }
};

// SUMIFS関数の実装（複数条件での合計）
export const SUMIFS: CustomFormula = {
  name: 'SUMIFS',
  pattern: /SUMIFS\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, argsStr] = matches;
    
    try {
      // 引数を解析
      const args: string[] = [];
      let current = '';
      let inQuotes = false;
      let depth = 0;
      
      for (let i = 0; i < argsStr.length; i++) {
        const char = argsStr[i];
        
        if (char === '"') {
          inQuotes = !inQuotes;
          current += char;
        } else if (!inQuotes && char === '(') {
          depth++;
          current += char;
        } else if (!inQuotes && char === ')') {
          depth--;
          current += char;
        } else if (!inQuotes && depth === 0 && char === ',') {
          args.push(current.trim());
          current = '';
        } else {
          current += char;
        }
      }
      if (current) {
        args.push(current.trim());
      }
      
      if (args.length < 3 || args.length % 2 === 0) {
        return FormulaError.VALUE;
      }
      
      // 最初の引数は合計範囲
      const sumRangeValues = getCellRangeValues(args[0], context);
      
      // 条件範囲と条件のペアを処理
      const conditionPairs: Array<{values: unknown[], criteria: string}> = [];
      for (let i = 1; i < args.length; i += 2) {
        const rangeValues = getCellRangeValues(args[i], context);
        const criteria = args[i + 1].replace(/^["']|["']$/g, '');
        
        if (rangeValues.length !== sumRangeValues.length) {
          return FormulaError.VALUE;
        }
        
        conditionPairs.push({ values: rangeValues, criteria });
      }
      
      // すべての条件を満たす行の合計を計算
      let sum = 0;
      for (let i = 0; i < sumRangeValues.length; i++) {
        let allConditionsMet = true;
        
        for (const pair of conditionPairs) {
          if (!evaluateCondition(pair.values[i], pair.criteria)) {
            allConditionsMet = false;
            break;
          }
        }
        
        if (allConditionsMet) {
          const value = Number(sumRangeValues[i]);
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

// COUNTIFS関数の実装（複数条件でのカウント）
export const COUNTIFS: CustomFormula = {
  name: 'COUNTIFS',
  pattern: /COUNTIFS\(([^,]+),\s*"([^"]+)"(?:,\s*([^,]+),\s*"([^"]+)")*\)/i,
  calculate: () => null // HyperFormulaが処理
};

// AVERAGEIFS関数の実装（複数条件での平均）
export const AVERAGEIFS: CustomFormula = {
  name: 'AVERAGEIFS',
  pattern: /AVERAGEIFS\(([^,]+),\s*([^,]+),\s*"([^"]+)"(?:,\s*([^,]+),\s*"([^"]+)")*\)/i,
  calculate: () => null // HyperFormulaが処理
};

// PRODUCT関数の実装（積を計算）
export const PRODUCT: CustomFormula = {
  name: 'PRODUCT',
  pattern: /PRODUCT\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, args] = matches;
    const numbers = parseArgumentsToNumbers(args, context);
    
    if (numbers.length === 0) {
      return 0;
    }
    
    return numbers.reduce((product, num) => product * num, 1);
  }
};

// MROUND関数の実装（倍数に丸める）
export const MROUND: CustomFormula = {
  name: 'MROUND',
  pattern: /MROUND\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [fullMatch, numberRef, multipleRef] = matches;
    
    const numberValue = getCellValue(numberRef.trim(), context);
    const multipleValue = getCellValue(multipleRef.trim(), context);
    
    const number = Number(numberValue ?? numberRef);
    const multiple = Number(multipleValue ?? multipleRef);
    
    if (isNaN(number) || isNaN(multiple)) {
      return FormulaError.VALUE;
    }
    
    if (multiple === 0) {
      return 0;
    }
    
    // 符号が異なる場合はエラー
    if ((number > 0 && multiple < 0) || (number < 0 && multiple > 0)) {
      return FormulaError.NUM;
    }
    
    const result = Math.round(number / multiple) * multiple;
    return result;
  }
};

// COMBIN関数の実装（組み合わせ数）
export const COMBIN: CustomFormula = {
  name: 'COMBIN',
  pattern: /COMBIN\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, nRef, kRef] = matches;
    
    const n = Number(getCellValue(nRef, context) ?? nRef);
    const k = Number(getCellValue(kRef, context) ?? kRef);
    
    if (isNaN(n) || isNaN(k)) {
      return FormulaError.VALUE;
    }
    
    if (!Number.isInteger(n) || !Number.isInteger(k)) {
      return FormulaError.NUM;
    }
    
    if (n < 0 || k < 0 || k > n) {
      return FormulaError.NUM;
    }
    
    if (k === 0 || k === n) {
      return 1;
    }
    
    // 計算効率を上げるため、k > n/2の場合はn-kを使う
    let k2 = k;
    if (k2 > n / 2) {
      k2 = n - k2;
    }
    
    let result = 1;
    for (let i = 0; i < k2; i++) {
      result = result * (n - i) / (i + 1);
    }
    
    return Math.round(result);
  }
};

// GCD関数の実装（最大公約数）
export const GCD: CustomFormula = {
  name: 'GCD',
  pattern: /GCD\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, args] = matches;
    const numbers = parseArgumentsToNumbers(args, context);
    
    if (numbers.length === 0) {
      return FormulaError.VALUE;
    }
    
    // すべて整数である必要がある
    const integers = numbers.map(n => {
      if (!Number.isInteger(n)) {
        return Math.floor(Math.abs(n));
      }
      return Math.abs(n);
    });
    
    // ユークリッドの互除法
    const gcd2 = (a: number, b: number): number => {
      return b === 0 ? a : gcd2(b, a % b);
    };
    
    let result = integers[0];
    for (let i = 1; i < integers.length; i++) {
      result = gcd2(result, integers[i]);
      if (result === 1) break; // 1以下にはならない
    }
    
    return result;
  }
};

// LCM関数の実装（最小公倍数）
export const LCM: CustomFormula = {
  name: 'LCM',
  pattern: /LCM\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, args] = matches;
    const numbers = parseArgumentsToNumbers(args, context);
    
    if (numbers.length === 0) {
      return FormulaError.VALUE;
    }
    
    // すべて整数である必要がある
    const integers = numbers.map(n => {
      if (!Number.isInteger(n)) {
        return Math.floor(Math.abs(n));
      }
      return Math.abs(n);
    });
    
    // GCDを使ってLCMを計算
    const gcd2 = (a: number, b: number): number => {
      return b === 0 ? a : gcd2(b, a % b);
    };
    
    const lcm2 = (a: number, b: number): number => {
      return (a * b) / gcd2(a, b);
    };
    
    let result = integers[0];
    for (let i = 1; i < integers.length; i++) {
      result = lcm2(result, integers[i]);
      if (!isFinite(result)) {
        return FormulaError.NUM;
      }
    }
    
    return result;
  }
};

// QUOTIENT関数の実装（商の整数部分）
export const QUOTIENT: CustomFormula = {
  name: 'QUOTIENT',
  pattern: /QUOTIENT\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, numeratorRef, denominatorRef] = matches;
    
    const numerator = Number(getCellValue(numeratorRef, context) ?? numeratorRef);
    const denominator = Number(getCellValue(denominatorRef, context) ?? denominatorRef);
    
    if (isNaN(numerator) || isNaN(denominator)) {
      return FormulaError.VALUE;
    }
    
    if (denominator === 0) {
      return FormulaError.DIV0;
    }
    
    return Math.trunc(numerator / denominator);
  }
};

// 双曲線関数の実装
// SINH関数（双曲線正弦）

// HyperFormulaでサポートされている未実装関数を追加

// SUMSQ関数（平方和）
export const SUMSQ: CustomFormula = {
  name: 'SUMSQ',
  pattern: /SUMSQ\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, args] = matches;
    const numbers = parseArgumentsToNumbers(args, context);
    
    if (numbers.length === 0) {
      return 0;
    }
    
    return numbers.reduce((sum, num) => sum + num * num, 0);
  }
};

// SUMPRODUCT関数（配列の積の和）
export const SUMPRODUCT: CustomFormula = {
  name: 'SUMPRODUCT',
  pattern: /SUMPRODUCT\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [fullMatch, argsStr] = matches;
    
    try {
      // 引数を解析（範囲を分割）
      const ranges = argsStr.split(',').map(r => r.trim());
      
      if (ranges.length === 0) {
        return FormulaError.VALUE;
      }
      
      // 各範囲の値を取得
      const rangeValues = ranges.map(range => getCellRangeValues(range, context));
      
      
      // 範囲が1つの場合は、その値の合計を返す
      if (rangeValues.length === 1) {
        return rangeValues[0].reduce((sum, val) => {
          const num = Number(val);
          return sum + (isNaN(num) ? 0 : num);
        }, 0);
      }
      
      // すべての範囲が同じ長さか確認
      const length = rangeValues[0].length;
      if (!rangeValues.every(values => values.length === length)) {
        return FormulaError.VALUE;
      }
      
      // 各位置で積を計算し、合計する
      let sum = 0;
      for (let i = 0; i < length; i++) {
        let product = 1;
        for (const values of rangeValues) {
          const num = Number(values[i]);
          if (!isNaN(num)) {
            product *= num;
          } else {
            product = 0;
            break;
          }
        }
        sum += product;
      }
      
      return sum;
    } catch (error) {
      return FormulaError.VALUE;
    }
  }
};

// EVEN関数（最も近い偶数に切り上げ）
export const EVEN: CustomFormula = {
  name: 'EVEN',
  pattern: /EVEN\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, numberRef] = matches;
    
    const number = Number(getCellValue(numberRef, context) ?? numberRef);
    
    if (isNaN(number)) {
      return FormulaError.VALUE;
    }
    
    if (number === 0) {
      return 0;
    }
    
    if (number > 0) {
      return Math.ceil(number / 2) * 2;
    } else {
      return Math.floor(number / 2) * 2;
    }
  }
};

// ODD関数（最も近い奇数に切り上げ）
export const ODD: CustomFormula = {
  name: 'ODD',
  pattern: /ODD\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, numberRef] = matches;
    
    const number = Number(getCellValue(numberRef, context) ?? numberRef);
    
    if (isNaN(number)) {
      return FormulaError.VALUE;
    }
    
    if (number === 0) {
      return 1;
    }
    
    const sign = Math.sign(number);
    const absNumber = Math.abs(number);
    const result = Math.ceil(absNumber);
    
    // 偶数の場合は1を足す
    if (result % 2 === 0) {
      return sign * (result + 1);
    } else {
      return sign * result;
    }
  }
};

// ARABIC関数（ローマ数字をアラビア数字に変換）
export const ARABIC: CustomFormula = {
  name: 'ARABIC',
  pattern: /ARABIC\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray) => {
    const romanStr = matches[1].trim().replace(/["']/g, '').toUpperCase();
    
    // ローマ数字の基本文字と対応する値
    const romanMap: { [key: string]: number } = {
      'I': 1, 'V': 5, 'X': 10, 'L': 50,
      'C': 100, 'D': 500, 'M': 1000
    };
    
    // 無効なローマ数字文字をチェック
    if (!/^[IVXLCDM]+$/.test(romanStr)) {
      return FormulaError.VALUE;
    }
    
    let result = 0;
    let prev = 0;
    
    // 右から左へ処理
    for (let i = romanStr.length - 1; i >= 0; i--) {
      const current = romanMap[romanStr[i]];
      
      if (current === undefined) {
        return FormulaError.VALUE;
      }
      
      // 前の文字より小さい場合は減算、そうでなければ加算
      if (current < prev) {
        result -= current;
      } else {
        result += current;
      }
      
      prev = current;
    }
    
    return result;
  }
};

// ROMAN関数（アラビア数字をローマ数字に変換）
export const ROMAN: CustomFormula = {
  name: 'ROMAN',
  pattern: /ROMAN\(([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, numberRef] = matches;
    
    const number = Math.floor(Number(getCellValue(numberRef, context) ?? numberRef));
    
    if (isNaN(number) || number < 0 || number > 3999) {
      return FormulaError.VALUE;
    }
    
    if (number === 0) {
      return '';
    }
    
    const romanNumerals = [
      { value: 1000, numeral: 'M' },
      { value: 900, numeral: 'CM' },
      { value: 500, numeral: 'D' },
      { value: 400, numeral: 'CD' },
      { value: 100, numeral: 'C' },
      { value: 90, numeral: 'XC' },
      { value: 50, numeral: 'L' },
      { value: 40, numeral: 'XL' },
      { value: 10, numeral: 'X' },
      { value: 9, numeral: 'IX' },
      { value: 5, numeral: 'V' },
      { value: 4, numeral: 'IV' },
      { value: 1, numeral: 'I' }
    ];
    
    let result = '';
    let remaining = number;
    
    for (const { value, numeral } of romanNumerals) {
      while (remaining >= value) {
        result += numeral;
        remaining -= value;
      }
    }
    
    return result;
  }
};

// COMBINA関数（重複組合せ）
export const COMBINA: CustomFormula = {
  name: 'COMBINA',
  pattern: /COMBINA\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, nRef, kRef] = matches;
    
    const n = Number(getCellValue(nRef, context) ?? nRef);
    const k = Number(getCellValue(kRef, context) ?? kRef);
    
    if (isNaN(n) || isNaN(k)) {
      return FormulaError.VALUE;
    }
    
    if (!Number.isInteger(n) || !Number.isInteger(k)) {
      return FormulaError.NUM;
    }
    
    if (n < 0 || k < 0) {
      return FormulaError.NUM;
    }
    
    if (n === 0 && k > 0) {
      return 0;
    }
    
    // 重複組合せの公式: C(n+k-1, k)
    const newN = n + k - 1;
    
    if (k === 0 || k === newN) {
      return 1;
    }
    
    // 計算効率を上げるため、k > newN/2の場合はnewN-kを使う
    let k2 = k;
    if (k2 > newN / 2) {
      k2 = newN - k2;
    }
    
    let result = 1;
    for (let i = 0; i < k2; i++) {
      result = result * (newN - i) / (i + 1);
    }
    
    return Math.round(result);
  }
};

// PERMUT関数（順列）
export const PERMUT: CustomFormula = {
  name: 'PERMUT',
  pattern: /PERMUT\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, nRef, kRef] = matches;
    
    const n = Number(getCellValue(nRef, context) ?? nRef);
    const k = Number(getCellValue(kRef, context) ?? kRef);
    
    if (isNaN(n) || isNaN(k)) {
      return FormulaError.VALUE;
    }
    
    if (!Number.isInteger(n) || !Number.isInteger(k)) {
      return FormulaError.NUM;
    }
    
    if (n < 0 || k < 0 || k > n) {
      return FormulaError.NUM;
    }
    
    if (k === 0) {
      return 1;
    }
    
    let result = 1;
    for (let i = 0; i < k; i++) {
      result *= (n - i);
    }
    
    return result;
  }
};

// PERMUTATIONA関数（重複順列）
export const PERMUTATIONA: CustomFormula = {
  name: 'PERMUTATIONA',
  pattern: /PERMUTATIONA\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray) => {
    const nStr = matches[1].trim();
    const kStr = matches[2].trim();
    
    // 数値の取得
    const n = parseFloat(nStr);
    const k = parseFloat(kStr);
    
    // 数値型チェック
    if (isNaN(n) || isNaN(k)) {
      return FormulaError.VALUE;
    }
    
    // 整数チェック
    if (!Number.isInteger(n) || !Number.isInteger(k)) {
      return FormulaError.NUM;
    }
    
    // 負数チェック
    if (n < 0 || k < 0) {
      return FormulaError.NUM;
    }
    
    // k = 0の場合
    if (k === 0) {
      return 1;
    }
    
    // n = 0の場合
    if (n === 0) {
      return 0;
    }
    
    // 重複順列の公式: n^k
    return Math.pow(n, k);
  }
};

// FACTDOUBLE関数（二重階乗）
export const FACTDOUBLE: CustomFormula = {
  name: 'FACTDOUBLE',
  pattern: /FACTDOUBLE\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, numberRef] = matches;
    
    const number = Number(getCellValue(numberRef, context) ?? numberRef);
    
    if (isNaN(number)) {
      return FormulaError.VALUE;
    }
    
    if (!Number.isInteger(number) || number < -1) {
      return FormulaError.NUM;
    }
    
    if (number > 300) {
      return FormulaError.NUM; // オーバーフロー防止
    }
    
    if (number === -1 || number === 0) {
      return 1;
    }
    
    let result = 1;
    for (let i = number; i > 0; i -= 2) {
      result *= i;
    }
    
    return result;
  }
};

// SQRTPI関数（π倍の平方根）
export const SQRTPI: CustomFormula = {
  name: 'SQRTPI',
  pattern: /SQRTPI\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, numberRef] = matches;
    
    const number = Number(getCellValue(numberRef, context) ?? numberRef);
    
    if (isNaN(number)) {
      return FormulaError.VALUE;
    }
    
    if (number < 0) {
      return FormulaError.NUM;
    }
    
    return Math.sqrt(number * Math.PI);
  }
};

// SUMX2MY2関数（x^2-y^2の和）
export const SUMX2MY2: CustomFormula = {
  name: 'SUMX2MY2',
  pattern: /SUMX2MY2\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, xRange, yRange] = matches;
    
    try {
      const xValues = getCellRangeValues(xRange.trim(), context);
      const yValues = getCellRangeValues(yRange.trim(), context);
      
      if (xValues.length !== yValues.length) {
        return FormulaError.NA;
      }
      
      let sum = 0;
      for (let i = 0; i < xValues.length; i++) {
        const x = Number(xValues[i]);
        const y = Number(yValues[i]);
        
        if (!isNaN(x) && !isNaN(y)) {
          sum += (x * x - y * y);
        }
      }
      
      return sum;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// SUMX2PY2関数（x^2+y^2の和）
export const SUMX2PY2: CustomFormula = {
  name: 'SUMX2PY2',
  pattern: /SUMX2PY2\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, xRange, yRange] = matches;
    
    try {
      const xValues = getCellRangeValues(xRange.trim(), context);
      const yValues = getCellRangeValues(yRange.trim(), context);
      
      if (xValues.length !== yValues.length) {
        return FormulaError.NA;
      }
      
      let sum = 0;
      for (let i = 0; i < xValues.length; i++) {
        const x = Number(xValues[i]);
        const y = Number(yValues[i]);
        
        if (!isNaN(x) && !isNaN(y)) {
          sum += (x * x + y * y);
        }
      }
      
      return sum;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// SUMXMY2関数（(x-y)^2の和）
export const SUMXMY2: CustomFormula = {
  name: 'SUMXMY2',
  pattern: /SUMXMY2\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, xRange, yRange] = matches;
    
    try {
      const xValues = getCellRangeValues(xRange.trim(), context);
      const yValues = getCellRangeValues(yRange.trim(), context);
      
      if (xValues.length !== yValues.length) {
        return FormulaError.NA;
      }
      
      let sum = 0;
      for (let i = 0; i < xValues.length; i++) {
        const x = Number(xValues[i]);
        const y = Number(yValues[i]);
        
        if (!isNaN(x) && !isNaN(y)) {
          const diff = x - y;
          sum += diff * diff;
        }
      }
      
      return sum;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// MULTINOMIAL関数（多項係数）
export const MULTINOMIAL: CustomFormula = {
  name: 'MULTINOMIAL',
  pattern: /MULTINOMIAL\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, args] = matches;
    const numbers = parseArgumentsToNumbers(args, context);
    
    if (numbers.length === 0) {
      return FormulaError.VALUE;
    }
    
    // すべて非負整数である必要がある
    for (const num of numbers) {
      if (!Number.isInteger(num) || num < 0) {
        return FormulaError.NUM;
      }
    }
    
    // 多項係数の計算: (n1 + n2 + ... + nk)! / (n1! * n2! * ... * nk!)
    const sum = numbers.reduce((a, b) => a + b, 0);
    
    if (sum > 170) {
      return FormulaError.NUM; // オーバーフロー防止
    }
    
    // 階乗を計算
    let numerator = 1;
    for (let i = 2; i <= sum; i++) {
      numerator *= i;
    }
    
    let denominator = 1;
    for (const num of numbers) {
      for (let i = 2; i <= num; i++) {
        denominator *= i;
      }
    }
    
    return numerator / denominator;
  }
};

// BASE関数（数値を指定した基数に変換）
export const BASE: CustomFormula = {
  name: 'BASE',
  pattern: /BASE\(([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, numberRef, radixRef, minLengthRef] = matches;
    
    const number = Math.floor(Number(getCellValue(numberRef, context) ?? numberRef));
    const radix = Number(getCellValue(radixRef, context) ?? radixRef);
    const minLength = minLengthRef ? Number(getCellValue(minLengthRef, context) ?? minLengthRef) : 0;
    
    if (isNaN(number) || isNaN(radix) || isNaN(minLength)) {
      return FormulaError.VALUE;
    }
    
    if (number < 0 || number > 2 ** 53 - 1) {
      return FormulaError.NUM;
    }
    
    if (radix < 2 || radix > 36 || !Number.isInteger(radix)) {
      return FormulaError.NUM;
    }
    
    if (minLength < 0 || minLength > 255) {
      return FormulaError.NUM;
    }
    
    let result = number.toString(radix).toUpperCase();
    
    // 最小長に満たない場合は0でパディング
    if (result.length < minLength) {
      result = '0'.repeat(minLength - result.length) + result;
    }
    
    return result;
  }
};

// DECIMAL関数（指定した基数の数値を10進数に変換）
export const DECIMAL: CustomFormula = {
  name: 'DECIMAL',
  pattern: /DECIMAL\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, textRef, radixRef] = matches;
    
    const text = String(getCellValue(textRef, context) ?? textRef).trim();
    const radix = Number(getCellValue(radixRef, context) ?? radixRef);
    
    if (isNaN(radix)) {
      return FormulaError.VALUE;
    }
    
    if (radix < 2 || radix > 36 || !Number.isInteger(radix)) {
      return FormulaError.NUM;
    }
    
    if (text.length === 0 || text.length > 255) {
      return FormulaError.VALUE;
    }
    
    // 基数に応じて有効な文字をチェック
    const validChars = '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ'.slice(0, radix);
    const upperText = text.toUpperCase();
    
    for (const char of upperText) {
      if (!validChars.includes(char)) {
        return FormulaError.NUM;
      }
    }
    
    const result = parseInt(upperText, radix);
    
    if (isNaN(result) || !isFinite(result)) {
      return FormulaError.NUM;
    }
    
    return result;
  }
};

// CEILING.MATH関数（数学的な切り上げ）
export const CEILING_MATH: CustomFormula = {
  name: 'CEILING.MATH',
  pattern: /CEILING\.MATH\(([^,)]+)(?:,\s*([^,)]+))?(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, numberRef, significanceRef, modeRef] = matches;
    
    const number = Number(getCellValue(numberRef, context) ?? numberRef);
    const significance = significanceRef ? Number(getCellValue(significanceRef, context) ?? significanceRef) : 1;
    const mode = modeRef ? Number(getCellValue(modeRef, context) ?? modeRef) : 0;
    
    if (isNaN(number) || isNaN(significance) || isNaN(mode)) {
      return FormulaError.VALUE;
    }
    
    if (significance === 0) {
      return 0;
    }
    
    const absSignificance = Math.abs(significance);
    
    if (number >= 0 || mode === 0) {
      return Math.ceil(number / absSignificance) * absSignificance;
    } else {
      return -Math.floor(Math.abs(number) / absSignificance) * absSignificance;
    }
  }
};

// CEILING.PRECISE関数（精密な切り上げ）
export const CEILING_PRECISE: CustomFormula = {
  name: 'CEILING.PRECISE',
  pattern: /CEILING\.PRECISE\(([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, numberRef, significanceRef] = matches;
    
    const number = Number(getCellValue(numberRef, context) ?? numberRef);
    const significance = significanceRef ? Number(getCellValue(significanceRef, context) ?? significanceRef) : 1;
    
    if (isNaN(number) || isNaN(significance)) {
      return FormulaError.VALUE;
    }
    
    if (significance === 0) {
      return 0;
    }
    
    const absSignificance = Math.abs(significance);
    return Math.ceil(number / absSignificance) * absSignificance;
  }
};

// FLOOR.MATH関数（数学的な切り下げ）
export const FLOOR_MATH: CustomFormula = {
  name: 'FLOOR.MATH',
  pattern: /FLOOR\.MATH\(([^,)]+)(?:,\s*([^,)]+))?(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, numberRef, significanceRef, modeRef] = matches;
    
    const number = Number(getCellValue(numberRef, context) ?? numberRef);
    const significance = significanceRef ? Number(getCellValue(significanceRef, context) ?? significanceRef) : 1;
    const mode = modeRef ? Number(getCellValue(modeRef, context) ?? modeRef) : 0;
    
    if (isNaN(number) || isNaN(significance) || isNaN(mode)) {
      return FormulaError.VALUE;
    }
    
    if (significance === 0) {
      return 0;
    }
    
    const absSignificance = Math.abs(significance);
    
    if (number >= 0 || mode === 0) {
      return Math.floor(number / absSignificance) * absSignificance;
    } else {
      return -Math.ceil(Math.abs(number) / absSignificance) * absSignificance;
    }
  }
};

// FLOOR.PRECISE関数（精密な切り下げ）
export const FLOOR_PRECISE: CustomFormula = {
  name: 'FLOOR.PRECISE',
  pattern: /FLOOR\.PRECISE\(([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, numberRef, significanceRef] = matches;
    
    const number = Number(getCellValue(numberRef, context) ?? numberRef);
    const significance = significanceRef ? Number(getCellValue(significanceRef, context) ?? significanceRef) : 1;
    
    if (isNaN(number) || isNaN(significance)) {
      return FormulaError.VALUE;
    }
    
    if (significance === 0) {
      return 0;
    }
    
    const absSignificance = Math.abs(significance);
    return Math.floor(number / absSignificance) * absSignificance;
  }
};

// ISO.CEILING関数（ISO標準の切り上げ）
export const ISO_CEILING: CustomFormula = {
  name: 'ISO.CEILING',
  pattern: /ISO\.CEILING\(([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, numberRef, significanceRef] = matches;
    
    const number = Number(getCellValue(numberRef, context) ?? numberRef);
    const significance = significanceRef ? Number(getCellValue(significanceRef, context) ?? significanceRef) : 1;
    
    if (isNaN(number) || isNaN(significance)) {
      return FormulaError.VALUE;
    }
    
    if (significance === 0) {
      return 0;
    }
    
    return Math.ceil(number / Math.abs(significance)) * Math.abs(significance);
  }
};

// SERIESSUM関数（べき級数を計算）
export const SERIESSUM: CustomFormula = {
  name: 'SERIESSUM',
  pattern: /SERIESSUM\(([^,]+),\s*([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: () => null // HyperFormulaに処理を委譲
};

// RANDARRAY関数（ランダム配列を生成）
export const RANDARRAY: CustomFormula = {
  name: 'RANDARRAY',
  pattern: /RANDARRAY\((?:([^,)]+))?(?:,\s*([^,)]+))?(?:,\s*([^,)]+))?(?:,\s*([^,)]+))?(?:,\s*([^)]+))?\)/i,
  calculate: (matches) => {
    const rows = matches[1] ? parseInt(matches[1]) : 1;
    const cols = matches[2] ? parseInt(matches[2]) : 1;
    const min = matches[3] ? parseFloat(matches[3]) : 0;
    const max = matches[4] ? parseFloat(matches[4]) : 1;
    const wholeNumber = matches[5] ? matches[5].toLowerCase() === 'true' : false;
    
    if (isNaN(rows) || isNaN(cols) || rows <= 0 || cols <= 0) return FormulaError.VALUE;
    if (isNaN(min) || isNaN(max) || min >= max) return FormulaError.VALUE;
    
    const result: number[][] = [];
    for (let i = 0; i < rows; i++) {
      const row: number[] = [];
      for (let j = 0; j < cols; j++) {
        const randomValue = Math.random() * (max - min) + min;
        row.push(wholeNumber ? Math.floor(randomValue) : randomValue);
      }
      result.push(row);
    }
    
    // 1x1の場合は単一の値を返す
    if (rows === 1 && cols === 1) {
      return result[0][0];
    }
    
    return result;
  }
};

// AGGREGATE関数用の計算ロジック
function performAggregateCalculation(functionNum: number, values: number[], k?: number): number | string {
  if (values.length === 0) return FormulaError.DIV0;
  
  switch (functionNum) {
    case 1: // AVERAGE
      return values.reduce((sum, val) => sum + val, 0) / values.length;
    case 2: // COUNT（数値のみカウント）
      return values.length;
    case 3: // COUNTA（空でない値をカウント）
      return values.length;
    case 4: // MAX
      return Math.max(...values);
    case 5: // MIN
      return Math.min(...values);
    case 6: // PRODUCT
      return values.reduce((prod, val) => prod * val, 1);
    case 9: // SUM
      return values.reduce((sum, val) => sum + val, 0);
    case 14: { // LARGE（k番目に大きい値）
      if (k === undefined || k <= 0 || k > values.length) return FormulaError.NUM;
      const sortedDesc = values.sort((a, b) => b - a);
      return sortedDesc[k - 1];
    }
    case 15: { // SMALL（k番目に小さい値）
      if (k === undefined || k <= 0 || k > values.length) return FormulaError.NUM;
      const sortedAsc = values.sort((a, b) => a - b);
      return sortedAsc[k - 1];
    }
    default:
      return FormulaError.VALUE;
  }
}

// AGGREGATE関数（集計関数、エラー値を除外）
export const AGGREGATE: CustomFormula = {
  name: 'AGGREGATE',
  pattern: /AGGREGATE\(([^,]+),\s*([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, functionNumRef, optionRef, arrayRef, kRef] = matches;
    
    const functionNum = Number(getCellValue(functionNumRef, context) ?? functionNumRef);
    const option = Number(getCellValue(optionRef, context) ?? optionRef);
    const k = kRef ? Number(getCellValue(kRef, context) ?? kRef) : undefined;
    
    if (isNaN(functionNum) || isNaN(option)) {
      return FormulaError.VALUE;
    }
    
    if (!Number.isInteger(functionNum) || !Number.isInteger(option)) {
      return FormulaError.NUM;
    }
    
    // 配列の値を取得
    const arrayValues = getCellRangeValues(arrayRef, context);
    let numbers = arrayValues.map(v => Number(v)).filter(n => !isNaN(n) && isFinite(n));
    
    // オプションに応じてエラー値や隠し行の処理
    if (option >= 4) {
      // エラー値を除外（オプション4-7はエラー値を無視）
      numbers = numbers.filter(n => !isNaN(n) && isFinite(n));
    }
    
    return performAggregateCalculation(functionNum, numbers, k);
  }
};

// SUBTOTAL関数（小計を計算）
export const SUBTOTAL: CustomFormula = {
  name: 'SUBTOTAL',
  pattern: /SUBTOTAL\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, functionNumRef, ref1] = matches;
    
    const functionNum = Number(getCellValue(functionNumRef, context) ?? functionNumRef);
    
    if (isNaN(functionNum) || !Number.isInteger(functionNum)) {
      return FormulaError.VALUE;
    }
    
    // 参照範囲の値を取得
    const values = getCellRangeValues(ref1, context);
    const numbers = values.map(v => Number(v)).filter(n => !isNaN(n) && isFinite(n));
    
    // 関数番号を基本関数に変換（100番台は非表示行を除外）
    const baseFunctionNum = functionNum > 100 ? functionNum - 100 : functionNum;
    
    // 基本的な計算を実行
    return performAggregateCalculation(baseFunctionNum, numbers);
  }
};

// MAXIFS関数（条件付き最大値）
export const MAXIFS: CustomFormula = {
  name: 'MAXIFS',
  pattern: /MAXIFS\(([^,]+),\s*([^,]+),\s*([^,]+)(?:,\s*([^,]+),\s*([^,]+))*\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, maxRange, ...criteriaArgs] = matches;
    
    try {
      // 最大値を求める範囲の値を取得
      const maxRangeValues = getCellRangeValues(maxRange.trim(), context);
      
      // 条件をペアに分ける
      const conditions: Array<{range: string, criteria: string}> = [];
      for (let i = 0; i < criteriaArgs.length; i += 2) {
        if (criteriaArgs[i] && criteriaArgs[i + 1]) {
          conditions.push({
            range: criteriaArgs[i].trim(),
            criteria: criteriaArgs[i + 1].replace(/^["']|["']$/g, '')
          });
        }
      }
      
      if (conditions.length === 0) {
        return FormulaError.VALUE;
      }
      
      const validValues: number[] = [];
      
      // 各値に対して条件をチェック
      for (let i = 0; i < maxRangeValues.length; i++) {
        let allConditionsMet = true;
        
        for (const condition of conditions) {
          const conditionRangeValues = getCellRangeValues(condition.range, context);
          if (i >= conditionRangeValues.length || !evaluateCondition(conditionRangeValues[i], condition.criteria)) {
            allConditionsMet = false;
            break;
          }
        }
        
        if (allConditionsMet) {
          const num = Number(maxRangeValues[i]);
          if (!isNaN(num)) {
            validValues.push(num);
          }
        }
      }
      
      if (validValues.length === 0) {
        return 0;
      }
      
      return Math.max(...validValues);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// MINIFS関数（条件付き最小値）
export const MINIFS: CustomFormula = {
  name: 'MINIFS',
  pattern: /MINIFS\(([^,]+),\s*([^,]+),\s*([^,]+)(?:,\s*([^,]+),\s*([^,]+))*\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, minRange, ...criteriaArgs] = matches;
    
    try {
      // 最小値を求める範囲の値を取得
      const minRangeValues = getCellRangeValues(minRange.trim(), context);
      
      // 条件をペアに分ける
      const conditions: Array<{range: string, criteria: string}> = [];
      for (let i = 0; i < criteriaArgs.length; i += 2) {
        if (criteriaArgs[i] && criteriaArgs[i + 1]) {
          conditions.push({
            range: criteriaArgs[i].trim(),
            criteria: criteriaArgs[i + 1].replace(/^["']|["']$/g, '')
          });
        }
      }
      
      if (conditions.length === 0) {
        return FormulaError.VALUE;
      }
      
      const validValues: number[] = [];
      
      // 各値に対して条件をチェック
      for (let i = 0; i < minRangeValues.length; i++) {
        let allConditionsMet = true;
        
        for (const condition of conditions) {
          const conditionRangeValues = getCellRangeValues(condition.range, context);
          if (i >= conditionRangeValues.length || !evaluateCondition(conditionRangeValues[i], condition.criteria)) {
            allConditionsMet = false;
            break;
          }
        }
        
        if (allConditionsMet) {
          const num = Number(minRangeValues[i]);
          if (!isNaN(num)) {
            validValues.push(num);
          }
        }
      }
      
      if (validValues.length === 0) {
        return 0;
      }
      
      return Math.min(...validValues);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// SEQUENCE関数（連続値を生成）
export const SEQUENCE: CustomFormula = {
  name: 'SEQUENCE',
  pattern: /SEQUENCE\(([^,)]+)(?:,\s*([^,)]+))?(?:,\s*([^,)]+))?(?:,\s*([^)]+))?\)/i,
  calculate: (matches) => {
    const rows = parseInt(matches[1]);
    const cols = matches[2] ? parseInt(matches[2]) : 1;
    const start = matches[3] ? parseFloat(matches[3]) : 1;
    const step = matches[4] ? parseFloat(matches[4]) : 1;
    
    if (isNaN(rows) || isNaN(cols) || rows <= 0 || cols <= 0) return FormulaError.VALUE;
    if (isNaN(start) || isNaN(step)) return FormulaError.VALUE;
    
    const result: number[][] = [];
    let currentValue = start;
    
    for (let i = 0; i < rows; i++) {
      const row: number[] = [];
      for (let j = 0; j < cols; j++) {
        row.push(currentValue);
        currentValue += step;
      }
      result.push(row);
    }
    
    // 1x1の場合は単一の値を返す
    if (rows === 1 && cols === 1) {
      return result[0][0];
    }
    
    return result;
  }
};

// 双曲線逆正弦（ASINH）

// 双曲線逆余弦（ACOSH）

// 双曲線逆正接（ATANH）

// 余割（CSC）
export const CSC: CustomFormula = {
  name: 'CSC',
  pattern: /CSC\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, angleRef] = matches;
    
    const angleValue = getCellValue(angleRef, context) ?? angleRef;
    const angle = Number(angleValue);
    
    if (isNaN(angle)) {
      return FormulaError.VALUE;
    }
    
    const sinValue = Math.sin(angle);
    
    if (Math.abs(sinValue) < 1e-10) {
      return FormulaError.DIV0;
    }
    
    return 1 / sinValue;
  }
};

// 正割（SEC）
export const SEC: CustomFormula = {
  name: 'SEC',
  pattern: /SEC\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, angleRef] = matches;
    
    const angleValue = getCellValue(angleRef, context) ?? angleRef;
    const angle = Number(angleValue);
    
    if (isNaN(angle)) {
      return FormulaError.VALUE;
    }
    
    const cosValue = Math.cos(angle);
    
    if (Math.abs(cosValue) < 1e-10) {
      return FormulaError.DIV0;
    }
    
    return 1 / cosValue;
  }
};

// 余接（COT）
export const COT: CustomFormula = {
  name: 'COT',
  pattern: /COT\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, angleRef] = matches;
    
    const angleValue = getCellValue(angleRef, context) ?? angleRef;
    const angle = Number(angleValue);
    
    if (isNaN(angle)) {
      return FormulaError.VALUE;
    }
    
    const tanValue = Math.tan(angle);
    
    if (Math.abs(tanValue) < 1e-10) {
      return FormulaError.DIV0;
    }
    
    return 1 / tanValue;
  }
};

// 逆余接（ACOT）
export const ACOT: CustomFormula = {
  name: 'ACOT',
  pattern: /ACOT\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, valueRef] = matches;
    
    const value = Number(getCellValue(valueRef, context) ?? valueRef);
    
    if (isNaN(value)) {
      return FormulaError.VALUE;
    }
    
    // acot(x) = atan(1/x) for x != 0
    // acot(0) = π/2
    if (value === 0) {
      return Math.PI / 2;
    }
    
    return Math.atan(1 / value);
  }
};

// 双曲線余割（CSCH）
export const CSCH: CustomFormula = {
  name: 'CSCH',
  pattern: /CSCH\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, valueRef] = matches;
    
    const value = Number(getCellValue(valueRef, context) ?? valueRef);
    
    if (isNaN(value)) {
      return FormulaError.VALUE;
    }
    
    const sinhValue = Math.sinh(value);
    
    if (Math.abs(sinhValue) < 1e-10) {
      return FormulaError.DIV0;
    }
    
    return 1 / sinhValue;
  }
};

// 双曲線正割（SECH）
export const SECH: CustomFormula = {
  name: 'SECH',
  pattern: /SECH\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, valueRef] = matches;
    
    const value = Number(getCellValue(valueRef, context) ?? valueRef);
    
    if (isNaN(value)) {
      return FormulaError.VALUE;
    }
    
    const coshValue = Math.cosh(value);
    
    if (Math.abs(coshValue) < 1e-10) {
      return FormulaError.DIV0;
    }
    
    return 1 / coshValue;
  }
};

// 双曲線余接（COTH）
export const COTH: CustomFormula = {
  name: 'COTH',
  pattern: /COTH\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, valueRef] = matches;
    
    const value = Number(getCellValue(valueRef, context) ?? valueRef);
    
    if (isNaN(value)) {
      return FormulaError.VALUE;
    }
    
    const tanhValue = Math.tanh(value);
    
    if (Math.abs(tanhValue) < 1e-10) {
      return FormulaError.DIV0;
    }
    
    return 1 / tanhValue;
  }
};

// 双曲線逆余接（ACOTH）
export const ACOTH: CustomFormula = {
  name: 'ACOTH',
  pattern: /ACOTH\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, valueRef] = matches;
    
    const value = Number(getCellValue(valueRef, context) ?? valueRef);
    
    if (isNaN(value)) {
      return FormulaError.VALUE;
    }
    
    if (Math.abs(value) <= 1) {
      return FormulaError.NUM;
    }
    
    // acoth(x) = 0.5 * ln((x+1)/(x-1))
    return 0.5 * Math.log((value + 1) / (value - 1));
  }
};