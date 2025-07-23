// 数学関数の実装

import type { CustomFormula, FormulaContext } from './types';
import { FormulaError } from './types';
import { getCellValue, getCellRangeValues } from './utils';

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

// SUMIF関数の実装
export const SUMIF: CustomFormula = {
  name: 'SUMIF',
  pattern: /SUMIF\(([^,]+),\s*"([^"]+)",\s*([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// COUNTIF関数の実装
export const COUNTIF: CustomFormula = {
  name: 'COUNTIF',
  pattern: /COUNTIF\(([^,]+),\s*"([^"]+)"\)/i,
  calculate: () => null // HyperFormulaが処理
};

// AVERAGEIF関数の実装
export const AVERAGEIF: CustomFormula = {
  name: 'AVERAGEIF',
  pattern: /AVERAGEIF\(([^,]+),\s*"([^"]+)",\s*([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
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
      return FormulaError.VALUE;
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
      return FormulaError.VALUE;
    }
    
    return Math.min(...numbers);
  }
};

// ROUND関数の実装
export const ROUND: CustomFormula = {
  name: 'ROUND',
  pattern: /ROUND\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, numberRef, digitsRef] = matches;
    
    const number = Number(getCellValue(numberRef, context) ?? numberRef);
    const digits = Number(getCellValue(digitsRef, context) ?? digitsRef);
    
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
    
    const base = Number(getCellValue(baseRef, context) ?? baseRef);
    const exponent = Number(getCellValue(exponentRef, context) ?? exponentRef);
    
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
export const SIN: CustomFormula = {
  name: 'SIN',
  pattern: /SIN\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, angleRef] = matches;
    
    const angleValue = getCellValue(angleRef, context) ?? angleRef;
    const angle = Number(angleValue);
    
    if (isNaN(angle)) {
      return FormulaError.VALUE;
    }
    
    return Math.sin(angle);
  }
};

// COS関数の実装（余弦）
export const COS: CustomFormula = {
  name: 'COS',
  pattern: /COS\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, angleRef] = matches;
    
    const angleValue = getCellValue(angleRef, context) ?? angleRef;
    const angle = Number(angleValue);
    
    if (isNaN(angle)) {
      return FormulaError.VALUE;
    }
    
    return Math.cos(angle);
  }
};

// TAN関数の実装（正接）
export const TAN: CustomFormula = {
  name: 'TAN',
  pattern: /TAN\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, angleRef] = matches;
    
    const angleValue = getCellValue(angleRef, context) ?? angleRef;
    const angle = Number(angleValue);
    
    if (isNaN(angle)) {
      return FormulaError.VALUE;
    }
    
    return Math.tan(angle);
  }
};

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
export const ASIN: CustomFormula = {
  name: 'ASIN',
  pattern: /ASIN\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, valueRef] = matches;
    
    const value = Number(getCellValue(valueRef, context) ?? valueRef);
    
    if (isNaN(value) || value < -1 || value > 1) {
      return FormulaError.NUM;
    }
    
    return Math.asin(value);
  }
};

// ACOS関数の実装（逆余弦）
export const ACOS: CustomFormula = {
  name: 'ACOS',
  pattern: /ACOS\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, valueRef] = matches;
    
    const value = Number(getCellValue(valueRef, context) ?? valueRef);
    
    if (isNaN(value) || value < -1 || value > 1) {
      return FormulaError.NUM;
    }
    
    return Math.acos(value);
  }
};

// ATAN関数の実装（逆正接）
export const ATAN: CustomFormula = {
  name: 'ATAN',
  pattern: /ATAN\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, valueRef] = matches;
    
    const value = Number(getCellValue(valueRef, context) ?? valueRef);
    
    if (isNaN(value)) {
      return FormulaError.VALUE;
    }
    
    return Math.atan(value);
  }
};

// ATAN2関数の実装（x,y座標から角度）
export const ATAN2: CustomFormula = {
  name: 'ATAN2',
  pattern: /ATAN2\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, xRef, yRef] = matches;
    
    const x = Number(getCellValue(xRef, context) ?? xRef);
    const y = Number(getCellValue(yRef, context) ?? yRef);
    
    if (isNaN(x) || isNaN(y)) {
      return FormulaError.VALUE;
    }
    
    if (x === 0 && y === 0) {
      return FormulaError.DIV0;
    }
    
    return Math.atan2(x, y);
  }
};

// ROUNDUP関数の実装（切り上げ）
export const ROUNDUP: CustomFormula = {
  name: 'ROUNDUP',
  pattern: /ROUNDUP\(([^,]+),\s*([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// ROUNDDOWN関数の実装（切り下げ）
export const ROUNDDOWN: CustomFormula = {
  name: 'ROUNDDOWN',
  pattern: /ROUNDDOWN\(([^,]+),\s*([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// CEILING関数の実装（基準値の倍数に切り上げ）
export const CEILING: CustomFormula = {
  name: 'CEILING',
  pattern: /CEILING\(([^,]+),\s*([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// FLOOR関数の実装（基準値の倍数に切り下げ）
export const FLOOR: CustomFormula = {
  name: 'FLOOR',
  pattern: /FLOOR\(([^,]+),\s*([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
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
  pattern: /SUMIFS\(([^,]+),\s*([^,]+),\s*"([^"]+)"(?:,\s*([^,]+),\s*"([^"]+)")*\)/i,
  calculate: () => null // HyperFormulaが処理
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
  calculate: () => null // HyperFormulaが処理
};

// MROUND関数の実装（倍数に丸める）
export const MROUND: CustomFormula = {
  name: 'MROUND',
  pattern: /MROUND\(([^,]+),\s*([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// COMBIN関数の実装（組み合わせ数）
export const COMBIN: CustomFormula = {
  name: 'COMBIN',
  pattern: /COMBIN\(([^,]+),\s*([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// GCD関数の実装（最大公約数）
export const GCD: CustomFormula = {
  name: 'GCD',
  pattern: /GCD\(([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// LCM関数の実装（最小公倍数）
export const LCM: CustomFormula = {
  name: 'LCM',
  pattern: /LCM\(([^)]+)\)/i,
  calculate: () => null // HyperFormulaに処理を委譲
};

// QUOTIENT関数の実装（商の整数部分）
export const QUOTIENT: CustomFormula = {
  name: 'QUOTIENT',
  pattern: /QUOTIENT\(([^,]+),\s*([^)]+)\)/i,
  calculate: () => null // HyperFormulaに処理を委譲
};

// 双曲線関数の実装
// SINH関数（双曲線正弦）
export const SINH: CustomFormula = {
  name: 'SINH',
  pattern: /SINH\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, valueRef] = matches;
    
    const value = Number(getCellValue(valueRef, context) ?? valueRef);
    
    if (isNaN(value)) {
      return FormulaError.VALUE;
    }
    
    return Math.sinh(value);
  }
};

// COSH関数（双曲線余弦）
export const COSH: CustomFormula = {
  name: 'COSH',
  pattern: /COSH\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, valueRef] = matches;
    
    const value = Number(getCellValue(valueRef, context) ?? valueRef);
    
    if (isNaN(value)) {
      return FormulaError.VALUE;
    }
    
    return Math.cosh(value);
  }
};

// TANH関数（双曲線正接）
export const TANH: CustomFormula = {
  name: 'TANH',
  pattern: /TANH\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, valueRef] = matches;
    
    const value = Number(getCellValue(valueRef, context) ?? valueRef);
    
    if (isNaN(value)) {
      return FormulaError.VALUE;
    }
    
    return Math.tanh(value);
  }
};

// HyperFormulaでサポートされている未実装関数を追加

// SUMSQ関数（平方和）
export const SUMSQ: CustomFormula = {
  name: 'SUMSQ',
  pattern: /SUMSQ\(([^)]+)\)/i,
  calculate: () => null // HyperFormulaに処理を委譲
};

// SUMPRODUCT関数（配列の積の和）
export const SUMPRODUCT: CustomFormula = {
  name: 'SUMPRODUCT',
  pattern: /SUMPRODUCT\(([^)]+)\)/i,
  calculate: () => null // HyperFormulaに処理を委譲
};

// EVEN関数（最も近い偶数に切り上げ）
export const EVEN: CustomFormula = {
  name: 'EVEN',
  pattern: /EVEN\(([^)]+)\)/i,
  calculate: () => null // HyperFormulaに処理を委譲
};

// ODD関数（最も近い奇数に切り上げ）
export const ODD: CustomFormula = {
  name: 'ODD',
  pattern: /ODD\(([^)]+)\)/i,
  calculate: () => null // HyperFormulaに処理を委譲
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
  calculate: () => null // HyperFormulaに処理を委譲
};

// COMBINA関数（重複組合せ）
export const COMBINA: CustomFormula = {
  name: 'COMBINA',
  pattern: /COMBINA\(([^,]+),\s*([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// PERMUT関数（順列）
export const PERMUT: CustomFormula = {
  name: 'PERMUT',
  pattern: /PERMUT\(([^,]+),\s*([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
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
  calculate: () => null // HyperFormulaに処理を委譲
};

// SQRTPI関数（π倍の平方根）
export const SQRTPI: CustomFormula = {
  name: 'SQRTPI',
  pattern: /SQRTPI\(([^)]+)\)/i,
  calculate: () => null // HyperFormulaに処理を委譲
};

// SUMX2MY2関数（x^2-y^2の和）
export const SUMX2MY2: CustomFormula = {
  name: 'SUMX2MY2',
  pattern: /SUMX2MY2\(([^,]+),\s*([^)]+)\)/i,
  calculate: () => null // HyperFormulaに処理を委譲
};

// SUMX2PY2関数（x^2+y^2の和）
export const SUMX2PY2: CustomFormula = {
  name: 'SUMX2PY2',
  pattern: /SUMX2PY2\(([^,]+),\s*([^)]+)\)/i,
  calculate: () => null // HyperFormulaに処理を委譲
};

// SUMXMY2関数（(x-y)^2の和）
export const SUMXMY2: CustomFormula = {
  name: 'SUMXMY2',
  pattern: /SUMXMY2\(([^,]+),\s*([^)]+)\)/i,
  calculate: () => null // HyperFormulaに処理を委譲
};

// MULTINOMIAL関数（多項係数）
export const MULTINOMIAL: CustomFormula = {
  name: 'MULTINOMIAL',
  pattern: /MULTINOMIAL\(([^)]+)\)/i,
  calculate: () => null // HyperFormulaに処理を委譲
};


// BASE関数（数値を指定した基数に変換）
export const BASE: CustomFormula = {
  name: 'BASE',
  pattern: /BASE\(([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: () => null // HyperFormulaに処理を委譲
};

// DECIMAL関数（指定した基数の数値を10進数に変換）
export const DECIMAL: CustomFormula = {
  name: 'DECIMAL',
  pattern: /DECIMAL\(([^,]+),\s*([^)]+)\)/i,
  calculate: () => null // HyperFormulaに処理を委譲
};

// SUBTOTAL関数（小計）
export const SUBTOTAL: CustomFormula = {
  name: 'SUBTOTAL',
  pattern: /SUBTOTAL\(([^,]+),\s*([^)]+)\)/i,
  calculate: () => null // HyperFormulaに処理を委譲
};

// AGGREGATE関数（集計関数、エラー値を除外）
export const AGGREGATE: CustomFormula = {
  name: 'AGGREGATE',
  pattern: /AGGREGATE\(([^,]+),\s*([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: () => null // HyperFormulaに処理を委譲
};

// CEILING.MATH関数（数学的な切り上げ）
export const CEILING_MATH: CustomFormula = {
  name: 'CEILING.MATH',
  pattern: /CEILING\.MATH\(([^,)]+)(?:,\s*([^,)]+))?(?:,\s*([^)]+))?\)/i,
  calculate: () => null // HyperFormulaに処理を委譲
};

// CEILING.PRECISE関数（精密な切り上げ）
export const CEILING_PRECISE: CustomFormula = {
  name: 'CEILING.PRECISE',
  pattern: /CEILING\.PRECISE\(([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: () => null // HyperFormulaに処理を委譲
};

// FLOOR.MATH関数（数学的な切り下げ）
export const FLOOR_MATH: CustomFormula = {
  name: 'FLOOR.MATH',
  pattern: /FLOOR\.MATH\(([^,)]+)(?:,\s*([^,)]+))?(?:,\s*([^)]+))?\)/i,
  calculate: () => null // HyperFormulaに処理を委譲
};

// FLOOR.PRECISE関数（精密な切り下げ）
export const FLOOR_PRECISE: CustomFormula = {
  name: 'FLOOR.PRECISE',
  pattern: /FLOOR\.PRECISE\(([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: () => null // HyperFormulaに処理を委譲
};

// ISO.CEILING関数（ISO標準の切り上げ）
export const ISO_CEILING: CustomFormula = {
  name: 'ISO.CEILING',
  pattern: /ISO\.CEILING\(([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: () => null // HyperFormulaに処理を委譲
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
export const ASINH: CustomFormula = {
  name: 'ASINH',
  pattern: /ASINH\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, valueRef] = matches;
    
    const value = Number(getCellValue(valueRef, context) ?? valueRef);
    
    if (isNaN(value)) {
      return FormulaError.VALUE;
    }
    
    return Math.asinh(value);
  }
};

// 双曲線逆余弦（ACOSH）
export const ACOSH: CustomFormula = {
  name: 'ACOSH',
  pattern: /ACOSH\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, valueRef] = matches;
    
    const value = Number(getCellValue(valueRef, context) ?? valueRef);
    
    if (isNaN(value) || value < 1) {
      return FormulaError.NUM;
    }
    
    return Math.acosh(value);
  }
};

// 双曲線逆正接（ATANH）
export const ATANH: CustomFormula = {
  name: 'ATANH',
  pattern: /ATANH\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, valueRef] = matches;
    
    const value = Number(getCellValue(valueRef, context) ?? valueRef);
    
    if (isNaN(value) || value <= -1 || value >= 1) {
      return FormulaError.NUM;
    }
    
    return Math.atanh(value);
  }
};

// 余割（CSC）
export const CSC: CustomFormula = {
  name: 'CSC',
  pattern: /CSC\(([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// 正割（SEC）
export const SEC: CustomFormula = {
  name: 'SEC',
  pattern: /SEC\(([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// 余接（COT）
export const COT: CustomFormula = {
  name: 'COT',
  pattern: /COT\(([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// 逆余接（ACOT）
export const ACOT: CustomFormula = {
  name: 'ACOT',
  pattern: /ACOT\(([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// 双曲線余割（CSCH）
export const CSCH: CustomFormula = {
  name: 'CSCH',
  pattern: /CSCH\(([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// 双曲線正割（SECH）
export const SECH: CustomFormula = {
  name: 'SECH',
  pattern: /SECH\(([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// 双曲線余接（COTH）
export const COTH: CustomFormula = {
  name: 'COTH',
  pattern: /COTH\(([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// 双曲線逆余接（ACOTH）
export const ACOTH: CustomFormula = {
  name: 'ACOTH',
  pattern: /ACOTH\(([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};