// 数学関数の実装

import type { CustomFormula } from './types';
import { FormulaError } from './types';
import { getCellValue } from './utils';

// SUMIF関数の実装
export const SUMIF: CustomFormula = {
  name: 'SUMIF',
  pattern: /SUMIF\(([^,]+),\s*"([^"]+)",\s*([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// COUNTIF関数の実装
export const COUNTIF: CustomFormula = {
  name: 'COUNTIF',
  pattern: /COUNTIF\(([^,]+),\s*"([^"]+)"\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// AVERAGEIF関数の実装
export const AVERAGEIF: CustomFormula = {
  name: 'AVERAGEIF',
  pattern: /AVERAGEIF\(([^,]+),\s*"([^"]+)",\s*([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// SUM関数の実装
export const SUM: CustomFormula = {
  name: 'SUM',
  pattern: /SUM\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// AVERAGE関数の実装
export const AVERAGE: CustomFormula = {
  name: 'AVERAGE',
  pattern: /AVERAGE\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// COUNT関数の実装
export const COUNT: CustomFormula = {
  name: 'COUNT',
  pattern: /COUNT\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// MAX関数の実装
export const MAX: CustomFormula = {
  name: 'MAX',
  pattern: /MAX\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// MIN関数の実装
export const MIN: CustomFormula = {
  name: 'MIN',
  pattern: /MIN\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// ROUND関数の実装
export const ROUND: CustomFormula = {
  name: 'ROUND',
  pattern: /ROUND\(([^,]+),\s*(\d+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// ABS関数の実装（絶対値）
export const ABS: CustomFormula = {
  name: 'ABS',
  pattern: /ABS\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// SQRT関数の実装（平方根）
export const SQRT: CustomFormula = {
  name: 'SQRT',
  pattern: /SQRT\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// POWER関数の実装（べき乗）
export const POWER: CustomFormula = {
  name: 'POWER',
  pattern: /POWER\(([^,]+),\s*([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// MOD関数の実装（剰余）
export const MOD: CustomFormula = {
  name: 'MOD',
  pattern: /MOD\(([^,]+),\s*([^)]+)\)/i,
  isSupported: false,
  calculate: (matches, context) => {
    const dividendRef = matches[1].trim();
    const divisorRef = matches[2].trim();
    
    let dividend: number, divisor: number;
    
    if (dividendRef.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(dividendRef, context);
      dividend = parseFloat(String(cellValue ?? '0'));
    } else {
      dividend = parseFloat(dividendRef);
    }
    
    if (divisorRef.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(divisorRef, context);
      divisor = parseFloat(String(cellValue ?? '0'));
    } else {
      divisor = parseFloat(divisorRef);
    }
    
    if (isNaN(dividend) || isNaN(divisor)) return FormulaError.VALUE;
    if (divisor === 0) return FormulaError.DIV0;
    return dividend % divisor;
  }
};

// INT関数の実装（整数部分）
export const INT: CustomFormula = {
  name: 'INT',
  pattern: /INT\(([^)]+)\)/i,
  isSupported: false,
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
  isSupported: false,
  calculate: (matches, context) => {
    const valueRef = matches[1].trim();
    const digitsRef = matches[2] ? matches[2].trim() : '0';
    
    let value: number, digits: number;
    
    if (valueRef.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(valueRef, context);
      value = parseFloat(String(cellValue ?? '0'));
    } else {
      value = parseFloat(valueRef);
    }
    
    if (digitsRef.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(digitsRef, context);
      digits = parseInt(String(cellValue ?? '0'));
    } else {
      digits = parseInt(digitsRef);
    }
    
    if (isNaN(value)) return FormulaError.VALUE;
    if (isNaN(digits)) return FormulaError.VALUE;
    
    const multiplier = Math.pow(10, digits);
    return Math.trunc(value * multiplier) / multiplier;
  }
};

// RAND関数の実装（0以上1未満の乱数）
export const RAND: CustomFormula = {
  name: 'RAND',
  pattern: /RAND\(\)/i,
  isSupported: false,
  calculate: () => {
    return Math.random();
  }
};

// RANDBETWEEN関数の実装（指定範囲の整数乱数）
export const RANDBETWEEN: CustomFormula = {
  name: 'RANDBETWEEN',
  pattern: /RANDBETWEEN\(([^,]+),\s*([^)]+)\)/i,
  isSupported: false,
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
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// DEGREES関数の実装（ラジアンを度に変換）
export const DEGREES: CustomFormula = {
  name: 'DEGREES',
  pattern: /DEGREES\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// RADIANS関数の実装（度をラジアンに変換）
export const RADIANS: CustomFormula = {
  name: 'RADIANS',
  pattern: /RADIANS\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// SIN関数の実装（正弦）
export const SIN: CustomFormula = {
  name: 'SIN',
  pattern: /SIN\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// COS関数の実装（余弦）
export const COS: CustomFormula = {
  name: 'COS',
  pattern: /COS\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// TAN関数の実装（正接）
export const TAN: CustomFormula = {
  name: 'TAN',
  pattern: /TAN\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// LOG関数の実装（対数）
export const LOG: CustomFormula = {
  name: 'LOG',
  pattern: /LOG\(([^,)]+)(?:,\s*([^)]+))?\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// LOG10関数の実装（常用対数）
export const LOG10: CustomFormula = {
  name: 'LOG10',
  pattern: /LOG10\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// LN関数の実装（自然対数）
export const LN: CustomFormula = {
  name: 'LN',
  pattern: /LN\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// EXP関数の実装（指数関数）
export const EXP: CustomFormula = {
  name: 'EXP',
  pattern: /EXP\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// ASIN関数の実装（逆正弦）
export const ASIN: CustomFormula = {
  name: 'ASIN',
  pattern: /ASIN\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// ACOS関数の実装（逆余弦）
export const ACOS: CustomFormula = {
  name: 'ACOS',
  pattern: /ACOS\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// ATAN関数の実装（逆正接）
export const ATAN: CustomFormula = {
  name: 'ATAN',
  pattern: /ATAN\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// ATAN2関数の実装（x,y座標から角度）
export const ATAN2: CustomFormula = {
  name: 'ATAN2',
  pattern: /ATAN2\(([^,]+),\s*([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// ROUNDUP関数の実装（切り上げ）
export const ROUNDUP: CustomFormula = {
  name: 'ROUNDUP',
  pattern: /ROUNDUP\(([^,]+),\s*([^)]+)\)/i,
  isSupported: false,
  calculate: (matches, context) => {
    const valueRef = matches[1].trim();
    const digitsRef = matches[2].trim();
    
    let value: number, digits: number;
    
    if (valueRef.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(valueRef, context);
      value = parseFloat(String(cellValue ?? '0'));
    } else {
      value = parseFloat(valueRef);
    }
    
    if (digitsRef.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(digitsRef, context);
      digits = parseInt(String(cellValue ?? '0'));
    } else {
      digits = parseInt(digitsRef);
    }
    
    if (isNaN(value) || isNaN(digits)) return FormulaError.VALUE;
    
    const multiplier = Math.pow(10, digits);
    return Math.ceil(value * multiplier) / multiplier;
  }
};

// ROUNDDOWN関数の実装（切り下げ）
export const ROUNDDOWN: CustomFormula = {
  name: 'ROUNDDOWN',
  pattern: /ROUNDDOWN\(([^,]+),\s*([^)]+)\)/i,
  isSupported: false,
  calculate: (matches, context) => {
    const valueRef = matches[1].trim();
    const digitsRef = matches[2].trim();
    
    let value: number, digits: number;
    
    if (valueRef.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(valueRef, context);
      value = parseFloat(String(cellValue ?? '0'));
    } else {
      value = parseFloat(valueRef);
    }
    
    if (digitsRef.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(digitsRef, context);
      digits = parseInt(String(cellValue ?? '0'));
    } else {
      digits = parseInt(digitsRef);
    }
    
    if (isNaN(value) || isNaN(digits)) return FormulaError.VALUE;
    
    const multiplier = Math.pow(10, digits);
    return Math.floor(value * multiplier) / multiplier;
  }
};

// CEILING関数の実装（基準値の倍数に切り上げ）
export const CEILING: CustomFormula = {
  name: 'CEILING',
  pattern: /CEILING\(([^,]+),\s*([^)]+)\)/i,
  isSupported: false,
  calculate: (matches, context) => {
    const valueRef = matches[1].trim();
    const significanceRef = matches[2].trim();
    
    let value: number, significance: number;
    
    if (valueRef.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(valueRef, context);
      value = parseFloat(String(cellValue ?? '0'));
    } else {
      value = parseFloat(valueRef);
    }
    
    if (significanceRef.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(significanceRef, context);
      significance = parseFloat(String(cellValue ?? '0'));
    } else {
      significance = parseFloat(significanceRef);
    }
    
    if (isNaN(value) || isNaN(significance)) return FormulaError.VALUE;
    if (significance === 0) return 0;
    if ((value > 0 && significance < 0) || (value < 0 && significance > 0)) return FormulaError.NUM;
    
    return Math.ceil(value / significance) * significance;
  }
};

// FLOOR関数の実装（基準値の倍数に切り下げ）
export const FLOOR: CustomFormula = {
  name: 'FLOOR',
  pattern: /FLOOR\(([^,]+),\s*([^)]+)\)/i,
  isSupported: false,
  calculate: (matches, context) => {
    const valueRef = matches[1].trim();
    const significanceRef = matches[2].trim();
    
    let value: number, significance: number;
    
    if (valueRef.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(valueRef, context);
      value = parseFloat(String(cellValue ?? '0'));
    } else {
      value = parseFloat(valueRef);
    }
    
    if (significanceRef.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(significanceRef, context);
      significance = parseFloat(String(cellValue ?? '0'));
    } else {
      significance = parseFloat(significanceRef);
    }
    
    if (isNaN(value) || isNaN(significance)) return FormulaError.VALUE;
    if (significance === 0) return 0;
    if ((value > 0 && significance < 0) || (value < 0 && significance > 0)) return FormulaError.NUM;
    
    return Math.floor(value / significance) * significance;
  }
};

// SIGN関数の実装（符号）
export const SIGN: CustomFormula = {
  name: 'SIGN',
  pattern: /SIGN\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// FACT関数の実装（階乗）
export const FACT: CustomFormula = {
  name: 'FACT',
  pattern: /FACT\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// SUMIFS関数の実装（複数条件での合計）
export const SUMIFS: CustomFormula = {
  name: 'SUMIFS',
  pattern: /SUMIFS\(([^,]+),\s*([^,]+),\s*"([^"]+)"(?:,\s*([^,]+),\s*"([^"]+)")*\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// COUNTIFS関数の実装（複数条件でのカウント）
export const COUNTIFS: CustomFormula = {
  name: 'COUNTIFS',
  pattern: /COUNTIFS\(([^,]+),\s*"([^"]+)"(?:,\s*([^,]+),\s*"([^"]+)")*\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// AVERAGEIFS関数の実装（複数条件での平均）
export const AVERAGEIFS: CustomFormula = {
  name: 'AVERAGEIFS',
  pattern: /AVERAGEIFS\(([^,]+),\s*([^,]+),\s*"([^"]+)"(?:,\s*([^,]+),\s*"([^"]+)")*\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// PRODUCT関数の実装（積を計算）
export const PRODUCT: CustomFormula = {
  name: 'PRODUCT',
  pattern: /PRODUCT\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// MROUND関数の実装（倍数に丸める）
export const MROUND: CustomFormula = {
  name: 'MROUND',
  pattern: /MROUND\(([^,]+),\s*([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// COMBIN関数の実装（組み合わせ数）
export const COMBIN: CustomFormula = {
  name: 'COMBIN',
  pattern: /COMBIN\(([^,]+),\s*([^)]+)\)/i,
  isSupported: false,
  calculate: (matches, context) => {
    const nRef = matches[1].trim();
    const kRef = matches[2].trim();
    
    let n: number, k: number;
    
    // n値を取得
    if (nRef.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(nRef, context);
      n = parseInt(String(cellValue ?? '0'));
    } else {
      n = parseInt(nRef);
    }
    
    // k値を取得
    if (kRef.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(kRef, context);
      k = parseInt(String(cellValue ?? '0'));
    } else {
      k = parseInt(kRef);
    }
    
    if (isNaN(n) || isNaN(k)) return FormulaError.VALUE;
    if (n < 0 || k < 0) return FormulaError.NUM;
    if (k > n) return 0;
    if (k === 0 || k === n) return 1;
    
    // 効率的な組み合わせ計算
    k = Math.min(k, n - k);
    let result = 1;
    for (let i = 0; i < k; i++) {
      result = result * (n - i) / (i + 1);
    }
    return Math.round(result);
  }
};

// PERMUT関数の実装（順列数）
export const PERMUT: CustomFormula = {
  name: 'PERMUT',
  pattern: /PERMUT\(([^,]+),\s*([^)]+)\)/i,
  isSupported: false,
  calculate: (matches, context) => {
    const nRef = matches[1].trim();
    const kRef = matches[2].trim();
    
    let n: number, k: number;
    
    // n値を取得
    if (nRef.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(nRef, context);
      n = parseInt(String(cellValue ?? '0'));
    } else {
      n = parseInt(nRef);
    }
    
    // k値を取得
    if (kRef.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(kRef, context);
      k = parseInt(String(cellValue ?? '0'));
    } else {
      k = parseInt(kRef);
    }
    
    if (isNaN(n) || isNaN(k)) return FormulaError.VALUE;
    if (n < 0 || k < 0) return FormulaError.NUM;
    if (k > n) return 0;
    
    let result = 1;
    for (let i = 0; i < k; i++) {
      result *= (n - i);
    }
    return result;
  }
};

// GCD関数の実装（最大公約数）
export const GCD: CustomFormula = {
  name: 'GCD',
  pattern: /GCD\(([^)]+)\)/i,
  isSupported: false,
  calculate: (matches, context) => {
    const args = matches[1].split(',').map(arg => arg.trim());
    const numbers: number[] = [];
    
    for (const arg of args) {
      let num: number;
      if (arg.match(/^[A-Z]+\d+$/)) {
        const cellValue = getCellValue(arg, context);
        num = parseInt(String(cellValue ?? '0'));
      } else {
        num = parseInt(arg);
      }
      
      if (isNaN(num)) return FormulaError.VALUE;
      if (num < 0) return FormulaError.NUM;
      numbers.push(num);
    }
    
    if (numbers.length === 0) return FormulaError.VALUE;
    
    // ユークリッドの互除法
    const gcd = (a: number, b: number): number => {
      return b === 0 ? a : gcd(b, a % b);
    };
    
    let result = numbers[0];
    for (let i = 1; i < numbers.length; i++) {
      result = gcd(result, numbers[i]);
    }
    return result;
  }
};

// LCM関数の実装（最小公倍数）
export const LCM: CustomFormula = {
  name: 'LCM',
  pattern: /LCM\(([^)]+)\)/i,
  isSupported: false,
  calculate: (matches, context) => {
    const args = matches[1].split(',').map(arg => arg.trim());
    const numbers: number[] = [];
    
    for (const arg of args) {
      let num: number;
      if (arg.match(/^[A-Z]+\d+$/)) {
        const cellValue = getCellValue(arg, context);
        num = parseInt(String(cellValue ?? '0'));
      } else {
        num = parseInt(arg);
      }
      
      if (isNaN(num)) return FormulaError.VALUE;
      if (num <= 0) return FormulaError.NUM;
      numbers.push(num);
    }
    
    if (numbers.length === 0) return FormulaError.VALUE;
    
    // ユークリッドの互除法でGCDを計算
    const gcd = (a: number, b: number): number => {
      return b === 0 ? a : gcd(b, a % b);
    };
    
    // LCM = (a * b) / GCD(a, b)
    const lcm = (a: number, b: number): number => {
      return (a * b) / gcd(a, b);
    };
    
    let result = numbers[0];
    for (let i = 1; i < numbers.length; i++) {
      result = lcm(result, numbers[i]);
    }
    return result;
  }
};

// QUOTIENT関数の実装（商の整数部分）
export const QUOTIENT: CustomFormula = {
  name: 'QUOTIENT',
  pattern: /QUOTIENT\(([^,]+),\s*([^)]+)\)/i,
  isSupported: false,
  calculate: (matches, context) => {
    const numeratorRef = matches[1].trim();
    const denominatorRef = matches[2].trim();
    
    let numerator: number, denominator: number;
    
    // 分子を取得
    if (numeratorRef.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(numeratorRef, context);
      numerator = parseFloat(String(cellValue ?? '0'));
    } else {
      numerator = parseFloat(numeratorRef);
    }
    
    // 分母を取得
    if (denominatorRef.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(denominatorRef, context);
      denominator = parseFloat(String(cellValue ?? '0'));
    } else {
      denominator = parseFloat(denominatorRef);
    }
    
    if (isNaN(numerator) || isNaN(denominator)) return FormulaError.VALUE;
    if (denominator === 0) return FormulaError.DIV0;
    
    return Math.trunc(numerator / denominator);
  }
};

// 双曲線関数の実装
// SINH関数（双曲線正弦）
export const SINH: CustomFormula = {
  name: 'SINH',
  pattern: /SINH\(([^)]+)\)/i,
  isSupported: false,
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
    return Math.sinh(value);
  }
};

// COSH関数（双曲線余弦）
export const COSH: CustomFormula = {
  name: 'COSH',
  pattern: /COSH\(([^)]+)\)/i,
  isSupported: false,
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
    return Math.cosh(value);
  }
};

// TANH関数（双曲線正接）
export const TANH: CustomFormula = {
  name: 'TANH',
  pattern: /TANH\(([^)]+)\)/i,
  isSupported: false,
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
    return Math.tanh(value);
  }
};