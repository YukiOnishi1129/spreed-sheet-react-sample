// 数学関数の実装

import type { CustomFormula } from './types';

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
  isSupported: false,
  calculate: (matches) => {
    const value = parseFloat(matches[1]);
    if (isNaN(value)) return '#VALUE!';
    return Math.abs(value);
  }
};

// SQRT関数の実装（平方根）
export const SQRT: CustomFormula = {
  name: 'SQRT',
  pattern: /SQRT\(([^)]+)\)/i,
  isSupported: false,
  calculate: (matches) => {
    const value = parseFloat(matches[1]);
    if (isNaN(value)) return '#VALUE!';
    if (value < 0) return '#NUM!';
    return Math.sqrt(value);
  }
};

// POWER関数の実装（べき乗）
export const POWER: CustomFormula = {
  name: 'POWER',
  pattern: /POWER\(([^,]+),\s*([^)]+)\)/i,
  isSupported: false,
  calculate: (matches) => {
    const base = parseFloat(matches[1]);
    const exponent = parseFloat(matches[2]);
    if (isNaN(base) || isNaN(exponent)) return '#VALUE!';
    return Math.pow(base, exponent);
  }
};

// MOD関数の実装（剰余）
export const MOD: CustomFormula = {
  name: 'MOD',
  pattern: /MOD\(([^,]+),\s*([^)]+)\)/i,
  isSupported: false,
  calculate: (matches) => {
    const dividend = parseFloat(matches[1]);
    const divisor = parseFloat(matches[2]);
    if (isNaN(dividend) || isNaN(divisor)) return '#VALUE!';
    if (divisor === 0) return '#DIV/0!';
    return dividend % divisor;
  }
};

// INT関数の実装（整数部分）
export const INT: CustomFormula = {
  name: 'INT',
  pattern: /INT\(([^)]+)\)/i,
  isSupported: false,
  calculate: (matches) => {
    const value = parseFloat(matches[1]);
    if (isNaN(value)) return '#VALUE!';
    return Math.floor(value);
  }
};

// TRUNC関数の実装（小数部分切り捨て）
export const TRUNC: CustomFormula = {
  name: 'TRUNC',
  pattern: /TRUNC\(([^,)]+)(?:,\s*([^)]+))?\)/i,
  isSupported: false,
  calculate: (matches) => {
    const value = parseFloat(matches[1]);
    const digits = matches[2] ? parseInt(matches[2]) : 0;
    if (isNaN(value)) return '#VALUE!';
    if (isNaN(digits)) return '#VALUE!';
    
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
  calculate: (matches) => {
    const min = parseInt(matches[1]);
    const max = parseInt(matches[2]);
    if (isNaN(min) || isNaN(max)) return '#VALUE!';
    if (min > max) return '#NUM!';
    return Math.floor(Math.random() * (max - min + 1)) + min;
  }
};

// PI関数の実装（円周率）
export const PI: CustomFormula = {
  name: 'PI',
  pattern: /PI\(\)/i,
  isSupported: false,
  calculate: () => {
    return Math.PI;
  }
};

// DEGREES関数の実装（ラジアンを度に変換）
export const DEGREES: CustomFormula = {
  name: 'DEGREES',
  pattern: /DEGREES\(([^)]+)\)/i,
  isSupported: false,
  calculate: (matches) => {
    const radians = parseFloat(matches[1]);
    if (isNaN(radians)) return '#VALUE!';
    return radians * (180 / Math.PI);
  }
};

// RADIANS関数の実装（度をラジアンに変換）
export const RADIANS: CustomFormula = {
  name: 'RADIANS',
  pattern: /RADIANS\(([^)]+)\)/i,
  isSupported: false,
  calculate: (matches) => {
    const degrees = parseFloat(matches[1]);
    if (isNaN(degrees)) return '#VALUE!';
    return degrees * (Math.PI / 180);
  }
};

// SIN関数の実装（正弦）
export const SIN: CustomFormula = {
  name: 'SIN',
  pattern: /SIN\(([^)]+)\)/i,
  isSupported: false,
  calculate: (matches) => {
    const angle = parseFloat(matches[1]);
    if (isNaN(angle)) return '#VALUE!';
    return Math.sin(angle);
  }
};

// COS関数の実装（余弦）
export const COS: CustomFormula = {
  name: 'COS',
  pattern: /COS\(([^)]+)\)/i,
  isSupported: false,
  calculate: (matches) => {
    const angle = parseFloat(matches[1]);
    if (isNaN(angle)) return '#VALUE!';
    return Math.cos(angle);
  }
};

// TAN関数の実装（正接）
export const TAN: CustomFormula = {
  name: 'TAN',
  pattern: /TAN\(([^)]+)\)/i,
  isSupported: false,
  calculate: (matches) => {
    const angle = parseFloat(matches[1]);
    if (isNaN(angle)) return '#VALUE!';
    return Math.tan(angle);
  }
};

// LOG関数の実装（対数）
export const LOG: CustomFormula = {
  name: 'LOG',
  pattern: /LOG\(([^,)]+)(?:,\s*([^)]+))?\)/i,
  isSupported: false,
  calculate: (matches) => {
    const value = parseFloat(matches[1]);
    const base = matches[2] ? parseFloat(matches[2]) : 10;
    
    if (isNaN(value) || isNaN(base)) return '#VALUE!';
    if (value <= 0 || base <= 0 || base === 1) return '#NUM!';
    
    return Math.log(value) / Math.log(base);
  }
};

// LOG10関数の実装（常用対数）
export const LOG10: CustomFormula = {
  name: 'LOG10',
  pattern: /LOG10\(([^)]+)\)/i,
  isSupported: false,
  calculate: (matches) => {
    const value = parseFloat(matches[1]);
    if (isNaN(value)) return '#VALUE!';
    if (value <= 0) return '#NUM!';
    return Math.log10(value);
  }
};

// LN関数の実装（自然対数）
export const LN: CustomFormula = {
  name: 'LN',
  pattern: /LN\(([^)]+)\)/i,
  isSupported: false,
  calculate: (matches) => {
    const value = parseFloat(matches[1]);
    if (isNaN(value)) return '#VALUE!';
    if (value <= 0) return '#NUM!';
    return Math.log(value);
  }
};

// EXP関数の実装（指数関数）
export const EXP: CustomFormula = {
  name: 'EXP',
  pattern: /EXP\(([^)]+)\)/i,
  isSupported: false,
  calculate: (matches) => {
    const value = parseFloat(matches[1]);
    if (isNaN(value)) return '#VALUE!';
    return Math.exp(value);
  }
};

// ASIN関数の実装（逆正弦）
export const ASIN: CustomFormula = {
  name: 'ASIN',
  pattern: /ASIN\(([^)]+)\)/i,
  isSupported: false,
  calculate: (matches) => {
    const value = parseFloat(matches[1]);
    if (isNaN(value)) return '#VALUE!';
    if (value < -1 || value > 1) return '#NUM!';
    return Math.asin(value);
  }
};

// ACOS関数の実装（逆余弦）
export const ACOS: CustomFormula = {
  name: 'ACOS',
  pattern: /ACOS\(([^)]+)\)/i,
  isSupported: false,
  calculate: (matches) => {
    const value = parseFloat(matches[1]);
    if (isNaN(value)) return '#VALUE!';
    if (value < -1 || value > 1) return '#NUM!';
    return Math.acos(value);
  }
};

// ATAN関数の実装（逆正接）
export const ATAN: CustomFormula = {
  name: 'ATAN',
  pattern: /ATAN\(([^)]+)\)/i,
  isSupported: false,
  calculate: (matches) => {
    const value = parseFloat(matches[1]);
    if (isNaN(value)) return '#VALUE!';
    return Math.atan(value);
  }
};

// ATAN2関数の実装（x,y座標から角度）
export const ATAN2: CustomFormula = {
  name: 'ATAN2',
  pattern: /ATAN2\(([^,]+),\s*([^)]+)\)/i,
  isSupported: false,
  calculate: (matches) => {
    const x = parseFloat(matches[1]);
    const y = parseFloat(matches[2]);
    if (isNaN(x) || isNaN(y)) return '#VALUE!';
    return Math.atan2(y, x);
  }
};

// ROUNDUP関数の実装（切り上げ）
export const ROUNDUP: CustomFormula = {
  name: 'ROUNDUP',
  pattern: /ROUNDUP\(([^,]+),\s*([^)]+)\)/i,
  isSupported: false,
  calculate: (matches) => {
    const value = parseFloat(matches[1]);
    const digits = parseInt(matches[2]);
    
    if (isNaN(value) || isNaN(digits)) return '#VALUE!';
    
    const multiplier = Math.pow(10, digits);
    return Math.ceil(value * multiplier) / multiplier;
  }
};

// ROUNDDOWN関数の実装（切り下げ）
export const ROUNDDOWN: CustomFormula = {
  name: 'ROUNDDOWN',
  pattern: /ROUNDDOWN\(([^,]+),\s*([^)]+)\)/i,
  isSupported: false,
  calculate: (matches) => {
    const value = parseFloat(matches[1]);
    const digits = parseInt(matches[2]);
    
    if (isNaN(value) || isNaN(digits)) return '#VALUE!';
    
    const multiplier = Math.pow(10, digits);
    return Math.floor(value * multiplier) / multiplier;
  }
};

// CEILING関数の実装（基準値の倍数に切り上げ）
export const CEILING: CustomFormula = {
  name: 'CEILING',
  pattern: /CEILING\(([^,]+),\s*([^)]+)\)/i,
  isSupported: false,
  calculate: (matches) => {
    const value = parseFloat(matches[1]);
    const significance = parseFloat(matches[2]);
    
    if (isNaN(value) || isNaN(significance)) return '#VALUE!';
    if (significance === 0) return 0;
    if ((value > 0 && significance < 0) || (value < 0 && significance > 0)) return '#NUM!';
    
    return Math.ceil(value / significance) * significance;
  }
};

// FLOOR関数の実装（基準値の倍数に切り下げ）
export const FLOOR: CustomFormula = {
  name: 'FLOOR',
  pattern: /FLOOR\(([^,]+),\s*([^)]+)\)/i,
  isSupported: false,
  calculate: (matches) => {
    const value = parseFloat(matches[1]);
    const significance = parseFloat(matches[2]);
    
    if (isNaN(value) || isNaN(significance)) return '#VALUE!';
    if (significance === 0) return 0;
    if ((value > 0 && significance < 0) || (value < 0 && significance > 0)) return '#NUM!';
    
    return Math.floor(value / significance) * significance;
  }
};

// SIGN関数の実装（符号）
export const SIGN: CustomFormula = {
  name: 'SIGN',
  pattern: /SIGN\(([^)]+)\)/i,
  isSupported: false,
  calculate: (matches) => {
    const value = parseFloat(matches[1]);
    if (isNaN(value)) return '#VALUE!';
    return value > 0 ? 1 : value < 0 ? -1 : 0;
  }
};

// FACT関数の実装（階乗）
export const FACT: CustomFormula = {
  name: 'FACT',
  pattern: /FACT\(([^)]+)\)/i,
  isSupported: false,
  calculate: (matches) => {
    const value = parseInt(matches[1]);
    if (isNaN(value)) return '#VALUE!';
    if (value < 0) return '#NUM!';
    if (value > 170) return '#NUM!'; // 階乗が大きすぎる場合
    
    let result = 1;
    for (let i = 2; i <= value; i++) {
      result *= i;
    }
    return result;
  }
};