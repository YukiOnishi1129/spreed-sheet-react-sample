// 三角関数の実装

import type { CustomFormula, FormulaContext } from './types';
import { FormulaError } from './types';
import { getCellValue } from './utils';

// SIN関数（正弦）
export const SIN: CustomFormula = {
  name: 'SIN',
  pattern: /SIN\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, angleRef] = matches;
    
    const angle = Number(getCellValue(angleRef, context) ?? angleRef);
    
    if (isNaN(angle)) {
      return FormulaError.VALUE;
    }
    
    return Math.sin(angle);
  }
};

// COS関数（余弦）
export const COS: CustomFormula = {
  name: 'COS',
  pattern: /COS\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, angleRef] = matches;
    
    const angle = Number(getCellValue(angleRef, context) ?? angleRef);
    
    if (isNaN(angle)) {
      return FormulaError.VALUE;
    }
    
    return Math.cos(angle);
  }
};

// TAN関数（正接）
export const TAN: CustomFormula = {
  name: 'TAN',
  pattern: /TAN\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, angleRef] = matches;
    
    const angle = Number(getCellValue(angleRef, context) ?? angleRef);
    
    if (isNaN(angle)) {
      return FormulaError.VALUE;
    }
    
    return Math.tan(angle);
  }
};

// ASIN関数（逆正弦）
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

// ACOS関数（逆余弦）
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

// ATAN関数（逆正接）
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

// ATAN2関数（x,y座標から角度を計算）
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
    
    return Math.atan2(y, x);
  }
};

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

// ASINH関数（双曲線逆正弦）
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

// ACOSH関数（双曲線逆余弦）
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

// ATANH関数（双曲線逆正接）
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

// DEGREES関数（ラジアンを度に変換）
export const DEGREES: CustomFormula = {
  name: 'DEGREES',
  pattern: /DEGREES\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, angleRef] = matches;
    
    const angle = Number(getCellValue(angleRef, context) ?? angleRef);
    
    if (isNaN(angle)) {
      return FormulaError.VALUE;
    }
    
    return angle * (180 / Math.PI);
  }
};

// RADIANS関数（度をラジアンに変換）
export const RADIANS: CustomFormula = {
  name: 'RADIANS',
  pattern: /RADIANS\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, angleRef] = matches;
    
    const angle = Number(getCellValue(angleRef, context) ?? angleRef);
    
    if (isNaN(angle)) {
      return FormulaError.VALUE;
    }
    
    return angle * (Math.PI / 180);
  }
};

// PI関数（円周率を返す）
export const PI: CustomFormula = {
  name: 'PI',
  pattern: /PI\(\)/i,
  calculate: () => {
    return Math.PI;
  }
};

// CSC関数（余割）
export const CSC: CustomFormula = {
  name: 'CSC',
  pattern: /CSC\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, angleRef] = matches;
    
    const angle = Number(getCellValue(angleRef, context) ?? angleRef);
    
    if (isNaN(angle)) {
      return FormulaError.VALUE;
    }
    
    const sinValue = Math.sin(angle);
    if (Math.abs(sinValue) < Number.EPSILON) {
      return FormulaError.DIV0;
    }
    
    return 1 / sinValue;
  }
};

// SEC関数（正割）
export const SEC: CustomFormula = {
  name: 'SEC',
  pattern: /SEC\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, angleRef] = matches;
    
    const angle = Number(getCellValue(angleRef, context) ?? angleRef);
    
    if (isNaN(angle)) {
      return FormulaError.VALUE;
    }
    
    const cosValue = Math.cos(angle);
    if (Math.abs(cosValue) < Number.EPSILON) {
      return FormulaError.DIV0;
    }
    
    return 1 / cosValue;
  }
};

// COT関数（余接）
export const COT: CustomFormula = {
  name: 'COT',
  pattern: /COT\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, angleRef] = matches;
    
    const angle = Number(getCellValue(angleRef, context) ?? angleRef);
    
    if (isNaN(angle)) {
      return FormulaError.VALUE;
    }
    
    const tanValue = Math.tan(angle);
    if (Math.abs(tanValue) < Number.EPSILON) {
      return FormulaError.DIV0;
    }
    
    return 1 / tanValue;
  }
};

// ACOT関数（逆余接）
export const ACOT: CustomFormula = {
  name: 'ACOT',
  pattern: /ACOT\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, valueRef] = matches;
    
    const value = Number(getCellValue(valueRef, context) ?? valueRef);
    
    if (isNaN(value)) {
      return FormulaError.VALUE;
    }
    
    // ACOT(x) = ATAN(1/x) for x != 0
    // ACOT(0) = π/2
    if (Math.abs(value) < Number.EPSILON) {
      return Math.PI / 2;
    }
    
    return Math.atan(1 / value);
  }
};

// CSCH関数（双曲線余割）
export const CSCH: CustomFormula = {
  name: 'CSCH',
  pattern: /CSCH\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, valueRef] = matches;
    
    const value = Number(getCellValue(valueRef, context) ?? valueRef);
    
    if (isNaN(value)) {
      return FormulaError.VALUE;
    }
    
    if (value === 0) {
      return FormulaError.DIV0;
    }
    
    return 1 / Math.sinh(value);
  }
};

// SECH関数（双曲線正割）
export const SECH: CustomFormula = {
  name: 'SECH',
  pattern: /SECH\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, valueRef] = matches;
    
    const value = Number(getCellValue(valueRef, context) ?? valueRef);
    
    if (isNaN(value)) {
      return FormulaError.VALUE;
    }
    
    return 1 / Math.cosh(value);
  }
};

// COTH関数（双曲線余接）
export const COTH: CustomFormula = {
  name: 'COTH',
  pattern: /COTH\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, valueRef] = matches;
    
    const value = Number(getCellValue(valueRef, context) ?? valueRef);
    
    if (isNaN(value)) {
      return FormulaError.VALUE;
    }
    
    if (value === 0) {
      return FormulaError.DIV0;
    }
    
    return 1 / Math.tanh(value);
  }
};

// ACOTH関数（双曲線逆余接）
export const ACOTH: CustomFormula = {
  name: 'ACOTH',
  pattern: /ACOTH\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, valueRef] = matches;
    
    const value = Number(getCellValue(valueRef, context) ?? valueRef);
    
    if (isNaN(value) || Math.abs(value) <= 1) {
      return FormulaError.NUM;
    }
    
    // ACOTH(x) = 0.5 * ln((x+1)/(x-1))
    return 0.5 * Math.log((value + 1) / (value - 1));
  }
};