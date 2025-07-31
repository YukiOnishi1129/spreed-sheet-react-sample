// 三角関数の実装（高精度版）

import type { CustomFormula, FormulaContext } from '../shared/types';
import { FormulaError } from '../shared/types';
import { getCellValue } from '../shared/utils';

// 高精度三角関数のためのユーティリティ
const HALF_PI = Math.PI / 2;
const TWO_PI = Math.PI * 2;

// 角度を[-π, π]の範囲に正規化
function normalizeAngle(angle: number): number {
  // 非常に大きな値の場合、modulo演算で精度を保つ
  if (Math.abs(angle) > 1e8) {
    angle = angle % TWO_PI;
  }
  
  // [-π, π]の範囲に調整
  while (angle > Math.PI) {
    angle -= TWO_PI;
  }
  while (angle < -Math.PI) {
    angle += TWO_PI;
  }
  
  return angle;
}

// 高精度sin計算（極端な値での精度向上）
function preciseSin(angle: number): number {
  // 無限大やNaNのチェック
  if (!isFinite(angle)) {
    return NaN;
  }
  
  // 角度を正規化
  const normalizedAngle = normalizeAngle(angle);
  
  // 特殊値の処理
  if (Math.abs(normalizedAngle) < Number.EPSILON) return 0;
  if (Math.abs(normalizedAngle - Math.PI) < Number.EPSILON) return 0;
  if (Math.abs(normalizedAngle - HALF_PI) < Number.EPSILON) return 1;
  if (Math.abs(normalizedAngle + HALF_PI) < Number.EPSILON) return -1;
  
  return Math.sin(normalizedAngle);
}

// 高精度cos計算
function preciseCos(angle: number): number {
  if (!isFinite(angle)) {
    return NaN;
  }
  
  const normalizedAngle = normalizeAngle(angle);
  
  // 特殊値の処理
  if (Math.abs(normalizedAngle) < Number.EPSILON) return 1;
  if (Math.abs(normalizedAngle - Math.PI) < Number.EPSILON) return -1;
  if (Math.abs(normalizedAngle - HALF_PI) < Number.EPSILON) return 0;
  if (Math.abs(normalizedAngle + HALF_PI) < Number.EPSILON) return 0;
  
  return Math.cos(normalizedAngle);
}

// 高精度tan計算
function preciseTan(angle: number): number {
  if (!isFinite(angle)) {
    return NaN;
  }
  
  const normalizedAngle = normalizeAngle(angle);
  
  // π/2 + nπ付近での特殊処理（無限大になる点）
  const halfPiCheck = Math.abs(Math.abs(normalizedAngle) - HALF_PI);
  if (halfPiCheck < 1e-15) {
    // Excel準拠：非常に大きな値を返すが無限大ではない
    return normalizedAngle > 0 ? 1e308 : -1e308;
  }
  
  // 特殊値の処理
  if (Math.abs(normalizedAngle) < Number.EPSILON) return 0;
  if (Math.abs(normalizedAngle - Math.PI) < Number.EPSILON) return 0;
  
  return Math.tan(normalizedAngle);
}

// 高精度逆三角関数
function preciseAsin(value: number): number {
  // 定義域チェック
  if (value < -1 || value > 1) {
    return NaN;
  }
  
  // 特殊値の処理
  if (Math.abs(value) < Number.EPSILON) return 0;
  if (Math.abs(value - 1) < Number.EPSILON) return HALF_PI;
  if (Math.abs(value + 1) < Number.EPSILON) return -HALF_PI;
  
  // 精度向上のため、1に近い値では別の計算方法を使用
  if (Math.abs(value) > 0.9) {
    if (value > 0) {
      return HALF_PI - Math.acos(value);
    } else {
      return -HALF_PI + Math.acos(-value);
    }
  }
  
  return Math.asin(value);
}

function preciseAcos(value: number): number {
  // 定義域チェック
  if (value < -1 || value > 1) {
    return NaN;
  }
  
  // 特殊値の処理
  if (Math.abs(value - 1) < Number.EPSILON) return 0;
  if (Math.abs(value + 1) < Number.EPSILON) return Math.PI;
  if (Math.abs(value) < Number.EPSILON) return HALF_PI;
  
  // 精度向上のため、1に近い値では別の計算方法を使用
  if (Math.abs(value) > 0.9) {
    if (value > 0) {
      return Math.asin(Math.sqrt(1 - value * value));
    } else {
      return Math.PI - Math.asin(Math.sqrt(1 - value * value));
    }
  }
  
  return Math.acos(value);
}

function preciseAtan(value: number): number {
  if (!isFinite(value)) {
    if (value > 0) return HALF_PI;
    if (value < 0) return -HALF_PI;
    return NaN;
  }
  
  // 特殊値の処理
  if (Math.abs(value) < Number.EPSILON) return 0;
  
  return Math.atan(value);
}

// 高精度双曲線関数（オーバーフロー対策）
function preciseSinh(value: number): number {
  if (!isFinite(value)) {
    return value;
  }
  
  // 大きな値でのオーバーフロー対策
  if (Math.abs(value) > 700) {
    const sign = value > 0 ? 1 : -1;
    return sign * Infinity;
  }
  
  // 非常に小さな値での精度向上
  if (Math.abs(value) < 1e-10) {
    return value; // sinh(x) ≈ x for small x
  }
  
  return Math.sinh(value);
}

function preciseCosh(value: number): number {
  if (!isFinite(value)) {
    return Math.abs(value) === Infinity ? Infinity : NaN;
  }
  
  // 大きな値でのオーバーフロー対策
  if (Math.abs(value) > 700) {
    return Infinity;
  }
  
  return Math.cosh(value);
}

function preciseTanh(value: number): number {
  if (!isFinite(value)) {
    if (value > 0) return 1;
    if (value < 0) return -1;
    return NaN;
  }
  
  // 大きな値では飽和値を返す
  if (value > 20) return 1;
  if (value < -20) return -1;
  
  // 非常に小さな値での精度向上
  if (Math.abs(value) < 1e-10) {
    return value; // tanh(x) ≈ x for small x
  }
  
  return Math.tanh(value);
}

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
    
    return preciseSin(angle);
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
    
    return preciseCos(angle);
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
    
    return preciseTan(angle);
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
    
    const result = preciseAsin(value);
    return isNaN(result) ? FormulaError.NUM : result;
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
    
    const result = preciseAcos(value);
    return isNaN(result) ? FormulaError.NUM : result;
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
    
    return preciseAtan(value);
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
    
    return preciseSinh(value);
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
    
    return preciseCosh(value);
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
    
    return preciseTanh(value);
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