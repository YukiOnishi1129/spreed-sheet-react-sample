// ベッセル関数の実装

import type { CustomFormula, FormulaContext, FormulaResult } from '../shared/types';
import { FormulaError } from '../shared/types';
import { getCellValue } from '../shared/utils';

// ベッセル関数の計算のための補助関数
function factorial(n: number): number {
  if (n < 0 || n !== Math.floor(n)) return NaN;
  if (n === 0 || n === 1) return 1;
  let result = 1;
  for (let i = 2; i <= n; i++) {
    result *= i;
  }
  return result;
}

// ガンマ関数（整数の場合）
function gamma(n: number): number {
  if (n > 0 && n === Math.floor(n)) {
    return factorial(n - 1);
  }
  // Stirlingの近似
  if (n > 10) {
    return Math.sqrt(2 * Math.PI / n) * Math.pow(n / Math.E, n);
  }
  // Lanczosの近似
  const g = 7;
  const coef = [
    0.99999999999980993,
    676.5203681218851,
    -1259.1392167224028,
    771.32342877765313,
    -176.61502916214059,
    12.507343278686905,
    -0.13857109526572012,
    9.9843695780195716e-6,
    1.5056327351493116e-7
  ];
  
  let z = n - 1;
  let x = coef[0];
  for (let i = 1; i < 9; i++) {
    x += coef[i] / (z + i);
  }
  
  let t = z + g + 0.5;
  return Math.sqrt(2 * Math.PI) * Math.pow(t, z + 0.5) * Math.exp(-t) * x;
}

// BESSELJ関数の実装（ベッセル関数Jn(x)）
export const BESSELJ: CustomFormula = {
  name: 'BESSELJ',
  pattern: /BESSELJ\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, xRef, nRef] = matches;
    
    try {
      const x = parseFloat(getCellValue(xRef.trim(), context)?.toString() ?? xRef.trim());
      const n = parseInt(getCellValue(nRef.trim(), context)?.toString() ?? nRef.trim());
      
      if (isNaN(x) || isNaN(n)) {
        return FormulaError.VALUE;
      }
      
      if (n < 0) {
        return FormulaError.NUM;
      }
      
      // ベッセル関数J_n(x)の級数展開
      // J_n(x) = Σ(k=0 to ∞) [(-1)^k / (k! Γ(n+k+1))] * (x/2)^(n+2k)
      let result = 0;
      const maxTerms = 100;
      
      for (let k = 0; k < maxTerms; k++) {
        const term = Math.pow(-1, k) / (factorial(k) * gamma(n + k + 1)) * 
                     Math.pow(x / 2, n + 2 * k);
        
        if (Math.abs(term) < 1e-15) break;
        result += term;
      }
      
      return result;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// BESSELY関数の実装（ベッセル関数Yn(x)）
export const BESSELY: CustomFormula = {
  name: 'BESSELY',
  pattern: /BESSELY\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, xRef, nRef] = matches;
    
    try {
      const x = parseFloat(getCellValue(xRef.trim(), context)?.toString() ?? xRef.trim());
      const n = parseInt(getCellValue(nRef.trim(), context)?.toString() ?? nRef.trim());
      
      if (isNaN(x) || isNaN(n)) {
        return FormulaError.VALUE;
      }
      
      if (x <= 0 || n < 0) {
        return FormulaError.NUM;
      }
      
      // Y_n(x)の計算（近似実装）
      // より正確な実装には数値積分が必要
      if (n === 0) {
        // Y_0(x)の近似
        if (x < 3) {
          return (2 / Math.PI) * (Math.log(x / 2) * besselJ0(x) + 
                 0.07832358 - 0.0420024 * x * x);
        } else {
          const sqrtx = Math.sqrt(2 / (Math.PI * x));
          return sqrtx * Math.sin(x - Math.PI / 4);
        }
      } else if (n === 1) {
        // Y_1(x)の近似
        if (x < 3) {
          return (2 / Math.PI) * (Math.log(x / 2) * besselJ1(x) - 1 / x);
        } else {
          const sqrtx = Math.sqrt(2 / (Math.PI * x));
          return sqrtx * Math.sin(x - 3 * Math.PI / 4);
        }
      }
      
      // 再帰関係を使用
      // Y_{n+1}(x) = (2n/x) * Y_n(x) - Y_{n-1}(x)
      let y0 = bessely0Approx(x);
      let y1 = bessely1Approx(x);
      let yn = y1;
      
      for (let i = 1; i < n; i++) {
        const temp = (2 * i / x) * y1 - y0;
        y0 = y1;
        y1 = temp;
        yn = temp;
      }
      
      return yn;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// BESSELI関数の実装（修正ベッセル関数In(x)）
export const BESSELI: CustomFormula = {
  name: 'BESSELI',
  pattern: /BESSELI\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, xRef, nRef] = matches;
    
    try {
      const x = parseFloat(getCellValue(xRef.trim(), context)?.toString() ?? xRef.trim());
      const n = parseInt(getCellValue(nRef.trim(), context)?.toString() ?? nRef.trim());
      
      if (isNaN(x) || isNaN(n)) {
        return FormulaError.VALUE;
      }
      
      if (n < 0) {
        return FormulaError.NUM;
      }
      
      // 修正ベッセル関数I_n(x)の級数展開
      // I_n(x) = Σ(k=0 to ∞) [1 / (k! Γ(n+k+1))] * (x/2)^(n+2k)
      let result = 0;
      const maxTerms = 100;
      
      for (let k = 0; k < maxTerms; k++) {
        const term = 1 / (factorial(k) * gamma(n + k + 1)) * 
                     Math.pow(x / 2, n + 2 * k);
        
        if (Math.abs(term) < 1e-15 * Math.abs(result)) break;
        result += term;
      }
      
      return result;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// BESSELK関数の実装（修正ベッセル関数Kn(x)）
export const BESSELK: CustomFormula = {
  name: 'BESSELK',
  pattern: /BESSELK\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, xRef, nRef] = matches;
    
    try {
      const x = parseFloat(getCellValue(xRef.trim(), context)?.toString() ?? xRef.trim());
      const n = parseInt(getCellValue(nRef.trim(), context)?.toString() ?? nRef.trim());
      
      if (isNaN(x) || isNaN(n)) {
        return FormulaError.VALUE;
      }
      
      if (x <= 0 || n < 0) {
        return FormulaError.NUM;
      }
      
      // K_n(x)の計算（近似実装）
      if (n === 0) {
        // K_0(x)の近似
        if (x <= 2) {
          const i0 = besseli0Approx(x);
          return -Math.log(x / 2) * i0 + 0.42278420;
        } else {
          return Math.sqrt(Math.PI / (2 * x)) * Math.exp(-x) * 
                 (1 + 0.125 / x + 0.125 * 0.125 / (2 * x * x));
        }
      } else if (n === 1) {
        // K_1(x)の近似
        if (x <= 2) {
          const i1 = besseli1Approx(x);
          return Math.log(x / 2) * i1 + 1 / x * (1 + 0.5 * x * x / 4);
        } else {
          return Math.sqrt(Math.PI / (2 * x)) * Math.exp(-x) * 
                 (1 + 0.375 / x + 0.375 * 0.375 / (2 * x * x));
        }
      }
      
      // 再帰関係を使用
      // K_{n+1}(x) = (2n/x) * K_n(x) + K_{n-1}(x)
      let k0 = besselk0Approx(x);
      let k1 = besselk1Approx(x);
      let kn = k1;
      
      for (let i = 1; i < n; i++) {
        const temp = (2 * i / x) * k1 + k0;
        k0 = k1;
        k1 = temp;
        kn = temp;
      }
      
      return kn;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// 補助関数
function besselJ0(x: number): number {
  // J_0(x)の計算
  let result = 0;
  for (let k = 0; k < 50; k++) {
    const term = Math.pow(-1, k) / (factorial(k) * factorial(k)) * 
                 Math.pow(x / 2, 2 * k);
    if (Math.abs(term) < 1e-15) break;
    result += term;
  }
  return result;
}

function besselJ1(x: number): number {
  // J_1(x)の計算
  let result = 0;
  for (let k = 0; k < 50; k++) {
    const term = Math.pow(-1, k) / (factorial(k) * factorial(k + 1)) * 
                 Math.pow(x / 2, 2 * k + 1);
    if (Math.abs(term) < 1e-15) break;
    result += term;
  }
  return result;
}

function bessely0Approx(x: number): number {
  if (x < 3) {
    return (2 / Math.PI) * (Math.log(x / 2) * besselJ0(x) + 
           0.07832358 - 0.0420024 * x * x);
  } else {
    const sqrtx = Math.sqrt(2 / (Math.PI * x));
    return sqrtx * Math.sin(x - Math.PI / 4);
  }
}

function bessely1Approx(x: number): number {
  if (x < 3) {
    return (2 / Math.PI) * (Math.log(x / 2) * besselJ1(x) - 1 / x);
  } else {
    const sqrtx = Math.sqrt(2 / (Math.PI * x));
    return sqrtx * Math.sin(x - 3 * Math.PI / 4);
  }
}

function besseli0Approx(x: number): number {
  // I_0(x)の計算
  let result = 0;
  for (let k = 0; k < 50; k++) {
    const term = 1 / (factorial(k) * factorial(k)) * Math.pow(x / 2, 2 * k);
    if (Math.abs(term) < 1e-15 * Math.abs(result)) break;
    result += term;
  }
  return result;
}

function besseli1Approx(x: number): number {
  // I_1(x)の計算
  let result = 0;
  for (let k = 0; k < 50; k++) {
    const term = 1 / (factorial(k) * factorial(k + 1)) * Math.pow(x / 2, 2 * k + 1);
    if (Math.abs(term) < 1e-15 * Math.abs(result)) break;
    result += term;
  }
  return result;
}

function besselk0Approx(x: number): number {
  if (x <= 2) {
    const i0 = besseli0Approx(x);
    return -Math.log(x / 2) * i0 + 0.42278420;
  } else {
    return Math.sqrt(Math.PI / (2 * x)) * Math.exp(-x) * 
           (1 + 0.125 / x + 0.125 * 0.125 / (2 * x * x));
  }
}

function besselk1Approx(x: number): number {
  if (x <= 2) {
    const i1 = besseli1Approx(x);
    return Math.log(x / 2) * i1 + 1 / x * (1 + 0.5 * x * x / 4);
  } else {
    return Math.sqrt(Math.PI / (2 * x)) * Math.exp(-x) * 
           (1 + 0.375 / x + 0.375 * 0.375 / (2 * x * x));
  }
}