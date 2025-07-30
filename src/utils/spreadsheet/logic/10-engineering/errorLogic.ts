// 誤差関数とその他のエンジニアリング関数

import type { CustomFormula, FormulaContext, FormulaResult } from '../shared/types';
import { FormulaError } from '../shared/types';
import { getCellValue } from '../shared/utils';

// ERF関数の実装（誤差関数）
export const ERF: CustomFormula = {
  name: 'ERF',
  pattern: /ERF\(([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, lowerRef, upperRef] = matches;
    
    try {
      const lower = parseFloat(getCellValue(lowerRef.trim(), context)?.toString() ?? lowerRef.trim());
      
      if (isNaN(lower)) {
        return FormulaError.VALUE;
      }
      
      if (upperRef) {
        const upper = parseFloat(getCellValue(upperRef.trim(), context)?.toString() ?? upperRef.trim());
        if (isNaN(upper)) {
          return FormulaError.VALUE;
        }
        // erf(upper) - erf(lower)
        return erfFunction(upper) - erfFunction(lower);
      } else {
        // erf(lower)
        return erfFunction(lower);
      }
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// ERF.PRECISE関数の実装（誤差関数・精密）
export const ERF_PRECISE: CustomFormula = {
  name: 'ERF.PRECISE',
  pattern: /ERF\.PRECISE\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, xRef] = matches;
    
    try {
      const x = parseFloat(getCellValue(xRef.trim(), context)?.toString() ?? xRef.trim());
      
      if (isNaN(x)) {
        return FormulaError.VALUE;
      }
      
      return erfFunction(x);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// ERFC関数の実装（相補誤差関数）
export const ERFC: CustomFormula = {
  name: 'ERFC',
  pattern: /ERFC\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, xRef] = matches;
    
    try {
      const x = parseFloat(getCellValue(xRef.trim(), context)?.toString() ?? xRef.trim());
      
      if (isNaN(x)) {
        return FormulaError.VALUE;
      }
      
      return erfcFunction(x);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// ERFC.PRECISE関数の実装（相補誤差関数・精密）
export const ERFC_PRECISE: CustomFormula = {
  name: 'ERFC.PRECISE',
  pattern: /ERFC\.PRECISE\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, xRef] = matches;
    
    try {
      const x = parseFloat(getCellValue(xRef.trim(), context)?.toString() ?? xRef.trim());
      
      if (isNaN(x)) {
        return FormulaError.VALUE;
      }
      
      return erfcFunction(x);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// DELTA関数の実装（クロネッカーのデルタ）
export const DELTA: CustomFormula = {
  name: 'DELTA',
  pattern: /DELTA\(([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, num1Ref, num2Ref] = matches;
    
    try {
      const num1 = parseFloat(getCellValue(num1Ref.trim(), context)?.toString() ?? num1Ref.trim());
      const num2 = num2Ref ? 
        parseFloat(getCellValue(num2Ref.trim(), context)?.toString() ?? num2Ref.trim()) : 0;
      
      if (isNaN(num1) || isNaN(num2)) {
        return FormulaError.VALUE;
      }
      
      // 数値が等しい場合は1、そうでない場合は0を返す
      return num1 === num2 ? 1 : 0;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// GESTEP関数の実装（ステップ関数）
export const GESTEP: CustomFormula = {
  name: 'GESTEP',
  pattern: /GESTEP\(([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, numRef, stepRef] = matches;
    
    try {
      const num = parseFloat(getCellValue(numRef.trim(), context)?.toString() ?? numRef.trim());
      const step = stepRef ? 
        parseFloat(getCellValue(stepRef.trim(), context)?.toString() ?? stepRef.trim()) : 0;
      
      if (isNaN(num) || isNaN(step)) {
        return FormulaError.VALUE;
      }
      
      // 数値がステップ以上の場合は1、そうでない場合は0を返す
      return num >= step ? 1 : 0;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// 誤差関数の計算
function erfFunction(x: number): number {
  // Special cases for exact values
  if (x === 0) return 0;
  if (x === Infinity) return 1;
  if (x === -Infinity) return -1;
  
  // For very large values
  if (Math.abs(x) > 6) {
    return x > 0 ? 1 : -1;
  }
  
  // Abramowitz and Stegun approximation
  const a1 = 0.254829592;
  const a2 = -0.284496736;
  const a3 = 1.421413741;
  const a4 = -1.453152027;
  const a5 = 1.061405429;
  const p = 0.3275911;
  
  const sign = x >= 0 ? 1 : -1;
  x = Math.abs(x);
  
  // For very small values, use series expansion: erf(x) ≈ 2x/√π
  if (x < 1e-8) {
    return sign * (2 * x / Math.sqrt(Math.PI));
  }
  
  const t = 1.0 / (1.0 + p * x);
  const y = 1.0 - (((((a5 * t + a4) * t) + a3) * t + a2) * t + a1) * t * Math.exp(-x * x);
  
  return sign * y;
}

// 相補誤差関数の計算
function erfcFunction(x: number): number {
  // Special cases for exact values
  if (x === 0) return 1;
  if (x === Infinity) return 0;
  if (x === -Infinity) return 2;
  
  // For very large positive values, use asymptotic expansion to avoid precision loss
  if (x > 8) {
    return Math.exp(-x * x) / (Math.sqrt(Math.PI) * x) * (1 - 1/(2 * x * x));
  }
  
  // For very large negative values
  if (x < -8) {
    return 2 - erfcFunction(-x);
  }
  
  return 1 - erfFunction(x);
}