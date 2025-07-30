// その他の統計関数

import type { CustomFormula, FormulaContext, FormulaResult } from '../shared/types';
import { FormulaError } from '../shared/types';
import { getCellValue } from '../shared/utils';

// 標準正規分布の累積分布関数（簡易実装）
function normalCDF(x: number): number {
  // Abramowitz and Stegun approximation
  const a1 = 0.254829592;
  const a2 = -0.284496736;
  const a3 = 1.421413741;
  const a4 = -1.453152027;
  const a5 = 1.061405429;
  const p = 0.3275911;
  
  const sign = x >= 0 ? 1 : -1;
  x = Math.abs(x) / Math.sqrt(2);
  const t = 1.0 / (1.0 + p * x);
  const y = 1.0 - (((((a5 * t + a4) * t) + a3) * t + a2) * t + a1) * t * Math.exp(-x * x);
  
  return 0.5 * (1.0 + sign * y);
}

/* Commented out unused functions
// 標準正規分布の逆関数
function normInv(p: number): number {
  if (p <= 0 || p >= 1) return NaN;
  
  // Abramowitz and Stegun approximation
  const a1 = -3.969683028665376e+01;
  const a2 = 2.209460984245205e+02;
  const a3 = -2.759285104469687e+02;
  const a4 = 1.383577518672690e+02;
  const a5 = -3.066479806614716e+01;
  const a6 = 2.506628277459239e+00;
  
  const b1 = -5.447609879822406e+01;
  const b2 = 1.615858368580409e+02;
  const b3 = -1.556989798598866e+02;
  const b4 = 6.680131188771972e+01;
  const b5 = -1.328068155288572e+01;
  
  const c1 = -7.784894002430293e-03;
  const c2 = -3.223964580411365e-01;
  const c3 = -2.400758277161838e+00;
  const c4 = -2.549732539343734e+00;
  const c5 = 4.374664141464968e+00;
  const c6 = 2.938163982698783e+00;
  
  const d1 = 7.784695709041462e-03;
  const d2 = 3.224671290700398e-01;
  const d3 = 2.445134137142996e+00;
  const d4 = 3.754408661907416e+00;
  
  const p_low = 0.02425;
  const p_high = 1 - p_low;
  
  let x: number;
  
  if (p < p_low) {
    const q = Math.sqrt(-2 * Math.log(p));
    x = (((((c1 * q + c2) * q + c3) * q + c4) * q + c5) * q + c6) /
        ((((d1 * q + d2) * q + d3) * q + d4) * q + 1);
  } else if (p <= p_high) {
    const q = p - 0.5;
    const r = q * q;
    x = (((((a1 * r + a2) * r + a3) * r + a4) * r + a5) * r + a6) * q /
        (((((b1 * r + b2) * r + b3) * r + b4) * r + b5) * r + 1);
  } else {
    const q = Math.sqrt(-2 * Math.log(1 - p));
    x = -(((((c1 * q + c2) * q + c3) * q + c4) * q + c5) * q + c6) /
         ((((d1 * q + d2) * q + d3) * q + d4) * q + 1);
  }
  
  return x;
}

// T分布の逆関数
function tInv(p: number, df: number): number {
  if (p <= 0 || p >= 1 || df <= 0) return NaN;
  
  // 初期推定値として正規分布の逆関数を使用
  let x = normInv(p);
  
  // Newton-Raphson法で精度を上げる
  const maxIter = 100;
  const tol = 1e-8;
  
  for (let i = 0; i < maxIter; i++) {
    const cdf = tCDF(x, df);
    const pdf = tPDF(x, df);
    
    const error = cdf - p;
    if (Math.abs(error) < tol) break;
    
    x = x - error / pdf;
  }
  
  return x;
}

// T分布の累積分布関数
function tCDF(t: number, df: number): number {
  const x = df / (df + t * t);
  const a = 0.5;
  const b = df / 2;
  
  if (t < 0) {
    return 0.5 * betaIncomplete(b, a, x);
  } else {
    return 1 - 0.5 * betaIncomplete(b, a, x);
  }
}

// T分布の確率密度関数
function tPDF(t: number, df: number): number {
  const gamma1 = gamma((df + 1) / 2);
  const gamma2 = gamma(df / 2);
  
  return gamma1 / (Math.sqrt(Math.PI * df) * gamma2) * Math.pow(1 + t * t / df, -(df + 1) / 2);
}
*/

/* ガンマ関数
function gamma(x: number): number {
  if (x <= 0) return NaN;
  
  // Stirlingの近似
  if (x >= 10) {
    return Math.sqrt(2 * Math.PI / x) * Math.pow(x / Math.E, x) * 
           (1 + 1 / (12 * x) + 1 / (288 * x * x));
  }
  
  // 小さい値の場合は再帰的に計算
  if (x < 1) {
    return gamma(x + 1) / x;
  }
  
  // 1 <= x < 10 の場合はLanczosの近似を使用
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
  
  let z = x - 1;
  let base = z + g + 0.5;
  let sum = coef[0];
  
  for (let i = 1; i < 9; i++) {
    sum += coef[i] / (z + i);
  }
  
  return Math.sqrt(2 * Math.PI) * Math.pow(base, z + 0.5) * Math.exp(-base) * sum;
}
*/

/* 不完全ベータ関数
function betaIncomplete(a: number, b: number, x: number): number {
  if (x < 0 || x > 1) return NaN;
  if (x === 0) return 0;
  if (x === 1) return 1;
  
  // 連分数展開を使用
  const lbeta = gamma(a) * gamma(b) / gamma(a + b);
  const front = Math.exp(Math.log(x) * a + Math.log(1 - x) * b) / (a * lbeta);
  
  const d = 1;
  const c = 1;
  let f = 1;
  
  for (let i = 0; i < 200; i++) {
    const m = i / 2;
    let numerator: number;
    
    if (i === 0) {
      numerator = 1;
    } else if (i % 2 === 0) {
      numerator = (m * (b - m) * x) / ((a + 2 * m - 1) * (a + 2 * m));
    } else {
      numerator = -((a + m) * (a + b + m) * x) / ((a + 2 * m) * (a + 2 * m + 1));
    }
    
    const d_new = 1 + numerator * d;
    const c_new = 1 + numerator / c;
    
    const cd = c_new * d_new;
    f = f * cd;
    
    if (Math.abs(cd - 1) < 1e-10) break;
  }
  
  return front * (f - 1);
}
*/

// FISHER関数の実装（フィッシャー変換）
export const FISHER: CustomFormula = {
  name: 'FISHER',
  pattern: /FISHER\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, valueRef] = matches;
    
    try {
      const x = parseFloat(getCellValue(valueRef.trim(), context)?.toString() ?? valueRef.trim());
      
      if (isNaN(x)) {
        return FormulaError.VALUE;
      }
      
      if (x <= -1 || x >= 1) {
        return FormulaError.NUM;
      }
      
      // フィッシャー変換: F(x) = 0.5 * ln((1+x)/(1-x))
      return 0.5 * Math.log((1 + x) / (1 - x));
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// FISHERINV関数の実装（フィッシャー変換の逆関数）
export const FISHERINV: CustomFormula = {
  name: 'FISHERINV',
  pattern: /FISHERINV\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, valueRef] = matches;
    
    try {
      const y = parseFloat(getCellValue(valueRef.trim(), context)?.toString() ?? valueRef.trim());
      
      if (isNaN(y)) {
        return FormulaError.VALUE;
      }
      
      // フィッシャー逆変換: F^(-1)(y) = (e^(2y) - 1) / (e^(2y) + 1)
      const exp2y = Math.exp(2 * y);
      return (exp2y - 1) / (exp2y + 1);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// PHI関数の実装（標準正規分布の密度関数）
export const PHI: CustomFormula = {
  name: 'PHI',
  pattern: /PHI\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, xRef] = matches;
    
    try {
      const x = parseFloat(getCellValue(xRef.trim(), context)?.toString() ?? xRef.trim());
      
      if (isNaN(x)) {
        return FormulaError.VALUE;
      }
      
      // 標準正規分布の密度関数: φ(x) = (1/√(2π)) * e^(-x²/2)
      return Math.exp(-x * x / 2) / Math.sqrt(2 * Math.PI);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// GAUSS関数の実装（標準正規分布の累積分布関数から0.5を引いた値）
export const GAUSS: CustomFormula = {
  name: 'GAUSS',
  pattern: /GAUSS\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, zRef] = matches;
    
    try {
      const z = parseFloat(getCellValue(zRef.trim(), context)?.toString() ?? zRef.trim());
      
      if (isNaN(z)) {
        return FormulaError.VALUE;
      }
      
      // GAUSS(z) = N(z) - 0.5
      return normalCDF(z) - 0.5;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// PROB関数の実装（確率の計算）
export const PROB: CustomFormula = {
  name: 'PROB',
  pattern: /PROB\(([^,]+),\s*([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, xRangeRef, probRangeRef, lowerRef, upperRef] = matches;
    
    try {
      // X値と確率の配列を取得
      const xValues: number[] = [];
      const probValues: number[] = [];
      
      // 範囲参照の場合
      if (xRangeRef.includes(':') && probRangeRef.includes(':')) {
        const xRange = xRangeRef.trim().match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
        const probRange = probRangeRef.trim().match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
        
        if (!xRange || !probRange) {
          return FormulaError.REF;
        }
        
        const [, xStartCol, xStartRowStr, xEndCol, xEndRowStr] = xRange;
        const xStartRow = parseInt(xStartRowStr) - 1;
        const xEndRow = parseInt(xEndRowStr) - 1;
        const xStartColIndex = xStartCol.charCodeAt(0) - 65;
        const xEndColIndex = xEndCol.charCodeAt(0) - 65;
        
        // X値を収集
        for (let row = xStartRow; row <= xEndRow; row++) {
          for (let col = xStartColIndex; col <= xEndColIndex; col++) {
            const cell = context.data[row]?.[col];
            if (cell) {
              const val = parseFloat(cell.value?.toString() ?? '');
              if (!isNaN(val)) {
                xValues.push(val);
              }
            }
          }
        }
        
        // 確率値を収集
        const [, pStartCol, pStartRowStr, , pEndRowStr] = probRange;
        const pStartRow = parseInt(pStartRowStr) - 1;
        const pEndRow = parseInt(pEndRowStr) - 1;
        const pStartColIndex = pStartCol.charCodeAt(0) - 65;
        
        for (let row = pStartRow; row <= pEndRow; row++) {
          for (let col = pStartColIndex; col <= pStartColIndex + (xEndColIndex - xStartColIndex); col++) {
            const cell = context.data[row]?.[col];
            if (cell) {
              const val = parseFloat(cell.value?.toString() ?? '');
              if (!isNaN(val)) {
                probValues.push(val);
              }
            }
          }
        }
      }
      
      if (xValues.length === 0 || xValues.length !== probValues.length) {
        return FormulaError.VALUE;
      }
      
      // 確率の合計が1になることを確認
      const probSum = probValues.reduce((sum, p) => sum + p, 0);
      if (Math.abs(probSum - 1) > 0.0001) {
        return FormulaError.NUM;
      }
      
      const lower = parseFloat(getCellValue(lowerRef.trim(), context)?.toString() ?? lowerRef.trim());
      const upper = upperRef ? parseFloat(getCellValue(upperRef.trim(), context)?.toString() ?? upperRef.trim()) : lower;
      
      if (isNaN(lower) || isNaN(upper)) {
        return FormulaError.VALUE;
      }
      
      // 指定範囲内の確率を合計
      let totalProb = 0;
      for (let i = 0; i < xValues.length; i++) {
        if (xValues[i] >= lower && xValues[i] <= upper) {
          totalProb += probValues[i];
        }
      }
      
      return totalProb;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// BINOM.INV関数の実装（二項分布の逆関数）
export const BINOM_INV: CustomFormula = {
  name: 'BINOM.INV',
  pattern: /BINOM\.INV\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, trialsRef, probRef, alphaRef] = matches;
    
    try {
      const trialsValue = getCellValue(trialsRef.trim(), context)?.toString() ?? trialsRef.trim();
      const trialsNum = parseFloat(trialsValue);
      const prob = parseFloat(getCellValue(probRef.trim(), context)?.toString() ?? probRef.trim());
      const alpha = parseFloat(getCellValue(alphaRef.trim(), context)?.toString() ?? alphaRef.trim());
      
      if (isNaN(trialsNum) || isNaN(prob) || isNaN(alpha)) {
        return FormulaError.VALUE;
      }
      
      if (trialsNum < 0 || !Number.isInteger(trialsNum) || prob < 0 || prob > 1 || alpha < 0 || alpha > 1) {
        return FormulaError.NUM;
      }
      
      const trials = Math.floor(trialsNum);
      
      // 二項分布の累積分布関数を使って逆関数を計算
      let cumulativeProb = 0;
      
      for (let k = 0; k <= trials; k++) {
        // 二項係数の計算
        let binomCoef = 1;
        for (let i = 0; i < k; i++) {
          binomCoef *= (trials - i) / (i + 1);
        }
        
        // 確率の計算
        const probK = binomCoef * Math.pow(prob, k) * Math.pow(1 - prob, trials - k);
        cumulativeProb += probK;
        
        if (cumulativeProb >= alpha) {
          return k;
        }
      }
      
      return trials;
    } catch {
      return FormulaError.VALUE;
    }
  }
};
