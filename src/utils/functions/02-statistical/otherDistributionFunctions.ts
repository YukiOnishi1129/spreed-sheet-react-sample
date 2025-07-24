// その他の分布関数

import type { CustomFormula, FormulaContext, FormulaResult } from '../shared/types';
import { FormulaError } from '../shared/types';
import { getCellValue } from '../shared/utils';

// ガンマ関数
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

// 不完全ベータ関数
function betaIncomplete(a: number, b: number, x: number): number {
  if (x < 0 || x > 1) return NaN;
  if (x === 0) return 0;
  if (x === 1) return 1;
  
  // 連分数展開を使用
  const lbeta = Math.log(gamma(a)) + Math.log(gamma(b)) - Math.log(gamma(a + b));
  const front = Math.exp(Math.log(x) * a + Math.log(1 - x) * b - lbeta) / a;
  
  let c = 1;
  let d = 1;
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
    
    d = 1 + numerator * d;
    if (Math.abs(d) < 1e-30) d = 1e-30;
    d = 1 / d;
    
    c = 1 + numerator / c;
    if (Math.abs(c) < 1e-30) c = 1e-30;
    
    const delta = c * d;
    f *= delta;
    
    if (Math.abs(delta - 1) < 1e-10) break;
  }
  
  return front * (f - 1);
}

// LOGNORMDIST関数の実装（対数正規分布）- レガシー版
export const LOGNORMDIST: CustomFormula = {
  name: 'LOGNORMDIST',
  pattern: /LOGNORMDIST\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, xRef, meanRef, stdRef] = matches;
    
    try {
      const x = parseFloat(getCellValue(xRef.trim(), context)?.toString() ?? xRef.trim());
      const mean = parseFloat(getCellValue(meanRef.trim(), context)?.toString() ?? meanRef.trim());
      const std = parseFloat(getCellValue(stdRef.trim(), context)?.toString() ?? stdRef.trim());
      
      if (isNaN(x) || isNaN(mean) || isNaN(std)) {
        return FormulaError.VALUE;
      }
      
      if (x <= 0 || std <= 0) {
        return FormulaError.NUM;
      }
      
      // 対数正規分布の累積分布関数
      const z = (Math.log(x) - mean) / std;
      
      // 標準正規分布の累積分布関数を使用
      const a1 = 0.254829592;
      const a2 = -0.284496736;
      const a3 = 1.421413741;
      const a4 = -1.453152027;
      const a5 = 1.061405429;
      const p = 0.3275911;
      
      const sign = z >= 0 ? 1 : -1;
      const absZ = Math.abs(z) / Math.sqrt(2);
      const t = 1.0 / (1.0 + p * absZ);
      const y = 1.0 - (((((a5 * t + a4) * t) + a3) * t + a2) * t + a1) * t * Math.exp(-absZ * absZ);
      
      return 0.5 * (1.0 + sign * y);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// LOGINV関数の実装（対数正規分布の逆関数）- レガシー版
export const LOGINV: CustomFormula = {
  name: 'LOGINV',
  pattern: /LOGINV\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, probRef, meanRef, stdRef] = matches;
    
    try {
      const prob = parseFloat(getCellValue(probRef.trim(), context)?.toString() ?? probRef.trim());
      const mean = parseFloat(getCellValue(meanRef.trim(), context)?.toString() ?? meanRef.trim());
      const std = parseFloat(getCellValue(stdRef.trim(), context)?.toString() ?? stdRef.trim());
      
      if (isNaN(prob) || isNaN(mean) || isNaN(std)) {
        return FormulaError.VALUE;
      }
      
      if (prob <= 0 || prob >= 1 || std <= 0) {
        return FormulaError.NUM;
      }
      
      // 標準正規分布の逆関数を使用
      function normInv(p: number): number {
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
      
      const z = normInv(prob);
      return Math.exp(mean + std * z);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// EXPONDIST関数の実装（指数分布）- レガシー版
export const EXPONDIST: CustomFormula = {
  name: 'EXPONDIST',
  pattern: /EXPONDIST\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, xRef, lambdaRef, cumulativeRef] = matches;
    
    try {
      const x = parseFloat(getCellValue(xRef.trim(), context)?.toString() ?? xRef.trim());
      const lambda = parseFloat(getCellValue(lambdaRef.trim(), context)?.toString() ?? lambdaRef.trim());
      const cumulative = getCellValue(cumulativeRef.trim(), context)?.toString().toLowerCase() === 'true';
      
      if (isNaN(x) || isNaN(lambda)) {
        return FormulaError.VALUE;
      }
      
      if (x < 0 || lambda <= 0) {
        return FormulaError.NUM;
      }
      
      if (cumulative) {
        return 1 - Math.exp(-lambda * x);
      } else {
        return lambda * Math.exp(-lambda * x);
      }
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// WEIBULL関数の実装（ワイブル分布）- レガシー版
export const WEIBULL: CustomFormula = {
  name: 'WEIBULL',
  pattern: /WEIBULL\(([^,]+),\s*([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, xRef, alphaRef, betaRef, cumulativeRef] = matches;
    
    try {
      const x = parseFloat(getCellValue(xRef.trim(), context)?.toString() ?? xRef.trim());
      const alpha = parseFloat(getCellValue(alphaRef.trim(), context)?.toString() ?? alphaRef.trim());
      const beta = parseFloat(getCellValue(betaRef.trim(), context)?.toString() ?? betaRef.trim());
      const cumulative = getCellValue(cumulativeRef.trim(), context)?.toString().toLowerCase() === 'true';
      
      if (isNaN(x) || isNaN(alpha) || isNaN(beta)) {
        return FormulaError.VALUE;
      }
      
      if (x < 0 || alpha <= 0 || beta <= 0) {
        return FormulaError.NUM;
      }
      
      if (cumulative) {
        return 1 - Math.exp(-Math.pow(x / beta, alpha));
      } else {
        return (alpha / beta) * Math.pow(x / beta, alpha - 1) * Math.exp(-Math.pow(x / beta, alpha));
      }
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// GAMMADIST関数の実装（ガンマ分布）- レガシー版
export const GAMMADIST: CustomFormula = {
  name: 'GAMMADIST',
  pattern: /GAMMADIST\(([^,]+),\s*([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, xRef, alphaRef, betaRef, cumulativeRef] = matches;
    
    try {
      const x = parseFloat(getCellValue(xRef.trim(), context)?.toString() ?? xRef.trim());
      const alpha = parseFloat(getCellValue(alphaRef.trim(), context)?.toString() ?? alphaRef.trim());
      const beta = parseFloat(getCellValue(betaRef.trim(), context)?.toString() ?? betaRef.trim());
      const cumulative = getCellValue(cumulativeRef.trim(), context)?.toString().toLowerCase() === 'true';
      
      if (isNaN(x) || isNaN(alpha) || isNaN(beta)) {
        return FormulaError.VALUE;
      }
      
      if (x < 0 || alpha <= 0 || beta <= 0) {
        return FormulaError.NUM;
      }
      
      if (cumulative) {
        // 不完全ガンマ関数を使用した累積分布関数
        let sum = 0;
        const scaledX = x / beta;
        
        // 級数展開による近似
        for (let n = 0; n < 100; n++) {
          const term = Math.pow(scaledX, alpha + n) * Math.exp(-scaledX) / (gamma(alpha + n + 1));
          sum += term;
          if (Math.abs(term) < 1e-10) break;
        }
        
        return sum;
      } else {
        // 確率密度関数
        return Math.pow(x, alpha - 1) * Math.exp(-x / beta) / (Math.pow(beta, alpha) * gamma(alpha));
      }
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// GAMMAINV関数の実装（ガンマ分布の逆関数）- レガシー版
export const GAMMAINV: CustomFormula = {
  name: 'GAMMAINV',
  pattern: /GAMMAINV\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, probRef, alphaRef, betaRef] = matches;
    
    try {
      const prob = parseFloat(getCellValue(probRef.trim(), context)?.toString() ?? probRef.trim());
      const alpha = parseFloat(getCellValue(alphaRef.trim(), context)?.toString() ?? alphaRef.trim());
      const beta = parseFloat(getCellValue(betaRef.trim(), context)?.toString() ?? betaRef.trim());
      
      if (isNaN(prob) || isNaN(alpha) || isNaN(beta)) {
        return FormulaError.VALUE;
      }
      
      if (prob <= 0 || prob >= 1 || alpha <= 0 || beta <= 0) {
        return FormulaError.NUM;
      }
      
      // Newton-Raphson法で逆関数を計算
      let x = alpha * beta; // 初期値
      const maxIterations = 100;
      const tolerance = 1e-8;
      
      for (let i = 0; i < maxIterations; i++) {
        // 累積分布関数の値を計算
        let cdf = 0;
        const scaledX = x / beta;
        
        for (let n = 0; n < 50; n++) {
          const term = Math.pow(scaledX, alpha + n) * Math.exp(-scaledX) / gamma(alpha + n + 1);
          cdf += term;
          if (Math.abs(term) < 1e-10) break;
        }
        
        // 確率密度関数の値を計算
        const pdf = Math.pow(x, alpha - 1) * Math.exp(-x / beta) / (Math.pow(beta, alpha) * gamma(alpha));
        
        const error = cdf - prob;
        if (Math.abs(error) < tolerance) {
          return x;
        }
        
        x = x - error / pdf;
      }
      
      return x;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// BETADIST関数の実装（ベータ分布）- レガシー版
export const BETADIST: CustomFormula = {
  name: 'BETADIST',
  pattern: /BETADIST\(([^,]+),\s*([^,]+),\s*([^,)]+)(?:,\s*([^,)]+))?(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, xRef, alphaRef, betaRef, aRef, bRef] = matches;
    
    try {
      const x = parseFloat(getCellValue(xRef.trim(), context)?.toString() ?? xRef.trim());
      const alpha = parseFloat(getCellValue(alphaRef.trim(), context)?.toString() ?? alphaRef.trim());
      const beta = parseFloat(getCellValue(betaRef.trim(), context)?.toString() ?? betaRef.trim());
      const a = aRef ? parseFloat(getCellValue(aRef.trim(), context)?.toString() ?? aRef.trim()) : 0;
      const b = bRef ? parseFloat(getCellValue(bRef.trim(), context)?.toString() ?? bRef.trim()) : 1;
      
      if (isNaN(x) || isNaN(alpha) || isNaN(beta) || isNaN(a) || isNaN(b)) {
        return FormulaError.VALUE;
      }
      
      if (x < a || x > b || alpha <= 0 || beta <= 0 || a >= b) {
        return FormulaError.NUM;
      }
      
      // 標準化
      const standardizedX = (x - a) / (b - a);
      
      // 不完全ベータ関数を使用
      return betaIncomplete(alpha, beta, standardizedX);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// BETAINV関数の実装（ベータ分布の逆関数）- レガシー版
export const BETAINV: CustomFormula = {
  name: 'BETAINV',
  pattern: /BETAINV\(([^,]+),\s*([^,]+),\s*([^,)]+)(?:,\s*([^,)]+))?(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, probRef, alphaRef, betaRef, aRef, bRef] = matches;
    
    try {
      const prob = parseFloat(getCellValue(probRef.trim(), context)?.toString() ?? probRef.trim());
      const alpha = parseFloat(getCellValue(alphaRef.trim(), context)?.toString() ?? alphaRef.trim());
      const beta = parseFloat(getCellValue(betaRef.trim(), context)?.toString() ?? betaRef.trim());
      const a = aRef ? parseFloat(getCellValue(aRef.trim(), context)?.toString() ?? aRef.trim()) : 0;
      const b = bRef ? parseFloat(getCellValue(bRef.trim(), context)?.toString() ?? bRef.trim()) : 1;
      
      if (isNaN(prob) || isNaN(alpha) || isNaN(beta) || isNaN(a) || isNaN(b)) {
        return FormulaError.VALUE;
      }
      
      if (prob <= 0 || prob >= 1 || alpha <= 0 || beta <= 0 || a >= b) {
        return FormulaError.NUM;
      }
      
      // Newton-Raphson法で逆関数を計算
      let x = 0.5; // 初期値（0～1の範囲）
      const maxIterations = 100;
      const tolerance = 1e-8;
      
      for (let i = 0; i < maxIterations; i++) {
        const cdf = betaIncomplete(alpha, beta, x);
        const error = cdf - prob;
        
        if (Math.abs(error) < tolerance) {
          return a + x * (b - a);
        }
        
        // ベータ分布の確率密度関数
        const pdf = Math.pow(x, alpha - 1) * Math.pow(1 - x, beta - 1) / 
                    (gamma(alpha) * gamma(beta) / gamma(alpha + beta));
        
        x = x - error / pdf;
        
        // 境界チェック
        if (x < 0) x = 0.001;
        if (x > 1) x = 0.999;
      }
      
      return a + x * (b - a);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// CHIDIST関数の実装（カイ二乗分布の右側確率）- レガシー版
export const CHIDIST: CustomFormula = {
  name: 'CHIDIST',
  pattern: /CHIDIST\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, xRef, dfRef] = matches;
    
    try {
      const x = parseFloat(getCellValue(xRef.trim(), context)?.toString() ?? xRef.trim());
      const df = parseInt(getCellValue(dfRef.trim(), context)?.toString() ?? dfRef.trim());
      
      if (isNaN(x) || isNaN(df)) {
        return FormulaError.VALUE;
      }
      
      if (x < 0 || df < 1 || df > 10e10) {
        return FormulaError.NUM;
      }
      
      // カイ二乗分布は、ガンマ分布の特殊な場合
      // χ²(df) = Gamma(df/2, 2)
      const alpha = df / 2;
      const beta = 2;
      
      // 右側確率を計算（1 - 累積分布関数）
      let sum = 0;
      const scaledX = x / beta;
      
      for (let n = 0; n < 100; n++) {
        const term = Math.pow(scaledX, alpha + n) * Math.exp(-scaledX) / gamma(alpha + n + 1);
        sum += term;
        if (Math.abs(term) < 1e-10) break;
      }
      
      return 1 - sum;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// CHIINV関数の実装（カイ二乗分布の逆関数）- レガシー版
export const CHIINV: CustomFormula = {
  name: 'CHIINV',
  pattern: /CHIINV\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, probRef, dfRef] = matches;
    
    try {
      const prob = parseFloat(getCellValue(probRef.trim(), context)?.toString() ?? probRef.trim());
      const df = parseInt(getCellValue(dfRef.trim(), context)?.toString() ?? dfRef.trim());
      
      if (isNaN(prob) || isNaN(df)) {
        return FormulaError.VALUE;
      }
      
      if (prob <= 0 || prob >= 1 || df < 1) {
        return FormulaError.NUM;
      }
      
      // Newton-Raphson法で逆関数を計算
      let x = df; // 初期値
      const maxIterations = 100;
      const tolerance = 1e-8;
      
      for (let i = 0; i < maxIterations; i++) {
        // 右側確率を計算
        const alpha = df / 2;
        const beta = 2;
        let sum = 0;
        const scaledX = x / beta;
        
        for (let n = 0; n < 50; n++) {
          const term = Math.pow(scaledX, alpha + n) * Math.exp(-scaledX) / gamma(alpha + n + 1);
          sum += term;
          if (Math.abs(term) < 1e-10) break;
        }
        
        const rightTailProb = 1 - sum;
        const error = rightTailProb - prob;
        
        if (Math.abs(error) < tolerance) {
          return x;
        }
        
        // 確率密度関数
        const pdf = Math.pow(x, df/2 - 1) * Math.exp(-x/2) / (Math.pow(2, df/2) * gamma(df/2));
        
        x = x + error / pdf; // 右側確率なので符号を反転
      }
      
      return x;
    } catch {
      return FormulaError.VALUE;
    }
  }
};