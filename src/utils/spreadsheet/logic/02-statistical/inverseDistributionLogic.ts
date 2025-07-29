// 逆分布関数の実装

import type { CustomFormula, FormulaContext, FormulaResult } from '../shared/types';
import { FormulaError } from '../shared/types';
import { getCellValue } from '../shared/utils';
import { inverseTDistribution as inverseTDist } from './distributionLogic';
// 標準正規分布の逆関数（近似）
// function normSInv(p: number): number {
//   if (p <= 0 || p >= 1) return NaN;
//   
//   // Abramowitz and Stegun approximation
//   const a1 = -3.969683028665376e+01;
//   const a2 = 2.209460984245205e+02;
//   const a3 = -2.759285104469687e+02;
//   const a4 = 1.383577518672690e+02;
//   const a5 = -3.066479806614716e+01;
//   const a6 = 2.506628277459239e+00;
//   
//   const b1 = -5.447609879822406e+01;
//   const b2 = 1.615858368580409e+02;
//   const b3 = -1.556989798598866e+02;
//   const b4 = 6.680131188771972e+01;
//   const b5 = -1.328068155288572e+01;
//   
//   const c1 = -7.784894002430293e-03;
//   const c2 = -3.223964580411365e-01;
//   const c3 = -2.400758277161838e+00;
//   const c4 = -2.549732539343734e+00;
//   const c5 = 4.374664141464968e+00;
//   const c6 = 2.938163982698783e+00;
//   
//   const d1 = 7.784695709041462e-03;
//   const d2 = 3.224671290700398e-01;
//   const d3 = 2.445134137142996e+00;
//   const d4 = 3.754408661907416e+00;
//   
//   const p_low = 0.02425;
//   const p_high = 1 - p_low;
//   
//   let x: number;
//   
//   if (p < p_low) {
//     const q = Math.sqrt(-2 * Math.log(p));
//     x = (((((c1 * q + c2) * q + c3) * q + c4) * q + c5) * q + c6) /
//         ((((d1 * q + d2) * q + d3) * q + d4) * q + 1);
//   } else if (p <= p_high) {
//     const q = p - 0.5;
//     const r = q * q;
//     x = (((((a1 * r + a2) * r + a3) * r + a4) * r + a5) * r + a6) * q /
//         (((((b1 * r + b2) * r + b3) * r + b4) * r + b5) * r + 1);
//   } else {
//     const q = Math.sqrt(-2 * Math.log(1 - p));
//     x = -(((((c1 * q + c2) * q + c3) * q + c4) * q + c5) * q + c6) /
//          ((((d1 * q + d2) * q + d3) * q + d4) * q + 1);
//   }
//   
//   return x;
// }

// ガンマ関数の対数（高精度計算用）
function logGamma(x: number): number {
  if (x <= 0) return NaN;
  
  // Stirlingの近似を使用
  if (x >= 10) {
    const series = 1 / (12 * x) - 1 / (360 * x * x * x) + 1 / (1260 * x * x * x * x * x);
    return (x - 0.5) * Math.log(x) - x + 0.5 * Math.log(2 * Math.PI) + series;
  }
  
  // 小さな値の場合は再帰的に計算
  return logGamma(x + 1) - Math.log(x);
}

// 正規化された不完全ベータ関数
function betaIncomplete(a: number, b: number, x: number): number {
  if (x < 0 || x > 1) return NaN;
  if (x === 0) return 0;
  if (x === 1) return 1;
  
  const lbeta = logGamma(a) + logGamma(b) - logGamma(a + b);
  const front = Math.exp(Math.log(x) * a + Math.log(1 - x) * b - lbeta) / a;
  
  // 連分数展開
  const TINY = 1e-30;
  let f = 1, c = 1, d = 0;
  
  for (let i = 0; i <= 200; i++) {
    const m = i / 2;
    
    let numerator;
    if (i === 0) {
      numerator = 1;
    } else if (i % 2 === 0) {
      numerator = (m * (b - m) * x) / ((a + 2 * m - 1) * (a + 2 * m));
    } else {
      numerator = -((a + m) * (a + b + m) * x) / ((a + 2 * m) * (a + 2 * m + 1));
    }
    
    d = 1 + numerator * d;
    if (Math.abs(d) < TINY) d = TINY;
    d = 1 / d;
    
    c = 1 + numerator / c;
    if (Math.abs(c) < TINY) c = TINY;
    
    f *= c * d;
    
    if (Math.abs(1 - c * d) < 1e-8) break;
  }
  
  return front * (f - 1);
}

// T分布の逆関数（Newton-Raphson法） - 未使用のため削除
// function tInv(p: number, df: number): number {
//   if (p <= 0 || p >= 1) return NaN;
//   if (df <= 0) return NaN;
//   
//   // 初期推定値（正規分布の逆関数を使用）
//   let x = normSInv(p);
//   
//   // Newton-Raphson法で精度を上げる
//   const maxIter = 100;
//   const tol = 1e-8;
//   
//   for (let i = 0; i < maxIter; i++) {
//     // T分布のCDF
//     const t = x;
//     const a = 0.5;
//     const b = df / 2;
//     const x2 = df / (df + t * t);
//     
//     let cdf;
//     if (t < 0) {
//       cdf = 0.5 * betaIncomplete(b, a, x2);
//     } else {
//       cdf = 1 - 0.5 * betaIncomplete(b, a, x2);
//     }
//     
//     // T分布のPDF
//     const pdf = Math.exp(logGamma((df + 1) / 2) - logGamma(df / 2)) / 
//                 Math.sqrt(Math.PI * df) * Math.pow(1 + t * t / df, -(df + 1) / 2);
//     
//     // Newton-Raphson更新
//     const error = cdf - p;
//     if (Math.abs(error) < tol) break;
//     
//     x = x - error / pdf;
//   }
//   
//   return x;
// }

// カイ二乗分布の逆関数
function chiSquareInv(p: number, df: number): number {
  if (p <= 0 || p >= 1) return NaN;
  if (df <= 0) return NaN;
  
  // 初期推定値
  let x = df;
  
  // Newton-Raphson法
  const maxIter = 100;
  const tol = 1e-8;
  
  for (let i = 0; i < maxIter; i++) {
    // ガンマ分布のCDF（カイ二乗分布はガンマ分布の特殊ケース）
    const cdf = gammaIncomplete(df / 2, x / 2);
    
    // ガンマ分布のPDF
    const pdf = Math.pow(x / 2, df / 2 - 1) * Math.exp(-x / 2) / 
                (2 * Math.exp(logGamma(df / 2)));
    
    const error = cdf - p;
    if (Math.abs(error) < tol) break;
    
    x = x - error / pdf;
    
    // 負の値を防ぐ
    if (x < 0) x = 0.01;
  }
  
  return x;
}

// 不完全ガンマ関数
function gammaIncomplete(a: number, x: number): number {
  if (x < 0) return 0;
  if (x === 0) return 0;
  
  if (x < a + 1) {
    // 級数展開
    let sum = 1 / a;
    let term = 1 / a;
    
    for (let i = 1; i < 100; i++) {
      term *= x / (a + i);
      sum += term;
      if (Math.abs(term) < Math.abs(sum) * 1e-10) break;
    }
    
    return sum * Math.exp(-x + a * Math.log(x) - logGamma(a));
  } else {
    // 連分数展開
    let b = x + 1 - a;
    let c = 1 / 1e-30;
    let d = 1 / b;
    let h = d;
    
    for (let i = 1; i < 100; i++) {
      const an = -i * (i - a);
      b += 2;
      d = an * d + b;
      if (Math.abs(d) < 1e-30) d = 1e-30;
      c = b + an / c;
      if (Math.abs(c) < 1e-30) c = 1e-30;
      d = 1 / d;
      const del = d * c;
      h *= del;
      if (Math.abs(del - 1) < 1e-10) break;
    }
    
    return 1 - Math.exp(-x + a * Math.log(x) - logGamma(a)) * h;
  }
}

// F分布の逆関数
function fInv(p: number, df1: number, df2: number): number {
  if (p <= 0 || p >= 1) return NaN;
  if (df1 <= 0 || df2 <= 0) return NaN;
  
  // Newton-Raphson法とビセクション法の組み合わせ
  const tol = 1e-8;
  const maxIter = 100;
  
  // 初期推定値の改善
  // F分布の平均値付近から開始
  let x = df2 > 2 ? df2 / (df2 - 2) : 1;
  
  // 適切な初期範囲を見つける
  let low = 0;
  let high = x * 10;
  
  // F分布のCDFを計算する関数
  const fCDF = (f: number): number => {
    const w = (df1 * f) / (df1 * f + df2);
    return betaIncomplete(df1 / 2, df2 / 2, w);
  };
  
  // 上限を探す
  while (fCDF(high) < p && high < 1e6) {
    low = high;
    high *= 2;
  }
  
  // 下限を探す
  low = 0;
  while (fCDF(low) < p * 0.01 && low < high) {
    low = high / 100;
    if (fCDF(low) > p) {
      high = low;
      low = 0;
      break;
    }
  }
  
  // Newton-Raphson法（高速収束）
  for (let iter = 0; iter < maxIter; iter++) {
    const cdf = fCDF(x);
    
    if (Math.abs(cdf - p) < tol) {
      return x;
    }
    
    // F分布のPDFを計算
    const logBeta = logGamma(df1 / 2) + logGamma(df2 / 2) - logGamma((df1 + df2) / 2);
    const pdf = Math.exp(
      Math.log(Math.pow(df1 / df2, df1 / 2)) +
      ((df1 / 2) - 1) * Math.log(x) -
      ((df1 + df2) / 2) * Math.log(1 + (df1 / df2) * x) -
      logBeta
    );
    
    // Newton-Raphson更新
    const delta = (cdf - p) / pdf;
    const newX = x - delta;
    
    // 範囲外の場合はビセクション法にフォールバック
    if (newX <= low || newX >= high || isNaN(newX)) {
      // ビセクション法の1ステップ
      const mid = (low + high) / 2;
      const midCDF = fCDF(mid);
      
      if (midCDF < p) {
        low = mid;
      } else {
        high = mid;
      }
      x = (low + high) / 2;
    } else {
      x = newX;
      
      // 範囲を更新
      if (cdf < p) {
        low = Math.max(low, x - Math.abs(delta) * 2);
      } else {
        high = Math.min(high, x + Math.abs(delta) * 2);
      }
    }
  }
  
  return x;
}

// T.INV関数の実装（T分布の逆関数）
export const T_INV: CustomFormula = {
  name: 'T.INV',
  pattern: /T\.INV\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, probRef, dfRef] = matches;
    
    try {
      const probability = parseFloat(getCellValue(probRef.trim(), context)?.toString() ?? probRef.trim());
      const degreesOfFreedom = parseFloat(getCellValue(dfRef.trim(), context)?.toString() ?? dfRef.trim());
      
      if (isNaN(probability) || isNaN(degreesOfFreedom)) {
        return FormulaError.VALUE;
      }
      
      if (probability <= 0 || probability >= 1) {
        return FormulaError.NUM;
      }
      
      if (degreesOfFreedom < 1) {
        return FormulaError.NUM;
      }
      
      return inverseTDist(probability, degreesOfFreedom);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// T.INV.2T関数の実装（両側T分布の逆関数）
export const T_INV_2T: CustomFormula = {
  name: 'T.INV.2T',
  pattern: /T\.INV\.2T\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, probRef, dfRef] = matches;
    
    try {
      const probability = parseFloat(getCellValue(probRef.trim(), context)?.toString() ?? probRef.trim());
      const degreesOfFreedom = parseFloat(getCellValue(dfRef.trim(), context)?.toString() ?? dfRef.trim());
      
      if (isNaN(probability) || isNaN(degreesOfFreedom)) {
        return FormulaError.VALUE;
      }
      
      if (probability <= 0 || probability >= 1) {
        return FormulaError.NUM;
      }
      
      if (degreesOfFreedom < 1) {
        return FormulaError.NUM;
      }
      
      // 両側検定なので、片側の確率は半分
      return Math.abs(inverseTDist(1 - probability / 2, degreesOfFreedom));
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// CHISQ.INV関数の実装（カイ二乗分布の逆関数）
export const CHISQ_INV: CustomFormula = {
  name: 'CHISQ.INV',
  pattern: /CHISQ\.INV\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, probRef, dfRef] = matches;
    
    try {
      const probability = parseFloat(getCellValue(probRef.trim(), context)?.toString() ?? probRef.trim());
      const degreesOfFreedom = parseFloat(getCellValue(dfRef.trim(), context)?.toString() ?? dfRef.trim());
      
      if (isNaN(probability) || isNaN(degreesOfFreedom)) {
        return FormulaError.VALUE;
      }
      
      if (probability < 0 || probability >= 1) {
        return FormulaError.NUM;
      }
      
      if (degreesOfFreedom < 1) {
        return FormulaError.NUM;
      }
      
      return chiSquareInv(probability, degreesOfFreedom);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// CHISQ.INV.RT関数の実装（右側カイ二乗分布の逆関数）
export const CHISQ_INV_RT: CustomFormula = {
  name: 'CHISQ.INV.RT',
  pattern: /CHISQ\.INV\.RT\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, probRef, dfRef] = matches;
    
    try {
      const probability = parseFloat(getCellValue(probRef.trim(), context)?.toString() ?? probRef.trim());
      const degreesOfFreedom = parseFloat(getCellValue(dfRef.trim(), context)?.toString() ?? dfRef.trim());
      
      if (isNaN(probability) || isNaN(degreesOfFreedom)) {
        return FormulaError.VALUE;
      }
      
      if (probability <= 0 || probability > 1) {
        return FormulaError.NUM;
      }
      
      if (degreesOfFreedom < 1) {
        return FormulaError.NUM;
      }
      
      // 右側確率なので、1から引く
      return chiSquareInv(1 - probability, degreesOfFreedom);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// F.INV関数の実装（F分布の逆関数）
export const F_INV: CustomFormula = {
  name: 'F.INV',
  pattern: /F\.INV\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, probRef, df1Ref, df2Ref] = matches;
    
    try {
      const probability = parseFloat(getCellValue(probRef.trim(), context)?.toString() ?? probRef.trim());
      const df1 = parseFloat(getCellValue(df1Ref.trim(), context)?.toString() ?? df1Ref.trim());
      const df2 = parseFloat(getCellValue(df2Ref.trim(), context)?.toString() ?? df2Ref.trim());
      
      if (isNaN(probability) || isNaN(df1) || isNaN(df2)) {
        return FormulaError.VALUE;
      }
      
      if (probability <= 0 || probability >= 1) {
        return FormulaError.NUM;
      }
      
      if (df1 < 1 || df2 < 1) {
        return FormulaError.NUM;
      }
      
      return fInv(probability, df1, df2);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// F.INV.RT関数の実装（右側F分布の逆関数）
export const F_INV_RT: CustomFormula = {
  name: 'F.INV.RT',
  pattern: /F\.INV\.RT\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, probRef, df1Ref, df2Ref] = matches;
    
    try {
      const probability = parseFloat(getCellValue(probRef.trim(), context)?.toString() ?? probRef.trim());
      const df1 = parseFloat(getCellValue(df1Ref.trim(), context)?.toString() ?? df1Ref.trim());
      const df2 = parseFloat(getCellValue(df2Ref.trim(), context)?.toString() ?? df2Ref.trim());
      
      if (isNaN(probability) || isNaN(df1) || isNaN(df2)) {
        return FormulaError.VALUE;
      }
      
      if (probability <= 0 || probability > 1) {
        return FormulaError.NUM;
      }
      
      if (df1 < 1 || df2 < 1) {
        return FormulaError.NUM;
      }
      
      // 右側確率なので、1から引く
      return fInv(1 - probability, df1, df2);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// BETA.INV関数の実装（ベータ分布の逆関数）
export const BETA_INV: CustomFormula = {
  name: 'BETA.INV',
  pattern: /BETA\.INV\(([^,]+),\s*([^,]+),\s*([^,]+)(?:,\s*([^,]+))?(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, probRef, alphaRef, betaRef, aRef, bRef] = matches;
    
    try {
      const probability = parseFloat(getCellValue(probRef.trim(), context)?.toString() ?? probRef.trim());
      const alpha = parseFloat(getCellValue(alphaRef.trim(), context)?.toString() ?? alphaRef.trim());
      const beta = parseFloat(getCellValue(betaRef.trim(), context)?.toString() ?? betaRef.trim());
      const a = aRef ? parseFloat(getCellValue(aRef.trim(), context)?.toString() ?? aRef.trim()) : 0;
      const b = bRef ? parseFloat(getCellValue(bRef.trim(), context)?.toString() ?? bRef.trim()) : 1;
      
      if (isNaN(probability) || isNaN(alpha) || isNaN(beta) || isNaN(a) || isNaN(b)) {
        return FormulaError.VALUE;
      }
      
      if (probability <= 0 || probability >= 1) {
        return FormulaError.NUM;
      }
      
      if (alpha <= 0 || beta <= 0) {
        return FormulaError.NUM;
      }
      
      if (a >= b) {
        return FormulaError.NUM;
      }
      
      // Newton-Raphson法でベータ分布の逆関数を計算
      let x = 0.5; // 初期推定値
      const maxIter = 100;
      const tol = 1e-8;
      
      for (let i = 0; i < maxIter; i++) {
        const cdf = betaIncomplete(alpha, beta, x);
        const pdf = Math.pow(x, alpha - 1) * Math.pow(1 - x, beta - 1) * 
                    Math.exp(logGamma(alpha + beta) - logGamma(alpha) - logGamma(beta));
        
        const error = cdf - probability;
        if (Math.abs(error) < tol) break;
        
        x = x - error / pdf;
        
        // 範囲を制限
        if (x < 0) x = 0.001;
        if (x > 1) x = 0.999;
      }
      
      // スケーリング
      return a + (b - a) * x;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// GAMMA.INV関数の実装（ガンマ分布の逆関数）
export const GAMMA_INV: CustomFormula = {
  name: 'GAMMA.INV',
  pattern: /GAMMA\.INV\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, probRef, alphaRef, betaRef] = matches;
    
    try {
      const probability = parseFloat(getCellValue(probRef.trim(), context)?.toString() ?? probRef.trim());
      const alpha = parseFloat(getCellValue(alphaRef.trim(), context)?.toString() ?? alphaRef.trim());
      const beta = parseFloat(getCellValue(betaRef.trim(), context)?.toString() ?? betaRef.trim());
      
      if (isNaN(probability) || isNaN(alpha) || isNaN(beta)) {
        return FormulaError.VALUE;
      }
      
      if (probability < 0 || probability >= 1) {
        return FormulaError.NUM;
      }
      
      if (alpha <= 0 || beta <= 0) {
        return FormulaError.NUM;
      }
      
      // カイ二乗分布の逆関数を使用（ガンマ分布との関係から）
      return beta * chiSquareInv(probability, 2 * alpha) / 2;
    } catch {
      return FormulaError.VALUE;
    }
  }
};