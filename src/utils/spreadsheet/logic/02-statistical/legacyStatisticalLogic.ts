// レガシー統計関数（互換性のために残されている関数）

import type { CustomFormula, FormulaContext, FormulaResult } from '../shared/types';
import { FormulaError } from '../shared/types';
import { getCellValue, getCellRangeValues, toNumber } from '../shared/utils';

// 標準正規分布の累積分布関数
function normalCDF(x: number): number {
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

// 標準正規分布の逆関数
function normInv(p: number): number {
  if (p <= 0 || p >= 1) return NaN;
  
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

// NORMDIST関数の実装（正規分布）- レガシー版
export const NORMDIST: CustomFormula = {
  name: 'NORMDIST',
  pattern: /NORMDIST\(([^,]+),\s*([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, xRef, meanRef, stdRef, cumulativeRef] = matches;
    
    try {
      const x = parseFloat(getCellValue(xRef.trim(), context)?.toString() ?? xRef.trim());
      const mean = parseFloat(getCellValue(meanRef.trim(), context)?.toString() ?? meanRef.trim());
      const std = parseFloat(getCellValue(stdRef.trim(), context)?.toString() ?? stdRef.trim());
      const cumulative = getCellValue(cumulativeRef.trim(), context)?.toString().toLowerCase() === 'true';
      
      if (isNaN(x) || isNaN(mean) || isNaN(std)) {
        return FormulaError.VALUE;
      }
      
      if (std <= 0) {
        return FormulaError.NUM;
      }
      
      const z = (x - mean) / std;
      
      if (cumulative) {
        return normalCDF(z);
      } else {
        return Math.exp(-z * z / 2) / (std * Math.sqrt(2 * Math.PI));
      }
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// NORMINV関数の実装（正規分布の逆関数）- レガシー版
export const NORMINV: CustomFormula = {
  name: 'NORMINV',
  pattern: /NORMINV\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
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
      
      return mean + std * normInv(prob);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// NORMSDIST関数の実装（標準正規分布）- レガシー版
export const NORMSDIST: CustomFormula = {
  name: 'NORMSDIST',
  pattern: /NORMSDIST\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, zRef] = matches;
    
    try {
      const z = parseFloat(getCellValue(zRef.trim(), context)?.toString() ?? zRef.trim());
      
      if (isNaN(z)) {
        return FormulaError.VALUE;
      }
      
      return normalCDF(z);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// NORMSINV関数の実装（標準正規分布の逆関数）- レガシー版
export const NORMSINV: CustomFormula = {
  name: 'NORMSINV',
  pattern: /NORMSINV\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, probRef] = matches;
    
    try {
      const prob = parseFloat(getCellValue(probRef.trim(), context)?.toString() ?? probRef.trim());
      
      if (isNaN(prob)) {
        return FormulaError.VALUE;
      }
      
      if (prob <= 0 || prob >= 1) {
        return FormulaError.NUM;
      }
      
      return normInv(prob);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// BINOMDIST関数の実装（二項分布）- レガシー版
export const BINOMDIST: CustomFormula = {
  name: 'BINOMDIST',
  pattern: /BINOMDIST\(([^,]+),\s*([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, numSuccessRef, trialsRef, probRef, cumulativeRef] = matches;
    
    try {
      const numSuccess = parseInt(getCellValue(numSuccessRef.trim(), context)?.toString() ?? numSuccessRef.trim());
      const trials = parseInt(getCellValue(trialsRef.trim(), context)?.toString() ?? trialsRef.trim());
      const prob = parseFloat(getCellValue(probRef.trim(), context)?.toString() ?? probRef.trim());
      const cumulative = getCellValue(cumulativeRef.trim(), context)?.toString().toLowerCase() === 'true';
      
      if (isNaN(numSuccess) || isNaN(trials) || isNaN(prob)) {
        return FormulaError.VALUE;
      }
      
      if (numSuccess < 0 || numSuccess > trials || prob < 0 || prob > 1) {
        return FormulaError.NUM;
      }
      
      if (cumulative) {
        let cumulativeProb = 0;
        for (let k = 0; k <= numSuccess; k++) {
          let binomCoef = 1;
          for (let i = 0; i < k; i++) {
            binomCoef *= (trials - i) / (i + 1);
          }
          cumulativeProb += binomCoef * Math.pow(prob, k) * Math.pow(1 - prob, trials - k);
        }
        return cumulativeProb;
      } else {
        let binomCoef = 1;
        for (let i = 0; i < numSuccess; i++) {
          binomCoef *= (trials - i) / (i + 1);
        }
        return binomCoef * Math.pow(prob, numSuccess) * Math.pow(1 - prob, trials - numSuccess);
      }
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// POISSON関数の実装（ポアソン分布）- レガシー版
export const POISSON: CustomFormula = {
  name: 'POISSON',
  pattern: /POISSON\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, xRef, meanRef, cumulativeRef] = matches;
    
    try {
      const x = parseInt(getCellValue(xRef.trim(), context)?.toString() ?? xRef.trim());
      const mean = parseFloat(getCellValue(meanRef.trim(), context)?.toString() ?? meanRef.trim());
      const cumulative = getCellValue(cumulativeRef.trim(), context)?.toString().toLowerCase() === 'true';
      
      if (isNaN(x) || isNaN(mean)) {
        return FormulaError.VALUE;
      }
      
      if (x < 0 || mean <= 0) {
        return FormulaError.NUM;
      }
      
      if (cumulative) {
        let sum = 0;
        for (let k = 0; k <= x; k++) {
          let factorial = 1;
          for (let i = 1; i <= k; i++) {
            factorial *= i;
          }
          sum += Math.pow(mean, k) * Math.exp(-mean) / factorial;
        }
        return sum;
      } else {
        let factorial = 1;
        for (let i = 1; i <= x; i++) {
          factorial *= i;
        }
        return Math.pow(mean, x) * Math.exp(-mean) / factorial;
      }
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// CRITBINOM関数の実装（二項分布の臨界値）
export const CRITBINOM: CustomFormula = {
  name: 'CRITBINOM',
  pattern: /CRITBINOM\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, trialsRef, probRef, alphaRef] = matches;
    
    try {
      const trials = parseInt(getCellValue(trialsRef.trim(), context)?.toString() ?? trialsRef.trim());
      const prob = parseFloat(getCellValue(probRef.trim(), context)?.toString() ?? probRef.trim());
      const alpha = parseFloat(getCellValue(alphaRef.trim(), context)?.toString() ?? alphaRef.trim());
      
      if (isNaN(trials) || isNaN(prob) || isNaN(alpha)) {
        return FormulaError.VALUE;
      }
      
      if (trials < 0 || prob < 0 || prob > 1 || alpha < 0 || alpha > 1) {
        return FormulaError.NUM;
      }
      
      // 二項分布の累積分布関数を使って臨界値を計算
      let cumulativeProb = 0;
      
      for (let k = 0; k <= trials; k++) {
        let binomCoef = 1;
        for (let i = 0; i < k; i++) {
          binomCoef *= (trials - i) / (i + 1);
        }
        
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

// CONFIDENCE関数の実装（信頼区間）- レガシー版
export const CONFIDENCE: CustomFormula = {
  name: 'CONFIDENCE',
  pattern: /CONFIDENCE\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, alphaRef, stdRef, sizeRef] = matches;
    
    try {
      const alpha = parseFloat(getCellValue(alphaRef.trim(), context)?.toString() ?? alphaRef.trim());
      const std = parseFloat(getCellValue(stdRef.trim(), context)?.toString() ?? stdRef.trim());
      const size = parseInt(getCellValue(sizeRef.trim(), context)?.toString() ?? sizeRef.trim());
      
      if (isNaN(alpha) || isNaN(std) || isNaN(size)) {
        return FormulaError.VALUE;
      }
      
      if (alpha <= 0 || alpha >= 1 || std <= 0 || size < 1) {
        return FormulaError.NUM;
      }
      
      // 正規分布の逆関数を使用
      const z = normInv(1 - alpha / 2);
      
      return z * std / Math.sqrt(size);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// ZTEST関数の実装（Z検定）- レガシー版
export const ZTEST: CustomFormula = {
  name: 'ZTEST',
  pattern: /ZTEST\(([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, arrayRef, xRef, sigmaRef] = matches;
    
    try {
      const x = parseFloat(getCellValue(xRef.trim(), context)?.toString() ?? xRef.trim());
      
      if (isNaN(x)) {
        return FormulaError.VALUE;
      }
      
      // 配列からデータを取得
      const values: number[] = [];
      
      if (arrayRef.includes(':')) {
        const rangeMatch = arrayRef.trim().match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
        if (!rangeMatch) {
          return FormulaError.REF;
        }
        
        const [, startCol, startRowStr, endCol, endRowStr] = rangeMatch;
        const startRow = parseInt(startRowStr) - 1;
        const endRow = parseInt(endRowStr) - 1;
        const startColIndex = startCol.charCodeAt(0) - 65;
        const endColIndex = endCol.charCodeAt(0) - 65;
        
        for (let row = startRow; row <= endRow; row++) {
          for (let col = startColIndex; col <= endColIndex; col++) {
            const cell = context.data[row]?.[col];
            if (cell && !isNaN(Number(cell.value))) {
              values.push(Number(cell.value));
            }
          }
        }
      }
      
      if (values.length === 0) {
        return FormulaError.NA;
      }
      
      // 平均と標準偏差を計算
      const mean = values.reduce((sum, val) => sum + val, 0) / values.length;
      
      let sigma: number;
      if (sigmaRef) {
        sigma = parseFloat(getCellValue(sigmaRef.trim(), context)?.toString() ?? sigmaRef.trim());
        if (isNaN(sigma) || sigma <= 0) {
          return FormulaError.NUM;
        }
      } else {
        // 標本標準偏差を計算
        const variance = values.reduce((sum, val) => sum + Math.pow(val - mean, 2), 0) / (values.length - 1);
        sigma = Math.sqrt(variance);
      }
      
      // Z統計量を計算
      const z = (x - mean) / (sigma / Math.sqrt(values.length));
      
      // 片側検定のp値を返す
      return 1 - normalCDF(z);
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
      
      if (alpha <= 0 || beta <= 0 || a >= b || x < a || x > b) {
        return FormulaError.NUM;
      }
      
      // 標準化
      const y = (x - a) / (b - a);
      
      // ベータ関数の計算（ガンマ関数を使用）
      const lnBeta = (p: number, q: number): number => {
        const lnGamma = (z: number): number => {
          const cof = [76.18009172947146, -86.5053203294168, 24.01409824083091,
                       -1.231739572450155, 0.001208650973866179, -0.000005395239384953];
          const x = z;
          let y = z;
          let tmp = x + 5.5;
          tmp -= (x + 0.5) * Math.log(tmp);
          let ser = 1.000000000190015;
          for (let j = 0; j < 6; j++) {
            ser += cof[j] / ++y;
          }
          return -tmp + Math.log(2.506628274631 * ser / x);
        };
        
        return lnGamma(p) + lnGamma(q) - lnGamma(p + q);
      };
      
      // 不完全ベータ関数の計算（連分数展開）
      const betaInc = (x: number, a: number, b: number): number => {
        if (x === 0 || x === 1) return x;
        
        const lnBetaVal = lnBeta(a, b);
        const factor = Math.exp(a * Math.log(x) + b * Math.log(1 - x) - lnBetaVal);
        
        if (x < (a + 1) / (a + b + 2)) {
          // 連分数展開
          const maxIter = 100;
          const eps = 1e-10;
          let c = 1;
          let d = 1 / (1 - (a + b) * x / (a + 1));
          let h = d;
          
          for (let m = 1; m <= maxIter; m++) {
            const m2 = 2 * m;
            let aa = m * (b - m) * x / ((a + m2 - 1) * (a + m2));
            d = 1 / (1 + aa * d);
            c = 1 + aa / c;
            h *= d * c;
            
            aa = -(a + m) * (a + b + m) * x / ((a + m2) * (a + m2 + 1));
            d = 1 / (1 + aa * d);
            c = 1 + aa / c;
            const del = d * c;
            h *= del;
            
            if (Math.abs(del - 1) < eps) break;
          }
          
          return factor * h / a;
        } else {
          // 対称性を利用
          return 1 - betaInc(1 - x, b, a);
        }
      };
      
      return betaInc(y, alpha, beta);
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
      
      if (prob < 0 || prob > 1 || alpha <= 0 || beta <= 0 || a >= b) {
        return FormulaError.NUM;
      }
      
      // ニュートン法で逆関数を計算
      let x = 0.5; // 初期値
      const maxIter = 100;
      const tol = 1e-10;
      
      for (let iter = 0; iter < maxIter; iter++) {
        // BETADIST相当の計算
        const betaDist = BETADIST.calculate(['', x.toString(), alpha.toString(), beta.toString(), '0', '1'], context) as number;
        const error = betaDist - prob;
        
        if (Math.abs(error) < tol) {
          return a + x * (b - a);
        }
        
        // ベータ分布の確率密度関数
        const pdf = Math.pow(x, alpha - 1) * Math.pow(1 - x, beta - 1) * 
                    Math.exp(-lnBeta(alpha, beta));
        
        x -= error / pdf;
        x = Math.max(0.0001, Math.min(0.9999, x));
      }
      
      return FormulaError.NUM;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// ベータ関数の対数を計算するヘルパー関数
function lnBeta(a: number, b: number): number {
  const lnGamma = (z: number): number => {
    const cof = [76.18009172947146, -86.5053203294168, 24.01409824083091,
                 -1.231739572450155, 0.1208650973866179e-2, -0.5395239384953e-5];
    const x = z;
    let y = z;
    let tmp = x + 5.5;
    tmp -= (x + 0.5) * Math.log(tmp);
    let ser = 1.000000000190015;
    for (let j = 0; j < 6; j++) {
      ser += cof[j] / ++y;
    }
    return -tmp + Math.log(2.506628274631 * ser / x);
  };
  
  return lnGamma(a) + lnGamma(b) - lnGamma(a + b);
}

// CHIDIST関数の実装（カイ二乗分布）- レガシー版
export const CHIDIST: CustomFormula = {
  name: 'CHIDIST',
  pattern: /CHIDIST\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, xRef, degFreedomRef] = matches;
    
    try {
      const x = parseFloat(getCellValue(xRef.trim(), context)?.toString() ?? xRef.trim());
      const degFreedom = parseInt(getCellValue(degFreedomRef.trim(), context)?.toString() ?? degFreedomRef.trim());
      
      if (isNaN(x) || isNaN(degFreedom)) {
        return FormulaError.VALUE;
      }
      
      if (x < 0 || degFreedom < 1 || degFreedom > 10e10) {
        return FormulaError.NUM;
      }
      
      // ガンマ関数を使ったカイ二乗分布の計算
      const gammaInc = (a: number, x: number): number => {
        if (x < 0 || a <= 0) return NaN;
        
        if (x < a + 1) {
          // 級数展開
          let sum = 1 / a;
          let term = 1 / a;
          const maxIter = 100;
          
          for (let n = 1; n < maxIter; n++) {
            term *= x / (a + n);
            sum += term;
            if (Math.abs(term) < Math.abs(sum) * 1e-10) break;
          }
          
          return sum * Math.exp(-x + a * Math.log(x) - lnGamma(a));
        } else {
          // 連分数展開
          const maxIter = 100;
          const eps = 1e-10;
          let b = x + 1 - a;
          let c = 1 / 1e-30;
          let d = 1 / b;
          let h = d;
          
          for (let i = 1; i <= maxIter; i++) {
            const an = -i * (i - a);
            b += 2;
            d = an * d + b;
            if (Math.abs(d) < 1e-30) d = 1e-30;
            c = b + an / c;
            if (Math.abs(c) < 1e-30) c = 1e-30;
            d = 1 / d;
            const del = d * c;
            h *= del;
            if (Math.abs(del - 1) < eps) break;
          }
          
          return 1 - Math.exp(-x + a * Math.log(x) - lnGamma(a)) * h;
        }
      };
      
      // カイ二乗分布の右側確率
      return 1 - gammaInc(degFreedom / 2, x / 2);
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
    const [, probRef, degFreedomRef] = matches;
    
    try {
      const prob = parseFloat(getCellValue(probRef.trim(), context)?.toString() ?? probRef.trim());
      const degFreedom = parseInt(getCellValue(degFreedomRef.trim(), context)?.toString() ?? degFreedomRef.trim());
      
      if (isNaN(prob) || isNaN(degFreedom)) {
        return FormulaError.VALUE;
      }
      
      if (prob < 0 || prob > 1 || degFreedom < 1 || degFreedom > 10e10) {
        return FormulaError.NUM;
      }
      
      // ニュートン法で逆関数を計算
      let x = degFreedom; // 初期値
      const maxIter = 100;
      const tol = 1e-10;
      
      for (let iter = 0; iter < maxIter; iter++) {
        const chiDist = CHIDIST.calculate(['', x.toString(), degFreedom.toString()], context) as number;
        const error = chiDist - prob;
        
        if (Math.abs(error) < tol) {
          return x;
        }
        
        // カイ二乗分布の確率密度関数
        const k = degFreedom / 2;
        const pdf = Math.pow(x, k - 1) * Math.exp(-x / 2) / (Math.pow(2, k) * Math.exp(lnGamma(k)));
        
        x += error / pdf;
        x = Math.max(0.0001, x);
      }
      
      return FormulaError.NUM;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// ガンマ関数の対数を計算するヘルパー関数
function lnGamma(z: number): number {
  const cof = [76.18009172947146, -86.5053203294168, 24.01409824083091,
               -1.231739572450155, 0.001208650973866179, -0.000005395239384953];
  const x = z;
  let y = z;
  let tmp = x + 5.5;
  tmp -= (x + 0.5) * Math.log(tmp);
  let ser = 1.000000000190015;
  for (let j = 0; j < 6; j++) {
    ser += cof[j] / ++y;
  }
  return -tmp + Math.log(2.506628274631 * ser / x);
}

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
        // 累積分布関数
        const gammaInc = (a: number, x: number): number => {
          if (x < 0 || a <= 0) return NaN;
          
          if (x < a + 1) {
            // 級数展開
            let sum = 1 / a;
            let term = 1 / a;
            const maxIter = 100;
            
            for (let n = 1; n < maxIter; n++) {
              term *= x / (a + n);
              sum += term;
              if (Math.abs(term) < Math.abs(sum) * 1e-10) break;
            }
            
            return sum * Math.exp(-x + a * Math.log(x) - lnGamma(a));
          } else {
            // 連分数展開
            const maxIter = 100;
            const eps = 1e-10;
            let b = x + 1 - a;
            let c = 1 / 1e-30;
            let d = 1 / b;
            let h = d;
            
            for (let i = 1; i <= maxIter; i++) {
              const an = -i * (i - a);
              b += 2;
              d = an * d + b;
              if (Math.abs(d) < 1e-30) d = 1e-30;
              c = b + an / c;
              if (Math.abs(c) < 1e-30) c = 1e-30;
              d = 1 / d;
              const del = d * c;
              h *= del;
              if (Math.abs(del - 1) < eps) break;
            }
            
            return 1 - Math.exp(-x + a * Math.log(x) - lnGamma(a)) * h;
          }
        };
        
        return gammaInc(alpha, x / beta);
      } else {
        // 確率密度関数
        return Math.pow(x, alpha - 1) * Math.exp(-x / beta) / (Math.pow(beta, alpha) * Math.exp(lnGamma(alpha)));
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
      
      if (prob < 0 || prob > 1 || alpha <= 0 || beta <= 0) {
        return FormulaError.NUM;
      }
      
      // ニュートン法で逆関数を計算
      let x = alpha * beta; // 初期値（平均値）
      const maxIter = 100;
      const tol = 1e-10;
      
      for (let iter = 0; iter < maxIter; iter++) {
        const gammaDist = GAMMADIST.calculate(['', x.toString(), alpha.toString(), beta.toString(), 'TRUE'], context) as number;
        const error = gammaDist - prob;
        
        if (Math.abs(error) < tol) {
          return x;
        }
        
        // ガンマ分布の確率密度関数
        const pdf = Math.pow(x, alpha - 1) * Math.exp(-x / beta) / (Math.pow(beta, alpha) * Math.exp(lnGamma(alpha)));
        
        x -= error / pdf;
        x = Math.max(0.0001, x);
      }
      
      return FormulaError.NUM;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// HYPGEOMDIST関数の実装（超幾何分布）- レガシー版
export const HYPGEOMDIST: CustomFormula = {
  name: 'HYPGEOMDIST',
  pattern: /HYPGEOMDIST\(([^,]+),\s*([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, sampleSRef, numberSampleRef, populationSRef, numberPopRef] = matches;
    
    try {
      const sampleS = parseInt(getCellValue(sampleSRef.trim(), context)?.toString() ?? sampleSRef.trim());
      const numberSample = parseInt(getCellValue(numberSampleRef.trim(), context)?.toString() ?? numberSampleRef.trim());
      const populationS = parseInt(getCellValue(populationSRef.trim(), context)?.toString() ?? populationSRef.trim());
      const numberPop = parseInt(getCellValue(numberPopRef.trim(), context)?.toString() ?? numberPopRef.trim());
      
      if (isNaN(sampleS) || isNaN(numberSample) || isNaN(populationS) || isNaN(numberPop)) {
        return FormulaError.VALUE;
      }
      
      if (sampleS < 0 || sampleS > numberSample || sampleS > populationS ||
          numberSample <= 0 || numberSample > numberPop ||
          populationS <= 0 || populationS > numberPop ||
          numberPop <= 0) {
        return FormulaError.NUM;
      }
      
      // 組み合わせ数を計算
      const combination = (n: number, k: number): number => {
        if (k > n || k < 0) return 0;
        if (k === 0 || k === n) return 1;
        
        k = Math.min(k, n - k);
        let result = 1;
        
        for (let i = 0; i < k; i++) {
          result *= (n - i) / (i + 1);
        }
        
        return result;
      };
      
      // 超幾何分布の確率質量関数
      const numerator = combination(populationS, sampleS) * combination(numberPop - populationS, numberSample - sampleS);
      const denominator = combination(numberPop, numberSample);
      
      return numerator / denominator;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// PERCENTRANK.INC関数の実装（パーセントランク - 含む）
export const PERCENTRANK_INC: CustomFormula = {
  name: 'PERCENTRANK.INC',
  pattern: /PERCENTRANK\.INC\(([^,)]+)(?:,\s*([^,)]+))?(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, arrayRef, xRef, significanceRef] = matches;
    
    try {
      // 配列を取得
      const values = getCellRangeValues(arrayRef.trim(), context);
      
      // xの値を取得
      const xValue = xRef ? getCellValue(xRef.trim(), context) : null;
      const x = toNumber(xValue) ?? (xRef ? parseFloat(xRef.trim()) : NaN);
      
      if (isNaN(x)) {
        return FormulaError.VALUE;
      }
      
      // 有効桁数（デフォルトは3）
      const significanceValue = significanceRef ? getCellValue(significanceRef.trim(), context) : 3;
      const significance = toNumber(significanceValue) ?? (significanceRef ? parseInt(significanceRef.trim()) : 3);
      
      if (significance < 1) {
        return FormulaError.NUM;
      }
      
      // 数値のみを抽出してソート
      const numericValues = values
        .map(v => toNumber(v))
        .filter(n => n !== null) as number[];
      
      if (numericValues.length === 0) {
        return FormulaError.NUM;
      }
      
      numericValues.sort((a, b) => a - b);
      
      const n = numericValues.length;
      const min = numericValues[0];
      const max = numericValues[n - 1];
      
      // xが範囲外の場合
      if (x < min || x > max) {
        return FormulaError.NA;
      }
      
      // xが配列内にある場合
      const exactIndex = numericValues.indexOf(x);
      if (exactIndex !== -1) {
        // 同じ値が複数ある場合は平均を取る
        let count = 1;
        let sumRank = exactIndex;
        
        // 前方向を確認
        let i = exactIndex - 1;
        while (i >= 0 && numericValues[i] === x) {
          sumRank += i;
          count++;
          i--;
        }
        
        // 後方向を確認
        i = exactIndex + 1;
        while (i < n && numericValues[i] === x) {
          sumRank += i;
          count++;
          i++;
        }
        
        const avgRank = sumRank / count;
        const percentRank = avgRank / (n - 1);
        
        // 有効桁数で丸める
        const factor = Math.pow(10, significance);
        return Math.round(percentRank * factor) / factor;
      }
      
      // xが配列内にない場合は補間
      let lowerIndex = -1;
      let upperIndex = -1;
      
      for (let i = 0; i < n - 1; i++) {
        if (numericValues[i] < x && x < numericValues[i + 1]) {
          lowerIndex = i;
          upperIndex = i + 1;
          break;
        }
      }
      
      if (lowerIndex === -1) {
        return FormulaError.NA;
      }
      
      // 線形補間
      const lowerValue = numericValues[lowerIndex];
      const upperValue = numericValues[upperIndex];
      const fraction = (x - lowerValue) / (upperValue - lowerValue);
      
      const lowerRank = lowerIndex / (n - 1);
      const upperRank = upperIndex / (n - 1);
      const percentRank = lowerRank + fraction * (upperRank - lowerRank);
      
      // 有効桁数で丸める
      const factor = Math.pow(10, significance);
      return Math.round(percentRank * factor) / factor;
    } catch {
      return FormulaError.VALUE;
    }
  }
};