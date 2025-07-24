// レガシー統計関数（互換性のために残されている関数）

import type { CustomFormula, FormulaContext, FormulaResult } from '../shared/types';
import { FormulaError } from '../shared/types';
import { getCellValue } from '../shared/utils';

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