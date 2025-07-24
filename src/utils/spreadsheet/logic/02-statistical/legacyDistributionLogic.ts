// レガシー分布関数（F分布、T分布など）

import type { CustomFormula, FormulaContext, FormulaResult } from '../shared/types';
import { FormulaError } from '../shared/types';
import { getCellValue } from '../shared/utils';

// ガンマ関数
function gamma(x: number): number {
  if (x <= 0) return NaN;
  
  if (x >= 10) {
    return Math.sqrt(2 * Math.PI / x) * Math.pow(x / Math.E, x) * 
           (1 + 1 / (12 * x) + 1 / (288 * x * x));
  }
  
  if (x < 1) {
    return gamma(x + 1) / x;
  }
  
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
  
  const z = x - 1;
  const base = z + g + 0.5;
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

// FDIST関数の実装（F分布の右側確率）- レガシー版
export const FDIST: CustomFormula = {
  name: 'FDIST',
  pattern: /FDIST\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, xRef, df1Ref, df2Ref] = matches;
    
    try {
      const x = parseFloat(getCellValue(xRef.trim(), context)?.toString() ?? xRef.trim());
      const df1 = parseInt(getCellValue(df1Ref.trim(), context)?.toString() ?? df1Ref.trim());
      const df2 = parseInt(getCellValue(df2Ref.trim(), context)?.toString() ?? df2Ref.trim());
      
      if (isNaN(x) || isNaN(df1) || isNaN(df2)) {
        return FormulaError.VALUE;
      }
      
      if (x < 0 || df1 < 1 || df2 < 1) {
        return FormulaError.NUM;
      }
      
      // F分布の累積分布関数はベータ分布を使用して計算
      const a = df1 / 2;
      const b = df2 / 2;
      const y = df1 * x / (df1 * x + df2);
      
      // 右側確率 = 1 - CDF
      return 1 - betaIncomplete(a, b, y);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// FINV関数の実装（F分布の逆関数）- レガシー版
export const FINV: CustomFormula = {
  name: 'FINV',
  pattern: /FINV\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, probRef, df1Ref, df2Ref] = matches;
    
    try {
      const prob = parseFloat(getCellValue(probRef.trim(), context)?.toString() ?? probRef.trim());
      const df1 = parseInt(getCellValue(df1Ref.trim(), context)?.toString() ?? df1Ref.trim());
      const df2 = parseInt(getCellValue(df2Ref.trim(), context)?.toString() ?? df2Ref.trim());
      
      if (isNaN(prob) || isNaN(df1) || isNaN(df2)) {
        return FormulaError.VALUE;
      }
      
      if (prob <= 0 || prob >= 1 || df1 < 1 || df2 < 1) {
        return FormulaError.NUM;
      }
      
      // Newton-Raphson法で逆関数を計算
      let x = df2 / df1; // 初期値
      const maxIterations = 100;
      const tolerance = 1e-8;
      
      for (let i = 0; i < maxIterations; i++) {
        const a = df1 / 2;
        const b = df2 / 2;
        const y = df1 * x / (df1 * x + df2);
        
        const rightTailProb = 1 - betaIncomplete(a, b, y);
        const error = rightTailProb - prob;
        
        if (Math.abs(error) < tolerance) {
          return x;
        }
        
        // F分布の確率密度関数
        const pdf = Math.sqrt(Math.pow(df1 * x, df1) * Math.pow(df2, df2) / 
                             Math.pow(df1 * x + df2, df1 + df2)) / 
                    (x * gamma(df1/2) * gamma(df2/2) / gamma((df1+df2)/2));
        
        x = x + error / pdf;
      }
      
      return x;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// TDIST関数の実装（t分布）- レガシー版
export const TDIST: CustomFormula = {
  name: 'TDIST',
  pattern: /TDIST\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, xRef, dfRef, tailsRef] = matches;
    
    try {
      const x = parseFloat(getCellValue(xRef.trim(), context)?.toString() ?? xRef.trim());
      const df = parseInt(getCellValue(dfRef.trim(), context)?.toString() ?? dfRef.trim());
      const tails = parseInt(getCellValue(tailsRef.trim(), context)?.toString() ?? tailsRef.trim());
      
      if (isNaN(x) || isNaN(df) || isNaN(tails)) {
        return FormulaError.VALUE;
      }
      
      if (x < 0 || df < 1 || (tails !== 1 && tails !== 2)) {
        return FormulaError.NUM;
      }
      
      // t分布の累積分布関数
      const a = df / 2;
      const b = 0.5;
      const y = df / (df + x * x);
      
      const prob = betaIncomplete(a, b, y);
      
      if (tails === 1) {
        // 片側検定
        return 0.5 * prob;
      } else {
        // 両側検定
        return prob;
      }
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// TINV関数の実装（t分布の逆関数）- レガシー版
export const TINV: CustomFormula = {
  name: 'TINV',
  pattern: /TINV\(([^,]+),\s*([^)]+)\)/i,
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
      
      // 両側検定のt値を計算
      // Newton-Raphson法
      let t = 0; // 初期値
      const maxIterations = 100;
      const tolerance = 1e-8;
      
      for (let i = 0; i < maxIterations; i++) {
        const a = df / 2;
        const b = 0.5;
        const y = df / (df + t * t);
        
        const cdf = betaIncomplete(a, b, y);
        const error = cdf - prob;
        
        if (Math.abs(error) < tolerance) {
          return Math.abs(t);
        }
        
        // t分布の確率密度関数
        const pdf = gamma((df + 1) / 2) / (Math.sqrt(Math.PI * df) * gamma(df / 2)) * 
                    Math.pow(1 + t * t / df, -(df + 1) / 2);
        
        t = t - error / (2 * pdf); // 両側なので2で割る
      }
      
      return Math.abs(t);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// FTEST関数の実装（F検定）- レガシー版
export const FTEST: CustomFormula = {
  name: 'FTEST',
  pattern: /FTEST\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, array1Ref, array2Ref] = matches;
    
    try {
      // 配列1のデータを取得
      const values1: number[] = [];
      const range1Match = array1Ref.trim().match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
      
      if (range1Match) {
        const [, startCol, startRowStr, endCol, endRowStr] = range1Match;
        const startRow = parseInt(startRowStr) - 1;
        const endRow = parseInt(endRowStr) - 1;
        const startColIndex = startCol.charCodeAt(0) - 65;
        const endColIndex = endCol.charCodeAt(0) - 65;
        
        for (let row = startRow; row <= endRow; row++) {
          for (let col = startColIndex; col <= endColIndex; col++) {
            const cell = context.data[row]?.[col];
            if (cell && !isNaN(Number(cell.value))) {
              values1.push(Number(cell.value));
            }
          }
        }
      }
      
      // 配列2のデータを取得
      const values2: number[] = [];
      const range2Match = array2Ref.trim().match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
      
      if (range2Match) {
        const [, startCol, startRowStr, endCol, endRowStr] = range2Match;
        const startRow = parseInt(startRowStr) - 1;
        const endRow = parseInt(endRowStr) - 1;
        const startColIndex = startCol.charCodeAt(0) - 65;
        const endColIndex = endCol.charCodeAt(0) - 65;
        
        for (let row = startRow; row <= endRow; row++) {
          for (let col = startColIndex; col <= endColIndex; col++) {
            const cell = context.data[row]?.[col];
            if (cell && !isNaN(Number(cell.value))) {
              values2.push(Number(cell.value));
            }
          }
        }
      }
      
      if (values1.length < 2 || values2.length < 2) {
        return FormulaError.DIV0;
      }
      
      // 分散を計算
      const mean1 = values1.reduce((sum, val) => sum + val, 0) / values1.length;
      const mean2 = values2.reduce((sum, val) => sum + val, 0) / values2.length;
      
      const var1 = values1.reduce((sum, val) => sum + Math.pow(val - mean1, 2), 0) / (values1.length - 1);
      const var2 = values2.reduce((sum, val) => sum + Math.pow(val - mean2, 2), 0) / (values2.length - 1);
      
      // F統計量
      const f = var1 > var2 ? var1 / var2 : var2 / var1;
      const df1 = var1 > var2 ? values1.length - 1 : values2.length - 1;
      const df2 = var1 > var2 ? values2.length - 1 : values1.length - 1;
      
      // p値（両側検定）
      const a = df1 / 2;
      const b = df2 / 2;
      const y = df1 * f / (df1 * f + df2);
      
      return 2 * (1 - betaIncomplete(a, b, y));
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// TTEST関数の実装（t検定）- レガシー版
export const TTEST: CustomFormula = {
  name: 'TTEST',
  pattern: /TTEST\(([^,]+),\s*([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, array1Ref, array2Ref, tailsRef, typeRef] = matches;
    
    try {
      const tails = parseInt(getCellValue(tailsRef.trim(), context)?.toString() ?? tailsRef.trim());
      const type = parseInt(getCellValue(typeRef.trim(), context)?.toString() ?? typeRef.trim());
      
      if (isNaN(tails) || isNaN(type)) {
        return FormulaError.VALUE;
      }
      
      if ((tails !== 1 && tails !== 2) || (type < 1 || type > 3)) {
        return FormulaError.NUM;
      }
      
      // 配列1のデータを取得
      const values1: number[] = [];
      const range1Match = array1Ref.trim().match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
      
      if (range1Match) {
        const [, startCol, startRowStr, endCol, endRowStr] = range1Match;
        const startRow = parseInt(startRowStr) - 1;
        const endRow = parseInt(endRowStr) - 1;
        const startColIndex = startCol.charCodeAt(0) - 65;
        const endColIndex = endCol.charCodeAt(0) - 65;
        
        for (let row = startRow; row <= endRow; row++) {
          for (let col = startColIndex; col <= endColIndex; col++) {
            const cell = context.data[row]?.[col];
            if (cell && !isNaN(Number(cell.value))) {
              values1.push(Number(cell.value));
            }
          }
        }
      }
      
      // 配列2のデータを取得
      const values2: number[] = [];
      const range2Match = array2Ref.trim().match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
      
      if (range2Match) {
        const [, startCol, startRowStr, endCol, endRowStr] = range2Match;
        const startRow = parseInt(startRowStr) - 1;
        const endRow = parseInt(endRowStr) - 1;
        const startColIndex = startCol.charCodeAt(0) - 65;
        const endColIndex = endCol.charCodeAt(0) - 65;
        
        for (let row = startRow; row <= endRow; row++) {
          for (let col = startColIndex; col <= endColIndex; col++) {
            const cell = context.data[row]?.[col];
            if (cell && !isNaN(Number(cell.value))) {
              values2.push(Number(cell.value));
            }
          }
        }
      }
      
      if (values1.length === 0 || values2.length === 0) {
        return FormulaError.NA;
      }
      
      // 平均と分散を計算
      const mean1 = values1.reduce((sum, val) => sum + val, 0) / values1.length;
      const mean2 = values2.reduce((sum, val) => sum + val, 0) / values2.length;
      
      let t: number;
      let df: number;
      
      if (type === 1) {
        // 対応のあるt検定
        if (values1.length !== values2.length) {
          return FormulaError.NA;
        }
        
        const diffs = values1.map((val, i) => val - values2[i]);
        const meanDiff = diffs.reduce((sum, val) => sum + val, 0) / diffs.length;
        const varDiff = diffs.reduce((sum, val) => sum + Math.pow(val - meanDiff, 2), 0) / (diffs.length - 1);
        
        t = meanDiff / Math.sqrt(varDiff / diffs.length);
        df = diffs.length - 1;
      } else {
        const var1 = values1.reduce((sum, val) => sum + Math.pow(val - mean1, 2), 0) / (values1.length - 1);
        const var2 = values2.reduce((sum, val) => sum + Math.pow(val - mean2, 2), 0) / (values2.length - 1);
        
        if (type === 2) {
          // 等分散を仮定した2標本t検定
          const pooledVar = ((values1.length - 1) * var1 + (values2.length - 1) * var2) / 
                           (values1.length + values2.length - 2);
          t = (mean1 - mean2) / Math.sqrt(pooledVar * (1/values1.length + 1/values2.length));
          df = values1.length + values2.length - 2;
        } else {
          // 等分散を仮定しない2標本t検定（Welchのt検定）
          t = (mean1 - mean2) / Math.sqrt(var1/values1.length + var2/values2.length);
          const v1 = var1 / values1.length;
          const v2 = var2 / values2.length;
          df = Math.pow(v1 + v2, 2) / (Math.pow(v1, 2)/(values1.length - 1) + Math.pow(v2, 2)/(values2.length - 1));
        }
      }
      
      // p値を計算
      const a = df / 2;
      const b = 0.5;
      const y = df / (df + t * t);
      const prob = betaIncomplete(a, b, y);
      
      return tails === 1 ? 0.5 * prob : prob;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// CHITEST関数の実装（カイ二乗検定）- レガシー版
export const CHITEST: CustomFormula = {
  name: 'CHITEST',
  pattern: /CHITEST\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, actualRef, expectedRef] = matches;
    
    try {
      // 実測値の配列を取得
      const actualValues: number[][] = [];
      const actualRange = actualRef.trim().match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
      
      if (!actualRange) {
        return FormulaError.REF;
      }
      
      const [, startCol, startRowStr, endCol, endRowStr] = actualRange;
      const startRow = parseInt(startRowStr) - 1;
      const endRow = parseInt(endRowStr) - 1;
      const startColIndex = startCol.charCodeAt(0) - 65;
      const endColIndex = endCol.charCodeAt(0) - 65;
      
      for (let row = startRow; row <= endRow; row++) {
        const rowData: number[] = [];
        for (let col = startColIndex; col <= endColIndex; col++) {
          const cell = context.data[row]?.[col];
          if (cell && !isNaN(Number(cell.value))) {
            rowData.push(Number(cell.value));
          } else {
            rowData.push(0);
          }
        }
        actualValues.push(rowData);
      }
      
      // 期待値の配列を取得
      const expectedValues: number[][] = [];
      const expectedRange = expectedRef.trim().match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
      
      if (!expectedRange) {
        return FormulaError.REF;
      }
      
      const [, expStartCol, expStartRowStr, expEndCol, expEndRowStr] = expectedRange;
      const expStartRow = parseInt(expStartRowStr) - 1;
      const expEndRow = parseInt(expEndRowStr) - 1;
      const expStartColIndex = expStartCol.charCodeAt(0) - 65;
      const expEndColIndex = expEndCol.charCodeAt(0) - 65;
      
      for (let row = expStartRow; row <= expEndRow; row++) {
        const rowData: number[] = [];
        for (let col = expStartColIndex; col <= expEndColIndex; col++) {
          const cell = context.data[row]?.[col];
          if (cell && !isNaN(Number(cell.value))) {
            rowData.push(Number(cell.value));
          } else {
            rowData.push(0);
          }
        }
        expectedValues.push(rowData);
      }
      
      // サイズチェック
      if (actualValues.length !== expectedValues.length || 
          actualValues[0].length !== expectedValues[0].length) {
        return FormulaError.NA;
      }
      
      // カイ二乗統計量を計算
      let chiSquare = 0;
      let df = 0;
      
      for (let i = 0; i < actualValues.length; i++) {
        for (let j = 0; j < actualValues[i].length; j++) {
          const observed = actualValues[i][j];
          const expected = expectedValues[i][j];
          
          if (expected > 0) {
            chiSquare += Math.pow(observed - expected, 2) / expected;
            df++;
          }
        }
      }
      
      df = Math.max(1, df - 1);
      
      // p値を計算（右側確率）
      let sum = 0;
      const alpha = df / 2;
      const beta = 2;
      const scaledX = chiSquare / beta;
      
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

// NEGBINOMDIST関数の実装（負の二項分布）- レガシー版
export const NEGBINOMDIST: CustomFormula = {
  name: 'NEGBINOMDIST',
  pattern: /NEGBINOMDIST\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, numFailuresRef, numSuccessRef, probRef] = matches;
    
    try {
      const numFailures = parseInt(getCellValue(numFailuresRef.trim(), context)?.toString() ?? numFailuresRef.trim());
      const numSuccess = parseInt(getCellValue(numSuccessRef.trim(), context)?.toString() ?? numSuccessRef.trim());
      const prob = parseFloat(getCellValue(probRef.trim(), context)?.toString() ?? probRef.trim());
      
      if (isNaN(numFailures) || isNaN(numSuccess) || isNaN(prob)) {
        return FormulaError.VALUE;
      }
      
      if (numFailures < 0 || numSuccess < 1 || prob <= 0 || prob >= 1) {
        return FormulaError.NUM;
      }
      
      // 負の二項分布の確率質量関数
      // P(X = k) = C(k+r-1, k) * p^r * (1-p)^k
      // ここで、k = numFailures, r = numSuccess, p = prob
      
      // 組み合わせ C(n,k) = n! / (k! * (n-k)!)
      let binomCoef = 1;
      for (let i = 0; i < numFailures; i++) {
        binomCoef *= (numFailures + numSuccess - 1 - i) / (i + 1);
      }
      
      return binomCoef * Math.pow(prob, numSuccess) * Math.pow(1 - prob, numFailures);
    } catch {
      return FormulaError.VALUE;
    }
  }
};