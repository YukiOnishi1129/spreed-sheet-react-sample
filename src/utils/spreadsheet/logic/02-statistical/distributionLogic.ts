// 分布関数の実装

import type { CustomFormula, FormulaContext } from '../shared/types';
import { FormulaError } from '../shared/types';
import { getCellValue, getCellRangeValues } from '../shared/utils';

// 数学定数とヘルパー関数
const SQRT_2PI = Math.sqrt(2 * Math.PI);
const SQRT_2 = Math.sqrt(2);

// 範囲を2次元配列として解析
function parseRange(range: string, context: FormulaContext): number[][] | null {
  const match = range.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
  if (!match) return null;
  
  const startCol = match[1].charCodeAt(0) - 65;
  const startRow = parseInt(match[2]) - 1;
  const endCol = match[3].charCodeAt(0) - 65;
  const endRow = parseInt(match[4]) - 1;
  
  const result: number[][] = [];
  for (let r = startRow; r <= endRow; r++) {
    const row: number[] = [];
    for (let c = startCol; c <= endCol; c++) {
      const value = context.data[r]?.[c]?.value ?? context.data[r]?.[c];
      row.push(Number(value));
    }
    result.push(row);
  }
  
  return result;
}

// カイ二乗分布の累積分布関数
function chiSquareCDF(x: number, df: number): number {
  if (x <= 0) return 0;
  
  // 不完全ガンマ関数を使用
  return lowerIncompleteGamma(df / 2, x / 2) / gamma(df / 2);
}

// 下側不完全ガンマ関数（改良版）
function lowerIncompleteGamma(a: number, x: number): number {
  if (x <= 0) return 0;
  
  // 級数展開を使用（小さいxの場合）
  if (x < a + 1) {
    let sum = 0;
    let term = 1 / a;
    let ap = a;
    const maxIter = 200; // 反復限界を増加
    const tolerance = 1e-15; // より厳密な収束判定
    
    for (let i = 0; i < maxIter; i++) {
      sum += term;
      ap += 1;
      term *= x / ap;
      // 相対誤差で判定
      if (Math.abs(term) < Math.abs(sum) * tolerance) break;
    }
    
    return sum * Math.exp(-x + a * Math.log(x) - logGamma(a + 1));
  }
  
  // 連分数展開を使用（大きいxの場合）
  return gamma(a) - upperIncompleteGamma(a, x);
}

// 上側不完全ガンマ関数（改良版）
function upperIncompleteGamma(a: number, x: number): number {
  if (x <= 0) return gamma(a);
  
  // 連分数展開（Lentz's algorithm）
  const fpmin = 1e-30;
  const maxIter = 200; // 反復限界を増加
  const tolerance = 1e-15; // より厳密な収束判定
  
  let b = x + 1 - a;
  let c = 1 / fpmin;
  let d = 1 / b;
  let h = d;
  
  for (let i = 1; i < maxIter; i++) {
    const an = -i * (i - a);
    b += 2;
    d = an * d + b;
    if (Math.abs(d) < fpmin) d = fpmin;
    c = b + an / c;
    if (Math.abs(c) < fpmin) c = fpmin;
    d = 1 / d;
    const del = d * c;
    h *= del;
    
    // 相対誤差で判定
    if (Math.abs(del - 1) < tolerance) break;
  }
  
  return Math.exp(-x + a * Math.log(x) - logGamma(a)) * h;
}

// F分布の累積分布関数
function fDistributionCDF(x: number, df1: number, df2: number): number {
  if (x <= 0) return 0;
  
  // ベータ関数を使用してF分布のCDFを計算
  const y = df1 * x / (df1 * x + df2);
  return incompleteBeta(y, df1 / 2, df2 / 2);
}

// 不完全ベータ関数（正規化）- 高精度実装
function incompleteBeta(x: number, a: number, b: number): number {
  if (x <= 0) return 0;
  if (x >= 1) return 1;
  
  // ベータ関数 B(a,b) の対数
  const logBeta = logGamma(a) + logGamma(b) - logGamma(a + b);
  
  // x^a * (1-x)^b / B(a,b) の計算
  const bt = Math.exp(a * Math.log(x) + b * Math.log(1 - x) - logBeta);
  
  // 収束を良くするために対称性を利用
  if (x < (a + 1) / (a + b + 2)) {
    // 前方級数展開
    return bt * betaContinuedFraction(x, a, b) / a;
  } else {
    // 対称性を使用: I_x(a,b) = 1 - I_{1-x}(b,a)
    return 1 - bt * betaContinuedFraction(1 - x, b, a) / b;
  }
}

// ベータ連分数（Lentz法による高精度実装）
function betaContinuedFraction(x: number, a: number, b: number): number {
  const maxIterations = 200;
  const fpmin = 1e-30;
  const eps = 1e-15;
  
  const qab = a + b;
  const qap = a + 1;
  const qam = a - 1;
  
  // 初期値設定
  let c = 1;
  let d = 1 - qab * x / qap;
  if (Math.abs(d) < fpmin) d = fpmin;
  d = 1 / d;
  let h = d;
  
  for (let m = 1; m <= maxIterations; m++) {
    const m2 = 2 * m;
    
    // 偶数項
    let aa = m * (b - m) * x / ((qam + m2) * (a + m2));
    d = 1 + aa * d;
    if (Math.abs(d) < fpmin) d = fpmin;
    c = 1 + aa / c;
    if (Math.abs(c) < fpmin) c = fpmin;
    d = 1 / d;
    h *= d * c;
    
    // 奇数項
    aa = -(a + m) * (qab + m) * x / ((a + m2) * (qap + m2));
    d = 1 + aa * d;
    if (Math.abs(d) < fpmin) d = fpmin;
    c = 1 + aa / c;
    if (Math.abs(c) < fpmin) c = fpmin;
    d = 1 / d;
    const del = d * c;
    h *= del;
    
    // 収束判定
    if (Math.abs(del - 1) < eps) break;
  }
  
  return h;
}

// ガンマ関数の近似（Lanczos近似）
function gamma(z: number): number {
  if (z < 0.5) {
    return Math.PI / (Math.sin(Math.PI * z) * gamma(1 - z));
  }
  
  z -= 1;
  const x = 0.9999999999998099;
  const coefficients = [
    676.5203681218851, -1259.1392167224028, 771.32342877765313,
    -176.61502916214059, 12.507343278686905, -0.1385710952657201,
    9.984369578019572e-6, 1.505632735149312e-7
  ];
  
  let result = x;
  for (let i = 0; i < coefficients.length; i++) {
    result += coefficients[i] / (z + i + 1);
  }
  
  const t = z + coefficients.length - 0.5;
  return Math.sqrt(2 * Math.PI) * Math.pow(t, z + 0.5) * Math.exp(-t) * result;
}

// ガンマ関数の対数
function logGamma(z: number): number {
  if (z < 0.5) {
    return Math.log(Math.PI) - Math.log(Math.sin(Math.PI * z)) - logGamma(1 - z);
  }
  
  z -= 1;
  const x = 0.9999999999998099;
  const coefficients = [
    676.5203681218851, -1259.1392167224028, 771.32342877765313,
    -176.61502916214059, 12.507343278686905, -0.1385710952657201,
    9.984369578019572e-6, 1.505632735149312e-7
  ];
  
  let sum = x;
  for (let i = 0; i < coefficients.length; i++) {
    sum += coefficients[i] / (z + i + 1);
  }
  
  const t = z + coefficients.length - 0.5;
  
  return 0.5 * Math.log(2 * Math.PI) + (z + 0.5) * Math.log(t) - t + Math.log(sum);
}

// 標準正規分布の累積分布関数（誤差関数を使用）
function standardNormalCDF(z: number): number {
  return 0.5 * (1 + erf(z / SQRT_2));
}

// 誤差関数の高精度実装
function erf(x: number): number {
  // 小さい値に対しては級数展開を使用
  if (Math.abs(x) < 1.5) {
    return erfSmall(x);
  }
  
  // 大きい値に対しては相補誤差関数を使用
  const sign = x >= 0 ? 1 : -1;
  return sign * (1 - erfc(Math.abs(x)));
}

// 小さい値用の誤差関数（Maclaurin級数）
function erfSmall(x: number): number {
  const x2 = x * x;
  let sum = x;
  let term = x;
  
  for (let n = 1; n < 30; n++) {
    term *= -x2 * (2 * n - 1) / (n * (2 * n + 1));
    sum += term;
    if (Math.abs(term) < Math.abs(sum) * 1e-15) break;
  }
  
  return (2 / Math.sqrt(Math.PI)) * sum;
}

// 相補誤差関数の高精度実装
function erfc(x: number): number {
  if (x < 0) return 2 - erfc(-x);
  
  // Abramowitz and Stegun 7.1.26の有理近似
  const a = [
    -1.26551223,
     1.00002368,
     0.37409196,
     0.09678418,
    -0.18628806,
     0.27886807,
    -1.13520398,
     1.48851587,
    -0.82215223,
     0.17087277
  ];
  
  const z = 1 / (1 + 0.5 * x);
  
  const t = z * Math.exp(-x * x + 
    a[0] + z * (a[1] + z * (a[2] + z * (a[3] + z * (a[4] + 
    z * (a[5] + z * (a[6] + z * (a[7] + z * (a[8] + z * a[9])))))))));
  
  return t;
}

// 標準正規分布の確率密度関数
function standardNormalPDF(z: number): number {
  return Math.exp(-0.5 * z * z) / SQRT_2PI;
}

// ニュートン・ラフソン法による逆関数の近似
function inverseStandardNormal(p: number): number {
  if (p <= 0 || p >= 1) {
    throw new Error('Probability must be between 0 and 1');
  }
  
  // 初期値の設定
  let x = 0;
  if (p < 0.5) {
    x = -Math.sqrt(-2 * Math.log(p));
  } else {
    x = Math.sqrt(-2 * Math.log(1 - p));
  }
  
  // ニュートン・ラフソン法で改善
  for (let i = 0; i < 10; i++) {
    const fx = standardNormalCDF(x) - p;
    const fpx = standardNormalPDF(x);
    if (Math.abs(fpx) < 1e-15) break;
    const newX = x - fx / fpx;
    if (Math.abs(newX - x) < 1e-10) break;
    x = newX;
  }
  
  return x;
}

// 階乗関数（大きな数に対してはStirlingの近似を使用）
function factorial(n: number): number {
  if (n === 0 || n === 1) return 1;
  if (n < 0 || !Number.isInteger(n)) return NaN;
  
  // 小さな値は直接計算
  if (n <= 20) {
    let result = 1;
    for (let i = 2; i <= n; i++) {
      result *= i;
    }
    return result;
  }
  
  // 大きな値はガンマ関数を使用
  return gamma(n + 1);
}

// 不完全ガンマ関数（正規化）- 高精度実装
function incompleteGamma(a: number, x: number): number {
  if (x <= 0 || a <= 0) return 0;
  
  // 正規化された下側不完全ガンマ関数を返す
  // P(a,x) = γ(a,x) / Γ(a)
  return gammaP(a, x) * gamma(a);
}

// 正規化下側不完全ガンマ関数 P(a,x) = γ(a,x) / Γ(a)
function gammaP(a: number, x: number): number {
  if (x <= 0 || a <= 0) return 0;
  
  // x < a + 1 の場合は級数展開を使用
  if (x < a + 1) {
    return gammaPSeries(a, x);
  } else {
    // それ以外は連分数展開を使用して Q(a,x) から計算
    return 1 - gammaQContinuedFraction(a, x);
  }
}

// 級数展開による計算
function gammaPSeries(a: number, x: number): number {
  const maxIterations = 200;
  const eps = 1e-15;
  
  if (x === 0) return 0;
  
  let ap = a;
  let sum = 1 / a;
  let del = sum;
  
  for (let n = 1; n <= maxIterations; n++) {
    ap += 1;
    del *= x / ap;
    sum += del;
    if (Math.abs(del) < Math.abs(sum) * eps) {
      break;
    }
  }
  
  return sum * Math.exp(-x + a * Math.log(x) - logGamma(a));
}

// 連分数展開による上側不完全ガンマ関数 Q(a,x) = Γ(a,x) / Γ(a)
function gammaQContinuedFraction(a: number, x: number): number {
  const maxIterations = 200;
  const fpmin = 1e-30;
  const eps = 1e-15;
  
  let b = x + 1 - a;
  let c = 1 / fpmin;
  let d = 1 / b;
  let h = d;
  
  for (let i = 1; i <= maxIterations; i++) {
    const an = -i * (i - a);
    b += 2;
    d = an * d + b;
    if (Math.abs(d) < fpmin) d = fpmin;
    c = b + an / c;
    if (Math.abs(c) < fpmin) c = fpmin;
    d = 1 / d;
    const del = d * c;
    h *= del;
    if (Math.abs(del - 1) < eps) {
      break;
    }
  }
  
  return Math.exp(-x + a * Math.log(x) - logGamma(a)) * h;
}

// t分布の累積分布関数
function tCDF(t: number, df: number): number {
  const x = df / (df + t * t);
  if (t >= 0) {
    return 1 - 0.5 * incompleteBeta(x, df / 2, 0.5);
  } else {
    return 0.5 * incompleteBeta(x, df / 2, 0.5);
  }
}

// t分布の逆関数（改良版）
export function inverseTDistribution(p: number, df: number): number {
  if (p <= 0 || p >= 1) return NaN;
  if (df <= 0) return NaN;
  
  // 改良された初期推定値
  const z = inverseStandardNormal(p);
  let x;
  
  if (df > 30) {
    // 大きい自由度の場合は正規分布に近似
    x = z;
  } else {
    // Cornish-Fisher展開による初期推定値
    const g1 = 0.25 * (1 + z * z) / df;
    const g2 = (1 / 96) * z * (3 + z * z) * (1 + 3 * z * z) / (df * df);
    x = z + g1 * z + g2;
  }
  
  // Newton-Raphson法で精度を上げる
  const maxIter = 50; // 良い初期値なので反復回数を減らせる
  const tol = 1e-12; // より高精度
  
  for (let i = 0; i < maxIter; i++) {
    const cdf = tCDF(x, df);
    
    // t分布のPDF
    const pdf = Math.exp(logGamma((df + 1) / 2) - logGamma(df / 2)) / 
                Math.sqrt(Math.PI * df) * Math.pow(1 + x * x / df, -(df + 1) / 2);
    
    // Newton-Raphson更新
    const error = cdf - p;
    if (Math.abs(error) < tol) break;
    
    // 更新量を制限して安定性を向上
    const delta = error / pdf;
    const maxDelta = Math.abs(x) + 1;
    if (Math.abs(delta) > maxDelta) {
      x = x - Math.sign(delta) * maxDelta;
    } else {
      x = x - delta;
    }
  }
  
  return x;
}

// F分布の累積分布関数（現在未使用）
// function fCDF(f: number, df1: number, df2: number): number {
//   const x = (df1 * f) / (df1 * f + df2);
//   return incompleteBeta(x, df1 / 2, df2 / 2);
// }

// カイ二乗分布の累積分布関数（現在未使用）
// function chiSquareCDF(chi: number, df: number): number {
//   return incompleteGamma(df / 2, chi / 2) / gamma(df / 2);
// }

// NORM.DIST関数（正規分布）
export const NORM_DIST: CustomFormula = {
  name: 'NORM.DIST',
  pattern: /NORM\.DIST\(([^,]+),\s*([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, xRef, meanRef, stdevRef, cumulativeRef] = matches;
    
    const x = Number(getCellValue(xRef.trim(), context) ?? xRef);
    const mean = Number(getCellValue(meanRef.trim(), context) ?? meanRef);
    const stdev = Number(getCellValue(stdevRef.trim(), context) ?? stdevRef);
    const cumulative = getCellValue(cumulativeRef.trim(), context) ?? cumulativeRef;
    
    if (isNaN(x) || isNaN(mean) || isNaN(stdev)) {
      return FormulaError.VALUE;
    }
    
    if (stdev <= 0) {
      return FormulaError.NUM;
    }
    
    const z = (x - mean) / stdev;
    const isCumulative = cumulative.toString().toLowerCase() === 'true' || cumulative === 1;
    
    if (isCumulative) {
      // 累積分布関数
      return standardNormalCDF(z);
    } else {
      // 確率密度関数
      return standardNormalPDF(z) / stdev;
    }
  }
};

// NORM.INV関数（正規分布の逆関数）
export const NORM_INV: CustomFormula = {
  name: 'NORM.INV',
  pattern: /NORM\.INV\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, probabilityRef, meanRef, stdevRef] = matches;
    
    const probability = Number(getCellValue(probabilityRef.trim(), context) ?? probabilityRef);
    const mean = Number(getCellValue(meanRef.trim(), context) ?? meanRef);
    const stdev = Number(getCellValue(stdevRef.trim(), context) ?? stdevRef);
    
    if (isNaN(probability) || isNaN(mean) || isNaN(stdev)) {
      return FormulaError.VALUE;
    }
    
    if (probability <= 0 || probability >= 1 || stdev <= 0) {
      return FormulaError.NUM;
    }
    
    try {
      const z = inverseStandardNormal(probability);
      return mean + z * stdev;
    } catch {
      return FormulaError.NUM;
    }
  }
};

// NORM.S.DIST関数（標準正規分布）
export const NORM_S_DIST: CustomFormula = {
  name: 'NORM.S.DIST',
  pattern: /NORM\.S\.DIST\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, zRef, cumulativeRef] = matches;
    
    const z = Number(getCellValue(zRef.trim(), context) ?? zRef);
    const cumulative = getCellValue(cumulativeRef.trim(), context) ?? cumulativeRef;
    
    if (isNaN(z)) {
      return FormulaError.VALUE;
    }
    
    const isCumulative = cumulative.toString().toLowerCase() === 'true' || cumulative === 1;
    
    if (isCumulative) {
      return standardNormalCDF(z);
    } else {
      return standardNormalPDF(z);
    }
  }
};

// NORM.S.INV関数（標準正規分布の逆関数）
export const NORM_S_INV: CustomFormula = {
  name: 'NORM.S.INV',
  pattern: /NORM\.S\.INV\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, probabilityRef] = matches;
    
    const probability = Number(getCellValue(probabilityRef.trim(), context) ?? probabilityRef);
    
    if (isNaN(probability)) {
      return FormulaError.VALUE;
    }
    
    if (probability <= 0 || probability >= 1) {
      return FormulaError.NUM;
    }
    
    try {
      return inverseStandardNormal(probability);
    } catch {
      return FormulaError.NUM;
    }
  }
};

// EXPON.DIST関数（指数分布）
export const EXPON_DIST: CustomFormula = {
  name: 'EXPON.DIST',
  pattern: /EXPON\.DIST\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, xRef, lambdaRef, cumulativeRef] = matches;
    
    const x = Number(getCellValue(xRef.trim(), context) ?? xRef);
    const lambda = Number(getCellValue(lambdaRef.trim(), context) ?? lambdaRef);
    const cumulative = getCellValue(cumulativeRef.trim(), context) ?? cumulativeRef;
    
    if (isNaN(x) || isNaN(lambda)) {
      return FormulaError.VALUE;
    }
    
    if (x < 0 || lambda <= 0) {
      return FormulaError.NUM;
    }
    
    const isCumulative = cumulative.toString().toLowerCase() === 'true' || cumulative === 1;
    
    if (isCumulative) {
      // 累積分布関数: 1 - e^(-λx)
      return 1 - Math.exp(-lambda * x);
    } else {
      // 確率密度関数: λe^(-λx)
      return lambda * Math.exp(-lambda * x);
    }
  }
};

// BINOM.DIST関数（二項分布）
export const BINOM_DIST: CustomFormula = {
  name: 'BINOM.DIST',
  pattern: /BINOM\.DIST\(([^,]+),\s*([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, successRef, trialsRef, probSRef, cumulativeRef] = matches;
    
    const success = Number(getCellValue(successRef.trim(), context) ?? successRef);
    const trials = Number(getCellValue(trialsRef.trim(), context) ?? trialsRef);
    const probS = Number(getCellValue(probSRef.trim(), context) ?? probSRef);
    const cumulative = getCellValue(cumulativeRef.trim(), context) ?? cumulativeRef;
    
    if (isNaN(success) || isNaN(trials) || isNaN(probS)) {
      return FormulaError.VALUE;
    }
    
    if (success < 0 || trials < 0 || !Number.isInteger(success) || !Number.isInteger(trials) ||
        success > trials || probS < 0 || probS > 1) {
      return FormulaError.NUM;
    }
    
    const isCumulative = cumulative.toString().toLowerCase() === 'true' || cumulative === 1;
    
    if (isCumulative) {
      // 累積分布関数
      let result = 0;
      for (let k = 0; k <= success; k++) {
        const binomCoeff = factorial(trials) / (factorial(k) * factorial(trials - k));
        result += binomCoeff * Math.pow(probS, k) * Math.pow(1 - probS, trials - k);
      }
      return result;
    } else {
      // 確率質量関数
      const binomCoeff = factorial(trials) / (factorial(success) * factorial(trials - success));
      return binomCoeff * Math.pow(probS, success) * Math.pow(1 - probS, trials - success);
    }
  }
};

// POISSON.DIST関数（ポアソン分布）
export const POISSON_DIST: CustomFormula = {
  name: 'POISSON.DIST',
  pattern: /POISSON\.DIST\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, xRef, meanRef, cumulativeRef] = matches;
    
    const x = Number(getCellValue(xRef.trim(), context) ?? xRef);
    const mean = Number(getCellValue(meanRef.trim(), context) ?? meanRef);
    const cumulative = getCellValue(cumulativeRef.trim(), context) ?? cumulativeRef;
    
    if (isNaN(x) || isNaN(mean)) {
      return FormulaError.VALUE;
    }
    
    if (x < 0 || !Number.isInteger(x) || mean <= 0) {
      return FormulaError.NUM;
    }
    
    const isCumulative = cumulative.toString().toLowerCase() === 'true' || cumulative === 1;
    
    if (isCumulative) {
      // 累積分布関数
      let result = 0;
      for (let k = 0; k <= x; k++) {
        result += (Math.pow(mean, k) * Math.exp(-mean)) / factorial(k);
      }
      return result;
    } else {
      // 確率質量関数: (λ^x * e^(-λ)) / x!
      return (Math.pow(mean, x) * Math.exp(-mean)) / factorial(x);
    }
  }
};

// GAMMA.DIST関数（ガンマ分布）
export const GAMMA_DIST: CustomFormula = {
  name: 'GAMMA.DIST',
  pattern: /GAMMA\.DIST\(([^,]+),\s*([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, xRef, alphaRef, betaRef, cumulativeRef] = matches;
    
    const x = Number(getCellValue(xRef.trim(), context) ?? xRef);
    const alpha = Number(getCellValue(alphaRef.trim(), context) ?? alphaRef);
    const beta = Number(getCellValue(betaRef.trim(), context) ?? betaRef);
    const cumulative = getCellValue(cumulativeRef.trim(), context) ?? cumulativeRef;
    
    if (isNaN(x) || isNaN(alpha) || isNaN(beta)) {
      return FormulaError.VALUE;
    }
    
    if (x < 0 || alpha <= 0 || beta <= 0) {
      return FormulaError.NUM;
    }
    
    const isCumulative = cumulative.toString().toLowerCase() === 'true' || cumulative === 1;
    
    if (isCumulative) {
      // 累積分布関数（不完全ガンマ関数を使用）
      return incompleteGamma(alpha, x / beta) / gamma(alpha);
    } else {
      // 確率密度関数: (x^(α-1) * e^(-x/β)) / (β^α * Γ(α))
      return Math.pow(x, alpha - 1) * Math.exp(-x / beta) / (Math.pow(beta, alpha) * gamma(alpha));
    }
  }
};

// LOGNORM.DIST関数（対数正規分布）
export const LOGNORM_DIST: CustomFormula = {
  name: 'LOGNORM.DIST',
  pattern: /LOGNORM\.DIST\(([^,]+),\s*([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, xRef, meanRef, stdevRef, cumulativeRef] = matches;
    
    const x = Number(getCellValue(xRef.trim(), context) ?? xRef);
    const mean = Number(getCellValue(meanRef.trim(), context) ?? meanRef);
    const stdev = Number(getCellValue(stdevRef.trim(), context) ?? stdevRef);
    const cumulative = getCellValue(cumulativeRef.trim(), context) ?? cumulativeRef;
    
    if (isNaN(x) || isNaN(mean) || isNaN(stdev)) {
      return FormulaError.VALUE;
    }
    
    if (x <= 0 || stdev <= 0) {
      return FormulaError.NUM;
    }
    
    const isCumulative = cumulative.toString().toLowerCase() === 'true' || cumulative === 1;
    const z = (Math.log(x) - mean) / stdev;
    
    if (isCumulative) {
      // 累積分布関数
      return standardNormalCDF(z);
    } else {
      // 確率密度関数
      return standardNormalPDF(z) / (x * stdev);
    }
  }
};

// LOGNORM.INV関数（対数正規分布の逆関数）
export const LOGNORM_INV: CustomFormula = {
  name: 'LOGNORM.INV',
  pattern: /LOGNORM\.INV\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, probabilityRef, meanRef, stdevRef] = matches;
    
    const probability = Number(getCellValue(probabilityRef.trim(), context) ?? probabilityRef);
    const mean = Number(getCellValue(meanRef.trim(), context) ?? meanRef);
    const stdev = Number(getCellValue(stdevRef.trim(), context) ?? stdevRef);
    
    if (isNaN(probability) || isNaN(mean) || isNaN(stdev)) {
      return FormulaError.VALUE;
    }
    
    if (probability <= 0 || probability >= 1 || stdev <= 0) {
      return FormulaError.NUM;
    }
    
    try {
      const z = inverseStandardNormal(probability);
      return Math.exp(mean + z * stdev);
    } catch {
      return FormulaError.NUM;
    }
  }
};

// WEIBULL.DIST関数（ワイブル分布）
export const WEIBULL_DIST: CustomFormula = {
  name: 'WEIBULL.DIST',
  pattern: /WEIBULL\.DIST\(([^,]+),\s*([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, xRef, alphaRef, betaRef, cumulativeRef] = matches;
    
    const x = Number(getCellValue(xRef.trim(), context) ?? xRef);
    const alpha = Number(getCellValue(alphaRef.trim(), context) ?? alphaRef);
    const beta = Number(getCellValue(betaRef.trim(), context) ?? betaRef);
    const cumulative = getCellValue(cumulativeRef.trim(), context) ?? cumulativeRef;
    
    if (isNaN(x) || isNaN(alpha) || isNaN(beta)) {
      return FormulaError.VALUE;
    }
    
    if (x < 0 || alpha <= 0 || beta <= 0) {
      return FormulaError.NUM;
    }
    
    const isCumulative = cumulative.toString().toLowerCase() === 'true' || cumulative === 1;
    
    if (isCumulative) {
      // 累積分布関数: 1 - e^(-(x/β)^α)
      return 1 - Math.exp(-Math.pow(x / beta, alpha));
    } else {
      // 確率密度関数: (α/β) * (x/β)^(α-1) * e^(-(x/β)^α)
      return (alpha / beta) * Math.pow(x / beta, alpha - 1) * Math.exp(-Math.pow(x / beta, alpha));
    }
  }
};

// ベータ関数の近似
function beta(a: number, b: number): number {
  return gamma(a) * gamma(b) / gamma(a + b);
}


// T.DIST関数（t分布・左側）
export const T_DIST: CustomFormula = {
  name: 'T.DIST',
  pattern: /T\.DIST\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, xRef, degFreedomRef, cumulativeRef] = matches;
    
    const x = Number(getCellValue(xRef.trim(), context) ?? xRef);
    const degFreedom = Number(getCellValue(degFreedomRef.trim(), context) ?? degFreedomRef);
    const cumulative = getCellValue(cumulativeRef.trim(), context) ?? cumulativeRef;
    
    if (isNaN(x) || isNaN(degFreedom)) {
      return FormulaError.VALUE;
    }
    
    if (degFreedom < 1 || !Number.isInteger(degFreedom)) {
      return FormulaError.NUM;
    }
    
    const isCumulative = cumulative.toString().toLowerCase() === 'true' || cumulative === 1;
    
    if (isCumulative) {
      // 累積分布関数（不完全ベータ関数を使用）
      // t分布のCDF: P(T ≤ x) = I_{x/(√(x²+ν))}(ν/2, 1/2) for x ≥ 0
      const sign = x < 0 ? -1 : 1;
      const absX = Math.abs(x);
      const z = degFreedom / (degFreedom + absX * absX);
      
      if (x < 0) {
        // 負の値の場合: P(T ≤ x) = 1/2 * I_z(ν/2, 1/2)
        return 0.5 * incompleteBeta(z, degFreedom / 2, 0.5);
      } else {
        // 正の値の場合: P(T ≤ x) = 1 - 1/2 * I_z(ν/2, 1/2)
        return 1 - 0.5 * incompleteBeta(z, degFreedom / 2, 0.5);
      }
    } else {
      // 確率密度関数
      const coefficient = gamma((degFreedom + 1) / 2) / (Math.sqrt(degFreedom * Math.PI) * gamma(degFreedom / 2));
      return coefficient * Math.pow(1 + (x * x) / degFreedom, -(degFreedom + 1) / 2);
    }
  }
};

// T.DIST.2T関数（t分布・両側）
export const T_DIST_2T: CustomFormula = {
  name: 'T.DIST.2T',
  pattern: /T\.DIST\.2T\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, xRef, degFreedomRef] = matches;
    
    const xOriginal = Number(getCellValue(xRef.trim(), context) ?? xRef);
    const degFreedom = Number(getCellValue(degFreedomRef.trim(), context) ?? degFreedomRef);
    
    if (isNaN(xOriginal) || isNaN(degFreedom)) {
      return FormulaError.VALUE;
    }
    
    if (xOriginal < 0) {
      return FormulaError.NUM;
    }
    
    if (degFreedom < 1 || !Number.isInteger(degFreedom)) {
      return FormulaError.NUM;
    }
    
    // T.DIST.2T(x, deg) = 2 * P(T > |x|) where T ~ t(deg)
    // For t-distribution with degrees of freedom df:
    // P(|T| > x) = 1 - I_(x²/(x²+df))(1/2, df/2) where I is the regularized incomplete beta function
    const x = Math.abs(xOriginal);
    
    if (x === 0) {
      return 1; // P(|T| > 0) = 1
    }
    
    // Calculate the regularized incomplete beta function
    // w = x² / (x² + df)
    const w = (x * x) / (x * x + degFreedom);
    const betaValue = incompleteBeta(w, 0.5, degFreedom / 2);
    
    // T.DIST.2T = 2 * P(T > |x|) = 1 - I_w(1/2, df/2)
    const result = 1 - betaValue;
    
    return Math.max(0, Math.min(1, result));
  }
};

// T.DIST.RT関数（t分布・右側）
export const T_DIST_RT: CustomFormula = {
  name: 'T.DIST.RT',
  pattern: /T\.DIST\.RT\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, xRef, degFreedomRef] = matches;
    
    const x = Number(getCellValue(xRef.trim(), context) ?? xRef);
    const degFreedom = Number(getCellValue(degFreedomRef.trim(), context) ?? degFreedomRef);
    
    if (isNaN(x) || isNaN(degFreedom)) {
      return FormulaError.VALUE;
    }
    
    if (degFreedom < 1 || !Number.isInteger(degFreedom)) {
      return FormulaError.NUM;
    }
    
    // 右側確率 = P(T > x)
    const z = degFreedom / (degFreedom + x * x);
    
    if (x < 0) {
      // 負の値の場合: P(T > x) = 1 - 1/2 * I_z(ν/2, 1/2)
      return 1 - 0.5 * incompleteBeta(z, degFreedom / 2, 0.5);
    } else {
      // 正の値の場合: P(T > x) = 1/2 * I_z(ν/2, 1/2)
      return 0.5 * incompleteBeta(z, degFreedom / 2, 0.5);
    }
  }
};

// CHISQ.DIST関数（カイ二乗分布）
export const CHISQ_DIST: CustomFormula = {
  name: 'CHISQ.DIST',
  pattern: /CHISQ\.DIST\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, xRef, degFreedomRef, cumulativeRef] = matches;
    
    const x = Number(getCellValue(xRef.trim(), context) ?? xRef);
    const degFreedom = Number(getCellValue(degFreedomRef.trim(), context) ?? degFreedomRef);
    const cumulative = getCellValue(cumulativeRef.trim(), context) ?? cumulativeRef;
    
    if (isNaN(x) || isNaN(degFreedom)) {
      return FormulaError.VALUE;
    }
    
    if (x < 0 || degFreedom < 1 || !Number.isInteger(degFreedom)) {
      return FormulaError.NUM;
    }
    
    const isCumulative = cumulative.toString().toLowerCase() === 'true' || cumulative === 1;
    
    if (isCumulative) {
      // 累積分布関数（不完全ガンマ関数の正規化）
      return incompleteGamma(degFreedom / 2, x / 2) / gamma(degFreedom / 2);
    } else {
      // 確率密度関数
      if (x === 0) {
        return degFreedom === 2 ? 0.5 : 0;
      }
      const coefficient = 1 / (Math.pow(2, degFreedom / 2) * gamma(degFreedom / 2));
      return coefficient * Math.pow(x, degFreedom / 2 - 1) * Math.exp(-x / 2);
    }
  }
};

// CHISQ.DIST.RT関数（カイ二乗分布・右側）
export const CHISQ_DIST_RT: CustomFormula = {
  name: 'CHISQ.DIST.RT',
  pattern: /CHISQ\.DIST\.RT\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, xRef, degFreedomRef] = matches;
    
    const x = Number(getCellValue(xRef.trim(), context) ?? xRef);
    const degFreedom = Number(getCellValue(degFreedomRef.trim(), context) ?? degFreedomRef);
    
    if (isNaN(x) || isNaN(degFreedom)) {
      return FormulaError.VALUE;
    }
    
    if (x < 0 || degFreedom < 1 || !Number.isInteger(degFreedom)) {
      return FormulaError.NUM;
    }
    
    // 右側確率 = 1 - 左側確率
    const leftTail = incompleteGamma(degFreedom / 2, x / 2) / gamma(degFreedom / 2);
    return 1 - leftTail;
  }
};

// F.DIST関数（F分布）
export const F_DIST: CustomFormula = {
  name: 'F.DIST',
  pattern: /F\.DIST\(([^,]+),\s*([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, xRef, deg1Ref, deg2Ref, cumulativeRef] = matches;
    
    const x = Number(getCellValue(xRef.trim(), context) ?? xRef);
    const deg1 = Number(getCellValue(deg1Ref.trim(), context) ?? deg1Ref);
    const deg2 = Number(getCellValue(deg2Ref.trim(), context) ?? deg2Ref);
    const cumulative = getCellValue(cumulativeRef.trim(), context) ?? cumulativeRef;
    
    if (isNaN(x) || isNaN(deg1) || isNaN(deg2)) {
      return FormulaError.VALUE;
    }
    
    if (x < 0 || deg1 < 1 || deg2 < 1 || !Number.isInteger(deg1) || !Number.isInteger(deg2)) {
      return FormulaError.NUM;
    }
    
    const isCumulative = cumulative.toString().toLowerCase() === 'true' || cumulative === 1;
    
    if (isCumulative) {
      // 累積分布関数（不完全ベータ関数を使用）
      const t = (deg1 * x) / (deg1 * x + deg2);
      return incompleteBeta(t, deg1 / 2, deg2 / 2);
    } else {
      // 確率密度関数
      if (x === 0) return 0;
      
      const coefficient = (gamma((deg1 + deg2) / 2) / (gamma(deg1 / 2) * gamma(deg2 / 2))) *
                         Math.pow(deg1 / deg2, deg1 / 2);
      const numerator = Math.pow(x, deg1 / 2 - 1);
      const denominator = Math.pow(1 + (deg1 * x) / deg2, (deg1 + deg2) / 2);
      
      return coefficient * numerator / denominator;
    }
  }
};

// F.DIST.RT関数（F分布・右側）
export const F_DIST_RT: CustomFormula = {
  name: 'F.DIST.RT',
  pattern: /F\.DIST\.RT\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, xRef, deg1Ref, deg2Ref] = matches;
    
    const x = Number(getCellValue(xRef.trim(), context) ?? xRef);
    const deg1 = Number(getCellValue(deg1Ref.trim(), context) ?? deg1Ref);
    const deg2 = Number(getCellValue(deg2Ref.trim(), context) ?? deg2Ref);
    
    if (isNaN(x) || isNaN(deg1) || isNaN(deg2)) {
      return FormulaError.VALUE;
    }
    
    if (x < 0 || deg1 < 1 || deg2 < 1 || !Number.isInteger(deg1) || !Number.isInteger(deg2)) {
      return FormulaError.NUM;
    }
    
    // 右側確率 = 1 - 左側確率
    const t = (deg1 * x) / (deg1 * x + deg2);
    const leftTail = incompleteBeta(t, deg1 / 2, deg2 / 2);
    return 1 - leftTail;
  }
};

// BETA.DIST関数（ベータ分布）
export const BETA_DIST: CustomFormula = {
  name: 'BETA.DIST',
  pattern: /BETA\.DIST\(([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,)]+)(?:,\s*([^,)]+))?(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, xRef, alphaRef, betaRef, cumulativeRef, ARef, BRef] = matches;
    
    const x = Number(getCellValue(xRef.trim(), context) ?? xRef);
    const alpha = Number(getCellValue(alphaRef.trim(), context) ?? alphaRef);
    const betaParam = Number(getCellValue(betaRef.trim(), context) ?? betaRef);
    const cumulative = getCellValue(cumulativeRef.trim(), context) ?? cumulativeRef;
    const A = ARef ? Number(getCellValue(ARef.trim(), context) ?? ARef) : 0;
    const B = BRef ? Number(getCellValue(BRef.trim(), context) ?? BRef) : 1;
    
    if (isNaN(x) || isNaN(alpha) || isNaN(betaParam)) {
      return FormulaError.VALUE;
    }
    
    if (alpha <= 0 || betaParam <= 0 || A >= B) {
      return FormulaError.NUM;
    }
    
    // 標準化
    const standardX = (x - A) / (B - A);
    if (standardX < 0 || standardX > 1) {
      return FormulaError.NUM;
    }
    
    const isCumulative = cumulative.toString().toLowerCase() === 'true' || cumulative === 1;
    
    if (isCumulative) {
      // 累積分布関数
      return incompleteBeta(standardX, alpha, betaParam);
    } else {
      // 確率密度関数
      if (standardX === 0 || standardX === 1) {
        if ((alpha === 1 && standardX === 0) || (betaParam === 1 && standardX === 1)) {
          return 1 / (B - A);
        }
        return 0;
      }
      
      const coefficient = 1 / (beta(alpha, betaParam) * (B - A));
      return coefficient * Math.pow(standardX, alpha - 1) * Math.pow(1 - standardX, betaParam - 1);
    }
  }
};

// NEGBINOM.DIST関数（負の二項分布）
export const NEGBINOM_DIST: CustomFormula = {
  name: 'NEGBINOM.DIST',
  pattern: /NEGBINOM\.DIST\(([^,]+),\s*([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, numberFRef, numberSRef, probabilitySRef, cumulativeRef] = matches;
    
    // ExcelのNEGBINOM.DISTの引数：
    // number_f: 失敗回数
    // number_s: 成功回数のしきい値
    // probability_s: 成功の確率
    const numberF = Number(getCellValue(numberFRef.trim(), context) ?? numberFRef);
    const numberS = Number(getCellValue(numberSRef.trim(), context) ?? numberSRef);
    const probabilityS = Number(getCellValue(probabilitySRef.trim(), context) ?? probabilitySRef);
    const cumulative = getCellValue(cumulativeRef.trim(), context) ?? cumulativeRef;
    
    if (isNaN(numberF) || isNaN(numberS) || isNaN(probabilityS)) {
      return FormulaError.VALUE;
    }
    
    if (numberF < 0 || numberS < 1 || !Number.isInteger(numberF) || 
        probabilityS <= 0 || probabilityS > 1) {
      return FormulaError.NUM;
    }
    
    const isCumulative = cumulative.toString().toLowerCase() === 'true' || cumulative === 1;
    
    if (isCumulative) {
      // 累積分布関数
      // I(x; r, p) = ベータ分布を使用
      return incompleteBeta(probabilityS, numberS, numberF + 1);
    } else {
      // 確率質量関数
      // P(X = k) = C(k+r-1, k) * p^r * (1-p)^k
      // ここで k = numberF, r = numberS, p = probabilityS
      // factorialが大きい数でオーバーフローする可能性があるため、
      // 対数を使用した計算を行う
      const logCoeff = logGamma(numberF + numberS) - logGamma(numberF + 1) - logGamma(numberS);
      const logResult = logCoeff + numberS * Math.log(probabilityS) + numberF * Math.log(1 - probabilityS);
      return Math.exp(logResult);
    }
  }
};

// HYPGEOM.DIST関数（超幾何分布）
export const HYPGEOM_DIST: CustomFormula = {
  name: 'HYPGEOM.DIST',
  pattern: /HYPGEOM\.DIST\(([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, sampleSRef, sampleNRef, popSRef, popNRef, cumulativeRef] = matches;
    
    const sampleS = Number(getCellValue(sampleSRef.trim(), context) ?? sampleSRef);
    const sampleN = Number(getCellValue(sampleNRef.trim(), context) ?? sampleNRef);
    const popS = Number(getCellValue(popSRef.trim(), context) ?? popSRef);
    const popN = Number(getCellValue(popNRef.trim(), context) ?? popNRef);
    const cumulative = getCellValue(cumulativeRef.trim(), context) ?? cumulativeRef;
    
    if (isNaN(sampleS) || isNaN(sampleN) || isNaN(popS) || isNaN(popN)) {
      return FormulaError.VALUE;
    }
    
    if (sampleS < 0 || sampleN < 0 || popS < 0 || popN < 0 ||
        !Number.isInteger(sampleS) || !Number.isInteger(sampleN) || 
        !Number.isInteger(popS) || !Number.isInteger(popN) ||
        sampleS > sampleN || popS > popN || sampleN > popN ||
        sampleS > popS || (sampleN - sampleS) > (popN - popS)) {
      return FormulaError.NUM;
    }
    
    const isCumulative = cumulative.toString().toLowerCase() === 'true' || cumulative === 1;
    
    
    // 組み合わせ計算（対数を使用してオーバーフローを防ぐ）
    const logCombination = (n: number, k: number): number => {
      if (k > n || k < 0) return -Infinity;
      if (k === 0 || k === n) return 0;
      if (k > n - k) k = n - k;
      
      let logResult = 0;
      for (let i = 0; i < k; i++) {
        logResult += Math.log(n - i) - Math.log(i + 1);
      }
      return logResult;
    };
    
    if (isCumulative) {
      // 累積分布関数
      let result = 0;
      const minK = Math.max(0, sampleN + popS - popN);
      
      for (let k = minK; k <= sampleS; k++) {
        const logProb = logCombination(popS, k) + 
                        logCombination(popN - popS, sampleN - k) - 
                        logCombination(popN, sampleN);
        result += Math.exp(logProb);
      }
      return result;
    } else {
      // 確率質量関数
      const logProb = logCombination(popS, sampleS) + 
                      logCombination(popN - popS, sampleN - sampleS) - 
                      logCombination(popN, sampleN);
      
      return Math.exp(logProb);
    }
  }
};


// === 検定・推定関数 ===

// CONFIDENCE.NORM関数（信頼区間・正規分布）
export const CONFIDENCE_NORM: CustomFormula = {
  name: 'CONFIDENCE.NORM',
  pattern: /CONFIDENCE\.NORM\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, alphaRef, stdDevRef, sizeRef] = matches;
    
    const alpha = Number(getCellValue(alphaRef.trim(), context) ?? alphaRef);
    const stdDev = Number(getCellValue(stdDevRef.trim(), context) ?? stdDevRef);
    const size = Number(getCellValue(sizeRef.trim(), context) ?? sizeRef);
    
    if (isNaN(alpha) || isNaN(stdDev) || isNaN(size)) {
      return FormulaError.VALUE;
    }
    
    if (alpha <= 0 || alpha >= 1 || stdDev <= 0 || size < 1 || !Number.isInteger(size)) {
      return FormulaError.NUM;
    }
    
    // 標準正規分布の逆関数を使用
    const z = Math.abs(inverseStandardNormal(1 - alpha / 2));
    
    // 信頼区間の幅
    return z * stdDev / Math.sqrt(size);
  }
};

// CONFIDENCE.T関数（信頼区間・t分布）
export const CONFIDENCE_T: CustomFormula = {
  name: 'CONFIDENCE.T',
  pattern: /CONFIDENCE\.T\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, alphaRef, stdDevRef, sizeRef] = matches;
    
    const alpha = Number(getCellValue(alphaRef.trim(), context) ?? alphaRef);
    const stdDev = Number(getCellValue(stdDevRef.trim(), context) ?? stdDevRef);
    const size = Number(getCellValue(sizeRef.trim(), context) ?? sizeRef);
    
    if (isNaN(alpha) || isNaN(stdDev) || isNaN(size)) {
      return FormulaError.VALUE;
    }
    
    if (alpha <= 0 || alpha >= 1 || stdDev <= 0 || size < 1) {
      return FormulaError.NUM;
    }
    
    // 自由度
    const df = size - 1;
    
    // t分布の逆関数を使用 (両側確率)
    const t = inverseTDistribution(1 - alpha / 2, df);
    
    // 信頼区間の幅
    return t * stdDev / Math.sqrt(size);
  }
};

// Z.TEST関数（z検定）
export const Z_TEST: CustomFormula = {
  name: 'Z.TEST',
  pattern: /Z\.TEST\(([^,]+),\s*([^,]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, arrayRef, muRef, sigmaRef] = matches;
    
    // データ範囲から値を取得
    const values = getCellRangeValues(arrayRef.trim(), context);
    const numbers = values.filter(v => typeof v === 'number' && !isNaN(v) && isFinite(v)) as number[];
    
    if (numbers.length === 0) {
      return FormulaError.NUM;
    }
    
    const mu = Number(getCellValue(muRef.trim(), context) ?? muRef);
    if (isNaN(mu)) {
      return FormulaError.VALUE;
    }
    
    // 標本平均を計算
    const sampleMean = numbers.reduce((sum, n) => sum + n, 0) / numbers.length;
    
    // 標準偏差を計算または取得
    let sigma: number;
    if (sigmaRef) {
      sigma = Number(getCellValue(sigmaRef.trim(), context) ?? sigmaRef);
      if (isNaN(sigma) || sigma <= 0) {
        return FormulaError.NUM;
      }
    } else {
      // 標本標準偏差を計算
      const variance = numbers.reduce((sum, n) => sum + Math.pow(n - sampleMean, 2), 0) / numbers.length;
      sigma = Math.sqrt(variance);
      if (sigma === 0) {
        return FormulaError.DIV0;
      }
    }
    
    // z統計量を計算
    const z = (sampleMean - mu) / (sigma / Math.sqrt(numbers.length));
    
    // 片側p値を返す（上側確率）
    return 1 - standardNormalCDF(z);
  }
};

// T.TEST関数（t検定）
export const T_TEST: CustomFormula = {
  name: 'T.TEST',
  pattern: /T\.TEST\(([^,]+),\s*([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, array1Ref, array2Ref, tailsRef, typeRef] = matches;
    
    // データ範囲から値を取得
    const values1 = getCellRangeValues(array1Ref.trim(), context);
    const values2 = getCellRangeValues(array2Ref.trim(), context);
    
    const numbers1 = values1.filter(v => typeof v === 'number' && !isNaN(v) && isFinite(v)) as number[];
    const numbers2 = values2.filter(v => typeof v === 'number' && !isNaN(v) && isFinite(v)) as number[];
    
    if (numbers1.length < 2 || numbers2.length < 2) {
      return FormulaError.NUM;
    }
    
    const tails = Number(getCellValue(tailsRef.trim(), context) ?? tailsRef);
    const type = Number(getCellValue(typeRef.trim(), context) ?? typeRef);
    
    if (isNaN(tails) || isNaN(type) || ![1, 2].includes(tails) || ![1, 2, 3].includes(type)) {
      return FormulaError.VALUE;
    }
    
    // 平均を計算
    const mean1 = numbers1.reduce((sum, n) => sum + n, 0) / numbers1.length;
    const mean2 = numbers2.reduce((sum, n) => sum + n, 0) / numbers2.length;
    
    // 分散を計算
    const var1 = numbers1.reduce((sum, n) => sum + Math.pow(n - mean1, 2), 0) / (numbers1.length - 1);
    const var2 = numbers2.reduce((sum, n) => sum + Math.pow(n - mean2, 2), 0) / (numbers2.length - 1);
    
    let t: number;
    let df: number;
    
    if (type === 1) {
      // 対応のあるt検定
      if (numbers1.length !== numbers2.length) {
        return FormulaError.NA;
      }
      
      const differences = numbers1.map((n, i) => n - numbers2[i]);
      const meanDiff = differences.reduce((sum, d) => sum + d, 0) / differences.length;
      const varDiff = differences.reduce((sum, d) => sum + Math.pow(d - meanDiff, 2), 0) / (differences.length - 1);
      
      t = meanDiff / (Math.sqrt(varDiff / differences.length));
      df = differences.length - 1;
      
    } else if (type === 2) {
      // 等分散を仮定した独立標本t検定
      const pooledVar = ((numbers1.length - 1) * var1 + (numbers2.length - 1) * var2) / 
                        (numbers1.length + numbers2.length - 2);
      t = (mean1 - mean2) / Math.sqrt(pooledVar * (1 / numbers1.length + 1 / numbers2.length));
      df = numbers1.length + numbers2.length - 2;
      
    } else {
      // 等分散を仮定しない独立標本t検定（Welchのt検定）
      t = (mean1 - mean2) / Math.sqrt(var1 / numbers1.length + var2 / numbers2.length);
      
      // Welch-Satterthwaiteの自由度
      const num = Math.pow(var1 / numbers1.length + var2 / numbers2.length, 2);
      const denom = Math.pow(var1 / numbers1.length, 2) / (numbers1.length - 1) + 
                    Math.pow(var2 / numbers2.length, 2) / (numbers2.length - 1);
      df = num / denom;
    }
    
    // t分布の累積分布関数を使用してp値を計算
    // tCDFは左側確率を返すので、両側検定の場合は調整が必要
    if (tails === 1) {
      // 片側検定
      if (t >= 0) {
        return 1 - tCDF(t, df);
      } else {
        return tCDF(t, df);
      }
    } else {
      // 両側検定
      return 2 * (1 - tCDF(Math.abs(t), df));
    }
  }
};

// F.TEST関数（F検定）
export const F_TEST: CustomFormula = {
  name: 'F.TEST',
  pattern: /F\.TEST\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, array1Ref, array2Ref] = matches;
    
    // データ範囲から値を取得
    const values1 = getCellRangeValues(array1Ref.trim(), context);
    const values2 = getCellRangeValues(array2Ref.trim(), context);
    
    const numbers1 = values1.filter(v => typeof v === 'number' && !isNaN(v) && isFinite(v)) as number[];
    const numbers2 = values2.filter(v => typeof v === 'number' && !isNaN(v) && isFinite(v)) as number[];
    
    if (numbers1.length < 2 || numbers2.length < 2) {
      return FormulaError.DIV0;
    }
    
    // 平均を計算
    const mean1 = numbers1.reduce((sum, n) => sum + n, 0) / numbers1.length;
    const mean2 = numbers2.reduce((sum, n) => sum + n, 0) / numbers2.length;
    
    // 分散を計算
    const var1 = numbers1.reduce((sum, n) => sum + Math.pow(n - mean1, 2), 0) / (numbers1.length - 1);
    const var2 = numbers2.reduce((sum, n) => sum + Math.pow(n - mean2, 2), 0) / (numbers2.length - 1);
    
    if (var1 === 0 || var2 === 0) {
      return FormulaError.DIV0;
    }
    
    // F統計量を計算
    let f: number;
    let df1: number;
    let df2: number;
    
    if (var1 >= var2) {
      f = var1 / var2;
      df1 = numbers1.length - 1;
      df2 = numbers2.length - 1;
    } else {
      f = var2 / var1;
      df1 = numbers2.length - 1;
      df2 = numbers1.length - 1;
    }
    
    // F分布の両側検定のp値を計算
    const p = fDistributionCDF(f, df1, df2);
    
    // 両側検定なので、小さい方の確率を2倍にする
    return 2 * Math.min(p, 1 - p);
  }
};

// CHISQ.TEST関数（カイ二乗検定）
export const CHISQ_TEST: CustomFormula = {
  name: 'CHISQ.TEST',
  pattern: /CHISQ\.TEST\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, actualRef, expectedRef] = matches;
    
    // 観測値と期待値の範囲を取得
    // parseRangeがマッチしない場合は、getCellRangeValuesを使用
    let actualRangeValues = parseRange(actualRef.trim(), context);
    let expectedRangeValues = parseRange(expectedRef.trim(), context);
    
    if (!actualRangeValues || !expectedRangeValues) {
      // 単一列のデータとして処理
      const actualValues = getCellRangeValues(actualRef.trim(), context);
      const expectedValues = getCellRangeValues(expectedRef.trim(), context);
      
      if (actualValues.length === 0 || expectedValues.length === 0) {
        return FormulaError.REF;
      }
      
      // 1列のデータを二次元配列に変換
      actualRangeValues = actualValues.map(v => [Number(v)]);
      expectedRangeValues = expectedValues.map(v => [Number(v)]);
    }
    
    // 2次元配列として処理
    const actualRows = actualRangeValues.length;
    const actualCols = actualRangeValues[0]?.length || 0;
    const expectedRows = expectedRangeValues.length;
    const expectedCols = expectedRangeValues[0]?.length || 0;
    
    if (actualRows !== expectedRows || actualCols !== expectedCols) {
      return FormulaError.NA;
    }
    
    // カイ二乗統計量を計算
    let chiSquare = 0;
    let validCells = 0;
    
    for (let i = 0; i < actualRows; i++) {
      for (let j = 0; j < actualCols; j++) {
        const actual = Number(actualRangeValues[i][j]);
        const expected = Number(expectedRangeValues[i][j]);
        
        if (!isNaN(actual) && !isNaN(expected) && expected > 0) {
          chiSquare += Math.pow(actual - expected, 2) / expected;
          validCells++;
        }
      }
    }
    
    if (validCells === 0) {
      return FormulaError.DIV0;
    }
    
    // 自由度の計算
    // 1次元データの場合は、データ数-1
    // 2次元データの場合は、(行数-1)×(列数-1)
    let df: number;
    if (actualCols === 1 && actualRows > 1) {
      // 1列のデータ
      df = actualRows - 1;
    } else if (actualRows === 1 && actualCols > 1) {
      // 1行のデータ
      df = actualCols - 1;
    } else {
      // 2次元データ
      df = (actualRows - 1) * (actualCols - 1);
    }
    
    if (df <= 0) {
      return FormulaError.NUM;
    }
    
    // カイ二乗分布の上側確率を計算（p値）
    return 1 - chiSquareCDF(chiSquare, df);
  }
};

// 内部関数をエクスポート（他のモジュールで使用するため）
export { incompleteBeta, gamma, logGamma };