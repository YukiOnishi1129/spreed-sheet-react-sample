// 分布関数の実装

import type { CustomFormula, FormulaContext } from '../shared/types';
import { FormulaError } from '../shared/types';
import { getCellValue } from '../shared/utils';

// 数学定数とヘルパー関数
const SQRT_2PI = Math.sqrt(2 * Math.PI);
const SQRT_2 = Math.sqrt(2);

// ガンマ関数の近似（Lanczos近似）
function gamma(z: number): number {
  if (z < 0.5) {
    return Math.PI / (Math.sin(Math.PI * z) * gamma(1 - z));
  }
  
  z -= 1;
  const x = 0.99999999999980993;
  const coefficients = [
    676.5203681218851, -1259.1392167224028, 771.32342877765313,
    -176.61502916214059, 12.507343278686905, -0.13857109526572012,
    9.9843695780195716e-6, 1.5056327351493116e-7
  ];
  
  let result = x;
  for (let i = 0; i < coefficients.length; i++) {
    result += coefficients[i] / (z + i + 1);
  }
  
  const t = z + coefficients.length - 0.5;
  return Math.sqrt(2 * Math.PI) * Math.pow(t, z + 0.5) * Math.exp(-t) * result;
}

// 標準正規分布の累積分布関数（誤差関数を使用）
function standardNormalCDF(z: number): number {
  return 0.5 * (1 + erf(z / SQRT_2));
}

// 誤差関数の近似
function erf(x: number): number {
  const a1 =  0.254829592;
  const a2 = -0.284496736;
  const a3 =  1.421413741;
  const a4 = -1.453152027;
  const a5 =  1.061405429;
  const p  =  0.3275911;

  const sign = x >= 0 ? 1 : -1;
  x = Math.abs(x);

  const t = 1.0 / (1.0 + p * x);
  const y = 1.0 - (((((a5 * t + a4) * t) + a3) * t + a2) * t + a1) * t * Math.exp(-x * x);

  return sign * y;
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

// 不完全ガンマ関数の近似
function incompleteGamma(a: number, x: number): number {
  if (x === 0) return 0;
  if (x < 0) return NaN;
  
  // 級数展開による近似
  let sum = 1;
  let term = 1;
  
  for (let n = 1; n < 100; n++) {
    term *= x / (a + n - 1);
    sum += term;
    if (Math.abs(term) < 1e-15) break;
  }
  
  return Math.pow(x, a) * Math.exp(-x) * sum;
}

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

// 不完全ベータ関数の近似
function incompleteBeta(x: number, a: number, b: number): number {
  if (x === 0) return 0;
  if (x === 1) return 1;
  
  // 連続分数による近似
  const lnBeta = Math.log(beta(a, b));
  let result = Math.exp(a * Math.log(x) + b * Math.log(1 - x) - lnBeta) / a;
  
  let term = result;
  for (let m = 1; m < 100; m++) {
    const numerator = m * (b - m) * x;
    const denominator = (a + 2 * m - 1) * (a + 2 * m);
    term *= numerator / denominator;
    result += term;
    
    if (Math.abs(term) < 1e-15) break;
  }
  
  return result;
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
      const t = (x + Math.sqrt(x * x + degFreedom)) / (2 * Math.sqrt(x * x + degFreedom));
      return incompleteBeta(t, 0.5, degFreedom / 2);
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
    
    const x = Math.abs(Number(getCellValue(xRef.trim(), context) ?? xRef));
    const degFreedom = Number(getCellValue(degFreedomRef.trim(), context) ?? degFreedomRef);
    
    if (isNaN(x) || isNaN(degFreedom)) {
      return FormulaError.VALUE;
    }
    
    if (degFreedom < 1 || !Number.isInteger(degFreedom)) {
      return FormulaError.NUM;
    }
    
    // 両側確率 = 2 * (1 - 左側確率)
    const t = (x + Math.sqrt(x * x + degFreedom)) / (2 * Math.sqrt(x * x + degFreedom));
    const leftTail = incompleteBeta(t, 0.5, degFreedom / 2);
    return 2 * (1 - leftTail);
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
    
    // 右側確率 = 1 - 左側確率
    const t = (x + Math.sqrt(x * x + degFreedom)) / (2 * Math.sqrt(x * x + degFreedom));
    const leftTail = incompleteBeta(t, 0.5, degFreedom / 2);
    return 1 - leftTail;
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
    const [, failuresRef, successesRef, probSRef, cumulativeRef] = matches;
    
    const failures = Number(getCellValue(failuresRef.trim(), context) ?? failuresRef);
    const successes = Number(getCellValue(successesRef.trim(), context) ?? successesRef);
    const probS = Number(getCellValue(probSRef.trim(), context) ?? probSRef);
    const cumulative = getCellValue(cumulativeRef.trim(), context) ?? cumulativeRef;
    
    if (isNaN(failures) || isNaN(successes) || isNaN(probS)) {
      return FormulaError.VALUE;
    }
    
    if (failures < 0 || successes < 1 || !Number.isInteger(failures) || !Number.isInteger(successes) ||
        probS <= 0 || probS >= 1) {
      return FormulaError.NUM;
    }
    
    const isCumulative = cumulative.toString().toLowerCase() === 'true' || cumulative === 1;
    
    if (isCumulative) {
      // 累積分布関数
      let result = 0;
      for (let k = 0; k <= failures; k++) {
        const binomCoeff = factorial(successes + k - 1) / (factorial(k) * factorial(successes - 1));
        result += binomCoeff * Math.pow(probS, successes) * Math.pow(1 - probS, k);
      }
      return result;
    } else {
      // 確率質量関数
      const binomCoeff = factorial(successes + failures - 1) / (factorial(failures) * factorial(successes - 1));
      return binomCoeff * Math.pow(probS, successes) * Math.pow(1 - probS, failures);
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
    
    // 組み合わせ計算のヘルパー
    const combination = (n: number, k: number): number => {
      if (k > n || k < 0) return 0;
      if (k === 0 || k === n) return 1;
      
      return factorial(n) / (factorial(k) * factorial(n - k));
    };
    
    if (isCumulative) {
      // 累積分布関数
      let result = 0;
      for (let k = Math.max(0, sampleN - (popN - popS)); k <= Math.min(sampleS, sampleN); k++) {
        const numerator = combination(popS, k) * combination(popN - popS, sampleN - k);
        const denominator = combination(popN, sampleN);
        result += numerator / denominator;
      }
      return result;
    } else {
      // 確率質量関数
      const numerator = combination(popS, sampleS) * combination(popN - popS, sampleN - sampleS);
      const denominator = combination(popN, sampleN);
      return numerator / denominator;
    }
  }
};


// === 検定・推定関数 ===

// CONFIDENCE.NORM関数（信頼区間・正規分布）
export const CONFIDENCE_NORM: CustomFormula = {
  name: 'CONFIDENCE.NORM',
  pattern: /CONFIDENCE\.NORM\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// CONFIDENCE.T関数（信頼区間・t分布）
export const CONFIDENCE_T: CustomFormula = {
  name: 'CONFIDENCE.T',
  pattern: /CONFIDENCE\.T\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// Z.TEST関数（z検定）
export const Z_TEST: CustomFormula = {
  name: 'Z.TEST',
  pattern: /Z\.TEST\(([^,]+),\s*([^,]+)(?:,\s*([^)]+))?\)/i,
  calculate: () => null // HyperFormulaが処理
};

// T.TEST関数（t検定）
export const T_TEST: CustomFormula = {
  name: 'T.TEST',
  pattern: /T\.TEST\(([^,]+),\s*([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// F.TEST関数（F検定）
export const F_TEST: CustomFormula = {
  name: 'F.TEST',
  pattern: /F\.TEST\(([^,]+),\s*([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// CHISQ.TEST関数（カイ二乗検定）
export const CHISQ_TEST: CustomFormula = {
  name: 'CHISQ.TEST',
  pattern: /CHISQ\.TEST\(([^,]+),\s*([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};