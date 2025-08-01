// 高度な統計関数の実装

import type { CustomFormula, FormulaContext } from '../shared/types';
import { FormulaError } from '../shared/types';
import { getCellValue, getCellRangeValues } from '../shared/utils';

// セル範囲から数値のみを抽出するヘルパー関数
function extractNumbersFromRange(rangeRef: string, context: FormulaContext): number[] {
  const numbers: number[] = [];
  
  if (rangeRef.includes(':')) {
    const values = getCellRangeValues(rangeRef, context);
    values.forEach(value => {
      // Skip null, undefined, empty string
      if (value === null || value === undefined || value === '') {
        return;
      }
      
      const num = typeof value === 'string' ? parseFloat(value) : Number(value);
      if (!isNaN(num) && isFinite(num)) {
        numbers.push(num);
      }
    });
  } else if (rangeRef.match(/^[A-Z]+\d+$/)) {
    const cellValue = getCellValue(rangeRef, context);
    // Skip null, undefined, empty string
    if (cellValue === null || cellValue === undefined || cellValue === '') {
      return numbers;
    }
    
    const num = typeof cellValue === 'string' ? parseFloat(cellValue) : Number(cellValue);
    if (!isNaN(num) && isFinite(num)) {
      numbers.push(num);
    }
  } else {
    const num = parseFloat(rangeRef);
    if (!isNaN(num) && isFinite(num)) {
      numbers.push(num);
    }
  }
  
  return numbers;
}

// COVARIANCE.P関数（母共分散）
export const COVARIANCE_P: CustomFormula = {
  name: 'COVARIANCE.P',
  pattern: /COVARIANCE\.P\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, array1Ref, array2Ref] = matches;
    
    const numbers1 = extractNumbersFromRange(array1Ref.trim(), context);
    const numbers2 = extractNumbersFromRange(array2Ref.trim(), context);
    
    // Check for empty arrays first
    if (numbers1.length === 0 || numbers2.length === 0) {
      return FormulaError.NA;
    }
    
    // Check for different array lengths
    if (numbers1.length !== numbers2.length) {
      return FormulaError.NA;
    }
    
    const mean1 = numbers1.reduce((sum, num) => sum + num, 0) / numbers1.length;
    const mean2 = numbers2.reduce((sum, num) => sum + num, 0) / numbers2.length;
    
    const covariance = numbers1.reduce((sum, num, i) => {
      return sum + (num - mean1) * (numbers2[i] - mean2);
    }, 0) / numbers1.length;
    
    return covariance;
  }
};

// COVARIANCE.S関数（標本共分散）
export const COVARIANCE_S: CustomFormula = {
  name: 'COVARIANCE.S',
  pattern: /COVARIANCE\.S\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, array1Ref, array2Ref] = matches;
    
    const numbers1 = extractNumbersFromRange(array1Ref.trim(), context);
    const numbers2 = extractNumbersFromRange(array2Ref.trim(), context);
    
    if (numbers1.length !== numbers2.length || numbers1.length < 2) {
      return FormulaError.NA;
    }
    
    const mean1 = numbers1.reduce((sum, num) => sum + num, 0) / numbers1.length;
    const mean2 = numbers2.reduce((sum, num) => sum + num, 0) / numbers2.length;
    
    const covariance = numbers1.reduce((sum, num, i) => {
      return sum + (num - mean1) * (numbers2[i] - mean2);
    }, 0) / (numbers1.length - 1);
    
    return covariance;
  }
};

// CORREL関数（相関係数）
export const CORREL: CustomFormula = {
  name: 'CORREL',
  pattern: /CORREL\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, array1Ref, array2Ref] = matches;
    
    const numbers1 = extractNumbersFromRange(array1Ref.trim(), context);
    const numbers2 = extractNumbersFromRange(array2Ref.trim(), context);
    
    if (numbers1.length !== numbers2.length || numbers1.length < 2) {
      return FormulaError.NA;
    }
    
    const mean1 = numbers1.reduce((sum, num) => sum + num, 0) / numbers1.length;
    const mean2 = numbers2.reduce((sum, num) => sum + num, 0) / numbers2.length;
    
    let numerator = 0;
    let sumSq1 = 0;
    let sumSq2 = 0;
    
    for (let i = 0; i < numbers1.length; i++) {
      const diff1 = numbers1[i] - mean1;
      const diff2 = numbers2[i] - mean2;
      numerator += diff1 * diff2;
      sumSq1 += diff1 * diff1;
      sumSq2 += diff2 * diff2;
    }
    
    const denominator = Math.sqrt(sumSq1 * sumSq2);
    if (denominator === 0) {
      return FormulaError.DIV0;
    }
    
    return numerator / denominator;
  }
};

// PEARSON関数（ピアソン相関係数）
export const PEARSON: CustomFormula = {
  name: 'PEARSON',
  pattern: /PEARSON\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    // PEARSON関数はCORREL関数と同じ計算
    if (CORREL.calculate) {
      return CORREL.calculate(matches, context);
    }
    return FormulaError.VALUE;
  }
};

// DEVSQ関数（偏差平方和）
export const DEVSQ: CustomFormula = {
  name: 'DEVSQ',
  pattern: /DEVSQ\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, args] = matches;
    const parts = args.split(',').map(part => part.trim());
    let numbers: number[] = [];
    
    for (const part of parts) {
      numbers = numbers.concat(extractNumbersFromRange(part, context));
    }
    
    // Check for empty array - should return NUM error
    if (numbers.length === 0) {
      return FormulaError.NUM;
    }
    
    const mean = numbers.reduce((sum, num) => sum + num, 0) / numbers.length;
    return numbers.reduce((sum, num) => sum + Math.pow(num - mean, 2), 0);
  }
};

// KURT関数（尖度）
export const KURT: CustomFormula = {
  name: 'KURT',
  pattern: /KURT\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, args] = matches;
    const parts = args.split(',').map(part => part.trim());
    let numbers: number[] = [];
    
    for (const part of parts) {
      numbers = numbers.concat(extractNumbersFromRange(part, context));
    }
    
    if (numbers.length < 4) {
      return FormulaError.DIV0;
    }
    
    const n = numbers.length;
    const mean = numbers.reduce((sum, num) => sum + num, 0) / n;
    
    // Excel の KURT 関数は標本の尖度を計算（n-1で正規化）
    const variance = numbers.reduce((sum, num) => sum + Math.pow(num - mean, 2), 0) / (n - 1);
    
    if (variance === 0) {
      return FormulaError.DIV0;
    }
    
    // 標準偏差
    const standardDev = Math.sqrt(variance);
    
    // 4次中心モーメントを計算
    const fourthMoment = numbers.reduce((sum, num) => {
      return sum + Math.pow(num - mean, 4);
    }, 0);
    
    // Excel の KURT 公式（過剰尖度）
    // KURT = [n(n+1) / ((n-1)(n-2)(n-3))] * [Σ(xi-mean)^4 / s^4] - [3(n-1)^2 / ((n-2)(n-3))]
    const factor1 = (n * (n + 1)) / ((n - 1) * (n - 2) * (n - 3));
    const normalizedFourthMoment = fourthMoment / Math.pow(standardDev, 4);
    const factor2 = (3 * Math.pow(n - 1, 2)) / ((n - 2) * (n - 3));
    
    return factor1 * normalizedFourthMoment - factor2;
  }
};

// SKEW関数（歪度）
export const SKEW: CustomFormula = {
  name: 'SKEW',
  pattern: /SKEW\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, args] = matches;
    const parts = args.split(',').map(part => part.trim());
    let numbers: number[] = [];
    
    for (const part of parts) {
      numbers = numbers.concat(extractNumbersFromRange(part, context));
    }
    
    if (numbers.length < 3) {
      return FormulaError.DIV0;
    }
    
    const n = numbers.length;
    const mean = numbers.reduce((sum, num) => sum + num, 0) / n;
    const variance = numbers.reduce((sum, num) => sum + Math.pow(num - mean, 2), 0) / (n - 1);
    
    if (variance === 0) {
      return FormulaError.DIV0;
    }
    
    const standardDev = Math.sqrt(variance);
    const thirdMoment = numbers.reduce((sum, num) => {
      return sum + Math.pow((num - mean) / standardDev, 3);
    }, 0);
    
    return (n / ((n - 1) * (n - 2))) * thirdMoment;
  }
};

// SKEW.P関数（歪度、母集団）
export const SKEW_P: CustomFormula = {
  name: 'SKEW.P',
  pattern: /SKEW\.P\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, args] = matches;
    const parts = args.split(',').map(part => part.trim());
    let numbers: number[] = [];
    
    for (const part of parts) {
      numbers = numbers.concat(extractNumbersFromRange(part, context));
    }
    
    if (numbers.length === 0) {
      return FormulaError.DIV0;
    }
    
    const n = numbers.length;
    const mean = numbers.reduce((sum, num) => sum + num, 0) / n;
    const variance = numbers.reduce((sum, num) => sum + Math.pow(num - mean, 2), 0) / n;
    
    if (variance === 0) {
      return FormulaError.DIV0;
    }
    
    const standardDev = Math.sqrt(variance);
    const thirdMoment = numbers.reduce((sum, num) => {
      return sum + Math.pow((num - mean) / standardDev, 3);
    }, 0) / n;
    
    return thirdMoment;
  }
};

// STANDARDIZE関数（標準化）
export const STANDARDIZE: CustomFormula = {
  name: 'STANDARDIZE',
  pattern: /STANDARDIZE\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, xRef, meanRef, stdevRef] = matches;
    
    const x = Number(getCellValue(xRef.trim(), context) ?? xRef);
    const mean = Number(getCellValue(meanRef.trim(), context) ?? meanRef);
    const stdev = Number(getCellValue(stdevRef.trim(), context) ?? stdevRef);
    
    if (isNaN(x) || isNaN(mean) || isNaN(stdev)) {
      return FormulaError.VALUE;
    }
    
    if (stdev <= 0) {
      return FormulaError.NUM;
    }
    
    return (x - mean) / stdev;
  }
};

// PERCENTILE.INC関数（パーセンタイル、境界値含む）
export const PERCENTILE_INC: CustomFormula = {
  name: 'PERCENTILE.INC',
  pattern: /PERCENTILE\.INC\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, arrayRef, kRef] = matches;
    
    const numbers = extractNumbersFromRange(arrayRef.trim(), context);
    const k = Number(getCellValue(kRef.trim(), context) ?? kRef);
    
    if (numbers.length === 0 || isNaN(k) || k < 0 || k > 1) {
      return FormulaError.NUM;
    }
    
    numbers.sort((a, b) => a - b);
    
    if (k === 0) return numbers[0];
    if (k === 1) return numbers[numbers.length - 1];
    
    const index = k * (numbers.length - 1);
    const lower = Math.floor(index);
    const upper = Math.ceil(index);
    const weight = index - lower;
    
    return numbers[lower] * (1 - weight) + numbers[upper] * weight;
  }
};

// PERCENTILE.EXC関数（パーセンタイル、境界値除く）
export const PERCENTILE_EXC: CustomFormula = {
  name: 'PERCENTILE.EXC',
  pattern: /PERCENTILE\.EXC\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, arrayRef, kRef] = matches;
    
    const numbers = extractNumbersFromRange(arrayRef.trim(), context);
    const k = Number(getCellValue(kRef.trim(), context) ?? kRef);
    
    if (numbers.length === 0 || isNaN(k) || k <= 0 || k >= 1) {
      return FormulaError.NUM;
    }
    
    numbers.sort((a, b) => a - b);
    
    // PERCENTILE.EXC uses (n+1)*k formula
    const position = k * (numbers.length + 1);
    
    // Check if position is valid for exclusive method
    if (position < 1 || position > numbers.length) {
      return FormulaError.NUM;
    }
    
    const index = position - 1; // Convert to 0-based index
    const lower = Math.floor(index);
    const upper = Math.ceil(index);
    const weight = index - lower;
    
    if (upper >= numbers.length) {
      return numbers[lower];
    }
    
    return numbers[lower] * (1 - weight) + numbers[upper] * weight;
  }
};

// QUARTILE.INC関数（四分位数、境界値含む）
export const QUARTILE_INC: CustomFormula = {
  name: 'QUARTILE.INC',
  pattern: /QUARTILE\.INC\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, arrayRef, quartRef] = matches;
    
    const numbers = extractNumbersFromRange(arrayRef.trim(), context);
    const quart = Number(getCellValue(quartRef.trim(), context) ?? quartRef);
    
    if (numbers.length === 0 || isNaN(quart) || ![0, 1, 2, 3, 4].includes(quart)) {
      return FormulaError.NUM;
    }
    
    const k = quart / 4;
    const newMatches = ['', arrayRef, k.toString()] as RegExpMatchArray;
    newMatches.index = 0;
    newMatches.input = '';
    if (PERCENTILE_INC.calculate) {
      return PERCENTILE_INC.calculate(newMatches, context);
    }
    return FormulaError.VALUE;
  }
};

// QUARTILE.EXC関数（四分位数、境界値除く）
export const QUARTILE_EXC: CustomFormula = {
  name: 'QUARTILE.EXC',
  pattern: /QUARTILE\.EXC\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, arrayRef, quartRef] = matches;
    
    const numbers = extractNumbersFromRange(arrayRef.trim(), context);
    const quart = Number(getCellValue(quartRef.trim(), context) ?? quartRef);
    
    if (numbers.length === 0 || isNaN(quart) || ![1, 2, 3].includes(quart)) {
      return FormulaError.NUM;
    }
    
    const k = quart / 4;
    const newMatches = ['', arrayRef, k.toString()] as RegExpMatchArray;
    newMatches.index = 0;
    newMatches.input = '';
    if (PERCENTILE_EXC.calculate) {
      return PERCENTILE_EXC.calculate(newMatches, context);
    }
    return FormulaError.VALUE;
  }
};

// PERCENTRANK関数（パーセント順位）
export const PERCENTRANK: CustomFormula = {
  name: 'PERCENTRANK',
  pattern: /PERCENTRANK\(([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, arrayRef, xRef, significanceRef] = matches;
    
    const numbers = extractNumbersFromRange(arrayRef.trim(), context);
    const x = Number(getCellValue(xRef.trim(), context) ?? xRef);
    const significance = significanceRef ? Number(getCellValue(significanceRef.trim(), context) ?? significanceRef) : 3;
    
    if (numbers.length === 0 || isNaN(x) || isNaN(significance) || significance < 1) {
      return FormulaError.NUM;
    }
    
    numbers.sort((a, b) => a - b);
    
    const minVal = numbers[0];
    const maxVal = numbers[numbers.length - 1];
    
    if (x < minVal || x > maxVal) {
      return FormulaError.NA;
    }
    
    if (x === minVal) return 0;
    if (x === maxVal) return 1;
    
    // 線形補間でパーセント順位を計算
    let lowerIndex = -1;
    let upperIndex = -1;
    
    for (let i = 0; i < numbers.length; i++) {
      if (numbers[i] <= x) {
        lowerIndex = i;
      }
      if (numbers[i] >= x && upperIndex === -1) {
        upperIndex = i;
        break;
      }
    }
    
    if (lowerIndex === upperIndex) {
      return lowerIndex / (numbers.length - 1);
    }
    
    const lowerValue = numbers[lowerIndex];
    const upperValue = numbers[upperIndex];
    const weight = (x - lowerValue) / (upperValue - lowerValue);
    
    const percentRank = (lowerIndex + weight) / (numbers.length - 1);
    
    return Math.round(percentRank * Math.pow(10, significance)) / Math.pow(10, significance);
  }
};

// GAMMALN関数（ガンマ関数の自然対数）
export const GAMMALN: CustomFormula = {
  name: 'GAMMALN',
  pattern: /GAMMALN\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, xRef] = matches;
    
    const cellValue = getCellValue(xRef.trim(), context);
    const x = Number(cellValue ?? xRef);
    
    if (isNaN(x)) {
      return FormulaError.VALUE;
    }
    
    if (x <= 0) {
      return FormulaError.NUM;
    }
    
    // 整数の場合は正確な値を計算
    if (Number.isInteger(x) && x <= 30) {
      // ln(Γ(n)) = ln((n-1)!)
      let result = 0;
      for (let i = 2; i < x; i++) {
        result += Math.log(i);
      }
      return result;
    }
    
    // Stirlingの近似を使用（大きい値の場合）
    if (x >= 10) {
      return (x - 0.5) * Math.log(x) - x + 0.5 * Math.log(2 * Math.PI) + 
             1 / (12 * x) - 1 / (360 * x * x * x) + 1 / (1260 * x * x * x * x * x);
    }
    
    // Lancczosの近似を使用（小さい値の場合）
    const g = 4.7421875;
    const coefficients = [
      0.9999999999999971,
      57.15623566586292,
      -59.59796035547549,
      14.13609797474175,
      -0.49191381609762,
      3.399464998481189e-5,
      4.652362892704858e-5,
      -9.837447530487956e-5,
      1.580887032249125e-4,
      -2.102644417241049e-4,
      2.174396181152126e-4,
      -1.643181065367639e-4,
      8.441822398385274e-5,
      -2.619083840158141e-5,
      3.689918265953162e-6
    ];
    
    const ag = 0.5 * Math.log(2 * Math.PI) - Math.log(x);
    let sum = coefficients[0];
    
    for (let k = 1; k < 15; k++) {
      sum += coefficients[k] / (x + k);
    }
    
    return ag + Math.log(sum) + (x - 0.5) * Math.log(x + g + 0.5) - (x + g + 0.5);
  }
};