// 統計関数の実装

import type { CustomFormula, FormulaContext } from './types';
import { FormulaError } from './types';
import { getCellValue, getCellRangeValues } from './utils';

// セル範囲から数値のみを抽出するヘルパー関数（改良版）
function extractNumbersFromRange(rangeRef: string, context: FormulaContext): number[] {
  const numbers: number[] = [];
  
  if (rangeRef.includes(':')) {
    // セル範囲の場合
    const values = getCellRangeValues(rangeRef, context);
    values.forEach(value => {
      const num = typeof value === 'string' ? parseFloat(value) : Number(value);
      if (!isNaN(num) && isFinite(num)) {
        numbers.push(num);
      }
    });
  } else if (rangeRef.match(/^[A-Z]+\d+$/)) {
    // 単一セル参照の場合
    const cellValue = getCellValue(rangeRef, context);
    const num = typeof cellValue === 'string' ? parseFloat(cellValue) : Number(cellValue);
    if (!isNaN(num) && isFinite(num)) {
      numbers.push(num);
    }
  } else {
    // 直接値の場合
    const num = parseFloat(rangeRef);
    if (!isNaN(num) && isFinite(num)) {
      numbers.push(num);
    }
  }
  
  return numbers;
}


// MEDIAN関数の実装（中央値）
export const MEDIAN: CustomFormula = {
  name: 'MEDIAN',
  pattern: /MEDIAN\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, args] = matches;
    const parts = args.split(',').map(part => part.trim());
    let numbers: number[] = [];
    
    for (const part of parts) {
      numbers = numbers.concat(extractNumbersFromRange(part, context));
    }
    
    if (numbers.length === 0) {
      return FormulaError.NUM;
    }
    
    numbers.sort((a, b) => a - b);
    const mid = Math.floor(numbers.length / 2);
    
    if (numbers.length % 2 === 0) {
      return (numbers[mid - 1] + numbers[mid]) / 2;
    } else {
      return numbers[mid];
    }
  }
};

// MODE関数の実装（最頻値）
export const MODE: CustomFormula = {
  name: 'MODE',
  pattern: /MODE\(([^)]+)\)/i,
  calculate: (matches, context) => {
    const rangeRef = matches[1].trim();
    const numbers = extractNumbersFromRange(rangeRef, context);
    
    if (numbers.length === 0) return FormulaError.NA;
    
    const frequency: { [key: number]: number } = {};
    numbers.forEach(num => {
      frequency[num] = (frequency[num] || 0) + 1;
    });
    
    let maxCount = 0;
    let mode = null;
    
    for (const [num, count] of Object.entries(frequency)) {
      if (count > maxCount) {
        maxCount = count;
        mode = parseFloat(num);
      }
    }
    
    if (maxCount < 2) return FormulaError.NA;
    return mode;
  }
};

// COUNTA関数の実装（空白以外のセル数）
export const COUNTA: CustomFormula = {
  name: 'COUNTA',
  pattern: /COUNTA\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, args] = matches;
    const parts = args.split(',').map(part => part.trim());
    let count = 0;
    
    for (const part of parts) {
      if (part.includes(':')) {
        // 範囲指定の場合
        const values = getCellRangeValues(part, context);
        for (const value of values) {
          if (value !== null && value !== undefined && value !== '') {
            count++;
          }
        }
      } else if (part.match(/^[A-Z]+\d+$/)) {
        // 単一セル参照の場合
        const cellValue = getCellValue(part, context);
        if (cellValue !== null && cellValue !== undefined && cellValue !== '') {
          count++;
        }
      } else {
        // 直接値の場合
        if (part !== '' && part !== 'null' && part !== 'undefined') {
          count++;
        }
      }
    }
    
    return count;
  }
};

// COUNTBLANK関数の実装（空白セルの個数）
export const COUNTBLANK: CustomFormula = {
  name: 'COUNTBLANK',
  pattern: /COUNTBLANK\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, rangeRef] = matches;
    
    if (rangeRef.includes(':')) {
      // 範囲指定の場合
      const values = getCellRangeValues(rangeRef, context);
      let blankCount = 0;
      
      for (const value of values) {
        if (value === null || value === undefined || value === '') {
          blankCount++;
        }
      }
      
      return blankCount;
    } else if (rangeRef.match(/^[A-Z]+\d+$/)) {
      // 単一セル参照の場合
      const cellValue = getCellValue(rangeRef, context);
      return (cellValue === null || cellValue === undefined || cellValue === '') ? 1 : 0;
    }
    
    return 0;
  }
};

// STDEV関数の実装（標準偏差・標本）
export const STDEV: CustomFormula = {
  name: 'STDEV',
  pattern: /STDEV\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, args] = matches;
    const parts = args.split(',').map(part => part.trim());
    let numbers: number[] = [];
    
    for (const part of parts) {
      numbers = numbers.concat(extractNumbersFromRange(part, context));
    }
    
    if (numbers.length < 2) {
      return FormulaError.DIV0;
    }
    
    // 標本標準偏差（n-1で割る）
    const mean = numbers.reduce((sum, num) => sum + num, 0) / numbers.length;
    const variance = numbers.reduce((sum, num) => sum + Math.pow(num - mean, 2), 0) / (numbers.length - 1);
    
    return Math.sqrt(variance);
  }
};

// VAR関数の実装（分散・標本）
export const VAR: CustomFormula = {
  name: 'VAR',
  pattern: /VAR\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, args] = matches;
    const parts = args.split(',').map(part => part.trim());
    let numbers: number[] = [];
    
    for (const part of parts) {
      numbers = numbers.concat(extractNumbersFromRange(part, context));
    }
    
    if (numbers.length < 2) {
      return FormulaError.DIV0;
    }
    
    // 標本分散（n-1で割る）
    const mean = numbers.reduce((sum, num) => sum + num, 0) / numbers.length;
    const variance = numbers.reduce((sum, num) => sum + Math.pow(num - mean, 2), 0) / (numbers.length - 1);
    
    return variance;
  }
};

// LARGE関数の実装（k番目に大きい値）
export const LARGE: CustomFormula = {
  name: 'LARGE',
  pattern: /LARGE\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, rangeRef, kRef] = matches;
    
    const numbers = extractNumbersFromRange(rangeRef, context);
    
    let k: number;
    if (kRef.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(kRef, context);
      k = parseInt(String(cellValue ?? '1'));
    } else {
      k = parseInt(kRef);
    }
    
    if (isNaN(k) || k < 1 || k > numbers.length) {
      return FormulaError.NUM;
    }
    
    // 降順にソート
    numbers.sort((a, b) => b - a);
    
    return numbers[k - 1];
  }
};

// SMALL関数の実装（k番目に小さい値）
export const SMALL: CustomFormula = {
  name: 'SMALL',
  pattern: /SMALL\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, rangeRef, kRef] = matches;
    
    const numbers = extractNumbersFromRange(rangeRef, context);
    
    let k: number;
    if (kRef.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(kRef, context);
      k = parseInt(String(cellValue ?? '1'));
    } else {
      k = parseInt(kRef);
    }
    
    if (isNaN(k) || k < 1 || k > numbers.length) {
      return FormulaError.NUM;
    }
    
    // 昇順にソート
    numbers.sort((a, b) => a - b);
    
    return numbers[k - 1];
  }
};

// RANK関数の実装（順位）
export const RANK: CustomFormula = {
  name: 'RANK',
  pattern: /RANK\(([^,]+),\s*([^,]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches, context) => {
    const valueRef = matches[1].trim();
    const rangeRef = matches[2].trim();
    const orderRef = matches[3] ? matches[3].trim() : '0';
    
    let value: number, order: number;
    
    if (valueRef.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(valueRef, context);
      value = parseFloat(String(cellValue ?? '0'));
    } else {
      value = parseFloat(valueRef);
    }
    
    if (orderRef.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(orderRef, context);
      order = parseInt(String(cellValue ?? '0'));
    } else {
      order = parseInt(orderRef);
    }
    
    if (isNaN(value)) return FormulaError.VALUE;
    
    const numbers = extractNumbersFromRange(rangeRef, context);
    if (numbers.length === 0) return FormulaError.NA;
    
    const sorted = order === 1 
      ? numbers.sort((a, b) => a - b) 
      : numbers.sort((a, b) => b - a);
    
    const rank = sorted.indexOf(value) + 1;
    return rank === 0 ? FormulaError.NA : rank;
  }
};

// CORREL関数の実装（相関係数）
export const CORREL: CustomFormula = {
  name: 'CORREL',
  pattern: /CORREL\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, array1Ref, array2Ref] = matches;
    
    const numbers1 = extractNumbersFromRange(array1Ref, context);
    const numbers2 = extractNumbersFromRange(array2Ref, context);
    
    if (numbers1.length !== numbers2.length || numbers1.length < 2) {
      return FormulaError.NA;
    }
    
    const n = numbers1.length;
    const mean1 = numbers1.reduce((sum, num) => sum + num, 0) / n;
    const mean2 = numbers2.reduce((sum, num) => sum + num, 0) / n;
    
    let numerator = 0;
    let sum1Squared = 0;
    let sum2Squared = 0;
    
    for (let i = 0; i < n; i++) {
      const diff1 = numbers1[i] - mean1;
      const diff2 = numbers2[i] - mean2;
      numerator += diff1 * diff2;
      sum1Squared += diff1 * diff1;
      sum2Squared += diff2 * diff2;
    }
    
    const denominator = Math.sqrt(sum1Squared * sum2Squared);
    
    if (denominator === 0) {
      return FormulaError.DIV0;
    }
    
    return numerator / denominator;
  }
};

// QUARTILE関数の実装（四分位数）
export const QUARTILE: CustomFormula = {
  name: 'QUARTILE',
  pattern: /QUARTILE\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, rangeRef, quartRef] = matches;
    
    let quart: number;
    if (quartRef.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(quartRef, context);
      quart = parseInt(String(cellValue ?? '0'));
    } else {
      quart = parseInt(quartRef);
    }
    
    if (isNaN(quart) || quart < 0 || quart > 4) {
      return FormulaError.NUM;
    }
    
    const numbers = extractNumbersFromRange(rangeRef, context);
    if (numbers.length === 0) {
      return FormulaError.NUM;
    }
    
    numbers.sort((a, b) => a - b);
    
    switch (quart) {
      case 0: // 最小値
        return numbers[0];
      case 1: // 第一四分位数
        return numbers[Math.floor((numbers.length - 1) * 0.25)];
      case 2: { // 中央値
        const mid = Math.floor(numbers.length / 2);
        if (numbers.length % 2 === 0) {
          return (numbers[mid - 1] + numbers[mid]) / 2;
        } else {
          return numbers[mid];
        }
      }
      case 3: // 第三四分位数
        return numbers[Math.floor((numbers.length - 1) * 0.75)];
      case 4: // 最大値
        return numbers[numbers.length - 1];
      default:
        return FormulaError.NUM;
    }
  }
};

// PERCENTILE関数の実装（パーセンタイル値）
export const PERCENTILE: CustomFormula = {
  name: 'PERCENTILE',
  pattern: /PERCENTILE\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches, context) => {
    const rangeRef = matches[1].trim();
    const kRef = matches[2].trim();
    
    let k: number;
    if (kRef.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(kRef, context);
      k = parseFloat(String(cellValue ?? '0'));
    } else {
      k = parseFloat(kRef);
    }
    
    if (isNaN(k) || k < 0 || k > 1) return FormulaError.NUM;
    
    const numbers = extractNumbersFromRange(rangeRef, context);
    if (numbers.length === 0) return FormulaError.NUM;
    
    const sorted = numbers.sort((a, b) => a - b);
    const n = sorted.length;
    const index = (n - 1) * k;
    
    const lower = Math.floor(index);
    const upper = Math.ceil(index);
    const weight = index % 1;
    
    if (upper >= n) return sorted[n - 1];
    return sorted[lower] * (1 - weight) + sorted[upper] * weight;
  }
};

// GEOMEAN関数の実装（幾何平均）
export const GEOMEAN: CustomFormula = {
  name: 'GEOMEAN',
  pattern: /GEOMEAN\(([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// HARMEAN関数の実装（調和平均）
export const HARMEAN: CustomFormula = {
  name: 'HARMEAN',
  pattern: /HARMEAN\(([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// TRIMMEAN関数の実装（トリム平均）
export const TRIMMEAN: CustomFormula = {
  name: 'TRIMMEAN',
  pattern: /TRIMMEAN\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches, context) => {
    const rangeRef = matches[1].trim();
    const percentRef = matches[2].trim();
    
    let percent: number;
    if (percentRef.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(percentRef, context);
      percent = parseFloat(String(cellValue ?? '0'));
    } else {
      percent = parseFloat(percentRef);
    }
    
    if (isNaN(percent) || percent < 0 || percent >= 1) return FormulaError.NUM;
    
    const numbers = extractNumbersFromRange(rangeRef, context);
    if (numbers.length === 0) return FormulaError.NUM;
    
    const sorted = numbers.sort((a, b) => a - b);
    const trimCount = Math.floor(sorted.length * percent / 2);
    
    if (trimCount * 2 >= sorted.length) return FormulaError.NUM;
    
    const trimmed = sorted.slice(trimCount, sorted.length - trimCount);
    const sum = trimmed.reduce((sum, num) => sum + num, 0);
    
    return sum / trimmed.length;
  }
};

// MAXIFS関数（条件付き最大値）
export const MAXIFS: CustomFormula = {
  name: 'MAXIFS',
  pattern: /MAXIFS\(([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// MINIFS関数（条件付き最小値）
export const MINIFS: CustomFormula = {
  name: 'MINIFS',
  pattern: /MINIFS\(([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// STDEV.S関数（標本標準偏差）
export const STDEV_S: CustomFormula = {
  name: 'STDEV.S',
  pattern: /STDEV\.S\(([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// STDEV.P関数（母集団標準偏差）
export const STDEV_P: CustomFormula = {
  name: 'STDEV.P',
  pattern: /STDEV\.P\(([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// STDEVA関数（文字列含む標本標準偏差）
export const STDEVA: CustomFormula = {
  name: 'STDEVA',
  pattern: /STDEVA\(([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// STDEVPA関数（文字列含む母集団標準偏差）
export const STDEVPA: CustomFormula = {
  name: 'STDEVPA',
  pattern: /STDEVPA\(([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// VAR.S関数（標本分散）
export const VAR_S: CustomFormula = {
  name: 'VAR.S',
  pattern: /VAR\.S\(([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// VAR.P関数（母集団分散）
export const VAR_P: CustomFormula = {
  name: 'VAR.P',
  pattern: /VAR\.P\(([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// VARA関数（文字列含む標本分散）
export const VARA: CustomFormula = {
  name: 'VARA',
  pattern: /VARA\(([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// VARPA関数（文字列含む母集団分散）
export const VARPA: CustomFormula = {
  name: 'VARPA',
  pattern: /VARPA\(([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// AVEDEV関数（平均偏差）
export const AVEDEV: CustomFormula = {
  name: 'AVEDEV',
  pattern: /AVEDEV\(([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// DEVSQ関数（偏差平方和）
export const DEVSQ: CustomFormula = {
  name: 'DEVSQ',
  pattern: /DEVSQ\(([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// KURT関数（尖度）
export const KURT: CustomFormula = {
  name: 'KURT',
  pattern: /KURT\(([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// SKEW関数（歪度）
export const SKEW: CustomFormula = {
  name: 'SKEW',
  pattern: /SKEW\(([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// SLOPE関数（回帰直線の傾き）
export const SLOPE: CustomFormula = {
  name: 'SLOPE',
  pattern: /SLOPE\(([^,]+),\s*([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// INTERCEPT関数（回帰直線の切片）
export const INTERCEPT: CustomFormula = {
  name: 'INTERCEPT',
  pattern: /INTERCEPT\(([^,]+),\s*([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// RSQ関数（決定係数）
export const RSQ: CustomFormula = {
  name: 'RSQ',
  pattern: /RSQ\(([^,]+),\s*([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// PEARSON関数（ピアソン相関係数）
export const PEARSON: CustomFormula = {
  name: 'PEARSON',
  pattern: /PEARSON\(([^,]+),\s*([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// STEYX関数（回帰の標準誤差）
export const STEYX: CustomFormula = {
  name: 'STEYX',
  pattern: /STEYX\(([^,]+),\s*([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// FISHER関数（フィッシャー変換）
export const FISHER: CustomFormula = {
  name: 'FISHER',
  pattern: /FISHER\(([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// FISHERINV関数（フィッシャー変換の逆関数）
export const FISHERINV: CustomFormula = {
  name: 'FISHERINV',
  pattern: /FISHERINV\(([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// MODE.SNGL関数（最頻値・単一）
export const MODE_SNGL: CustomFormula = {
  name: 'MODE.SNGL',
  pattern: /MODE\.SNGL\(([^)]+)\)/i,
  calculate: (matches, context) => {
    const rangeRef = matches[1].trim();
    const numbers = extractNumbersFromRange(rangeRef, context);
    
    if (numbers.length === 0) return FormulaError.NA;
    
    const frequency: { [key: number]: number } = {};
    numbers.forEach(num => {
      frequency[num] = (frequency[num] || 0) + 1;
    });
    
    let maxCount = 0;
    let mode = null;
    
    for (const [num, count] of Object.entries(frequency)) {
      if (count > maxCount && count >= 2) {
        maxCount = count;
        mode = parseFloat(num);
      }
    }
    
    if (maxCount < 2) return FormulaError.NA;
    return mode;
  }
};

// MODE.MULT関数（最頻値・複数）- 簡略化実装
export const MODE_MULT: CustomFormula = {
  name: 'MODE.MULT',
  pattern: /MODE\.MULT\(([^)]+)\)/i,
  calculate: (matches, context) => {
    const rangeRef = matches[1].trim();
    const numbers = extractNumbersFromRange(rangeRef, context);
    
    if (numbers.length === 0) return FormulaError.NA;
    
    const frequency: { [key: number]: number } = {};
    numbers.forEach(num => {
      frequency[num] = (frequency[num] || 0) + 1;
    });
    
    let maxCount = 0;
    for (const count of Object.values(frequency)) {
      if (count > maxCount) maxCount = count;
    }
    
    if (maxCount < 2) return FormulaError.NA;
    
    const modes = Object.keys(frequency)
      .filter(key => frequency[parseFloat(key)] === maxCount)
      .map(key => parseFloat(key))
      .sort((a, b) => a - b);
    
    // 複数値を返すのは複雑なので、最小値を返す
    return modes[0];
  }
};

// RANK.EQ関数（順位・同順位は最小）
export const RANK_EQ: CustomFormula = {
  name: 'RANK.EQ',
  pattern: /RANK\.EQ\(([^,]+),\s*([^,]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches, context) => {
    const valueRef = matches[1].trim();
    const rangeRef = matches[2].trim();
    const orderRef = matches[3] ? matches[3].trim() : '0';
    
    let value: number, order: number;
    
    if (valueRef.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(valueRef, context);
      value = parseFloat(String(cellValue ?? '0'));
    } else {
      value = parseFloat(valueRef);
    }
    
    if (orderRef.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(orderRef, context);
      order = parseInt(String(cellValue ?? '0'));
    } else {
      order = parseInt(orderRef);
    }
    
    if (isNaN(value)) return FormulaError.VALUE;
    
    const numbers = extractNumbersFromRange(rangeRef, context);
    if (numbers.length === 0) return FormulaError.NA;
    
    const sorted = order === 1 
      ? numbers.sort((a, b) => a - b) 
      : numbers.sort((a, b) => b - a);
    
    const rank = sorted.findIndex(num => num === value) + 1;
    return rank === 0 ? FormulaError.NA : rank;
  }
};

// RANK.AVG関数（順位・同順位は平均）
export const RANK_AVG: CustomFormula = {
  name: 'RANK.AVG',
  pattern: /RANK\.AVG\(([^,]+),\s*([^,]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches, context) => {
    const valueRef = matches[1].trim();
    const rangeRef = matches[2].trim();
    const orderRef = matches[3] ? matches[3].trim() : '0';
    
    let value: number, order: number;
    
    if (valueRef.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(valueRef, context);
      value = parseFloat(String(cellValue ?? '0'));
    } else {
      value = parseFloat(valueRef);
    }
    
    if (orderRef.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(orderRef, context);
      order = parseInt(String(cellValue ?? '0'));
    } else {
      order = parseInt(orderRef);
    }
    
    if (isNaN(value)) return FormulaError.VALUE;
    
    const numbers = extractNumbersFromRange(rangeRef, context);
    if (numbers.length === 0) return FormulaError.NA;
    
    const sorted = order === 1 
      ? numbers.sort((a, b) => a - b) 
      : numbers.sort((a, b) => b - a);
    
    const indices = sorted.reduce((acc: number[], num, index) => {
      if (num === value) acc.push(index + 1);
      return acc;
    }, []);
    
    if (indices.length === 0) return FormulaError.NA;
    
    const avgRank = indices.reduce((sum, rank) => sum + rank, 0) / indices.length;
    return avgRank;
  }
};

// PERCENTILE.INC関数（境界値含むパーセンタイル）
export const PERCENTILE_INC: CustomFormula = {
  name: 'PERCENTILE.INC',
  pattern: /PERCENTILE\.INC\(([^,]+),\s*([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// PERCENTILE.EXC関数（境界値除くパーセンタイル）
export const PERCENTILE_EXC: CustomFormula = {
  name: 'PERCENTILE.EXC',
  pattern: /PERCENTILE\.EXC\(([^,]+),\s*([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// QUARTILE.INC関数（境界値含む四分位数）
export const QUARTILE_INC: CustomFormula = {
  name: 'QUARTILE.INC',
  pattern: /QUARTILE\.INC\(([^,]+),\s*([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// QUARTILE.EXC関数（境界値除く四分位数）
export const QUARTILE_EXC: CustomFormula = {
  name: 'QUARTILE.EXC',
  pattern: /QUARTILE\.EXC\(([^,]+),\s*([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// STANDARDIZE関数（標準化）
export const STANDARDIZE: CustomFormula = {
  name: 'STANDARDIZE',
  pattern: /STANDARDIZE\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// COVARIANCE.P関数（母共分散）
export const COVARIANCE_P: CustomFormula = {
  name: 'COVARIANCE.P',
  pattern: /COVARIANCE\.P\(([^,]+),\s*([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// COVARIANCE.S関数（標本共分散）
export const COVARIANCE_S: CustomFormula = {
  name: 'COVARIANCE.S',
  pattern: /COVARIANCE\.S\(([^,]+),\s*([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// SKEW.P関数（母集団歪度）
export const SKEW_P: CustomFormula = {
  name: 'SKEW.P',
  pattern: /SKEW\.P\(([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// PROB関数（確率計算）- 手動実装
export const PROB: CustomFormula = {
  name: 'PROB',
  pattern: /PROB\(([^,]+),\s*([^,]+)(?:,\s*([^,]+))?(?:,\s*([^)]+))?\)/i,
  calculate: () => {
    // 簡略化実装 - 基本的なエラーハンドリングのみ
    return FormulaError.VALUE;
  }
};

// PERCENTRANK関数（パーセント順位）
export const PERCENTRANK: CustomFormula = {
  name: 'PERCENTRANK',
  pattern: /PERCENTRANK\(([^,]+),\s*([^,]+)(?:,\s*([^)]+))?\)/i,
  calculate: () => null // HyperFormulaが処理
};

// PERCENTRANK.INC関数（パーセント順位・境界値含む）
export const PERCENTRANK_INC: CustomFormula = {
  name: 'PERCENTRANK.INC',
  pattern: /PERCENTRANK\.INC\(([^,]+),\s*([^,]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches, context) => {
    const rangeRef = matches[1].trim();
    const xRef = matches[2].trim();
    const significanceRef = matches[3] ? matches[3].trim() : '3';
    
    let x: number, significance: number;
    
    if (xRef.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(xRef, context);
      x = parseFloat(String(cellValue ?? '0'));
    } else {
      x = parseFloat(xRef);
    }
    
    if (significanceRef.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(significanceRef, context);
      significance = parseInt(String(cellValue ?? '3'));
    } else {
      significance = parseInt(significanceRef);
    }
    
    if (isNaN(x)) return FormulaError.VALUE;
    if (isNaN(significance) || significance < 1) significance = 3;
    
    const numbers = extractNumbersFromRange(rangeRef, context);
    if (numbers.length === 0) return FormulaError.NA;
    
    const sorted = numbers.sort((a, b) => a - b);
    const n = sorted.length;
    
    // xが範囲外の場合
    if (x < sorted[0] || x > sorted[n - 1]) return FormulaError.NA;
    
    // 完全一致の場合
    const exactIndex = sorted.indexOf(x);
    if (exactIndex !== -1) {
      const rank = exactIndex / (n - 1);
      return Math.round(rank * Math.pow(10, significance)) / Math.pow(10, significance);
    }
    
    // 線形補間
    let lowerIndex = 0;
    let upperIndex = n - 1;
    
    for (let i = 0; i < n - 1; i++) {
      if (sorted[i] <= x && x <= sorted[i + 1]) {
        lowerIndex = i;
        upperIndex = i + 1;
        break;
      }
    }
    
    const lowerRank = lowerIndex / (n - 1);
    const upperRank = upperIndex / (n - 1);
    const weight = (x - sorted[lowerIndex]) / (sorted[upperIndex] - sorted[lowerIndex]);
    const rank = lowerRank + weight * (upperRank - lowerRank);
    
    return Math.round(rank * Math.pow(10, significance)) / Math.pow(10, significance);
  }
};

// PERCENTRANK.EXC関数（パーセント順位・境界値除く）
export const PERCENTRANK_EXC: CustomFormula = {
  name: 'PERCENTRANK.EXC',
  pattern: /PERCENTRANK\.EXC\(([^,]+),\s*([^,]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches, context) => {
    const rangeRef = matches[1].trim();
    const xRef = matches[2].trim();
    const significanceRef = matches[3] ? matches[3].trim() : '3';
    
    let x: number, significance: number;
    
    if (xRef.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(xRef, context);
      x = parseFloat(String(cellValue ?? '0'));
    } else {
      x = parseFloat(xRef);
    }
    
    if (significanceRef.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(significanceRef, context);
      significance = parseInt(String(cellValue ?? '3'));
    } else {
      significance = parseInt(significanceRef);
    }
    
    if (isNaN(x)) return FormulaError.VALUE;
    if (isNaN(significance) || significance < 1) significance = 3;
    
    const numbers = extractNumbersFromRange(rangeRef, context);
    if (numbers.length === 0) return FormulaError.NA;
    
    const sorted = numbers.sort((a, b) => a - b);
    const n = sorted.length;
    
    // xが境界値の場合はNA
    if (x <= sorted[0] || x >= sorted[n - 1]) return FormulaError.NA;
    
    // 完全一致の場合
    const exactIndex = sorted.indexOf(x);
    if (exactIndex !== -1) {
      const rank = exactIndex / (n + 1);
      return Math.round(rank * Math.pow(10, significance)) / Math.pow(10, significance);
    }
    
    // 線形補間
    let lowerIndex = 0;
    let upperIndex = n - 1;
    
    for (let i = 0; i < n - 1; i++) {
      if (sorted[i] <= x && x <= sorted[i + 1]) {
        lowerIndex = i;
        upperIndex = i + 1;
        break;
      }
    }
    
    const lowerRank = (lowerIndex + 1) / (n + 1);
    const upperRank = (upperIndex + 1) / (n + 1);
    const weight = (x - sorted[lowerIndex]) / (sorted[upperIndex] - sorted[lowerIndex]);
    const rank = lowerRank + weight * (upperRank - lowerRank);
    
    return Math.round(rank * Math.pow(10, significance)) / Math.pow(10, significance);
  }
};

// PHI関数（標準正規分布の密度関数）
export const PHI: CustomFormula = {
  name: 'PHI',
  pattern: /PHI\(([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// GAUSS関数（ガウス関数）
export const GAUSS: CustomFormula = {
  name: 'GAUSS',
  pattern: /GAUSS\(([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};