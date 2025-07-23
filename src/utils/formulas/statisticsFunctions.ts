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
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// MODE関数の実装（最頻値）
export const MODE: CustomFormula = {
  name: 'MODE',
  pattern: /MODE\(([^)]+)\)/i,
  isSupported: false,
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
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// COUNTBLANK関数の実装（空白セルの個数）
export const COUNTBLANK: CustomFormula = {
  name: 'COUNTBLANK',
  pattern: /COUNTBLANK\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// STDEV関数の実装（標準偏差・標本）
export const STDEV: CustomFormula = {
  name: 'STDEV',
  pattern: /STDEV\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// VAR関数の実装（分散・標本）
export const VAR: CustomFormula = {
  name: 'VAR',
  pattern: /VAR\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// LARGE関数の実装（k番目に大きい値）
export const LARGE: CustomFormula = {
  name: 'LARGE',
  pattern: /LARGE\(([^,]+),\s*([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// SMALL関数の実装（k番目に小さい値）
export const SMALL: CustomFormula = {
  name: 'SMALL',
  pattern: /SMALL\(([^,]+),\s*([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// RANK関数の実装（順位）
export const RANK: CustomFormula = {
  name: 'RANK',
  pattern: /RANK\(([^,]+),\s*([^,]+)(?:,\s*([^)]+))?\)/i,
  isSupported: false,
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
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// QUARTILE関数の実装（四分位数）
export const QUARTILE: CustomFormula = {
  name: 'QUARTILE',
  pattern: /QUARTILE\(([^,]+),\s*([^)]+)\)/i,
  isSupported: false,
  calculate: (matches, context) => {
    const rangeRef = matches[1].trim();
    const quartRef = matches[2].trim();
    
    let quart: number;
    if (quartRef.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(quartRef, context);
      quart = parseInt(String(cellValue ?? '0'));
    } else {
      quart = parseInt(quartRef);
    }
    
    if (isNaN(quart) || quart < 0 || quart > 4) return FormulaError.NUM;
    
    const numbers = extractNumbersFromRange(rangeRef, context);
    if (numbers.length === 0) return FormulaError.NUM;
    
    const sorted = numbers.sort((a, b) => a - b);
    const n = sorted.length;
    
    switch (quart) {
      case 0: 
        return sorted[0]; // 最小値
      case 1: {
        // 第1四分位数
        const q1Index = (n - 1) * 0.25;
        return sorted[Math.floor(q1Index)] + (q1Index % 1) * (sorted[Math.ceil(q1Index)] - sorted[Math.floor(q1Index)]);
      }
      case 2: {
        // 中央値
        const middle = Math.floor(n / 2);
        return n % 2 === 0 ? (sorted[middle - 1] + sorted[middle]) / 2 : sorted[middle];
      }
      case 3: {
        // 第3四分位数
        const q3Index = (n - 1) * 0.75;
        return sorted[Math.floor(q3Index)] + (q3Index % 1) * (sorted[Math.ceil(q3Index)] - sorted[Math.floor(q3Index)]);
      }
      case 4: 
        return sorted[n - 1]; // 最大値
      default: 
        return FormulaError.NUM;
    }
  }
};

// PERCENTILE関数の実装（パーセンタイル値）
export const PERCENTILE: CustomFormula = {
  name: 'PERCENTILE',
  pattern: /PERCENTILE\(([^,]+),\s*([^)]+)\)/i,
  isSupported: false,
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
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// HARMEAN関数の実装（調和平均）
export const HARMEAN: CustomFormula = {
  name: 'HARMEAN',
  pattern: /HARMEAN\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// TRIMMEAN関数の実装（トリム平均）
export const TRIMMEAN: CustomFormula = {
  name: 'TRIMMEAN',
  pattern: /TRIMMEAN\(([^,]+),\s*([^)]+)\)/i,
  isSupported: false,
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
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// MINIFS関数（条件付き最小値）
export const MINIFS: CustomFormula = {
  name: 'MINIFS',
  pattern: /MINIFS\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// STDEV.S関数（標本標準偏差）
export const STDEV_S: CustomFormula = {
  name: 'STDEV.S',
  pattern: /STDEV\.S\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// STDEV.P関数（母集団標準偏差）
export const STDEV_P: CustomFormula = {
  name: 'STDEV.P',
  pattern: /STDEV\.P\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// STDEVA関数（文字列含む標本標準偏差）
export const STDEVA: CustomFormula = {
  name: 'STDEVA',
  pattern: /STDEVA\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// STDEVPA関数（文字列含む母集団標準偏差）
export const STDEVPA: CustomFormula = {
  name: 'STDEVPA',
  pattern: /STDEVPA\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// VAR.S関数（標本分散）
export const VAR_S: CustomFormula = {
  name: 'VAR.S',
  pattern: /VAR\.S\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// VAR.P関数（母集団分散）
export const VAR_P: CustomFormula = {
  name: 'VAR.P',
  pattern: /VAR\.P\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// VARA関数（文字列含む標本分散）
export const VARA: CustomFormula = {
  name: 'VARA',
  pattern: /VARA\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// VARPA関数（文字列含む母集団分散）
export const VARPA: CustomFormula = {
  name: 'VARPA',
  pattern: /VARPA\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// AVEDEV関数（平均偏差）
export const AVEDEV: CustomFormula = {
  name: 'AVEDEV',
  pattern: /AVEDEV\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// DEVSQ関数（偏差平方和）
export const DEVSQ: CustomFormula = {
  name: 'DEVSQ',
  pattern: /DEVSQ\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// KURT関数（尖度）
export const KURT: CustomFormula = {
  name: 'KURT',
  pattern: /KURT\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// SKEW関数（歪度）
export const SKEW: CustomFormula = {
  name: 'SKEW',
  pattern: /SKEW\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// SLOPE関数（回帰直線の傾き）
export const SLOPE: CustomFormula = {
  name: 'SLOPE',
  pattern: /SLOPE\(([^,]+),\s*([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// INTERCEPT関数（回帰直線の切片）
export const INTERCEPT: CustomFormula = {
  name: 'INTERCEPT',
  pattern: /INTERCEPT\(([^,]+),\s*([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// RSQ関数（決定係数）
export const RSQ: CustomFormula = {
  name: 'RSQ',
  pattern: /RSQ\(([^,]+),\s*([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// PEARSON関数（ピアソン相関係数）
export const PEARSON: CustomFormula = {
  name: 'PEARSON',
  pattern: /PEARSON\(([^,]+),\s*([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// STEYX関数（回帰の標準誤差）
export const STEYX: CustomFormula = {
  name: 'STEYX',
  pattern: /STEYX\(([^,]+),\s*([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// FISHER関数（フィッシャー変換）
export const FISHER: CustomFormula = {
  name: 'FISHER',
  pattern: /FISHER\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// FISHERINV関数（フィッシャー変換の逆関数）
export const FISHERINV: CustomFormula = {
  name: 'FISHERINV',
  pattern: /FISHERINV\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// MODE.SNGL関数（最頻値・単一）
export const MODE_SNGL: CustomFormula = {
  name: 'MODE.SNGL',
  pattern: /MODE\.SNGL\(([^)]+)\)/i,
  isSupported: false, // 手動実装
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
  isSupported: false, // 手動実装
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
  isSupported: false, // 手動実装
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
  isSupported: false, // 手動実装
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
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// PERCENTILE.EXC関数（境界値除くパーセンタイル）
export const PERCENTILE_EXC: CustomFormula = {
  name: 'PERCENTILE.EXC',
  pattern: /PERCENTILE\.EXC\(([^,]+),\s*([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// QUARTILE.INC関数（境界値含む四分位数）
export const QUARTILE_INC: CustomFormula = {
  name: 'QUARTILE.INC',
  pattern: /QUARTILE\.INC\(([^,]+),\s*([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// QUARTILE.EXC関数（境界値除く四分位数）
export const QUARTILE_EXC: CustomFormula = {
  name: 'QUARTILE.EXC',
  pattern: /QUARTILE\.EXC\(([^,]+),\s*([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// STANDARDIZE関数（標準化）
export const STANDARDIZE: CustomFormula = {
  name: 'STANDARDIZE',
  pattern: /STANDARDIZE\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// COVARIANCE.P関数（母共分散）
export const COVARIANCE_P: CustomFormula = {
  name: 'COVARIANCE.P',
  pattern: /COVARIANCE\.P\(([^,]+),\s*([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// COVARIANCE.S関数（標本共分散）
export const COVARIANCE_S: CustomFormula = {
  name: 'COVARIANCE.S',
  pattern: /COVARIANCE\.S\(([^,]+),\s*([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// SKEW.P関数（母集団歪度）
export const SKEW_P: CustomFormula = {
  name: 'SKEW.P',
  pattern: /SKEW\.P\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// PROB関数（確率計算）- 手動実装
export const PROB: CustomFormula = {
  name: 'PROB',
  pattern: /PROB\(([^,]+),\s*([^,]+)(?:,\s*([^,]+))?(?:,\s*([^)]+))?\)/i,
  isSupported: false, // 手動実装
  calculate: () => {
    // 簡略化実装 - 基本的なエラーハンドリングのみ
    return FormulaError.VALUE;
  }
};

// PERCENTRANK関数（パーセント順位）
export const PERCENTRANK: CustomFormula = {
  name: 'PERCENTRANK',
  pattern: /PERCENTRANK\(([^,]+),\s*([^,]+)(?:,\s*([^)]+))?\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// PERCENTRANK.INC関数（パーセント順位・境界値含む）
export const PERCENTRANK_INC: CustomFormula = {
  name: 'PERCENTRANK.INC',
  pattern: /PERCENTRANK\.INC\(([^,]+),\s*([^,]+)(?:,\s*([^)]+))?\)/i,
  isSupported: false, // 手動実装
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
  isSupported: false, // 手動実装
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
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// GAUSS関数（ガウス関数）
export const GAUSS: CustomFormula = {
  name: 'GAUSS',
  pattern: /GAUSS\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};