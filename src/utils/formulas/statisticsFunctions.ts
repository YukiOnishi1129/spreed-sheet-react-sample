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
  isSupported: false,
  calculate: (matches, context) => {
    const rangeRef = matches[1].trim();
    let count = 0;
    
    if (rangeRef.includes(':')) {
      // セル範囲の場合
      const values = getCellRangeValues(rangeRef, context);
      values.forEach(value => {
        if (value !== null && value !== undefined && value !== '') {
          count++;
        }
      });
    } else if (rangeRef.match(/^[A-Z]+\d+$/)) {
      // 単一セル参照の場合
      const cellValue = getCellValue(rangeRef, context);
      if (cellValue !== null && cellValue !== undefined && cellValue !== '') {
        count = 1;
      }
    }
    
    return count;
  }
};

// COUNTBLANK関数の実装（空白セルの個数）
export const COUNTBLANK: CustomFormula = {
  name: 'COUNTBLANK',
  pattern: /COUNTBLANK\(([^)]+)\)/i,
  isSupported: false,
  calculate: (matches, context) => {
    const rangeRef = matches[1].trim();
    let count = 0;
    
    if (rangeRef.includes(':')) {
      // セル範囲の場合
      const values = getCellRangeValues(rangeRef, context);
      values.forEach(value => {
        if (value === null || value === undefined || value === '') {
          count++;
        }
      });
    } else if (rangeRef.match(/^[A-Z]+\d+$/)) {
      // 単一セル参照の場合
      const cellValue = getCellValue(rangeRef, context);
      if (cellValue === null || cellValue === undefined || cellValue === '') {
        count = 1;
      }
    }
    
    return count;
  }
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
      case 0: return sorted[0]; // 最小値
      case 1: // 第1四分位数
        const q1Index = (n - 1) * 0.25;
        return sorted[Math.floor(q1Index)] + (q1Index % 1) * (sorted[Math.ceil(q1Index)] - sorted[Math.floor(q1Index)]);
      case 2: // 中央値
        const middle = Math.floor(n / 2);
        return n % 2 === 0 ? (sorted[middle - 1] + sorted[middle]) / 2 : sorted[middle];
      case 3: // 第3四分位数
        const q3Index = (n - 1) * 0.75;
        return sorted[Math.floor(q3Index)] + (q3Index % 1) * (sorted[Math.ceil(q3Index)] - sorted[Math.floor(q3Index)]);
      case 4: return sorted[n - 1]; // 最大値
      default: return FormulaError.NUM;
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