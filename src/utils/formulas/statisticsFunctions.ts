// 統計関数の実装

import type { CustomFormula, CellData, FormulaContext } from './types';
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

// 配列から数値のみを抽出するヘルパー関数（下位互換のため保持）
function extractNumbers(data: CellData[][]): number[] {
  const numbers: number[] = [];
  data.forEach(row => {
    row.forEach(cell => {
      if (!cell) return;
      let value;
      if (typeof cell === 'object' && 'value' in cell) {
        value = typeof cell.value === 'string' ? parseFloat(cell.value) : cell.value;
      } else {
        value = typeof cell === 'string' ? parseFloat(String(cell)) : Number(cell);
      }
      if (typeof value === 'number' && !isNaN(value)) {
        numbers.push(value);
      }
    });
  });
  return numbers;
}

// MEDIAN関数の実装（中央値）
export const MEDIAN: CustomFormula = {
  name: 'MEDIAN',
  pattern: /MEDIAN\(([^)]+)\)/i,
  isSupported: false,
  calculate: (matches, context) => {
    const rangeRef = matches[1].trim();
    const numbers = extractNumbersFromRange(rangeRef, context);
    
    if (numbers.length === 0) return FormulaError.NUM;
    
    const sorted = numbers.sort((a, b) => a - b);
    const middle = Math.floor(sorted.length / 2);
    
    if (sorted.length % 2 === 0) {
      return (sorted[middle - 1] + sorted[middle]) / 2;
    } else {
      return sorted[middle];
    }
  }
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
  isSupported: false,
  calculate: (matches, context) => {
    const rangeRef = matches[1].trim();
    const numbers = extractNumbersFromRange(rangeRef, context);
    
    if (numbers.length < 2) return FormulaError.DIV0;
    
    const mean = numbers.reduce((sum, num) => sum + num, 0) / numbers.length;
    const squaredDiffs = numbers.map(num => Math.pow(num - mean, 2));
    const variance = squaredDiffs.reduce((sum, diff) => sum + diff, 0) / (numbers.length - 1);
    
    return Math.sqrt(variance);
  }
};

// VAR関数の実装（分散・標本）
export const VAR: CustomFormula = {
  name: 'VAR',
  pattern: /VAR\(([^)]+)\)/i,
  isSupported: false,
  calculate: (matches, context) => {
    const rangeRef = matches[1].trim();
    const numbers = extractNumbersFromRange(rangeRef, context);
    
    if (numbers.length < 2) return FormulaError.DIV0;
    
    const mean = numbers.reduce((sum, num) => sum + num, 0) / numbers.length;
    const squaredDiffs = numbers.map(num => Math.pow(num - mean, 2));
    const variance = squaredDiffs.reduce((sum, diff) => sum + diff, 0) / (numbers.length - 1);
    
    return variance;
  }
};

// LARGE関数の実装（k番目に大きい値）
export const LARGE: CustomFormula = {
  name: 'LARGE',
  pattern: /LARGE\(([^,]+),\s*([^)]+)\)/i,
  isSupported: false,
  calculate: (matches, context) => {
    const rangeRef = matches[1].trim();
    const kRef = matches[2].trim();
    
    let k: number;
    if (kRef.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(kRef, context);
      k = parseInt(String(cellValue ?? '1'));
    } else {
      k = parseInt(kRef);
    }
    
    if (isNaN(k) || k < 1) return FormulaError.NUM;
    
    const numbers = extractNumbersFromRange(rangeRef, context);
    if (numbers.length === 0) return FormulaError.NUM;
    if (k > numbers.length) return FormulaError.NUM;
    
    const sorted = numbers.sort((a, b) => b - a);
    return sorted[k - 1];
  }
};

// SMALL関数の実装（k番目に小さい値）
export const SMALL: CustomFormula = {
  name: 'SMALL',
  pattern: /SMALL\(([^,]+),\s*([^)]+)\)/i,
  isSupported: false,
  calculate: (matches, context) => {
    const rangeRef = matches[1].trim();
    const kRef = matches[2].trim();
    
    let k: number;
    if (kRef.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(kRef, context);
      k = parseInt(String(cellValue ?? '1'));
    } else {
      k = parseInt(kRef);
    }
    
    if (isNaN(k) || k < 1) return FormulaError.NUM;
    
    const numbers = extractNumbersFromRange(rangeRef, context);
    if (numbers.length === 0) return FormulaError.NUM;
    if (k > numbers.length) return FormulaError.NUM;
    
    const sorted = numbers.sort((a, b) => a - b);
    return sorted[k - 1];
  }
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