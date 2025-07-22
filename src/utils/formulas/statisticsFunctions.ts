// 統計関数の実装

import type { CustomFormula, CellData } from './types';

// 配列から数値のみを抽出するヘルパー関数
function extractNumbers(data: CellData[][]): number[] {
  const numbers: number[] = [];
  data.forEach(row => {
    row.forEach(cell => {
      const value = typeof cell.value === 'string' ? parseFloat(cell.value) : cell.value;
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
    const numbers = extractNumbers(context.data);
    if (numbers.length === 0) return '#NUM!';
    
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
    const numbers = extractNumbers(context.data);
    if (numbers.length === 0) return '#N/A';
    
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
    
    if (maxCount < 2) return '#N/A';
    return mode;
  }
};

// COUNTA関数の実装（空白以外のセル数）
export const COUNTA: CustomFormula = {
  name: 'COUNTA',
  pattern: /COUNTA\(([^)]+)\)/i,
  isSupported: false,
  calculate: (matches, context) => {
    let count = 0;
    context.data.forEach(row => {
      row.forEach(cell => {
        if (cell.value !== null && cell.value !== undefined && cell.value !== '') {
          count++;
        }
      });
    });
    return count;
  }
};

// COUNTBLANK関数の実装（空白セルの個数）
export const COUNTBLANK: CustomFormula = {
  name: 'COUNTBLANK',
  pattern: /COUNTBLANK\(([^)]+)\)/i,
  isSupported: false,
  calculate: (matches, context) => {
    let count = 0;
    context.data.forEach(row => {
      row.forEach(cell => {
        if (cell.value === null || cell.value === undefined || cell.value === '') {
          count++;
        }
      });
    });
    return count;
  }
};

// STDEV関数の実装（標準偏差・標本）
export const STDEV: CustomFormula = {
  name: 'STDEV',
  pattern: /STDEV\(([^)]+)\)/i,
  isSupported: false,
  calculate: (matches, context) => {
    const numbers = extractNumbers(context.data);
    if (numbers.length < 2) return '#DIV/0!';
    
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
    const numbers = extractNumbers(context.data);
    if (numbers.length < 2) return '#DIV/0!';
    
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
    const k = parseInt(matches[2]);
    if (isNaN(k) || k < 1) return '#NUM!';
    
    const numbers = extractNumbers(context.data);
    if (numbers.length === 0) return '#NUM!';
    if (k > numbers.length) return '#NUM!';
    
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
    const k = parseInt(matches[2]);
    if (isNaN(k) || k < 1) return '#NUM!';
    
    const numbers = extractNumbers(context.data);
    if (numbers.length === 0) return '#NUM!';
    if (k > numbers.length) return '#NUM!';
    
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
    const value = parseFloat(matches[1]);
    const order = matches[3] ? parseInt(matches[3]) : 0; // 0=降順, 1=昇順
    
    if (isNaN(value)) return '#VALUE!';
    
    const numbers = extractNumbers(context.data);
    if (numbers.length === 0) return '#N/A';
    
    const sorted = order === 1 
      ? numbers.sort((a, b) => a - b) 
      : numbers.sort((a, b) => b - a);
    
    const rank = sorted.indexOf(value) + 1;
    return rank === 0 ? '#N/A' : rank;
  }
};