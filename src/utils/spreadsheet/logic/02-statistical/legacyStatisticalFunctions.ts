// 旧版統計関数の実装

import type { CustomFormula, FormulaContext } from '../shared/types';
import { FormulaError } from '../shared/types';
import { getCellValue, getCellRangeValues } from '../shared/utils';

// PERCENTILE関数の実装（パーセンタイル - 旧版）
export const PERCENTILE: CustomFormula = {
  name: 'PERCENTILE',
  pattern: /PERCENTILE\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, rangeRef, kRef] = matches;
    
    const values = getCellRangeValues(rangeRef, context);
    const k = Number(getCellValue(kRef, context) ?? kRef);
    
    if (isNaN(k) || k < 0 || k > 1) {
      return FormulaError.VALUE;
    }
    
    const numbers = values
      .map(v => Number(v))
      .filter(n => !isNaN(n))
      .sort((a, b) => a - b);
    
    if (numbers.length === 0) {
      return FormulaError.NUM;
    }
    
    const pos = k * (numbers.length - 1);
    const lower = Math.floor(pos);
    const upper = Math.ceil(pos);
    const weight = pos - lower;
    
    if (lower === upper) {
      return numbers[lower];
    }
    
    return numbers[lower] * (1 - weight) + numbers[upper] * weight;
  }
};

// PERCENTRANK関数の実装（パーセント順位 - 旧版）
export const PERCENTRANK: CustomFormula = {
  name: 'PERCENTRANK',
  pattern: /PERCENTRANK\(([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, rangeRef, xRef, significanceRef] = matches;
    
    const values = getCellRangeValues(rangeRef, context);
    const x = Number(getCellValue(xRef, context) ?? xRef);
    const significance = significanceRef ? Number(getCellValue(significanceRef, context) ?? significanceRef) : 3;
    
    if (isNaN(x) || isNaN(significance)) {
      return FormulaError.VALUE;
    }
    
    const numbers = values
      .map(v => Number(v))
      .filter(n => !isNaN(n))
      .sort((a, b) => a - b);
    
    if (numbers.length === 0) {
      return FormulaError.NUM;
    }
    
    // xが範囲外の場合
    if (x < numbers[0] || x > numbers[numbers.length - 1]) {
      return FormulaError.NA;
    }
    
    // xが配列内にある場合
    const exactIndex = numbers.indexOf(x);
    if (exactIndex !== -1) {
      const rank = exactIndex / (numbers.length - 1);
      return Number(rank.toFixed(significance));
    }
    
    // 補間が必要な場合
    let lowerIndex = 0;
    let upperIndex = numbers.length - 1;
    
    for (let i = 0; i < numbers.length - 1; i++) {
      if (numbers[i] <= x && numbers[i + 1] >= x) {
        lowerIndex = i;
        upperIndex = i + 1;
        break;
      }
    }
    
    const lowerValue = numbers[lowerIndex];
    const upperValue = numbers[upperIndex];
    const weight = (x - lowerValue) / (upperValue - lowerValue);
    
    const lowerRank = lowerIndex / (numbers.length - 1);
    const upperRank = upperIndex / (numbers.length - 1);
    const rank = lowerRank + weight * (upperRank - lowerRank);
    
    return Number(rank.toFixed(significance));
  }
};

// QUARTILE関数の実装（四分位数 - 旧版）
export const QUARTILE: CustomFormula = {
  name: 'QUARTILE',
  pattern: /QUARTILE\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, rangeRef, quartileRef] = matches;
    
    const values = getCellRangeValues(rangeRef, context);
    const quartile = Number(getCellValue(quartileRef, context) ?? quartileRef);
    
    if (isNaN(quartile) || quartile < 0 || quartile > 4 || quartile !== Math.floor(quartile)) {
      return FormulaError.VALUE;
    }
    
    const numbers = values
      .map(v => Number(v))
      .filter(n => !isNaN(n))
      .sort((a, b) => a - b);
    
    if (numbers.length === 0) {
      return FormulaError.NUM;
    }
    
    // PERCENTILE関数を使用して計算
    const k = quartile / 4;
    const pos = k * (numbers.length - 1);
    const lower = Math.floor(pos);
    const upper = Math.ceil(pos);
    const weight = pos - lower;
    
    if (lower === upper) {
      return numbers[lower];
    }
    
    return numbers[lower] * (1 - weight) + numbers[upper] * weight;
  }
};

// RANK関数の実装（順位 - 旧版）
export const RANK: CustomFormula = {
  name: 'RANK',
  pattern: /RANK\(([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, numberRef, refRef, orderRef] = matches;
    
    const number = Number(getCellValue(numberRef, context) ?? numberRef);
    const values = getCellRangeValues(refRef, context);
    const order = orderRef ? Number(getCellValue(orderRef, context) ?? orderRef) : 0;
    
    if (isNaN(number) || isNaN(order)) {
      return FormulaError.VALUE;
    }
    
    const numbers = values
      .map(v => Number(v))
      .filter(n => !isNaN(n));
    
    if (numbers.length === 0) {
      return FormulaError.NUM;
    }
    
    // 数値が配列内に存在しない場合
    if (!numbers.includes(number)) {
      return FormulaError.NA;
    }
    
    // 降順（デフォルト）または昇順でソート
    const sorted = [...numbers].sort((a, b) => order === 0 ? b - a : a - b);
    
    // 順位を返す（1ベース）
    return sorted.indexOf(number) + 1;
  }
};

// STDEV関数の実装（標準偏差 - 旧版）
export const STDEV: CustomFormula = {
  name: 'STDEV',
  pattern: /STDEV\((.+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, argsString] = matches;
    const args = argsString.split(',').map(arg => arg.trim());
    
    let numbers: number[] = [];
    
    for (const arg of args) {
      if (arg.includes(':')) {
        // 範囲指定の場合
        const rangeValues = getCellRangeValues(arg, context);
        const rangeNumbers = rangeValues
          .map(v => Number(v))
          .filter(n => !isNaN(n));
        numbers = numbers.concat(rangeNumbers);
      } else {
        // 単一値の場合
        const value = getCellValue(arg, context) ?? arg;
        const num = Number(value);
        if (!isNaN(num)) {
          numbers.push(num);
        }
      }
    }
    
    if (numbers.length < 2) {
      return FormulaError.DIV0;
    }
    
    // 平均を計算
    const mean = numbers.reduce((sum, n) => sum + n, 0) / numbers.length;
    
    // 分散を計算（標本分散）
    const variance = numbers.reduce((sum, n) => sum + Math.pow(n - mean, 2), 0) / (numbers.length - 1);
    
    // 標準偏差を返す
    return Math.sqrt(variance);
  }
};

// STDEVP関数の実装（標準偏差・母集団 - 旧版）
export const STDEVP: CustomFormula = {
  name: 'STDEVP',
  pattern: /STDEVP\((.+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, argsString] = matches;
    const args = argsString.split(',').map(arg => arg.trim());
    
    let numbers: number[] = [];
    
    for (const arg of args) {
      if (arg.includes(':')) {
        // 範囲指定の場合
        const rangeValues = getCellRangeValues(arg, context);
        const rangeNumbers = rangeValues
          .map(v => Number(v))
          .filter(n => !isNaN(n));
        numbers = numbers.concat(rangeNumbers);
      } else {
        // 単一値の場合
        const value = getCellValue(arg, context) ?? arg;
        const num = Number(value);
        if (!isNaN(num)) {
          numbers.push(num);
        }
      }
    }
    
    if (numbers.length === 0) {
      return FormulaError.NUM;
    }
    
    // 平均を計算
    const mean = numbers.reduce((sum, n) => sum + n, 0) / numbers.length;
    
    // 分散を計算（母集団分散）
    const variance = numbers.reduce((sum, n) => sum + Math.pow(n - mean, 2), 0) / numbers.length;
    
    // 標準偏差を返す
    return Math.sqrt(variance);
  }
};

// VAR関数の実装（分散 - 旧版）
export const VAR: CustomFormula = {
  name: 'VAR',
  pattern: /VAR\((.+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, argsString] = matches;
    const args = argsString.split(',').map(arg => arg.trim());
    
    let numbers: number[] = [];
    
    for (const arg of args) {
      if (arg.includes(':')) {
        // 範囲指定の場合
        const rangeValues = getCellRangeValues(arg, context);
        const rangeNumbers = rangeValues
          .map(v => Number(v))
          .filter(n => !isNaN(n));
        numbers = numbers.concat(rangeNumbers);
      } else {
        // 単一値の場合
        const value = getCellValue(arg, context) ?? arg;
        const num = Number(value);
        if (!isNaN(num)) {
          numbers.push(num);
        }
      }
    }
    
    if (numbers.length < 2) {
      return FormulaError.DIV0;
    }
    
    // 平均を計算
    const mean = numbers.reduce((sum, n) => sum + n, 0) / numbers.length;
    
    // 標本分散を返す
    return numbers.reduce((sum, n) => sum + Math.pow(n - mean, 2), 0) / (numbers.length - 1);
  }
};

// VARP関数の実装（分散・母集団 - 旧版）
export const VARP: CustomFormula = {
  name: 'VARP',
  pattern: /VARP\((.+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, argsString] = matches;
    const args = argsString.split(',').map(arg => arg.trim());
    
    let numbers: number[] = [];
    
    for (const arg of args) {
      if (arg.includes(':')) {
        // 範囲指定の場合
        const rangeValues = getCellRangeValues(arg, context);
        const rangeNumbers = rangeValues
          .map(v => Number(v))
          .filter(n => !isNaN(n));
        numbers = numbers.concat(rangeNumbers);
      } else {
        // 単一値の場合
        const value = getCellValue(arg, context) ?? arg;
        const num = Number(value);
        if (!isNaN(num)) {
          numbers.push(num);
        }
      }
    }
    
    if (numbers.length === 0) {
      return FormulaError.NUM;
    }
    
    // 平均を計算
    const mean = numbers.reduce((sum, n) => sum + n, 0) / numbers.length;
    
    // 母集団分散を返す
    return numbers.reduce((sum, n) => sum + Math.pow(n - mean, 2), 0) / numbers.length;
  }
};