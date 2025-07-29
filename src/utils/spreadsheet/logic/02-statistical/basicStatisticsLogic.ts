// 基本統計関数の実装

import type { CustomFormula, FormulaContext } from '../shared/types';
import { FormulaError } from '../shared/types';
import { getCellValue, getCellRangeValues } from '../shared/utils';

// セル範囲から数値のみを抽出するヘルパー関数
function extractNumbersFromRange(rangeRef: string, context: FormulaContext): number[] {
  const numbers: number[] = [];
  
  if (rangeRef.includes(':')) {
    // セル範囲の場合
    const values = getCellRangeValues(rangeRef, context);
    
    values.forEach((value) => {
      // Handle object wrapped values
      let actualValue = value;
      if (value && typeof value === 'object' && 'value' in value) {
        actualValue = value.value;
      }
      
      const num = typeof actualValue === 'string' ? parseFloat(actualValue) : Number(actualValue);
      if (!isNaN(num) && isFinite(num)) {
        numbers.push(num);
      }
    });
  } else if (rangeRef.match(/^\$?[A-Z]+\$?\d+$/)) {
    // 単一セル参照の場合（絶対参照も含む）
    const cellValue = getCellValue(rangeRef, context);
    let actualValue = cellValue;
    if (cellValue && typeof cellValue === 'object' && 'value' in cellValue) {
      actualValue = cellValue.value;
    }
    
    const num = typeof actualValue === 'string' ? parseFloat(actualValue) : Number(actualValue);
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

// AVERAGE関数（平均値）
export const AVERAGE: CustomFormula = {
  name: 'AVERAGE',
  pattern: /AVERAGE\(([^)]+)\)/i,
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
    
    return numbers.reduce((sum, num) => sum + num, 0) / numbers.length;
  }
};

// COUNT関数（数値のセル数をカウント）
export const COUNT: CustomFormula = {
  name: 'COUNT',
  pattern: /COUNT\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, args] = matches;
    const parts = args.split(',').map(part => part.trim());
    let count = 0;
    
    for (const part of parts) {
      const numbers = extractNumbersFromRange(part, context);
      count += numbers.length;
    }
    
    return count;
  }
};

// COUNTA関数（空でないセル数をカウント）
export const COUNTA: CustomFormula = {
  name: 'COUNTA',
  pattern: /COUNTA\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, args] = matches;
    const parts = args.split(',').map(part => part.trim());
    let count = 0;
    
    for (const part of parts) {
      if (part.includes(':')) {
        const values = getCellRangeValues(part, context);
        values.forEach(value => {
          if (value !== null && value !== undefined && value !== '') {
            count++;
          }
        });
      } else if (part.match(/^[A-Z]+\d+$/)) {
        const cellValue = getCellValue(part, context);
        if (cellValue !== null && cellValue !== undefined && cellValue !== '') {
          count++;
        }
      } else {
        // 直接値の場合
        if (part !== '') {
          count++;
        }
      }
    }
    
    return count;
  }
};

// COUNTBLANK関数（空のセル数をカウント）
export const COUNTBLANK: CustomFormula = {
  name: 'COUNTBLANK',
  pattern: /COUNTBLANK\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, rangeRef] = matches;
    
    if (!rangeRef.includes(':')) {
      return FormulaError.VALUE;
    }
    
    const values = getCellRangeValues(rangeRef, context);
    let blankCount = 0;
    
    values.forEach(value => {
      // Handle object wrapped values
      let actualValue = value;
      if (value && typeof value === 'object' && 'value' in value) {
        actualValue = value.value;
      }
      
      const isBlank = actualValue === null || actualValue === undefined || actualValue === '';
      if (isBlank) {
        blankCount++;
      }
    });
    
    return blankCount;
  }
};

// MAX関数（最大値）
export const MAX: CustomFormula = {
  name: 'MAX',
  pattern: /MAX\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, args] = matches;
    const parts = args.split(',').map(part => part.trim());
    let numbers: number[] = [];
    
    for (const part of parts) {
      numbers = numbers.concat(extractNumbersFromRange(part, context));
    }
    
    if (numbers.length === 0) {
      return 0;
    }
    
    return Math.max(...numbers);
  }
};

// MIN関数（最小値）
export const MIN: CustomFormula = {
  name: 'MIN',
  pattern: /MIN\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, args] = matches;
    const parts = args.split(',').map(part => part.trim());
    let numbers: number[] = [];
    
    for (const part of parts) {
      numbers = numbers.concat(extractNumbersFromRange(part, context));
    }
    
    if (numbers.length === 0) {
      return 0;
    }
    
    return Math.min(...numbers);
  }
};

// MEDIAN関数（中央値）
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

// MODE関数（最頻値）
export const MODE: CustomFormula = {
  name: 'MODE',
  pattern: /MODE\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, args] = matches;
    const parts = args.split(',').map(part => part.trim());
    let numbers: number[] = [];
    
    for (const part of parts) {
      numbers = numbers.concat(extractNumbersFromRange(part, context));
    }
    
    if (numbers.length === 0) return FormulaError.NA;
    
    const frequency: { [key: number]: number } = {};
    numbers.forEach(num => {
      frequency[num] = (frequency[num] || 0) + 1;
    });
    
    let maxCount = 0;
    let mode: number | null = null;
    
    for (const [numStr, count] of Object.entries(frequency)) {
      if (count > maxCount) {
        maxCount = count;
        mode = parseFloat(numStr);
      }
    }
    
    return maxCount > 1 ? mode : FormulaError.NA;
  }
};

// STDEV.S関数（標準偏差、標本）
export const STDEV_S: CustomFormula = {
  name: 'STDEV.S',
  pattern: /STDEV\.S\(([^)]+)\)/i,
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
    
    const mean = numbers.reduce((sum, num) => sum + num, 0) / numbers.length;
    const variance = numbers.reduce((sum, num) => sum + Math.pow(num - mean, 2), 0) / (numbers.length - 1);
    
    return Math.sqrt(variance);
  }
};

// STDEV.P関数（標準偏差、母集団）
export const STDEV_P: CustomFormula = {
  name: 'STDEV.P',
  pattern: /STDEV\.P\(([^)]+)\)/i,
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
    
    const mean = numbers.reduce((sum, num) => sum + num, 0) / numbers.length;
    const variance = numbers.reduce((sum, num) => sum + Math.pow(num - mean, 2), 0) / numbers.length;
    
    return Math.sqrt(variance);
  }
};

// VAR.S関数（分散、標本）
export const VAR_S: CustomFormula = {
  name: 'VAR.S',
  pattern: /VAR\.S\(([^)]+)\)/i,
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
    
    const mean = numbers.reduce((sum, num) => sum + num, 0) / numbers.length;
    return numbers.reduce((sum, num) => sum + Math.pow(num - mean, 2), 0) / (numbers.length - 1);
  }
};

// VAR.P関数（分散、母集団）
export const VAR_P: CustomFormula = {
  name: 'VAR.P',
  pattern: /VAR\.P\(([^)]+)\)/i,
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
    
    const mean = numbers.reduce((sum, num) => sum + num, 0) / numbers.length;
    return numbers.reduce((sum, num) => sum + Math.pow(num - mean, 2), 0) / numbers.length;
  }
};

// AVEDEV関数（平均偏差）
export const AVEDEV: CustomFormula = {
  name: 'AVEDEV',
  pattern: /AVEDEV\(([^)]+)\)/i,
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
    
    const mean = numbers.reduce((sum, num) => sum + num, 0) / numbers.length;
    const avedev = numbers.reduce((sum, num) => sum + Math.abs(num - mean), 0) / numbers.length;
    
    return avedev;
  }
};

// GEOMEAN関数（幾何平均）
export const GEOMEAN: CustomFormula = {
  name: 'GEOMEAN',
  pattern: /GEOMEAN\(([^)]+)\)/i,
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
    
    // 全ての値が正数である必要がある
    if (numbers.some(num => num <= 0)) {
      return FormulaError.NUM;
    }
    
    const product = numbers.reduce((prod, num) => prod * num, 1);
    return Math.pow(product, 1 / numbers.length);
  }
};

// HARMEAN関数（調和平均）
export const HARMEAN: CustomFormula = {
  name: 'HARMEAN',
  pattern: /HARMEAN\(([^)]+)\)/i,
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
    
    // 全ての値が正数である必要がある
    if (numbers.some(num => num <= 0)) {
      return FormulaError.NUM;
    }
    
    const reciprocalSum = numbers.reduce((sum, num) => sum + (1 / num), 0);
    return numbers.length / reciprocalSum;
  }
};

// LARGE関数（k番目に大きい値）
export const LARGE: CustomFormula = {
  name: 'LARGE',
  pattern: /LARGE\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, arrayRef, kRef] = matches;
    
    const numbers = extractNumbersFromRange(arrayRef.trim(), context);
    
    // k値を取得（数値の直接入力も考慮）
    let k;
    const kTrimmed = kRef.trim();
    const directK = parseFloat(kTrimmed);
    if (!isNaN(directK)) {
      k = directK;
    } else {
      const kValue = getCellValue(kTrimmed, context);
      k = Number(kValue);
      if (isNaN(k)) {
        k = Number(kTrimmed);
      }
    }
    
    
    if (numbers.length === 0) {
      return FormulaError.NUM;
    }
    
    if (isNaN(k) || k <= 0 || k > numbers.length || !Number.isInteger(k)) {
      return FormulaError.NUM;
    }
    
    numbers.sort((a, b) => b - a); // 降順ソート
    return numbers[k - 1];
  }
};

// SMALL関数（k番目に小さい値）
export const SMALL: CustomFormula = {
  name: 'SMALL',
  pattern: /SMALL\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, arrayRef, kRef] = matches;
    
    const numbers = extractNumbersFromRange(arrayRef.trim(), context);
    const k = Number(getCellValue(kRef.trim(), context) ?? kRef);
    
    if (numbers.length === 0 || isNaN(k) || k <= 0 || k > numbers.length || !Number.isInteger(k)) {
      return FormulaError.NUM;
    }
    
    numbers.sort((a, b) => a - b); // 昇順ソート
    return numbers[k - 1];
  }
};

// RANK関数（順位）
export const RANK: CustomFormula = {
  name: 'RANK',
  pattern: /RANK\(([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, numberRef, arrayRef, orderRef] = matches;
    
    const number = Number(getCellValue(numberRef.trim(), context) ?? numberRef);
    const numbers = extractNumbersFromRange(arrayRef.trim(), context);
    const order = orderRef ? Number(getCellValue(orderRef.trim(), context) ?? orderRef) : 0;
    
    if (isNaN(number) || numbers.length === 0) {
      return FormulaError.NA;
    }
    
    if (!numbers.includes(number)) {
      return FormulaError.NA;
    }
    
    // 降順（order = 0）または昇順（order != 0）
    const sorted = order === 0 ? 
      [...numbers].sort((a, b) => b - a) : 
      [...numbers].sort((a, b) => a - b);
    
    return sorted.indexOf(number) + 1;
  }
};

// TRIMMEAN関数（トリム平均）
export const TRIMMEAN: CustomFormula = {
  name: 'TRIMMEAN',
  pattern: /TRIMMEAN\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, arrayRef, percentRef] = matches;
    
    const numbers = extractNumbersFromRange(arrayRef.trim(), context);
    const percent = Number(getCellValue(percentRef.trim(), context) ?? percentRef);
    
    if (numbers.length === 0 || isNaN(percent) || percent < 0 || percent >= 1) {
      return FormulaError.NUM;
    }
    
    const sorted = [...numbers].sort((a, b) => a - b);
    const trimCount = Math.floor(sorted.length * percent / 2);
    
    if (trimCount * 2 >= sorted.length) {
      return FormulaError.NUM;
    }
    
    const trimmed = sorted.slice(trimCount, sorted.length - trimCount);
    return trimmed.reduce((sum, num) => sum + num, 0) / trimmed.length;
  }
};

// STDEVA関数（文字列含む標本標準偏差）
export const STDEVA: CustomFormula = {
  name: 'STDEVA',
  pattern: /STDEVA\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, args] = matches;
    const parts = args.split(',').map(part => part.trim());
    const values: number[] = [];
    
    for (const part of parts) {
      if (part.includes(':')) {
        const rangeValues = getCellRangeValues(part, context);
        rangeValues.forEach(value => {
          if (typeof value === 'number') {
            values.push(value);
          } else if (typeof value === 'string') {
            const num = parseFloat(value);
            if (!isNaN(num)) {
              values.push(num);
            } else {
              values.push(0); // 文字列は0として扱う
            }
          } else if (typeof value === 'boolean') {
            values.push(value ? 1 : 0);
          }
        });
      } else {
        const value = getCellValue(part, context) ?? part;
        if (typeof value === 'number') {
          values.push(value);
        } else if (typeof value === 'string') {
          const num = parseFloat(value);
          values.push(isNaN(num) ? 0 : num);
        } else if (typeof value === 'boolean') {
          values.push(value ? 1 : 0);
        }
      }
    }
    
    if (values.length < 2) {
      return FormulaError.DIV0;
    }
    
    const mean = values.reduce((sum, val) => sum + val, 0) / values.length;
    const variance = values.reduce((sum, val) => sum + Math.pow(val - mean, 2), 0) / (values.length - 1);
    
    return Math.sqrt(variance);
  }
};

// STDEVPA関数（文字列含む母集団標準偏差）
export const STDEVPA: CustomFormula = {
  name: 'STDEVPA',
  pattern: /STDEVPA\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, args] = matches;
    const parts = args.split(',').map(part => part.trim());
    const values: number[] = [];
    
    for (const part of parts) {
      if (part.includes(':')) {
        const rangeValues = getCellRangeValues(part, context);
        rangeValues.forEach(value => {
          if (typeof value === 'number') {
            values.push(value);
          } else if (typeof value === 'string') {
            const num = parseFloat(value);
            if (!isNaN(num)) {
              values.push(num);
            } else {
              values.push(0); // 文字列は0として扱う
            }
          } else if (typeof value === 'boolean') {
            values.push(value ? 1 : 0);
          }
        });
      } else {
        const value = getCellValue(part, context) ?? part;
        if (typeof value === 'number') {
          values.push(value);
        } else if (typeof value === 'string') {
          const num = parseFloat(value);
          values.push(isNaN(num) ? 0 : num);
        } else if (typeof value === 'boolean') {
          values.push(value ? 1 : 0);
        }
      }
    }
    
    if (values.length === 0) {
      return FormulaError.DIV0;
    }
    
    const mean = values.reduce((sum, val) => sum + val, 0) / values.length;
    const variance = values.reduce((sum, val) => sum + Math.pow(val - mean, 2), 0) / values.length;
    
    return Math.sqrt(variance);
  }
};

// VARA関数（文字列含む標本分散）
export const VARA: CustomFormula = {
  name: 'VARA',
  pattern: /VARA\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, args] = matches;
    const parts = args.split(',').map(part => part.trim());
    const values: number[] = [];
    
    for (const part of parts) {
      if (part.includes(':')) {
        const rangeValues = getCellRangeValues(part, context);
        rangeValues.forEach(value => {
          if (typeof value === 'number') {
            values.push(value);
          } else if (typeof value === 'string') {
            const num = parseFloat(value);
            if (!isNaN(num)) {
              values.push(num);
            } else {
              values.push(0);
            }
          } else if (typeof value === 'boolean') {
            values.push(value ? 1 : 0);
          }
        });
      } else {
        const value = getCellValue(part, context) ?? part;
        if (typeof value === 'number') {
          values.push(value);
        } else if (typeof value === 'string') {
          const num = parseFloat(value);
          values.push(isNaN(num) ? 0 : num);
        } else if (typeof value === 'boolean') {
          values.push(value ? 1 : 0);
        }
      }
    }
    
    if (values.length < 2) {
      return FormulaError.DIV0;
    }
    
    const mean = values.reduce((sum, val) => sum + val, 0) / values.length;
    return values.reduce((sum, val) => sum + Math.pow(val - mean, 2), 0) / (values.length - 1);
  }
};

// VARPA関数（文字列含む母集団分散）
export const VARPA: CustomFormula = {
  name: 'VARPA',
  pattern: /VARPA\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, args] = matches;
    const parts = args.split(',').map(part => part.trim());
    const values: number[] = [];
    
    for (const part of parts) {
      if (part.includes(':')) {
        const rangeValues = getCellRangeValues(part, context);
        rangeValues.forEach(value => {
          if (typeof value === 'number') {
            values.push(value);
          } else if (typeof value === 'string') {
            const num = parseFloat(value);
            if (!isNaN(num)) {
              values.push(num);
            } else {
              values.push(0);
            }
          } else if (typeof value === 'boolean') {
            values.push(value ? 1 : 0);
          }
        });
      } else {
        const value = getCellValue(part, context) ?? part;
        if (typeof value === 'number') {
          values.push(value);
        } else if (typeof value === 'string') {
          const num = parseFloat(value);
          values.push(isNaN(num) ? 0 : num);
        } else if (typeof value === 'boolean') {
          values.push(value ? 1 : 0);
        }
      }
    }
    
    if (values.length === 0) {
      return FormulaError.DIV0;
    }
    
    const mean = values.reduce((sum, val) => sum + val, 0) / values.length;
    return values.reduce((sum, val) => sum + Math.pow(val - mean, 2), 0) / values.length;
  }
};