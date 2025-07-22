// 情報関数の実装

import type { CustomFormula } from './types';
import { FormulaError } from './types';
import { getCellValue } from './utils';

// ISBLANK関数の実装（空白セルか判定）
export const ISBLANK: CustomFormula = {
  name: 'ISBLANK',
  pattern: /ISBLANK\(([^)]+)\)/i,
  isSupported: false,
  calculate: (matches, context) => {
    const cellRef = matches[1].trim();
    
    // セル参照の場合
    if (cellRef.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(cellRef, context);
      return cellValue === null || cellValue === undefined || cellValue === '';
    }
    
    // 直接値の場合
    const value = cellRef.replace(/"/g, '');
    return value === '';
  }
};

// ISERROR関数の実装（エラー値か判定）
export const ISERROR: CustomFormula = {
  name: 'ISERROR',
  pattern: /ISERROR\(([^)]+)\)/i,
  isSupported: false,
  calculate: (matches, context) => {
    const valueRef = matches[1].trim();
    
    // セル参照の場合
    if (valueRef.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(valueRef, context);
      return typeof cellValue === 'string' && cellValue.startsWith('#') && cellValue.endsWith('!');
    }
    
    // 直接値の場合
    return valueRef.startsWith('#') && valueRef.endsWith('!');
  }
};

// ISNA関数の実装（#N/Aエラーか判定）
export const ISNA: CustomFormula = {
  name: 'ISNA',
  pattern: /ISNA\(([^)]+)\)/i,
  isSupported: false,
  calculate: (matches, context) => {
    const valueRef = matches[1].trim();
    
    // セル参照の場合
    if (valueRef.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(valueRef, context);
      return cellValue === '#N/A' || cellValue === '#N/A!';
    }
    
    // 直接値の場合
    return valueRef === '#N/A' || valueRef === '#N/A!';
  }
};

// ISTEXT関数の実装（文字列か判定）
export const ISTEXT: CustomFormula = {
  name: 'ISTEXT',
  pattern: /ISTEXT\(([^)]+)\)/i,
  isSupported: false,
  calculate: (matches, context) => {
    const valueRef = matches[1].trim();
    
    // セル参照の場合
    if (valueRef.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(valueRef, context);
      return typeof cellValue === 'string' && isNaN(parseFloat(cellValue));
    }
    
    // 直接値の場合
    if (valueRef.startsWith('"') && valueRef.endsWith('"')) {
      return true; // 明示的に文字列
    }
    
    // 数値でない文字列
    return isNaN(parseFloat(valueRef));
  }
};

// ISNUMBER関数の実装（数値か判定）
export const ISNUMBER: CustomFormula = {
  name: 'ISNUMBER',
  pattern: /ISNUMBER\(([^)]+)\)/i,
  isSupported: false,
  calculate: (matches, context) => {
    const valueRef = matches[1].trim();
    
    // セル参照の場合
    if (valueRef.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(valueRef, context);
      return typeof cellValue === 'number' || (typeof cellValue === 'string' && !isNaN(parseFloat(cellValue)));
    }
    
    // 直接値の場合
    return !isNaN(parseFloat(valueRef));
  }
};

// ISLOGICAL関数の実装（論理値か判定）
export const ISLOGICAL: CustomFormula = {
  name: 'ISLOGICAL',
  pattern: /ISLOGICAL\(([^)]+)\)/i,
  isSupported: false,
  calculate: (matches, context) => {
    const valueRef = matches[1].trim();
    
    // セル参照の場合
    if (valueRef.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(valueRef, context);
      return typeof cellValue === 'boolean';
    }
    
    // 直接値の場合
    const value = valueRef.replace(/"/g, '').toUpperCase();
    return value === 'TRUE' || value === 'FALSE';
  }
};

// ISEVEN関数の実装（偶数か判定）
export const ISEVEN: CustomFormula = {
  name: 'ISEVEN',
  pattern: /ISEVEN\(([^)]+)\)/i,
  isSupported: false,
  calculate: (matches, context) => {
    const valueRef = matches[1].trim();
    let value: number;
    
    // セル参照の場合
    if (valueRef.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(valueRef, context);
      value = typeof cellValue === 'number' ? cellValue : parseFloat(String(cellValue));
    } else {
      value = parseFloat(valueRef);
    }
    
    if (isNaN(value)) return FormulaError.VALUE;
    return Math.floor(value) % 2 === 0;
  }
};

// ISODD関数の実装（奇数か判定）
export const ISODD: CustomFormula = {
  name: 'ISODD',
  pattern: /ISODD\(([^)]+)\)/i,
  isSupported: false,
  calculate: (matches, context) => {
    const valueRef = matches[1].trim();
    let value: number;
    
    // セル参照の場合
    if (valueRef.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(valueRef, context);
      value = typeof cellValue === 'number' ? cellValue : parseFloat(String(cellValue));
    } else {
      value = parseFloat(valueRef);
    }
    
    if (isNaN(value)) return FormulaError.VALUE;
    return Math.floor(value) % 2 === 1;
  }
};

// TYPE関数の実装（データ型を返す）
export const TYPE: CustomFormula = {
  name: 'TYPE',
  pattern: /TYPE\(([^)]+)\)/i,
  isSupported: false,
  calculate: (matches, context) => {
    const valueRef = matches[1].trim();
    let value: unknown;
    
    // セル参照の場合
    if (valueRef.match(/^[A-Z]+\d+$/)) {
      value = getCellValue(valueRef, context);
    } else {
      // 直接値の場合
      if (valueRef.startsWith('"') && valueRef.endsWith('"')) {
        value = valueRef.slice(1, -1);
      } else if (!isNaN(parseFloat(valueRef))) {
        value = parseFloat(valueRef);
      } else {
        value = valueRef;
      }
    }
    
    // Excel TYPE関数の戻り値
    // 1=数値, 2=テキスト, 4=論理値, 16=エラー値, 64=配列
    if (typeof value === 'number') return 1;
    if (typeof value === 'string') {
      if (value.startsWith('#') && value.endsWith('!')) return 16; // エラー
      return 2; // テキスト
    }
    if (typeof value === 'boolean') return 4;
    if (Array.isArray(value)) return 64;
    
    return 2; // デフォルトはテキスト
  }
};

// N関数の実装（数値に変換）
export const N: CustomFormula = {
  name: 'N',
  pattern: /N\(([^)]+)\)/i,
  isSupported: false,
  calculate: (matches, context) => {
    const valueRef = matches[1].trim();
    let value: unknown;
    
    // セル参照の場合
    if (valueRef.match(/^[A-Z]+\d+$/)) {
      value = getCellValue(valueRef, context);
    } else {
      // 直接値の場合
      if (valueRef.startsWith('"') && valueRef.endsWith('"')) {
        value = valueRef.slice(1, -1);
      } else {
        value = valueRef;
      }
    }
    
    // 数値変換
    if (typeof value === 'number') return value;
    if (typeof value === 'boolean') return value ? 1 : 0;
    if (typeof value === 'string') {
      const numValue = parseFloat(value);
      return isNaN(numValue) ? 0 : numValue;
    }
    
    return 0;
  }
};