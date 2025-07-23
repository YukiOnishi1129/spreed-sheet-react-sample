// 情報関数の実装

import type { CustomFormula, FormulaContext } from './types';
import { FormulaError } from './types';
import { getCellValue } from './utils';

// ISBLANK関数の実装（空白セルか判定）
export const ISBLANK: CustomFormula = {
  name: 'ISBLANK',
  pattern: /ISBLANK\(([^)]+)\)/i,
  calculate: (matches, context) => {
    const valueRef = matches[1].trim();
    const value = getCellValue(valueRef, context);
    
    return value === null || value === undefined || value === '';
  }
};

// ISERROR関数の実装（エラー値か判定）
export const ISERROR: CustomFormula = {
  name: 'ISERROR',
  pattern: /ISERROR\(([^)]+)\)/i,
  calculate: (matches, context) => {
    const valueRef = matches[1].trim();
    const value = getCellValue(valueRef, context);
    
    return typeof value === 'string' && value.startsWith('#') && value.endsWith('!');
  }
};

// ISNA関数の実装（#N/Aエラーか判定）
export const ISNA: CustomFormula = {
  name: 'ISNA',
  pattern: /ISNA\(([^)]+)\)/i,
  calculate: (matches, context) => {
    const valueRef = matches[1].trim();
    const value = getCellValue(valueRef, context);
    
    return value === FormulaError.NA;
  }
};

// ISTEXT関数の実装（文字列か判定）
export const ISTEXT: CustomFormula = {
  name: 'ISTEXT',
  pattern: /ISTEXT\(([^)]+)\)/i,
  calculate: (matches, context) => {
    const valueRef = matches[1].trim();
    const value = getCellValue(valueRef, context);
    
    return typeof value === 'string' && !(value.startsWith('#') && value.endsWith('!'));
  }
};

// ISNUMBER関数の実装（数値か判定）
export const ISNUMBER: CustomFormula = {
  name: 'ISNUMBER',
  pattern: /ISNUMBER\(([^)]+)\)/i,
  calculate: (matches, context) => {
    const valueRef = matches[1].trim();
    const value = getCellValue(valueRef, context);
    
    return typeof value === 'number' && !isNaN(value);
  }
};

// ISLOGICAL関数の実装（論理値か判定）
export const ISLOGICAL: CustomFormula = {
  name: 'ISLOGICAL',
  pattern: /ISLOGICAL\(([^)]+)\)/i,
  calculate: (matches, context) => {
    const valueRef = matches[1].trim();
    const value = getCellValue(valueRef, context);
    
    return typeof value === 'boolean';
  }
};

// ISEVEN関数の実装（偶数か判定）
export const ISEVEN: CustomFormula = {
  name: 'ISEVEN',
  pattern: /ISEVEN\(([^)]+)\)/i,
  calculate: (matches, context) => {
    const valueRef = matches[1].trim();
    const value = getCellValue(valueRef, context) ?? valueRef;
    const num = Number(value);
    
    if (isNaN(num)) {
      return FormulaError.VALUE;
    }
    
    return Math.floor(num) % 2 === 0;
  }
};

// ISODD関数の実装（奇数か判定）
export const ISODD: CustomFormula = {
  name: 'ISODD',
  pattern: /ISODD\(([^)]+)\)/i,
  calculate: (matches, context) => {
    const valueRef = matches[1].trim();
    const value = getCellValue(valueRef, context) ?? valueRef;
    const num = Number(value);
    
    if (isNaN(num)) {
      return FormulaError.VALUE;
    }
    
    return Math.floor(num) % 2 !== 0;
  }
};

// TYPE関数の実装（データ型を返す）
export const TYPE: CustomFormula = {
  name: 'TYPE',
  pattern: /TYPE\(([^)]+)\)/i,
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