// 情報関数の実装

import type { CustomFormula, FormulaResult } from '../shared/types';
import { FormulaError } from '../shared/types';
import { getCellValue } from '../shared/utils';

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

// ISERR関数の実装（#N/A以外のエラー値か判定）
export const ISERR: CustomFormula = {
  name: 'ISERR',
  pattern: /ISERR\(([^)]+)\)/i,
  calculate: (matches, context) => {
    const valueRef = matches[1].trim();
    const value = getCellValue(valueRef, context);
    
    return typeof value === 'string' && value.startsWith('#') && value.endsWith('!') && value !== FormulaError.NA;
  }
};

// ISNONTEXT関数の実装（文字列以外か判定）
export const ISNONTEXT: CustomFormula = {
  name: 'ISNONTEXT',
  pattern: /ISNONTEXT\(([^)]+)\)/i,
  calculate: (matches, context) => {
    const valueRef = matches[1].trim();
    const value = getCellValue(valueRef, context);
    
    return !(typeof value === 'string' && !(value.startsWith('#') && value.endsWith('!')));
  }
};

// ISREF関数の実装（参照か判定）
export const ISREF: CustomFormula = {
  name: 'ISREF',
  pattern: /ISREF\(([^)]+)\)/i,
  calculate: (matches) => {
    const valueRef = matches[1].trim();
    
    // セル参照パターンをチェック
    return /^[A-Z]+\d+$/.test(valueRef) || /^[A-Z]+\d+:[A-Z]+\d+$/.test(valueRef);
  }
};

// ISFORMULA関数の実装（数式か判定）
export const ISFORMULA: CustomFormula = {
  name: 'ISFORMULA',
  pattern: /ISFORMULA\(([^)]+)\)/i,
  calculate: (matches, context) => {
    const valueRef = matches[1].trim();
    
    if (valueRef.match(/^[A-Z]+\d+$/)) {
      const rowMatch = valueRef.match(/\d+/);
      if (!rowMatch) return false;
      const cellData = context.data[parseInt(rowMatch[0]) - 1];
      if (cellData) {
        const colMatch = valueRef.match(/^[A-Z]+/);
        if (!colMatch) return false;
        const colIndex = colMatch[0].charCodeAt(0) - 65;
        const cell = cellData[colIndex];
        return typeof cell === 'object' && cell !== null && 'formula' in cell;
      }
    }
    
    return false;
  }
};

// NA関数の実装（#N/Aエラーを返す）
export const NA: CustomFormula = {
  name: 'NA',
  pattern: /NA\(\)/i,
  calculate: () => {
    return FormulaError.NA;
  }
};

// ERROR.TYPE関数の実装（エラーの種類を返す）
export const ERROR_TYPE: CustomFormula = {
  name: 'ERROR.TYPE',
  pattern: /ERROR\.TYPE\(([^)]+)\)/i,
  calculate: (matches, context) => {
    const valueRef = matches[1].trim();
    const value = getCellValue(valueRef, context);
    
    if (typeof value !== 'string' || !value.startsWith('#') || !value.endsWith('!')) {
      return FormulaError.NA;
    }
    
    // Excelのエラータイプ番号
    switch (value) {
      case FormulaError.NULL: return 1;
      case FormulaError.DIV0: return 2;
      case FormulaError.VALUE: return 3;
      case FormulaError.REF: return 4;
      case FormulaError.NAME: return 5;
      case FormulaError.NUM: return 6;
      case FormulaError.NA: return 7;
      default: return FormulaError.NA;
    }
  }
};

// INFO関数の実装（システム情報を返す）
export const INFO: CustomFormula = {
  name: 'INFO',
  pattern: /INFO\(([^)]+)\)/i,
  calculate: (matches, context) => {
    const typeRef = matches[1].trim();
    const cellValue = getCellValue(typeRef, context);
    const infoType = String(cellValue ?? typeRef.replace(/^['"]|['"]$/g, ''));
    
    switch (infoType.toLowerCase()) {
      case 'directory':
        return '/'; // デフォルトディレクトリ
      case 'numfile':
        return 1; // ファイル数
      case 'origin':
        return '$A$1'; // 原点
      case 'osversion':
        return 'Web'; // OS版本
      case 'recalc':
        return 'Automatic'; // 再計算モード
      case 'release':
        return '1.0'; // リリース版本
      case 'system':
        return 'Web'; // システム
      case 'version':
        return '1.0'; // 版本
      default:
        return FormulaError.VALUE;
    }
  }
};

// SHEET関数の実装（シート番号を返す）
export const SHEET: CustomFormula = {
  name: 'SHEET',
  pattern: /SHEET\(([^)]*)\)/i,
  calculate: () => {
    // 現在のシート番号を返す（固定値として1を返す）
    return 1;
  }
};

// SHEETS関数の実装（シート数を返す）
export const SHEETS: CustomFormula = {
  name: 'SHEETS',
  pattern: /SHEETS\(([^)]*)\)/i,
  calculate: () => {
    // シート数を返す（固定値として1を返す）
    return 1;
  }
};

// CELL関数の実装（セル情報を返す）
export const CELL: CustomFormula = {
  name: 'CELL',
  pattern: /CELL\(([^,]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches, context): FormulaResult => {
    const infoTypeRef = matches[1].trim();
    const referenceRef = matches[2]?.trim() || 'A1';
    
    const infoTypeCellValue = getCellValue(infoTypeRef, context);
    const infoType = String(infoTypeCellValue ?? infoTypeRef.replace(/^['"]|['"]$/g, ''));
    const cellValue = getCellValue(referenceRef, context);
    
    switch (infoType.toLowerCase()) {
      case 'address': {
        return `$${referenceRef}`;
      }
      case 'col': {
        const colMatch = referenceRef.match(/^([A-Z]+)/);
        if (colMatch) {
          const col = colMatch[1].split('').reduce((acc, char) => acc * 26 + char.charCodeAt(0) - 64, 0);
          return col;
        }
        return 1;
      }
      case 'color':
        return 0; // 色なし
      case 'contents': {
        const result: FormulaResult = cellValue !== null && cellValue !== undefined ? (cellValue as FormulaResult) : '';
        return result;
      }
      case 'filename':
        return 'spreadsheet.xlsx';
      case 'format':
        return 'G'; // 一般書式
      case 'parentheses':
        return 0; // 括弧なし
      case 'prefix':
        return ''; // プレフィックスなし
      case 'protect':
        return 0; // 保護なし
      case 'row': {
        const rowMatch = referenceRef.match(/(\d+)$/);
        return rowMatch ? parseInt(rowMatch[1]) : 1;
      }
      case 'type':
        if (typeof cellValue === 'string' && cellValue.startsWith('#')) return 'e'; // エラー
        if (typeof cellValue === 'string') return 'l'; // ラベル
        if (typeof cellValue === 'number') return 'v'; // 値
        return 'b'; // 空白
      case 'width':
        return 10; // 列幅
      default:
        return FormulaError.VALUE;
    }
  }
};