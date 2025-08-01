// 情報関数の実装

import type { CustomFormula, FormulaResult } from '../shared/types';
import { FormulaError } from '../shared/types';
import { getCellValue } from '../shared/utils';

// ISBLANK関数の実装（空白セルか判定）
export const ISBLANK: CustomFormula = {
  name: 'ISBLANK',
  pattern: /\bISBLANK\(([^)]+)\)/i,
  calculate: (matches, context) => {
    const valueRef = matches[1].trim();
    const value = getCellValue(valueRef, context);
    
    // getCellValueは { value: null } を返す場合があるので、その場合も考慮
    if (typeof value === 'object' && value !== null && 'value' in value) {
      return value.value === null || value.value === undefined || value.value === '';
    }
    
    // 範囲外のセル参照（#REF!エラー）も空白として扱う
    if (value === FormulaError.REF) {
      return true;
    }
    
    return value === null || value === undefined || value === '';
  }
};

// ISERROR関数の実装（エラー値か判定）
export const ISERROR: CustomFormula = {
  name: 'ISERROR',
  pattern: /\bISERROR\(([^)]+)\)/i,
  calculate: (matches, context) => {
    const valueRef = matches[1].trim();
    const value = getCellValue(valueRef, context);
    
    // FormulaError.NAは'#N/A'で、'!'で終わらない
    return typeof value === 'string' && value.startsWith('#') && (value.endsWith('!') || value === FormulaError.NA || value === '#N/A');
  }
};

// ISNA関数の実装（#N/Aエラーか判定）
export const ISNA: CustomFormula = {
  name: 'ISNA',
  pattern: /\bISNA\(([^)]+)\)/i,
  calculate: (matches, context) => {
    const valueRef = matches[1].trim();
    const value = getCellValue(valueRef, context);
    
    
    // Check the error constant - FormulaError.NA is '#N/A' (without !)
    return value === FormulaError.NA || value === '#N/A';
  }
};

// ISTEXT関数の実装（文字列か判定）
export const ISTEXT: CustomFormula = {
  name: 'ISTEXT',
  pattern: /\bISTEXT\(([^)]+)\)/i,
  calculate: (matches, context) => {
    const valueRef = matches[1].trim();
    const value = getCellValue(valueRef, context);
    
    return typeof value === 'string' && !(value.startsWith('#') && value.endsWith('!'));
  }
};

// ISNUMBER関数の実装（数値か判定）  
export const ISNUMBER: CustomFormula = {
  name: 'ISNUMBER',
  pattern: /\bISNUMBER\(([^)]+)\)/i,
  calculate: (matches, context) => {
    const valueRef = matches[1].trim();
    const value = getCellValue(valueRef, context);
    
    // Excel ISNUMBER: only returns true for actual numeric values, not string numbers
    return typeof value === 'number' && !isNaN(value) && isFinite(value);
  }
};

// ISLOGICAL関数の実装（論理値か判定）
export const ISLOGICAL: CustomFormula = {
  name: 'ISLOGICAL',
  pattern: /\bISLOGICAL\(([^)]+)\)/i,
  calculate: (matches, context) => {
    const valueRef = matches[1].trim();
    const value = getCellValue(valueRef, context);
    
    return typeof value === 'boolean';
  }
};

// ISEVEN関数の実装（偶数か判定）
export const ISEVEN: CustomFormula = {
  name: 'ISEVEN',
  pattern: /\bISEVEN\(([^)]+)\)/i,
  calculate: (matches, context) => {
    const valueRef = matches[1].trim();
    const value = getCellValue(valueRef, context) ?? valueRef;
    const num = Number(value);
    
    if (isNaN(num)) {
      return FormulaError.VALUE;
    }
    
    // Excelでは値を切り捨てて判定（例：2.5 → 2 → 偶数）
    return Math.floor(Math.abs(num)) % 2 === 0;
  }
};

// ISODD関数の実装（奇数か判定）
export const ISODD: CustomFormula = {
  name: 'ISODD',
  pattern: /\bISODD\(([^)]+)\)/i,
  calculate: (matches, context) => {
    const valueRef = matches[1].trim();
    const value = getCellValue(valueRef, context) ?? valueRef;
    const num = Number(value);
    
    if (isNaN(num)) {
      return FormulaError.VALUE;
    }
    
    // Excelでは値を切り捨てて判定（例：2.5 → 2 → 偶数、3.5 → 3 → 奇数）
    return Math.floor(Math.abs(num)) % 2 !== 0;
  }
};

// TYPE関数の実装（データ型を返す）
export const TYPE: CustomFormula = {
  name: 'TYPE',
  pattern: /(?<!ERROR\.)\bTYPE\(([^)]+)\)/i,
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
  pattern: /\bN\(([^)]+)\)/i,
  calculate: (matches, context) => {
    const valueRef = matches[1].trim();
    const value = getCellValue(valueRef, context);
    
    // 数値変換 - Excel N関数の仕様に従う
    if (typeof value === 'number') return value;
    if (typeof value === 'boolean') return value ? 1 : 0;
    // 文字列の場合は常に0を返す（Excel仕様）
    return 0;
  }
};

// ISERR関数の実装（#N/A以外のエラー値か判定）
export const ISERR: CustomFormula = {
  name: 'ISERR',
  pattern: /\bISERR\(([^)]+)\)/i,
  calculate: (matches, context) => {
    const valueRef = matches[1].trim();
    const value = getCellValue(valueRef, context);
    
    return typeof value === 'string' && value.startsWith('#') && value.endsWith('!') && value !== FormulaError.NA;
  }
};

// ISNONTEXT関数の実装（文字列以外か判定）
export const ISNONTEXT: CustomFormula = {
  name: 'ISNONTEXT',
  pattern: /\bISNONTEXT\(([^)]+)\)/i,
  calculate: (matches, context) => {
    const valueRef = matches[1].trim();
    const value = getCellValue(valueRef, context);
    
    return !(typeof value === 'string' && !(value.startsWith('#') && value.endsWith('!')));
  }
};

// ISREF関数の実装（参照か判定）
export const ISREF: CustomFormula = {
  name: 'ISREF',
  pattern: /\bISREF\(([^)]+)\)/i,
  calculate: (matches) => {
    const valueRef = matches[1].trim();
    
    // セル参照パターンをチェック
    return /^[A-Z]+\d+$/.test(valueRef) || /^[A-Z]+\d+:[A-Z]+\d+$/.test(valueRef);
  }
};

// ISFORMULA関数の実装（数式か判定）
export const ISFORMULA: CustomFormula = {
  name: 'ISFORMULA',
  pattern: /\bISFORMULA\(([^)]+)\)/i,
  calculate: (matches, context) => {
    const valueRef = matches[1].trim();
    
    if (valueRef.match(/^[A-Z]+\d+$/)) {
      const rowMatch = valueRef.match(/\d+/);
      if (!rowMatch) return false;
      const row = parseInt(rowMatch[0]) - 1;
      if (row < 0 || row >= context.data.length) return false;
      
      const colMatch = valueRef.match(/^[A-Z]+/);
      if (!colMatch) return false;
      const colIndex = colMatch[0].charCodeAt(0) - 65;
      
      const cellData = context.data[row];
      if (!cellData || colIndex < 0 || colIndex >= cellData.length) return false;
      
      const cell = cellData[colIndex];
      
      // Check if cell has a formula property and is not just a value
      if (cell && typeof cell === 'object' && 'formula' in cell && cell.formula) {
        return true;
      }
      
      // For raw cell values, check if it's a string that starts with '='
      if (typeof cell === 'object' && cell && 'value' in cell && !('formula' in cell)) {
        const cellData = cell as { value: unknown };
        const value = cellData.value;
        return typeof value === 'string' && value.startsWith('=');
      }
      
      // If it's just a plain value (number, string, etc.), it's not a formula
      return false;
    }
    
    return false;
  }
};

// NA関数の実装（#N/Aエラーを返す）
export const NA: CustomFormula = {
  name: 'NA',
  pattern: /\bNA\(\)/i,
  calculate: () => {
    return FormulaError.NA;
  }
};

// ERROR.TYPE関数の実装（エラーの種類を返す）
export const ERROR_TYPE: CustomFormula = {
  name: 'ERROR.TYPE',
  pattern: /\bERROR\.TYPE\(([^)]+)\)/i,
  calculate: (matches, context) => {
    const valueRef = matches[1].trim();
    const value = getCellValue(valueRef, context);
    
    // Check if the value is an error
    if (typeof value !== 'string' || !value.startsWith('#')) {
      return FormulaError.NA;
    }
    
    // Excel ERROR.TYPE function return values
    switch (value) {
      case FormulaError.NULL:
      case '#NULL!': return 1;
      case FormulaError.DIV0:
      case '#DIV/0!': return 2;
      case FormulaError.VALUE:
      case '#VALUE!': return 3;
      case FormulaError.REF:
      case '#REF!': return 4;
      case FormulaError.NAME:
      case '#NAME!':
      case '#NAME?': return 5;
      case FormulaError.NUM:
      case '#NUM!': return 6;
      case FormulaError.NA:
      case '#N/A': return 7;
      default: return FormulaError.NA;
    }
  }
};

// INFO関数の実装（システム情報を返す）
export const INFO: CustomFormula = {
  name: 'INFO',
  pattern: /\bINFO\(([^)]+)\)/i,
  calculate: (matches, context) => {
    const typeRef = matches[1].trim();
    
    // まず直接文字列として扱う（引用符を削除）
    let infoType = typeRef.replace(/^["']|["']$/g, '');
    
    // セル参照の場合のみgetCellValueを使う
    if (typeRef.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(typeRef, context);
      infoType = String(cellValue ?? '');
    }
    
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
  pattern: /\bSHEET\(([^)]*)\)/i,
  calculate: () => {
    // 現在のシート番号を返す（固定値として1を返す）
    return 1;
  }
};

// SHEETS関数の実装（シート数を返す）
export const SHEETS: CustomFormula = {
  name: 'SHEETS',
  pattern: /\bSHEETS\(([^)]*)\)/i,
  calculate: () => {
    // シート数を返す（固定値として1を返す）
    return 1;
  }
};

// CELL関数の実装（セル情報を返す）
export const CELL: CustomFormula = {
  name: 'CELL',
  pattern: /\bCELL\(([^,]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches, context): FormulaResult => {
    const infoTypeRef = matches[1].trim();
    const referenceRef = matches[2]?.trim() || 'A1';
    
    // まず直接文字列として扱う（引用符を削除）
    let infoType = infoTypeRef.replace(/^["']|["']$/g, '');
    
    // セル参照の場合のみgetCellValueを使う
    if (infoTypeRef.match(/^[A-Z]+\d+$/)) {
      const infoTypeCellValue = getCellValue(infoTypeRef, context);
      infoType = String(infoTypeCellValue ?? '');
    }
    
    const cellValue = getCellValue(referenceRef, context);
    
    switch (infoType.toLowerCase()) {
      case 'address': {
        // Excelのアドレス形式：列も行も絶対参照
        const colMatch = referenceRef.match(/^([A-Z]+)/);
        const rowMatch = referenceRef.match(/(\d+)$/);
        if (colMatch && rowMatch) {
          return `$${colMatch[1]}$${rowMatch[1]}`;
        }
        return '$A$1';
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

// ISBETWEEN関数の実装（値が範囲内か判定）- Google Sheets専用
export const ISBETWEEN: CustomFormula = {
  name: 'ISBETWEEN',
  pattern: /\bISBETWEEN\(([^,]+),\s*([^,]+),\s*([^,)]+)(?:,\s*([^,)]+))?(?:,\s*([^)]+))?\)/i,
  calculate: (matches, context): FormulaResult => {
    const [, valueRef, lowerRef, upperRef, lowerInclusiveRef, upperInclusiveRef] = matches;
    
    const value = getCellValue(valueRef.trim(), context);
    const lower = getCellValue(lowerRef.trim(), context);
    const upper = getCellValue(upperRef.trim(), context);
    
    // デフォルトは両方含む
    const lowerInclusive = lowerInclusiveRef ? 
      getCellValue(lowerInclusiveRef.trim(), context)?.toString().toLowerCase() !== 'false' : true;
    const upperInclusive = upperInclusiveRef ? 
      getCellValue(upperInclusiveRef.trim(), context)?.toString().toLowerCase() !== 'false' : true;
    
    // 数値に変換
    const numValue = Number(value);
    const numLower = Number(lower);
    const numUpper = Number(upper);
    
    if (isNaN(numValue) || isNaN(numLower) || isNaN(numUpper)) {
      return FormulaError.VALUE;
    }
    
    // 範囲チェック
    const lowerCheck = lowerInclusive ? numValue >= numLower : numValue > numLower;
    const upperCheck = upperInclusive ? numValue <= numUpper : numValue < numUpper;
    
    return lowerCheck && upperCheck;
  }
};