// テキスト関数の実装

import type { CustomFormula, FormulaContext } from './types';
import { FormulaError } from './types';
import { getCellValue, getCellRangeValues } from './utils';

// CONCATENATE関数の実装
export const CONCATENATE: CustomFormula = {
  name: 'CONCATENATE',
  pattern: /CONCATENATE\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// CONCAT関数の実装
export const CONCAT: CustomFormula = {
  name: 'CONCAT',
  pattern: /CONCAT\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// LEFT関数の実装
export const LEFT: CustomFormula = {
  name: 'LEFT',
  pattern: /LEFT\(([^,]+),\s*(\d+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// RIGHT関数の実装
export const RIGHT: CustomFormula = {
  name: 'RIGHT',
  pattern: /RIGHT\(([^,]+),\s*(\d+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// MID関数の実装
export const MID: CustomFormula = {
  name: 'MID',
  pattern: /MID\(([^,]+),\s*(\d+),\s*(\d+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// LEN関数の実装
export const LEN: CustomFormula = {
  name: 'LEN',
  pattern: /LEN\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// UPPER関数の実装
export const UPPER: CustomFormula = {
  name: 'UPPER',
  pattern: /UPPER\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// LOWER関数の実装
export const LOWER: CustomFormula = {
  name: 'LOWER',
  pattern: /LOWER\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// TRIM関数の実装
export const TRIM: CustomFormula = {
  name: 'TRIM',
  pattern: /TRIM\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// SUBSTITUTE関数の実装
export const SUBSTITUTE: CustomFormula = {
  name: 'SUBSTITUTE',
  pattern: /SUBSTITUTE\(([^,]+),\s*"([^"]*)",\s*"([^"]*)"(?:,\s*(\d+))?\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// FIND関数の実装
export const FIND: CustomFormula = {
  name: 'FIND',
  pattern: /FIND\(([^,]+),\s*([^,)]+)(?:,\s*(\d+))?\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// SEARCH関数の実装
export const SEARCH: CustomFormula = {
  name: 'SEARCH',
  pattern: /SEARCH\(([^,]+),\s*([^,)]+)(?:,\s*(\d+))?\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// TEXTJOIN関数の実装（手動実装が必要）
export const TEXTJOIN: CustomFormula = {
  name: 'TEXTJOIN',
  pattern: /TEXTJOIN\("([^"]*)",\s*(TRUE|FALSE),\s*([^)]+)\)/i,
  isSupported: false, // HyperFormulaでサポートされていない可能性
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, delimiter, ignoreEmpty, rangeOrValues] = matches;
    const shouldIgnoreEmpty = ignoreEmpty.toUpperCase() === 'TRUE';
    
    console.log('TEXTJOIN計算:', { delimiter, ignoreEmpty, rangeOrValues });
    
    let values: unknown[] = [];
    
    // セル範囲の場合
    if (rangeOrValues.match(/^[A-Z]+\d+:[A-Z]+\d+$/)) {
      values = getCellRangeValues(rangeOrValues, context);
    } else {
      // 個別のセル参照や値の場合
      const parts = rangeOrValues.split(',');
      for (const part of parts) {
        const trimmed = part.trim();
        if (trimmed.match(/^[A-Z]+\d+$/)) {
          const cellValue = getCellValue(trimmed, context);
          values.push(cellValue);
        } else if (trimmed.startsWith('"') && trimmed.endsWith('"')) {
          values.push(trimmed.slice(1, -1));
        } else {
          values.push(trimmed);
        }
      }
    }
    
    // 空の値を除外するかどうか
    if (shouldIgnoreEmpty) {
      values = values.filter(v => v !== null && v !== undefined && v !== '');
    }
    
    // 文字列に変換して結合
    return values.map(v => String(v)).join(delimiter);
  }
};

// SPLIT関数の実装（スプレッドシート特有、手動実装が必要）
export const SPLIT: CustomFormula = {
  name: 'SPLIT',
  pattern: /SPLIT\(([^,]+),\s*"([^"]*)"(?:,\s*(TRUE|FALSE))?(?:,\s*(TRUE|FALSE))?\)/i,
  isSupported: false, // Excel標準ではない（Google Sheets特有）
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, textRef, delimiter, , removeEmptyText] = matches;
    
    let text: string;
    if (textRef.match(/^[A-Z]+\d+$/)) {
      text = String(getCellValue(textRef, context) ?? '');
    } else if (textRef.startsWith('"') && textRef.endsWith('"')) {
      text = textRef.slice(1, -1);
    } else {
      text = String(textRef);
    }
    
    // 簡単な分割実装（最初の結果のみ返す）
    const result = text.split(delimiter);
    
    if (removeEmptyText && removeEmptyText.toUpperCase() === 'TRUE') {
      return result.filter(s => s !== '')[0] || '';
    }
    
    return result[0] || '';
  }
};

// PROPER関数の実装（先頭大文字変換） - stringFunctions.tsから移動
export const PROPER: CustomFormula = {
  name: 'PROPER',
  pattern: /PROPER\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// VALUE関数の実装（文字列を数値に変換） - stringFunctions.tsから移動
export const VALUE: CustomFormula = {
  name: 'VALUE',
  pattern: /VALUE\(([^)]+)\)/i,
  isSupported: false,
  calculate: (matches) => {
    const text = matches[1].replace(/"/g, '');
    const number = parseFloat(text);
    
    if (isNaN(number)) return FormulaError.VALUE;
    return number;
  }
};

// TEXT関数の実装（数値を書式付き文字列に変換） - stringFunctions.tsから移動
export const TEXT: CustomFormula = {
  name: 'TEXT',
  pattern: /TEXT\(([^,]+),\s*([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// REPT関数の実装（文字列繰り返し） - stringFunctions.tsから移動
export const REPT: CustomFormula = {
  name: 'REPT',
  pattern: /REPT\(([^,]+),\s*([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// REPLACE関数の実装（位置指定文字置換）
export const REPLACE: CustomFormula = {
  name: 'REPLACE',
  pattern: /REPLACE\(([^,]+),\s*([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  isSupported: false,
  calculate: (matches, context) => {
    let oldText = matches[1].trim();
    const startNumRef = matches[2].trim();
    const numCharsRef = matches[3].trim();
    let newText = matches[4].trim();
    
    // 元の文字列を取得
    if (oldText.startsWith('"') && oldText.endsWith('"')) {
      oldText = oldText.slice(1, -1);
    } else if (oldText.match(/^[A-Z]+\d+$/)) {
      oldText = String(getCellValue(oldText, context) ?? '');
    }
    
    // 開始位置を取得
    let startNum: number;
    if (startNumRef.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(startNumRef, context);
      startNum = parseInt(String(cellValue ?? '1'));
    } else {
      startNum = parseInt(startNumRef);
    }
    
    // 文字数を取得
    let numChars: number;
    if (numCharsRef.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(numCharsRef, context);
      numChars = parseInt(String(cellValue ?? '0'));
    } else {
      numChars = parseInt(numCharsRef);
    }
    
    // 新しい文字列を取得
    if (newText.startsWith('"') && newText.endsWith('"')) {
      newText = newText.slice(1, -1);
    } else if (newText.match(/^[A-Z]+\d+$/)) {
      newText = String(getCellValue(newText, context) ?? '');
    }
    
    if (isNaN(startNum) || isNaN(numChars)) return FormulaError.VALUE;
    if (startNum < 1) return FormulaError.VALUE;
    
    const before = oldText.substring(0, startNum - 1);
    const after = oldText.substring(startNum - 1 + numChars);
    
    return before + newText + after;
  }
};

// CHAR関数の実装（文字コードから文字を返す）
export const CHAR: CustomFormula = {
  name: 'CHAR',
  pattern: /CHAR\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// CODE関数の実装（文字から文字コードを返す）
export const CODE: CustomFormula = {
  name: 'CODE',
  pattern: /CODE\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// EXACT関数の実装（文字列が同一か判定）
export const EXACT: CustomFormula = {
  name: 'EXACT',
  pattern: /EXACT\(([^,]+),\s*([^)]+)\)/i,
  isSupported: false,
  calculate: (matches, context) => {
    let text1 = matches[1].trim();
    let text2 = matches[2].trim();
    
    // 最初の文字列を取得
    if (text1.startsWith('"') && text1.endsWith('"')) {
      text1 = text1.slice(1, -1);
    } else if (text1.match(/^[A-Z]+\d+$/)) {
      text1 = String(getCellValue(text1, context) ?? '');
    }
    
    // 2番目の文字列を取得
    if (text2.startsWith('"') && text2.endsWith('"')) {
      text2 = text2.slice(1, -1);
    } else if (text2.match(/^[A-Z]+\d+$/)) {
      text2 = String(getCellValue(text2, context) ?? '');
    }
    
    return text1 === text2;
  }
};

// CLEAN関数の実装（印刷不可文字を削除）
export const CLEAN: CustomFormula = {
  name: 'CLEAN',
  pattern: /CLEAN\(([^)]+)\)/i,
  isSupported: false,
  calculate: (matches, context) => {
    let text = matches[1].trim();
    
    if (text.startsWith('"') && text.endsWith('"')) {
      text = text.slice(1, -1);
    } else if (text.match(/^[A-Z]+\d+$/)) {
      text = String(getCellValue(text, context) ?? '');
    }
    
    // 印刷不可文字（制御文字）を削除
    return text.replace(/[\x00-\x1F]/g, '');
  }
};

// T関数の実装（文字列を返す）
export const T: CustomFormula = {
  name: 'T',
  pattern: /T\(([^)]+)\)/i,
  isSupported: false,
  calculate: (matches, context) => {
    let value = matches[1].trim();
    
    if (value.startsWith('"') && value.endsWith('"')) {
      return value.slice(1, -1);
    } else if (value.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(value, context);
      return typeof cellValue === 'string' ? cellValue : '';
    }
    
    // 数値の場合は空文字を返す
    if (!isNaN(parseFloat(value))) {
      return '';
    }
    
    return String(value);
  }
};

// FIXED関数の実装（固定小数点表示）
export const FIXED: CustomFormula = {
  name: 'FIXED',
  pattern: /FIXED\(([^,]+)(?:,\s*([^,]+))?(?:,\s*([^)]+))?\)/i,
  isSupported: false,
  calculate: (matches, context) => {
    const numberRef = matches[1].trim();
    const decimalsRef = matches[2]?.trim() || '2';
    const noCommasRef = matches[3]?.trim() || 'FALSE';
    
    let number: number, decimals: number, noCommas: boolean;
    
    // 数値を取得
    if (numberRef.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(numberRef, context);
      number = parseFloat(String(cellValue ?? '0'));
    } else {
      number = parseFloat(numberRef);
    }
    
    // 小数点以下桁数を取得
    if (decimalsRef.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(decimalsRef, context);
      decimals = parseInt(String(cellValue ?? '2'));
    } else {
      decimals = parseInt(decimalsRef);
    }
    
    // カンマ区切りなしフラグを取得
    if (noCommasRef.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(noCommasRef, context);
      noCommas = Boolean(cellValue);
    } else {
      noCommas = noCommasRef.toUpperCase() === 'TRUE' || noCommasRef === '1';
    }
    
    if (isNaN(number) || isNaN(decimals)) return FormulaError.VALUE;
    if (decimals < 0) decimals = 0;
    
    let result = number.toFixed(decimals);
    
    if (!noCommas) {
      // カンマ区切りを追加
      const parts = result.split('.');
      parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',');
      result = parts.join('.');
    }
    
    return result;
  }
};