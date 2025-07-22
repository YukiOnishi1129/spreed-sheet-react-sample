// テキスト関数の実装

import type { CustomFormula, FormulaContext } from './types';
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
  isSupported: false,
  calculate: (matches) => {
    const text = matches[1].replace(/"/g, '');
    return text.replace(/\w\S*/g, (txt) => 
      txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase()
    );
  }
};

// VALUE関数の実装（文字列を数値に変換） - stringFunctions.tsから移動
export const VALUE: CustomFormula = {
  name: 'VALUE',
  pattern: /VALUE\(([^)]+)\)/i,
  isSupported: false,
  calculate: (matches) => {
    const text = matches[1].replace(/"/g, '');
    const number = parseFloat(text);
    
    if (isNaN(number)) return '#VALUE!';
    return number;
  }
};

// TEXT関数の実装（数値を書式付き文字列に変換） - stringFunctions.tsから移動
export const TEXT: CustomFormula = {
  name: 'TEXT',
  pattern: /TEXT\(([^,]+),\s*([^)]+)\)/i,
  isSupported: false,
  calculate: (matches) => {
    const value = parseFloat(matches[1]);
    const format = matches[2].replace(/"/g, '');
    
    if (isNaN(value)) return '#VALUE!';
    
    // 簡単な書式対応（完全なExcel書式ではない）
    if (format.includes('%')) {
      return (value * 100).toFixed(2) + '%';
    } else if (format.includes('0.00')) {
      return value.toFixed(2);
    } else if (format.includes('0')) {
      return Math.round(value).toString();
    }
    
    return value.toString();
  }
};

// REPT関数の実装（文字列繰り返し） - stringFunctions.tsから移動
export const REPT: CustomFormula = {
  name: 'REPT',
  pattern: /REPT\(([^,]+),\s*([^)]+)\)/i,
  isSupported: false,
  calculate: (matches) => {
    const text = matches[1].replace(/"/g, '');
    const times = parseInt(matches[2]);
    
    if (isNaN(times) || times < 0) return '#VALUE!';
    if (times > 32767) return '#VALUE!';
    
    return text.repeat(times);
  }
};