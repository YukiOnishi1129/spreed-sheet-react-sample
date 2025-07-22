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
  calculate: (matches, context) => {
    console.log('TEXT関数実行:', { matches });
    
    const valueRef = matches[1].trim();
    const formatRef = matches[2].trim();
    
    // 値の取得
    let value: number;
    if (valueRef.match(/^[A-Z]+\d+$/)) {
      // セル参照
      const cellValue = getCellValue(valueRef, context);
      value = parseFloat(String(cellValue ?? '0'));
    } else {
      value = parseFloat(valueRef);
    }
    
    // フォーマットの取得
    let format = formatRef;
    if (format.startsWith('"') && format.endsWith('"')) {
      format = format.slice(1, -1);
    }
    
    console.log('TEXT計算:', { value, format, valueRef, formatRef });
    
    if (isNaN(value)) {
      console.error('TEXT: 数値変換失敗', { valueRef, value });
      return '#VALUE!';
    }
    
    // 簡単な書式対応（完全なExcel書式ではない）
    let result: string;
    if (format.includes('¥') || format.includes('￥')) {
      // 通貨形式
      result = '¥' + value.toLocaleString();
    } else if (format.includes('%')) {
      result = (value * 100).toFixed(2) + '%';
    } else if (format.includes('0.00')) {
      result = value.toFixed(2);
    } else if (format.includes('0,000')) {
      result = value.toLocaleString();
    } else if (format.includes('0')) {
      result = Math.round(value).toString();
    } else {
      result = value.toString();
    }
    
    console.log('TEXT結果:', { result });
    return result;
  }
};

// REPT関数の実装（文字列繰り返し） - stringFunctions.tsから移動
export const REPT: CustomFormula = {
  name: 'REPT',
  pattern: /REPT\(([^,]+),\s*([^)]+)\)/i,
  isSupported: false,
  calculate: (matches, context) => {
    console.log('REPT関数実行:', { matches });
    
    let text = matches[1].trim();
    const timesRef = matches[2].trim();
    
    // テキスト部分の処理
    if (text.startsWith('"') && text.endsWith('"')) {
      // 引用符で囲まれた文字列
      text = text.slice(1, -1);
    } else if (text.match(/^[A-Z]+\d+$/)) {
      // セル参照
      text = String(getCellValue(text, context) ?? '');
    }
    
    // 繰り返し回数の処理
    let times: number;
    if (timesRef.match(/^[A-Z]+\d+$/)) {
      // セル参照
      const cellValue = getCellValue(timesRef, context);
      times = parseInt(String(cellValue ?? '0'));
    } else {
      times = parseInt(timesRef);
    }
    
    console.log('REPT計算:', { text, times, originalText: matches[1], originalTimes: matches[2] });
    
    if (isNaN(times) || times < 0) {
      console.error('REPT: 無効な回数', times);
      return '#VALUE!';
    }
    if (times > 32767) {
      console.error('REPT: 回数が上限を超過', times);
      return '#VALUE!';
    }
    
    const result = text.repeat(times);
    console.log('REPT結果:', result);
    return result;
  }
};