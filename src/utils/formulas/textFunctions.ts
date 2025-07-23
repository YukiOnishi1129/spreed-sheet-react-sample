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
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// CLEAN関数の実装（印刷不可文字を削除）
export const CLEAN: CustomFormula = {
  name: 'CLEAN',
  pattern: /CLEAN\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// T関数の実装（文字列を返す）
export const T: CustomFormula = {
  name: 'T',
  pattern: /T\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
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

// NUMBERVALUE関数の実装（文字列を数値に変換・地域設定対応）
export const NUMBERVALUE: CustomFormula = {
  name: 'NUMBERVALUE',
  pattern: /NUMBERVALUE\(([^,]+)(?:,\s*([^,]+))?(?:,\s*([^)]+))?\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// DOLLAR関数の実装（通貨書式に変換）
export const DOLLAR: CustomFormula = {
  name: 'DOLLAR',
  pattern: /DOLLAR\(([^,]+)(?:,\s*([^)]+))?\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// UNICHAR関数の実装（Unicode文字を返す）
export const UNICHAR: CustomFormula = {
  name: 'UNICHAR',
  pattern: /UNICHAR\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// UNICODE関数の実装（Unicode値を返す）
export const UNICODE: CustomFormula = {
  name: 'UNICODE',
  pattern: /UNICODE\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// LENB関数の実装（バイト数を返す）
export const LENB: CustomFormula = {
  name: 'LENB',
  pattern: /LENB\(([^)]+)\)/i,
  isSupported: false,
  calculate: (matches, context) => {
    let text = matches[1].trim();
    
    // セル参照の場合
    if (text.match(/^[A-Z]+\d+$/)) {
      text = String(getCellValue(text, context) ?? '');
    } else if (text.startsWith('"') && text.endsWith('"')) {
      text = text.slice(1, -1);
    }
    
    // UTF-8でのバイト数を計算
    const encoder = new TextEncoder();
    return encoder.encode(text).length;
  }
};

// FINDB関数の実装（バイト位置を検索）
export const FINDB: CustomFormula = {
  name: 'FINDB',
  pattern: /FINDB\(([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
  isSupported: false,
  calculate: (matches, context) => {
    let findText = matches[1].trim();
    let withinText = matches[2].trim();
    const startNumRef = matches[3]?.trim() || '1';
    
    // 検索文字列を取得
    if (findText.startsWith('"') && findText.endsWith('"')) {
      findText = findText.slice(1, -1);
    } else if (findText.match(/^[A-Z]+\d+$/)) {
      findText = String(getCellValue(findText, context) ?? '');
    }
    
    // 対象文字列を取得
    if (withinText.match(/^[A-Z]+\d+$/)) {
      withinText = String(getCellValue(withinText, context) ?? '');
    } else if (withinText.startsWith('"') && withinText.endsWith('"')) {
      withinText = withinText.slice(1, -1);
    }
    
    // 開始位置を取得
    let startNum: number;
    if (startNumRef.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(startNumRef, context);
      startNum = parseInt(String(cellValue ?? '1'));
    } else {
      startNum = parseInt(startNumRef);
    }
    
    if (isNaN(startNum) || startNum < 1) return FormulaError.VALUE;
    
    const encoder = new TextEncoder();
    const withinBytes = encoder.encode(withinText);
    const findBytes = encoder.encode(findText);
    
    // バイト単位で検索
    for (let i = startNum - 1; i <= withinBytes.length - findBytes.length; i++) {
      let match = true;
      for (let j = 0; j < findBytes.length; j++) {
        if (withinBytes[i + j] !== findBytes[j]) {
          match = false;
          break;
        }
      }
      if (match) {
        return i + 1; // 1ベースで返す
      }
    }
    
    return FormulaError.VALUE;
  }
};

// SEARCHB関数の実装（バイト位置を検索・大文字小文字区別なし）
export const SEARCHB: CustomFormula = {
  name: 'SEARCHB',
  pattern: /SEARCHB\(([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
  isSupported: false,
  calculate: (matches, context) => {
    let findText = matches[1].trim();
    let withinText = matches[2].trim();
    const startNumRef = matches[3]?.trim() || '1';
    
    // 検索文字列を取得
    if (findText.startsWith('"') && findText.endsWith('"')) {
      findText = findText.slice(1, -1);
    } else if (findText.match(/^[A-Z]+\d+$/)) {
      findText = String(getCellValue(findText, context) ?? '');
    }
    
    // 対象文字列を取得
    if (withinText.match(/^[A-Z]+\d+$/)) {
      withinText = String(getCellValue(withinText, context) ?? '');
    } else if (withinText.startsWith('"') && withinText.endsWith('"')) {
      withinText = withinText.slice(1, -1);
    }
    
    // 開始位置を取得
    let startNum: number;
    if (startNumRef.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(startNumRef, context);
      startNum = parseInt(String(cellValue ?? '1'));
    } else {
      startNum = parseInt(startNumRef);
    }
    
    if (isNaN(startNum) || startNum < 1) return FormulaError.VALUE;
    
    // 大文字小文字を区別しない
    const encoder = new TextEncoder();
    const withinBytes = encoder.encode(withinText.toLowerCase());
    const findBytes = encoder.encode(findText.toLowerCase());
    
    // バイト単位で検索
    for (let i = startNum - 1; i <= withinBytes.length - findBytes.length; i++) {
      let match = true;
      for (let j = 0; j < findBytes.length; j++) {
        if (withinBytes[i + j] !== findBytes[j]) {
          match = false;
          break;
        }
      }
      if (match) {
        return i + 1; // 1ベースで返す
      }
    }
    
    return FormulaError.VALUE;
  }
};

// REPLACEB関数の実装（バイト単位で置換）
export const REPLACEB: CustomFormula = {
  name: 'REPLACEB',
  pattern: /REPLACEB\(([^,]+),\s*([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  isSupported: false,
  calculate: (matches, context) => {
    let oldText = matches[1].trim();
    const startNumRef = matches[2].trim();
    const numBytesRef = matches[3].trim();
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
    
    // バイト数を取得
    let numBytes: number;
    if (numBytesRef.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(numBytesRef, context);
      numBytes = parseInt(String(cellValue ?? '0'));
    } else {
      numBytes = parseInt(numBytesRef);
    }
    
    // 新しい文字列を取得
    if (newText.startsWith('"') && newText.endsWith('"')) {
      newText = newText.slice(1, -1);
    } else if (newText.match(/^[A-Z]+\d+$/)) {
      newText = String(getCellValue(newText, context) ?? '');
    }
    
    if (isNaN(startNum) || isNaN(numBytes)) return FormulaError.VALUE;
    if (startNum < 1 || numBytes < 0) return FormulaError.VALUE;
    
    const encoder = new TextEncoder();
    const decoder = new TextDecoder();
    const oldBytes = encoder.encode(oldText);
    const newBytes = encoder.encode(newText);
    
    // バイト配列を操作
    const resultBytes = new Uint8Array(oldBytes.length - numBytes + newBytes.length);
    let pos = 0;
    
    // 開始位置までコピー
    for (let i = 0; i < startNum - 1 && i < oldBytes.length; i++) {
      resultBytes[pos++] = oldBytes[i];
    }
    
    // 新しい文字列をコピー
    for (let i = 0; i < newBytes.length; i++) {
      resultBytes[pos++] = newBytes[i];
    }
    
    // 残りをコピー
    for (let i = startNum - 1 + numBytes; i < oldBytes.length; i++) {
      resultBytes[pos++] = oldBytes[i];
    }
    
    return decoder.decode(resultBytes.slice(0, pos));
  }
};

// TEXTBEFORE関数の実装（区切り文字の前を抽出）
export const TEXTBEFORE: CustomFormula = {
  name: 'TEXTBEFORE',
  pattern: /TEXTBEFORE\(([^,]+),\s*([^,)]+)(?:,\s*([^,)]+))?(?:,\s*([^,)]+))?(?:,\s*([^,)]+))?(?:,\s*([^)]+))?\)/i,
  isSupported: false,
  calculate: (matches, context) => {
    let text = matches[1].trim();
    let delimiter = matches[2].trim();
    const instanceNum = matches[3]?.trim() || '1';
    const matchMode = matches[4]?.trim() || '0';
    const ifNotFound = matches[6]?.trim() || '#N/A';
    
    // テキストを取得
    if (text.match(/^[A-Z]+\d+$/)) {
      text = String(getCellValue(text, context) ?? '');
    } else if (text.startsWith('"') && text.endsWith('"')) {
      text = text.slice(1, -1);
    }
    
    // 区切り文字を取得
    if (delimiter.startsWith('"') && delimiter.endsWith('"')) {
      delimiter = delimiter.slice(1, -1);
    } else if (delimiter.match(/^[A-Z]+\d+$/)) {
      delimiter = String(getCellValue(delimiter, context) ?? '');
    }
    
    // インスタンス番号を取得
    let instance = parseInt(instanceNum);
    if (isNaN(instance)) instance = 1;
    
    // マッチモードを取得（0: 大文字小文字を区別、1: 区別しない）
    const caseSensitive = matchMode === '0';
    
    // 検索用のテキストと区切り文字
    const searchText = caseSensitive ? text : text.toLowerCase();
    const searchDelimiter = caseSensitive ? delimiter : delimiter.toLowerCase();
    
    // 区切り文字の位置を検索
    const positions: number[] = [];
    let pos = 0;
    while ((pos = searchText.indexOf(searchDelimiter, pos)) !== -1) {
      positions.push(pos);
      pos += searchDelimiter.length;
    }
    
    if (positions.length === 0) {
      // 区切り文字が見つからない場合
      if (ifNotFound === '#N/A') {
        return FormulaError.NA;
      }
      return ifNotFound.replace(/^"|"$/g, '');
    }
    
    // インスタンス番号に基づいて位置を選択
    let targetPos: number;
    if (instance > 0) {
      if (instance > positions.length) {
        if (ifNotFound === '#N/A') {
          return FormulaError.NA;
        }
        return ifNotFound.replace(/^"|"$/g, '');
      }
      targetPos = positions[instance - 1];
    } else {
      // 負の場合は後ろから数える
      const absInstance = Math.abs(instance);
      if (absInstance > positions.length) {
        if (ifNotFound === '#N/A') {
          return FormulaError.NA;
        }
        return ifNotFound.replace(/^"|"$/g, '');
      }
      targetPos = positions[positions.length - absInstance];
    }
    
    return text.substring(0, targetPos);
  }
};

// TEXTAFTER関数の実装（区切り文字の後を抽出）
export const TEXTAFTER: CustomFormula = {
  name: 'TEXTAFTER',
  pattern: /TEXTAFTER\(([^,]+),\s*([^,)]+)(?:,\s*([^,)]+))?(?:,\s*([^,)]+))?(?:,\s*([^,)]+))?(?:,\s*([^)]+))?\)/i,
  isSupported: false,
  calculate: (matches, context) => {
    let text = matches[1].trim();
    let delimiter = matches[2].trim();
    const instanceNum = matches[3]?.trim() || '1';
    const matchMode = matches[4]?.trim() || '0';
    const ifNotFound = matches[6]?.trim() || '#N/A';
    
    // テキストを取得
    if (text.match(/^[A-Z]+\d+$/)) {
      text = String(getCellValue(text, context) ?? '');
    } else if (text.startsWith('"') && text.endsWith('"')) {
      text = text.slice(1, -1);
    }
    
    // 区切り文字を取得
    if (delimiter.startsWith('"') && delimiter.endsWith('"')) {
      delimiter = delimiter.slice(1, -1);
    } else if (delimiter.match(/^[A-Z]+\d+$/)) {
      delimiter = String(getCellValue(delimiter, context) ?? '');
    }
    
    // インスタンス番号を取得
    let instance = parseInt(instanceNum);
    if (isNaN(instance)) instance = 1;
    
    // マッチモードを取得（0: 大文字小文字を区別、1: 区別しない）
    const caseSensitive = matchMode === '0';
    
    // 検索用のテキストと区切り文字
    const searchText = caseSensitive ? text : text.toLowerCase();
    const searchDelimiter = caseSensitive ? delimiter : delimiter.toLowerCase();
    
    // 区切り文字の位置を検索
    const positions: number[] = [];
    let pos = 0;
    while ((pos = searchText.indexOf(searchDelimiter, pos)) !== -1) {
      positions.push(pos);
      pos += searchDelimiter.length;
    }
    
    if (positions.length === 0) {
      // 区切り文字が見つからない場合
      if (ifNotFound === '#N/A') {
        return FormulaError.NA;
      }
      return ifNotFound.replace(/^"|"$/g, '');
    }
    
    // インスタンス番号に基づいて位置を選択
    let targetPos: number;
    if (instance > 0) {
      if (instance > positions.length) {
        if (ifNotFound === '#N/A') {
          return FormulaError.NA;
        }
        return ifNotFound.replace(/^"|"$/g, '');
      }
      targetPos = positions[instance - 1];
    } else {
      // 負の場合は後ろから数える
      const absInstance = Math.abs(instance);
      if (absInstance > positions.length) {
        if (ifNotFound === '#N/A') {
          return FormulaError.NA;
        }
        return ifNotFound.replace(/^"|"$/g, '');
      }
      targetPos = positions[positions.length - absInstance];
    }
    
    return text.substring(targetPos + delimiter.length);
  }
};

// TEXTSPLIT関数の実装（文字列を分割）
export const TEXTSPLIT: CustomFormula = {
  name: 'TEXTSPLIT',
  pattern: /TEXTSPLIT\(([^,]+),\s*([^,)]+)(?:,\s*([^,)]+))?(?:,\s*([^,)]+))?(?:,\s*([^,)]+))?(?:,\s*([^)]+))?\)/i,
  isSupported: false,
  calculate: (matches, context) => {
    let text = matches[1].trim();
    let colDelimiter = matches[2].trim();
    const ignoreEmpty = matches[4]?.trim() || 'FALSE';
    
    // テキストを取得
    if (text.match(/^[A-Z]+\d+$/)) {
      text = String(getCellValue(text, context) ?? '');
    } else if (text.startsWith('"') && text.endsWith('"')) {
      text = text.slice(1, -1);
    }
    
    // 列区切り文字を取得
    if (colDelimiter.startsWith('"') && colDelimiter.endsWith('"')) {
      colDelimiter = colDelimiter.slice(1, -1);
    } else if (colDelimiter.match(/^[A-Z]+\d+$/)) {
      colDelimiter = String(getCellValue(colDelimiter, context) ?? '');
    }
    
    // 空白を無視するかどうか
    const shouldIgnoreEmpty = ignoreEmpty.toUpperCase() === 'TRUE';
    
    // 分割実行（簡略化版：最初の要素のみ返す）
    const parts = text.split(colDelimiter);
    
    if (shouldIgnoreEmpty) {
      const filtered = parts.filter(p => p !== '');
      return filtered[0] || '';
    }
    
    return parts[0] || '';
  }
};

// ASC関数の実装（全角を半角に変換）
export const ASC: CustomFormula = {
  name: 'ASC',
  pattern: /ASC\(([^)]+)\)/i,
  isSupported: false,
  calculate: (matches, context) => {
    let text = matches[1].trim();
    
    // テキストを取得
    if (text.match(/^[A-Z]+\d+$/)) {
      text = String(getCellValue(text, context) ?? '');
    } else if (text.startsWith('"') && text.endsWith('"')) {
      text = text.slice(1, -1);
    }
    
    // 全角から半角への変換マップ
    const fullToHalf: { [key: string]: string } = {
      '　': ' ', '！': '!', '＂': '"', '＃': '#', '＄': '$', '％': '%', '＆': '&', '＇': "'",
      '（': '(', '）': ')', '＊': '*', '＋': '+', '，': ',', '－': '-', '．': '.', '／': '/',
      '０': '0', '１': '1', '２': '2', '３': '3', '４': '4', '５': '5', '６': '6', '７': '7',
      '８': '8', '９': '9', '：': ':', '；': ';', '＜': '<', '＝': '=', '＞': '>', '？': '?',
      '＠': '@', 'Ａ': 'A', 'Ｂ': 'B', 'Ｃ': 'C', 'Ｄ': 'D', 'Ｅ': 'E', 'Ｆ': 'F', 'Ｇ': 'G',
      'Ｈ': 'H', 'Ｉ': 'I', 'Ｊ': 'J', 'Ｋ': 'K', 'Ｌ': 'L', 'Ｍ': 'M', 'Ｎ': 'N', 'Ｏ': 'O',
      'Ｐ': 'P', 'Ｑ': 'Q', 'Ｒ': 'R', 'Ｓ': 'S', 'Ｔ': 'T', 'Ｕ': 'U', 'Ｖ': 'V', 'Ｗ': 'W',
      'Ｘ': 'X', 'Ｙ': 'Y', 'Ｚ': 'Z', '［': '[', '＼': '\\', '］': ']', '＾': '^', '＿': '_',
      '｀': '`', 'ａ': 'a', 'ｂ': 'b', 'ｃ': 'c', 'ｄ': 'd', 'ｅ': 'e', 'ｆ': 'f', 'ｇ': 'g',
      'ｈ': 'h', 'ｉ': 'i', 'ｊ': 'j', 'ｋ': 'k', 'ｌ': 'l', 'ｍ': 'm', 'ｎ': 'n', 'ｏ': 'o',
      'ｐ': 'p', 'ｑ': 'q', 'ｒ': 'r', 'ｓ': 's', 'ｔ': 't', 'ｕ': 'u', 'ｖ': 'v', 'ｗ': 'w',
      'ｘ': 'x', 'ｙ': 'y', 'ｚ': 'z', '｛': '{', '｜': '|', '｝': '}', '～': '~',
      // カタカナ
      'ア': 'ｱ', 'イ': 'ｲ', 'ウ': 'ｳ', 'エ': 'ｴ', 'オ': 'ｵ',
      'カ': 'ｶ', 'キ': 'ｷ', 'ク': 'ｸ', 'ケ': 'ｹ', 'コ': 'ｺ',
      'サ': 'ｻ', 'シ': 'ｼ', 'ス': 'ｽ', 'セ': 'ｾ', 'ソ': 'ｿ',
      'タ': 'ﾀ', 'チ': 'ﾁ', 'ツ': 'ﾂ', 'テ': 'ﾃ', 'ト': 'ﾄ',
      'ナ': 'ﾅ', 'ニ': 'ﾆ', 'ヌ': 'ﾇ', 'ネ': 'ﾈ', 'ノ': 'ﾉ',
      'ハ': 'ﾊ', 'ヒ': 'ﾋ', 'フ': 'ﾌ', 'ヘ': 'ﾍ', 'ホ': 'ﾎ',
      'マ': 'ﾏ', 'ミ': 'ﾐ', 'ム': 'ﾑ', 'メ': 'ﾒ', 'モ': 'ﾓ',
      'ヤ': 'ﾔ', 'ユ': 'ﾕ', 'ヨ': 'ﾖ',
      'ラ': 'ﾗ', 'リ': 'ﾘ', 'ル': 'ﾙ', 'レ': 'ﾚ', 'ロ': 'ﾛ',
      'ワ': 'ﾜ', 'ヲ': 'ｦ', 'ン': 'ﾝ',
      'ガ': 'ｶﾞ', 'ギ': 'ｷﾞ', 'グ': 'ｸﾞ', 'ゲ': 'ｹﾞ', 'ゴ': 'ｺﾞ',
      'ザ': 'ｻﾞ', 'ジ': 'ｼﾞ', 'ズ': 'ｽﾞ', 'ゼ': 'ｾﾞ', 'ゾ': 'ｿﾞ',
      'ダ': 'ﾀﾞ', 'ヂ': 'ﾁﾞ', 'ヅ': 'ﾂﾞ', 'デ': 'ﾃﾞ', 'ド': 'ﾄﾞ',
      'バ': 'ﾊﾞ', 'ビ': 'ﾋﾞ', 'ブ': 'ﾌﾞ', 'ベ': 'ﾍﾞ', 'ボ': 'ﾎﾞ',
      'パ': 'ﾊﾟ', 'ピ': 'ﾋﾟ', 'プ': 'ﾌﾟ', 'ペ': 'ﾍﾟ', 'ポ': 'ﾎﾟ',
      'ァ': 'ｧ', 'ィ': 'ｨ', 'ゥ': 'ｩ', 'ェ': 'ｪ', 'ォ': 'ｫ',
      'ッ': 'ｯ', 'ャ': 'ｬ', 'ュ': 'ｭ', 'ョ': 'ｮ',
      'ー': 'ｰ', '・': '･', '「': '｢', '」': '｣'
    };
    
    let result = '';
    for (const char of text) {
      result += fullToHalf[char] || char;
    }
    
    return result;
  }
};

// JIS関数の実装（半角を全角に変換）
export const JIS: CustomFormula = {
  name: 'JIS',
  pattern: /JIS\(([^)]+)\)/i,
  isSupported: false,
  calculate: (matches, context) => {
    let text = matches[1].trim();
    
    // テキストを取得
    if (text.match(/^[A-Z]+\d+$/)) {
      text = String(getCellValue(text, context) ?? '');
    } else if (text.startsWith('"') && text.endsWith('"')) {
      text = text.slice(1, -1);
    }
    
    // 半角から全角への変換マップ
    const halfToFull: { [key: string]: string } = {
      ' ': '　', '!': '！', '"': '＂', '#': '＃', '$': '＄', '%': '％', '&': '＆', "'": '＇',
      '(': '（', ')': '）', '*': '＊', '+': '＋', ',': '，', '-': '－', '.': '．', '/': '／',
      '0': '０', '1': '１', '2': '２', '3': '３', '4': '４', '5': '５', '6': '６', '7': '７',
      '8': '８', '9': '９', ':': '：', ';': '；', '<': '＜', '=': '＝', '>': '＞', '?': '？',
      '@': '＠', 'A': 'Ａ', 'B': 'Ｂ', 'C': 'Ｃ', 'D': 'Ｄ', 'E': 'Ｅ', 'F': 'Ｆ', 'G': 'Ｇ',
      'H': 'Ｈ', 'I': 'Ｉ', 'J': 'Ｊ', 'K': 'Ｋ', 'L': 'Ｌ', 'M': 'Ｍ', 'N': 'Ｎ', 'O': 'Ｏ',
      'P': 'Ｐ', 'Q': 'Ｑ', 'R': 'Ｒ', 'S': 'Ｓ', 'T': 'Ｔ', 'U': 'Ｕ', 'V': 'Ｖ', 'W': 'Ｗ',
      'X': 'Ｘ', 'Y': 'Ｙ', 'Z': 'Ｚ', '[': '［', '\\': '＼', ']': '］', '^': '＾', '_': '＿',
      '`': '｀', 'a': 'ａ', 'b': 'ｂ', 'c': 'ｃ', 'd': 'ｄ', 'e': 'ｅ', 'f': 'ｆ', 'g': 'ｇ',
      'h': 'ｈ', 'i': 'ｉ', 'j': 'ｊ', 'k': 'ｋ', 'l': 'ｌ', 'm': 'ｍ', 'n': 'ｎ', 'o': 'ｏ',
      'p': 'ｐ', 'q': 'ｑ', 'r': 'ｒ', 's': 'ｓ', 't': 'ｔ', 'u': 'ｕ', 'v': 'ｖ', 'w': 'ｗ',
      'x': 'ｘ', 'y': 'ｙ', 'z': 'ｚ', '{': '｛', '|': '｜', '}': '｝', '~': '～',
      // 半角カタカナ
      'ｱ': 'ア', 'ｲ': 'イ', 'ｳ': 'ウ', 'ｴ': 'エ', 'ｵ': 'オ',
      'ｶ': 'カ', 'ｷ': 'キ', 'ｸ': 'ク', 'ｹ': 'ケ', 'ｺ': 'コ',
      'ｻ': 'サ', 'ｼ': 'シ', 'ｽ': 'ス', 'ｾ': 'セ', 'ｿ': 'ソ',
      'ﾀ': 'タ', 'ﾁ': 'チ', 'ﾂ': 'ツ', 'ﾃ': 'テ', 'ﾄ': 'ト',
      'ﾅ': 'ナ', 'ﾆ': 'ニ', 'ﾇ': 'ヌ', 'ﾈ': 'ネ', 'ﾉ': 'ノ',
      'ﾊ': 'ハ', 'ﾋ': 'ヒ', 'ﾌ': 'フ', 'ﾍ': 'ヘ', 'ﾎ': 'ホ',
      'ﾏ': 'マ', 'ﾐ': 'ミ', 'ﾑ': 'ム', 'ﾒ': 'メ', 'ﾓ': 'モ',
      'ﾔ': 'ヤ', 'ﾕ': 'ユ', 'ﾖ': 'ヨ',
      'ﾗ': 'ラ', 'ﾘ': 'リ', 'ﾙ': 'ル', 'ﾚ': 'レ', 'ﾛ': 'ロ',
      'ﾜ': 'ワ', 'ｦ': 'ヲ', 'ﾝ': 'ン',
      'ｧ': 'ァ', 'ｨ': 'ィ', 'ｩ': 'ゥ', 'ｪ': 'ェ', 'ｫ': 'ォ',
      'ｯ': 'ッ', 'ｬ': 'ャ', 'ｭ': 'ュ', 'ｮ': 'ョ',
      'ｰ': 'ー', '･': '・', '｢': '「', '｣': '」',
      'ﾞ': '゛', 'ﾟ': '゜'
    };
    
    // 濁点・半濁点の処理
    let result = '';
    for (let i = 0; i < text.length; i++) {
      const char = text[i];
      const nextChar = text[i + 1];
      
      if (nextChar === 'ﾞ' || nextChar === 'ﾟ') {
        // 濁点・半濁点付きの文字を処理
        const combined = char + nextChar;
        const dakutenMap: { [key: string]: string } = {
          'ｶﾞ': 'ガ', 'ｷﾞ': 'ギ', 'ｸﾞ': 'グ', 'ｹﾞ': 'ゲ', 'ｺﾞ': 'ゴ',
          'ｻﾞ': 'ザ', 'ｼﾞ': 'ジ', 'ｽﾞ': 'ズ', 'ｾﾞ': 'ゼ', 'ｿﾞ': 'ゾ',
          'ﾀﾞ': 'ダ', 'ﾁﾞ': 'ヂ', 'ﾂﾞ': 'ヅ', 'ﾃﾞ': 'デ', 'ﾄﾞ': 'ド',
          'ﾊﾞ': 'バ', 'ﾋﾞ': 'ビ', 'ﾌﾞ': 'ブ', 'ﾍﾞ': 'ベ', 'ﾎﾞ': 'ボ',
          'ﾊﾟ': 'パ', 'ﾋﾟ': 'ピ', 'ﾌﾟ': 'プ', 'ﾍﾟ': 'ペ', 'ﾎﾟ': 'ポ'
        };
        if (dakutenMap[combined]) {
          result += dakutenMap[combined];
          i++; // 次の文字をスキップ
          continue;
        }
      }
      
      result += halfToFull[char] || char;
    }
    
    return result;
  }
};

// DBCS関数の実装（JISと同じ、半角を全角に変換）
export const DBCS: CustomFormula = {
  name: 'DBCS',
  pattern: /DBCS\(([^)]+)\)/i,
  isSupported: false,
  calculate: JIS.calculate // JIS関数と同じ実装
};

// PHONETIC関数の実装（ふりがなを抽出）
export const PHONETIC: CustomFormula = {
  name: 'PHONETIC',
  pattern: /PHONETIC\(([^)]+)\)/i,
  isSupported: false,
  calculate: (matches, context) => {
    const reference = matches[1].trim();
    
    // セル参照の場合
    if (reference.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(reference, context);
      // 実際のふりがな抽出は実装が複雑なため、元の値を返す
      return String(cellValue ?? '');
    }
    
    // 直接値の場合も元の値を返す
    if (reference.startsWith('"') && reference.endsWith('"')) {
      return reference.slice(1, -1);
    }
    
    return reference;
  }
};

// BAHTTEXT関数の実装（数値をタイ語文字列に変換）
export const BAHTTEXT: CustomFormula = {
  name: 'BAHTTEXT',
  pattern: /BAHTTEXT\(([^)]+)\)/i,
  isSupported: false,
  calculate: (matches, context) => {
    const numberRef = matches[1].trim();
    
    // 数値を取得
    let number: number;
    if (numberRef.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(numberRef, context);
      number = parseFloat(String(cellValue ?? '0'));
    } else {
      number = parseFloat(numberRef);
    }
    
    if (isNaN(number)) return FormulaError.VALUE;
    
    // 簡略化実装：タイバーツ形式の文字列を返す
    const baht = Math.floor(number);
    const satang = Math.round((number - baht) * 100);
    
    // 実際のタイ語変換は複雑なため、簡略化した表現を返す
    if (satang > 0) {
      return `${baht} บาท ${satang} สตางค์`;
    } else {
      return `${baht} บาท`;
    }
  }
};