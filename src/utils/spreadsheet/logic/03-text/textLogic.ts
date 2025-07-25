// テキスト関数の実装

import type { CustomFormula, FormulaContext } from '../shared/types';
import { FormulaError } from '../shared/types';
import { getCellValue, getCellRangeValues } from '../shared/utils';
import { matchFormula } from '../index';

// CONCATENATE関数の実装
export const CONCATENATE: CustomFormula = {
  name: 'CONCATENATE',
  pattern: /CONCATENATE\((.+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, argsString] = matches;
    
    // 引数を正しく分割（括弧内のカンマは無視）
    const parts: string[] = [];
    let currentPart = '';
    let parenDepth = 0;
    let inQuotes = false;
    
    for (let i = 0; i < argsString.length; i++) {
      const char = argsString[i];
      
      if (char === '"' && argsString[i - 1] !== '\\') {
        inQuotes = !inQuotes;
      }
      
      if (!inQuotes) {
        if (char === '(') parenDepth++;
        else if (char === ')') parenDepth--;
        else if (char === ',' && parenDepth === 0) {
          parts.push(currentPart.trim());
          currentPart = '';
          continue;
        }
      }
      
      currentPart += char;
    }
    
    if (currentPart.trim()) {
      parts.push(currentPart.trim());
    }
    
    // 各部分を評価して結合
    let result = '';
    for (const part of parts) {
      let value: string;
      
      // セル参照の場合
      if (part.match(/^[A-Z]+\d+$/)) {
        value = String(getCellValue(part, context) ?? '');
      }
      // 文字列リテラルの場合
      else if (part.startsWith('"') && part.endsWith('"')) {
        value = part.slice(1, -1);
      }
      // 関数呼び出しの場合（LEFT、RIGHTなど）
      else if (part.match(/^[A-Z]+\(/i)) {
        // matchFormulaを使用してネストした関数を計算
        const matchResult = matchFormula(part);
        
        if (matchResult) {
          const nestedResult = matchResult.function.calculate(matchResult.matches, context);
          value = String(nestedResult);
        } else {
          value = '#VALUE!';
        }
      }
      // その他
      else {
        value = String(part);
      }
      
      result += value;
    }
    
    return result;
  }
};

// CONCAT関数の実装
export const CONCAT: CustomFormula = {
  name: 'CONCAT',
  pattern: /CONCAT\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, args] = matches;
    const parts = args.split(',').map(part => part.trim());
    let result = '';
    
    for (const part of parts) {
      // 範囲指定（A1:B2のような形式）かチェック
      if (part.includes(':')) {
        const rangeValues = getCellRangeValues(part, context);
        for (const value of rangeValues) {
          result += String(value ?? '');
        }
      } else {
        let value: string;
        if (part.match(/^[A-Z]+\d+$/)) {
          value = String(getCellValue(part, context) ?? '');
        } else if (part.startsWith('"') && part.endsWith('"')) {
          value = part.slice(1, -1);
        } else {
          value = String(part);
        }
        result += value;
      }
    }
    
    return result;
  }
};

// LEFT関数の実装
export const LEFT: CustomFormula = {
  name: 'LEFT',
  pattern: /LEFT\(([^,]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, textRef, numCharsRef] = matches;
    
    let text: string;
    if (textRef.match(/^[A-Z]+\d+$/)) {
      text = String(getCellValue(textRef, context) ?? '');
    } else if (textRef.startsWith('"') && textRef.endsWith('"')) {
      text = textRef.slice(1, -1);
    } else {
      text = String(textRef);
    }
    
    let numChars = 1; // デフォルト値
    if (numCharsRef) {
      // 演算式の場合（FIND("@",A7)-1など）
      if (numCharsRef.includes('-') || numCharsRef.includes('+') || numCharsRef.includes('*') || numCharsRef.includes('/')) {
        // 簡単な演算を処理（FIND("@",A7)-1のような形式）
        const parts = numCharsRef.split(/([+\-*/])/);
        if (parts.length === 3) {
          const [leftPart, operator, rightPart] = parts;
          let leftValue: number;
          
          // 左辺が関数呼び出しの場合
          if (leftPart.trim().match(/^[A-Z]+\(/i)) {
            const matchResult = matchFormula(leftPart.trim());
            if (matchResult) {
              const findResult = matchResult.function.calculate(matchResult.matches, context);
              leftValue = parseInt(String(findResult));
            } else {
              leftValue = parseInt(leftPart.trim());
            }
          } else {
            leftValue = parseInt(leftPart.trim());
          }
          
          const rightValue = parseInt(rightPart.trim());
          
          if (!isNaN(leftValue) && !isNaN(rightValue)) {
            switch (operator) {
              case '+': numChars = leftValue + rightValue; break;
              case '-': numChars = leftValue - rightValue; break;
              case '*': numChars = leftValue * rightValue; break;
              case '/': numChars = Math.floor(leftValue / rightValue); break;
            }
          }
        }
      }
      // セル参照の場合
      else if (numCharsRef.match(/^[A-Z]+\d+$/)) {
        const cellValue = getCellValue(numCharsRef, context);
        numChars = parseInt(String(cellValue ?? '1'));
      } else {
        numChars = parseInt(numCharsRef);
      }
    }
    
    if (isNaN(numChars) || numChars < 0) {
      return FormulaError.VALUE;
    }
    
    return text.substring(0, numChars);
  }
};

// RIGHT関数の実装
export const RIGHT: CustomFormula = {
  name: 'RIGHT',
  pattern: /RIGHT\(([^,]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, textRef, numCharsRef] = matches;
    
    let text: string;
    if (textRef.match(/^[A-Z]+\d+$/)) {
      text = String(getCellValue(textRef, context) ?? '');
    } else if (textRef.startsWith('"') && textRef.endsWith('"')) {
      text = textRef.slice(1, -1);
    } else {
      text = String(textRef);
    }
    
    let numChars = 1; // デフォルト値
    if (numCharsRef) {
      if (numCharsRef.match(/^[A-Z]+\d+$/)) {
        const cellValue = getCellValue(numCharsRef, context);
        numChars = parseInt(String(cellValue ?? '1'));
      } else {
        numChars = parseInt(numCharsRef);
      }
    }
    
    if (isNaN(numChars) || numChars < 0) {
      return FormulaError.VALUE;
    }
    
    return text.substring(Math.max(0, text.length - numChars));
  }
};

// MID関数の実装
export const MID: CustomFormula = {
  name: 'MID',
  pattern: /MID\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, textRef, startNumRef, numCharsRef] = matches;
    
    let text: string;
    if (textRef.match(/^[A-Z]+\d+$/)) {
      text = String(getCellValue(textRef, context) ?? '');
    } else if (textRef.startsWith('"') && textRef.endsWith('"')) {
      text = textRef.slice(1, -1);
    } else {
      text = String(textRef);
    }
    
    let startNum: number;
    if (startNumRef.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(startNumRef, context);
      startNum = parseInt(String(cellValue ?? '1'));
    } else {
      startNum = parseInt(startNumRef);
    }
    
    let numChars: number;
    if (numCharsRef.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(numCharsRef, context);
      numChars = parseInt(String(cellValue ?? '0'));
    } else {
      numChars = parseInt(numCharsRef);
    }
    
    if (isNaN(startNum) || isNaN(numChars) || startNum < 1 || numChars < 0) {
      return FormulaError.VALUE;
    }
    
    return text.substring(startNum - 1, startNum - 1 + numChars);
  }
};

// LEN関数の実装
export const LEN: CustomFormula = {
  name: 'LEN',
  pattern: /LEN\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, textRef] = matches;
    
    let text: string;
    if (textRef.match(/^[A-Z]+\d+$/)) {
      text = String(getCellValue(textRef, context) ?? '');
    } else if (textRef.startsWith('"') && textRef.endsWith('"')) {
      text = textRef.slice(1, -1);
    } else {
      text = String(textRef);
    }
    
    return text.length;
  }
};

// UPPER関数の実装
export const UPPER: CustomFormula = {
  name: 'UPPER',
  pattern: /UPPER\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, textRef] = matches;
    
    let text: string;
    if (textRef.match(/^[A-Z]+\d+$/)) {
      text = String(getCellValue(textRef, context) ?? '');
    } else if (textRef.startsWith('"') && textRef.endsWith('"')) {
      text = textRef.slice(1, -1);
    } else {
      text = String(textRef);
    }
    
    return text.toUpperCase();
  }
};

// LOWER関数の実装
export const LOWER: CustomFormula = {
  name: 'LOWER',
  pattern: /LOWER\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, textRef] = matches;
    
    let text: string;
    if (textRef.match(/^[A-Z]+\d+$/)) {
      text = String(getCellValue(textRef, context) ?? '');
    } else if (textRef.startsWith('"') && textRef.endsWith('"')) {
      text = textRef.slice(1, -1);
    } else {
      text = String(textRef);
    }
    
    return text.toLowerCase();
  }
};

// TRIM関数の実装
export const TRIM: CustomFormula = {
  name: 'TRIM',
  pattern: /TRIM\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, textRef] = matches;
    
    let text: string;
    if (textRef.match(/^[A-Z]+\d+$/)) {
      text = String(getCellValue(textRef, context) ?? '');
    } else if (textRef.startsWith('"') && textRef.endsWith('"')) {
      text = textRef.slice(1, -1);
    } else {
      text = String(textRef);
    }
    
    // 先頭と末尾の空白を削除し、複数の連続する空白を単一の空白に変換
    return text.replace(/\s+/g, ' ').trim();
  }
};

// SUBSTITUTE関数の実装
export const SUBSTITUTE: CustomFormula = {
  name: 'SUBSTITUTE',
  pattern: /SUBSTITUTE\(([^,]+),\s*([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, textRef, oldTextRef, newTextRef, instanceNumRef] = matches;
    
    let text: string;
    if (textRef.match(/^[A-Z]+\d+$/)) {
      text = String(getCellValue(textRef, context) ?? '');
    } else if (textRef.startsWith('"') && textRef.endsWith('"')) {
      text = textRef.slice(1, -1);
    } else {
      text = String(textRef);
    }
    
    let oldText: string;
    if (oldTextRef.startsWith('"') && oldTextRef.endsWith('"')) {
      oldText = oldTextRef.slice(1, -1);
    } else if (oldTextRef.match(/^[A-Z]+\d+$/)) {
      oldText = String(getCellValue(oldTextRef, context) ?? '');
    } else {
      oldText = String(oldTextRef);
    }
    
    let newText: string;
    if (newTextRef.startsWith('"') && newTextRef.endsWith('"')) {
      newText = newTextRef.slice(1, -1);
    } else if (newTextRef.match(/^[A-Z]+\d+$/)) {
      newText = String(getCellValue(newTextRef, context) ?? '');
    } else {
      newText = String(newTextRef);
    }
    
    // インスタンス番号の取得
    let instanceNum: number | undefined;
    if (instanceNumRef) {
      if (instanceNumRef.match(/^[A-Z]+\d+$/)) {
        const cellValue = getCellValue(instanceNumRef, context);
        instanceNum = parseInt(String(cellValue ?? '1'));
      } else {
        instanceNum = parseInt(instanceNumRef);
      }
      
      if (isNaN(instanceNum) || instanceNum < 1) {
        return FormulaError.VALUE;
      }
    }
    
    if (instanceNum) {
      // 特定のインスタンスのみ置換
      let count = 0;
      let pos = 0;
      while ((pos = text.indexOf(oldText, pos)) !== -1) {
        count++;
        if (count === instanceNum) {
          return text.substring(0, pos) + newText + text.substring(pos + oldText.length);
        }
        pos += oldText.length;
      }
      return text; // 特定のインスタンスが見つからない場合
    } else {
      // すべて置換
      return text.replaceAll(oldText, newText);
    }
  }
};

// FIND関数の実装
export const FIND: CustomFormula = {
  name: 'FIND',
  pattern: /FIND\(([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, findTextRef, withinTextRef, startNumRef] = matches;
    
    let findText: string;
    if (findTextRef.startsWith('"') && findTextRef.endsWith('"')) {
      findText = findTextRef.slice(1, -1);
    } else if (findTextRef.match(/^[A-Z]+\d+$/)) {
      findText = String(getCellValue(findTextRef, context) ?? '');
    } else {
      findText = String(findTextRef);
    }
    
    let withinText: string;
    if (withinTextRef.match(/^[A-Z]+\d+$/)) {
      withinText = String(getCellValue(withinTextRef, context) ?? '');
    } else if (withinTextRef.startsWith('"') && withinTextRef.endsWith('"')) {
      withinText = withinTextRef.slice(1, -1);
    } else {
      withinText = String(withinTextRef);
    }
    
    let startNum = 1; // デフォルト値
    if (startNumRef) {
      if (startNumRef.match(/^[A-Z]+\d+$/)) {
        const cellValue = getCellValue(startNumRef, context);
        startNum = parseInt(String(cellValue ?? '1'));
      } else {
        startNum = parseInt(startNumRef);
      }
    }
    
    if (isNaN(startNum) || startNum < 1) {
      return FormulaError.VALUE;
    }
    
    const index = withinText.indexOf(findText, startNum - 1);
    if (index === -1) {
      return FormulaError.VALUE;
    }
    
    return index + 1; // 1ベースで返す
  }
};

// SEARCH関数の実装
export const SEARCH: CustomFormula = {
  name: 'SEARCH',
  pattern: /SEARCH\(([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, findTextRef, withinTextRef, startNumRef] = matches;
    
    let findText: string;
    if (findTextRef.startsWith('"') && findTextRef.endsWith('"')) {
      findText = findTextRef.slice(1, -1).toLowerCase();
    } else if (findTextRef.match(/^[A-Z]+\d+$/)) {
      findText = String(getCellValue(findTextRef, context) ?? '').toLowerCase();
    } else {
      findText = String(findTextRef).toLowerCase();
    }
    
    let withinText: string;
    if (withinTextRef.match(/^[A-Z]+\d+$/)) {
      withinText = String(getCellValue(withinTextRef, context) ?? '').toLowerCase();
    } else if (withinTextRef.startsWith('"') && withinTextRef.endsWith('"')) {
      withinText = withinTextRef.slice(1, -1).toLowerCase();
    } else {
      withinText = String(withinTextRef).toLowerCase();
    }
    
    let startNum = 1; // デフォルト値
    if (startNumRef) {
      if (startNumRef.match(/^[A-Z]+\d+$/)) {
        const cellValue = getCellValue(startNumRef, context);
        startNum = parseInt(String(cellValue ?? '1'));
      } else {
        startNum = parseInt(startNumRef);
      }
    }
    
    if (isNaN(startNum) || startNum < 1) {
      return FormulaError.VALUE;
    }
    
    // ワイルドカード処理（?と*）
    const pattern = findText
      .replace(/[.*+?^${}()|[\]\\]/g, '\\$&') // エスケープ
      .replace(/\\\?/g, '.') // ?を.に変換
      .replace(/\\\*/g, '.*'); // *を.*に変換
    
    const regex = new RegExp(pattern);
    const searchText = withinText.substring(startNum - 1);
    const match = searchText.match(regex);
    
    if (!match?.index) {
      return FormulaError.VALUE;
    }
    
    return startNum + match.index; // 1ベースで返す
  }
};

// TEXTJOIN関数の実装（手動実装が必要）
export const TEXTJOIN: CustomFormula = {
  name: 'TEXTJOIN',
  pattern: /TEXTJOIN\("([^"]*)",\s*(TRUE|FALSE),\s*([^)]+)\)/i,
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
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, textRef] = matches;
    
    let text: string;
    if (textRef.match(/^[A-Z]+\d+$/)) {
      text = String(getCellValue(textRef, context) ?? '');
    } else if (textRef.startsWith('"') && textRef.endsWith('"')) {
      text = textRef.slice(1, -1);
    } else {
      text = String(textRef);
    }
    
    // 各単語の先頭を大文字に、他を小文字に変換
    return text.toLowerCase().replace(/\b\w/g, (char) => char.toUpperCase());
  }
};

// VALUE関数の実装（文字列を数値に変換） - stringFunctions.tsから移動
export const VALUE: CustomFormula = {
  name: 'VALUE',
  pattern: /VALUE\(([^)]+)\)/i,
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
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, valueRef, formatRef] = matches;
    
    // 値を取得
    let value: number | string;
    if (valueRef.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(valueRef, context);
      value = (cellValue as number | string) ?? 0;
    } else {
      value = parseFloat(valueRef);
      if (isNaN(value)) {
        value = valueRef;
      }
    }
    
    // フォーマットを取得
    let format: string;
    if (formatRef.startsWith('"') && formatRef.endsWith('"')) {
      format = formatRef.slice(1, -1);
    } else if (formatRef.match(/^[A-Z]+\d+$/)) {
      format = String(getCellValue(formatRef, context) ?? '');
    } else {
      format = String(formatRef);
    }
    
    // 簡単なフォーマット処理
    if (typeof value === 'number') {
      if (format.includes('#,##0')) {
        return value.toLocaleString();
      } else if (format.includes('0.00')) {
        return value.toFixed(2);
      } else if (format.includes('0.0')) {
        return value.toFixed(1);
      } else if (format.includes('0')) {
        return Math.round(value).toString();
      } else if (format.includes('%')) {
        return (value * 100).toFixed(0) + '%';
      }
    }
    
    return String(value);
  }
};

// REPT関数の実装（文字列繰り返し） - stringFunctions.tsから移動
export const REPT: CustomFormula = {
  name: 'REPT',
  pattern: /REPT\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, textRef, numberTimesRef] = matches;
    
    let text: string;
    if (textRef.match(/^[A-Z]+\d+$/)) {
      text = String(getCellValue(textRef, context) ?? '');
    } else if (textRef.startsWith('"') && textRef.endsWith('"')) {
      text = textRef.slice(1, -1);
    } else {
      text = String(textRef);
    }
    
    let numberTimes: number;
    if (numberTimesRef.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(numberTimesRef, context);
      numberTimes = parseInt(String(cellValue ?? '0'));
    } else {
      numberTimes = parseInt(numberTimesRef);
    }
    
    if (isNaN(numberTimes) || numberTimes < 0) {
      return FormulaError.VALUE;
    }
    
    if (numberTimes === 0) {
      return '';
    }
    
    // 最大長を制限してメモリ不足を防ぐ
    if (text.length * numberTimes > 32767) {
      return FormulaError.VALUE;
    }
    
    return text.repeat(numberTimes);
  }
};

// REPLACE関数の実装（位置指定文字置換）
export const REPLACE: CustomFormula = {
  name: 'REPLACE',
  pattern: /REPLACE\(([^,]+),\s*([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
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
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, numberRef] = matches;
    
    let number: number;
    if (numberRef.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(numberRef, context);
      number = parseInt(String(cellValue ?? '0'));
    } else {
      number = parseInt(numberRef);
    }
    
    if (isNaN(number) || number < 1 || number > 255) {
      return FormulaError.VALUE;
    }
    
    return String.fromCharCode(number);
  }
};

// CODE関数の実装（文字から文字コードを返す）
export const CODE: CustomFormula = {
  name: 'CODE',
  pattern: /CODE\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, textRef] = matches;
    
    let text: string;
    if (textRef.match(/^[A-Z]+\d+$/)) {
      text = String(getCellValue(textRef, context) ?? '');
    } else if (textRef.startsWith('"') && textRef.endsWith('"')) {
      text = textRef.slice(1, -1);
    } else {
      text = String(textRef);
    }
    
    if (text.length === 0) {
      return FormulaError.VALUE;
    }
    
    return text.charCodeAt(0);
  }
};

// EXACT関数の実装（文字列が同一か判定）
export const EXACT: CustomFormula = {
  name: 'EXACT',
  pattern: /EXACT\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, text1Ref, text2Ref] = matches;
    
    let text1: string;
    if (text1Ref.match(/^[A-Z]+\d+$/)) {
      text1 = String(getCellValue(text1Ref, context) ?? '');
    } else if (text1Ref.startsWith('"') && text1Ref.endsWith('"')) {
      text1 = text1Ref.slice(1, -1);
    } else {
      text1 = String(text1Ref);
    }
    
    let text2: string;
    if (text2Ref.match(/^[A-Z]+\d+$/)) {
      text2 = String(getCellValue(text2Ref, context) ?? '');
    } else if (text2Ref.startsWith('"') && text2Ref.endsWith('"')) {
      text2 = text2Ref.slice(1, -1);
    } else {
      text2 = String(text2Ref);
    }
    
    return text1 === text2;
  }
};

// CLEAN関数の実装（印刷不可文字を削除）
export const CLEAN: CustomFormula = {
  name: 'CLEAN',
  pattern: /CLEAN\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, textRef] = matches;
    
    let text: string;
    if (textRef.match(/^[A-Z]+\d+$/)) {
      text = String(getCellValue(textRef, context) ?? '');
    } else if (textRef.startsWith('"') && textRef.endsWith('"')) {
      text = textRef.slice(1, -1);
    } else {
      text = String(textRef);
    }
    
    // 印刷不可文字（文字コード 0-31 および 127）を削除
    return text.split('').filter(char => {
      const code = char.charCodeAt(0);
      return !(code >= 0 && code <= 31) && code !== 127;
    }).join('');
  }
};

// T関数の実装（文字列を返す）
export const T: CustomFormula = {
  name: 'T',
  pattern: /T\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, valueRef] = matches;
    
    let value: unknown;
    if (valueRef.match(/^[A-Z]+\d+$/)) {
      value = getCellValue(valueRef, context);
    } else if (valueRef.startsWith('"') && valueRef.endsWith('"')) {
      value = valueRef.slice(1, -1);
    } else {
      value = valueRef;
    }
    
    // 文字列の場合はそのまま返し、それ以外は空文字列を返す
    if (typeof value === 'string') {
      return value;
    }
    
    return '';
  }
};

// FIXED関数の実装（固定小数点表示）
export const FIXED: CustomFormula = {
  name: 'FIXED',
  pattern: /FIXED\(([^,]+)(?:,\s*([^,]+))?(?:,\s*([^)]+))?\)/i,
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
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, textRef, decimalSeparatorRef, groupSeparatorRef] = matches;
    
    let text: string;
    if (textRef.match(/^[A-Z]+\d+$/)) {
      text = String(getCellValue(textRef, context) ?? '');
    } else if (textRef.startsWith('"') && textRef.endsWith('"')) {
      text = textRef.slice(1, -1);
    } else {
      text = String(textRef);
    }
    
    let decimalSeparator = '.';
    if (decimalSeparatorRef) {
      if (decimalSeparatorRef.startsWith('"') && decimalSeparatorRef.endsWith('"')) {
        decimalSeparator = decimalSeparatorRef.slice(1, -1);
      } else if (decimalSeparatorRef.match(/^[A-Z]+\d+$/)) {
        decimalSeparator = String(getCellValue(decimalSeparatorRef, context) ?? '.');
      } else {
        decimalSeparator = decimalSeparatorRef;
      }
    }
    
    let groupSeparator = ',';
    if (groupSeparatorRef) {
      if (groupSeparatorRef.startsWith('"') && groupSeparatorRef.endsWith('"')) {
        groupSeparator = groupSeparatorRef.slice(1, -1);
      } else if (groupSeparatorRef.match(/^[A-Z]+\d+$/)) {
        groupSeparator = String(getCellValue(groupSeparatorRef, context) ?? ',');
      } else {
        groupSeparator = groupSeparatorRef;
      }
    }
    
    // グループ区切り文字を削除し、小数点を正規化
    let normalizedText = text;
    if (groupSeparator) {
      normalizedText = normalizedText.replace(new RegExp('\\' + groupSeparator, 'g'), '');
    }
    if (decimalSeparator !== '.') {
      normalizedText = normalizedText.replace(decimalSeparator, '.');
    }
    
    // パーセント記号の処理
    let isPercentage = false;
    if (normalizedText.includes('%')) {
      isPercentage = true;
      normalizedText = normalizedText.replace('%', '');
    }
    
    const number = parseFloat(normalizedText);
    if (isNaN(number)) {
      return FormulaError.VALUE;
    }
    
    return isPercentage ? number / 100 : number;
  }
};

// DOLLAR関数の実装（通貨書式に変換）
export const DOLLAR: CustomFormula = {
  name: 'DOLLAR',
  pattern: /DOLLAR\(([^,]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, numberRef, decimalsRef] = matches;
    
    let number: number;
    if (numberRef.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(numberRef, context);
      number = parseFloat(String(cellValue ?? '0'));
    } else {
      number = parseFloat(numberRef);
    }
    
    let decimals = 2; // デフォルト値
    if (decimalsRef) {
      if (decimalsRef.match(/^[A-Z]+\d+$/)) {
        const cellValue = getCellValue(decimalsRef, context);
        decimals = parseInt(String(cellValue ?? '2'));
      } else {
        decimals = parseInt(decimalsRef);
      }
    }
    
    if (isNaN(number) || isNaN(decimals)) {
      return FormulaError.VALUE;
    }
    
    if (decimals < 0) {
      decimals = 0;
    }
    
    // ドルサインとカンマ区切り付きのフォーマット
    const formatted = number.toFixed(decimals);
    const parts = formatted.split('.');
    parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    
    return '$' + parts.join('.');
  }
};

// UNICHAR関数の実装（Unicode文字を返す）
export const UNICHAR: CustomFormula = {
  name: 'UNICHAR',
  pattern: /UNICHAR\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, numberRef] = matches;
    
    let number: number;
    if (numberRef.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(numberRef, context);
      number = parseInt(String(cellValue ?? '0'));
    } else {
      number = parseInt(numberRef);
    }
    
    if (isNaN(number) || number < 1 || number > 1114111) {
      return FormulaError.VALUE;
    }
    
    try {
      return String.fromCodePoint(number);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// UNICODE関数の実装（Unicode値を返す）
export const UNICODE: CustomFormula = {
  name: 'UNICODE',
  pattern: /UNICODE\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, textRef] = matches;
    
    let text: string;
    if (textRef.match(/^[A-Z]+\d+$/)) {
      text = String(getCellValue(textRef, context) ?? '');
    } else if (textRef.startsWith('"') && textRef.endsWith('"')) {
      text = textRef.slice(1, -1);
    } else {
      text = String(textRef);
    }
    
    if (text.length === 0) {
      return FormulaError.VALUE;
    }
    
    return text.codePointAt(0) ?? 0;
  }
};

// LENB関数の実装（バイト数を返す）
export const LENB: CustomFormula = {
  name: 'LENB',
  pattern: /LENB\(([^)]+)\)/i,
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
  calculate: JIS.calculate // JIS関数と同じ実装
};

// PHONETIC関数の実装（ふりがなを抽出）
export const PHONETIC: CustomFormula = {
  name: 'PHONETIC',
  pattern: /PHONETIC\(([^)]+)\)/i,
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