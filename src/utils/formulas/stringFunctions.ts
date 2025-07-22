// 文字列関数の実装

import type { CustomFormula } from './types';

// CONCATENATE関数の実装（文字列結合）
export const CONCATENATE: CustomFormula = {
  name: 'CONCATENATE',
  pattern: /CONCATENATE\(([^)]+)\)/i,
  isSupported: false,
  calculate: (matches) => {
    const args = matches[1].split(',').map(arg => arg.trim().replace(/"/g, ''));
    return args.join('');
  }
};

// LEFT関数の実装（左から文字を抽出）
export const LEFT: CustomFormula = {
  name: 'LEFT',
  pattern: /LEFT\(([^,]+)(?:,\s*([^)]+))?\)/i,
  isSupported: false,
  calculate: (matches) => {
    const text = matches[1].replace(/"/g, '');
    const numChars = matches[2] ? parseInt(matches[2]) : 1;
    
    if (isNaN(numChars) || numChars < 0) return '#VALUE!';
    return text.substring(0, numChars);
  }
};

// RIGHT関数の実装（右から文字を抽出）
export const RIGHT: CustomFormula = {
  name: 'RIGHT',
  pattern: /RIGHT\(([^,]+)(?:,\s*([^)]+))?\)/i,
  isSupported: false,
  calculate: (matches) => {
    const text = matches[1].replace(/"/g, '');
    const numChars = matches[2] ? parseInt(matches[2]) : 1;
    
    if (isNaN(numChars) || numChars < 0) return '#VALUE!';
    return text.substring(text.length - numChars);
  }
};

// MID関数の実装（中間の文字を抽出）
export const MID: CustomFormula = {
  name: 'MID',
  pattern: /MID\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  isSupported: false,
  calculate: (matches) => {
    const text = matches[1].replace(/"/g, '');
    const startPos = parseInt(matches[2]);
    const numChars = parseInt(matches[3]);
    
    if (isNaN(startPos) || isNaN(numChars) || startPos < 1 || numChars < 0) return '#VALUE!';
    return text.substring(startPos - 1, startPos - 1 + numChars);
  }
};

// LEN関数の実装（文字数）
export const LEN: CustomFormula = {
  name: 'LEN',
  pattern: /LEN\(([^)]+)\)/i,
  isSupported: false,
  calculate: (matches) => {
    const text = matches[1].replace(/"/g, '');
    return text.length;
  }
};

// UPPER関数の実装（大文字変換）
export const UPPER: CustomFormula = {
  name: 'UPPER',
  pattern: /UPPER\(([^)]+)\)/i,
  isSupported: false,
  calculate: (matches) => {
    const text = matches[1].replace(/"/g, '');
    return text.toUpperCase();
  }
};

// LOWER関数の実装（小文字変換）
export const LOWER: CustomFormula = {
  name: 'LOWER',
  pattern: /LOWER\(([^)]+)\)/i,
  isSupported: false,
  calculate: (matches) => {
    const text = matches[1].replace(/"/g, '');
    return text.toLowerCase();
  }
};

// PROPER関数の実装（先頭大文字変換）
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

// TRIM関数の実装（余分なスペース削除）
export const TRIM: CustomFormula = {
  name: 'TRIM',
  pattern: /TRIM\(([^)]+)\)/i,
  isSupported: false,
  calculate: (matches) => {
    const text = matches[1].replace(/"/g, '');
    return text.replace(/\s+/g, ' ').trim();
  }
};

// SUBSTITUTE関数の実装（文字置換）
export const SUBSTITUTE: CustomFormula = {
  name: 'SUBSTITUTE',
  pattern: /SUBSTITUTE\(([^,]+),\s*([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
  isSupported: false,
  calculate: (matches) => {
    const text = matches[1].replace(/"/g, '');
    const oldText = matches[2].replace(/"/g, '');
    const newText = matches[3].replace(/"/g, '');
    const instanceNum = matches[4] ? parseInt(matches[4]) : null;
    
    if (instanceNum !== null) {
      // 特定の出現箇所のみ置換
      if (isNaN(instanceNum) || instanceNum < 1) return '#VALUE!';
      
      let count = 0;
      return text.replace(new RegExp(oldText.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g'), (match) => {
        count++;
        return count === instanceNum ? newText : match;
      });
    } else {
      // すべて置換
      return text.replace(new RegExp(oldText.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g'), newText);
    }
  }
};

// FIND関数の実装（文字位置検索・大小区別あり）
export const FIND: CustomFormula = {
  name: 'FIND',
  pattern: /FIND\(([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
  isSupported: false,
  calculate: (matches) => {
    const findText = matches[1].replace(/"/g, '');
    const withinText = matches[2].replace(/"/g, '');
    const startNum = matches[3] ? parseInt(matches[3]) : 1;
    
    if (isNaN(startNum) || startNum < 1) return '#VALUE!';
    
    const searchFrom = startNum - 1;
    const position = withinText.indexOf(findText, searchFrom);
    
    return position === -1 ? '#VALUE!' : position + 1;
  }
};

// SEARCH関数の実装（文字位置検索・大小区別なし）
export const SEARCH: CustomFormula = {
  name: 'SEARCH',
  pattern: /SEARCH\(([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
  isSupported: false,
  calculate: (matches) => {
    const findText = matches[1].replace(/"/g, '').toLowerCase();
    const withinText = matches[2].replace(/"/g, '').toLowerCase();
    const startNum = matches[3] ? parseInt(matches[3]) : 1;
    
    if (isNaN(startNum) || startNum < 1) return '#VALUE!';
    
    const searchFrom = startNum - 1;
    const position = withinText.indexOf(findText, searchFrom);
    
    return position === -1 ? '#VALUE!' : position + 1;
  }
};

// VALUE関数の実装（文字列を数値に変換）
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

// TEXT関数の実装（数値を書式付き文字列に変換）
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

// REPT関数の実装（文字列繰り返し）
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