// 数学関数の実装

import type { CustomFormula } from './types';
import { FormulaError } from './types';
import { getCellValue } from './utils';

// SUMIF関数の実装
export const SUMIF: CustomFormula = {
  name: 'SUMIF',
  pattern: /SUMIF\(([^,]+),\s*"([^"]+)",\s*([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// COUNTIF関数の実装
export const COUNTIF: CustomFormula = {
  name: 'COUNTIF',
  pattern: /COUNTIF\(([^,]+),\s*"([^"]+)"\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// AVERAGEIF関数の実装
export const AVERAGEIF: CustomFormula = {
  name: 'AVERAGEIF',
  pattern: /AVERAGEIF\(([^,]+),\s*"([^"]+)",\s*([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// SUM関数の実装
export const SUM: CustomFormula = {
  name: 'SUM',
  pattern: /SUM\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// AVERAGE関数の実装
export const AVERAGE: CustomFormula = {
  name: 'AVERAGE',
  pattern: /AVERAGE\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// COUNT関数の実装
export const COUNT: CustomFormula = {
  name: 'COUNT',
  pattern: /COUNT\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// MAX関数の実装
export const MAX: CustomFormula = {
  name: 'MAX',
  pattern: /MAX\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// MIN関数の実装
export const MIN: CustomFormula = {
  name: 'MIN',
  pattern: /MIN\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// ROUND関数の実装
export const ROUND: CustomFormula = {
  name: 'ROUND',
  pattern: /ROUND\(([^,]+),\s*(\d+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// ABS関数の実装（絶対値）
export const ABS: CustomFormula = {
  name: 'ABS',
  pattern: /ABS\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// SQRT関数の実装（平方根）
export const SQRT: CustomFormula = {
  name: 'SQRT',
  pattern: /SQRT\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// POWER関数の実装（べき乗）
export const POWER: CustomFormula = {
  name: 'POWER',
  pattern: /POWER\(([^,]+),\s*([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// MOD関数の実装（剰余）
export const MOD: CustomFormula = {
  name: 'MOD',
  pattern: /MOD\(([^,]+),\s*([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// INT関数の実装（整数部分）
export const INT: CustomFormula = {
  name: 'INT',
  pattern: /INT\(([^)]+)\)/i,
  isSupported: false,
  calculate: (matches, context) => {
    const valueRef = matches[1].trim();
    let value: number;
    
    if (valueRef.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(valueRef, context);
      value = parseFloat(String(cellValue ?? '0'));
    } else {
      value = parseFloat(valueRef);
    }
    
    if (isNaN(value)) return FormulaError.VALUE;
    return Math.floor(value);
  }
};

// TRUNC関数の実装（小数部分切り捨て）
export const TRUNC: CustomFormula = {
  name: 'TRUNC',
  pattern: /TRUNC\(([^,)]+)(?:,\s*([^)]+))?\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// RAND関数の実装（0以上1未満の乱数）
export const RAND: CustomFormula = {
  name: 'RAND',
  pattern: /RAND\(\)/i,
  isSupported: false,
  calculate: () => {
    return Math.random();
  }
};

// RANDBETWEEN関数の実装（指定範囲の整数乱数）
export const RANDBETWEEN: CustomFormula = {
  name: 'RANDBETWEEN',
  pattern: /RANDBETWEEN\(([^,]+),\s*([^)]+)\)/i,
  isSupported: false,
  calculate: (matches, context) => {
    const minRef = matches[1].trim();
    const maxRef = matches[2].trim();
    
    let min: number, max: number;
    
    if (minRef.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(minRef, context);
      min = parseInt(String(cellValue ?? '0'));
    } else {
      min = parseInt(minRef);
    }
    
    if (maxRef.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(maxRef, context);
      max = parseInt(String(cellValue ?? '0'));
    } else {
      max = parseInt(maxRef);
    }
    
    if (isNaN(min) || isNaN(max)) return FormulaError.VALUE;
    if (min > max) return FormulaError.NUM;
    return Math.floor(Math.random() * (max - min + 1)) + min;
  }
};

// PI関数の実装（円周率）
export const PI: CustomFormula = {
  name: 'PI',
  pattern: /PI\(\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// DEGREES関数の実装（ラジアンを度に変換）
export const DEGREES: CustomFormula = {
  name: 'DEGREES',
  pattern: /DEGREES\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// RADIANS関数の実装（度をラジアンに変換）
export const RADIANS: CustomFormula = {
  name: 'RADIANS',
  pattern: /RADIANS\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// SIN関数の実装（正弦）
export const SIN: CustomFormula = {
  name: 'SIN',
  pattern: /SIN\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// COS関数の実装（余弦）
export const COS: CustomFormula = {
  name: 'COS',
  pattern: /COS\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// TAN関数の実装（正接）
export const TAN: CustomFormula = {
  name: 'TAN',
  pattern: /TAN\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// LOG関数の実装（対数）
export const LOG: CustomFormula = {
  name: 'LOG',
  pattern: /LOG\(([^,)]+)(?:,\s*([^)]+))?\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// LOG10関数の実装（常用対数）
export const LOG10: CustomFormula = {
  name: 'LOG10',
  pattern: /LOG10\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// LN関数の実装（自然対数）
export const LN: CustomFormula = {
  name: 'LN',
  pattern: /LN\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// EXP関数の実装（指数関数）
export const EXP: CustomFormula = {
  name: 'EXP',
  pattern: /EXP\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// ASIN関数の実装（逆正弦）
export const ASIN: CustomFormula = {
  name: 'ASIN',
  pattern: /ASIN\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// ACOS関数の実装（逆余弦）
export const ACOS: CustomFormula = {
  name: 'ACOS',
  pattern: /ACOS\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// ATAN関数の実装（逆正接）
export const ATAN: CustomFormula = {
  name: 'ATAN',
  pattern: /ATAN\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// ATAN2関数の実装（x,y座標から角度）
export const ATAN2: CustomFormula = {
  name: 'ATAN2',
  pattern: /ATAN2\(([^,]+),\s*([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// ROUNDUP関数の実装（切り上げ）
export const ROUNDUP: CustomFormula = {
  name: 'ROUNDUP',
  pattern: /ROUNDUP\(([^,]+),\s*([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// ROUNDDOWN関数の実装（切り下げ）
export const ROUNDDOWN: CustomFormula = {
  name: 'ROUNDDOWN',
  pattern: /ROUNDDOWN\(([^,]+),\s*([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// CEILING関数の実装（基準値の倍数に切り上げ）
export const CEILING: CustomFormula = {
  name: 'CEILING',
  pattern: /CEILING\(([^,]+),\s*([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// FLOOR関数の実装（基準値の倍数に切り下げ）
export const FLOOR: CustomFormula = {
  name: 'FLOOR',
  pattern: /FLOOR\(([^,]+),\s*([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// SIGN関数の実装（符号）
export const SIGN: CustomFormula = {
  name: 'SIGN',
  pattern: /SIGN\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// FACT関数の実装（階乗）
export const FACT: CustomFormula = {
  name: 'FACT',
  pattern: /FACT\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// SUMIFS関数の実装（複数条件での合計）
export const SUMIFS: CustomFormula = {
  name: 'SUMIFS',
  pattern: /SUMIFS\(([^,]+),\s*([^,]+),\s*"([^"]+)"(?:,\s*([^,]+),\s*"([^"]+)")*\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// COUNTIFS関数の実装（複数条件でのカウント）
export const COUNTIFS: CustomFormula = {
  name: 'COUNTIFS',
  pattern: /COUNTIFS\(([^,]+),\s*"([^"]+)"(?:,\s*([^,]+),\s*"([^"]+)")*\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// AVERAGEIFS関数の実装（複数条件での平均）
export const AVERAGEIFS: CustomFormula = {
  name: 'AVERAGEIFS',
  pattern: /AVERAGEIFS\(([^,]+),\s*([^,]+),\s*"([^"]+)"(?:,\s*([^,]+),\s*"([^"]+)")*\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// PRODUCT関数の実装（積を計算）
export const PRODUCT: CustomFormula = {
  name: 'PRODUCT',
  pattern: /PRODUCT\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// MROUND関数の実装（倍数に丸める）
export const MROUND: CustomFormula = {
  name: 'MROUND',
  pattern: /MROUND\(([^,]+),\s*([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// COMBIN関数の実装（組み合わせ数）
export const COMBIN: CustomFormula = {
  name: 'COMBIN',
  pattern: /COMBIN\(([^,]+),\s*([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// PERMUT関数の実装（順列数）
export const PERMUT: CustomFormula = {
  name: 'PERMUT',
  pattern: /PERMUT\(([^,]+),\s*([^)]+)\)/i,
  isSupported: false,
  calculate: (matches, context) => {
    const nRef = matches[1].trim();
    const kRef = matches[2].trim();
    
    let n: number, k: number;
    
    // n値を取得
    if (nRef.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(nRef, context);
      n = parseInt(String(cellValue ?? '0'));
    } else {
      n = parseInt(nRef);
    }
    
    // k値を取得
    if (kRef.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(kRef, context);
      k = parseInt(String(cellValue ?? '0'));
    } else {
      k = parseInt(kRef);
    }
    
    if (isNaN(n) || isNaN(k)) return FormulaError.VALUE;
    if (n < 0 || k < 0) return FormulaError.NUM;
    if (k > n) return 0;
    
    let result = 1;
    for (let i = 0; i < k; i++) {
      result *= (n - i);
    }
    return result;
  }
};

// GCD関数の実装（最大公約数）
export const GCD: CustomFormula = {
  name: 'GCD',
  pattern: /GCD\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// LCM関数の実装（最小公倍数）
export const LCM: CustomFormula = {
  name: 'LCM',
  pattern: /LCM\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaに処理を委譲
};

// QUOTIENT関数の実装（商の整数部分）
export const QUOTIENT: CustomFormula = {
  name: 'QUOTIENT',
  pattern: /QUOTIENT\(([^,]+),\s*([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaに処理を委譲
};

// 双曲線関数の実装
// SINH関数（双曲線正弦）
export const SINH: CustomFormula = {
  name: 'SINH',
  pattern: /SINH\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// COSH関数（双曲線余弦）
export const COSH: CustomFormula = {
  name: 'COSH',
  pattern: /COSH\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// TANH関数（双曲線正接）
export const TANH: CustomFormula = {
  name: 'TANH',
  pattern: /TANH\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// HyperFormulaでサポートされている未実装関数を追加

// SUMSQ関数（平方和）
export const SUMSQ: CustomFormula = {
  name: 'SUMSQ',
  pattern: /SUMSQ\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaに処理を委譲
};

// SUMPRODUCT関数（配列の積の和）
export const SUMPRODUCT: CustomFormula = {
  name: 'SUMPRODUCT',
  pattern: /SUMPRODUCT\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaに処理を委譲
};

// EVEN関数（最も近い偶数に切り上げ）
export const EVEN: CustomFormula = {
  name: 'EVEN',
  pattern: /EVEN\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaに処理を委譲
};

// ODD関数（最も近い奇数に切り上げ）
export const ODD: CustomFormula = {
  name: 'ODD',
  pattern: /ODD\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaに処理を委譲
};

// ARABIC関数（ローマ数字をアラビア数字に変換）
export const ARABIC: CustomFormula = {
  name: 'ARABIC',
  pattern: /ARABIC\(([^)]+)\)/i,
  isSupported: false,
  calculate: (matches: RegExpMatchArray) => {
    const romanStr = matches[1].trim().replace(/["']/g, '').toUpperCase();
    
    // ローマ数字の基本文字と対応する値
    const romanMap: { [key: string]: number } = {
      'I': 1, 'V': 5, 'X': 10, 'L': 50,
      'C': 100, 'D': 500, 'M': 1000
    };
    
    // 無効なローマ数字文字をチェック
    if (!/^[IVXLCDM]+$/.test(romanStr)) {
      return FormulaError.VALUE;
    }
    
    let result = 0;
    let prev = 0;
    
    // 右から左へ処理
    for (let i = romanStr.length - 1; i >= 0; i--) {
      const current = romanMap[romanStr[i]];
      
      if (current === undefined) {
        return FormulaError.VALUE;
      }
      
      // 前の文字より小さい場合は減算、そうでなければ加算
      if (current < prev) {
        result -= current;
      } else {
        result += current;
      }
      
      prev = current;
    }
    
    return result;
  }
};

// ROMAN関数（アラビア数字をローマ数字に変換）
export const ROMAN: CustomFormula = {
  name: 'ROMAN',
  pattern: /ROMAN\(([^,)]+)(?:,\s*([^)]+))?\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaに処理を委譲
};

// COMBINA関数（重複組合せ）
export const COMBINA: CustomFormula = {
  name: 'COMBINA',
  pattern: /COMBINA\(([^,]+),\s*([^)]+)\)/i,
  isSupported: false,
  calculate: (matches: RegExpMatchArray) => {
    const nStr = matches[1].trim();
    const kStr = matches[2].trim();
    
    // 数値の取得
    const n = parseFloat(nStr);
    const k = parseFloat(kStr);
    
    // 数値型チェック
    if (typeof n !== 'number' || typeof k !== 'number') {
      return FormulaError.VALUE;
    }
    
    // 整数チェック
    if (!Number.isInteger(n) || !Number.isInteger(k)) {
      return FormulaError.NUM;
    }
    
    // 負数チェック
    if (n < 0 || k < 0) {
      return FormulaError.NUM;
    }
    
    // k = 0の場合
    if (k === 0) {
      return 1;
    }
    
    // 重複組合せの公式: C(n+k-1, k) = (n+k-1)! / (k! * (n-1)!)
    // これは C(n+k-1, k) と同じ
    const numerator = n + k - 1;
    
    if (numerator < k) {
      return 0;
    }
    
    // 効率的な組合せ計算
    let result = 1;
    for (let i = 0; i < Math.min(k, numerator - k); i++) {
      result = result * (numerator - i) / (i + 1);
    }
    
    return Math.round(result);
  }
};

// FACTDOUBLE関数（二重階乗）
export const FACTDOUBLE: CustomFormula = {
  name: 'FACTDOUBLE',
  pattern: /FACTDOUBLE\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaに処理を委譲
};

// SQRTPI関数（π倍の平方根）
export const SQRTPI: CustomFormula = {
  name: 'SQRTPI',
  pattern: /SQRTPI\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaに処理を委譲
};

// SUMX2MY2関数（x^2-y^2の和）
export const SUMX2MY2: CustomFormula = {
  name: 'SUMX2MY2',
  pattern: /SUMX2MY2\(([^,]+),\s*([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaに処理を委譲
};

// SUMX2PY2関数（x^2+y^2の和）
export const SUMX2PY2: CustomFormula = {
  name: 'SUMX2PY2',
  pattern: /SUMX2PY2\(([^,]+),\s*([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaに処理を委譲
};

// SUMXMY2関数（(x-y)^2の和）
export const SUMXMY2: CustomFormula = {
  name: 'SUMXMY2',
  pattern: /SUMXMY2\(([^,]+),\s*([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaに処理を委譲
};

// MULTINOMIAL関数（多項係数）
export const MULTINOMIAL: CustomFormula = {
  name: 'MULTINOMIAL',
  pattern: /MULTINOMIAL\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaに処理を委譲
};

// PERMUTATIONA関数（重複順列）
export const PERMUTATIONA: CustomFormula = {
  name: 'PERMUTATIONA',
  pattern: /PERMUTATIONA\(([^,]+),\s*([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaに処理を委譲
};

// BASE関数（数値を指定した基数に変換）
export const BASE: CustomFormula = {
  name: 'BASE',
  pattern: /BASE\(([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaに処理を委譲
};

// DECIMAL関数（指定した基数の数値を10進数に変換）
export const DECIMAL: CustomFormula = {
  name: 'DECIMAL',
  pattern: /DECIMAL\(([^,]+),\s*([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaに処理を委譲
};

// SUBTOTAL関数（小計）
export const SUBTOTAL: CustomFormula = {
  name: 'SUBTOTAL',
  pattern: /SUBTOTAL\(([^,]+),\s*([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaに処理を委譲
};

// AGGREGATE関数（集計関数、エラー値を除外）
export const AGGREGATE: CustomFormula = {
  name: 'AGGREGATE',
  pattern: /AGGREGATE\(([^,]+),\s*([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaに処理を委譲
};

// CEILING.MATH関数（数学的な切り上げ）
export const CEILING_MATH: CustomFormula = {
  name: 'CEILING.MATH',
  pattern: /CEILING\.MATH\(([^,)]+)(?:,\s*([^,)]+))?(?:,\s*([^)]+))?\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaに処理を委譲
};

// CEILING.PRECISE関数（精密な切り上げ）
export const CEILING_PRECISE: CustomFormula = {
  name: 'CEILING.PRECISE',
  pattern: /CEILING\.PRECISE\(([^,)]+)(?:,\s*([^)]+))?\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaに処理を委譲
};

// FLOOR.MATH関数（数学的な切り下げ）
export const FLOOR_MATH: CustomFormula = {
  name: 'FLOOR.MATH',
  pattern: /FLOOR\.MATH\(([^,)]+)(?:,\s*([^,)]+))?(?:,\s*([^)]+))?\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaに処理を委譲
};

// FLOOR.PRECISE関数（精密な切り下げ）
export const FLOOR_PRECISE: CustomFormula = {
  name: 'FLOOR.PRECISE',
  pattern: /FLOOR\.PRECISE\(([^,)]+)(?:,\s*([^)]+))?\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaに処理を委譲
};

// ISO.CEILING関数（ISO標準の切り上げ）
export const ISO_CEILING: CustomFormula = {
  name: 'ISO.CEILING',
  pattern: /ISO\.CEILING\(([^,)]+)(?:,\s*([^)]+))?\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaに処理を委譲
};

// SERIESSUM関数（べき級数を計算）
export const SERIESSUM: CustomFormula = {
  name: 'SERIESSUM',
  pattern: /SERIESSUM\(([^,]+),\s*([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaに処理を委譲
};

// RANDARRAY関数（ランダム配列を生成）
export const RANDARRAY: CustomFormula = {
  name: 'RANDARRAY',
  pattern: /RANDARRAY\((?:([^,)]+))?(?:,\s*([^,)]+))?(?:,\s*([^,)]+))?(?:,\s*([^,)]+))?(?:,\s*([^)]+))?\)/i,
  isSupported: false, // HyperFormulaでサポートされていない（手動実装）
  calculate: (matches) => {
    const rows = matches[1] ? parseInt(matches[1]) : 1;
    const cols = matches[2] ? parseInt(matches[2]) : 1;
    const min = matches[3] ? parseFloat(matches[3]) : 0;
    const max = matches[4] ? parseFloat(matches[4]) : 1;
    const wholeNumber = matches[5] ? matches[5].toLowerCase() === 'true' : false;
    
    if (isNaN(rows) || isNaN(cols) || rows <= 0 || cols <= 0) return FormulaError.VALUE;
    if (isNaN(min) || isNaN(max) || min >= max) return FormulaError.VALUE;
    
    const result: number[][] = [];
    for (let i = 0; i < rows; i++) {
      const row: number[] = [];
      for (let j = 0; j < cols; j++) {
        const randomValue = Math.random() * (max - min) + min;
        row.push(wholeNumber ? Math.floor(randomValue) : randomValue);
      }
      result.push(row);
    }
    
    // 1x1の場合は単一の値を返す
    if (rows === 1 && cols === 1) {
      return result[0][0];
    }
    
    return result;
  }
};

// SEQUENCE関数（連続値を生成）
export const SEQUENCE: CustomFormula = {
  name: 'SEQUENCE',
  pattern: /SEQUENCE\(([^,)]+)(?:,\s*([^,)]+))?(?:,\s*([^,)]+))?(?:,\s*([^)]+))?\)/i,
  isSupported: false, // HyperFormulaでサポートされていない（手動実装）
  calculate: (matches) => {
    const rows = parseInt(matches[1]);
    const cols = matches[2] ? parseInt(matches[2]) : 1;
    const start = matches[3] ? parseFloat(matches[3]) : 1;
    const step = matches[4] ? parseFloat(matches[4]) : 1;
    
    if (isNaN(rows) || isNaN(cols) || rows <= 0 || cols <= 0) return FormulaError.VALUE;
    if (isNaN(start) || isNaN(step)) return FormulaError.VALUE;
    
    const result: number[][] = [];
    let currentValue = start;
    
    for (let i = 0; i < rows; i++) {
      const row: number[] = [];
      for (let j = 0; j < cols; j++) {
        row.push(currentValue);
        currentValue += step;
      }
      result.push(row);
    }
    
    // 1x1の場合は単一の値を返す
    if (rows === 1 && cols === 1) {
      return result[0][0];
    }
    
    return result;
  }
};