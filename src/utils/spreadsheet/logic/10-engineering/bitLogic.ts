// ビット演算関数の実装

import type { CustomFormula, FormulaContext, FormulaResult } from '../shared/types';
import { FormulaError } from '../shared/types';
import { getCellValue } from '../shared/utils';

// BITAND関数の実装（ビット単位AND）
export const BITAND: CustomFormula = {
  name: 'BITAND',
  pattern: /BITAND\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, num1Ref, num2Ref] = matches;
    
    try {
      const num1 = Math.trunc(parseFloat(getCellValue(num1Ref.trim(), context)?.toString() ?? num1Ref.trim()));
      const num2 = Math.trunc(parseFloat(getCellValue(num2Ref.trim(), context)?.toString() ?? num2Ref.trim()));
      
      if (isNaN(num1) || isNaN(num2)) {
        return FormulaError.VALUE;
      }
      
      // 負の数または大きすぎる数のチェック
      if (num1 < 0 || num2 < 0 || num1 > 281474976710655 || num2 > 281474976710655) {
        return FormulaError.NUM;
      }
      
      // JavaScriptのビット演算子は32ビットなので、BigIntを使用
      if (num1 > 0xFFFFFFFF || num2 > 0xFFFFFFFF) {
        return Number(BigInt(num1) & BigInt(num2));
      }
      return num1 & num2;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// BITOR関数の実装（ビット単位OR）
export const BITOR: CustomFormula = {
  name: 'BITOR',
  pattern: /BITOR\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, num1Ref, num2Ref] = matches;
    
    try {
      const num1 = Math.trunc(parseFloat(getCellValue(num1Ref.trim(), context)?.toString() ?? num1Ref.trim()));
      const num2 = Math.trunc(parseFloat(getCellValue(num2Ref.trim(), context)?.toString() ?? num2Ref.trim()));
      
      if (isNaN(num1) || isNaN(num2)) {
        return FormulaError.VALUE;
      }
      
      // 負の数または大きすぎる数のチェック
      if (num1 < 0 || num2 < 0 || num1 > 281474976710655 || num2 > 281474976710655) {
        return FormulaError.NUM;
      }
      
      // JavaScriptのビット演算子は32ビットなので、BigIntを使用
      if (num1 > 0xFFFFFFFF || num2 > 0xFFFFFFFF) {
        return Number(BigInt(num1) | BigInt(num2));
      }
      return num1 | num2;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// BITXOR関数の実装（ビット単位XOR）
export const BITXOR: CustomFormula = {
  name: 'BITXOR',
  pattern: /BITXOR\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, num1Ref, num2Ref] = matches;
    
    try {
      const num1 = Math.trunc(parseFloat(getCellValue(num1Ref.trim(), context)?.toString() ?? num1Ref.trim()));
      const num2 = Math.trunc(parseFloat(getCellValue(num2Ref.trim(), context)?.toString() ?? num2Ref.trim()));
      
      if (isNaN(num1) || isNaN(num2)) {
        return FormulaError.VALUE;
      }
      
      // 負の数または大きすぎる数のチェック
      if (num1 < 0 || num2 < 0 || num1 > 281474976710655 || num2 > 281474976710655) {
        return FormulaError.NUM;
      }
      
      // JavaScriptのビット演算子は32ビットなので、BigIntを使用
      if (num1 > 0xFFFFFFFF || num2 > 0xFFFFFFFF) {
        return Number(BigInt(num1) ^ BigInt(num2));
      }
      return num1 ^ num2;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// BITLSHIFT関数の実装（ビット左シフト）
export const BITLSHIFT: CustomFormula = {
  name: 'BITLSHIFT',
  pattern: /BITLSHIFT\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, numRef, shiftRef] = matches;
    
    try {
      const num = Math.trunc(parseFloat(getCellValue(numRef.trim(), context)?.toString() ?? numRef.trim()));
      const shift = Math.trunc(parseFloat(getCellValue(shiftRef.trim(), context)?.toString() ?? shiftRef.trim()));
      
      if (isNaN(num) || isNaN(shift)) {
        return FormulaError.VALUE;
      }
      
      // 負の数または大きすぎる数のチェック
      if (num < 0 || num > 281474976710655) {
        return FormulaError.NUM;
      }
      
      // シフト量の制限
      if (shift < 0 || shift > 53) {
        return FormulaError.NUM;
      }
      
      // 左シフト
      const result = num * Math.pow(2, shift);
      // オーバーフローチェック
      if (result > 281474976710655) {
        return FormulaError.NUM;
      }
      return Math.trunc(result);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// BITRSHIFT関数の実装（ビット右シフト）
export const BITRSHIFT: CustomFormula = {
  name: 'BITRSHIFT',
  pattern: /BITRSHIFT\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, numRef, shiftRef] = matches;
    
    try {
      const num = Math.trunc(parseFloat(getCellValue(numRef.trim(), context)?.toString() ?? numRef.trim()));
      const shift = Math.trunc(parseFloat(getCellValue(shiftRef.trim(), context)?.toString() ?? shiftRef.trim()));
      
      if (isNaN(num) || isNaN(shift)) {
        return FormulaError.VALUE;
      }
      
      // 負の数または大きすぎる数のチェック
      if (num < 0 || num > 281474976710655) {
        return FormulaError.NUM;
      }
      
      // シフト量の制限
      if (shift < 0 || shift > 53) {
        return FormulaError.NUM;
      }
      
      // 右シフト（論理シフト）
      return Math.trunc(num / Math.pow(2, shift));
    } catch {
      return FormulaError.VALUE;
    }
  }
};