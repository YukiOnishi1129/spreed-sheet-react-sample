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
      const num1 = parseInt(getCellValue(num1Ref.trim(), context)?.toString() ?? num1Ref.trim());
      const num2 = parseInt(getCellValue(num2Ref.trim(), context)?.toString() ?? num2Ref.trim());
      
      if (isNaN(num1) || isNaN(num2)) {
        return FormulaError.VALUE;
      }
      
      // 負の数または大きすぎる数のチェック
      if (num1 < 0 || num2 < 0 || num1 > 281474976710655 || num2 > 281474976710655) {
        return FormulaError.NUM;
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
      const num1 = parseInt(getCellValue(num1Ref.trim(), context)?.toString() ?? num1Ref.trim());
      const num2 = parseInt(getCellValue(num2Ref.trim(), context)?.toString() ?? num2Ref.trim());
      
      if (isNaN(num1) || isNaN(num2)) {
        return FormulaError.VALUE;
      }
      
      // 負の数または大きすぎる数のチェック
      if (num1 < 0 || num2 < 0 || num1 > 281474976710655 || num2 > 281474976710655) {
        return FormulaError.NUM;
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
      const num1 = parseInt(getCellValue(num1Ref.trim(), context)?.toString() ?? num1Ref.trim());
      const num2 = parseInt(getCellValue(num2Ref.trim(), context)?.toString() ?? num2Ref.trim());
      
      if (isNaN(num1) || isNaN(num2)) {
        return FormulaError.VALUE;
      }
      
      // 負の数または大きすぎる数のチェック
      if (num1 < 0 || num2 < 0 || num1 > 281474976710655 || num2 > 281474976710655) {
        return FormulaError.NUM;
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
      const num = parseInt(getCellValue(numRef.trim(), context)?.toString() ?? numRef.trim());
      const shift = parseInt(getCellValue(shiftRef.trim(), context)?.toString() ?? shiftRef.trim());
      
      if (isNaN(num) || isNaN(shift)) {
        return FormulaError.VALUE;
      }
      
      // 負の数または大きすぎる数のチェック
      if (num < 0 || num > 281474976710655) {
        return FormulaError.NUM;
      }
      
      // シフト量の制限
      if (Math.abs(shift) > 53) {
        return FormulaError.NUM;
      }
      
      if (shift >= 0) {
        // 左シフト
        const result = num << shift;
        // オーバーフローチェック
        if (result > 281474976710655) {
          return FormulaError.NUM;
        }
        return result;
      } else {
        // 負の値の場合は右シフト
        return num >>> (-shift);
      }
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
      const num = parseInt(getCellValue(numRef.trim(), context)?.toString() ?? numRef.trim());
      const shift = parseInt(getCellValue(shiftRef.trim(), context)?.toString() ?? shiftRef.trim());
      
      if (isNaN(num) || isNaN(shift)) {
        return FormulaError.VALUE;
      }
      
      // 負の数または大きすぎる数のチェック
      if (num < 0 || num > 281474976710655) {
        return FormulaError.NUM;
      }
      
      // シフト量の制限
      if (Math.abs(shift) > 53) {
        return FormulaError.NUM;
      }
      
      if (shift >= 0) {
        // 右シフト（論理シフト）
        return num >>> shift;
      } else {
        // 負の値の場合は左シフト
        const result = num << (-shift);
        // オーバーフローチェック
        if (result > 281474976710655) {
          return FormulaError.NUM;
        }
        return result;
      }
    } catch {
      return FormulaError.VALUE;
    }
  }
};