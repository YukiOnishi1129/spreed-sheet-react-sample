// 論理関数の実装

import type { CustomFormula, FormulaContext } from './types';
import { FormulaError } from './types';
import { getCellValue } from './utils';

// IF関数の実装
export const IF: CustomFormula = {
  name: 'IF',
  pattern: /IF\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// AND関数の実装
export const AND: CustomFormula = {
  name: 'AND',
  pattern: /AND\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// OR関数の実装
export const OR: CustomFormula = {
  name: 'OR',
  pattern: /OR\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// NOT関数の実装
export const NOT: CustomFormula = {
  name: 'NOT',
  pattern: /NOT\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// IFS関数の実装（手動実装が必要）
export const IFS: CustomFormula = {
  name: 'IFS',
  pattern: /IFS\(([^)]+)\)/i,
  isSupported: false, // HyperFormulaでサポートされていない可能性
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const argsString = matches[1];
    
    // 引数をパースする（カンマで分割、ただし関数内のカンマは無視）
    const args: string[] = [];
    let currentArg = '';
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
          args.push(currentArg.trim());
          currentArg = '';
          continue;
        }
      }
      
      currentArg += char;
    }
    
    if (currentArg.trim()) {
      args.push(currentArg.trim());
    }
    
    // 条件と値のペアを評価
    for (let i = 0; i < args.length; i += 2) {
      if (i + 1 >= args.length) break;
      
      const conditionArg = args[i];
      const valueArg = args[i + 1];
      
      // 条件を評価（簡単な実装）
      let conditionResult = false;
      
      // セル参照の場合
      if (conditionArg.match(/^[A-Z]+\d+$/)) {
        const cellValue = getCellValue(conditionArg, context);
        conditionResult = Boolean(cellValue);
      } else {
        // 簡単な比較演算子の処理
        const comparisonMatch = conditionArg.match(/([A-Z]+\d+|"[^"]*"|\d+)\s*([><=!]+)\s*([A-Z]+\d+|"[^"]*"|\d+)/);
        if (comparisonMatch) {
          let [, left, operator, right] = comparisonMatch;
          
          // セル参照の場合は値を取得
          if (left.match(/^[A-Z]+\d+$/)) {
            left = getCellValue(left, context);
          } else if (left.startsWith('"') && left.endsWith('"')) {
            left = left.slice(1, -1);
          }
          
          if (right.match(/^[A-Z]+\d+$/)) {
            right = getCellValue(right, context);
          } else if (right.startsWith('"') && right.endsWith('"')) {
            right = right.slice(1, -1);
          }
          
          // 数値比較の場合
          const leftNum = parseFloat(left);
          const rightNum = parseFloat(right);
          
          if (!isNaN(leftNum) && !isNaN(rightNum)) {
            switch (operator) {
              case '>': conditionResult = leftNum > rightNum; break;
              case '<': conditionResult = leftNum < rightNum; break;
              case '>=': conditionResult = leftNum >= rightNum; break;
              case '<=': conditionResult = leftNum <= rightNum; break;
              case '=': conditionResult = leftNum === rightNum; break;
              case '<>':
              case '!=': conditionResult = leftNum !== rightNum; break;
            }
          } else {
            // 文字列比較
            switch (operator) {
              case '=': conditionResult = String(left) === String(right); break;
              case '<>':
              case '!=': conditionResult = String(left) !== String(right); break;
            }
          }
        }
      }
      
      if (conditionResult) {
        // 戻り値を処理
        if (valueArg.match(/^[A-Z]+\d+$/)) {
          return getCellValue(valueArg, context);
        } else if (valueArg.startsWith('"') && valueArg.endsWith('"')) {
          return valueArg.slice(1, -1);
        } else {
          const num = parseFloat(valueArg);
          return !isNaN(num) ? num : valueArg;
        }
      }
    }
    
    return FormulaError.NA;
  }
};