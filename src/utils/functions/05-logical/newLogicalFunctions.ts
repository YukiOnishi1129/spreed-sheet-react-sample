// 新しい論理関数の実装（SWITCH, LET）

import type { CustomFormula, FormulaContext, FormulaResult } from '../shared/types';
import { FormulaError } from '../shared/types';
import { getCellValue } from '../shared/utils';

// SWITCH関数の実装（現代的なネストIFの代替）
export const SWITCH: CustomFormula = {
  name: 'SWITCH',
  pattern: /SWITCH\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
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
    
    if (args.length < 3) {
      return FormulaError.VALUE;
    }
    
    // 最初の引数は比較する式
    const expressionArg = args[0];
    let expressionValue: unknown;
    
    // セル参照の場合
    if (expressionArg.match(/^[A-Z]+\d+$/)) {
      expressionValue = getCellValue(expressionArg, context);
    } else if (expressionArg.startsWith('"') && expressionArg.endsWith('"')) {
      expressionValue = expressionArg.slice(1, -1);
    } else {
      const num = parseFloat(expressionArg);
      expressionValue = !isNaN(num) ? num : expressionArg;
    }
    
    // 値とケースのペアを評価
    let defaultValue: unknown = null;
    const hasDefault = args.length % 2 === 0; // 偶数個の引数の場合はデフォルト値あり
    
    if (hasDefault) {
      const defaultArg = args[args.length - 1];
      if (defaultArg.match(/^[A-Z]+\d+$/)) {
        defaultValue = getCellValue(defaultArg, context);
      } else if (defaultArg.startsWith('"') && defaultArg.endsWith('"')) {
        defaultValue = defaultArg.slice(1, -1);
      } else {
        const num = parseFloat(defaultArg);
        defaultValue = !isNaN(num) ? num : defaultArg;
      }
    }
    
    // ケースを確認（1番目から開始、ペアで処理）
    const casePairs = hasDefault ? args.slice(1, -1) : args.slice(1);
    
    for (let i = 0; i < casePairs.length; i += 2) {
      if (i + 1 >= casePairs.length) break;
      
      const caseArg = casePairs[i];
      const valueArg = casePairs[i + 1];
      
      let caseValue: unknown;
      
      // ケース値を取得
      if (caseArg.match(/^[A-Z]+\d+$/)) {
        caseValue = getCellValue(caseArg, context);
      } else if (caseArg.startsWith('"') && caseArg.endsWith('"')) {
        caseValue = caseArg.slice(1, -1);
      } else {
        const num = parseFloat(caseArg);
        caseValue = !isNaN(num) ? num : caseArg;
      }
      
      // 一致チェック
      if (String(expressionValue) === String(caseValue)) {
        // 戻り値を処理
        if (valueArg.match(/^[A-Z]+\d+$/)) {
          return getCellValue(valueArg, context) as FormulaResult;
        } else if (valueArg.startsWith('"') && valueArg.endsWith('"')) {
          return valueArg.slice(1, -1);
        } else {
          const num = parseFloat(valueArg);
          return !isNaN(num) ? num : valueArg;
        }
      }
    }
    
    // 一致するケースがない場合
    if (hasDefault) {
      return defaultValue as FormulaResult;
    }
    
    return FormulaError.NA;
  }
};

// LET関数の実装（変数を定義して式で利用）
export const LET: CustomFormula = {
  name: 'LET',
  pattern: /LET\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
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
    
    if (args.length < 3 || args.length % 2 === 0) {
      return FormulaError.VALUE; // 奇数個の引数が必要
    }
    
    // 変数の定義を処理
    const variables = new Map<string, unknown>();
    
    // 最後の引数は計算式、それ以外は変数名と値のペア
    const formula = args[args.length - 1];
    const varPairs = args.slice(0, -1);
    
    // 変数を定義
    for (let i = 0; i < varPairs.length; i += 2) {
      const varName = varPairs[i].trim();
      const varValueArg = varPairs[i + 1];
      
      let varValue: unknown;
      
      // 値を取得
      if (varValueArg.match(/^[A-Z]+\d+$/)) {
        varValue = getCellValue(varValueArg, context);
      } else if (varValueArg.startsWith('"') && varValueArg.endsWith('"')) {
        varValue = varValueArg.slice(1, -1);
      } else {
        const num = parseFloat(varValueArg);
        varValue = !isNaN(num) ? num : varValueArg;
      }
      
      variables.set(varName, varValue);
    }
    
    // 数式内の変数を置換
    let processedFormula = formula;
    
    for (const [varName, varValue] of variables) {
      // 変数名を値に置換（単語境界を考慮）
      const regex = new RegExp(`\\b${varName}\\b`, 'g');
      processedFormula = processedFormula.replace(regex, String(varValue));
    }
    
    // 簡単な数式評価（基本的な算術演算のみ対応）
    try {
      // セキュリティのため、eval は使わず簡単な演算のみ対応
      // 加算、減算、乗算、除算のみサポート
      const mathExpression = processedFormula.replace(/[^0-9+\-*/().\s]/g, '');
      
      if (mathExpression === processedFormula) {
        // 安全な数式の場合のみ評価
        const result = Function(`"use strict"; return (${mathExpression})`)();
        return result;
      } else {
        // セル参照が含まれる場合
        if (processedFormula.match(/^[A-Z]+\d+$/)) {
          return getCellValue(processedFormula, context) as FormulaResult;
        } else if (processedFormula.startsWith('"') && processedFormula.endsWith('"')) {
          return processedFormula.slice(1, -1);
        } else {
          const num = parseFloat(processedFormula);
          return !isNaN(num) ? num : processedFormula;
        }
      }
    } catch {
      return FormulaError.VALUE;
    }
  }
};