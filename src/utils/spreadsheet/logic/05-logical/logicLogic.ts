// 論理関数の実装

import type { CustomFormula, FormulaContext, FormulaResult } from '../shared/types';
import { FormulaError } from '../shared/types';
import { getCellValue, evaluateExpression, evaluateFormulaWithErrorCheck } from '../shared/utils';
import { matchFormula } from '../index';

// 引数を論理値に変換するユーティリティ関数
function parseArgumentsToLogical(args: string, context: FormulaContext): boolean[] {
  const values: boolean[] = [];
  
  // 引数をカンマで分割
  const parts = args.split(',').map(part => part.trim());
  
  for (const part of parts) {
    const cellValue = getCellValue(part, context) ?? part;
    
    // 数値の場合：0はfalse、それ以外はtrue
    if (typeof cellValue === 'number') {
      values.push(cellValue !== 0);
    }
    // 文字列の場合："TRUE"/"FALSE"または"1"/"0"をチェック
    else if (typeof cellValue === 'string') {
      const str = cellValue.toLowerCase();
      if (str === 'true' || str === '1') {
        values.push(true);
      } else if (str === 'false' || str === '0') {
        values.push(false);
      } else {
        // 文字列は一般的にtrueとして扱う（空文字列以外）
        values.push(cellValue !== '');
      }
    }
    // 論理値の場合
    else if (typeof cellValue === 'boolean') {
      values.push(cellValue);
    }
    // nullまたはundefinedの場合
    else if (cellValue === null || cellValue === undefined) {
      values.push(true); // nullをtruthyとして扱う
    }
    else {
      values.push(Boolean(cellValue));
    }
  }
  
  return values;
}

// 単一の値を論理値に変換
function toLogical(value: unknown): boolean {
  if (typeof value === 'number') {
    return value !== 0;
  }
  if (typeof value === 'string') {
    const str = value.toLowerCase();
    if (str === 'true' || str === '1') return true;
    if (str === 'false' || str === '0') return false;
    return value !== '';
  }
  if (typeof value === 'boolean') {
    return value;
  }
  if (value === null || value === undefined) {
    return true; // nullをtruthyとして扱う
  }
  return Boolean(value);
}

// IF関数の結果を評価する改善された関数
function evaluateFormulaResult(resultStr: string, context: FormulaContext): FormulaResult {
  const trimmed = resultStr.trim();
  
  // 引用符付き文字列の場合
  if ((trimmed.startsWith('"') && trimmed.endsWith('"')) ||
      (trimmed.startsWith("'") && trimmed.endsWith("'"))) {
    return trimmed.slice(1, -1);
  }
  
  // エスケープされた引用符の場合
  if (trimmed.startsWith('\\"') && trimmed.endsWith('\\"')) {
    return trimmed.slice(2, -2);
  }
  
  // セル参照の場合
  if (trimmed.match(/^[A-Z]+\d+$/)) {
    const cellValue = getCellValue(trimmed, context);
    return cellValue as FormulaResult;
  }
  
  // 数値の場合
  const num = parseFloat(trimmed);
  if (!isNaN(num) && /^[+-]?\d*\.?\d*([eE][+-]?\d+)?$/.test(trimmed)) {
    return num;
  }
  
  // 論理値の場合
  if (trimmed.toLowerCase() === 'true') return true;
  if (trimmed.toLowerCase() === 'false') return false;
  
  // 関数呼び出しの場合（ネストされた関数）
  if (trimmed.match(/^[A-Z_]+\s*\(/i)) {
    try {
      const result = matchFormula(trimmed, context);
      if (result === null || result === undefined) {
        return FormulaError.VALUE;
      }
      return result;
    } catch (error) {
      // 関数評価エラーの場合、エラーを伝播
      if (typeof error === 'object' && error !== null && 'toString' in error) {
        const errorStr = error.toString();
        if (errorStr.includes('VALUE')) return FormulaError.VALUE;
        if (errorStr.includes('REF')) return FormulaError.REF;
        if (errorStr.includes('NUM')) return FormulaError.NUM;
        if (errorStr.includes('DIV')) return FormulaError.DIV0;
        if (errorStr.includes('NA')) return FormulaError.NA;
      }
      return FormulaError.VALUE;
    }
  }
  
  // 算術式の場合
  if (trimmed.match(/[+\-*/()]/)) {
    try {
      const result = evaluateExpression(trimmed, context);
      return result;
    } catch (error) {
      throw FormulaError.VALUE;
    }
  }
  
  // その他の場合は文字列として返す
  return trimmed;
}

// IF関数の実装
export const IF: CustomFormula = {
  name: 'IF',
  pattern: /IF\s*\((.+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const fullArgs = matches[1];
    
    // 引数を適切に分割する（引用符内のカンマを無視）
    const args: string[] = [];
    let currentArg = '';
    let inQuotes = false;
    let quoteChar = '';
    let parenDepth = 0;
    
    for (let i = 0; i < fullArgs.length; i++) {
      const char = fullArgs[i];
      const prevChar = i > 0 ? fullArgs[i - 1] : '';
      
      // エスケープされた引用符を処理
      if (char === '"' && prevChar !== '\\') {
        if (!inQuotes) {
          inQuotes = true;
          quoteChar = '"';
        } else if (quoteChar === '"') {
          inQuotes = false;
        }
      }
      
      // 括弧の深さを追跡
      if (!inQuotes) {
        if (char === '(') parenDepth++;
        else if (char === ')') parenDepth--;
      }
      
      // カンマで引数を分割（引用符内や括弧内でない場合）
      if (char === ',' && !inQuotes && parenDepth === 0) {
        args.push(currentArg.trim());
        currentArg = '';
      } else {
        currentArg += char;
      }
    }
    
    // 最後の引数を追加
    if (currentArg.trim()) {
      args.push(currentArg.trim());
    }
    
    // IF関数は3つの引数が必要
    if (args.length !== 3) {
      return FormulaError.VALUE;
    }
    
    const [conditionStr, trueValueStr, falseValueStr] = args;
    
    // 条件を評価（比較演算子を含む場合の処理）
    let condition: boolean;
    
    // 比較演算子のパターン
    const comparisonMatch = conditionStr.match(/^(.+?)(>=|<=|<>|!=|>|<|=)(.+)$/);
    
    if (comparisonMatch) {
      const [, leftExpr, operator, rightExpr] = comparisonMatch;
      
      // 左辺と右辺の値を取得
      let leftValue: string | number | unknown;
      let rightValue: string | number | unknown;
      
      // 左辺と右辺の値を改善された方法で取得
      try {
        leftValue = evaluateFormulaResult(leftExpr.trim(), context);
      } catch (error) {
        leftValue = leftExpr.trim();
      }
      
      try {
        rightValue = evaluateFormulaResult(rightExpr.trim(), context);
      } catch (error) {
        rightValue = rightExpr.trim();
      }
      
      // 数値に変換を試みる
      const leftNum = typeof leftValue === 'number' ? leftValue : parseFloat(String(leftValue));
      const rightNum = typeof rightValue === 'number' ? rightValue : parseFloat(String(rightValue));
      
      // 比較演算を実行
      switch (operator) {
        case '>=':
          condition = !isNaN(leftNum) && !isNaN(rightNum) ? leftNum >= rightNum : String(leftValue) >= String(rightValue);
          break;
        case '<=':
          condition = !isNaN(leftNum) && !isNaN(rightNum) ? leftNum <= rightNum : String(leftValue) <= String(rightValue);
          break;
        case '>':
          condition = !isNaN(leftNum) && !isNaN(rightNum) ? leftNum > rightNum : String(leftValue) > String(rightValue);
          break;
        case '<':
          condition = !isNaN(leftNum) && !isNaN(rightNum) ? leftNum < rightNum : String(leftValue) < String(rightValue);
          break;
        case '<>':
        case '!=':
          condition = !isNaN(leftNum) && !isNaN(rightNum) ? leftNum !== rightNum : leftValue !== rightValue;
          break;
        case '=':
          condition = !isNaN(leftNum) && !isNaN(rightNum) ? leftNum === rightNum : leftValue === rightValue;
          break;
        default:
          condition = false;
      }
    } else {
      // 比較演算子がない場合は、より包括的に評価
      try {
        const conditionResult = evaluateFormulaResult(conditionStr, context);
        condition = toLogical(conditionResult);
      } catch (error) {
        // エラーの場合はfalseとして扱う
        condition = false;
      }
    }
    
    // 結果の値を取得（改善された処理）
    const resultStr = condition ? trueValueStr : falseValueStr;
    
    try {
      // より包括的な結果評価
      return evaluateFormulaResult(resultStr, context);
    } catch (error) {
      // エラーが発生した場合、元の文字列を返すかFormulaErrorを返す
      if (error === FormulaError.VALUE || error === FormulaError.REF || error === FormulaError.NUM || 
          error === FormulaError.DIV0 || error === FormulaError.NA || error === FormulaError.CALC) {
        return error;
      }
      return resultStr as FormulaResult;
    }
  }
};

// AND関数の実装
export const AND: CustomFormula = {
  name: 'AND',
  pattern: /AND\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, args] = matches;
    const logicalValues = parseArgumentsToLogical(args, context);
    
    if (logicalValues.length === 0) {
      return FormulaError.VALUE;
    }
    
    // すべてがtrueの場合のみtrue
    return logicalValues.every(value => value === true);
  }
};

// OR関数の実装
export const OR: CustomFormula = {
  name: 'OR',
  pattern: /\bOR\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, args] = matches;
    const logicalValues = parseArgumentsToLogical(args, context);
    
    if (logicalValues.length === 0) {
      return FormulaError.VALUE;
    }
    
    // いずれかがtrueの場合にtrue
    return logicalValues.some(value => value === true);
  }
};

// NOT関数の実装
export const NOT: CustomFormula = {
  name: 'NOT',
  pattern: /NOT\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, valueRef] = matches;
    
    const value = getCellValue(valueRef, context) ?? valueRef;
    const logicalValue = toLogical(value);
    
    return !logicalValue;
  }
};

// IFS関数の実装（手動実装が必要）
export const IFS: CustomFormula = {
  name: 'IFS',
  pattern: /IFS\(((?:[^()]|\([^)]*\))*)\)/i,
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
    
    
    // 条件と値のペアを評価
    for (let i = 0; i < args.length; i += 2) {
      if (i + 1 >= args.length) break;
      
      const conditionArg = args[i];
      const valueArg = args[i + 1];
      
      // 条件を評価（簡単な実装）
      let conditionResult = false;
      
      // TRUEやFALSEの場合
      if (conditionArg.toUpperCase() === 'TRUE') {
        conditionResult = true;
      } else if (conditionArg.toUpperCase() === 'FALSE') {
        conditionResult = false;
      }
      // "1"や"0"の場合
      else if (conditionArg === '"1"' || conditionArg === '1') {
        conditionResult = true;
      } else if (conditionArg === '"0"' || conditionArg === '0') {
        conditionResult = false;
      }
      // セル参照の場合
      else if (conditionArg.match(/^[A-Z]+\d+$/)) {
        const cellValue = getCellValue(conditionArg, context);
        conditionResult = Boolean(cellValue);
      } else {
        // 簡単な比較演算子の処理
        const comparisonMatch = conditionArg.match(/([A-Z]+\d+|"[^"]*"|\d+(?:\.\d+)?)\s*([><=!]+)\s*([A-Z]+\d+|"[^"]*"|\d+(?:\.\d+)?)/);
        if (comparisonMatch) {
          const [, leftParam, operator, rightParam] = comparisonMatch;
          let left = leftParam;
          let right = rightParam;
          
          // セル参照の場合は値を取得
          if (left.match(/^[A-Z]+\d+$/)) {
            const cellVal = getCellValue(left, context);
            left = cellVal !== null && cellVal !== undefined ? String(cellVal) : '0';
          } else if (left.startsWith('"') && left.endsWith('"')) {
            left = left.slice(1, -1);
          }
          
          if (right.match(/^[A-Z]+\d+$/)) {
            const cellVal = getCellValue(right, context);
            right = cellVal !== null && cellVal !== undefined ? String(cellVal) : '0';
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
          return getCellValue(valueArg, context) as FormulaResult;
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

// XOR関数の実装（排他的論理和）
export const XOR: CustomFormula = {
  name: 'XOR',
  pattern: /XOR\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const args = matches[1].split(',').map(arg => arg.trim());
    let trueCount = 0;
    
    for (const arg of args) {
      let value = false;
      
      // セル参照の場合
      if (arg.match(/^[A-Z]+\d+$/)) {
        const cellValue = getCellValue(arg, context);
        value = toLogical(cellValue);
      } else {
        // 値の評価
        const trimmed = arg.replace(/"/g, '');
        if (trimmed === 'TRUE' || trimmed === '1') {
          value = true;
        } else if (trimmed === 'FALSE' || trimmed === '0') {
          value = false;
        } else {
          const numValue = parseFloat(trimmed);
          value = !isNaN(numValue) ? numValue !== 0 : trimmed.length > 0;
        }
      }
      
      if (value) {
        trueCount++;
      }
    }
    
    // XORは奇数個の引数がtrueの場合にtrueを返す
    return trueCount % 2 === 1;
  }
};

// TRUE関数の実装
export const TRUE: CustomFormula = {
  name: 'TRUE',
  pattern: /TRUE\(\)/i,
  calculate: (): FormulaResult => {
    return true;
  }
};

// FALSE関数の実装
export const FALSE: CustomFormula = {
  name: 'FALSE',
  pattern: /FALSE\(\)/i,
  calculate: (): FormulaResult => {
    return false;
  }
};

// IFERROR関数の実装（エラー時の値指定）
export const IFERROR: CustomFormula = {
  name: 'IFERROR',
  pattern: /IFERROR\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const fullArgs = matches[1];
    
    // 引数を分割（ネストされた関数の括弧を考慮）
    const args: string[] = [];
    let currentArg = '';
    let parenDepth = 0;
    let inQuotes = false;
    
    for (let i = 0; i < fullArgs.length; i++) {
      const char = fullArgs[i];
      
      if (char === '"' && fullArgs[i - 1] !== '\\') {
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
    
    if (args.length !== 2) {
      return FormulaError.VALUE;
    }
    
    const [valueExpr, errorValueExpr] = args;
    
    // エラー時の値を取得（引用符を除去）
    let errorValue: FormulaResult;
    if (errorValueExpr.startsWith('"') && errorValueExpr.endsWith('"')) {
      errorValue = errorValueExpr.slice(1, -1);
    } else if (errorValueExpr.match(/^[A-Z]+\d+$/)) {
      errorValue = getCellValue(errorValueExpr, context) as FormulaResult;
    } else {
      const num = parseFloat(errorValueExpr);
      errorValue = !isNaN(num) ? num : errorValueExpr;
    }
    
    // 値を評価
    let result: FormulaResult;
    
    // セル参照だけの場合の処理
    if (valueExpr.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(valueExpr, context);
      if (typeof cellValue === 'string' && cellValue.startsWith('#')) {
        return errorValue;
      }
      return cellValue as FormulaResult;
    }
    
    // 関数を含む式の場合
    const nestedMatch = matchFormula(valueExpr);
    if (nestedMatch) {
      result = nestedMatch.function.calculate(nestedMatch.matches, context);
      if (typeof result === 'string' && result.startsWith('#')) {
        return errorValue;
      }
      return result;
    }
    
    // A2/B2のような式を評価する必要がある
    result = evaluateFormulaWithErrorCheck(valueExpr, context);
    
    // エラーチェック
    if (typeof result === 'string' && result.startsWith('#')) {
      return errorValue;
    }
    
    return result;
  }
};

// IFNA関数の実装（#N/Aエラー時の値）
export const IFNA: CustomFormula = {
  name: 'IFNA',
  pattern: /IFNA\s*\((.+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const fullArgs = matches[1];
    
    // 引数を分割（ネストされた関数の括弧を考慮）
    const args: string[] = [];
    let currentArg = '';
    let parenDepth = 0;
    let inQuotes = false;
    let quoteChar = '';
    
    for (let i = 0; i < fullArgs.length; i++) {
      const char = fullArgs[i];
      const prevChar = i > 0 ? fullArgs[i - 1] : '';
      
      // エスケープされた引用符を処理
      if ((char === '"' || char === "'") && prevChar !== '\\') {
        if (!inQuotes) {
          inQuotes = true;
          quoteChar = char;
        } else if (char === quoteChar) {
          inQuotes = false;
        }
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
    
    if (args.length !== 2) {
      return FormulaError.VALUE;
    }
    
    const [valueExpr, naValueExpr] = args;
    
    // #N/Aエラー時の値を取得（引用符を除去）
    let naValue: FormulaResult;
    if (naValueExpr.startsWith('"') && naValueExpr.endsWith('"')) {
      naValue = naValueExpr.slice(1, -1);
    } else if (naValueExpr.match(/^[A-Z]+\d+$/)) {
      naValue = getCellValue(naValueExpr, context) as FormulaResult;
    } else {
      const num = parseFloat(naValueExpr);
      naValue = !isNaN(num) ? num : naValueExpr;
    }
    
    // 値を評価
    let result: FormulaResult;
    
    // セル参照だけの場合の処理
    if (valueExpr.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(valueExpr, context);
      if (cellValue === '#N/A' || cellValue === FormulaError.NA) {
        return naValue;
      }
      return cellValue as FormulaResult;
    }
    
    // 関数を含む式の場合（VLOOKUP、MATCH、INDEX等）
    const nestedMatch = matchFormula(valueExpr);
    if (nestedMatch) {
      result = nestedMatch.function.calculate(nestedMatch.matches, context);
      if (result === '#N/A' || result === FormulaError.NA) {
        return naValue;
      }
      return result;
    }
    
    // その他の式の評価
    result = evaluateFormulaWithErrorCheck(valueExpr, context);
    if (result === '#N/A' || result === FormulaError.NA) {
      return naValue;
    }
    
    return result;
  }
};