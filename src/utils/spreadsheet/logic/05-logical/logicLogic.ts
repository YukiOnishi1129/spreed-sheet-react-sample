// 論理関数の実装

import type { CustomFormula, FormulaContext, FormulaResult } from '../shared/types';
import { FormulaError } from '../shared/types';
import { getCellValue, evaluateExpression } from '../shared/utils';

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
      values.push(false);
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
    return false;
  }
  return Boolean(value);
}

// IF関数の実装
export const IF: CustomFormula = {
  name: 'IF',
  pattern: /^IF\s*\((.+)\)$/i,
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
      const leftValue = getCellValue(leftExpr.trim(), context) ?? leftExpr.trim();
      const rightValue = getCellValue(rightExpr.trim(), context) ?? rightExpr.trim();
      
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
      // 比較演算子がない場合は、値を論理値として評価
      const conditionValue = getCellValue(conditionStr, context) ?? conditionStr;
      condition = toLogical(conditionValue);
    }
    
    // 結果の値を取得
    const resultStr = condition ? trueValueStr : falseValueStr;
    
    // 引用符付き文字列の場合は直接処理
    if (resultStr.startsWith('"') && resultStr.endsWith('"')) {
      return resultStr.slice(1, -1);
    }
    
    // エスケープされた引用符の場合
    if (resultStr.startsWith('\\"') && resultStr.endsWith('\\"')) {
      return resultStr.slice(2, -2);
    }
    
    // セル参照の場合
    if (resultStr.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(resultStr, context);
      if (cellValue === FormulaError.REF) {
        return FormulaError.REF;
      }
      return cellValue as FormulaResult;
    }
    
    // その他の場合（数値または式）
    const num = parseFloat(resultStr);
    if (!isNaN(num)) {
      return num;
    }
    
    // 算術式の可能性をチェック
    if (resultStr.includes('*') || resultStr.includes('/') || resultStr.includes('+') || resultStr.includes('-')) {
      const evaluated = evaluateExpression(resultStr, context);
      return evaluated;
    }
    
    return resultStr as FormulaResult;
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
  pattern: /OR\(([^)]+)\)/i,
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
    
    console.log('IFS args:', args);
    
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
            console.log(`IFS: Cell ${left} value:`, cellVal);
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
    
    console.log('IFS: No condition matched, returning #N/A');
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
        value = Boolean(cellValue);
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
  pattern: /IFERROR\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const valueArg = matches[1].trim();
    const errorValue = matches[2].replace(/"/g, '').trim();
    
    // セル参照の場合
    if (valueArg.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(valueArg, context);
      // エラーチェック（文字列でエラーを表現）
      if (typeof cellValue === 'string' && cellValue.startsWith('#') && cellValue.endsWith('!')) {
        return errorValue;
      }
      return cellValue as FormulaResult;
    }
    
    // 値がエラーかどうかをチェック
    if (valueArg.startsWith('#') && valueArg.endsWith('!')) {
      return errorValue;
    }
    
    // 数値変換を試す
    const numValue = parseFloat(valueArg);
    if (!isNaN(numValue)) {
      return numValue;
    }
    
    return valueArg.replace(/"/g, '');
  }
};

// IFNA関数の実装（#N/Aエラー時の値）
export const IFNA: CustomFormula = {
  name: 'IFNA',
  pattern: /IFNA\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const valueArg = matches[1].trim();
    const naValue = matches[2].replace(/"/g, '').trim();
    
    // セル参照の場合
    if (valueArg.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(valueArg, context);
      // #N/Aエラーチェック
      if (cellValue === '#N/A' || cellValue === '#N/A!') {
        return naValue;
      }
      return cellValue as FormulaResult;
    }
    
    // #N/Aエラーの場合のみ置換
    if (valueArg === '#N/A' || valueArg === '#N/A!') {
      return naValue;
    }
    
    // 数値変換を試す
    const numValue = parseFloat(valueArg);
    if (!isNaN(numValue)) {
      return numValue;
    }
    
    return valueArg.replace(/"/g, '');
  }
};