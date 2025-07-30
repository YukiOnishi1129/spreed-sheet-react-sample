import type { CustomFormula, FormulaContext } from '../shared/types';
import { getCellValue } from '../shared/utils';

// 基本的な算術演算子
export const MULTIPLY_OPERATOR: CustomFormula = {
  name: '*',
  pattern: /^(.+)\*(.+)$/,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, left, right] = matches;
    
    // 左辺の値を取得
    const leftValue = evaluateExpression(left.trim(), context);
    const rightValue = evaluateExpression(right.trim(), context);
    
    // 算術演算に有効な値かチェック
    if (!isValidForArithmetic(leftValue) || !isValidForArithmetic(rightValue)) {
      return '#VALUE!';
    }
    
    const leftNum = Number(leftValue);
    const rightNum = Number(rightValue);
    
    if (isNaN(leftNum) || isNaN(rightNum)) {
      return '#VALUE!';
    }
    
    return leftNum * rightNum;
  }
};

export const DIVIDE_OPERATOR: CustomFormula = {
  name: '/',
  pattern: /^(.+)\/(.+)$/,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, left, right] = matches;
    
    const leftValue = evaluateExpression(left.trim(), context);
    const rightValue = evaluateExpression(right.trim(), context);
    
    // 算術演算に有効な値かチェック
    if (!isValidForArithmetic(leftValue) || !isValidForArithmetic(rightValue)) {
      return '#VALUE!';
    }
    
    const leftNum = Number(leftValue);
    const rightNum = Number(rightValue);
    
    if (isNaN(leftNum) || isNaN(rightNum)) {
      return '#VALUE!';
    }
    
    if (rightNum === 0) {
      return '#DIV/0!';
    }
    
    return leftNum / rightNum;
  }
};

export const ADD_OPERATOR: CustomFormula = {
  name: '+',
  pattern: /^(.+)\+(.+)$/,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, left, right] = matches;
    
    const leftValue = evaluateExpression(left.trim(), context);
    const rightValue = evaluateExpression(right.trim(), context);
    
    // 算術演算に有効な値かチェック
    if (!isValidForArithmetic(leftValue) || !isValidForArithmetic(rightValue)) {
      return '#VALUE!';
    }
    
    const leftNum = Number(leftValue);
    const rightNum = Number(rightValue);
    
    if (isNaN(leftNum) || isNaN(rightNum)) {
      return '#VALUE!';
    }
    
    return leftNum + rightNum;
  }
};

export const SUBTRACT_OPERATOR: CustomFormula = {
  name: '-',
  pattern: /^(.+)-(.+)$/,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, left, right] = matches;
    
    const leftValue = evaluateExpression(left.trim(), context);
    const rightValue = evaluateExpression(right.trim(), context);
    
    // 算術演算に有効な値かチェック
    if (!isValidForArithmetic(leftValue) || !isValidForArithmetic(rightValue)) {
      return '#VALUE!';
    }
    
    const leftNum = Number(leftValue);
    const rightNum = Number(rightValue);
    
    if (isNaN(leftNum) || isNaN(rightNum)) {
      return '#VALUE!';
    }
    
    return leftNum - rightNum;
  }
};

// 比較演算子
export const GREATER_THAN_OR_EQUAL: CustomFormula = {
  name: '>=',
  pattern: /^(.+)>=(.+)$/,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, left, right] = matches;
    
    const leftValue = evaluateExpression(left.trim(), context);
    const rightValue = evaluateExpression(right.trim(), context);
    
    // 数値比較の場合は数値として比較
    if (isValidForArithmetic(leftValue) && isValidForArithmetic(rightValue)) {
      const leftNum = Number(leftValue);
      const rightNum = Number(rightValue);
      return leftNum >= rightNum;
    }
    
    // どちらかが数値でない場合はfalseを返す
    return false;
  }
};

export const LESS_THAN_OR_EQUAL: CustomFormula = {
  name: '<=',
  pattern: /^(.+)<=(.+)$/,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, left, right] = matches;
    
    const leftValue = evaluateExpression(left.trim(), context);
    const rightValue = evaluateExpression(right.trim(), context);
    
    // 数値比較の場合は数値として比較
    if (isValidForArithmetic(leftValue) && isValidForArithmetic(rightValue)) {
      const leftNum = Number(leftValue);
      const rightNum = Number(rightValue);
      return leftNum <= rightNum;
    }
    
    // どちらかが数値でない場合はfalseを返す
    return false;
  }
};

export const GREATER_THAN: CustomFormula = {
  name: '>',
  pattern: /^(.+)>(.+)$/,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, left, right] = matches;
    
    const leftValue = evaluateExpression(left.trim(), context);
    const rightValue = evaluateExpression(right.trim(), context);
    
    // 数値比較の場合は数値として比較
    if (isValidForArithmetic(leftValue) && isValidForArithmetic(rightValue)) {
      const leftNum = Number(leftValue);
      const rightNum = Number(rightValue);
      return leftNum > rightNum;
    }
    
    // どちらかが数値でない場合はfalseを返す
    return false;
  }
};

export const LESS_THAN: CustomFormula = {
  name: '<',
  pattern: /^(.+)<(.+)$/,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, left, right] = matches;
    
    const leftValue = evaluateExpression(left.trim(), context);
    const rightValue = evaluateExpression(right.trim(), context);
    
    // 数値比較の場合は数値として比較
    if (isValidForArithmetic(leftValue) && isValidForArithmetic(rightValue)) {
      const leftNum = Number(leftValue);
      const rightNum = Number(rightValue);
      return leftNum < rightNum;
    }
    
    // どちらかが数値でない場合はfalseを返す
    return false;
  }
};

export const EQUAL: CustomFormula = {
  name: '=',
  pattern: /^(.+)=(.+)$/,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, left, right] = matches;
    
    const leftValue = evaluateExpression(left.trim(), context);
    const rightValue = evaluateExpression(right.trim(), context);
    
    return leftValue === rightValue;
  }
};

export const NOT_EQUAL: CustomFormula = {
  name: '<>',
  pattern: /^(.+)<>(.+)$/,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, left, right] = matches;
    
    const leftValue = evaluateExpression(left.trim(), context);
    const rightValue = evaluateExpression(right.trim(), context);
    
    return leftValue !== rightValue;
  }
};

// 式を評価する補助関数
function evaluateExpression(expression: string, context: FormulaContext): string | number {
  // セル参照の場合
  if (/^[A-Z]+\d+$/.test(expression)) {
    const value = getCellValue(expression, context);
    // FormulaErrorの場合はそのまま返す
    if (typeof value === 'object' && value !== null) {
      return String(value);
    }
    return value as string | number;
  }
  
  // 数値の場合
  const num = Number(expression);
  if (!isNaN(num)) {
    return num;
  }
  
  // その他の場合は文字列として返す
  return expression;
}

// 値が算術演算に有効かチェックする補助関数
function isValidForArithmetic(value: unknown): boolean {
  // null、undefinedは無効（Excelでは#VALUE!エラー）
  if (value === null || value === undefined) {
    return false;
  }
  // 空文字列は有効（Excelでは0として扱われる）
  if (value === '') {
    return true;
  }
  // 数値または数値に変換できる文字列は有効
  if (typeof value === 'number') {
    return !isNaN(value);
  }
  if (typeof value === 'string') {
    const num = Number(value);
    return !isNaN(num);
  }
  return false;
}