import type { CustomFormula, CellData, FormulaContext } from '../shared/types';
import { getCellValue } from '../shared/utils';

// 基本的な算術演算子
export const MULTIPLY_OPERATOR: CustomFormula = {
  name: '*',
  pattern: /^(.+)\*(.+)$/,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, left, right] = matches;
    
    // 左辺の値を取得
    const leftValue = evaluateExpression(left.trim(), context);
    const leftNum = Number(leftValue);
    
    // 右辺の値を取得
    const rightValue = evaluateExpression(right.trim(), context);
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
    const leftNum = Number(leftValue);
    
    const rightValue = evaluateExpression(right.trim(), context);
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
    const leftNum = Number(leftValue);
    
    const rightValue = evaluateExpression(right.trim(), context);
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
    const leftNum = Number(leftValue);
    
    const rightValue = evaluateExpression(right.trim(), context);
    const rightNum = Number(rightValue);
    
    if (isNaN(leftNum) || isNaN(rightNum)) {
      return '#VALUE!';
    }
    
    return leftNum - rightNum;
  }
};

// 式を評価する補助関数
function evaluateExpression(expression: string, context: FormulaContext): string | number {
  // セル参照の場合
  if (/^[A-Z]+\d+$/.test(expression)) {
    return getCellValue(expression, context);
  }
  
  // 数値の場合
  const num = Number(expression);
  if (!isNaN(num)) {
    return num;
  }
  
  // その他の場合は文字列として返す
  return expression;
}