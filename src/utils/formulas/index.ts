// Excel関数のモジュール化されたインデックス

import type { CustomFormula } from './types';

// 各カテゴリから関数をインポート
import { DATEDIF, NETWORKDAYS, TODAY } from './dateFunctions';
import { SUMIF, COUNTIF, AVERAGEIF, SUM, AVERAGE, COUNT, MAX, MIN, ROUND } from './mathFunctions';
import { VLOOKUP, HLOOKUP, INDEX, MATCH, LOOKUP, XLOOKUP } from './lookupFunctions';
import { IF, AND, OR, NOT, IFS } from './logicFunctions';
import { CONCATENATE, CONCAT, LEFT, RIGHT, MID, LEN, UPPER, LOWER, TRIM, SUBSTITUTE, FIND, SEARCH, TEXTJOIN, SPLIT } from './textFunctions';

// すべての関数を配列にまとめる
export const ALL_FUNCTIONS: CustomFormula[] = [
  // 日付関数
  DATEDIF,
  NETWORKDAYS,
  TODAY,
  
  // 数学関数
  SUMIF,
  COUNTIF,
  AVERAGEIF,
  SUM,
  AVERAGE,
  COUNT,
  MAX,
  MIN,
  ROUND,
  
  // 検索関数
  VLOOKUP,
  HLOOKUP,
  INDEX,
  MATCH,
  LOOKUP,
  XLOOKUP,
  
  // 論理関数
  IF,
  AND,
  OR,
  NOT,
  IFS,
  
  // テキスト関数
  CONCATENATE,
  CONCAT,
  LEFT,
  RIGHT,
  MID,
  LEN,
  UPPER,
  LOWER,
  TRIM,
  SUBSTITUTE,
  FIND,
  SEARCH,
  TEXTJOIN,
  SPLIT
];

// カテゴリ別の関数分類
export const FUNCTION_CATEGORIES = {
  date: [DATEDIF, NETWORKDAYS, TODAY],
  math: [SUMIF, COUNTIF, AVERAGEIF, SUM, AVERAGE, COUNT, MAX, MIN, ROUND],
  lookup: [VLOOKUP, HLOOKUP, INDEX, MATCH, LOOKUP, XLOOKUP],
  logic: [IF, AND, OR, NOT, IFS],
  text: [CONCATENATE, CONCAT, LEFT, RIGHT, MID, LEN, UPPER, LOWER, TRIM, SUBSTITUTE, FIND, SEARCH, TEXTJOIN, SPLIT]
};

// HyperFormulaでサポートされていない関数（手動計算が必要）
export const UNSUPPORTED_FUNCTIONS = ALL_FUNCTIONS.filter(f => f.isSupported === false);

// HyperFormulaでサポートされている関数
export const SUPPORTED_FUNCTIONS = ALL_FUNCTIONS.filter(f => f.isSupported !== false);

// 関数名で検索
export const findFunction = (name: string): CustomFormula | undefined => {
  return ALL_FUNCTIONS.find(f => f.name.toUpperCase() === name.toUpperCase());
};

// 関数のパターンマッチング
export const matchFormula = (formula: string): { function: CustomFormula; matches: RegExpMatchArray } | null => {
  for (const func of ALL_FUNCTIONS) {
    const matches = formula.match(func.pattern);
    if (matches) {
      return { function: func, matches };
    }
  }
  return null;
};

// 関数タイプの判定（色分けのため）
export const getFunctionType = (functionName: string): string => {
  const name = functionName.toUpperCase();
  
  if (FUNCTION_CATEGORIES.date.some(f => f.name === name)) return 'date';
  if (FUNCTION_CATEGORIES.math.some(f => f.name === name)) return 'math';
  if (FUNCTION_CATEGORIES.lookup.some(f => f.name === name)) return 'lookup';
  if (FUNCTION_CATEGORIES.logic.some(f => f.name === name)) return 'logic';
  if (FUNCTION_CATEGORIES.text.some(f => f.name === name)) return 'text';
  
  return 'other';
};

// エクスポート
export * from './types';
export * from './utils';
export * from './dateFunctions';
export * from './mathFunctions';
export * from './lookupFunctions';
export * from './logicFunctions';
export * from './textFunctions';