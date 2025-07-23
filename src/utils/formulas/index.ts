// Excel関数のモジュール化されたインデックス

import type { CustomFormula } from './types';

// 各カテゴリから関数をインポート
import { DATEDIF, NETWORKDAYS, TODAY, NOW, DATE, YEAR, MONTH, DAY, WEEKDAY, DAYS, EDATE, EOMONTH, TIME, HOUR, MINUTE, SECOND, WEEKNUM, DAYS360, YEARFRAC } from './dateFunctions';
import { 
  SUMIF, COUNTIF, AVERAGEIF, SUM, AVERAGE, COUNT, MAX, MIN, ROUND,
  ABS, SQRT, POWER, MOD, INT, TRUNC, RAND, RANDBETWEEN, PI, DEGREES, RADIANS,
  SIN, COS, TAN, LOG, LOG10, LN, EXP, ASIN, ACOS, ATAN, ATAN2, ROUNDUP, ROUNDDOWN,
  CEILING, FLOOR, SIGN, FACT, SUMIFS, COUNTIFS, AVERAGEIFS, PRODUCT, MROUND,
  COMBIN, PERMUT, GCD, LCM, QUOTIENT, SINH, COSH, TANH,
  SUMSQ, SUMPRODUCT, EVEN, ODD, ARABIC, ROMAN, COMBINA, FACTDOUBLE, SQRTPI
} from './mathFunctions';
import { VLOOKUP, HLOOKUP, INDEX, MATCH, LOOKUP, XLOOKUP } from './lookupFunctions';
import { IF, AND, OR, NOT, IFS, XOR, TRUE, FALSE, IFERROR, IFNA } from './logicFunctions';
import { 
  CONCATENATE, CONCAT, LEFT, RIGHT, MID, LEN, UPPER, LOWER, TRIM, SUBSTITUTE, FIND, SEARCH, TEXTJOIN, SPLIT,
  PROPER, VALUE, TEXT, REPT, REPLACE, CHAR, CODE, EXACT, CLEAN, T, FIXED
} from './textFunctions';
import { 
  MEDIAN, MODE, COUNTA, COUNTBLANK, STDEV, VAR, LARGE, SMALL, RANK,
  CORREL, QUARTILE, PERCENTILE, GEOMEAN, HARMEAN, TRIMMEAN
} from './statisticsFunctions';
import {
  ISBLANK, ISERROR, ISNA, ISTEXT, ISNUMBER, ISLOGICAL, ISEVEN, ISODD, TYPE, N
} from './informationFunctions';
import { PMT, PV, FV, NPV, IRR, PPMT, IPMT } from './financialFunctions';

// すべての関数を配列にまとめる
export const ALL_FUNCTIONS = [
  // 日付関数
  DATEDIF,
  NETWORKDAYS,
  TODAY,
  NOW,
  DATE,
  YEAR,
  MONTH,
  DAY,
  WEEKDAY,
  DAYS,
  EDATE,
  EOMONTH,
  TIME,
  HOUR,
  MINUTE,
  SECOND,
  WEEKNUM,
  DAYS360,
  YEARFRAC,
  
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
  ABS,
  SQRT,
  POWER,
  MOD,
  INT,
  TRUNC,
  RAND,
  RANDBETWEEN,
  PI,
  DEGREES,
  RADIANS,
  SIN,
  COS,
  TAN,
  LOG,
  LOG10,
  LN,
  EXP,
  ASIN,
  ACOS,
  ATAN,
  ATAN2,
  ROUNDUP,
  ROUNDDOWN,
  CEILING,
  FLOOR,
  SIGN,
  FACT,
  SUMIFS,
  COUNTIFS,
  AVERAGEIFS,
  PRODUCT,
  MROUND,
  COMBIN,
  PERMUT,
  GCD,
  LCM,
  QUOTIENT,
  SINH,
  COSH,
  TANH,
  SUMSQ,
  SUMPRODUCT,
  EVEN,
  ODD,
  ARABIC,
  ROMAN,
  COMBINA,
  FACTDOUBLE,
  SQRTPI,
  
  // 統計関数
  MEDIAN,
  MODE,
  COUNTA,
  COUNTBLANK,
  STDEV,
  VAR,
  LARGE,
  SMALL,
  RANK,
  CORREL,
  QUARTILE,
  PERCENTILE,
  GEOMEAN,
  HARMEAN,
  TRIMMEAN,
  
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
  XOR,
  TRUE,
  FALSE,
  IFERROR,
  IFNA,
  
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
  SPLIT,
  PROPER,
  VALUE,
  TEXT,
  REPT,
  REPLACE,
  CHAR,
  CODE,
  EXACT,
  CLEAN,
  T,
  FIXED,
  
  // 情報関数
  ISBLANK,
  ISERROR,
  ISNA,
  ISTEXT,
  ISNUMBER,
  ISLOGICAL,
  ISEVEN,
  ISODD,
  TYPE,
  N,
  
  // 財務関数
  PMT,
  PV,
  FV,
  NPV,
  IRR,
  PPMT,
  IPMT
] as CustomFormula[];

// カテゴリ別の関数分類
export const FUNCTION_CATEGORIES = {
  date: [DATEDIF, NETWORKDAYS, TODAY, NOW, DATE, YEAR, MONTH, DAY, WEEKDAY, DAYS, EDATE, EOMONTH, TIME, HOUR, MINUTE, SECOND, WEEKNUM, DAYS360, YEARFRAC] as CustomFormula[],
  math: [SUMIF, COUNTIF, AVERAGEIF, SUM, AVERAGE, COUNT, MAX, MIN, ROUND, ABS, SQRT, POWER, MOD, INT, TRUNC, RAND, RANDBETWEEN, PI, DEGREES, RADIANS, SIN, COS, TAN, LOG, LOG10, LN, EXP, ASIN, ACOS, ATAN, ATAN2, ROUNDUP, ROUNDDOWN, CEILING, FLOOR, SIGN, FACT, SUMIFS, COUNTIFS, AVERAGEIFS, PRODUCT, MROUND, COMBIN, PERMUT, GCD, LCM, QUOTIENT, SINH, COSH, TANH, SUMSQ, SUMPRODUCT, EVEN, ODD, ARABIC, ROMAN, COMBINA, FACTDOUBLE, SQRTPI] as CustomFormula[],
  statistics: [MEDIAN, MODE, COUNTA, COUNTBLANK, STDEV, VAR, LARGE, SMALL, RANK, CORREL, QUARTILE, PERCENTILE, GEOMEAN, HARMEAN, TRIMMEAN] as CustomFormula[],
  lookup: [VLOOKUP, HLOOKUP, INDEX, MATCH, LOOKUP, XLOOKUP] as CustomFormula[],
  logic: [IF, AND, OR, NOT, IFS, XOR, TRUE, FALSE, IFERROR, IFNA] as CustomFormula[],
  text: [CONCATENATE, CONCAT, LEFT, RIGHT, MID, LEN, UPPER, LOWER, TRIM, SUBSTITUTE, FIND, SEARCH, TEXTJOIN, SPLIT, PROPER, VALUE, TEXT, REPT, REPLACE, CHAR, CODE, EXACT, CLEAN, T, FIXED] as CustomFormula[],
  information: [ISBLANK, ISERROR, ISNA, ISTEXT, ISNUMBER, ISLOGICAL, ISEVEN, ISODD, TYPE, N] as CustomFormula[],
  financial: [PMT, PV, FV, NPV, IRR, PPMT, IPMT] as CustomFormula[]
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
  if (FUNCTION_CATEGORIES.statistics.some(f => f.name === name)) return 'statistics';
  if (FUNCTION_CATEGORIES.lookup.some(f => f.name === name)) return 'lookup';
  if (FUNCTION_CATEGORIES.logic.some(f => f.name === name)) return 'logic';
  if (FUNCTION_CATEGORIES.text.some(f => f.name === name)) return 'text';
  if (FUNCTION_CATEGORIES.information.some(f => f.name === name)) return 'information';
  if (FUNCTION_CATEGORIES.financial.some(f => f.name === name)) return 'financial';
  
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
export * from './statisticsFunctions';
export * from './informationFunctions';
export * from './financialFunctions';