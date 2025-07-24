// Excel関数のモジュール化されたインデックス

import type { CustomFormula } from './shared/types';

// 各カテゴリから関数をインポート
import { DATEDIF, NETWORKDAYS, TODAY, NOW, DATE, YEAR, MONTH, DAY, WEEKDAY, DAYS, EDATE, EOMONTH, TIME, HOUR, MINUTE, SECOND, WEEKNUM, DAYS360, YEARFRAC, DATEVALUE, TIMEVALUE, ISOWEEKNUM, NETWORKDAYS_INTL, WORKDAY, WORKDAY_INTL } from './04-datetime/dateFunctions';

import { 
  SUMIF, COUNTIF, AVERAGEIF, SUM, AVERAGE, COUNT, MAX, MIN, ROUND,
  ABS, SQRT, POWER, MOD, INT, TRUNC, RAND, RANDBETWEEN,
  LOG, LOG10, LN, EXP, ROUNDUP, ROUNDDOWN,
  CEILING, FLOOR, SIGN, FACT, SUMIFS, COUNTIFS, AVERAGEIFS, PRODUCT, MROUND,
  COMBIN, PERMUT, GCD, LCM, QUOTIENT,
  SUMSQ, SUMPRODUCT, EVEN, ODD, ARABIC, ROMAN, COMBINA, FACTDOUBLE, SQRTPI,
  SUMX2MY2, SUMX2PY2, SUMXMY2, MULTINOMIAL, PERMUTATIONA, BASE, DECIMAL, SUBTOTAL,
  AGGREGATE, CEILING_MATH, CEILING_PRECISE, FLOOR_MATH, FLOOR_PRECISE, ISO_CEILING, SERIESSUM,
  RANDARRAY, SEQUENCE
} from './01-math-trigonometry/mathFunctions';

import {
  SIN, COS, TAN, ASIN, ACOS, ATAN, ATAN2, SINH, COSH, TANH,
  ASINH, ACOSH, ATANH, CSC, SEC, COT, ACOT, CSCH, SECH, COTH, ACOTH
} from './01-math-trigonometry/trigonometricFunctions';

import { 
  MEDIAN, MODE, COUNTA, COUNTBLANK, LARGE, SMALL, RANK,
  GEOMEAN, HARMEAN, TRIMMEAN
} from './02-statistical/basicStatisticsFunctions';

import {
  DEVSQ, KURT, SKEW, PEARSON, PERCENTRANK,
  STANDARDIZE, COVARIANCE_P, COVARIANCE_S, SKEW_P,
  PERCENTILE_INC, PERCENTILE_EXC, QUARTILE_INC, QUARTILE_EXC
} from './02-statistical/advancedStatisticsFunctions';

import {
  NORM_DIST, NORM_INV, NORM_S_DIST, NORM_S_INV, LOGNORM_DIST, LOGNORM_INV,
  T_DIST, T_DIST_2T, T_DIST_RT, CHISQ_DIST, CHISQ_DIST_RT,
  F_DIST, F_DIST_RT, BETA_DIST,
  GAMMA_DIST, EXPON_DIST, WEIBULL_DIST, BINOM_DIST,
  NEGBINOM_DIST, POISSON_DIST, HYPGEOM_DIST, CONFIDENCE_NORM,
  CONFIDENCE_T, Z_TEST, T_TEST, F_TEST, CHISQ_TEST
} from './02-statistical/distributionFunctions';

import {
  T_INV, T_INV_2T, CHISQ_INV, CHISQ_INV_RT, F_INV, F_INV_RT, BETA_INV, GAMMA_INV
} from './02-statistical/inverseDistributionFunctions';

import { 
  CONCATENATE, CONCAT, LEFT, RIGHT, MID, LEN, UPPER, LOWER, TRIM, SUBSTITUTE, FIND, SEARCH, TEXTJOIN, SPLIT,
  PROPER, VALUE, TEXT, REPT, REPLACE, CHAR, CODE, EXACT, CLEAN, T, FIXED, NUMBERVALUE, DOLLAR, UNICHAR, UNICODE,
  LENB, FINDB, SEARCHB, REPLACEB, TEXTBEFORE, TEXTAFTER, TEXTSPLIT, ASC, JIS, DBCS, PHONETIC, BAHTTEXT
} from './03-text/textFunctions';

import { IF, AND, OR, NOT, IFS, XOR, TRUE, FALSE, IFERROR, IFNA } from './05-logical/logicFunctions';
import { SWITCH, LET } from './05-logical/newLogicalFunctions';

import { VLOOKUP, HLOOKUP, INDEX, MATCH, LOOKUP, XLOOKUP, OFFSET, INDIRECT, CHOOSE, TRANSPOSE, FILTER, SORT, UNIQUE } from './06-lookup-reference/lookupFunctions';

import { JOIN, ARRAYFORMULA, QUERY, REGEXMATCH, REGEXEXTRACT, REGEXREPLACE, FLATTEN } from './16-google-sheets/googleSheetsFunctions';

import { 
  ISBLANK, ISERROR, ISNA, ISTEXT, ISNUMBER, ISLOGICAL, ISEVEN, ISODD, TYPE, N,
  ISERR, ISNONTEXT, ISREF, ISFORMULA, NA, ERROR_TYPE, INFO, SHEET, SHEETS, CELL
} from './07-information/informationFunctions';

import { DSUM, DAVERAGE, DCOUNT, DCOUNTA, DMAX, DMIN, DPRODUCT, DGET } from './08-database/databaseFunctions';

import { PMT, PV, FV, NPV, IRR, PPMT, IPMT, RATE } from './09-financial/financialFunctions';

import { CONVERT, BIN2DEC, DEC2BIN, HEX2DEC, DEC2HEX, BIN2HEX, HEX2BIN, OCT2DEC, DEC2OCT } from './13-engineering/engineeringFunctions';

import { MDETERM, MINVERSE, MMULT, MUNIT } from './15-others/matrixFunctions';

// すべての関数を配列にまとめる
export const ALL_FUNCTIONS = [
  // 01. 数学・三角関数
  ...Object.values({
    SUMIF, COUNTIF, AVERAGEIF, SUM, AVERAGE, COUNT, MAX, MIN, ROUND,
    ABS, SQRT, POWER, MOD, INT, TRUNC, RAND, RANDBETWEEN,
    LOG, LOG10, LN, EXP, ROUNDUP, ROUNDDOWN,
    CEILING, FLOOR, SIGN, FACT, SUMIFS, COUNTIFS, AVERAGEIFS, PRODUCT, MROUND,
    COMBIN, PERMUT, GCD, LCM, QUOTIENT,
    SUMSQ, SUMPRODUCT, EVEN, ODD, ARABIC, ROMAN, COMBINA, FACTDOUBLE, SQRTPI,
    SUMX2MY2, SUMX2PY2, SUMXMY2, MULTINOMIAL, PERMUTATIONA, BASE, DECIMAL, SUBTOTAL,
    AGGREGATE, CEILING_MATH, CEILING_PRECISE, FLOOR_MATH, FLOOR_PRECISE, ISO_CEILING, SERIESSUM,
    RANDARRAY, SEQUENCE,
    SIN, COS, TAN, ASIN, ACOS, ATAN, ATAN2, SINH, COSH, TANH,
    ASINH, ACOSH, ATANH, CSC, SEC, COT, ACOT, CSCH, SECH, COTH, ACOTH
  }),

  // 02. 統計関数
  ...Object.values({
    MEDIAN, MODE, COUNTA, COUNTBLANK, LARGE, SMALL, RANK,
    GEOMEAN, HARMEAN, TRIMMEAN,
    DEVSQ, KURT, SKEW, PEARSON, PERCENTRANK,
    STANDARDIZE, COVARIANCE_P, COVARIANCE_S, SKEW_P,
    PERCENTILE_INC, PERCENTILE_EXC, QUARTILE_INC, QUARTILE_EXC,
    NORM_DIST, NORM_INV, NORM_S_DIST, NORM_S_INV, LOGNORM_DIST, LOGNORM_INV,
    T_DIST, T_DIST_2T, T_DIST_RT, CHISQ_DIST, CHISQ_DIST_RT,
    F_DIST, F_DIST_RT, BETA_DIST,
    GAMMA_DIST, EXPON_DIST, WEIBULL_DIST, BINOM_DIST,
    NEGBINOM_DIST, POISSON_DIST, HYPGEOM_DIST, CONFIDENCE_NORM,
    CONFIDENCE_T, Z_TEST, T_TEST, F_TEST, CHISQ_TEST,
    T_INV, T_INV_2T, CHISQ_INV, CHISQ_INV_RT, F_INV, F_INV_RT, BETA_INV, GAMMA_INV
  }),

  // 03. 文字列操作関数
  ...Object.values({
    CONCATENATE, CONCAT, LEFT, RIGHT, MID, LEN, UPPER, LOWER, TRIM, SUBSTITUTE, FIND, SEARCH, TEXTJOIN,
    PROPER, VALUE, TEXT, REPT, REPLACE, CHAR, CODE, EXACT, CLEAN, T, FIXED, NUMBERVALUE, DOLLAR, UNICHAR, UNICODE,
    LENB, FINDB, SEARCHB, REPLACEB, TEXTBEFORE, TEXTAFTER, TEXTSPLIT, ASC, JIS, DBCS, PHONETIC, BAHTTEXT
  }),

  // 04. 日付・時刻関数
  ...Object.values({
    DATEDIF, NETWORKDAYS, TODAY, NOW, DATE, YEAR, MONTH, DAY, WEEKDAY, DAYS, EDATE, EOMONTH, TIME, HOUR, MINUTE, SECOND, WEEKNUM, DAYS360, YEARFRAC, DATEVALUE, TIMEVALUE, ISOWEEKNUM, NETWORKDAYS_INTL, WORKDAY, WORKDAY_INTL
  }),

  // 05. 論理関数
  ...Object.values({
    IF, AND, OR, NOT, IFS, XOR, TRUE, FALSE, IFERROR, IFNA, SWITCH, LET
  }),

  // 06. 検索・参照関数
  ...Object.values({
    VLOOKUP, HLOOKUP, INDEX, MATCH, LOOKUP, XLOOKUP, OFFSET, INDIRECT, CHOOSE, TRANSPOSE, FILTER, SORT, UNIQUE
  }),

  // 07. 情報関数
  ...Object.values({
    ISBLANK, ISERROR, ISNA, ISTEXT, ISNUMBER, ISLOGICAL, ISEVEN, ISODD, TYPE, N,
    ISERR, ISNONTEXT, ISREF, ISFORMULA, NA, ERROR_TYPE, INFO, SHEET, SHEETS, CELL
  }),

  // 08. データベース関数
  ...Object.values({
    DSUM, DAVERAGE, DCOUNT, DCOUNTA, DMAX, DMIN, DPRODUCT, DGET
  }),

  // 09. 財務関数
  ...Object.values({
    PMT, PV, FV, NPV, IRR, PPMT, IPMT, RATE
  }),

  // 13. エンジニアリング関数
  ...Object.values({
    CONVERT, BIN2DEC, DEC2BIN, HEX2DEC, DEC2HEX, BIN2HEX, HEX2BIN, OCT2DEC, DEC2OCT
  }),

  // 15. その他 (行列関数など)
  ...Object.values({
    MDETERM, MINVERSE, MMULT, MUNIT
  }),

  // 16. Google Sheets関数
  ...Object.values({
    JOIN, ARRAYFORMULA, QUERY, REGEXMATCH, REGEXEXTRACT, REGEXREPLACE, FLATTEN
  })
] as CustomFormula[];

// カテゴリ別の関数分類
export const FUNCTION_CATEGORIES = {
  mathTrigonometry: [
    SUMIF, COUNTIF, AVERAGEIF, SUM, AVERAGE, COUNT, MAX, MIN, ROUND,
    ABS, SQRT, POWER, MOD, INT, TRUNC, RAND, RANDBETWEEN,
    LOG, LOG10, LN, EXP, ROUNDUP, ROUNDDOWN,
    CEILING, FLOOR, SIGN, FACT, SUMIFS, COUNTIFS, AVERAGEIFS, PRODUCT, MROUND,
    COMBIN, PERMUT, GCD, LCM, QUOTIENT,
    SUMSQ, SUMPRODUCT, EVEN, ODD, ARABIC, ROMAN, COMBINA, FACTDOUBLE, SQRTPI,
    SUMX2MY2, SUMX2PY2, SUMXMY2, MULTINOMIAL, PERMUTATIONA, BASE, DECIMAL, SUBTOTAL,
    AGGREGATE, CEILING_MATH, CEILING_PRECISE, FLOOR_MATH, FLOOR_PRECISE, ISO_CEILING, SERIESSUM,
    RANDARRAY, SEQUENCE,
    SIN, COS, TAN, ASIN, ACOS, ATAN, ATAN2, SINH, COSH, TANH,
    ASINH, ACOSH, ATANH, CSC, SEC, COT, ACOT, CSCH, SECH, COTH, ACOTH
  ] as CustomFormula[],
  statistical: [
    MEDIAN, MODE, COUNTA, COUNTBLANK, LARGE, SMALL, RANK,
    GEOMEAN, HARMEAN, TRIMMEAN,
    DEVSQ, KURT, SKEW, PEARSON, PERCENTRANK,
    STANDARDIZE, COVARIANCE_P, COVARIANCE_S, SKEW_P,
    PERCENTILE_INC, PERCENTILE_EXC, QUARTILE_INC, QUARTILE_EXC,
    NORM_DIST, NORM_INV, NORM_S_DIST, NORM_S_INV, LOGNORM_DIST, LOGNORM_INV,
    T_DIST, T_DIST_2T, T_DIST_RT, CHISQ_DIST, CHISQ_DIST_RT,
    F_DIST, F_DIST_RT, BETA_DIST,
    GAMMA_DIST, EXPON_DIST, WEIBULL_DIST, BINOM_DIST,
    NEGBINOM_DIST, POISSON_DIST, HYPGEOM_DIST, CONFIDENCE_NORM,
    CONFIDENCE_T, Z_TEST, T_TEST, F_TEST, CHISQ_TEST,
    T_INV, T_INV_2T, CHISQ_INV, CHISQ_INV_RT, F_INV, F_INV_RT, BETA_INV, GAMMA_INV
  ] as CustomFormula[],
  text: [
    CONCATENATE, CONCAT, LEFT, RIGHT, MID, LEN, UPPER, LOWER, TRIM, SUBSTITUTE, FIND, SEARCH, TEXTJOIN, SPLIT,
    PROPER, VALUE, TEXT, REPT, REPLACE, CHAR, CODE, EXACT, CLEAN, T, FIXED, NUMBERVALUE, DOLLAR, UNICHAR, UNICODE,
    LENB, FINDB, SEARCHB, REPLACEB, TEXTBEFORE, TEXTAFTER, TEXTSPLIT, ASC, JIS, DBCS, PHONETIC, BAHTTEXT
  ] as CustomFormula[],
  datetime: [
    DATEDIF, NETWORKDAYS, TODAY, NOW, DATE, YEAR, MONTH, DAY, WEEKDAY, DAYS, EDATE, EOMONTH, TIME, HOUR, MINUTE, SECOND, WEEKNUM, DAYS360, YEARFRAC, DATEVALUE, TIMEVALUE, ISOWEEKNUM, NETWORKDAYS_INTL, WORKDAY, WORKDAY_INTL
  ] as CustomFormula[],
  logical: [IF, AND, OR, NOT, IFS, XOR, TRUE, FALSE, IFERROR, IFNA, SWITCH, LET] as CustomFormula[],
  lookupReference: [VLOOKUP, HLOOKUP, INDEX, MATCH, LOOKUP, XLOOKUP, OFFSET, INDIRECT, CHOOSE, TRANSPOSE, FILTER, SORT, UNIQUE] as CustomFormula[],
  information: [
    ISBLANK, ISERROR, ISNA, ISTEXT, ISNUMBER, ISLOGICAL, ISEVEN, ISODD, TYPE, N,
    ISERR, ISNONTEXT, ISREF, ISFORMULA, NA, ERROR_TYPE, INFO, SHEET, SHEETS, CELL
  ] as CustomFormula[],
  database: [DSUM, DAVERAGE, DCOUNT, DCOUNTA, DMAX, DMIN, DPRODUCT, DGET] as CustomFormula[],
  financial: [PMT, PV, FV, NPV, IRR, PPMT, IPMT, RATE] as CustomFormula[],
  engineering: [CONVERT, BIN2DEC, DEC2BIN, HEX2DEC, DEC2HEX, BIN2HEX, HEX2BIN, OCT2DEC, DEC2OCT] as CustomFormula[],
  matrix: [MDETERM, MINVERSE, MMULT, MUNIT] as CustomFormula[],
  googleSheets: [JOIN, ARRAYFORMULA, QUERY, REGEXMATCH, REGEXEXTRACT, REGEXREPLACE, FLATTEN] as CustomFormula[]
};

// 全ての関数（HyperFormulaは使用しない）
export const MANUAL_FUNCTIONS = ALL_FUNCTIONS;

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
  
  if (FUNCTION_CATEGORIES.mathTrigonometry.some(f => f.name === name)) return 'math';
  if (FUNCTION_CATEGORIES.statistical.some(f => f.name === name)) return 'statistics';
  if (FUNCTION_CATEGORIES.text.some(f => f.name === name)) return 'text';
  if (FUNCTION_CATEGORIES.datetime.some(f => f.name === name)) return 'date';
  if (FUNCTION_CATEGORIES.logical.some(f => f.name === name)) return 'logic';
  if (FUNCTION_CATEGORIES.lookupReference.some(f => f.name === name)) return 'lookup';
  if (FUNCTION_CATEGORIES.information.some(f => f.name === name)) return 'information';
  if (FUNCTION_CATEGORIES.database.some(f => f.name === name)) return 'database';
  if (FUNCTION_CATEGORIES.financial.some(f => f.name === name)) return 'financial';
  if (FUNCTION_CATEGORIES.engineering.some(f => f.name === name)) return 'engineering';
  if (FUNCTION_CATEGORIES.matrix.some(f => f.name === name)) return 'matrix';
  if (FUNCTION_CATEGORIES.googleSheets.some(f => f.name === name)) return 'googleSheets';
  
  return 'other';
};

// エクスポート
export * from './shared/types';
export * from './shared/utils';
export { 
  SUMIF, COUNTIF, AVERAGEIF, SUM, AVERAGE, COUNT, MAX, MIN, ROUND,
  ABS, SQRT, POWER, MOD, INT, TRUNC, RAND, RANDBETWEEN,
  LOG, LOG10, LN, EXP, ROUNDUP, ROUNDDOWN,
  CEILING, FLOOR, SIGN, FACT, SUMIFS, COUNTIFS, AVERAGEIFS, PRODUCT, MROUND,
  COMBIN, PERMUT, GCD, LCM, QUOTIENT,
  SUMSQ, SUMPRODUCT, EVEN, ODD, ARABIC, ROMAN, COMBINA, FACTDOUBLE, SQRTPI,
  SUMX2MY2, SUMX2PY2, SUMXMY2, MULTINOMIAL, PERMUTATIONA, BASE, DECIMAL, SUBTOTAL,
  AGGREGATE, CEILING_MATH, CEILING_PRECISE, FLOOR_MATH, FLOOR_PRECISE, ISO_CEILING, SERIESSUM,
  RANDARRAY, SEQUENCE
} from './01-math-trigonometry/mathFunctions';
export * from './01-math-trigonometry/trigonometricFunctions';
export * from './02-statistical/basicStatisticsFunctions';
export * from './02-statistical/advancedStatisticsFunctions';
export * from './02-statistical/distributionFunctions';
export * from './02-statistical/inverseDistributionFunctions';
export * from './03-text/textFunctions';
export * from './04-datetime/dateFunctions';
export * from './05-logical/logicFunctions';
export * from './05-logical/newLogicalFunctions';
export * from './06-lookup-reference/lookupFunctions';
export * from './07-information/informationFunctions';
export * from './08-database/databaseFunctions';
export * from './09-financial/financialFunctions';
export * from './13-engineering/engineeringFunctions';
export * from './15-others/matrixFunctions';
export * from './16-google-sheets/googleSheetsFunctions';