// Google Sheets specific functions

import type { CustomFormula, FormulaResult } from '../shared/types';

// JOIN関数の実装（配列を文字列に結合）
export const JOIN: CustomFormula = {
  name: 'JOIN',
  pattern: /\bJOIN\(([^,]+),\s*(.+)\)/i,
  calculate: (): FormulaResult => {
    return '#N/A - Google Sheets specific functions not supported';
  }
};

// ARRAYFORMULA関数の実装（配列数式を作成）
export const ARRAYFORMULA: CustomFormula = {
  name: 'ARRAYFORMULA',
  pattern: /ARRAYFORMULA\((.+)\)/i,
  calculate: (): FormulaResult => {
    return '#N/A - Google Sheets specific functions not supported';
  }
};

// QUERY関数の実装（簡易版 - SQLライクなクエリ）
export const QUERY: CustomFormula = {
  name: 'QUERY',
  pattern: /QUERY\(([^,]+)(?:,\s*"([^"]*)")?(?:,\s*(-?\d+))?\)/i,
  calculate: (): FormulaResult => {
    return '#N/A - Google Sheets specific functions not supported';
  }
};


// REGEXMATCH関数の実装（正規表現マッチング）
export const REGEXMATCH: CustomFormula = {
  name: 'REGEXMATCH',
  pattern: /REGEXMATCH\(([^,]+),\s*"([^"]+)"\)/i,
  calculate: (): FormulaResult => {
    return '#N/A - Google Sheets specific functions not supported';
  }
};

// REGEXEXTRACT関数の実装（正規表現による抽出）
export const REGEXEXTRACT: CustomFormula = {
  name: 'REGEXEXTRACT',
  pattern: /REGEXEXTRACT\(([^,]+),\s*"([^"]+)"\)/i,
  calculate: (): FormulaResult => {
    return '#N/A - Google Sheets specific functions not supported';
  }
};

// REGEXREPLACE関数の実装（正規表現による置換）
export const REGEXREPLACE: CustomFormula = {
  name: 'REGEXREPLACE',
  pattern: /REGEXREPLACE\(([^,]+),\s*"([^"]+)",\s*"([^"]*)"\)/i,
  calculate: (): FormulaResult => {
    return '#N/A - Google Sheets specific functions not supported';
  }
};

// FLATTEN関数の実装（配列を1次元化）
export const FLATTEN: CustomFormula = {
  name: 'FLATTEN',
  pattern: /FLATTEN\((.+)\)/i,
  calculate: (): FormulaResult => {
    return '#N/A - Google Sheets specific functions not supported';
  }
};