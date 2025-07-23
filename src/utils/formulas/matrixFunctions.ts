// 行列関数の実装

import type { CustomFormula } from './types';

// MDETERM関数（行列式を計算）
export const MDETERM: CustomFormula = {
  name: 'MDETERM',
  pattern: /MDETERM\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// MINVERSE関数（逆行列を計算）
export const MINVERSE: CustomFormula = {
  name: 'MINVERSE',
  pattern: /MINVERSE\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// MMULT関数（行列の積を計算）
export const MMULT: CustomFormula = {
  name: 'MMULT',
  pattern: /MMULT\(([^,]+),\s*([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// MUNIT関数（単位行列を作成）
export const MUNIT: CustomFormula = {
  name: 'MUNIT',
  pattern: /MUNIT\(([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};