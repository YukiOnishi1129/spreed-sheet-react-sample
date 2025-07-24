// 財務関数の実装

import type { CustomFormula } from '../shared/types';
import { FormulaError } from '../shared/types';
import { getCellValue, getCellRangeValues } from '../shared/utils';

// PMT関数の実装（定期支払額）
export const PMT: CustomFormula = {
  name: 'PMT',
  pattern: /PMT\(([^,]+),\s*([^,]+),\s*([^,]+)(?:,\s*([^,]+))?(?:,\s*([^)]+))?\)/i,
  calculate: () => null // HyperFormulaが処理
};

// PV関数の実装（現在価値）
export const PV: CustomFormula = {
  name: 'PV',
  pattern: /PV\(([^,]+),\s*([^,]+),\s*([^,]+)(?:,\s*([^,]+))?(?:,\s*([^)]+))?\)/i,
  calculate: () => null // HyperFormulaが処理
};

// FV関数の実装（将来価値）
export const FV: CustomFormula = {
  name: 'FV',
  pattern: /FV\(([^,]+),\s*([^,]+),\s*([^,]+)(?:,\s*([^,]+))?(?:,\s*([^)]+))?\)/i,
  calculate: () => null // HyperFormulaが処理
};

// NPV関数の実装（正味現在価値）
export const NPV: CustomFormula = {
  name: 'NPV',
  pattern: /NPV\(([^,]+),\s*(.+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// IRR関数の実装（内部収益率）
export const IRR: CustomFormula = {
  name: 'IRR',
  pattern: /IRR\(([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches, context) => {
    const valuesRef = matches[1].trim();
    const guessRef = matches[2]?.trim();
    
    let guess = 0.1; // デフォルトの推定値
    
    // 推定値を取得（オプション）
    if (guessRef) {
      if (guessRef.match(/^[A-Z]+\d+$/)) {
        const cellValue = getCellValue(guessRef, context);
        guess = parseFloat(String(cellValue ?? '0.1'));
      } else {
        guess = parseFloat(guessRef);
      }
    }
    
    // 値の配列を取得
    const values: number[] = [];
    
    if (valuesRef.includes(':')) {
      // セル範囲の場合
      const rangeValues = getCellRangeValues(valuesRef, context);
      rangeValues.forEach(value => {
        const num = typeof value === 'string' ? parseFloat(value) : Number(value);
        if (!isNaN(num)) {
          values.push(num);
        }
      });
    } else {
      // 複数の引数の場合
      const args = valuesRef.split(',').map(arg => arg.trim());
      for (const arg of args) {
        if (arg.match(/^[A-Z]+\d+$/)) {
          const cellValue = getCellValue(arg, context);
          const num = typeof cellValue === 'string' ? parseFloat(cellValue) : Number(cellValue);
          if (!isNaN(num)) {
            values.push(num);
          }
        } else {
          const num = parseFloat(arg);
          if (!isNaN(num)) {
            values.push(num);
          }
        }
      }
    }
    
    if (values.length === 0) return FormulaError.VALUE;
    
    // 正負の値が両方必要
    let hasPositive = false;
    let hasNegative = false;
    for (const value of values) {
      if (value > 0) hasPositive = true;
      if (value < 0) hasNegative = true;
    }
    if (!hasPositive || !hasNegative) return FormulaError.NUM;
    
    // ニュートン・ラフソン法でIRRを計算
    let rate = guess;
    const maxIterations = 100;
    const precision = 0.00000001;
    
    for (let i = 0; i < maxIterations; i++) {
      let npv = 0;
      let derivative = 0;
      
      for (let j = 0; j < values.length; j++) {
        const pv = values[j] / Math.pow(1 + rate, j);
        npv += pv;
        derivative -= j * pv / (1 + rate);
      }
      
      if (Math.abs(npv) < precision) {
        return rate;
      }
      
      const newRate = rate - npv / derivative;
      
      if (Math.abs(newRate - rate) < precision) {
        return newRate;
      }
      
      rate = newRate;
    }
    
    return FormulaError.NUM; // 収束しない場合
  }
};

// PPMT関数の実装（元金返済額）
export const PPMT: CustomFormula = {
  name: 'PPMT',
  pattern: /PPMT\(([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,]+)(?:,\s*([^,]+))?(?:,\s*([^)]+))?\)/i,
  calculate: () => null // HyperFormulaが処理
};

// IPMT関数の実装（利息支払額）
export const IPMT: CustomFormula = {
  name: 'IPMT',
  pattern: /IPMT\(([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,]+)(?:,\s*([^,]+))?(?:,\s*([^)]+))?\)/i,
  calculate: () => null // HyperFormulaが処理
};