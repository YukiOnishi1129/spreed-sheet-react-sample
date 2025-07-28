// 財務関数の実装

import type { CustomFormula, FormulaContext, FormulaResult } from '../shared/types';
import { FormulaError } from '../shared/types';
import { getCellValue, getCellRangeValues } from '../shared/utils';

// PMT関数の実装（定期支払額）
export const PMT: CustomFormula = {
  name: 'PMT',
  pattern: /PMT\(([^,)]+),\s*([^,)]+),\s*([^,)]+)(?:,\s*([^,)]+))?(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, rateRef, nperRef, pvRef, fvRef, typeRef] = matches;
    
    try {
      // 金利（期間金利）
      const rate = parseFloat(getCellValue(rateRef.trim(), context)?.toString() ?? rateRef.trim());
      
      // 期間数
      const nper = parseFloat(getCellValue(nperRef.trim(), context)?.toString() ?? nperRef.trim());
      // 現在価値（借入額）
      let pv: number;
      const pvValue = pvRef.trim();
      if (pvValue.startsWith('-') && pvValue.length > 1) {
        const cellRef = pvValue.substring(1);
        const cellValue = getCellValue(cellRef, context);
        pv = cellValue !== null && cellValue !== undefined ? -parseFloat(cellValue.toString()) : NaN;
      } else {
        pv = parseFloat(getCellValue(pvValue, context)?.toString() ?? pvValue);
      }
      
      // 将来価値（デフォルト0）
      const fv = fvRef ? parseFloat(getCellValue(fvRef.trim(), context)?.toString() ?? fvRef.trim()) : 0;
      // 支払タイプ（0=期末、1=期初、デフォルト0）
      const type = typeRef ? parseFloat(getCellValue(typeRef.trim(), context)?.toString() ?? typeRef.trim()) : 0;
      
      if (isNaN(rate) || isNaN(nper) || isNaN(pv)) {
        return FormulaError.VALUE;
      }
      
      if (nper === 0) {
        return FormulaError.DIV0;
      }
      
      let pmt: number;
      
      if (rate === 0) {
        // 金利が0の場合
        pmt = -(pv + fv) / nper;
      } else {
        // 通常の計算
        const factor = Math.pow(1 + rate, nper);
        pmt = -(pv * factor + fv) * rate / (factor - 1);
        
        if (type === 1) {
          // 期初払いの場合
          pmt = pmt / (1 + rate);
        }
      }
      
      return pmt;
    } catch (error) {
      return FormulaError.VALUE;
    }
  }
};

// PV関数の実装（現在価値）
export const PV: CustomFormula = {
  name: 'PV',
  pattern: /PV\(([^,]+),\s*([^,]+),\s*([^,]+)(?:,\s*([^,]+))?(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, rateRef, nperRef, pmtRef, fvRef, typeRef] = matches;
    
    try {
      // 金利（期間金利）
      const rate = parseFloat(getCellValue(rateRef.trim(), context)?.toString() ?? rateRef.trim());
      // 期間数
      const nper = parseFloat(getCellValue(nperRef.trim(), context)?.toString() ?? nperRef.trim());
      // 定期支払額
      let pmt: number;
      const pmtValue = pmtRef.trim();
      if (pmtValue.startsWith('-') && pmtValue.length > 1) {
        const cellRef = pmtValue.substring(1);
        const cellValue = getCellValue(cellRef, context);
        pmt = cellValue !== null && cellValue !== undefined ? -parseFloat(cellValue.toString()) : NaN;
      } else {
        pmt = parseFloat(getCellValue(pmtValue, context)?.toString() ?? pmtValue);
      }
      // 将来価値（デフォルト0）
      const fv = fvRef ? parseFloat(getCellValue(fvRef.trim(), context)?.toString() ?? fvRef.trim()) : 0;
      // 支払タイプ（0=期末、1=期初、デフォルト0）
      const type = typeRef ? parseFloat(getCellValue(typeRef.trim(), context)?.toString() ?? typeRef.trim()) : 0;
      
      if (isNaN(rate) || isNaN(nper) || isNaN(pmt)) {
        return FormulaError.VALUE;
      }
      
      let pv: number;
      
      if (rate === 0) {
        // 金利が0の場合
        pv = -pmt * nper - fv;
      } else {
        // 通常の計算
        const factor = Math.pow(1 + rate, nper);
        let pvAnnuity = pmt * (1 - 1 / factor) / rate;
        
        if (type === 1) {
          // 期初払いの場合
          pvAnnuity *= (1 + rate);
        }
        
        pv = -pvAnnuity - fv / factor;
      }
      
      return pv;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// FV関数の実装（将来価値）
export const FV: CustomFormula = {
  name: 'FV',
  pattern: /FV\(([^,]+),\s*([^,]+),\s*([^,]+)(?:,\s*([^,]+))?(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, rateRef, nperRef, pmtRef, pvRef, typeRef] = matches;
    
    try {
      // 金利（期間金利）
      const rate = parseFloat(getCellValue(rateRef.trim(), context)?.toString() ?? rateRef.trim());
      // 期間数
      const nper = parseFloat(getCellValue(nperRef.trim(), context)?.toString() ?? nperRef.trim());
      // 定期支払額
      let pmt: number;
      const pmtValue = pmtRef.trim();
      if (pmtValue.startsWith('-') && pmtValue.length > 1) {
        const cellRef = pmtValue.substring(1);
        const cellValue = getCellValue(cellRef, context);
        pmt = cellValue !== null && cellValue !== undefined ? -parseFloat(cellValue.toString()) : NaN;
      } else {
        pmt = parseFloat(getCellValue(pmtValue, context)?.toString() ?? pmtValue);
      }
      // 現在価値（デフォルト0）
      const pv = pvRef ? parseFloat(getCellValue(pvRef.trim(), context)?.toString() ?? pvRef.trim()) : 0;
      // 支払タイプ（0=期末、1=期初、デフォルト0）
      const type = typeRef ? parseFloat(getCellValue(typeRef.trim(), context)?.toString() ?? typeRef.trim()) : 0;
      
      if (isNaN(rate) || isNaN(nper) || isNaN(pmt)) {
        return FormulaError.VALUE;
      }
      
      let fv: number;
      
      if (rate === 0) {
        // 金利が0の場合
        fv = -pv - pmt * nper;
      } else {
        // 通常の計算
        const factor = Math.pow(1 + rate, nper);
        let fvAnnuity = pmt * (factor - 1) / rate;
        
        if (type === 1) {
          // 期初払いの場合
          fvAnnuity *= (1 + rate);
        }
        
        fv = -pv * factor - fvAnnuity;
      }
      
      return fv;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// NPV関数の実装（正味現在価値）
export const NPV: CustomFormula = {
  name: 'NPV',
  pattern: /\bNPV\(([^,]+),\s*(.+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, rateRef, valuesRef] = matches;
    
    try {
      // 割引率
      const rate = parseFloat(getCellValue(rateRef.trim(), context)?.toString() ?? rateRef.trim());
      
      if (isNaN(rate)) {
        return FormulaError.VALUE;
      }
      
      // キャッシュフローの値を取得
      const values: number[] = [];
      
      if (valuesRef.includes(':')) {
        // セル範囲の場合
        const rangeValues = getCellRangeValues(valuesRef.trim(), context);
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
      
      if (values.length === 0) {
        return FormulaError.VALUE;
      }
      
      // NPVを計算
      let npv = 0;
      for (let i = 0; i < values.length; i++) {
        npv += values[i] / Math.pow(1 + rate, i + 1);
      }
      
      return npv;
    } catch {
      return FormulaError.VALUE;
    }
  }
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
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, rateRef, perRef, nperRef, pvRef, fvRef, typeRef] = matches;
    
    try {
      // 金利（期間金利）
      const rate = parseFloat(getCellValue(rateRef.trim(), context)?.toString() ?? rateRef.trim());
      // 対象期間
      const per = parseFloat(getCellValue(perRef.trim(), context)?.toString() ?? perRef.trim());
      // 総期間数
      const nper = parseFloat(getCellValue(nperRef.trim(), context)?.toString() ?? nperRef.trim());
      // 現在価値（借入額）
      const pv = parseFloat(getCellValue(pvRef.trim(), context)?.toString() ?? pvRef.trim());
      // 将来価値（デフォルト0）
      const fv = fvRef ? parseFloat(getCellValue(fvRef.trim(), context)?.toString() ?? fvRef.trim()) : 0;
      // 支払タイプ（0=期末、1=期初、デフォルト0）
      const type = typeRef ? parseFloat(getCellValue(typeRef.trim(), context)?.toString() ?? typeRef.trim()) : 0;
      
      if (isNaN(rate) || isNaN(per) || isNaN(nper) || isNaN(pv)) {
        return FormulaError.VALUE;
      }
      
      if (per < 1 || per > nper) {
        return FormulaError.NUM;
      }
      
      if (rate === 0) {
        // 金利が0の場合
        return -(pv + fv) / nper;
      }
      
      // PMTを計算
      const factor = Math.pow(1 + rate, nper);
      let pmt = -(pv * factor + fv) * rate / (factor - 1);
      
      if (type === 1) {
        pmt = pmt / (1 + rate);
      }
      
      // IPMT（利息部分）を計算
      let ipmt: number;
      if (type === 1 && per === 1) {
        ipmt = 0;
      } else {
        const adjustedPer = type === 1 ? per - 1 : per;
        const fvAtPeriod = pv * Math.pow(1 + rate, adjustedPer) + pmt * (Math.pow(1 + rate, adjustedPer) - 1) / rate;
        ipmt = -fvAtPeriod * rate;
      }
      
      // PPMT = PMT - IPMT
      return pmt - ipmt;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// IPMT関数の実装（利息支払額）
export const IPMT: CustomFormula = {
  name: 'IPMT',
  pattern: /IPMT\(([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,]+)(?:,\s*([^,]+))?(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, rateRef, perRef, nperRef, pvRef, fvRef, typeRef] = matches;
    
    try {
      // 金利（期間金利）
      const rate = parseFloat(getCellValue(rateRef.trim(), context)?.toString() ?? rateRef.trim());
      // 対象期間
      const per = parseFloat(getCellValue(perRef.trim(), context)?.toString() ?? perRef.trim());
      // 総期間数
      const nper = parseFloat(getCellValue(nperRef.trim(), context)?.toString() ?? nperRef.trim());
      // 現在価値（借入額）
      const pv = parseFloat(getCellValue(pvRef.trim(), context)?.toString() ?? pvRef.trim());
      // 将来価値（デフォルト0）
      const fv = fvRef ? parseFloat(getCellValue(fvRef.trim(), context)?.toString() ?? fvRef.trim()) : 0;
      // 支払タイプ（0=期末、1=期初、デフォルト0）
      const type = typeRef ? parseFloat(getCellValue(typeRef.trim(), context)?.toString() ?? typeRef.trim()) : 0;
      
      if (isNaN(rate) || isNaN(per) || isNaN(nper) || isNaN(pv)) {
        return FormulaError.VALUE;
      }
      
      if (per < 1 || per > nper) {
        return FormulaError.NUM;
      }
      
      if (rate === 0) {
        // 金利が0の場合、利息はゼロ
        return 0;
      }
      
      // PMTを計算
      const factor = Math.pow(1 + rate, nper);
      let pmt = -(pv * factor + fv) * rate / (factor - 1);
      
      if (type === 1) {
        pmt = pmt / (1 + rate);
      }
      
      // 利息部分を計算
      if (type === 1 && per === 1) {
        return 0;
      }
      
      const adjustedPer = type === 1 ? per - 1 : per;
      const fvAtPeriod = pv * Math.pow(1 + rate, adjustedPer) + pmt * (Math.pow(1 + rate, adjustedPer) - 1) / rate;
      
      return -fvAtPeriod * rate;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// RATE関数の実装（金利計算）
export const RATE: CustomFormula = {
  name: 'RATE',
  pattern: /RATE\(([^,]+),\s*([^,]+),\s*([^,]+)(?:,\s*([^,]+))?(?:,\s*([^,]+))?(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, nperRef, pmtRef, pvRef, fvRef, typeRef, guessRef] = matches;
    
    try {
      // 期間数
      const nper = parseFloat(getCellValue(nperRef.trim(), context)?.toString() ?? nperRef.trim());
      // 定期支払額
      const pmt = parseFloat(getCellValue(pmtRef.trim(), context)?.toString() ?? pmtRef.trim());
      // 現在価値
      const pv = parseFloat(getCellValue(pvRef.trim(), context)?.toString() ?? pvRef.trim());
      // 将来価値（デフォルト0）
      const fv = fvRef ? parseFloat(getCellValue(fvRef.trim(), context)?.toString() ?? fvRef.trim()) : 0;
      // 支払タイプ（0=期末、1=期初、デフォルト0）
      const type = typeRef ? parseFloat(getCellValue(typeRef.trim(), context)?.toString() ?? typeRef.trim()) : 0;
      // 推定値（デフォルト0.1）
      const guess = guessRef ? parseFloat(getCellValue(guessRef.trim(), context)?.toString() ?? guessRef.trim()) : 0.1;
      
      if (isNaN(nper) || isNaN(pmt) || isNaN(pv) || nper <= 0) {
        return FormulaError.VALUE;
      }
      
      // ニュートン・ラフソン法で金利を求める
      let rate = guess;
      const maxIterations = 100;
      const precision = 0.00000001;
      
      for (let i = 0; i < maxIterations; i++) {
        let f: number;
        let df: number;
        
        if (rate === 0) {
          f = pv + pmt * nper + fv;
          df = nper;
        } else {
          const factor = Math.pow(1 + rate, nper);
          const annuityFactor = (factor - 1) / rate;
          
          f = pv + pmt * annuityFactor + fv / factor;
          if (type === 1) {
            f += pmt * rate * annuityFactor;
          }
          
          // 導関数
          df = pmt * (nper * factor - annuityFactor) / (rate * rate * factor) - nper * fv / (factor * (1 + rate));
          if (type === 1) {
            df += pmt * (annuityFactor + rate * (nper * factor - annuityFactor) / (rate * rate * factor));
          }
        }
        
        if (Math.abs(f) < precision) {
          return rate;
        }
        
        if (Math.abs(df) < precision) {
          return FormulaError.NUM;
        }
        
        const newRate = rate - f / df;
        
        if (Math.abs(newRate - rate) < precision) {
          return newRate;
        }
        
        rate = newRate;
        
        // 金利が負の場合や極端に大きい場合は収束しない
        if (rate < -1 || rate > 100) {
          return FormulaError.NUM;
        }
      }
      
      return FormulaError.NUM; // 収束しない場合
    } catch {
      return FormulaError.VALUE;
    }
  }
};