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
      const pv = parseFloat(getCellValue(pvRef.trim(), context)?.toString() ?? pvRef.trim());
      
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
        // Excel-compatible PMT calculation - exact Excel formula
        // PMT = ((PV * rate) + (FV * rate / ((1 + rate)^n - 1))) / (1 - (1 + rate)^(-n)) * -1
        
        if (type === 0) {
          // 期末払い (ordinary annuity) - exact Excel PMT formula
          // Excel formula: PMT = (PV * rate + FV * rate / ((1 + rate)^n - 1)) / (1 - (1 + rate)^(-n)) * -1
          const pvFactor = (1 - Math.pow(1 + rate, -nper)) / rate;
          const fvFactor = Math.pow(1 + rate, -nper);
          pmt = -(pv / pvFactor + fv * fvFactor);
        } else {
          // 期初払い (annuity due) - Excel formula for beginning payments
          const pvFactor = (1 - Math.pow(1 + rate, -nper)) / rate * (1 + rate);
          const fvFactor = Math.pow(1 + rate, -nper);
          pmt = -(pv / pvFactor + fv * fvFactor / (1 + rate));
        }
      }
      
      return pmt;
    } catch {
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
      const pmt = parseFloat(getCellValue(pmtRef.trim(), context)?.toString() ?? pmtRef.trim());
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
        pv = -(pmt * nper + fv);
      } else {
        // 通常の計算
        const factor = Math.pow(1 + rate, nper);
        pv = -(pmt * (1 + rate * type) * (factor - 1) / rate + fv) / factor;
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
      const pmt = parseFloat(getCellValue(pmtRef.trim(), context)?.toString() ?? pmtRef.trim());
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
        fv = -(pv + pmt * nper);
      } else {
        // 通常の計算
        const factor = Math.pow(1 + rate, nper);
        fv = -(pv * factor + pmt * (1 + rate * type) * (factor - 1) / rate);
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
        // 複数の引数の場合 - handle the entire remaining string as comma-separated values
        // Split by comma but respect nested parentheses
        const args: string[] = [];
        let current = '';
        let depth = 0;
        
        for (let i = 0; i < valuesRef.length; i++) {
          const char = valuesRef[i];
          if (char === '(') depth++;
          else if (char === ')') depth--;
          
          if (char === ',' && depth === 0) {
            args.push(current.trim());
            current = '';
          } else {
            current += char;
          }
        }
        if (current) args.push(current.trim());
        
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
      // Excel's NPV assumes all cash flows happen at the end of periods
      // The first value is at the end of period 1, not at time 0
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
    
    try {
      let guess = 0.1; // デフォルトの推定値
      
      // 推定値を取得（オプション）
      if (guessRef) {
        if (guessRef.match(/^[A-Z]+\d+$/)) {
          const cellValue = getCellValue(guessRef, context);
          guess = parseFloat(String(cellValue ?? '0.1'));
        } else {
          guess = parseFloat(guessRef);
        }
        
        if (isNaN(guess)) {
          return FormulaError.VALUE;
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
    
    // Excel-compatible IRR calculation using Secant method for better convergence
    let rate = guess;
    let prevRate = guess * 1.1;
    const maxIterations = 100;
    const tolerance = 0.00000001;
    
    for (let i = 0; i < maxIterations; i++) {
      let npv = 0;
      let prevNpv = 0;
      
      // Calculate NPV at current rate
      for (let j = 0; j < values.length; j++) {
        npv += values[j] / Math.pow(1 + rate, j + 1);
      }
      
      if (Math.abs(npv) < tolerance) {
        return rate;
      }
      
      // Calculate NPV at previous rate for secant method
      for (let j = 0; j < values.length; j++) {
        prevNpv += values[j] / Math.pow(1 + prevRate, j + 1);
      }
      
      if (Math.abs(npv - prevNpv) < tolerance) {
        return rate;
      }
      
      // Secant method update
      const newRate = rate - npv * (rate - prevRate) / (npv - prevNpv);
      
      if (Math.abs(newRate - rate) < tolerance) {
        return newRate;
      }
      
      // Update rates for next iteration
      prevRate = rate;
      
      // Improved bounds checking
      if (newRate <= -0.999) {
        rate = -0.999;
      } else if (newRate > 10) {
        rate = 10;
      } else {
        rate = newRate;
      }
    }
    
      return FormulaError.NUM; // 収束しない場合
    } catch {
      return FormulaError.VALUE;
    }
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
      
      // PMTを計算 (use Excel-compatible calculation)
      let pmt: number;
      if (type === 0) {
        // 期末払い (ordinary annuity) - exact Excel PMT formula
        const pvFactor = (1 - Math.pow(1 + rate, -nper)) / rate;
        const fvFactor = Math.pow(1 + rate, -nper);
        pmt = -(pv / pvFactor + fv * fvFactor);
      } else {
        // 期初払い (annuity due) - Excel formula for beginning payments
        const pvFactor = (1 - Math.pow(1 + rate, -nper)) / rate * (1 + rate);
        const fvFactor = Math.pow(1 + rate, -nper);
        pmt = -(pv / pvFactor + fv * fvFactor / (1 + rate));
      }
      
      // 指定期間の開始時点での残高を計算
      let balance = pv;
      if (type === 0) {
        // 期末払いの場合
        for (let i = 1; i < per; i++) {
          const ipmt = -balance * rate;
          const ppmt = pmt - ipmt;
          balance = balance + ppmt;
        }
        const ipmt = -balance * rate;
        return pmt - ipmt;
      } else {
        // 期初払いの場合
        if (per === 1) {
          return pmt; // 初回は全額元金
        }
        for (let i = 1; i < per; i++) {
          const ppmt = pmt;
          balance = balance + ppmt;
          const ipmt = -balance * rate;
          balance = balance + ipmt;
        }
        const ipmt = -balance * rate;
        return pmt - ipmt;
      }
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
      
      // PMTを計算 (use Excel-compatible calculation)
      let pmt: number;
      if (type === 0) {
        // 期末払い (ordinary annuity) - exact Excel PMT formula
        const pvFactor = (1 - Math.pow(1 + rate, -nper)) / rate;
        const fvFactor = Math.pow(1 + rate, -nper);
        pmt = -(pv / pvFactor + fv * fvFactor);
      } else {
        // 期初払い (annuity due) - Excel formula for beginning payments
        const pvFactor = (1 - Math.pow(1 + rate, -nper)) / rate * (1 + rate);
        const fvFactor = Math.pow(1 + rate, -nper);
        pmt = -(pv / pvFactor + fv * fvFactor / (1 + rate));
      }
      
      // 指定期間の開始時点での残高を計算して利息を求める
      let balance = pv;
      if (type === 0) {
        // 期末払いの場合
        for (let i = 1; i < per; i++) {
          const ipmt = -balance * rate;
          const ppmt = pmt - ipmt;
          balance = balance + ppmt;
        }
        return -balance * rate;
      } else {
        // 期初払いの場合
        if (per === 1) {
          return 0; // 初回は利息なし
        }
        for (let i = 1; i < per; i++) {
          const ppmt = pmt;
          balance = balance + ppmt;
          const ipmt = -balance * rate;
          balance = balance + ipmt;
        }
        return -balance * rate;
      }
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
        
        if (Math.abs(rate) < precision) {
          f = pv + pmt * nper + fv;
          df = -nper * (nper - 1) * pmt / 2 - nper * pv;
        } else {
          const factor = Math.pow(1 + rate, nper);
          
          if (type === 0) {
            // 期末払い
            f = pv * factor + pmt * (factor - 1) / rate + fv;
            df = nper * pv * Math.pow(1 + rate, nper - 1) + 
                 pmt * ((nper * Math.pow(1 + rate, nper - 1) * rate - (factor - 1)) / (rate * rate));
          } else {
            // 期初払い
            f = pv * factor + pmt * (1 + rate) * (factor - 1) / rate + fv;
            df = nper * pv * Math.pow(1 + rate, nper - 1) + 
                 pmt * (1 + rate) * ((nper * Math.pow(1 + rate, nper - 1) * rate - (factor - 1)) / (rate * rate)) +
                 pmt * (factor - 1) / rate;
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
        
        // Prevent rate from going too negative or too high
        if (newRate <= -0.99999) {
          rate = -0.99999;
        } else if (newRate > 10) {
          rate = 10;
        } else {
          rate = newRate;
        }
      }
      
      return FormulaError.NUM; // 収束しない場合
    } catch {
      return FormulaError.VALUE;
    }
  }
};