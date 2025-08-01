// 追加の財務関数

import type { CustomFormula, FormulaContext, FormulaResult } from '../shared/types';
import { FormulaError } from '../shared/types';
import { getCellValue, getCellRangeValues } from '../shared/utils';

// 値または式を評価する
function evaluateValueOrExpression(value: string, context: FormulaContext): number {
  // 簡単な四則演算を処理（例: A2/12, A2*12）
  const operatorMatch = value.match(/^([^+\-*/]+)([+\-*/])(.+)$/);
  if (operatorMatch) {
    const [, left, operator, right] = operatorMatch;
    const leftValue = evaluateValueOrExpression(left.trim(), context);
    const rightValue = evaluateValueOrExpression(right.trim(), context);
    
    switch (operator) {
      case '+': return leftValue + rightValue;
      case '-': return leftValue - rightValue;
      case '*': return leftValue * rightValue;
      case '/': return rightValue !== 0 ? leftValue / rightValue : NaN;
    }
  }
  
  // 負の符号で始まる場合の処理（例: -A2）
  if (value.startsWith('-')) {
    const baseValue = value.substring(1);
    const result = evaluateValueOrExpression(baseValue, context);
    return -result;
  }
  
  // セル参照かどうか確認
  const cellValue = getCellValue(value, context);
  if (cellValue !== value) {
    // セル参照だった場合
    return parseFloat(cellValue?.toString() ?? '0');
  }
  
  // 数値として解析
  return parseFloat(value);
}

// NPER関数の実装（支払回数）
export const NPER: CustomFormula = {
  name: 'NPER',
  pattern: /NPER\(([^,]+),\s*([^,]+),\s*([^,)]+)(?:,\s*([^,)]+))?(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, rateRef, pmtRef, pvRef, fvRef, typeRef] = matches;
    
    try {
      const rate = evaluateValueOrExpression(rateRef.trim(), context);
      const pmt = evaluateValueOrExpression(pmtRef.trim(), context);
      const pv = evaluateValueOrExpression(pvRef.trim(), context);
      const fv = fvRef ? evaluateValueOrExpression(fvRef.trim(), context) : 0;
      const type = typeRef ? Math.round(evaluateValueOrExpression(typeRef.trim(), context)) : 0;
      
      if (isNaN(rate) || isNaN(pmt) || isNaN(pv)) {
        return FormulaError.VALUE;
      }
      
      if (type !== 0 && type !== 1) {
        return FormulaError.VALUE;
      }
      
      // 利率が0の場合の特殊処理
      if (rate === 0) {
        if (pmt === 0) {
          return FormulaError.NUM;
        }
        return -(pv + fv) / pmt;
      }
      
      // NPER = log((pmt * (1 + rate * type) - fv * rate) / (pmt * (1 + rate * type) + pv * rate)) / log(1 + rate)
      const numerator = pmt * (1 + rate * type) - fv * rate;
      const denominator = pmt * (1 + rate * type) + pv * rate;
      
      if (denominator === 0 || numerator / denominator <= 0) {
        return FormulaError.NUM;
      }
      
      return Math.log(numerator / denominator) / Math.log(1 + rate);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// XNPV関数の実装（正味現在価値：日付指定）
export const XNPV: CustomFormula = {
  name: 'XNPV',
  pattern: /XNPV\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, rateRef, valuesRef, datesRef] = matches;
    
    try {
      const rate = evaluateValueOrExpression(rateRef.trim(), context);
      
      if (isNaN(rate)) {
        return FormulaError.VALUE;
      }
      
      // 値と日付の配列を取得
      const values: number[] = [];
      const dates: number[] = [];
      
      // 範囲参照の場合
      if (valuesRef.includes(':') && datesRef.includes(':')) {
        // 値を収集
        const valuesArray = getCellRangeValues(valuesRef.trim(), context);
        valuesArray.forEach(value => {
          const num = typeof value === 'string' ? parseFloat(value) : Number(value);
          if (!isNaN(num)) {
            values.push(num);
          }
        });
        
        // 日付を収集
        const datesArray = getCellRangeValues(datesRef.trim(), context);
        datesArray.forEach(value => {
          const num = typeof value === 'string' ? parseFloat(value) : Number(value);
          if (!isNaN(num)) {
            dates.push(num);
          }
        });
      } else {
        // If not a proper range format, return REF error
        return FormulaError.REF;
      }
      
      if (values.length === 0 || values.length !== dates.length) {
        return FormulaError.VALUE;
      }
      
      
      // 最初の日付を基準にNPVを計算
      const firstDate = dates[0];
      let npv = 0;
      
      // When rate is 0, just sum the values
      if (rate === 0) {
        for (let i = 0; i < values.length; i++) {
          npv += values[i];
        }
      } else {
        for (let i = 0; i < values.length; i++) {
          const daysDiff = dates[i] - firstDate;
          const yearsDiff = daysDiff / 365;
          npv += values[i] / Math.pow(1 + rate, yearsDiff);
        }
      }
      
      return npv;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// XIRR関数の実装（内部収益率：日付指定）
export const XIRR: CustomFormula = {
  name: 'XIRR',
  pattern: /XIRR\(([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, valuesRef, datesRef, guessRef] = matches;
    
    try {
      let guess = guessRef ? parseFloat(getCellValue(guessRef.trim(), context)?.toString() ?? guessRef.trim()) : 0.1;
      
      if (isNaN(guess)) {
        guess = 0.1;
      }
      
      // 値と日付の配列を取得（XNPVと同様）
      const values: number[] = [];
      const dates: number[] = [];
      
      // 範囲参照の場合
      if (valuesRef.includes(':') && datesRef.includes(':')) {
        // 値を収集
        const valuesArray = getCellRangeValues(valuesRef.trim(), context);
        valuesArray.forEach(value => {
          const num = typeof value === 'string' ? parseFloat(value) : Number(value);
          if (!isNaN(num)) {
            values.push(num);
          }
        });
        
        // 日付を収集
        const datesArray = getCellRangeValues(datesRef.trim(), context);
        datesArray.forEach(value => {
          const num = typeof value === 'string' ? parseFloat(value) : Number(value);
          if (!isNaN(num)) {
            dates.push(num);
          }
        });
      } else {
        // If not a proper range format, return REF error
        return FormulaError.REF;
      }
      
      if (values.length === 0 || values.length !== dates.length) {
        return FormulaError.VALUE;
      }
      
      // Newton-Raphson法でIRRを計算
      let rate = guess;
      const firstDate = dates[0];
      const maxIterations = 100;
      const tolerance = 1e-10;
      
      for (let iter = 0; iter < maxIterations; iter++) {
        let npv = 0;
        let dnpv = 0;
        
        for (let i = 0; i < values.length; i++) {
          const yearsDiff = (dates[i] - firstDate) / 365;
          const factor = Math.pow(1 + rate, yearsDiff);
          npv += values[i] / factor;
          dnpv -= values[i] * yearsDiff / (factor * (1 + rate));
        }
        
        if (Math.abs(npv) < tolerance) {
          return rate;
        }
        
        if (dnpv === 0) {
          return FormulaError.NUM;
        }
        
        rate = rate - npv / dnpv;
        
        // 発散チェック
        if (rate < -0.99 || rate > 10) {
          return FormulaError.NUM;
        }
      }
      
      return FormulaError.NUM;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// MIRR関数の実装（修正内部収益率）
export const MIRR: CustomFormula = {
  name: 'MIRR',
  pattern: /MIRR\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, valuesRef, financeRateRef, reinvestRateRef] = matches;
    
    try {
      const financeRate = parseFloat(getCellValue(financeRateRef.trim(), context)?.toString() ?? financeRateRef.trim());
      const reinvestRate = parseFloat(getCellValue(reinvestRateRef.trim(), context)?.toString() ?? reinvestRateRef.trim());
      
      if (isNaN(financeRate) || isNaN(reinvestRate)) {
        return FormulaError.VALUE;
      }
      
      // 値の配列を取得
      const values: number[] = [];
      
      if (valuesRef.includes(':')) {
        const valuesArray = getCellRangeValues(valuesRef.trim(), context);
        valuesArray.forEach(value => {
          const num = typeof value === 'string' ? parseFloat(value) : Number(value);
          if (!isNaN(num)) {
            values.push(num);
          }
        });
      }
      
      if (values.length < 2) {
        return FormulaError.VALUE;
      }
      
      // 正と負のキャッシュフローを分離
      let positiveNPV = 0;
      let negativeNPV = 0;
      const n = values.length;
      
      for (let i = 0; i < n; i++) {
        if (values[i] > 0) {
          positiveNPV += values[i] / Math.pow(1 + reinvestRate, i);
        } else if (values[i] < 0) {
          negativeNPV += values[i] / Math.pow(1 + financeRate, i);
        }
      }
      
      if (negativeNPV === 0 || positiveNPV === 0) {
        return FormulaError.DIV0;
      }
      
      // MIRR = ((FV of positive cash flows / PV of negative cash flows)^(1/n)) - 1
      const fvPositive = positiveNPV * Math.pow(1 + reinvestRate, n - 1);
      const pvNegative = -negativeNPV;
      
      return Math.pow(fvPositive / pvNegative, 1 / (n - 1)) - 1;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// SLN関数の実装（定額法による減価償却）
export const SLN: CustomFormula = {
  name: 'SLN',
  pattern: /SLN\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, costRef, salvageRef, lifeRef] = matches;
    
    try {
      const cost = parseFloat(getCellValue(costRef.trim(), context)?.toString() ?? costRef.trim());
      const salvage = parseFloat(getCellValue(salvageRef.trim(), context)?.toString() ?? salvageRef.trim());
      const life = parseFloat(getCellValue(lifeRef.trim(), context)?.toString() ?? lifeRef.trim());
      
      if (isNaN(cost) || isNaN(salvage) || isNaN(life)) {
        return FormulaError.VALUE;
      }
      
      if (life <= 0) {
        return FormulaError.NUM;
      }
      
      // 定額法: (取得価額 - 残存価額) / 耐用年数
      return (cost - salvage) / life;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// SYD関数の実装（級数法による減価償却）
export const SYD: CustomFormula = {
  name: 'SYD',
  pattern: /SYD\(([^,]+),\s*([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, costRef, salvageRef, lifeRef, perRef] = matches;
    
    try {
      const cost = parseFloat(getCellValue(costRef.trim(), context)?.toString() ?? costRef.trim());
      const salvage = parseFloat(getCellValue(salvageRef.trim(), context)?.toString() ?? salvageRef.trim());
      const life = parseFloat(getCellValue(lifeRef.trim(), context)?.toString() ?? lifeRef.trim());
      const per = parseFloat(getCellValue(perRef.trim(), context)?.toString() ?? perRef.trim());
      
      if (isNaN(cost) || isNaN(salvage) || isNaN(life) || isNaN(per)) {
        return FormulaError.VALUE;
      }
      
      if (life <= 0 || per <= 0 || per > life) {
        return FormulaError.NUM;
      }
      
      // 級数法: (取得価額 - 残存価額) * (耐用年数 - 期 + 1) * 2 / (耐用年数 * (耐用年数 + 1))
      return (cost - salvage) * (life - per + 1) * 2 / (life * (life + 1));
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// DB関数の実装（定率法による減価償却）
export const DB: CustomFormula = {
  name: 'DB',
  pattern: /\bDB\(([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, costRef, salvageRef, lifeRef, periodRef, monthRef] = matches;
    
    try {
      const cost = parseFloat(getCellValue(costRef.trim(), context)?.toString() ?? costRef.trim());
      const salvage = parseFloat(getCellValue(salvageRef.trim(), context)?.toString() ?? salvageRef.trim());
      const life = parseFloat(getCellValue(lifeRef.trim(), context)?.toString() ?? lifeRef.trim());
      const period = parseFloat(getCellValue(periodRef.trim(), context)?.toString() ?? periodRef.trim());
      const month = monthRef ? parseFloat(getCellValue(monthRef.trim(), context)?.toString() ?? monthRef.trim()) : 12;
      
      if (isNaN(cost) || isNaN(salvage) || isNaN(life) || isNaN(period)) {
        return FormulaError.VALUE;
      }
      
      if (cost < 0 || salvage < 0 || life <= 0 || period < 0 || period > life || month <= 0 || month > 12) {
        return FormulaError.NUM;
      }
      
      // 定率法の率を計算
      const rate = 1 - Math.pow(salvage / cost, 1 / life);
      const roundedRate = Math.round(rate * 1000) / 1000;
      
      let depreciation = 0;
      let totalDepreciation = 0;
      let bookValue = cost;
      
      // 各期の減価償却を計算
      for (let i = 1; i <= period; i++) {
        if (i === 1) {
          // 第1期は月割り計算
          depreciation = bookValue * roundedRate * month / 12;
        } else if (i === Math.floor(life) + 1) {
          // 最終期は残額
          depreciation = bookValue - salvage;
        } else {
          depreciation = bookValue * roundedRate;
        }
        
        totalDepreciation += depreciation;
        bookValue = cost - totalDepreciation;
        
        // 簿価が残存価額を下回らないようにする
        if (bookValue < salvage) {
          depreciation -= (salvage - bookValue);
          bookValue = salvage;
        }
      }
      
      return depreciation;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// DDB関数の実装（倍率法による減価償却）
export const DDB: CustomFormula = {
  name: 'DDB',
  pattern: /\bDDB\(([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, costRef, salvageRef, lifeRef, periodRef, factorRef] = matches;
    
    try {
      const cost = parseFloat(getCellValue(costRef.trim(), context)?.toString() ?? costRef.trim());
      const salvage = parseFloat(getCellValue(salvageRef.trim(), context)?.toString() ?? salvageRef.trim());
      const life = parseFloat(getCellValue(lifeRef.trim(), context)?.toString() ?? lifeRef.trim());
      const period = parseFloat(getCellValue(periodRef.trim(), context)?.toString() ?? periodRef.trim());
      const factor = factorRef ? parseFloat(getCellValue(factorRef.trim(), context)?.toString() ?? factorRef.trim()) : 2;
      
      if (isNaN(cost) || isNaN(salvage) || isNaN(life) || isNaN(period) || isNaN(factor)) {
        return FormulaError.VALUE;
      }
      
      if (cost < 0 || salvage < 0 || life <= 0 || period < 0 || period > life || factor <= 0) {
        return FormulaError.NUM;
      }
      
      // 倍率法の率
      const rate = factor / life;
      
      let bookValue = cost;
      let depreciation = 0;
      
      for (let i = 1; i <= period; i++) {
        // 減価償却費 = 簿価 × 率
        depreciation = bookValue * rate;
        
        // 簿価が残存価額を下回らないようにする
        if (bookValue - depreciation < salvage) {
          depreciation = bookValue - salvage;
        }
        
        if (i < period) {
          bookValue -= depreciation;
        }
      }
      
      return depreciation;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// VDB関数の実装（可変定率法による減価償却）
export const VDB: CustomFormula = {
  name: 'VDB',
  pattern: /VDB\(([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,)]+)(?:,\s*([^,)]+))?(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, costRef, salvageRef, lifeRef, startPeriodRef, endPeriodRef, factorRef, noSwitchRef] = matches;
    
    try {
      const cost = parseFloat(getCellValue(costRef.trim(), context)?.toString() ?? costRef.trim());
      const salvage = parseFloat(getCellValue(salvageRef.trim(), context)?.toString() ?? salvageRef.trim());
      const life = parseFloat(getCellValue(lifeRef.trim(), context)?.toString() ?? lifeRef.trim());
      const startPeriod = parseFloat(getCellValue(startPeriodRef.trim(), context)?.toString() ?? startPeriodRef.trim());
      const endPeriod = parseFloat(getCellValue(endPeriodRef.trim(), context)?.toString() ?? endPeriodRef.trim());
      const factor = factorRef ? parseFloat(getCellValue(factorRef.trim(), context)?.toString() ?? factorRef.trim()) : 2;
      const noSwitch = noSwitchRef ? getCellValue(noSwitchRef.trim(), context)?.toString().toLowerCase() === 'true' : false;
      
      if (isNaN(cost) || isNaN(salvage) || isNaN(life) || isNaN(startPeriod) || isNaN(endPeriod) || isNaN(factor)) {
        return FormulaError.VALUE;
      }
      
      if (cost < 0 || salvage < 0 || life <= 0 || startPeriod < 0 || endPeriod > life || startPeriod >= endPeriod || factor <= 0) {
        return FormulaError.NUM;
      }
      
      let totalDepreciation = 0;
      let bookValue = cost;
      const ddbRate = factor / life;
      const slnDepreciation = (cost - salvage) / life;
      
      for (let period = 0; period < endPeriod; period++) {
        let periodDepreciation: number;
        
        // DDBまたはSLNのうち大きい方を使用（noSwitchがfalseの場合）
        const ddbDepreciation = bookValue * ddbRate;
        
        if (noSwitch) {
          periodDepreciation = ddbDepreciation;
        } else {
          periodDepreciation = Math.max(ddbDepreciation, slnDepreciation);
        }
        
        // 簿価が残存価額を下回らないようにする
        if (bookValue - periodDepreciation < salvage) {
          periodDepreciation = bookValue - salvage;
        }
        
        // 指定された期間内の減価償却を合計
        if (period >= startPeriod) {
          if (period < Math.floor(startPeriod)) {
            // 開始期間が小数の場合、按分計算
            const fraction = startPeriod - Math.floor(startPeriod);
            totalDepreciation += periodDepreciation * (1 - fraction);
          } else if (period >= Math.floor(endPeriod)) {
            // 終了期間が小数の場合、按分計算
            const fraction = endPeriod - Math.floor(endPeriod);
            totalDepreciation += periodDepreciation * fraction;
          } else {
            totalDepreciation += periodDepreciation;
          }
        }
        
        bookValue -= periodDepreciation;
      }
      
      return totalDepreciation;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// PDURATION関数の実装（投資期間）
export const PDURATION: CustomFormula = {
  name: 'PDURATION',
  pattern: /PDURATION\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, rateRef, pvRef, fvRef] = matches;
    
    try {
      const rate = evaluateValueOrExpression(rateRef.trim(), context);
      const pv = parseFloat(getCellValue(pvRef.trim(), context)?.toString() ?? pvRef.trim());
      const fv = parseFloat(getCellValue(fvRef.trim(), context)?.toString() ?? fvRef.trim());
      
      if (isNaN(rate) || isNaN(pv) || isNaN(fv)) {
        return FormulaError.VALUE;
      }
      
      if (rate <= 0 || pv <= 0 || fv <= 0) {
        return FormulaError.NUM;
      }
      
      // PDURATION = log(FV/PV) / log(1 + rate)
      return Math.log(fv / pv) / Math.log(1 + rate);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// RRI関数の実装（投資の成長率）
export const RRI: CustomFormula = {
  name: 'RRI',
  pattern: /RRI\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, nperRef, pvRef, fvRef] = matches;
    
    try {
      const nper = parseFloat(getCellValue(nperRef.trim(), context)?.toString() ?? nperRef.trim());
      const pv = parseFloat(getCellValue(pvRef.trim(), context)?.toString() ?? pvRef.trim());
      const fv = parseFloat(getCellValue(fvRef.trim(), context)?.toString() ?? fvRef.trim());
      
      if (isNaN(nper) || isNaN(pv) || isNaN(fv)) {
        return FormulaError.VALUE;
      }
      
      if (nper <= 0 || pv === 0) {
        return FormulaError.NUM;
      }
      
      // RRI = (FV/PV)^(1/NPER) - 1
      return Math.pow(Math.abs(fv / pv), 1 / nper) - 1;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// ISPMT関数の実装（元金均等返済の利息）
export const ISPMT: CustomFormula = {
  name: 'ISPMT',
  pattern: /ISPMT\(([^,]+),\s*([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, rateRef, perRef, nperRef, pvRef] = matches;
    
    try {
      const rate = evaluateValueOrExpression(rateRef.trim(), context);
      const per = parseInt(getCellValue(perRef.trim(), context)?.toString() ?? perRef.trim());
      const nper = parseInt(getCellValue(nperRef.trim(), context)?.toString() ?? nperRef.trim());
      const pv = parseFloat(getCellValue(pvRef.trim(), context)?.toString() ?? pvRef.trim());
      
      if (isNaN(rate) || isNaN(per) || isNaN(nper) || isNaN(pv)) {
        return FormulaError.VALUE;
      }
      
      if (nper === 0) {
        return FormulaError.NUM;
      }
      
      if (per < 1 || per > nper) {
        return FormulaError.NUM;
      }
      
      // 元金均等返済の利息 = 元金 × 利率 × (1 - (期 - 1) / 期間数)
      // これは残債に対する利息を計算
      const remainingBalance = pv * (1 - (per - 1) / nper);
      return -remainingBalance * rate;
    } catch {
      return FormulaError.VALUE;
    }
  }
};