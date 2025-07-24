// 追加の財務関数

import type { CustomFormula, FormulaContext, FormulaResult } from '../shared/types';
import { FormulaError } from '../shared/types';
import { getCellValue } from '../shared/utils';

// NPER関数の実装（支払回数）
export const NPER: CustomFormula = {
  name: 'NPER',
  pattern: /NPER\(([^,]+),\s*([^,]+),\s*([^,)]+)(?:,\s*([^,)]+))?(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, rateRef, pmtRef, pvRef, fvRef, typeRef] = matches;
    
    try {
      const rate = parseFloat(getCellValue(rateRef.trim(), context)?.toString() ?? rateRef.trim());
      const pmt = parseFloat(getCellValue(pmtRef.trim(), context)?.toString() ?? pmtRef.trim());
      const pv = parseFloat(getCellValue(pvRef.trim(), context)?.toString() ?? pvRef.trim());
      const fv = fvRef ? parseFloat(getCellValue(fvRef.trim(), context)?.toString() ?? fvRef.trim()) : 0;
      const type = typeRef ? parseInt(getCellValue(typeRef.trim(), context)?.toString() ?? typeRef.trim()) : 0;
      
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
      const rate = parseFloat(getCellValue(rateRef.trim(), context)?.toString() ?? rateRef.trim());
      
      if (isNaN(rate)) {
        return FormulaError.VALUE;
      }
      
      // 値と日付の配列を取得
      const values: number[] = [];
      const dates: number[] = [];
      
      // 範囲参照の場合
      if (valuesRef.includes(':') && datesRef.includes(':')) {
        const valuesRange = valuesRef.trim().match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
        const datesRange = datesRef.trim().match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
        
        if (!valuesRange || !datesRange) {
          return FormulaError.REF;
        }
        
        const [, vStartCol, vStartRowStr, vEndCol, vEndRowStr] = valuesRange;
        const [, dStartCol, dStartRowStr, , dEndRowStr] = datesRange;
        
        const vStartRow = parseInt(vStartRowStr) - 1;
        const vEndRow = parseInt(vEndRowStr) - 1;
        const vStartColIndex = vStartCol.charCodeAt(0) - 65;
        const vEndColIndex = vEndCol.charCodeAt(0) - 65;
        
        const dStartRow = parseInt(dStartRowStr) - 1;
        const dEndRow = parseInt(dEndRowStr) - 1;
        const dStartColIndex = dStartCol.charCodeAt(0) - 65;
        
        // 値を収集
        for (let row = vStartRow; row <= vEndRow; row++) {
          for (let col = vStartColIndex; col <= vEndColIndex; col++) {
            const cell = context.data[row]?.[col];
            if (cell) {
              const val = parseFloat(cell.value?.toString() ?? '');
              if (!isNaN(val)) {
                values.push(val);
              }
            }
          }
        }
        
        // 日付を収集
        for (let row = dStartRow; row <= dEndRow; row++) {
          for (let col = dStartColIndex; col <= dStartColIndex + (vEndColIndex - vStartColIndex); col++) {
            const cell = context.data[row]?.[col];
            if (cell) {
              const dateVal = parseFloat(cell.value?.toString() ?? '');
              if (!isNaN(dateVal)) {
                dates.push(dateVal);
              }
            }
          }
        }
      }
      
      if (values.length === 0 || values.length !== dates.length) {
        return FormulaError.VALUE;
      }
      
      // 最初の日付を基準にNPVを計算
      const firstDate = dates[0];
      let npv = 0;
      
      for (let i = 0; i < values.length; i++) {
        const daysDiff = dates[i] - firstDate;
        const yearsDiff = daysDiff / 365;
        npv += values[i] / Math.pow(1 + rate, yearsDiff);
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
        const valuesRange = valuesRef.trim().match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
        const datesRange = datesRef.trim().match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
        
        if (!valuesRange || !datesRange) {
          return FormulaError.REF;
        }
        
        const [, vStartCol, vStartRowStr, vEndCol, vEndRowStr] = valuesRange;
        const [, dStartCol, dStartRowStr, , dEndRowStr] = datesRange;
        
        const vStartRow = parseInt(vStartRowStr) - 1;
        const vEndRow = parseInt(vEndRowStr) - 1;
        const vStartColIndex = vStartCol.charCodeAt(0) - 65;
        const vEndColIndex = vEndCol.charCodeAt(0) - 65;
        
        const dStartRow = parseInt(dStartRowStr) - 1;
        const dEndRow = parseInt(dEndRowStr) - 1;
        const dStartColIndex = dStartCol.charCodeAt(0) - 65;
        
        // 値を収集
        for (let row = vStartRow; row <= vEndRow; row++) {
          for (let col = vStartColIndex; col <= vEndColIndex; col++) {
            const cell = context.data[row]?.[col];
            if (cell) {
              const val = parseFloat(cell.value?.toString() ?? '');
              if (!isNaN(val)) {
                values.push(val);
              }
            }
          }
        }
        
        // 日付を収集
        for (let row = dStartRow; row <= dEndRow; row++) {
          for (let col = dStartColIndex; col <= dStartColIndex + (vEndColIndex - vStartColIndex); col++) {
            const cell = context.data[row]?.[col];
            if (cell) {
              const dateVal = parseFloat(cell.value?.toString() ?? '');
              if (!isNaN(dateVal)) {
                dates.push(dateVal);
              }
            }
          }
        }
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
        const rangeMatch = valuesRef.trim().match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
        if (!rangeMatch) {
          return FormulaError.REF;
        }
        
        const [, startCol, startRowStr, endCol, endRowStr] = rangeMatch;
        const startRow = parseInt(startRowStr) - 1;
        const endRow = parseInt(endRowStr) - 1;
        const startColIndex = startCol.charCodeAt(0) - 65;
        const endColIndex = endCol.charCodeAt(0) - 65;
        
        for (let row = startRow; row <= endRow; row++) {
          for (let col = startColIndex; col <= endColIndex; col++) {
            const cell = context.data[row]?.[col];
            if (cell) {
              const val = parseFloat(cell.value?.toString() ?? '');
              if (!isNaN(val)) {
                values.push(val);
              }
            }
          }
        }
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
  pattern: /DB\(([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
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
  pattern: /DDB\(([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
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
      const rate = parseFloat(getCellValue(rateRef.trim(), context)?.toString() ?? rateRef.trim());
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