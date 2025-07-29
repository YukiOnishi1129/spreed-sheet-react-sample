// 変則期間証券関連の財務関数

import type { CustomFormula, FormulaContext, FormulaResult } from '../shared/types';
import { FormulaError } from '../shared/types';
import { getCellValue } from '../shared/utils';
import { parseDate } from '../shared/dateUtils';

// 日数計算規約（統一版）
function calculateDaysFraction(startDate: Date, endDate: Date, basis: number = 0): number {
  const earlier = startDate <= endDate ? startDate : endDate;
  const later = startDate <= endDate ? endDate : startDate;
  
  switch (basis) {
    case 0: // 30/360 US (NASD)
      return days30_360US(earlier, later);
    case 1: // Actual/actual - 実際の日数を返す
      return Math.floor((later.getTime() - earlier.getTime()) / (24 * 60 * 60 * 1000));
    case 2: // Actual/360 - 実際の日数を返す
      return Math.floor((later.getTime() - earlier.getTime()) / (24 * 60 * 60 * 1000));
    case 3: // Actual/365 - 実際の日数を返す
      return Math.floor((later.getTime() - earlier.getTime()) / (24 * 60 * 60 * 1000));
    case 4: // 30/360 European
      return days30_360European(earlier, later);
    default:
      return days30_360US(earlier, later);
  }
}

// 30/360 US (NASD) 方式の日数計算
function days30_360US(startDate: Date, endDate: Date): number {
  let startYear = startDate.getFullYear();
  let startMonth = startDate.getMonth() + 1;
  let startDay = startDate.getDate();
  
  let endYear = endDate.getFullYear();
  let endMonth = endDate.getMonth() + 1;
  let endDay = endDate.getDate();
  
  // US 30/360の特殊ルール
  if (startDay === getDaysInMonth(startYear, startMonth)) {
    startDay = 30;
  }
  
  if (endDay === getDaysInMonth(endYear, endMonth) && startDay < 30) {
    endDay = getDaysInMonth(endYear, endMonth);
  } else if (endDay === getDaysInMonth(endYear, endMonth) && startDay === 30) {
    endDay = 30;
  }
  
  if (startDay === 31 || (startMonth === 2 && startDay === getDaysInMonth(startYear, startMonth))) {
    startDay = 30;
  }
  
  if (endDay === 31 && startDay === 30) {
    endDay = 30;
  }
  
  return (endYear - startYear) * 360 + (endMonth - startMonth) * 30 + (endDay - startDay);
}

// 30/360 European方式の日数計算
function days30_360European(startDate: Date, endDate: Date): number {
  let startYear = startDate.getFullYear();
  let startMonth = startDate.getMonth() + 1;
  let startDay = startDate.getDate();
  
  let endYear = endDate.getFullYear();
  let endMonth = endDate.getMonth() + 1;
  let endDay = endDate.getDate();
  
  // European 30/360のルール
  if (startDay === 31) startDay = 30;
  if (endDay === 31) endDay = 30;
  
  return (endYear - startYear) * 360 + (endMonth - startMonth) * 30 + (endDay - startDay);
}

// ヘルパー関数：月の日数を取得
function getDaysInMonth(year: number, month: number): number {
  return new Date(year, month, 0).getDate();
}

// 年間日数を取得
function getDaysInYear(basis: number, year?: number): number {
  switch (basis) {
    case 0: // 30/360 US
    case 4: // 30/360 European
      return 360;
    case 1: // Actual/actual
      if (year) {
        return ((year % 4 === 0 && year % 100 !== 0) || (year % 400 === 0)) ? 366 : 365;
      }
      return 365; // デフォルト
    case 2: // Actual/360
      return 360;
    case 3: // Actual/365
      return 365;
    default:
      return 360;
  }
}


// ODDFPRICE関数の実装（変則初回期の価格）
export const ODDFPRICE: CustomFormula = {
  name: 'ODDFPRICE',
  pattern: /ODDFPRICE\(([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, settlementRef, maturityRef, issueRef, firstCouponRef, rateRef, yieldRef, redemptionRef, freqRef, basisRef] = matches;
    
    try {
      const settlement = parseDate(getCellValue(settlementRef.trim(), context)?.toString() ?? settlementRef.trim());
      const maturity = parseDate(getCellValue(maturityRef.trim(), context)?.toString() ?? maturityRef.trim());
      const issue = parseDate(getCellValue(issueRef.trim(), context)?.toString() ?? issueRef.trim());
      const firstCoupon = parseDate(getCellValue(firstCouponRef.trim(), context)?.toString() ?? firstCouponRef.trim());
      const rate = parseFloat(getCellValue(rateRef.trim(), context)?.toString() ?? rateRef.trim());
      const yld = parseFloat(getCellValue(yieldRef.trim(), context)?.toString() ?? yieldRef.trim());
      const redemption = parseFloat(getCellValue(redemptionRef.trim(), context)?.toString() ?? redemptionRef.trim());
      const frequency = parseInt(getCellValue(freqRef.trim(), context)?.toString() ?? freqRef.trim());
      const basis = basisRef ? parseInt(getCellValue(basisRef.trim(), context)?.toString() ?? basisRef.trim()) : 0;
      
      if (!settlement || !maturity || !issue || !firstCoupon) {
        return FormulaError.VALUE;
      }
      
      if (isNaN(rate) || isNaN(yld) || isNaN(redemption) || isNaN(frequency)) {
        return FormulaError.NUM;
      }
      
      if (rate < 0 || yld < 0 || redemption <= 0) {
        return FormulaError.NUM;
      }
      
      if (![1, 2, 4].includes(frequency)) {
        return FormulaError.NUM;
      }
      
      if (settlement >= maturity || issue >= settlement || firstCoupon <= issue || firstCoupon >= maturity) {
        return FormulaError.NUM;
      }
      
      // 簡易実装：変則初回期の価格計算
      const yearDays = getDaysInYear(basis, settlement.getFullYear());
      const r = yld / frequency;
      const coupon = rate * redemption / frequency;
      
      // 変則初回期のクーポン計算
      const daysIF = calculateDaysFraction(issue, firstCoupon, basis);
      const daysSF = calculateDaysFraction(settlement, firstCoupon, basis);
      const normalPeriod = yearDays / frequency;
      
      let price = 0;
      
      // 変則初回期のクーポン
      const oddCoupon = coupon * daysIF / normalPeriod;
      price += oddCoupon / Math.pow(1 + r, daysSF / normalPeriod);
      
      // 通常期のクーポンと償還額
      const n = Math.ceil((calculateDaysFraction(firstCoupon, maturity, 1) / 365) * frequency);
      for (let i = 1; i <= n; i++) {
        const discountFactor = Math.pow(1 + r, i + daysSF / normalPeriod);
        if (i < n) {
          price += coupon / discountFactor;
        } else {
          price += (coupon + redemption) / discountFactor;
        }
      }
      
      return price;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// ODDFYIELD関数の実装（変則初回期の利回り）
export const ODDFYIELD: CustomFormula = {
  name: 'ODDFYIELD',
  pattern: /ODDFYIELD\(([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, settlementRef, maturityRef, issueRef, firstCouponRef, rateRef, priceRef, redemptionRef, freqRef, basisRef] = matches;
    
    try {
      const settlement = parseDate(getCellValue(settlementRef.trim(), context)?.toString() ?? settlementRef.trim());
      const maturity = parseDate(getCellValue(maturityRef.trim(), context)?.toString() ?? maturityRef.trim());
      const issue = parseDate(getCellValue(issueRef.trim(), context)?.toString() ?? issueRef.trim());
      const firstCoupon = parseDate(getCellValue(firstCouponRef.trim(), context)?.toString() ?? firstCouponRef.trim());
      const rate = parseFloat(getCellValue(rateRef.trim(), context)?.toString() ?? rateRef.trim());
      const price = parseFloat(getCellValue(priceRef.trim(), context)?.toString() ?? priceRef.trim());
      const redemption = parseFloat(getCellValue(redemptionRef.trim(), context)?.toString() ?? redemptionRef.trim());
      const frequency = parseInt(getCellValue(freqRef.trim(), context)?.toString() ?? freqRef.trim());
      const basis = basisRef ? parseInt(getCellValue(basisRef.trim(), context)?.toString() ?? basisRef.trim()) : 0;
      
      if (!settlement || !maturity || !issue || !firstCoupon) {
        return FormulaError.VALUE;
      }
      
      if (isNaN(rate) || isNaN(price) || isNaN(redemption) || isNaN(frequency)) {
        return FormulaError.NUM;
      }
      
      if (rate < 0 || price <= 0 || redemption <= 0) {
        return FormulaError.NUM;
      }
      
      if (![1, 2, 4].includes(frequency)) {
        return FormulaError.NUM;
      }
      
      if (settlement >= maturity || issue >= settlement || firstCoupon <= issue || firstCoupon >= maturity) {
        return FormulaError.NUM;
      }
      
      // Newton-Raphson法で利回りを計算（簡易実装）
      let yld = rate; // 初期値
      const maxIterations = 100;
      const tolerance = 0.0000001;
      
      const yearDays = getDaysInYear(basis);
      const coupon = rate * redemption / frequency;
      const daysIF = calculateDaysFraction(issue, firstCoupon, basis);
      const daysSF = calculateDaysFraction(settlement, firstCoupon, basis);
      const normalPeriod = yearDays / frequency;
      const oddCoupon = coupon * daysIF / normalPeriod;
      const n = Math.ceil((actualDays(firstCoupon, maturity) / 365) * frequency);
      
      for (let iter = 0; iter < maxIterations; iter++) {
        const r = yld / frequency;
        
        // 価格を計算
        let calcPrice = oddCoupon / Math.pow(1 + r, daysSF / normalPeriod);
        
        for (let i = 1; i <= n; i++) {
          const discountFactor = Math.pow(1 + r, i + daysSF / normalPeriod);
          if (i < n) {
            calcPrice += coupon / discountFactor;
          } else {
            calcPrice += (coupon + redemption) / discountFactor;
          }
        }
        
        const diff = calcPrice - price;
        
        if (Math.abs(diff) < tolerance) {
          return yld;
        }
        
        // 導関数を近似計算して更新
        yld = yld - diff * 0.01; // 簡易的な更新
      }
      
      return yld;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// ODDLPRICE関数の実装（変則最終期の価格）
export const ODDLPRICE: CustomFormula = {
  name: 'ODDLPRICE',
  pattern: /ODDLPRICE\(([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, settlementRef, maturityRef, lastCouponRef, rateRef, yieldRef, redemptionRef, freqRef, basisRef] = matches;
    
    try {
      const settlement = parseDate(getCellValue(settlementRef.trim(), context)?.toString() ?? settlementRef.trim());
      const maturity = parseDate(getCellValue(maturityRef.trim(), context)?.toString() ?? maturityRef.trim());
      const lastCoupon = parseDate(getCellValue(lastCouponRef.trim(), context)?.toString() ?? lastCouponRef.trim());
      const rate = parseFloat(getCellValue(rateRef.trim(), context)?.toString() ?? rateRef.trim());
      const yld = parseFloat(getCellValue(yieldRef.trim(), context)?.toString() ?? yieldRef.trim());
      const redemption = parseFloat(getCellValue(redemptionRef.trim(), context)?.toString() ?? redemptionRef.trim());
      const frequency = parseInt(getCellValue(freqRef.trim(), context)?.toString() ?? freqRef.trim());
      const basis = basisRef ? parseInt(getCellValue(basisRef.trim(), context)?.toString() ?? basisRef.trim()) : 0;
      
      if (!settlement || !maturity || !lastCoupon) {
        return FormulaError.VALUE;
      }
      
      if (isNaN(rate) || isNaN(yld) || isNaN(redemption) || isNaN(frequency)) {
        return FormulaError.NUM;
      }
      
      if (rate < 0 || yld < 0 || redemption <= 0) {
        return FormulaError.NUM;
      }
      
      if (![1, 2, 4].includes(frequency)) {
        return FormulaError.NUM;
      }
      
      if (settlement >= maturity || lastCoupon >= maturity || lastCoupon <= settlement) {
        return FormulaError.NUM;
      }
      
      // 簡易実装：変則最終期の価格計算
      const yearDays = getDaysInYear(basis);
      const r = yld / frequency;
      const coupon = rate * redemption / frequency;
      
      // 変則最終期のクーポン計算
      const daysLM = calculateDaysFraction(lastCoupon, maturity, basis);
      const daysSM = calculateDaysFraction(settlement, maturity, basis);
      const normalPeriod = yearDays / frequency;
      
      // 変則最終期のクーポンと償還額
      const oddCoupon = coupon * daysLM / normalPeriod;
      const price = (oddCoupon + redemption) / Math.pow(1 + r, daysSM / normalPeriod);
      
      return price;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// ODDLYIELD関数の実装（変則最終期の利回り）
export const ODDLYIELD: CustomFormula = {
  name: 'ODDLYIELD',
  pattern: /ODDLYIELD\(([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, settlementRef, maturityRef, lastCouponRef, rateRef, priceRef, redemptionRef, freqRef, basisRef] = matches;
    
    try {
      const settlement = parseDate(getCellValue(settlementRef.trim(), context)?.toString() ?? settlementRef.trim());
      const maturity = parseDate(getCellValue(maturityRef.trim(), context)?.toString() ?? maturityRef.trim());
      const lastCoupon = parseDate(getCellValue(lastCouponRef.trim(), context)?.toString() ?? lastCouponRef.trim());
      const rate = parseFloat(getCellValue(rateRef.trim(), context)?.toString() ?? rateRef.trim());
      const price = parseFloat(getCellValue(priceRef.trim(), context)?.toString() ?? priceRef.trim());
      const redemption = parseFloat(getCellValue(redemptionRef.trim(), context)?.toString() ?? redemptionRef.trim());
      const frequency = parseInt(getCellValue(freqRef.trim(), context)?.toString() ?? freqRef.trim());
      const basis = basisRef ? parseInt(getCellValue(basisRef.trim(), context)?.toString() ?? basisRef.trim()) : 0;
      
      if (!settlement || !maturity || !lastCoupon) {
        return FormulaError.VALUE;
      }
      
      if (isNaN(rate) || isNaN(price) || isNaN(redemption) || isNaN(frequency)) {
        return FormulaError.NUM;
      }
      
      if (rate < 0 || price <= 0 || redemption <= 0) {
        return FormulaError.NUM;
      }
      
      if (![1, 2, 4].includes(frequency)) {
        return FormulaError.NUM;
      }
      
      if (settlement >= maturity || lastCoupon >= maturity || lastCoupon <= settlement) {
        return FormulaError.NUM;
      }
      
      // 簡易実装：利回りを直接計算
      const yearDays = getDaysInYear(basis);
      const coupon = rate * redemption / frequency;
      const daysLM = calculateDaysFraction(lastCoupon, maturity, basis);
      const daysSM = calculateDaysFraction(settlement, maturity, basis);
      const normalPeriod = yearDays / frequency;
      
      // 変則最終期のクーポン
      const oddCoupon = coupon * daysLM / normalPeriod;
      
      // 利回り = ((変則クーポン + 償還額) / 価格)^(正規期間/経過日数) - 1) × 頻度
      const yld = (Math.pow((oddCoupon + redemption) / price, normalPeriod / daysSM) - 1) * frequency;
      
      return yld;
    } catch {
      return FormulaError.VALUE;
    }
  }
};