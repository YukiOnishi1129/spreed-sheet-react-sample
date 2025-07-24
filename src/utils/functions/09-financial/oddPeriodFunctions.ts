// 変則期間証券関連の財務関数

import type { CustomFormula, FormulaContext, FormulaResult } from '../shared/types';
import { FormulaError } from '../shared/types';
import { getCellValue } from '../shared/utils';
import { parseDate } from '../shared/dateUtils';

// 実際の日数計算
function actualDays(startDate: Date, endDate: Date): number {
  const diffTime = endDate.getTime() - startDate.getTime();
  return Math.floor(diffTime / (1000 * 60 * 60 * 24));
}

// 30/360方式での日数計算
function days360(startDate: Date, endDate: Date, european: boolean = false): number {
  let startDay = startDate.getDate();
  let startMonth = startDate.getMonth() + 1;
  let startYear = startDate.getFullYear();
  let endDay = endDate.getDate();
  let endMonth = endDate.getMonth() + 1;
  let endYear = endDate.getFullYear();

  if (!european) {
    // US方式
    if (startDay === 31) startDay = 30;
    if (endDay === 31 && startDay === 30) endDay = 30;
  } else {
    // European方式
    if (startDay === 31) startDay = 30;
    if (endDay === 31) endDay = 30;
  }

  return (endYear - startYear) * 360 + (endMonth - startMonth) * 30 + (endDay - startDay);
}

// 日付ベースの日数計算
function getDayCount(startDate: Date, endDate: Date, basis: number): number {
  switch (basis) {
    case 0: // 30/360 US
      return days360(startDate, endDate, false);
    case 1: // Actual/actual
    case 2: // Actual/360
    case 3: // Actual/365
      return actualDays(startDate, endDate);
    case 4: // 30/360 European
      return days360(startDate, endDate, true);
    default:
      return days360(startDate, endDate, false);
  }
}

// 年間日数を取得
function getYearDays(basis: number): number {
  switch (basis) {
    case 0: // 30/360 US
    case 4: // 30/360 European
      return 360;
    case 1: // Actual/actual
    case 3: // Actual/365
      return 365;
    case 2: // Actual/360
      return 360;
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
      const yearDays = getYearDays(basis);
      const r = yld / frequency;
      const coupon = rate * redemption / frequency;
      
      // 変則初回期のクーポン計算
      const daysIF = getDayCount(issue, firstCoupon, basis);
      const daysSF = getDayCount(settlement, firstCoupon, basis);
      const normalPeriod = yearDays / frequency;
      
      let price = 0;
      
      // 変則初回期のクーポン
      const oddCoupon = coupon * daysIF / normalPeriod;
      price += oddCoupon / Math.pow(1 + r, daysSF / normalPeriod);
      
      // 通常期のクーポンと償還額
      const n = Math.ceil((actualDays(firstCoupon, maturity) / 365) * frequency);
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
      
      const yearDays = getYearDays(basis);
      const coupon = rate * redemption / frequency;
      const daysIF = getDayCount(issue, firstCoupon, basis);
      const daysSF = getDayCount(settlement, firstCoupon, basis);
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
      const yearDays = getYearDays(basis);
      const r = yld / frequency;
      const coupon = rate * redemption / frequency;
      
      // 変則最終期のクーポン計算
      const daysLM = getDayCount(lastCoupon, maturity, basis);
      const daysSM = getDayCount(settlement, maturity, basis);
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
      const yearDays = getYearDays(basis);
      const coupon = rate * redemption / frequency;
      const daysLM = getDayCount(lastCoupon, maturity, basis);
      const daysSM = getDayCount(settlement, maturity, basis);
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